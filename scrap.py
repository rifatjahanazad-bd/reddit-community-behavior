"""Reddit scraper (Selenium + new Reddit UI) — human-readable guide.

What it does (in plain English):
  1. Asks you for categories (e.g. "Gaming") and group links (e.g. r/gaming URLs).
  2. Opens each group page and scrolls down to collect the newest posts.
  3. Opens every post page to read its title, text, score, and comments.
  4. Saves one Excel file per group into the output/ folder.

Example run:
  How many categories? 1
  Category 1 name: Gaming
     [Gaming] group link 1/3: https://www.reddit.com/r/gaming/
  -> output/Gaming__r_gaming.xlsx

Each Excel row = one post, with columns like post_url, username, title,
content, post_score, comment counts, and post_status (normal/stickied/locked).

Requires: Chrome + matching chromedriver, pandas, selenium, openpyxl.
Run:  python scrap.py
"""

# Standard library: JSON parsing, file paths, regex, delays, timestamps, URLs.
import json
import os
import re
import time
import random
from datetime import datetime
from urllib.parse import urlparse

# Third-party: pandas for Excel export, Selenium for browser automation.
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# ----------------------------------------------------------------------------
# SETTINGS — tweak these before a run
# ----------------------------------------------------------------------------
CHROMEDRIVER_PATH = r""        # e.g. r"C:\tools\chromedriver.exe". Leave empty to let Selenium find it.
POSTS_PER_GROUP = 10           # how many newest posts to scrape per group link
LINKS_PER_CATEGORY = 3         # how many group links are asked per category
HEADLESS = False               # keep False: new Reddit is more likely to block headless browsers
FETCH_PROFILE = False          # True = also look up each author's karma / account age (slower, may be blocked)
DELAY_RANGE = (2.0, 4.0)       # random pause (seconds) between page loads (politeness + anti-block)
SCROLL_DELAY = (1.5, 2.5)      # pause after each scroll so new posts/comments can load
MAX_LISTING_STALLS = 4         # stop scrolling after this many scrolls that add no new posts
COMMENT_SCROLLS = 5            # how many times to scroll down a post page to load more comments
SAVE_EVERY = 10                # write the Excel file every N posts, so a crash loses nothing
OUTPUT_DIR = "output"          # folder for the per-group .xlsx files
DEBUG = False                  # True = print attribute names of the first post (helps fix selectors)


# ----------------------------------------------------------------------------
# HELPERS — small utilities used by the steps below
# ----------------------------------------------------------------------------
def pause():
    """Sleep a random short time to look human and avoid rate limits."""
    time.sleep(random.uniform(*DELAY_RANGE))


def make_driver():
    """Create a Chrome WebDriver with anti-automation flags disabled."""
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1280,900")
    # Hide the "automation controlled" blink feature so Reddit blocks less often.
    opts.add_argument("--disable-blink-features=AutomationControlled")
    service = Service(CHROMEDRIVER_PATH) if CHROMEDRIVER_PATH else Service()
    return webdriver.Chrome(service=service, options=opts)


def to_new_reddit(url):
    """Any reddit link -> https://www.reddit.com/<same path>."""
    url = url.strip()
    if not url.startswith("http"):
        url = "https://" + url
    parsed = urlparse(url)
    query = f"?{parsed.query}" if parsed.query else ""
    return f"https://www.reddit.com{parsed.path}{query}"


def safe_name(text):
    """Convert arbitrary text into a safe, friendly filename for Excel files.

    Operating systems dislike characters like '/', ':', '?', or spaces in filenames.
    This turns strings like 'Tech & Gadgets / 2024!' into 'Tech_Gadgets_2024'.

    Steps:
      1. Replace any special characters or spaces with an underscore.
      2. Trim leading/trailing underscores for a cleaner look.
      3. Keep up to 60 characters to avoid hitting operating system path limits.
      4. Fall back to 'group' if the name ends up empty.
    """
    if not text:
        return "group"

    # Step 1: Replace spaces, slashes, and symbols with underscores
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "_", str(text))

    # Step 2: Strip extra underscores from edges (e.g., '_Gaming_' -> 'Gaming')
    cleaned = cleaned.strip("_")

    # Step 3: Limit length so long category/group names don't create overly long file paths
    cleaned = cleaned[:60]

    # Step 4: Return clean name, or fallback if string had only special characters
    return cleaned if cleaned else "group"


def to_int(value, default=0):
    """Parse Reddit count strings like '1,234' safely; fall back to default."""
    try:
        return int(str(value).replace(",", "").strip())
    except (ValueError, TypeError):
        return default


def classify_post(title, flair, body="", promoted=False):
    """Categorize a Reddit post based on keywords in its title and flair.

    This assigns a high-level label to help organize and analyze posts:
      - 'promote'  -> Showcases, self-promotion, projects, launches, deals, giveaways
      - 'help'     -> Troubleshooting, bug reports, errors, asking for technical assistance
      - 'question' -> Inquiries, discussions (ends with '?' or starts with 'How', 'Why', etc.)
      - 'fun'      -> Memes, jokes, humor, casual fluff, shitposts
      - 'general'  -> Regular news, stories, and general discussions that don't match above

    Why the checking order matters:
      1. 'promote' comes FIRST:
         Example: "I built a new app! Any feedback?"
         -> Ends with a '?', but the real intent is self-promotion, not a simple question.
      2. 'help' comes SECOND:
         Example: "Need help: why is this crashing?"
         -> Starts with 'why', but the intent is troubleshooting, not a casual discussion.
      3. 'question' comes THIRD:
         Example: "What are your favorite Reddit communities?"
      4. 'fun' comes FOURTH:
         Example: "Relatable meme for Monday morning lol"
      5. 'general' is the FALLBACK:
         Any post that does not trigger the specific keyword buckets above.
    """
    # Standardize title and flair to lowercase so comparisons are case-insensitive
    clean_title = str(title or "").strip().lower()
    combined_text = f"{clean_title} {str(flair or '').lower()}"

    # ------------------------------------------------------------------------
    # 1. Self-Promotion, Projects & Deals ('promote')
    # Triggers when creators share their apps, projects, channels, or discounts.
    # ------------------------------------------------------------------------
    promo_keywords = (
        # Project showcases & creations
        "my project", "i made", "i built", "i created", "i developed",
        "check out", "launch", "giveaway", "side project", "beta test",
        "open source project", "feedback on my",
        # Marketing, channels & social links
        "discount", "promo", "subscribe", "my channel", "my app", "my website",
        "my newsletter", "my blog", "my video", "[promo", "self promotion",
        "self-promotion", "advert", "patreon", "kickstarter", "coupon",
        "affiliate", "sale ", "free access"
    )
    if promoted or any(keyword in combined_text for keyword in promo_keywords):
        return "promote"

    # ------------------------------------------------------------------------
    # 2. Seeking Assistance & Troubleshooting ('help')
    # Triggers when something is broken, throwing errors, or someone is stuck.
    # ------------------------------------------------------------------------
    help_keywords = (
        "help", "need ", "issue", "error", "problem", "stuck",
        "advice", "support", "can't", "cannot", "not working",
        "bug", "crash", "crashes", "freezing", "broken",
        "won't work", "wont work", "failed", "how do i fix",
        "anyone know why", "troubleshoot", "tech support"
    )
    if any(keyword in combined_text for keyword in help_keywords):
        return "help"

    # ------------------------------------------------------------------------
    # 3. Questions & Inquiries ('question')
    # Triggers if the post ends with '?', contains 'question', or starts with
    # common English question words.
    # ------------------------------------------------------------------------
    question_starters = (
        "how ", "what ", "why ", "when ", "where ",
        "which ", "who ", "is ", "are ", "can ",
        "does ", "do ", "should ", "could ", "would ",
        "has anyone ", "have you ", "anyone else ", "is there a ",
        "thoughts on ", "opinion on "
    )
    has_question_mark = clean_title.endswith("?")
    starts_with_question_word = clean_title.startswith(question_starters)
    has_question_flair_or_tag = "question" in combined_text

    if has_question_mark or starts_with_question_word or has_question_flair_or_tag:
        return "question"

    # ------------------------------------------------------------------------
    # 4. Humor, Memes & Entertainment ('fun')
    # Casual community entertainment, jokes, and memes.
    # ------------------------------------------------------------------------
    fun_keywords = (
        # Classic humor & laughter
        "meme", "lol", "funny", "humor", "humour", "joke", "fun ",
        "lmao", "lmfao", "rofl", "haha", "hahaha", "hilarious", "comedy",
        # Reddit tropes & slang
        "shitpost", "shitposting", "shit post", "me_irl", "me irl", "meirl",
        "bruh", "silly", "fluff", "cursed", "blursed", "wholesome",
        # Satire & sketches
        "satire", "parody", "comic", "skit", "pun intended", "puns ",
        # Flairs and brackets
        "[fun", "[meme]", "[humor]", "[humour]", "[joke]", "[fluff]", "[wholesome]"
    )

    if any(keyword in combined_text for keyword in fun_keywords):
        return "fun"

    # ------------------------------------------------------------------------
    # 5. Default Category ('general')
    # if other sections don't match, we are marking it as general category
    return "general"


# ----------------------------------------------------------------------------
# JAVASCRIPT SNIPPETS — how we read data from new Reddit's page.
# New Reddit shows posts as <shreddit-post> and comments as <shreddit-comment>.
# Instead of parsing visible text (which changes often), we read each
# element's built-in attributes (permalink, author, score, ...), which is
# more stable. Each snippet below returns plain data back to Python.
# ----------------------------------------------------------------------------
# LISTING_JS: runs on a group page, returns one small dict per visible post.
LISTING_JS = """
return Array.from(document.querySelectorAll('shreddit-post')).map(p => ({
    permalink: p.getAttribute('permalink'),
    author: p.getAttribute('author'),
    score: p.getAttribute('score'),
    comments: p.getAttribute('comment-count'),
    created: p.getAttribute('created-timestamp'),
    format: p.getAttribute('post-type'),
    pinned: p.hasAttribute('is-pinned') || p.hasAttribute('pinned') ||
            (p.textContent || '').includes('Pinned by'),
    locked: p.hasAttribute('is-locked') || p.hasAttribute('locked')
}));
"""

POST_JS = """
# POST_JS: runs on a single post page, returns the full detail of that post
# (title, author, body text, flairs, scores). Returns null if no post found.
const p = document.querySelector('shreddit-post');
if (!p) return null;
const textBody = p.querySelector('[slot="text-body"]');
const flair = p.querySelector('[slot="post-flair"]') || p.querySelector('a[href*="flair_name"]');
return {
    title: p.getAttribute('post-title') || '',
    author: p.getAttribute('author') || '',
    score: p.getAttribute('score'),
    comments: p.getAttribute('comment-count'),
    created: p.getAttribute('created-timestamp'),
    format: p.getAttribute('post-type'),
    body: textBody ? textBody.textContent.trim() : '',
    postFlair: flair ? flair.textContent.trim() : '',
    userFlair: p.getAttribute('author-flair-text') || '',
    pinned: p.hasAttribute('is-pinned') || p.hasAttribute('pinned'),
    locked: p.hasAttribute('is-locked') || p.hasAttribute('locked'),
    attributes: Array.from(p.attributes).map(a => a.name)
};
"""

COMMENTS_JS = """
# COMMENTS_JS: runs on a post page, returns [score, depth] for every loaded
# comment so Python can count them (depth 0 = top-level reply).
return Array.from(document.querySelectorAll('shreddit-comment')).map(c => [
    c.getAttribute('score'),
    c.getAttribute('depth')
]);
"""


# ----------------------------------------------------------------------------
# STEP 1: collect post stubs from a group listing (infinite scroll)
# Scrolls the subreddit page until `limit` unique permalinks are gathered
# or scrolling stops yielding new posts (MAX_LISTING_STALLS).
# ----------------------------------------------------------------------------
def collect_post_stubs(driver, group_url, limit):
    """Visit one group page and collect up to `limit` post links (stubs).

    Args:
        driver: the shared Chrome browser.
        group_url: e.g. https://www.reddit.com/r/gaming/.
        limit: max posts to collect (POSTS_PER_GROUP).

    Returns:
        A list of small dicts with permalink/author/score/comment count,
        used later by scrape_post() to open each post.
    """
    driver.get(to_new_reddit(group_url))
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "shreddit-post"))
        )
    except Exception:
        print("   ! no posts found (private / banned / NSFW or login wall / blocked?)")
        return []

    stubs, seen, stalls = [], set(), 0  # stubs=list, seen=dedup set, stalls=empty-scroll counter

    while len(stubs) < limit and stalls < MAX_LISTING_STALLS:
        added = 0
        # Read all currently loaded <shreddit-post> elements via JS.
        for item in driver.execute_script(LISTING_JS):
            permalink = item.get("permalink")
            if not permalink or permalink in seen:
                continue
            seen.add(permalink)
            stubs.append({
                "permalink": permalink,
                "author": item.get("author") or "[deleted]",
                "score": to_int(item.get("score")),
                "reported_comments": to_int(item.get("comments")),
                "created": item.get("created"),
                "format": item.get("format") or "",
                "pinned": bool(item.get("pinned")),
                "locked": bool(item.get("locked")),
            })
            added += 1

        stalls = 0 if added else stalls + 1  # reset on progress, else count a stall
        print(f"   collected {min(len(stubs), limit)}/{limit} posts")

        # Trigger infinite scroll, then wait for new posts to render.
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(random.uniform(*SCROLL_DELAY))

    return stubs[:limit]


# ----------------------------------------------------------------------------
# STEP 2: scrape one post page
# Opens the post permalink, reads the <shreddit-post> detail + counts
# visible <shreddit-comment> nodes, and builds one flat row for Excel.
# ----------------------------------------------------------------------------
def get_profile(driver, username, cache):
    """Optional: karma + account age through Reddit's public about.json (may be blocked)."""
    if username in cache:
        return cache[username]
    info = {"link_karma": None, "comment_karma": None, "account_created": None}
    if username and username != "[deleted]":
        try:
            driver.get(f"https://www.reddit.com/user/{username}/about.json")
            data = json.loads(driver.find_element(By.TAG_NAME, "body").text)["data"]
            info["link_karma"] = data.get("link_karma")
            info["comment_karma"] = data.get("comment_karma")
            created = data.get("created_utc")
            info["account_created"] = datetime.fromtimestamp(created).isoformat() if created else None
            pause()
        except Exception:
            pass
    cache[username] = info
    return info


def load_more_comments(driver):
    """Scroll down a few times so Reddit lazy-loads more of the comment tree."""
    for _ in range(COMMENT_SCROLLS):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(random.uniform(*SCROLL_DELAY))


def scrape_post(driver, stub, category, group_name, profile_cache):
    """Open one post page and return a flat dict row for the Excel export.

    Args:
        stub: one entry from collect_post_stubs() (has the permalink).
        category: your label, e.g. "Gaming" (becomes the Excel 'category' column).
        group_name: short id like "r_gaming" (becomes the 'group' column).
        profile_cache: dict that remembers author lookups to avoid refetching.

    Returns:
        One dict with identity (post_url, posted_at), author (username,
        flairs), content (title, content, format, type), and engagement
        (scores, comment counts, status) fields.
    """
    # Go to the post page and wait for the web-component post to render.
    driver.get(f"https://www.reddit.com{stub['permalink']}")
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "shreddit-post"))
    )

    # Read post detail (title/body/flair/scores) via JS.
    post = driver.execute_script(POST_JS)
    if not post:
        raise ValueError("post element not found")

    if DEBUG:
        print("   [debug] shreddit-post attributes:", post["attributes"])

    # Prefer live post data, fall back to listing stub when fields are empty.
    author = post["author"] or stub["author"]
    title = post["title"].strip()
    body = post["body"]

    # comments: scroll to lazy-load, then aggregate visible comment nodes.
    load_more_comments(driver)
    comments = driver.execute_script(COMMENTS_JS)
    loaded = len(comments)  # how many comment nodes actually rendered
    top_level = sum(1 for _, depth in comments if to_int(depth, -1) == 0)
    comment_votes = sum(to_int(score) for score, _ in comments)
    reported = to_int(post["comments"], stub["reported_comments"])  # Reddit's total count

    # status flags: pinned/locked/deleted markers for later analysis.
    status = []
    if stub["pinned"] or post["pinned"]:
        status.append("stickied")
    if stub["locked"] or post["locked"]:
        status.append("locked")
    if author == "[deleted]" or body in ("[removed]", "[deleted]"):
        status.append("removed/deleted")

    # One flat row per post — each key below becomes an Excel column header.
    row = {
        # -- identity: which post and when --
        "category": category,  # your own label, e.g. "Gaming"
        "group": group_name,  # short group id, e.g. "r_gaming"
        "post_url": f"https://www.reddit.com{stub['permalink']}",
        "posted_at": post["created"] or stub["created"],  # Reddit's timestamp
        # -- author: who posted, including flair badges --
        "username": author,  # "[deleted]" when the account is gone
        "username_length": len(author),  # handy for username-length analysis
        "user_flair": post["userFlair"],  # author's badge in this group
        "post_flair": post["postFlair"],  # moderator/category tag on the post
        # -- content: what the post says --
        "title": title,
        "title_length": len(title),  # character count of the title
        "content": body,  # text body (empty for image/video/link posts)
        "content_length": len(body),
        "content_format": post["format"] or stub["format"],      # text / image / video / link / gallery
        "post_score": to_int(post["score"], stub["score"]),  # upvotes minus downvotes
        "post_type": classify_post(title, post["postFlair"], body),  # promote/help/question/fun/general
        "post_status": ", ".join(status) if status else "normal",  # normal, stickied, locked, removed/deleted
        # -- engagement: how people reacted --
        "total_comment_votes": comment_votes,  # sum of scores of loaded comments
        "total_comments_with_replies": reported,  # Reddit's official total (incl. replies)
        "total_comments_without_replies": top_level,  # only depth-0 comments we saw
        "comments_loaded": loaded,  # how many comments actually rendered for us
        "more_comments_unloaded": loaded < reported,  # True when Reddit hid some replies
        "scraped_at": datetime.now().isoformat(timespec="seconds"),  # when we scraped it
    }

    # Optional enrichment: author karma/account age (off by default).
    if FETCH_PROFILE:
        row.update(get_profile(driver, author, profile_cache))

    return row


# ----------------------------------------------------------------------------
# MAIN — user input, per-group orchestration, output paths
# ----------------------------------------------------------------------------
def ask_inputs():
    """Prompt for categories and their group links; return {category: [links]}."""
    categories = {}
    n = to_int(input("How many categories? "), 1)
    for i in range(1, n + 1):
        name = input(f"Category {i} name: ").strip() or f"category_{i}"
        links = []
        for j in range(1, LINKS_PER_CATEGORY + 1):
            link = input(f"   [{name}] group link {j}/{LINKS_PER_CATEGORY}: ").strip()
            if link:
                links.append(link)
        categories[name] = links
    return categories


def get_group_name(link):
    """Derive a short group id from a URL, e.g. https://www.reddit.com/r/python/ -> r_python."""
    path = urlparse(to_new_reddit(link)).path
    return path.strip("/").replace("/", "_")


def get_output_path(category, group_name):
    """Build the per-group Excel path: output/<category>__<group>.xlsx."""
    file_name = f"{safe_name(category)}__{safe_name(group_name)}.xlsx"
    return os.path.join(OUTPUT_DIR, file_name)


def save_to_excel(rows, out_path):
    """Write collected row dicts to Excel (no-op when there is nothing yet)."""
    if rows:
        pd.DataFrame(rows).to_excel(out_path, index=False)


def scrape_group(driver, category, link, profile_cache):
    """Scrape all posts of ONE group link and save them to its own Excel file.

    Steps: (1) collect post links, (2) open each post with crash-safe
    autosave every SAVE_EVERY posts, (3) write the final Excel file.
    """
    group_name = get_group_name(link)
    out_path = get_output_path(category, group_name)
    print(f"\n=== {category} | {link} ===")

    # Step 1: get the list of posts
    posts = collect_post_stubs(driver, link, POSTS_PER_GROUP)
    total = len(posts)

    # Step 2: open every post and extract its data
    rows = []
    for number, post in enumerate(posts, start=1):
        short_url = post["permalink"][:70]

        try:
            row = scrape_post(driver, post, category, group_name, profile_cache)
            rows.append(row)
            print(f"   [{number}/{total}] ok      {short_url}")
        except Exception as error:
            print(f"   [{number}/{total}] FAILED  {short_url}  ({type(error).__name__})")

        # Autosave every few posts so a crash does not lose everything
        if number % SAVE_EVERY == 0:
            save_to_excel(rows, out_path)

        pause()

    # Step 3: final save
    if rows:
        save_to_excel(rows, out_path)
        print(f"   saved {len(rows)} posts -> {out_path}")
    else:
        print("   nothing scraped for this group")


def main():
    """Entry point: prompt for input, open one browser, scrape every group link."""
    categories = ask_inputs()
    os.makedirs(OUTPUT_DIR, exist_ok=True)  # create output/ if missing

    driver = make_driver()  # single browser reused for the whole run
    profile_cache = {}  # memoize author lookups so each user is fetched once

    try:
        for category, links in categories.items():
            for link in links:
                scrape_group(driver, category, link, profile_cache)
    finally:
        driver.quit()   # always close the browser, even after an error


if __name__ == "__main__":
    main()