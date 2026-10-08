"""Reddit scraper (Selenium + new Reddit UI).

Flow: ask for categories/links -> collect post stubs from each group listing
(infinite scroll on <shreddit-post>) -> open every post -> extract title,
body, scores, comment counts -> save one Excel file per group in output/.

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
    """Make a filesystem-safe name for the Excel file (alphanumeric + _-)."""
    return re.sub(r"[^A-Za-z0-9_-]+", "_", text).strip("_")[:60] or "group"


def to_int(value, default=0):
    """Parse Reddit count strings like '1,234' safely; fall back to default."""
    try:
        return int(str(value).replace(",", "").strip())
    except (ValueError, TypeError):
        return default


def classify_post(title, flair, body, promoted=False):
    """Rough keyword-based type: promote / help / question / fun / general."""
    text = f"{title} {flair}".lower()
    if promoted or any(k in text for k in (
            "my project", "i made", "i built", "check out", "launch", "giveaway",
            "discount", "promo", "subscribe", "my channel", "my app", "my website",
            "[promo", "self promotion", "self-promotion", "advert")):
        return "promote"
    if any(k in text for k in ("help", "need ", "issue", "error", "problem", "stuck",
                               "advice", "support", "can't", "cannot", "not working")):
        return "help"
    t = title.strip().lower()
    if t.endswith("?") or t.startswith(("how ", "what ", "why ", "when ", "where ",
                                         "which ", "who ", "is ", "are ", "can ",
                                         "does ", "do ", "should ")) or "question" in text:
        return "question"
    if any(k in text for k in ("meme", "lol", "funny", "humor", "humour", "joke",
                               "fun ", "shitpost", "[fun")):
        return "fun"
    return "general"


# ----------------------------------------------------------------------------
# JAVASCRIPT SNIPPETS (read the data from the <shreddit-...> elements)
# New Reddit renders posts/comments as web components; we read attributes
# directly instead of parsing fragile HTML/text.
# ----------------------------------------------------------------------------
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
    """Visit a group listing and return up to `limit` lightweight post stubs."""
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
    """Open one post page and return a flat dict row for the Excel export."""
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

    # One flat row per post — column names become the Excel headers.
    row = {
        "category": category,
        "group": group_name,
        "post_url": f"https://www.reddit.com{stub['permalink']}",
        "posted_at": post["created"] or stub["created"],
        "username": author,
        "username_length": len(author),
        "user_flair": post["userFlair"],
        "post_flair": post["postFlair"],
        "title": title,
        "title_length": len(title),
        "content": body,
        "content_length": len(body),
        "content_format": post["format"] or stub["format"],      # text / image / video / link / gallery
        "post_score": to_int(post["score"], stub["score"]),
        "post_type": classify_post(title, post["postFlair"], body),
        "post_status": ", ".join(status) if status else "normal",
        "total_comment_votes": comment_votes,
        "total_comments_with_replies": reported,
        "total_comments_without_replies": top_level,
        "comments_loaded": loaded,
        "more_comments_unloaded": loaded < reported,
        "scraped_at": datetime.now().isoformat(timespec="seconds"),
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
    """Scrape all posts of ONE group link and save them to its own Excel file."""
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