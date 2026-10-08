# Reddit Community Behavior

How do post length, description, and post type behave across Reddit communities?
This project scrapes Reddit groups, then cleans, analyzes, and visualizes the
answers in a Streamlit dashboard.

## Goal

Compare communities (e.g. Gaming vs. Memes, or any categories you choose) on:

- **Post length** — do longer titles/bodies get more score or comments?
- **Post description** — how do flairs, text vs. media, and status
  (stickied/locked/removed) relate to engagement?
- **Post type** — how do text / image / video / link / gallery posts and the
  keyword label (promote / help / question / fun / general) behave per community?

## Project plan

| Phase | Status | What happens |
|-------|--------|--------------|
| 1. Scrape | ✅ Done (current) | `scrap.py` collects posts via Selenium → one Excel file per group in `output/` |
| 2. Clean | 🔲 Next | Normalize timestamps, dedup posts, handle deleted authors/bodies, standardize types |
| 3. Analyze | 🔲 Planned | Length-vs-engagement, type mix per community, flair/status effects, posting-time patterns |
| 4. Dashboard | 🔲 Planned | Streamlit UI: filters by community, charts, post explorer, downloadable tables |

## How scraping works (Phase 1)

1. You enter categories and group links (3 links per category by default).
2. For each link, the scraper scrolls the group page and collects post stubs
   from `<shreddit-post>` elements (permalink, author, score, comment count).
3. It opens every post page, reads title/body/flairs/scores, scrolls to load
   comments, and counts visible `<shreddit-comment>` nodes.
4. Rows are autosaved every `SAVE_EVERY` posts and finally written to
   `output/<category>__<group>.xlsx` — one file per group, so a crash never
   loses everything.

New Reddit renders content as web components, so the scraper reads element
attributes via JavaScript instead of parsing fragile HTML text.

## Output schema (one row = one post)

- Identity: `category`, `group`, `post_url`, `posted_at`, `scraped_at`
- Author: `username`, `username_length`, `user_flair`, (+ `link_karma`,
  `comment_karma`, `account_created` when `FETCH_PROFILE=True`)
- Content: `title`, `title_length`, `content`, `content_length`,
  `content_format` (text/image/video/link/gallery), `post_flair`,
  `post_type` (promote/help/question/fun/general), `post_status`
- Engagement: `post_score`, `total_comment_votes`,
  `total_comments_with_replies`, `total_comments_without_replies`,
  `comments_loaded`, `more_comments_unloaded`

## Setup

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

Requires Chrome + a matching chromedriver (Selenium finds it automatically
if it's on your PATH, or set `CHROMEDRIVER_PATH` in `scrap.py`).

## Usage

```bash
source .venv/bin/activate
python scrap.py
```

Example session:

```text
How many categories? 1
Category 1 name: Gaming
   [Gaming] group link 1/3: https://www.reddit.com/r/gaming/
```

Key settings at the top of `scrap.py`:

- `POSTS_PER_GROUP = 10` — posts scraped per group link
- `HEADLESS = False` — keep False, new Reddit blocks headless browsers more often
- `FETCH_PROFILE = False` — author karma/age lookup (slower, may be blocked)
- `SAVE_EVERY = 10` — autosave interval

## Limitations

- Only publicly visible posts; private/banned/NSFW/login-walled groups return nothing.
- Comment counts reflect what rendered after `COMMENT_SCROLLS` scrolls —
  `more_comments_unloaded` flags hidden replies.
- `post_type` is a rough keyword guess, not ground truth.
- Be polite: random delays are built in; don't crank them to zero.

## Roadmap details

- **Clean:** unify `posted_at` timestamps, drop exact-duplicate posts, keep
  deleted/removed markers as nulls (never invent data).
- **Analyze:** title/content length vs. score/comments, type mix and virality
  per community, flair and status effects, hour/weekday patterns.
- **Dashboard:** Streamlit pages for overview, length analysis, type analysis,
  and a searchable post table with Excel/CSV downloads.
