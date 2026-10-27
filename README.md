# SwimSwam Workout Index

Builds a formatted Excel index of SwimSwam's Daily Swim Coach Workout posts
(#100–#1136) with a clickable link to each one.

## What it collects

Each row contains: **Workout #, Title, URL, Purpose, Target age group, Target level, Weeks until target meet, Course**

The main sheet is filtered to **Course: 25 Yards**; an **All rows** sheet keeps
everything with a Status column (ok / not found / blocked / no metadata found).

## Index files

- [`INDEX.md`](INDEX.md): browsable table with a link to every workout (renders on GitHub)
- [`data/workout_index.csv`](data/workout_index.csv): the same index as CSV

The committed copies list the links only; Purpose, age group, level, weeks and Course
fill in when the scraper runs. Each run also uploads fresh copies as artifacts.

## How to run

### Option A: Run manually on GitHub (recommended)

1. Go to the **Actions** tab in this repo
2. Click **SwimSwam Workout Index** in the left sidebar
3. Click **Run workflow**
4. Set the start/end numbers and course filter (defaults: 100, 1136, 25 Yards)
5. Click the green **Run workflow** button
6. When it finishes (about 1.5–2 hours for the full range), download the Excel file from the run's **Artifacts** section

### Option B: Run locally

```
pip install -r requirements.txt
CONTACT="you@example.com" START=100 END=1136 python scraper.py
```

### Option C: Pages saved from your browser (if the site blocks the scraper)

Save workout pages (Ctrl/Cmd+S, "Webpage, HTML only") into a folder, with the
workout number at the end of the filename (e.g. `daily-swim-coach-workout-1136.html`):

```
LOCAL_DIR=./saved_pages python scraper.py
```

## Configuration

- `START` / `END`: workout number range (default 100–1136)
- `COURSE_FILTER`: text the Course field must contain for the main sheet (default `25 Yards`; blank keeps all)
- `DELAY_MIN` / `DELAY_MAX`: seconds between requests (default 4–8)
- `CONTACT`: your email or repo URL, included in the User-Agent
- `LOCAL_DIR`: parse saved HTML files instead of fetching
- `OUTPUT`: output filename (default `swimswam_workouts.xlsx`)

## Notes

- For personal use. Only links and metadata are collected; workout text is not saved.
- The scraper obeys robots.txt, uses an honest User-Agent, and **stops after 3
  consecutive 403 responses** instead of trying to get around bot detection.
  SwimSwam may block GitHub's servers; if so, use Option C or ask SwimSwam for access.
- The page layout hasn't been verified. The parser looks for the five labels
  generically ("Purpose: X", or the label on one line and the value on the next).
  If rows come back "no metadata found", adjust `LABELS` in `scraper.py`.
- GitHub Actions gives 2,000 free minutes/month on private repos; a full run uses roughly 100.
