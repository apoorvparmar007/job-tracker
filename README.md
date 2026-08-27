# Germany Job Tracker (Data Scientist / AI Engineer)

Automated LinkedIn job tracker for Apoorv Parmar, scored against his resume
profile using Gemini. Runs on a GitHub Actions schedule, fetches new postings
via Apify, excludes known recruiting-mill / non-fit sources, dedupes against
previously seen jobs, and keeps a running spreadsheet
(`data/job_matches.xlsx`).

## What it does every run

1. Searches LinkedIn (via Apify's `curious_coder/linkedin-jobs-scraper`) for
   "Data Scientist" and "AI Engineer" roles in Germany posted in the last 24
   hours.
2. Drops postings from known recruiting-mill sources (e.g. "Jobs Ai", "Jack &
   Jill") and non-fit roles (PhD-only academic posts, internships/Werkstudent)
   *before* spending any Gemini or Apify credit on them.
3. Scores every remaining job 0-100 using **Gemini** (`gemini-flash-latest`)
   against Apoorv's resume profile — weighing GenAI/RAG/LangChain/LangGraph/
   Azure OpenAI heavily as his core specialization, classical ML as
   secondary, and adjusting for seniority mismatch. Falls back to a simple
   keyword scorer if Gemini is unavailable, so a single API hiccup doesn't
   stop the run (those rows are clearly labeled in the Reason column).
4. Detects German-language requirement and sponsorship signals directly from
   the job description text (pattern-based, not Gemini — these are simple
   enough that an LLM call isn't needed).
5. Compares against `data/seen_jobs.json` (job IDs already tracked) and adds
   **only genuinely new postings** as new rows.
6. Re-checks a bounded number of previously tracked, still-"Open" listings
   each run (plain HTTP GET, no extra Apify cost) and marks any that now show
   a "no longer accepting applications" marker as Closed — greyed out in the
   sheet, not deleted.
7. Commits the updated spreadsheet and seen-jobs list back to the repo.

## One-time setup

### 1. Create the GitHub repo

Push this folder's contents to a new **private** GitHub repository (private
recommended — the spreadsheet contains your personal job-search data).

```bash
cd job-tracker
git init
git add .
git commit -m "Initial commit: Germany job tracker"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

### 2. Add your API tokens as repo secrets

You need two secrets:

1. **Apify token** — from the Apify Console: **Settings → Integrations → API
   token**.
2. **Gemini API key** — from [Google AI Studio](https://aistudio.google.com/apikey).

In your GitHub repo: **Settings → Secrets and variables → Actions → New
repository secret**. Add both:

- Name: `APIFY_TOKEN`, Value: your Apify token
- Name: `GEMINI_API_KEY`, Value: your Gemini API key

No other secrets are needed — the workflow uses the repo's own
`GITHUB_TOKEN` (automatically provided) to commit the updated spreadsheet
back.

**If `GEMINI_API_KEY` is missing or the Gemini call fails for any reason**
(network issue, quota exhausted, malformed response), the script
automatically falls back to a simple keyword-based scorer rather than
crashing the run — those rows are clearly labeled
`[FALLBACK SCORER - Gemini call failed]` in the Reason column so you know to
treat that particular score as lower-confidence.

### 3. Confirm Actions permissions allow the bot to push

**Settings → Actions → General → Workflow permissions** → select **"Read and
write permissions"**. Without this, the workflow can run and score jobs but
will fail to commit the updated spreadsheet back to the repo.

### 4. (Optional) Change the schedule

The workflow is set to run every 6 hours (`0 */6 * * *`). To run every 3
hours instead, edit `.github/workflows/job-tracker.yml`:

```yaml
- cron: '0 */3 * * *'
```

GitHub Actions cron times are UTC and best-effort (can be delayed a few
minutes under load) — this is a GitHub platform limitation, not something
this code controls.

### 5. Run it once manually to confirm it works

Go to the **Actions** tab in your repo → **Germany Job Tracker** workflow →
**Run workflow** (this uses the `workflow_dispatch` trigger already in the
YAML). Check the run logs, then download `data/job_matches.xlsx` from the
repo to inspect it.

## How the Gemini scoring works

Each job's title, company, and description are sent to Gemini alongside a
fixed description of Apoorv's resume (`RESUME_PROFILE` in `tracker.py`) and a
scoring rubric (`SCORING_INSTRUCTIONS`) that tells it how to weigh GenAI/LLM
specialization vs. classical ML vs. seniority mismatch. The request uses a
JSON schema (`SCORE_RESPONSE_SCHEMA`) so Gemini's response is always a
well-formed `{score, reason}` object — no fragile text parsing.

This replaced an earlier pure keyword-matching approach, which under- and
over-scored several edge cases (e.g. treating "some exposure to LLMs is a
plus" as nearly as strong a match as a genuine GenAI-specialist role). An LLM
judge handles this kind of nuance far better than fixed keyword weights.

To adjust scoring behavior over time, edit in `tracker.py`:
- `RESUME_PROFILE` — update as Apoorv's resume or focus areas evolve.
- `SCORING_INSTRUCTIONS` — the rubric Gemini is given. Tighten or loosen this
  if you notice scores drifting too generous or too harsh.

The model used is `gemini-flash-latest` — a stable alias that always points
to Google's current recommended Flash-tier model, chosen deliberately over a
pinned version number (like `gemini-2.5-flash`) so this script doesn't break
when Google deprecates a specific dated model down the line.

## Updating the exclusion rules

Exclusion is deliberately kept as fast, deterministic pre-filtering in
`tracker.py` — no reason to spend a Gemini call scoring a posting from a
known recruiting mill:

- `EXCLUDED_COMPANY_SUBSTRINGS` — add a company name substring here to
  permanently exclude it from future runs (e.g. "jobs ai", "jack & jill").
- `EXCLUDED_TITLE_PATTERNS` — regex patterns matched against the job title
  (PhD-only posts, Werkstudent/internship roles, etc.).

After editing, commit and push — the next scheduled (or manually triggered)
run uses the new rules.

## Limitations (read before trusting this unattended)

- **Gemini scoring is much better than keyword matching, but still not
  infallible.** It can misjudge an unusually phrased or very short job
  description. Treat the score as a strong triage signal, not gospel — skim
  the `Reason` column before ruling something out.
- **The fallback scorer is deliberately crude.** If you see
  `[FALLBACK SCORER]` in the Reason column, that score was NOT produced by
  Gemini (API key missing, quota exhausted, or a request failure) — verify
  that job manually rather than trusting the number.
- **"Still Open" checking depends on LinkedIn's page markup staying stable.**
  It looks for phrases like "no longer accepting applications" in the raw
  page HTML. If LinkedIn changes this wording or markup, the check will
  silently stop catching closures — worth spot-checking every few weeks
  against a listing you know has closed.
- **German-requirement and sponsorship detection are pattern-based**, not
  guaranteed accurate. Always verify on the actual job posting before ruling
  a role in or out.
- **Both Apify and Gemini usage cost money** beyond any free tier — running
  every 3-6 hours will use meaningfully more quota than the one-off runs from
  before. Check your plan limits on both before deploying a tight schedule.
