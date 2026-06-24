# Resume Generator

A local web app that generates an **ATS-optimized resume tailored to any job description**
using the OpenAI API, rendered as a styled PDF (Bitter + Roboto, the `resume_format.py` look).

## How it works

You keep a **fixed set of facts** once: your name + contact, your companies (name,
location, start/end dates), and your university (name, location, dates). For each job you
paste the job description, and the AI writes everything else tailored to that job:

- A clean job-title **SUBTITLE** (internal level codes / team names stripped).
- A seniority-matched **Summary** with technical keywords highlighted.
- 4–6 **Skills** categories chosen to fit the job.
- For each company: a JD-appropriate **title**, a one-line **description**, and quantified
  **achievement bullets** (10 / 8 / 6 minimum, newest company first).
- A **degree** name that fits the job.

Your company names, dates, locations, and school are always kept **verbatim** — only the
prose is generated. The generation prompt lives in `app.py` (`PROMPT_TEMPLATE`).

> Note: this follows an aggressive ATS-optimization prompt that *invents* tailored
> achievements and metrics for each role. Review the output before sending, and regenerate
> if any bullet slips in a percentage (the prompt forbids them but the model occasionally
> adds one).

## Setup

1. Put your OpenAI API key in `.env` (copy from `.env.example`):
   ```
   OPENAI_API_KEY=sk-...your key...
   OPENAI_MODEL=gpt-4o
   ```
   ⚠️ If a key was ever shared, revoke it at https://platform.openai.com/api-keys and use a new one.

2. Dependencies (already installed for your user; reinstall if needed):
   ```
   python3 -m pip install --user -r requirements.txt
   ```
   Fonts **Bitter** and **Roboto** must be installed for the PDF to match (both are set up
   in `~/.local/share/fonts`).

## Run

```
python3 app.py
```
Open http://127.0.0.1:5000

1. **My Profile** → fill in your name/contact (your companies + school are pre-filled).
   Save.
2. **Generate** → paste a job description → **Generate tailored resume**.
3. Review the preview → **Download PDF**.

## Files

- `app.py` — Flask app, the generation prompt, plain-text parser, fact overlay, PDF export
- `profile.json` — your fixed facts (git-ignored)
- `templates/resume.html` — the Bitter/Roboto resume layout
- `templates/profile.html` — profile editor
- `static/style.css` — UI styling
