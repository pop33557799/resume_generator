"""
Resume Generator — produces an ATS-optimized resume tailored to a job description.

Model (per the reference prompt):
  - You keep a FIXED set of facts: name + contact, your companies (name, location,
    start/end dates), and your university (name, location, dates).
  - You paste a job description.
  - The OpenAI API writes everything else, tailored to that job: a clean job-title
    SUBTITLE, a seniority-matched Summary, 4-6 Skills categories, and for each company
    a JD-appropriate TITLE, a one-line DESCRIPTION, and quantified achievement bullets.
    It also picks a DEGREE name that fits the job.
  - Dates, locations, and the school are taken VERBATIM from your profile. Company
    names can come from the profile or from AI-chosen real, U.S.-based employers.
  - Output is rendered to the Bitter/Roboto PDF format defined in resume_format.py.
"""

import json
import math
import os
import re
import html
import tempfile
import time
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook

from dotenv import load_dotenv
from flask import (
    Flask, abort, flash, redirect, render_template, request,
    send_from_directory, session, url_for,
)
from markupsafe import Markup
from openai import OpenAI, RateLimitError
from weasyprint import HTML

load_dotenv()

BASE_DIR = Path(__file__).parent
LEGACY_PROFILE_PATH = BASE_DIR / "profile.json"   # single-profile file (auto-migrated)
PROFILES_DIR = BASE_DIR / "profiles"              # one <id>.json per profile
ACTIVE_PROFILE_FILE = BASE_DIR / ".active_profile"
RESUMES_DIR = BASE_DIR / "resumes"
LEGACY_RESUMES_LOG = BASE_DIR / "resumes_log.xlsx"   # pre-multi-profile shared log
LOG_HEADERS = ["Date", "Job Title", "URL", "Company", "Filename", "Progress", "Note"]


def resumes_log_path(pid: str) -> Path:
    """Each profile keeps its own log file: resumes_log_<id>.xlsx."""
    return BASE_DIR / f"resumes_log_{pid}.xlsx"


def _migrate_legacy_log() -> None:
    """One-time: assign the old shared resumes_log.xlsx to the first profile."""
    if not LEGACY_RESUMES_LOG.exists():
        return
    pid = get_active_id()
    if pid and not resumes_log_path(pid).exists():
        os.replace(LEGACY_RESUMES_LOG, resumes_log_path(pid))
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o")
NVIDIA_MODEL = os.getenv("NVIDIA_MODEL", "deepseek-ai/deepseek-v4.1-flash")
NVIDIA_BASE_URL = "https://integrate.api.nvidia.com/v1"
# NVIDIA-hosted models often default to a short output cap (~1024 tokens),
# which would cut the resume off; ask for enough room explicitly. DeepSeek's
# thinking plus a resume measured up to ~6,200 tokens, so leave headroom.
NVIDIA_MAX_TOKENS = 12000
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "openai/gpt-4o")
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
# OpenRouter refuses a request unless the balance covers max_tokens up front, so
# keep the cap close to a full resume (GPT-4o writes one in ~1,100 tokens)
# instead of the model's maximum (16,384) — a low-credit key can't afford that.
OPENROUTER_MAX_TOKENS = 2000
AI_PROVIDERS = ("openai", "nvidia", "openrouter")
# Models offered in the Generate page's "Model" dropdown, per provider. Each
# provider's model from .env is the default and is added to its list if missing.
MODEL_CHOICES = {
    "openai": ["gpt-4o", "gpt-4o-mini", "gpt-6-luna", "gpt-5.6-terra", "gpt-5.6-luna"],
    "nvidia": ["deepseek-ai/deepseek-v4.1-flash", "nvidia/nemotron-3.5-lightning-30b-a3b",
               "moonshotai/kimi-k3"],
    "openrouter": ["openai/gpt-4o", "openai/gpt-4o-mini", "openai/gpt-6-luna",
                   "openai/gpt-5.6-terra", "openai/gpt-5.6-luna",
                   "deepseek/deepseek-v4.1-flash", "anthropic/claude-sonnet-5.5"],
}
DEFAULT_MODELS = {"openai": MODEL, "nvidia": NVIDIA_MODEL, "openrouter": OPENROUTER_MODEL}
for _provider, _default in DEFAULT_MODELS.items():
    if _default not in MODEL_CHOICES[_provider]:
        MODEL_CHOICES[_provider].insert(0, _default)


def chosen_model(provider: str, model: str) -> str:
    """The requested model if it's offered for this provider, else its default."""
    return model if model in MODEL_CHOICES[provider] else DEFAULT_MODELS[provider]

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", "dev-secret-change-me")

_MONTHS = ["", "January", "February", "March", "April", "May", "June",
           "July", "August", "September", "October", "November", "December"]


# ---------------------------------------------------------------------------
# Profile persistence
# ---------------------------------------------------------------------------

def empty_profile() -> dict:
    return {"name": "", "industry": "", "email": "", "phone": "", "location": "",
            "links": [], "companies": [], "education": []}


def _profile_path(pid: str) -> Path:
    return PROFILES_DIR / f"{pid}.json"


def _unique_pid(name: str) -> str:
    base = _slug(name).lower() or "profile"
    pid, n = base, 2
    while _profile_path(pid).exists():
        pid, n = f"{base}_{n}", n + 1
    return pid


def _migrate_legacy_profile() -> None:
    """One-time: move the old single profile.json into profiles/ on first run."""
    PROFILES_DIR.mkdir(exist_ok=True)
    if any(PROFILES_DIR.glob("*.json")):
        return
    if LEGACY_PROFILE_PATH.exists():
        data = json.loads(LEGACY_PROFILE_PATH.read_text(encoding="utf-8"))
        pid = _unique_pid(data.get("name") or "profile")
        save_profile(data, pid)
        set_active_id(pid)
        LEGACY_PROFILE_PATH.rename(BASE_DIR / "profile.migrated.json")


def list_profiles() -> list:
    """[{id, name}] for every saved profile, sorted by name."""
    _migrate_legacy_profile()
    out = []
    for p in PROFILES_DIR.glob("*.json"):
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        industry = (data.get("industry") or "").strip()
        name = data.get("name") or "(unnamed)"
        label = f"{name} - {industry}" if industry else name
        out.append({"id": p.stem, "name": name, "industry": industry, "label": label})
    out.sort(key=lambda x: x["label"].lower())
    return out


def list_profile_names(profiles: list | None = None) -> list:
    """Unique users, each pointing at that user's first saved profile."""
    profiles = profiles or list_profiles()
    names, seen = [], set()
    for p in profiles:
        if p["name"] in seen:
            continue
        seen.add(p["name"])
        names.append({"name": p["name"], "id": p["id"]})
    return names


def profiles_for_name(name: str, profiles: list | None = None) -> list:
    """All saved profiles for one user/name. Kept for old per-industry logs."""
    profiles = profiles or list_profiles()
    return [p for p in profiles if p["name"] == name]


def get_active_id() -> str | None:
    if ACTIVE_PROFILE_FILE.exists():
        pid = ACTIVE_PROFILE_FILE.read_text(encoding="utf-8").strip()
        if pid and _profile_path(pid).exists():
            return pid
    profiles = list_profiles()
    return profiles[0]["id"] if profiles else None


def set_active_id(pid: str | None) -> None:
    ACTIVE_PROFILE_FILE.write_text(pid or "", encoding="utf-8")


def load_profile(pid: str | None = None) -> dict:
    """Load a profile by id; None = the active profile. Empty if none exist."""
    _migrate_legacy_profile()
    if pid is None:
        pid = get_active_id()
    if pid and _profile_path(pid).exists():
        data = json.loads(_profile_path(pid).read_text(encoding="utf-8"))
        return {**empty_profile(), **data}
    return empty_profile()


def save_profile(profile: dict, pid: str) -> None:
    PROFILES_DIR.mkdir(exist_ok=True)
    _profile_path(pid).write_text(
        json.dumps(profile, indent=2, ensure_ascii=False), encoding="utf-8")


def create_profile(name: str) -> str:
    pid = _unique_pid(name or "profile")
    save_profile({**empty_profile(), "name": name, "industry": ""}, pid)
    return pid


def delete_profile(pid: str) -> None:
    p = _profile_path(pid)
    if p.exists():
        p.unlink()


def parse_profile_form(form) -> dict:
    profile = empty_profile()
    for key in ("name", "email", "phone", "location"):
        profile[key] = form.get(key, "").strip()

    for label, url in zip(form.getlist("link_label"), form.getlist("link_url")):
        if label.strip() or url.strip():
            profile["links"].append({"label": label.strip(), "url": url.strip()})

    names = form.getlist("comp_name")
    locs = form.getlist("comp_location")
    starts = form.getlist("comp_start")
    ends = form.getlist("comp_end")
    n_companies = max(len(names), len(locs), len(starts), len(ends))
    for i in range(n_companies):
        name = names[i].strip() if i < len(names) else ""
        location = locs[i].strip() if i < len(locs) else ""
        start = starts[i].strip() if i < len(starts) else ""
        end = ends[i].strip() if i < len(ends) else ""
        if not (name or location or start or end):
            continue
        profile["companies"].append({
            "name": name or f"Experience {len(profile['companies']) + 1}",
            "location": location,
            "start": start,
            "end": end,
        })

    unis = form.getlist("edu_university")
    elocs = form.getlist("edu_location")
    estarts = form.getlist("edu_start")
    eends = form.getlist("edu_end")
    for i in range(len(unis)):
        if not unis[i].strip():
            continue
        profile["education"].append({
            "university": unis[i].strip(),
            "location": elocs[i].strip() if i < len(elocs) else "",
            "start": estarts[i].strip() if i < len(estarts) else "",
            "end": eends[i].strip() if i < len(eends) else "",
        })
    return profile


# ---------------------------------------------------------------------------
# Date / keyword formatting
# ---------------------------------------------------------------------------

def fmt_date(value: str, is_end: bool = False) -> str:
    """MM/YYYY -> 'Month YYYY'; YYYY stays; blank/Present -> 'Present' for an end."""
    value = (value or "").strip()
    if not value or value.lower() == "present":
        return "Present" if is_end else value
    m = re.match(r"^(\d{1,2})[/-](\d{4})$", value)
    if m:
        month = int(m.group(1))
        if 1 <= month <= 12:
            return f"{_MONTHS[month]} {m.group(2)}"
    return value


def fmt_range(start: str, end: str) -> str:
    s, e = fmt_date(start), fmt_date(end, is_end=True)
    return f"{s} – {e}" if s and e else (s or e)


# The prompt forbids percentages, but the model occasionally emits a delta like
# "by 20%". Strip that clause and tidy the result (keeps endorsed figures such as
# "99.9% uptime", which aren't deltas).
_PERCENT_RE = re.compile(r"\s*\bby\s+\d+(?:\.\d+)?\s*(?:%|percent\b)", re.IGNORECASE)


def strip_percentages(text: str) -> str:
    if not text:
        return text
    cleaned = _PERCENT_RE.sub("", text)
    cleaned = re.sub(r"\s+([.,;])", r"\1", cleaned)   # " ." -> "."
    cleaned = re.sub(r"\s{2,}", " ", cleaned)
    return cleaned.strip()


def build_warnings(tailored: dict) -> list:
    """Flag likely parse/format problems so the user can regenerate instead of
    silently shipping a sparse resume."""
    warnings = []
    if not tailored["summary"]:
        warnings.append("The Summary came back empty — the model output may be malformed.")
    if not tailored["skills"]:
        warnings.append("No skill categories were parsed.")
    minimums = [10, 8, 6] + [4] * 12
    for i, e in enumerate(tailored["experience"]):
        n = len(e["bullets"])
        if n == 0:
            warnings.append(f"{e['company']}: no bullet points were parsed.")
        elif i < len(minimums) and n < minimums[i]:
            warnings.append(f"{e['company']}: only {n} bullets (expected at least {minimums[i]}).")
    return warnings


def render_keywords(text: str) -> Markup:
    """Escape HTML, then turn <<keyword>> markup into bold spans."""
    if not text:
        return Markup("")
    escaped = html.escape(text)
    escaped = re.sub(r"&lt;&lt;(.+?)&gt;&gt;", r'<strong class="kw">\1</strong>', escaped)
    return Markup(escaped)


app.jinja_env.filters["kw"] = render_keywords


# ---------------------------------------------------------------------------
# Prompt (verbatim from the reference generator; {jobDesc}/{companiesText}/
# {educationText} placeholders are filled in build_prompt()).
# ---------------------------------------------------------------------------

PROMPT_TEMPLATE = r"""
Please make resume with perfect match with "__JOBDESC__".

Target/hiring company to avoid as an experience employer: __TARGETCOMPANY__.
Experience company-name source selected by user: __COMPANYSOURCE__.

CRITICAL — Job title and seniority matching:
- Read the job description carefully. Identify the core job role and seniority level.
- The FIRST line of your output must be: SUBTITLE: <a realistic, clean job title based on the job description>
- IMPORTANT: The SUBTITLE and TITLE fields must be realistic job titles that a real company would use. Strip out any internal level codes (like "L5", "L6", "IC4", "P3"), team/department names (like "- Messaging", "- Platform", "- Growth"), and company-specific jargon. For example:
  * JD says "Data Engineer (L5) - Messaging" -> use "Senior Data Engineer"
  * JD says "Software Engineer II - Payments" -> use "Software Engineer"
  * JD says "Staff SWE, Infrastructure" -> use "Staff Software Engineer"
  * JD says "Frontend Engineer (L4) - Growth" -> use "Frontend Engineer"
  Keep only the standard industry job title that would appear on LinkedIn or a real offer letter.
- The Summary must match the seniority level. If the job says "Junior", do NOT say "9 years of experience as a senior engineer". Instead say something appropriate like "A motivated developer with hands-on experience in..." matching the tone and level of the job description.
- For TITLE in each company: make each job title closely related to the job description role. Use variations that show career progression toward that role. For example if the job is for a Senior Data Engineer, use titles like "Senior Data Engineer", "Data Engineer", "Junior Data Engineer", "Data Analyst" — NOT "Senior Architect" or "Lead Platform Engineer". The most recent company title should be the closest match to the target role. NEVER include level codes, team names, or internal designations in any TITLE field.

Using the style of professional resume (clear headers, concise summary, bullet-pointed achievements with quantified results), generate a complete resume for this candidate.
Format requirements:
- Clear section headers: Summary, Skills, Professional Experience, Education.
- Do NOT include a Certifications section at all.
- Skills: Organize into 4-6 categories that make sense for THIS specific job. Choose category names that logically group the skills required in the job description. Each category label must accurately describe its contents — do NOT force unrelated skills into a mismatched category.
  Rules for skill categories:
  1. Read the job description and identify the major skill areas it requires.
  2. Create category names that genuinely fit those areas (e.g., for a Jira consultant role use "Project Management Tools", "Scripting & Automation", "Data Management" — NOT "Cloud, DevOps & Tools").
  3. Every skill listed under a category must logically belong there. "Microsoft Office Suite" is NOT a Cloud/DevOps tool. "API integrations" is NOT a Cloud/DevOps tool.
  4. Only list actual technologies, tools, platforms, languages, and frameworks — not vague concepts like "Data structures" or "Information Technology".
  5. Include all relevant technologies mentioned in the job description.
  Format: one category per line as "Category Name: skill1, skill2, skill3"
Do NOT use markdown, asterisks, or special characters like ** or ##. Use plain text only.
Do NOT include placeholder lines like [Candidate's Name] or [Phone Number].
Remove all blank lines between skill categories.

Writing style:
- Use clear, professional language with easy-to-read sentences.
- Avoid using asterisks (*) or markdown symbols; use plain text.
- Focus on achievements, impact, and technical depth to attract client's focus.
- SPELLING & GRAMMAR: Your output must be 100% free of spelling errors, grammatical mistakes, and typos. Double-check every word before finalizing. Use correct verb tenses (past tense for previous roles, present tense only for current role). Ensure subject-verb agreement, correct prepositions, and proper punctuation throughout. Common mistakes to avoid: "it's" vs "its", "effect" vs "affect", "insure" vs "ensure", run-on sentences, comma splices, and sentence fragments.
- IMPACT-FIRST writing: Every bullet and every sentence in the Summary must demonstrate measurable impact. Do NOT describe what you were "responsible for" — describe what you ACHIEVED and the result. Recruiters scan for impact numbers within the first 3 words after the verb. Structure: [Verb] [thing] -> [result with number]. Example: "Reduced API response time from 4s to 200ms by refactoring query logic and adding <<Redis>> caching layer, serving 15K daily active users."
- Do NOT use any percentage numbers like "30 percent", "20 percent", "by 40%", "by 25%". Senior engineers describe impact with absolute numbers instead (e.g., "from 12s to 800ms", "processing 3M records daily", "across 40 microservices", "serving 10K concurrent users"). Use counts, durations, throughputs, or scale — never percentages.
- Every bullet MUST mention specific technologies, tools, or frameworks by name.
- Every bullet MUST include a measurable outcome with absolute numbers (not percentages).
- Write like a senior engineer describing real project work, not generic responsibilities.
- IMPORTANT: In the Summary section AND in bullet points, wrap every technical keyword (technology name, framework, tool, protocol, service name, language) inside double angle brackets like <<Kafka>>, <<React>>, <<PostgreSQL>>, <<Kubernetes>>, <<REST API>>, <<CI/CD>>. This is used for formatting only. Example: "Built a streaming pipeline with <<Apache Kafka>> and <<Spark Structured Streaming>>, processing 3M events daily."
- Do NOT use <<>> in the SUBTITLE, COMPANY, DATES, DESCRIPTION, or TITLE fields. Only use <<>> inside the Summary text and the dash (-) bullet point lines.

CRITICAL — Avoid buzzwords and repetition:
- Do NOT use these overused cliché words/phrases ANYWHERE in the resume (Summary, bullets, descriptions — nowhere): "synergy", "leverage", "utilize", "spearheaded", "dynamic", "innovative", "cutting-edge", "best-in-class", "robust", "scalable solutions", "strategic", "proactive", "fast-paced environment", "go-to person", "think outside the box", "team player", "results-driven", "self-starter", "detail-oriented", "passion", "thrive", "adept at", "proven track record", "presentation skills", "strong communication skills", "highly skilled", "extensive experience", "seasoned professional", "committed to", "dedicated to", "responsible for", "in charge of", "various", "numerous", "several projects", "advanced knowledge", "expert in".
- The Summary must contain ZERO vague phrases. Every sentence must state a specific technology, a concrete number, or a measurable domain scope — never generic self-descriptions like "proven track record" or "adept at problem-solving".
- Instead use direct, concrete language. Say exactly what you did and what happened.
- ABSOLUTE RULE — ZERO repeated action verbs: Every single bullet point across the ENTIRE resume must start with a UNIQUE action verb. No verb may appear more than ONCE in the entire document — not even across different companies. Before writing each bullet, check every previous bullet to make sure you have not already used that verb.
- BANNED frequently-overused verbs (do NOT use these AT ALL): "Enhanced", "Improved", "Contributed", "Managed", "Assisted", "Helped", "Worked", "Handled", "Supported", "Utilized", "Developed". These are weak, generic verbs that ATS scanners and resume reviewers penalize.
- Instead, use strong, specific action verbs. Here is a bank of 50+ verbs — pick each one ONLY ONCE:
  Architected, Engineered, Constructed, Formulated, Created, Configured, Established, Assembled, Deployed, Orchestrated, Migrated, Automated, Redesigned, Refactored, Streamlined, Consolidated, Integrated, Pioneered, Implemented, Executed, Transformed, Accelerated, Parallelized, Optimized, Eliminated, Reduced, Resolved, Diagnosed, Overhauled, Modernized, Decoupled, Provisioned, Instrumented, Benchmarked, Standardized, Documented, Authored, Directed, Mentored, Coordinated, Negotiated, Presented, Delivered, Launched, Scaled, Built, Programmed, Coded, Prototyped, Containerized, Virtualized, Partitioned, Indexed, Normalized, Catalogued, Secured, Hardened, Audited, Monitored, Tested, Validated, Certified.
- NEVER repeat the same phrase or sentence structure more than once. Vary sentence length and pattern.
- Do NOT start more than one bullet with the same word across all companies combined.
- In the Summary, avoid generic filler phrases. Every sentence must state a concrete skill, domain, or accomplishment.

CRITICAL — Strong accomplishment-driven bullets with HARD DATA:
- Hiring managers at the senior level look for hard data above everything else. Every bullet MUST contain at least one concrete number. Not vague words like "significantly" or "greatly" — actual numbers.
- Every bullet must follow this formula: [Action verb] + [what you built/did] + [with what technology] + [measurable result with a number].
- Types of hard data to include: record counts (3M rows, 500K events/day), latency (from 12s to 800ms), user counts (serving 10K concurrent users), system scale (across 40 microservices, 15 data sources), time saved (reduced from 3 hours to 10 minutes), team/project scope (led team of 6, across 4 squads), SLA targets (four-nines uptime, p99 latency under 200ms), data volume (processing 2TB daily), infrastructure (deployed to 8-node cluster).
- REALISTIC & CONSISTENT numbers: the figures you invent must be believable for the company's scale and the role's seniority — a junior role at a small firm must NOT claim 5M events per day or a team of 50. Match the magnitude to the seniority (junior: modest scope; senior/staff: large scale and leadership). Keep the numbers within a single bullet consistent with each other, and vary them across bullets — do NOT make every metric a suspiciously round figure (mix values like "1.8M", "240ms", "37 services", "6-person squad").
- BAD (no data): "Managed data pipelines" or "Contributed to system improvements" or "Improved system performance significantly"
- GOOD (hard data): "Architected a real-time ingestion pipeline using <<Apache Kafka>> and <<Spark Structured Streaming>>, processing 5M events daily across 12 partitioned topics with sub-second latency"
- BAD (vague): "Enhanced data quality across the organization"
- GOOD (specific): "Built an automated data validation framework using <<Great Expectations>> and <<Python>>, catching 150+ schema violations per week across 30 production tables"
- Every role must demonstrate IMPACT. For senior roles: show leadership, architecture decisions, scale. For mid roles: show ownership, technical depth, team contributions. For junior roles: show learning speed, hands-on delivery, initiative.
- Each company section should read as a compelling narrative of accomplishments — not a list of generic duties or responsibilities.

ATS (Applicant Tracking System) Optimization — this resume WILL be scored by AI resume scanners. You MUST maximize keyword match score:

STEP 1 — Keyword extraction (do this mentally before writing):
Read the entire job description and extract EVERY keyword, phrase, and term into these categories:
a) Job title variations (e.g., "Software Engineer", "SWE", "Developer")
b) Hard skills & technologies (e.g., "Python", "AWS", "Docker", "Kubernetes")
c) Methodologies & processes (e.g., "Agile", "Scrum", "Kanban", "SDLC", "DevOps", "TDD", "BDD", "CI/CD")
d) Domain/industry terms (e.g., "data pipeline", "ETL", "microservices", "cloud migration", "distributed systems", "machine learning")
e) Certifications & standards mentioned (e.g., "AWS Certified", "ISO 27001", "SOC 2", "PCI DSS", "HIPAA")
f) Soft skill phrases the JD uses (e.g., "cross-functional collaboration", "stakeholder communication", "technical mentorship", "code review")
g) Business context terms (e.g., "SaaS", "B2B", "e-commerce", "fintech", "healthcare", "supply chain")
h) Action/responsibility phrases (e.g., "troubleshoot", "root cause analysis", "capacity planning", "incident response", "performance tuning")

STEP 2 — Keyword placement rules:
- Use a realistic, clean version of the job title (without internal level codes, team names, or company-specific jargon) in the SUBTITLE and the most recent company TITLE.
- EVERY keyword from categories (a) through (h) must appear at least ONCE somewhere in the resume — in Summary, Skills, bullet points, or company descriptions.
- Use the EXACT phrasing from the job description. If the JD says "RESTful APIs", write "RESTful APIs" — not "REST services" or "API development". If the JD says "cross-functional teams", write "cross-functional teams" — not "working with other teams".
- Include both the spelled-out form AND abbreviation where applicable (e.g., "Continuous Integration/Continuous Deployment (CI/CD)", "Amazon Web Services (AWS)", "Software Development Life Cycle (SDLC)").
- Front-load the most important keywords in the Summary (first 2-3 sentences). The Summary should contain at least 60% of the JD's key terms.
- In Skills, list every technology, tool, methodology, framework, and platform mentioned in the job description. Also include closely related industry-standard tools that someone in this role would use, even if not explicitly listed.
- Distribute keywords across ALL sections — do not cluster them only in Skills. Bullet points should naturally embed 2-4 JD keywords each.
- Use standard section headers: Summary, Skills, Professional Experience, Education.

STEP 3 — Industry keyword coverage:
- Beyond what the JD explicitly lists, include standard industry keywords that hiring managers in this field expect. For example, a backend engineer role should mention "API design", "database optimization", "system design", "load balancing", "caching", "logging and monitoring" even if the JD doesn't spell out every one.
- If the JD mentions a cloud provider (AWS/Azure/GCP), include 3-5 specific services from that provider in your bullets (e.g., for AWS: "EC2", "S3", "Lambda", "RDS", "CloudWatch").
- If the JD mentions "Agile" or "Scrum", reference sprint planning, retrospectives, or iterative development in at least one bullet.
- If the JD mentions collaboration or leadership, include bullets about code reviews, mentoring junior developers, or leading technical discussions.

Summary:
Write 4-6 impactful sentences. Match the tone to the job description seniority:
- If the job is senior/staff level: mention years of experience (use the number __YOE__) and deep expertise.
- If the job is mid-level: mention solid experience and growing expertise.
- If the job is junior/entry level: focus on enthusiasm, hands-on skills, and foundational knowledge — do NOT mention "9 years of experience" for a junior role.
Always mention the specific tech skills required in the job description.
Wrap technical keywords in <<>> in the Summary too, e.g., "Experienced in <<Python>>, <<Jira>>, and <<REST API>> integrations."

Skills:

Professional Experience:
For each company, use this EXACT structured format (one field per line):
COMPANY: Company Name, Remote
DATES: Month YYYY – Month YYYY (e.g., "February 2022 – January 2024", "March 2019 – May 2022". Always use full month name, NOT numeric format like "2024-02")
CRITICAL — DATES must be taken VERBATIM from the company list provided below. Use the exact From and To dates as given. Do NOT invent, round, or change any date. If the To value is "Present" or empty, write "Present".
DESCRIPTION: One-sentence company description that highlights the parts of the company's actual business most relevant to the job description. Frame the company's real work using keywords and domain language from the job posting, but never invent products, services, or lines of business. For example, if the job is about "data engineering", describe the company as "A technology firm specializing in enterprise data platform solutions and cloud-based analytics infrastructure" — not just "A software company".
TITLE: Detailed Job Title
- bullet point 1
- bullet point 2
...
IMPORTANT: For the COMPANY field, always use the company name followed by "Remote" — do NOT use City or State. Example: "COMPANY: Cibirix, Remote"

MANDATORY bullet count per role — you MUST meet these minimums or the output is invalid:
- Most recent company (latest dates): MINIMUM 10 bullet points, ideally 12. Do NOT write fewer than 10.
- Second most recent company: MINIMUM 8 bullet points, ideally 10. Do NOT write fewer than 8.
- Third company: MINIMUM 6 bullet points. Do NOT write fewer than 6.
- Fourth and older companies: MINIMUM 4 bullet points.
Count your bullets for each company before moving to the next. If a company has fewer than the minimum, add more bullets before proceeding.
This is the most important formatting rule — recruiters judge experience depth by bullet count.

Write bullet points for each role. Each bullet MUST:
1. Start with a strong action verb (Architected, Engineered, Optimized, Orchestrated, Migrated, Automated, Implemented, Redesigned, Built, Constructed, Configured, Established, Formulated, Consolidated, Refactored, Streamlined, Integrated, Deployed).
2. Name specific technologies wrapped in <<>> (e.g., "using <<Kafka>>, <<Spark>>, and <<Airflow>>").
3. Describe the technical problem or project scope concretely (e.g., "real-time event streaming pipeline processing 5M events per day" not just "data pipeline").
4. End with a measurable result using absolute numbers — NEVER percentages. Use counts, durations, throughput, or scale (e.g., "reducing query time from 12s to 800ms", "serving 10K concurrent users", "across 40 microservices").
Do NOT write vague bullets like "Improved system performance" — always specify what system, what technology, and what measurable result.
ABSOLUTE RULE: The resume must contain ZERO percentages. No "40%", no "25 percent", no "by 50%", no "increase of 30%". If you catch yourself writing a percentage, replace it with an absolute number (e.g., instead of "improved speed by 40%" write "improved speed from 5s to 800ms"). Scan your entire output and remove every single percentage before finishing.

Example of GOOD bullets:
- Architected a real-time data ingestion pipeline using <<Apache Kafka>> and <<Spark Structured Streaming>>, processing 3M events daily with sub-second latency across 12 partitioned topics.
- Migrated 15 legacy monolithic <<ETL>> workflows to <<Apache Airflow>> DAGs on <<AWS MWAA>>, cutting pipeline recovery time from 3 hours to under 10 minutes.
- Built a medallion architecture (Bronze/Silver/Gold) in <<Microsoft Fabric>> Lakehouse, enabling self-service analytics for 50 business users and reducing report generation from 4 hours to 15 minutes.

OVERRIDE — use this company-name rule instead of any earlier company-name rule:
- For OpenAI company-name source, choose a different REAL, VERIFIABLE employer company for each experience slot below. Never invent, blend, or alter a company name.
- Each company must actually exist and be headquartered in the United States.
- Each company must be easy to verify: it has an official website and a LinkedIn company page, so a Google search for the exact name finds it right away. A Wikipedia article is NOT required.
- Each company must have been operating during that slot's full date range (founded before the From date, not shut down or acquired before the To date).
- The company names must fit the job description's industry/domain.
- Each company must be a lesser-known, small-to-mid-sized company: roughly 50 to 1,000 employees, usually privately held, respected inside its niche but not a household name. Do NOT pick Fortune 500 companies, large publicly traded companies, unicorn startups, or any brand an average person would recognize (for example Google, Amazon, Apple, Meta, Microsoft, Netflix, Salesforce, Oracle, IBM, Adobe, Uber, Airbnb, Stripe, Shopify, Zillow, Chewy, Paycom). Also avoid one-person shops and brand-new startups.
- Use these as a guide — good places to find such companies: regional software and IT services firms, niche B2B SaaS vendors, digital agencies, consultancies, and specialized technology vendors in the job's industry (for example a regional healthcare IT vendor for a healthcare role, or a logistics software firm for a supply-chain role).
- Write each company's exact official name and spelling as shown on its website and LinkedIn page. Do not add or drop words like "Solutions", "Technologies", or "Group" unless they are part of the official name.
- Only pick a company you are certain is real. If you are unsure whether a smaller company exists, pick a different small-to-mid-sized company you are certain about — never fall back to a large, well-known brand.
- For a real company, the DESCRIPTION must be factually accurate about what that company actually does; emphasize the parts of its real business related to the job, but never make up products or lines of business.
- Do NOT use the hiring company or target company from the job description.
- Do NOT use the target/hiring company listed above.
- Do NOT use any company name that appears in the job description.
- Do NOT use generic experience slot labels as final company names.
- If the selected company-name source is OpenAI, do NOT use the profile company names as final COMPANY names; choose replacement real employers that follow the rules above.
- If the selected company-name source is Profile, keep the profile company names in the COMPANY fields.
- Use the exact number of experience slots below. Keep the dates and order from those slots.
- Output exactly the company/experience count shown below. Do not add, skip, or merge experience blocks.

__COMPANIESTEXT__

Please show experience sorted by latest company.

Education:
For each degree, use this EXACT structured format (one field per line):
UNIVERSITY: University Name, City, State
GRADUATION: Month YYYY
DEGREE: Full degree name (e.g., Bachelor of Science in Computer Science)
COURSEWORK: course1 • course2 • course3

__EDUCATIONTEXT__
"""


def _parse_ym(value: str):
    value = (value or "").strip()
    m = re.match(r"^(\d{1,2})[/-](\d{4})$", value)
    if m:
        return int(m.group(2)), int(m.group(1))
    m = re.match(r"^(\d{4})$", value)
    if m:
        return int(m.group(1)), 1
    return None


def years_of_experience(profile: dict) -> int:
    """Total career span (earliest start -> latest end) from the real company
    dates, rounded to the nearest year. Replaces the prompt's hardcoded '9'."""
    today = date.today()
    starts, ends = [], []
    for c in profile.get("companies", []):
        s = _parse_ym(c.get("start", ""))
        if s:
            starts.append(s)
        end_raw = (c.get("end", "") or "").strip()
        if end_raw.lower() in ("", "present"):
            ends.append((today.year, today.month))
        else:
            ends.append(_parse_ym(end_raw) or (today.year, today.month))
    if not starts:
        return 9
    sy, sm = min(starts)
    ey, em = max(ends)
    months = (ey - sy) * 12 + (em - sm)
    years = months // 12 + (1 if months % 12 >= 6 else 0)
    return max(1, years)


def build_companies_text(profile: dict) -> str:
    lines = []
    lines.append(f"Total experience/company count: {len(profile['companies'])}.")
    for i, c in enumerate(profile["companies"], start=1):
        lines.append(
            f"- Experience slot {i}: "
            f"Profile company name: {c.get('name', '')}. "
            f"Location to keep: {c.get('location', '')}. "
            f"From: {fmt_date(c.get('start',''))}, "
            f"To: {fmt_date(c.get('end',''), is_end=True)}"
        )
    return "\n".join(lines)


def build_education_text(profile: dict) -> str:
    lines = []
    for e in profile["education"]:
        lines.append(
            f"- {e['university']}, {e.get('location', '')}, "
            f"{fmt_date(e.get('start',''))} - {fmt_date(e.get('end',''), is_end=True)}. "
            f"Choose a DEGREE that best fits the job description."
        )
    return "\n".join(lines)


def build_prompt(profile: dict, job_description: str, target_company: str = "",
                 company_source: str = "profile", template: str = PROMPT_TEMPLATE) -> str:
    source_label = "OpenAI" if company_source == "openai" else "Profile"
    return (template
            .replace("__JOBDESC__", job_description)
            .replace("__TARGETCOMPANY__", target_company or "Not provided; infer it from the job description if present")
            .replace("__COMPANYSOURCE__", source_label)
            .replace("__YOE__", str(years_of_experience(profile)))
            .replace("__COMPANIESTEXT__", build_companies_text(profile))
            .replace("__EDUCATIONTEXT__", build_education_text(profile)))


# ---------------------------------------------------------------------------
# Per-model prompts (edited on the Prompts page)
# ---------------------------------------------------------------------------
# prompts.json maps "default" or "model:<name>" to {"system", "prompt",
# "updated"}. A model uses its own prompt, else the "default" one, else the
# built-in SYSTEM_PROMPT + PROMPT_TEMPLATE. Prompts belong to the AI model, not
# the provider: OpenAI's "gpt-4o" and OpenRouter's "openai/gpt-4o" share one.

PROMPTS_PATH = BASE_DIR / "prompts.json"
DEFAULT_PROMPT_KEY = "default"
SYSTEM_PROMPT = ("You are an expert resume writer. Follow the user's instructions exactly "
                 "and output plain text only, in the precise structure requested.")
# Shown on the Prompts page; build_prompt() fills these in.
PROMPT_PLACEHOLDERS = [
    ("__JOBDESC__", "The job description pasted on the Generate page"),
    ("__TARGETCOMPANY__", "The company name typed on the Generate page"),
    ("__COMPANYSOURCE__", "\"Profile\" or \"OpenAI\" (the company-names choice)"),
    ("__YOE__", "Years of experience, from the profile's dates"),
    ("__COMPANIESTEXT__", "One line per profile company: name, location, dates"),
    ("__EDUCATIONTEXT__", "One line per profile school: name, location, dates"),
]


def model_name(model: str) -> str:
    """The AI model without its provider prefix: "openai/gpt-4o" -> "gpt-4o",
    "deepseek-ai/deepseek-v4.1-flash" -> "deepseek-v4.1-flash"."""
    return model.rsplit("/", 1)[-1].lstrip("~")


def prompt_key(provider: str, model: str) -> str:
    return f"model:{model_name(model)}"


def prompt_models() -> dict:
    """{prompt key: [(provider, model id), ...]} for every offered model, so a
    model reachable through several providers appears once."""
    groups = {}
    for provider, models in MODEL_CHOICES.items():
        for m in models:
            groups.setdefault(prompt_key(provider, m), []).append((provider, m))
    return groups


def prompt_keys() -> list:
    """Every key the Prompts page can edit: the default, then each model."""
    return [DEFAULT_PROMPT_KEY] + list(prompt_models())


def load_prompts() -> dict:
    try:
        data = json.loads(PROMPTS_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_prompts(prompts: dict) -> None:
    # Write to a temp file first so a crash can't leave a half-written file.
    tmp = PROMPTS_PATH.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(prompts, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, PROMPTS_PATH)


def resolve_prompt(provider: str, model: str, prompts: dict = None) -> tuple:
    """(system message, prompt template, source) for a model, where source is
    "model", "default" or "built-in"."""
    prompts = load_prompts() if prompts is None else prompts
    for key, source in ((prompt_key(provider, model), "model"), (DEFAULT_PROMPT_KEY, "default")):
        entry = prompts.get(key) or {}
        if entry.get("prompt", "").strip():
            return entry.get("system", "").strip() or SYSTEM_PROMPT, entry["prompt"], source
    return SYSTEM_PROMPT, PROMPT_TEMPLATE, "built-in"


# ---------------------------------------------------------------------------
# Parse the model's plain-text resume into structured data
# ---------------------------------------------------------------------------

_HEADERS = {"summary": "summary", "skills": "skills",
            "professional experience": "experience", "experience": "experience",
            "education": "education"}


def _field(line: str) -> str:
    return line.split(":", 1)[1].strip() if ":" in line else ""


def _strip_kw(text: str) -> str:
    """Remove <<>> markup. Used for fields that should render as plain text
    (skills, titles, descriptions) — the markup belongs only in summary/bullets."""
    return (text or "").replace("<<", "").replace(">>", "")


def response_text(resp) -> str:
    """Model reply text, minus any <think>...</think> reasoning block that some
    open models (DeepSeek, Kimi, Nemotron, ...) put before the answer."""
    text = resp.choices[0].message.content or ""
    return re.sub(r"(?s)^\s*<think>.*?</think>\s*", "", text)


def parse_resume_text(text: str) -> dict:
    subtitle = ""
    summary_parts, skills, experience, education = [], [], [], []
    section = None
    cur_exp = cur_edu = None

    def flush():
        nonlocal cur_exp, cur_edu
        if cur_exp:
            experience.append(cur_exp); cur_exp = None
        if cur_edu:
            education.append(cur_edu); cur_edu = None

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue

        if line.upper().startswith("SUBTITLE:"):
            subtitle = _field(line)
            continue

        # A section header is a line that is JUST the header word, optionally
        # followed by a colon and nothing else (e.g. "Summary" or "Skills:").
        parts = line.split(":", 1)
        head = parts[0].strip().lower()
        after = parts[1].strip() if len(parts) > 1 else ""
        if head in _HEADERS and not after:
            flush()
            section = _HEADERS[head]
            continue

        if section == "summary":
            summary_parts.append(line)

        elif section == "skills":
            if ":" in line:
                label, items = line.split(":", 1)
                vals = [_strip_kw(s).strip() for s in items.split(",") if s.strip()]
                if vals:
                    skills.append({"label": _strip_kw(label).strip(), "skills": vals})

        elif section == "experience":
            U = line.upper()
            if U.startswith("COMPANY:"):
                flush()
                cur_exp = {"company": _field(line), "dates": "", "description": "",
                           "title": "", "bullets": []}
            elif cur_exp is None:
                continue
            elif U.startswith("DATES:"):
                cur_exp["dates"] = _field(line)
            elif U.startswith("DESCRIPTION:"):
                cur_exp["description"] = _strip_kw(_field(line))
            elif U.startswith("TITLE:"):
                cur_exp["title"] = _strip_kw(_field(line))
            elif line[0] in "-•*":
                cur_exp["bullets"].append(line[1:].strip())

        elif section == "education":
            U = line.upper()
            if U.startswith("UNIVERSITY:"):
                flush()
                cur_edu = {"university": _field(line), "graduation": "",
                           "degree": "", "coursework": ""}
            elif cur_edu is None:
                continue
            elif U.startswith("GRADUATION:"):
                cur_edu["graduation"] = _field(line)
            elif U.startswith("DEGREE:"):
                cur_edu["degree"] = _strip_kw(_field(line))
            elif U.startswith("COURSEWORK:"):
                cur_edu["coursework"] = _strip_kw(_field(line))

    flush()
    return {"subtitle": subtitle, "summary": " ".join(summary_parts),
            "skills": skills, "experience": experience, "education": education}


def _norm_company(name: str) -> str:
    """Normalize a company name for matching: take the part before the first
    comma (drops the ', Remote' the model appends) and keep only alphanumerics.
    'BYLTAX, Remote' and 'byltax' both become 'byltax'."""
    head = (name or "").split(",", 1)[0]
    return re.sub(r"[^a-z0-9]+", "", head.lower())


def clean_company_name(name: str) -> str:
    """Use only the company name part from model output like 'Acme, Remote'."""
    return (name or "").split(",", 1)[0].strip()


def invalid_openai_company_names(parsed_exp: list, profile: dict,
                                 target_company: str = "") -> list:
    """Problems that make AI-chosen company names unsafe to use."""
    problems = []
    profile_names = {_norm_company(c.get("name", "")) for c in profile["companies"]}
    profile_names.discard("")
    target = _norm_company(target_company)
    seen = set()
    if len(parsed_exp) < len(profile["companies"]):
        problems.append(
            f"returned {len(parsed_exp)} company blocks for "
            f"{len(profile['companies'])} profile companies"
        )
    for i, c in enumerate(profile["companies"]):
        ai = parsed_exp[i] if i < len(parsed_exp) else {}
        company = clean_company_name(ai.get("company", ""))
        norm = _norm_company(company)
        if not norm:
            problems.append(f"slot {i + 1} has no generated company name")
        elif norm in profile_names:
            problems.append(f"slot {i + 1} reused profile company '{c.get('name', '')}'")
        elif target and norm == target:
            problems.append(f"slot {i + 1} reused target company '{target_company}'")
        elif norm in seen:
            problems.append(f"slot {i + 1} duplicated generated company '{company}'")
        seen.add(norm)
    return problems


def ai_client(provider: str) -> tuple:
    """Return (client, model) for the chosen provider. NVIDIA and OpenRouter speak
    the OpenAI API, so the same client works with a different base URL."""
    if provider == "nvidia":
        api_key = os.getenv("NVIDIA_API_KEY")
        if not api_key or api_key.startswith("nvapi-paste"):
            raise RuntimeError("No NVIDIA API key configured. Edit .env and add NVIDIA_API_KEY "
                               "(free key at https://build.nvidia.com).")
        return OpenAI(api_key=api_key, base_url=NVIDIA_BASE_URL), NVIDIA_MODEL
    if provider == "openrouter":
        api_key = os.getenv("OPENROUTER_API_KEY")
        if not api_key or api_key.startswith("sk-or-paste"):
            raise RuntimeError("No OpenRouter API key configured. Edit .env and add OPENROUTER_API_KEY.")
        # Optional OpenRouter header: names the app on openrouter.ai's activity page.
        return OpenAI(api_key=api_key, base_url=OPENROUTER_BASE_URL,
                      default_headers={"X-OpenRouter-Title": "Resume Generator"}), OPENROUTER_MODEL
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key or api_key.startswith("sk-paste"):
        raise RuntimeError("No OpenAI API key configured. Edit .env and add your key.")
    return OpenAI(api_key=api_key), MODEL


def create_chat(client, **kwargs):
    """chat.completions.create, but wait out short per-minute rate limits
    (e.g. "Please try again in 7.65s") instead of failing."""
    for attempt in range(3):
        try:
            return client.chat.completions.create(**kwargs)
        except RateLimitError as exc:
            m = re.search(r"try again in (?:(\d+)m(?!s))?([\d.]+)(ms|s)", str(exc))
            wait = None
            if m:
                secs = float(m.group(2)) / (1000 if m.group(3) == "ms" else 1)
                wait = int(m.group(1) or 0) * 60 + secs + 1
            if wait is None or wait > 65 or attempt == 2:
                raise
            time.sleep(wait)


def tailor_resume(profile: dict, job_description: str, target_company: str = "",
                  company_source: str = "profile", provider: str = "openai",
                  model: str = "") -> dict:
    client, default_model = ai_client(provider)
    model = model or default_model
    system_prompt, template, _ = resolve_prompt(provider, model)
    limits = {}
    # The GPT-5.x/6.x reasoning families only accept the default temperature (1)
    # and reject any other value, unlike gpt-4o.
    if not re.search(r"gpt-[5-9]", model):
        limits["temperature"] = 0.6
    if provider == "openrouter":
        # GPT-4o writes a resume in ~1,100 tokens; reasoning models (GPT-5.x/6.x,
        # Claude) also spend tokens thinking, so give them more room — DeepSeek
        # thinks the longest (measured up to ~6,200 tokens incl. the resume).
        if "gpt-4o" in model:
            limits["max_tokens"] = OPENROUTER_MAX_TOKENS
        else:
            limits["max_tokens"] = 12000 if "deepseek" in model else 8000
    elif provider == "nvidia":
        limits["max_tokens"] = NVIDIA_MAX_TOKENS
        if "nemotron" in model:
            # Skip Nemotron's reasoning pass: ~3x faster, same resume quality.
            limits["extra_body"] = {"chat_template_kwargs": {"enable_thinking": False}}

    # Fast models occasionally return a reply the parser can't read (no Summary
    # or Skills section); ask once more rather than saving a broken PDF.
    for _ in range(2):
        resp = create_chat(
            client,
            model=model,
            **limits,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": build_prompt(profile, job_description, target_company,
                                                         company_source, template)},
            ],
        )
        parsed = parse_resume_text(response_text(resp))
        if parsed["summary"] and parsed["skills"]:
            break
    if company_source == "openai":
        problems = invalid_openai_company_names(
            parsed["experience"], profile, target_company)
        if problems:
            retry_prompt = (
                build_prompt(profile, job_description, target_company, company_source, template)
                + "\n\nSTRICT CORRECTION REQUIRED:\n"
                + "The previous company names were invalid because: "
                + "; ".join(problems)
                + ". Generate the full resume again. Use exactly "
                + str(len(profile["companies"]))
                + " experience/company blocks. Do not use profile company names, "
                + "generic slot labels, duplicate company names, or the target/hiring company. "
                + "Every company must be a real, verifiable, U.S.-headquartered, lesser-known "
                + "small-to-mid-sized company (not a large or household-name brand) "
                + "written with its exact official name — never an invented name."
            )
            resp = create_chat(
                client,
                model=model,
                **limits,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": retry_prompt},
                ],
            )
            parsed = parse_resume_text(response_text(resp))
            problems = invalid_openai_company_names(
                parsed["experience"], profile, target_company)
            if problems:
                raise RuntimeError(
                    "OpenAI company-name generation failed: " + "; ".join(problems)
                )

    # Overlay AI prose onto the FIXED facts (company/dates/location/school verbatim).
    # Match each profile company to the AI block with the SAME name rather than
    # trusting list position — if the model reorders/drops a company, position
    # matching would attach bullets to the wrong employer. Fall back to position
    # only when no name matches (and that positional block isn't already taken).
    parsed_exp = parsed["experience"]
    experience = []
    for i, c in enumerate(profile["companies"]):
        ai = parsed_exp[i] if i < len(parsed_exp) else {}
        ai_company = clean_company_name(ai.get("company", ""))
        company = ai_company if company_source == "openai" and ai_company else c["name"]
        experience.append({
            "company": company,
            "location": c.get("location", ""),
            "dates": fmt_range(c.get("start", ""), c.get("end", "")),
            "title": ai.get("title", ""),
            "description": ai.get("description", ""),
            "bullets": [strip_percentages(b) for b in ai.get("bullets", [])],
        })

    education = []
    for i, e in enumerate(profile["education"]):
        ai = parsed["education"][i] if i < len(parsed["education"]) else {}
        education.append({
            "university": e["university"],
            "location": e.get("location", ""),
            "dates": fmt_range(e.get("start", ""), e.get("end", "")),
            "degree": ai.get("degree", ""),
            "coursework": ai.get("coursework", ""),
        })

    return {"subtitle": parsed["subtitle"],
            "summary": strip_percentages(parsed["summary"]),
            "skills": parsed["skills"], "experience": experience, "education": education}


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    active = get_active_id()
    company_source = session.get("company_source", "profile")
    if company_source not in ("profile", "openai"):
        company_source = "profile"
    ai_provider = session.get("ai_provider", "openai")
    if ai_provider not in AI_PROVIDERS:
        ai_provider = "openai"
    print("ai_provider====" + ai_provider)
    ai_models = {p: chosen_model(p, session.get(f"ai_model_{p}", "")) for p in AI_PROVIDERS}
    saved_prompts = load_prompts()
    # For each "provider:model": which prompt it uses ("model" / "default" /
    # "built-in") and its Prompts-page key, shown under the Model dropdown.
    prompt_sources = {f"{p}:{m}": {"source": resolve_prompt(p, m, saved_prompts)[2],
                                   "key": prompt_key(p, m)}
                      for p, models in MODEL_CHOICES.items() for m in models}
    return render_template("index.html", profiles=list_profile_names(list_profiles()),
                           active=active, profile=load_profile(active),
                           company_source=company_source, ai_provider=ai_provider,
                           ai_models=ai_models, model_choices=MODEL_CHOICES,
                           prompt_sources=prompt_sources)


@app.route("/profile", methods=["GET", "POST"])
def profile():
    if request.method == "POST":
        pid = request.form.get("profile_id", "").strip() or get_active_id()
        if not pid:  # no profile to save into — create one from the form name
            pid = create_profile(request.form.get("name", "").strip() or "profile")
        save_profile(parse_profile_form(request.form), pid)
        set_active_id(pid)
        flash("Profile saved.", "success")
        return redirect(url_for("profile", id=pid))

    profiles = list_profiles()
    pid = request.args.get("id", "").strip() or get_active_id()
    if pid:
        set_active_id(pid)
    profile_data = load_profile(pid)
    names = list_profile_names(profiles)
    return render_template("profile.html", profiles=profiles, current=pid,
                           profile=profile_data, names=names)


@app.route("/profile/new", methods=["POST"])
def profile_new():
    pid = create_profile(request.form.get("new_name", "").strip() or "New profile")
    set_active_id(pid)
    flash("Profile created — fill in the details and save.", "success")
    return redirect(url_for("profile", id=pid))


@app.route("/profile/delete", methods=["POST"])
def profile_delete():
    pid = request.form.get("profile_id", "").strip()
    if pid:
        delete_profile(pid)
        flash("Profile deleted.", "success")
    remaining = list_profiles()
    set_active_id(remaining[0]["id"] if remaining else None)
    return redirect(url_for("profile"))


def render_pdf(profile_data: dict, tailored: dict) -> bytes:
    html_doc = render_template("resume.html", profile=profile_data, tailored=tailored,
                               preview=False)
    return HTML(string=html_doc, base_url=str(BASE_DIR)).write_pdf()


def _slug(text: str) -> str:
    """Filesystem-safe filename segment: alphanumerics, runs collapsed to '_'."""
    return re.sub(r"[^A-Za-z0-9]+", "_", text or "").strip("_")


def build_filename(profile_data: dict, company: str) -> str:
    """e.g. 'Spencer_Aguas_Acme_1735649895.pdf', or
    'Spencer_Aguas_1735649895.pdf' when no company is given. The Unix timestamp
    keeps each download unique; if two resumes land in the same second, a numeric
    suffix is added so a PDF is never silently overwritten."""
    name = _slug(profile_data.get("name", "")) or "resume"
    stamp = int(time.time())
    company = _slug(company)
    base = "_".join([name] + ([company] if company else []) + [str(stamp)])
    candidate = base + ".pdf"
    n = 2
    while (RESUMES_DIR / candidate).exists():
        candidate = f"{base}_{n}.pdf"
        n += 1
    return candidate


def jd_filename(pdf_filename: str) -> str:
    """The job-description text file that pairs with a resume PDF: same base name,
    '.txt' extension (e.g. 'Spencer_Aguas_Acme_123.pdf' -> '..._123.txt')."""
    return Path(pdf_filename).with_suffix(".txt").name


def save_job_description(pdf_filename: str, job_description: str, company: str,
                         job_title: str, job_link: str) -> None:
    """Save the job description next to its resume PDF as a self-documenting .txt,
    so you can always see what each resume was tailored to. Best-effort: a write
    failure here must never lose the generated PDF."""
    try:
        header = [
            f"Company:   {company or '—'}",
            f"Job title: {job_title or '—'}",
            f"Job link:  {job_link or '—'}",
            f"Saved:     {time.strftime('%Y-%m-%d %H:%M')}",
            f"Resume:    {pdf_filename}",
            "=" * 60,
            "",
        ]
        text = "\n".join(header) + job_description.strip() + "\n"
        (RESUMES_DIR / jd_filename(pdf_filename)).write_text(text, encoding="utf-8")
    except OSError as exc:
        app.logger.warning("Could not save job description: %s", exc)


def _save_workbook(wb, path: Path) -> None:
    """Save atomically: write to a unique temp file in the same folder, then
    os.replace it into place. A crash or concurrent write can never leave a
    half-written / truncated workbook.

    The temp name is unique per save (mkstemp) so two requests writing the same
    profile's log at once don't clobber a shared temp file, and a stale temp from
    a past crash can never block future saves.

    On Windows the final replace can hit a lock on the destination and raise
    PermissionError — WinError 5 (access denied, e.g. OneDrive/antivirus) or
    WinError 32 (sharing violation, the .xlsx is open in Excel). Transient locks
    clear in milliseconds, so retry with a short backoff. A persistent lock still
    raises, with a clear message telling the user to close the file."""
    fd, tmp_name = tempfile.mkstemp(
        dir=str(path.parent), prefix=path.stem + "_", suffix=".tmp.xlsx")
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        wb.save(tmp)
        last_err = None
        for attempt in range(6):
            try:
                os.replace(tmp, path)
                return  # success — tmp no longer exists, finally is a no-op
            except PermissionError as e:  # WinError 5 / 32 on a locked destination
                last_err = e
                time.sleep(0.25 * (attempt + 1))
        raise PermissionError(
            f"Could not save {path.name}: the file is locked. Close it in Excel "
            f"(and consider moving the project out of OneDrive) and try again."
        ) from last_err
    finally:
        if tmp.exists():  # only on failure; clean up the orphaned temp file
            try:
                os.remove(tmp)
            except OSError:
                pass


def log_resume(pid: str, job_title: str, url: str, company: str, filename: str,
               progress: str = "", note: str = "") -> None:
    """Append a row to the profile's log (resumes_log_<pid>.xlsx). Columns: Date,
    Job Title, URL, Company, Filename, Progress, Note. Never breaks the download."""
    path = resumes_log_path(pid)
    try:
        if path.exists():
            wb = load_workbook(path)
            ws = wb.active
            if ws.max_row == 0 or [c.value for c in ws[1]] != LOG_HEADERS:
                ws.insert_rows(1)
                for col, header in enumerate(LOG_HEADERS, start=1):
                    ws.cell(row=1, column=col, value=header)
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Resumes Log"
            ws.append(LOG_HEADERS)
        ws.append([time.strftime("%Y-%m-%d %H:%M:%S"), job_title, url, company,
                   filename, progress, note])
        _save_workbook(wb, path)
    except Exception as exc:  # noqa: BLE001 — never block the download on logging
        app.logger.warning("Could not write resumes log: %s", exc)


def resume_sort_key(date_value: str, filename: str, row: int) -> tuple:
    """Newest-first sort key. Prefer the Unix timestamp embedded in filenames."""
    m = re.search(r"_(\d{10})(?:_\d+)?\.pdf$", filename or "")
    if m:
        return (int(m.group(1)), row)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return (int(time.mktime(time.strptime(str(date_value), fmt))), row)
        except (TypeError, ValueError):
            pass
    return (0, row)


def read_log_rows(pid: str) -> list:
    """A profile's logged resumes as dicts, plus whether each PDF still exists."""
    if not pid:
        return []
    path = resumes_log_path(pid)
    if not path.exists():
        return []
    # read_only=True keeps the file handle open until wb.close(); on Windows a
    # leaked handle locks the file and the next write fails (WinError 32). Always
    # close it in a finally so the workbook is never left open.
    wb = None
    rows = []
    try:
        wb = load_workbook(path, read_only=True)
        ws = wb.active
        for idx, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not r or all(c is None for c in r):
                continue
            r = list(r) + [None] * (7 - len(r))
            filename = str(r[4] or "")
            progress = str(r[5] or "").strip()
            date_value = str(r[0] or "")
            rows.append({
                "row": idx,  # actual worksheet row, used to update Progress in place
                "profile_id": pid,
                "sort_key": resume_sort_key(date_value, filename, idx),
                "date": date_value,
                "job_title": str(r[1] or ""),
                "url": str(r[2] or ""),
                "company": str(r[3] or ""),
                "filename": filename,
                "progress": progress,
                # Match any case/variant ("REJECTED", "Rejected (phone only)", ...)
                "is_rejected": progress.upper().startswith("REJECT"),
                "note": str(r[6] or ""),
                "exists": bool(filename) and (RESUMES_DIR / filename).exists(),
                # The paired job-description text file, if it was saved alongside.
                "jd": jd_filename(filename) if filename else "",
                "jd_exists": bool(filename)
                and (RESUMES_DIR / jd_filename(filename)).exists(),
            })
    except Exception as exc:  # noqa: BLE001
        app.logger.warning("Could not read resumes log: %s", exc)
        return []
    finally:
        if wb is not None:
            wb.close()
    return rows


# ---------------------------------------------------------------------------
# Statistics — how many resumes were generated per day (/stats page)
# ---------------------------------------------------------------------------

# Range choices for the daily chart: (key, label, number of days, None = all time)
STATS_RANGES = (
    ("7", "Last 7 days", 7),
    ("14", "Last 14 days", 14),
    ("30", "Last 30 days", 30),
    ("90", "Last 90 days", 90),
    ("all", "All time", None),
)
STATS_RANGE_DEFAULT = "30"
STATS_MONTHS = 12          # months shown in the monthly chart
STATS_MAX_BARS = 730       # never render more than ~2 years of daily bars


def _log_date(value: str) -> str:
    """A log row's Date cell is 'YYYY-MM-DD HH:MM:SS' (older rows are just
    'YYYY-MM-DD'). Return the YYYY-MM-DD part, or '' when it isn't a date."""
    m = re.match(r"^\s*(\d{4}-\d{2}-\d{2})", str(value or ""))
    return m.group(1) if m else ""


def daily_counts(rows: list) -> dict:
    """{'YYYY-MM-DD': count} over a list of log rows."""
    counts = {}
    for r in rows:
        day = _log_date(r.get("date", ""))
        if day:
            counts[day] = counts.get(day, 0) + 1
    return counts


def monthly_counts(counts: dict) -> dict:
    """{'YYYY-MM': count} rolled up from the daily counts."""
    months = {}
    for day, n in counts.items():
        key = day[:7]
        months[key] = months.get(key, 0) + n
    return months


def build_daily_series(counts: dict, days: int | None, today: date) -> list:
    """One entry per day, oldest → newest, with the count and bar height (% of
    peak). days=None means all time: from the first logged day through today."""
    if days is None:
        first = date.fromisoformat(min(counts)) if counts else today
        span = (today - first).days + 1
    else:
        span = days
    span = min(max(span, 1), STATS_MAX_BARS)
    start = today - timedelta(days=span - 1)
    series = []
    for i in range(span):
        day = start + timedelta(days=i)
        key = day.isoformat()
        series.append({"date": key, "label": day.strftime("%d/%m"),
                       "count": counts.get(key, 0)})
    peak = max((s["count"] for s in series), default=0)
    for s in series:
        s["pct"] = round(s["count"] / peak * 100) if peak else 0
        s["is_today"] = s["date"] == today.isoformat()
    return series


def build_monthly_series(counts: dict, months: int, today: date) -> list:
    """The last `months` calendar months (oldest → newest), with counts and bar
    height (% of peak)."""
    keys, y, m = [], today.year, today.month
    for _ in range(months):
        keys.append(f"{y:04d}-{m:02d}")
        m -= 1
        if m == 0:
            y, m = y - 1, 12
    series = [{"key": k, "label": f"{_MONTHS[int(k[5:7])][:3]} {k[2:4]}",
               "count": counts.get(k, 0)} for k in reversed(keys)]
    peak = max((s["count"] for s in series), default=0)
    for s in series:
        s["pct"] = round(s["count"] / peak * 100) if peak else 0
        s["is_current"] = s["key"] == f"{today.year:04d}-{today.month:02d}"
    return series


def stats_summary(counts: dict, rows: list, today: date) -> dict:
    """Headline numbers for the stats page."""
    def window(days: int) -> int:
        start = today - timedelta(days=days - 1)
        return sum(n for day, n in counts.items()
                   if start <= date.fromisoformat(day) <= today)

    best_day, best_count = "", 0
    for day, n in counts.items():
        if n > best_count or (n == best_count and day > best_day):
            best_day, best_count = day, n

    # Consecutive days with at least one resume, ending today (or yesterday).
    streak, cursor = 0, today
    if counts.get(cursor.isoformat(), 0) == 0:
        cursor -= timedelta(days=1)
    while counts.get(cursor.isoformat(), 0) > 0:
        streak += 1
        cursor -= timedelta(days=1)

    companies = {(r.get("company") or "").strip().lower()
                 for r in rows if (r.get("company") or "").strip()}
    return {
        "total": len(rows),
        "today": counts.get(today.isoformat(), 0),
        "yesterday": counts.get((today - timedelta(days=1)).isoformat(), 0),
        "week": window(7),
        "month": window(30),
        "avg_day": round(window(30) / 30, 1),
        "active_days": len(counts),
        "best_day": best_day,
        "best_day_label": (f"{best_day[8:10]}/{best_day[5:7]}/{best_day[0:4]}"
                           if best_day else "—"),
        "best_count": best_count,
        "streak": streak,
        "rejected": sum(1 for r in rows if r.get("is_rejected")),
        "companies": len(companies),
    }


@app.route("/stats")
def stats():
    """Daily generation counts + charts for the selected name's resumes."""
    _migrate_legacy_log()
    profiles = list_profiles()
    pid = request.args.get("profile", "").strip() or get_active_id()
    if pid:
        set_active_id(pid)
    current_profile = load_profile(pid)
    current_name = current_profile.get("name", "")
    names = list_profile_names(profiles)

    # Same scope as the Resumes page: every log belonging to this name.
    rows = []
    for p in profiles_for_name(current_name, profiles):
        rows.extend(read_log_rows(p["id"]))

    today = date.today()
    counts = daily_counts(rows)

    range_key = request.args.get("range", "").strip() or STATS_RANGE_DEFAULT
    range_days = {key: days for key, _label, days in STATS_RANGES}
    if range_key not in range_days:
        range_key = STATS_RANGE_DEFAULT
    range_label = next(label for key, label, _d in STATS_RANGES if key == range_key)

    daily = build_daily_series(counts, range_days[range_key], today)
    monthly = build_monthly_series(monthly_counts(counts), STATS_MONTHS, today)

    # The table follows the selected range so it always agrees with the chart.
    in_range = {s["date"] for s in daily if s["count"]}
    daily_table = [
        {"date": day, "weekday": date.fromisoformat(day).strftime("%a"),
         "count": counts[day]}
        for day in sorted(in_range, reverse=True)
    ]
    range_total = sum(row["count"] for row in daily_table)
    for row in daily_table:
        row["pct"] = round(row["count"] / range_total * 100) if range_total else 0

    return render_template(
        "stats.html", summary=stats_summary(counts, rows, today),
        daily=daily, daily_peak=max((s["count"] for s in daily), default=0),
        monthly=monthly, daily_table=daily_table, range_total=range_total,
        ranges=STATS_RANGES, range_key=range_key, range_label=range_label,
        today=today, profiles=profiles, current=pid, names=names,
        current_name=current_name)


@app.route("/prompts")
def prompts():
    saved = load_prompts()
    keys = prompt_keys()
    current = request.args.get("key", "")
    if current not in keys:
        current = DEFAULT_PROMPT_KEY

    choices = [(k, "Default (all models)" if k == DEFAULT_PROMPT_KEY else k.split(":", 1)[1])
               for k in keys]
    if current in saved:
        status = "Custom prompt"
    elif current == DEFAULT_PROMPT_KEY:
        status = "Built-in prompt"
    else:
        status = "Uses the default prompt"

    # Show the prompt this entry uses now, so editing starts from it.
    if current in saved:
        prompt_text = saved[current].get("prompt", "")
    elif current == DEFAULT_PROMPT_KEY:
        prompt_text = PROMPT_TEMPLATE
    else:
        provider, model = prompt_models()[current][0]
        prompt_text = resolve_prompt(provider, model, saved)[1]

    return render_template("prompts.html", choices=choices, current=current,
                           saved_keys=set(saved), status=status, prompt_text=prompt_text,
                           is_custom=current in saved, placeholders=PROMPT_PLACEHOLDERS)


@app.route("/prompts/save", methods=["POST"])
def save_prompt():
    key = request.form.get("key", "")
    if key not in prompt_keys():
        abort(400)
    saved = load_prompts()
    if request.form.get("action") == "reset":
        saved.pop(key, None)
        save_prompts(saved)
        flash("Reset to the default prompt.", "success")
        return redirect(url_for("prompts", key=key))

    prompt_text = request.form.get("prompt", "").replace("\r\n", "\n")
    if not prompt_text.strip():
        flash("The prompt is empty — nothing saved.", "error")
        return redirect(url_for("prompts", key=key))
    saved[key] = {"prompt": prompt_text, "updated": datetime.now().strftime("%Y-%m-%d %H:%M")}
    save_prompts(saved)
    if "__JOBDESC__" in prompt_text:
        flash("Prompt saved.", "success")
    else:
        flash("Prompt saved — but it has no __JOBDESC__, so the job description won't be sent.",
              "error")
    return redirect(url_for("prompts", key=key))


@app.route("/generate", methods=["POST"])
def generate():
    job_description = request.form.get("job_description", "").strip()
    if not job_description:
        flash("Please paste a job description first.", "error")
        return redirect(url_for("index"))

    pid = request.form.get("profile_id", "").strip() or get_active_id()
    profile_data = load_profile(pid)
    if pid:
        set_active_id(pid)  # remember the last profile used
    if not profile_data["companies"]:
        flash("Add your experience slots on the Profile page first.", "error")
        return redirect(url_for("profile"))

    company = request.form.get("company_name", "").strip()
    job_link = request.form.get("job_link", "").strip()
    job_title = request.form.get("job_title", "").strip()
    company_source = request.form.get("company_source", "profile").strip()
    if company_source not in ("profile", "openai"):
        company_source = "profile"
    session["company_source"] = company_source
    ai_provider = request.form.get("ai_provider", "openai").strip()
    if ai_provider not in AI_PROVIDERS:
        ai_provider = "openai"
    session["ai_provider"] = ai_provider
    # Each provider has its own Model dropdown (ai_model_<provider>); remember
    # the pick per provider so switching providers keeps each one's choice.
    ai_model = chosen_model(ai_provider, request.form.get(f"ai_model_{ai_provider}", "").strip())
    session[f"ai_model_{ai_provider}"] = ai_model

    try:
        tailored = tailor_resume(profile_data, job_description, company, company_source,
                                 ai_provider, ai_model)
        for w in build_warnings(tailored):
            app.logger.warning("Resume parse: %s", w)
        pdf_bytes = render_pdf(profile_data, tailored)
    except Exception as exc:  # noqa: BLE001
        flash(f"Generation failed: {exc}", "error")
        return redirect(url_for("index"))

    # Save the finished PDF into the project's resumes/ folder (not the browser
    # Downloads folder), then return to the page with a confirmation.
    filename = build_filename(profile_data, company)

    RESUMES_DIR.mkdir(exist_ok=True)
    (RESUMES_DIR / filename).write_bytes(pdf_bytes)
    save_job_description(filename, job_description, company, job_title, job_link)
    log_resume(pid, job_title, job_link, company, filename)

    flash(f"Saved to resumes/{filename}", "success")
    return redirect(url_for("index"))


RESUMES_PER_PAGE_CHOICES = (20, 50, 100, 500)
RESUMES_PER_PAGE_DEFAULT = RESUMES_PER_PAGE_CHOICES[0]


def _resumes_redirect(pid: str, query: str, page: int = None, per_page: int = None):
    args = {}
    if query:
        args["q"] = query
    if pid:
        args["profile"] = pid
    if page and page > 1:
        args["page"] = page
    if per_page and per_page != RESUMES_PER_PAGE_DEFAULT:
        args["per_page"] = per_page
    return redirect(url_for("resumes", **args))


@app.route("/resumes")
def resumes():
    _migrate_legacy_log()
    profiles = list_profiles()
    pid = request.args.get("profile", "").strip() or get_active_id()
    if pid:
        set_active_id(pid)
    query = request.args.get("q", "").strip()
    current_profile = load_profile(pid)
    current_name = current_profile.get("name", "")
    names = list_profile_names(profiles)
    selected_profiles = profiles_for_name(current_name, profiles)
    rows = []
    for p in selected_profiles:
        for row in read_log_rows(p["id"]):
            row["profile_id"] = p["id"]
            rows.append(row)
    rows.sort(key=lambda r: r["sort_key"], reverse=True)
    total = len(rows)
    if query:
        q = query.lower()
        rows = [r for r in rows
                if q in r["job_title"].lower()
                or q in r["company"].lower()
                or q in r["url"].lower()]
    shown = len(rows)

    per_page = request.args.get("per_page", type=int) or RESUMES_PER_PAGE_DEFAULT
    if per_page not in RESUMES_PER_PAGE_CHOICES:
        per_page = RESUMES_PER_PAGE_DEFAULT
    page_count = max(1, math.ceil(shown / per_page)) if shown else 1
    page = request.args.get("page", type=int) or 1
    page = min(max(page, 1), page_count)
    start = (page - 1) * per_page
    rows = rows[start:start + per_page]

    return render_template("resumes.html", rows=rows, q=query, total=total,
                           shown=shown, profiles=profiles, current=pid,
                           names=names, current_name=current_name,
                           page=page, page_count=page_count, per_page=per_page,
                           per_page_choices=RESUMES_PER_PAGE_CHOICES)


PROGRESS_VALUES = ("", "Rejected")


@app.route("/resumes/progress", methods=["POST"])
def update_progress():
    """Set a row's Progress to '' or 'Rejected' (column 6) in the profile's log."""
    row = request.form.get("row", type=int)
    progress = request.form.get("progress", "").strip()
    query = request.form.get("q", "").strip()
    pid = request.form.get("profile_id", "").strip() or get_active_id()
    page = request.form.get("page", type=int)
    per_page = request.form.get("per_page", type=int)
    if progress not in PROGRESS_VALUES:
        progress = ""
    path = resumes_log_path(pid) if pid else None
    if row and row >= 2 and path and path.exists():
        try:
            wb = load_workbook(path)
            ws = wb.active
            if row <= ws.max_row:
                # Assign .value directly: ws.cell(..., value=None) is a no-op in
                # openpyxl (it won't clear an existing value), so blanking fails.
                ws.cell(row=row, column=6).value = progress or None
                _save_workbook(wb, path)
        except Exception as exc:  # noqa: BLE001
            app.logger.warning("Could not update progress: %s", exc)
            flash("Could not update progress (is the Excel file open elsewhere?).", "error")
    return _resumes_redirect(pid, query, page, per_page)


@app.route("/resumes/note", methods=["POST"])
def update_note():
    """Set a row's free-text Note (column 7) in the profile's log."""
    row = request.form.get("row", type=int)
    note = request.form.get("note", "").strip()
    query = request.form.get("q", "").strip()
    pid = request.form.get("profile_id", "").strip() or get_active_id()
    page = request.form.get("page", type=int)
    per_page = request.form.get("per_page", type=int)
    path = resumes_log_path(pid) if pid else None
    if row and row >= 2 and path and path.exists():
        try:
            wb = load_workbook(path)
            ws = wb.active
            if row <= ws.max_row:
                ws.cell(row=row, column=7).value = note or None
                _save_workbook(wb, path)
        except Exception as exc:  # noqa: BLE001
            app.logger.warning("Could not update note: %s", exc)
            flash("Could not update note (is the Excel file open elsewhere?).", "error")
    return _resumes_redirect(pid, query, page, per_page)


@app.route("/resumes/delete", methods=["POST"])
def delete_resume():
    """Delete a log row (and its PDF if saved in resumes/) from the profile's log."""
    row = request.form.get("row", type=int)
    query = request.form.get("q", "").strip()
    pid = request.form.get("profile_id", "").strip() or get_active_id()
    page = request.form.get("page", type=int)
    per_page = request.form.get("per_page", type=int)
    path = resumes_log_path(pid) if pid else None
    if row and row >= 2 and path and path.exists():
        try:
            wb = load_workbook(path)
            ws = wb.active
            if row <= ws.max_row:
                filename = ws.cell(row=row, column=5).value
                ws.delete_rows(row, 1)
                _save_workbook(wb, path)
                # Remove the saved PDF and its companion job-description .txt too
                # (Path(...).name blocks path traversal).
                if filename:
                    safe = Path(str(filename)).name
                    for f in (RESUMES_DIR / safe, RESUMES_DIR / jd_filename(safe)):
                        if f.exists():
                            f.unlink()
                flash("Deleted resume entry.", "success")
        except Exception as exc:  # noqa: BLE001
            app.logger.warning("Could not delete row: %s", exc)
            flash("Could not delete (is the Excel file open elsewhere?).", "error")
    return _resumes_redirect(pid, query, page, per_page)


@app.route("/resumes/file/<path:filename>")
def resume_file(filename):
    """Serve a saved PDF from resumes/. send_from_directory blocks path traversal."""
    if not (RESUMES_DIR / filename).exists():
        abort(404)
    return send_from_directory(RESUMES_DIR, filename)


if __name__ == "__main__":
    # host="0.0.0.0" listens on all interfaces so the app is reachable via the LAN IP, not just 127.0.0.1.
    app.run(debug=True, host="0.0.0.0", port=int(os.getenv("PORT", 5000)), use_reloader=False)
