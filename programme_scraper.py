"""MVP crawler and extractor for inbound short-term programme pages.

Reads input_urls.xlsx and writes output_programmes.xlsx with candidate pages,
OpenAI-extracted programme rows, and a run log. The functions in this file are
used by both the command-line script and the simple Windows GUI app.
"""
from __future__ import annotations

import json
import os
import re
import time
from collections import deque
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Callable, Dict, List, Optional, Set, Tuple
from urllib.parse import urldefrag, urljoin, urlparse
from urllib.robotparser import RobotFileParser

import pandas as pd
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openai import OpenAI

INPUT_FILE = "input_urls.xlsx"
OUTPUT_FILE = "output_programmes.xlsx"
CRAWL_DEPTH_LIMIT = 2
MAX_PAGES_PER_UNIVERSITY = 50
REQUEST_DELAY_SECONDS = 1.5
RELEVANCE_THRESHOLD = 60
USER_AGENT = "ProgrammeFinderMVP/1.0 (+https://example.local; polite educational crawler)"
DEFAULT_OPENAI_MODEL = "gpt-4o-mini"
ProgressCallback = Optional[Callable[[str], None]]

POSITIVE_KEYWORDS = [
    "short-term programme", "short-term program", "summer programme", "summer program",
    "winter programme", "winter program", "Japanese language programme",
    "Japanese language program", "Japanese studies", "study in Japan", "inbound programme",
    "inbound program", "international students", "overseas students", "partner universities",
    "custom programme", "custom program", "留学生", "短期受入", "短期プログラム",
    "日本語プログラム", "サマープログラム", "ウィンタープログラム", "海外大学向け",
    "受入プログラム", "協定校向け",
]
NEGATIVE_KEYWORDS = [
    "outbound", "study abroad for Japanese students", "outgoing exchange", "派遣留学",
    "海外留学", "日本人学生向け", "本学学生向け",
]
AVOID_URL_PARTS = ["/login", "signin", "auth", "wp-login", "logout", "contact", "inquiry"]
OUTPUT_PROGRAMME_COLUMNS = [
    "University", "Programme Name", "Country or Region", "Programme Type", "Target Participants",
    "Period", "Duration", "Eligibility", "Programme Fee", "Capacity", "Application Deadline",
    "Academic Year", "Summary", "Source URL", "Evidence Text", "Check Status",
    "Notes for Human Review",
]
INPUT_COLUMNS = ["University", "Start URL", "Notes"]
CANDIDATE_COLUMNS = [
    "University", "Page Title", "URL", "Relevance Score",
    "Found Keywords", "Reason", "PDF Links", "Notes",
]
RUN_LOG_COLUMNS = [
    "University", "Start URL", "Status", "Pages Visited",
    "Candidate Pages Found", "Programmes Extracted", "Error Message",
    "Run DateTime",
]
SHEET_CANDIDATE_PAGES = "Candidate Pages"
SHEET_EXTRACTED_PROGRAMMES = "Extracted Programmes"
SHEET_RUN_LOG = "Run Log"


@dataclass
class PageResult:
    title: str
    url: str
    score: int
    found_keywords: str
    reason: str
    pdf_links: str
    notes: str
    text: str


@dataclass
class RobotsPolicy:
    """robots.txt rules plus whether they were actually available."""

    parser: RobotFileParser
    available: bool
    reason: str = ""


def report(progress_callback: ProgressCallback, message: str) -> None:
    """Send a progress message to the GUI or print nothing in CLI mode."""
    if progress_callback:
        progress_callback(message)


def normalise_url(url: str) -> str:
    """Resolve fragments and standardise a URL enough to avoid duplicate visits."""
    clean, _ = urldefrag(str(url).strip())
    parsed = urlparse(clean)
    if not parsed.scheme:
        clean = "https://" + clean
        parsed = urlparse(clean)
    return parsed._replace(fragment="").geturl().rstrip("/")


def related_domain(hostname: str) -> str:
    """Return a practical base domain, handling common Japanese university domains."""
    parts = hostname.lower().split(".")
    if len(parts) >= 3 and parts[-2] in {"ac", "co", "go", "or", "ne"} and parts[-1] == "jp":
        return ".".join(parts[-3:])
    if len(parts) >= 2:
        return ".".join(parts[-2:])
    return hostname.lower()


def is_related_url(url: str, start_host: str) -> bool:
    """Keep crawling on the same host or clearly related university subdomains."""
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        return False
    host = parsed.hostname.lower()
    base = related_domain(start_host)
    return host == start_host.lower() or host.endswith("." + base) or host == base


def should_skip_url(url: str) -> bool:
    """Avoid login/contact-like URLs and non-web links.

    This MVP only reads public HTML pages. It does not submit forms,
    enter login areas, or intentionally collect personal data.
    """
    lower = url.lower()
    return any(part in lower for part in AVOID_URL_PARTS) or lower.startswith("mailto:") or lower.startswith("tel:")


def get_robot_policy(start_url: str) -> RobotsPolicy:
    """Read robots.txt when possible, with a safe polite fallback.

    If robots.txt is readable, the crawler respects it. If robots.txt cannot be
    read because of a network error, timeout, URL error, or server-side 5xx
    response, this MVP does not treat that as "block every page". Instead, it
    continues politely using the strict existing safeguards: same/related-domain
    only, depth and page limits, request delay, and login/form avoidance.
    """
    parsed = urlparse(start_url)
    robots_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    rp = RobotFileParser()
    rp.set_url(robots_url)
    try:
        response = requests.get(robots_url, headers={"User-Agent": USER_AGENT}, timeout=10)
        if 200 <= response.status_code < 300:
            rp.parse(response.text.splitlines())
            return RobotsPolicy(rp, True)
        if 400 <= response.status_code < 500:
            return RobotsPolicy(rp, False, f"robots.txt not present or unavailable ({response.status_code})")
        return RobotsPolicy(rp, False, f"robots.txt server error ({response.status_code})")
    except requests.RequestException as exc:
        return RobotsPolicy(rp, False, f"robots.txt could not be read: {exc}")


def fetch_html(url: str, robots_policy: RobotsPolicy) -> Optional[str]:
    if robots_policy.available and not robots_policy.parser.can_fetch(USER_AGENT, url):
        return None
    headers = {"User-Agent": USER_AGENT}
    response = requests.get(url, headers=headers, timeout=20)
    response.raise_for_status()
    content_type = response.headers.get("Content-Type", "").lower()
    if "text/html" not in content_type and "application/xhtml" not in content_type:
        return None
    return response.text


def extract_text_and_links(html: str, page_url: str) -> Tuple[str, str, List[str], List[str]]:
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg", "form"]):
        tag.decompose()
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
    links: List[str] = []
    pdfs: List[str] = []
    for anchor in soup.find_all("a", href=True):
        href = urljoin(page_url, anchor["href"])
        href = normalise_url(href)
        if href.lower().endswith(".pdf"):
            pdfs.append(href)
        else:
            links.append(href)
    return title, text, sorted(set(links)), sorted(set(pdfs))


def score_page(title: str, text: str, url: str) -> Tuple[int, List[str], str, str]:
    haystack = f"{title} {url} {text[:8000]}".lower()
    positives = [kw for kw in POSITIVE_KEYWORDS if kw.lower() in haystack]
    negatives = [kw for kw in NEGATIVE_KEYWORDS if kw.lower() in haystack]
    score = min(100, len(positives) * 18)
    if any(term in haystack for term in ["application", "fee", "schedule", "eligibility", "duration"]):
        score += 10
    if any(term in haystack for term in ["international", "overseas", "留学生", "受入"]):
        score += 10
    score -= len(negatives) * 25
    score = max(0, min(100, score))
    reason = "Positive keywords found" if positives else "No strong positive keyword found"
    if negatives:
        reason += "; negative outbound keywords also found"
    notes = "PDF review needed" if url.lower().endswith(".pdf") else ""
    return score, positives, reason, notes


def crawl_university(university: str, start_url: str, progress_callback: ProgressCallback = None) -> Tuple[List[PageResult], int, Optional[str]]:
    start_url = normalise_url(start_url)
    start_host = urlparse(start_url).hostname or ""
    robots_policy = get_robot_policy(start_url)
    if not robots_policy.available and robots_policy.reason:
        report(progress_callback, f"{university}: {robots_policy.reason}; continuing politely with strict crawl limits")
    queue = deque([(start_url, 0)])
    visited: Set[str] = set()
    candidates: List[PageResult] = []
    error_message = None

    while queue and len(visited) < MAX_PAGES_PER_UNIVERSITY:
        url, depth = queue.popleft()
        if url in visited or should_skip_url(url) or not is_related_url(url, start_host):
            continue
        visited.add(url)
        report(progress_callback, f"{university}: visiting page {len(visited)}/{MAX_PAGES_PER_UNIVERSITY}: {url}")
        try:
            html = fetch_html(url, robots_policy)
            time.sleep(REQUEST_DELAY_SECONDS)
            if html is None:
                continue
            title, text, links, pdf_links = extract_text_and_links(html, url)
            score, keywords, reason, notes = score_page(title, text, url)
            if pdf_links:
                notes = "; ".join(filter(None, [notes, "PDF review needed"]))
            if score >= 20 or pdf_links:
                candidates.append(PageResult(title, url, score, ", ".join(keywords), reason, "; ".join(pdf_links), notes, text[:12000]))
            if depth < CRAWL_DEPTH_LIMIT:
                for link in links:
                    if link not in visited and is_related_url(link, start_host) and not should_skip_url(link):
                        queue.append((link, depth + 1))
        except Exception as exc:
            error_message = str(exc)
            report(progress_callback, f"Warning: could not read {url}: {exc}")
            continue
    candidates.sort(key=lambda p: p.score, reverse=True)
    return candidates, len(visited), error_message


def read_input_file(input_path: str) -> pd.DataFrame:
    """Read input_urls.xlsx and verify the required MVP columns are present."""
    inputs = pd.read_excel(input_path).fillna("")
    missing_columns = [column for column in INPUT_COLUMNS if column not in inputs.columns]
    if missing_columns:
        missing = ", ".join(missing_columns)
        required = ", ".join(INPUT_COLUMNS)
        raise ValueError(f"{input_path} is missing required column(s): {missing}. Required columns: {required}.")
    return inputs


def extract_programmes_with_openai(client: OpenAI, model: str, university: str, page: PageResult) -> Tuple[List[Dict[str, str]], Optional[str]]:
    prompt = f"""
Extract inbound short-term programme information for international students or overseas partner universities from this page.
Return valid JSON only, with this shape: {{"programmes": [{{"Programme Name":"", "Country or Region":"", "Programme Type":"", "Target Participants":"", "Period":"", "Duration":"", "Eligibility":"", "Programme Fee":"", "Capacity":"", "Application Deadline":"", "Academic Year":"", "Summary":"", "Evidence Text":"", "Notes for Human Review":""}}]}}.
Do not guess. Leave missing fields blank. Ignore outbound programmes for Japanese/domestic students. If details appear old or unclear, mention that in Notes for Human Review.
University: {university}
URL: {page.url}
Page title: {page.title}
Page text:
{page.text[:10000]}
"""
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            response_format={"type": "json_object"},
        )
        content = response.choices[0].message.content or "{}"
        data = json.loads(content)
        programmes = data.get("programmes", [])
        if not isinstance(programmes, list):
            return [], "OpenAI JSON did not contain a programmes list"
        rows: List[Dict[str, str]] = []
        for item in programmes:
            if not isinstance(item, dict):
                continue
            row = {column: "" for column in OUTPUT_PROGRAMME_COLUMNS}
            row["University"] = university
            row["Source URL"] = page.url
            row["Check Status"] = "Needs human review"
            for key in row:
                if key in item and item[key] is not None:
                    row[key] = str(item[key])
            if not row["Programme Name"] and not row["Summary"]:
                continue
            rows.append(row)
        return rows, None
    except Exception as exc:
        return [], str(exc)


def run_scraper(
    input_path: str = INPUT_FILE,
    output_path: str = OUTPUT_FILE,
    api_key: Optional[str] = None,
    model: Optional[str] = None,
    progress_callback: ProgressCallback = None,
) -> str:
    """Run the MVP scraper and return the created Excel file path.

    The GUI passes paths and API key values directly. The command-line version can
    still use `.env` values, so daily Windows users do not need PowerShell.
    """
    load_dotenv()
    input_path = input_path or os.getenv("INPUT_FILE", INPUT_FILE)
    output_path = output_path or os.getenv("OUTPUT_FILE", OUTPUT_FILE)
    model = model or os.getenv("OPENAI_MODEL", DEFAULT_OPENAI_MODEL)
    api_key = api_key or os.getenv("OPENAI_API_KEY")
    client = OpenAI(api_key=api_key) if api_key else None

    report(progress_callback, f"Reading input file: {input_path}")
    inputs = read_input_file(input_path)
    candidate_rows: List[Dict[str, object]] = []
    programme_rows: List[Dict[str, str]] = []
    log_rows: List[Dict[str, object]] = []

    for index, source in inputs.iterrows():
        university = str(source.get("University", "")).strip()
        start_url = str(source.get("Start URL", "")).strip()
        status = "Completed"
        error_parts: List[str] = []
        candidates: List[PageResult] = []
        pages_visited = 0
        report(progress_callback, f"Starting row {index + 1}: {university or '(missing university)'}")
        try:
            if not university or not start_url:
                raise ValueError("University and Start URL are required for each input row")
            candidates, pages_visited, crawl_error = crawl_university(university, start_url, progress_callback)
            if crawl_error:
                error_parts.append(crawl_error)
            for page in candidates:
                candidate_rows.append({
                    "University": university, "Page Title": page.title, "URL": page.url,
                    "Relevance Score": page.score, "Found Keywords": page.found_keywords,
                    "Reason": page.reason, "PDF Links": page.pdf_links, "Notes": page.notes,
                })
            extraction_pages = [p for p in candidates if p.score >= RELEVANCE_THRESHOLD]
            if not extraction_pages:
                extraction_pages = candidates[:3]
            if client:
                report(progress_callback, f"{university}: extracting from {len(extraction_pages)} candidate page(s)")
                for page in extraction_pages:
                    if page.score < 20:
                        continue
                    rows, extract_error = extract_programmes_with_openai(client, model, university, page)
                    programme_rows.extend(rows)
                    if extract_error:
                        error_parts.append(f"{page.url}: {extract_error}")
            else:
                error_parts.append("OPENAI_API_KEY not set; skipped OpenAI extraction")
                status = "Completed with warnings"
            report(progress_callback, f"Finished {university}: {len(candidates)} candidate page(s)")
        except Exception as exc:
            status = "Failed"
            error_parts.append(str(exc))
            report(progress_callback, f"Error for {university or start_url}: {exc}")
        log_rows.append({
            "University": university, "Start URL": start_url, "Status": status,
            "Pages Visited": pages_visited, "Candidate Pages Found": len(candidates),
            "Programmes Extracted": len([r for r in programme_rows if r.get("University") == university]),
            "Error Message": " | ".join(error_parts),
            "Run DateTime": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
        })

    report(progress_callback, f"Writing output file: {output_path}")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(candidate_rows, columns=CANDIDATE_COLUMNS).to_excel(writer, sheet_name=SHEET_CANDIDATE_PAGES, index=False)
        pd.DataFrame(programme_rows, columns=OUTPUT_PROGRAMME_COLUMNS).to_excel(writer, sheet_name=SHEET_EXTRACTED_PROGRAMMES, index=False)
        pd.DataFrame(log_rows, columns=RUN_LOG_COLUMNS).to_excel(writer, sheet_name=SHEET_RUN_LOG, index=False)
    report(progress_callback, f"Done. Created {output_path}")
    return output_path


def main() -> None:
    output_path = run_scraper()
    print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
