"""Crawler and extractor for inbound short-term programme pages.

Reads input_urls.xlsx and writes output_programmes.xlsx with candidate pages,
extracted programme rows, and a run log.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import time
from collections import deque
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Callable, Dict, List, Optional, Set, Tuple
from urllib.parse import unquote, urldefrag, urljoin, urlparse
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
FALLBACK_PAGE_LIMIT = 10
USER_AGENT = "ProgrammeFinderMVP/1.0 (+https://example.local; polite educational crawler)"

SEED_URLS = [
    "https://en.ritsumei.ac.jp/admissions-e/short-term-programs",
    "https://en.ritsumei.ac.jp/admissions/short-term-programs/japanese-language/",
    "https://en.ritsumei.ac.jp/international-e/inbound/",
    "https://rsjprwjp.com/",
]

HIGH_PRIORITY_KEYWORDS = [
    "RSJP", "RWJP", "RSJP Express", "RWJP Express", "Short-Term Programs",
    "Short-Term Non Degree Programs", "Japanese Language", "Japanese Language Program",
    "Customizable Programs", "Inbound Programs", "Programme", "Program",
    "Programme Fee", "Program Fee", "Application Deadline", "Eligibility",
    "Certificate", "Buddy", "Study in Kyoto Program", "SKP",
]
LOW_PRIORITY_KEYWORDS = ["international students", "留学生", "study abroad", "admissions", "campus life"]
NEGATIVE_KEYWORDS = [
    "outbound", "study abroad for Japanese students", "outgoing exchange", "派遣留学",
    "海外留学", "日本人学生向け", "本学学生向け",
]
AVOID_URL_PARTS = ["/login", "signin", "auth", "wp-login", "logout", "contact", "inquiry"]
IMPORTANT_FIELDS = [
    "Programme Type", "Target Students", "Language", "Duration / Period", "Programme Dates",
    "Application Deadline", "Programme Fee", "Housing", "Credits / Certificate",
    "Main Contents", "Eligibility",
]
OUTPUT_PROGRAMME_COLUMNS = [
    "University", "Programme Name", "Programme Type", "Institution / Department",
    "Target Students", "Language", "Duration / Period", "Programme Dates",
    "Application Deadline", "Programme Fee", "Housing", "Credits / Certificate",
    "Main Contents", "Eligibility", "Source URL", "Source Page Title",
    "Confidence Score", "Missing Fields", "Review Status", "Notes",
]
INPUT_COLUMNS = ["University", "Start URL", "Notes"]
CANDIDATE_COLUMNS = [
    "University", "URL", "Page Title", "Candidate Score", "Matched Keywords",
    "Candidate Type", "Reason", "Needs Review", "PDF Links", "Notes",
]
RUN_LOG_COLUMNS = [
    "University", "Start URL", "Status", "Start Time", "End Time", "Pages Visited",
    "Candidate Pages Found", "Programmes Extracted", "Warnings", "Error Details",
]


@dataclass
class PageContent:
    title: str
    headings: List[str]
    text: str
    links: List[Tuple[str, str]]
    pdf_links: List[Tuple[str, str]]


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
    headings: List[str]
    candidate_type: str = "HTML"
    needs_review: str = "No"


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


def get_start_urls(start_url: str) -> List[str]:
    """Load crawl seeds, prioritising Ritsumeikan-specific programme pages."""
    start = normalise_url(start_url)
    host = (urlparse(start).hostname or "").lower()
    urls = [start]
    if "ritsumei.ac.jp" in host:
        urls = SEED_URLS + [start]
    return list(dict.fromkeys(normalise_url(url) for url in urls))


def get_robot_parser(start_url: str) -> RobotFileParser:
    parsed = urlparse(start_url)
    robots_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    rp = RobotFileParser()
    rp.set_url(robots_url)
    try:
        response = requests.get(robots_url, headers={"User-Agent": USER_AGENT}, timeout=10)
        response.encoding = response.apparent_encoding or response.encoding
        if response.status_code < 500:
            response.raise_for_status()
            rp.parse(response.text.splitlines())
            return rp
    except requests.RequestException:
        pass

    # Safe robots.txt fallback: network errors, timeouts, URL errors, and 5xx
    # server responses mean we could not learn the site's rules, not that every
    # page is disallowed. Continue only under the crawler's strict safeguards:
    # same/related domains, depth and page limits, polite delay, and avoidance
    # of login, form, contact, and personal-data collection areas.
    rp.parse(["User-agent: *", "Allow: /"])
    return rp


def fetch_page(url: str, robot_parser: RobotFileParser) -> Optional[str]:
    """Fetch an HTML page and decode Japanese/English text reliably."""
    if not robot_parser.can_fetch(USER_AGENT, url):
        return None
    response = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=20)
    response.raise_for_status()
    content_type = response.headers.get("Content-Type", "").lower()
    if "text/html" not in content_type and "application/xhtml" not in content_type:
        return None
    if response.apparent_encoding:
        response.encoding = response.apparent_encoding
    elif not response.encoding:
        response.encoding = "utf-8"
    return response.text


def extract_page_content(html: str, page_url: str) -> PageContent:
    """Extract readable title, headings, body text, HTML links, and PDF links."""
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg", "form"]):
        tag.decompose()
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    headings = [tag.get_text(" ", strip=True) for tag in soup.find_all(["h1", "h2", "h3"]) if tag.get_text(" ", strip=True)]
    text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
    links: List[Tuple[str, str]] = []
    pdfs: List[Tuple[str, str]] = []
    for anchor in soup.find_all("a", href=True):
        href = normalise_url(urljoin(page_url, anchor["href"]))
        label = re.sub(r"\s+", " ", anchor.get_text(" ", strip=True))
        if href.lower().endswith(".pdf"):
            pdfs.append((href, label))
        else:
            links.append((href, label))
    return PageContent(title, headings, text, sorted(set(links)), sorted(set(pdfs)))


def score_candidate_page(title: str, text: str, url: str, link_text: str = "", is_pdf: bool = False) -> Tuple[int, List[str], str, str]:
    """Score programme relevance with emphasis on specific inbound programme terms."""
    url_lower = url.lower()
    title_lower = title.lower()
    haystack = f"{title} {link_text} {url} {text[:10000]}".lower()
    matched_high = [kw for kw in HIGH_PRIORITY_KEYWORDS if kw.lower() in haystack]
    matched_low = [kw for kw in LOW_PRIORITY_KEYWORDS if kw.lower() in haystack]
    negatives = [kw for kw in NEGATIVE_KEYWORDS if kw.lower() in haystack]
    score = len(matched_high) * 14 + len(matched_low) * 4
    reasons: List[str] = []
    if matched_high:
        reasons.append("specific programme keywords matched")
    if matched_low:
        reasons.append("general international keywords matched")
    boosts = {
        "short-term-programs": 18,
        "japanese-language": 18,
        "inbound": 14,
    }
    for term, points in boosts.items():
        if term in url_lower:
            score += points
            reasons.append(f"URL contains {term}")
    if "rsjp" in url_lower or "rsjp" in title_lower:
        score += 25
        reasons.append("URL/title contains RSJP")
    if "rwjp" in url_lower or "rwjp" in title_lower:
        score += 25
        reasons.append("URL/title contains RWJP")
    if "en.ritsumei.ac.jp" in url_lower:
        score += 10
        reasons.append("Ritsumeikan English site")
    if "rsjprwjp.com" in url_lower:
        score += 20
        reasons.append("RSJP/RWJP programme site")
    if any(term in haystack for term in ["application deadline", "programme fee", "program fee", "eligibility", "certificate", "buddy"]):
        score += 12
        reasons.append("programme detail fields found")
    if is_pdf:
        score = max(5, int(score * 0.7))
        reasons.append("PDF review needed")
    if negatives:
        score -= len(negatives) * 20
        reasons.append("negative outbound keywords found")
    score = max(0, min(100, score))
    notes = "PDF review needed" if is_pdf else ""
    return score, matched_high + matched_low, "; ".join(reasons) or "No strong programme signals", notes


def crawl_university(university: str, start_url: str) -> Tuple[List[PageResult], int, Optional[str]]:
    seeds = get_start_urls(start_url)
    start_host = urlparse(normalise_url(start_url)).hostname or ""
    allowed_seed_hosts = {urlparse(seed).hostname or "" for seed in seeds}
    robot_parsers: Dict[str, RobotFileParser] = {}
    queue = deque((seed, 0) for seed in seeds)
    visited: Set[str] = set()
    candidates_by_url: Dict[str, PageResult] = {}
    error_message = None

    while queue and len(visited) < MAX_PAGES_PER_UNIVERSITY:
        url, depth = queue.popleft()
        host = urlparse(url).hostname or ""
        if url in visited or should_skip_url(url) or not (is_related_url(url, start_host) or host in allowed_seed_hosts):
            continue
        visited.add(url)
        try:
            robot_parser = robot_parsers.setdefault(host, get_robot_parser(url))
            html = fetch_page(url, robot_parser)
            time.sleep(REQUEST_DELAY_SECONDS)
            if html is None:
                continue
            content = extract_page_content(html, url)
            score, keywords, reason, notes = score_candidate_page(content.title, content.text, url)
            if score >= 20:
                candidates_by_url[url] = PageResult(
                    content.title, url, score, ", ".join(keywords), reason, "", notes,
                    content.text[:12000], content.headings, "HTML", "No" if score >= RELEVANCE_THRESHOLD else "Yes",
                )
            for pdf_url, pdf_label in content.pdf_links:
                pdf_score, pdf_keywords, pdf_reason, pdf_notes = score_candidate_page(pdf_label, "", pdf_url, pdf_label, is_pdf=True)
                if pdf_score >= 10:
                    candidates_by_url[pdf_url] = PageResult(
                        pdf_label or unquote(pdf_url.rsplit("/", 1)[-1]), pdf_url, pdf_score,
                        ", ".join(pdf_keywords), pdf_reason, "", pdf_notes, "", [], "PDF", "Yes",
                    )
            if depth < CRAWL_DEPTH_LIMIT:
                for link, _label in content.links:
                    link_host = urlparse(link).hostname or ""
                    if link not in visited and (is_related_url(link, start_host) or link_host in allowed_seed_hosts) and not should_skip_url(link):
                        queue.append((link, depth + 1))
        except Exception as exc:
            error_message = str(exc)
            continue

    if not candidates_by_url:
        for seed in seeds:
            seed_score, seed_keywords, seed_reason, seed_notes = score_candidate_page("", "", seed)
            if seed_score >= 20:
                inferred_title = unquote(urlparse(seed).path.strip("/").split("/")[-1]).replace("-", " ").title() or urlparse(seed).netloc
                candidates_by_url[seed] = PageResult(
                    inferred_title, seed, seed_score, ", ".join(seed_keywords),
                    seed_reason + "; seed URL fallback because page could not be fetched",
                    "", seed_notes or "Fetch failed; review source page manually", "", [], "HTML", "Yes",
                )
    candidates = sorted(candidates_by_url.values(), key=lambda p: p.score, reverse=True)
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


def infer_programme_name(page: PageResult) -> Tuple[str, str]:
    """Infer a programme name from explicit title/headings or recognizable URL terms."""
    visible_text = " ".join([page.title, *page.headings])
    combined = " ".join([visible_text, page.url])
    known = [
        "RSJP Express", "RWJP Express", "RSJP", "RWJP", "Study in Kyoto Program", "SKP",
        "Short-Term Non Degree Programs", "Short-Term Programs", "Japanese Language Programs",
        "Japanese Language Program", "Customizable Programs", "Inbound Programs",
    ]
    for name in known:
        if name.lower() in visible_text.lower():
            return name, ""
    for name in known:
        if name.lower() in page.url.lower():
            return name, "Inferred from page title, heading, or URL"
    for heading in page.headings:
        if heading and len(heading) <= 120:
            return heading, "Inferred from page title, heading, or URL"
    if page.title:
        return page.title[:120], "Inferred from page title, heading, or URL"
    slug = unquote(urlparse(page.url).path.strip("/").split("/")[-1]).replace("-", " ").replace("_", " ").strip()
    return (slug.title() if slug else page.url), "Inferred from page title, heading, or URL"


def calculate_missing_fields(row: Dict[str, str]) -> str:
    """List important output fields that are blank or Unknown."""
    return ", ".join(field for field in IMPORTANT_FIELDS if not row.get(field) or row.get(field) == "Unknown")


def review_status(confidence: int, missing_fields: str) -> str:
    if confidence >= 75 and len([f for f in missing_fields.split(", ") if f]) <= 4:
        return "Likely valid"
    if confidence < 35:
        return "Low confidence"
    return "Needs human review"


def extract_programme_info_from_page(university: str, page: PageResult, item: Optional[Dict[str, object]] = None) -> Dict[str, str]:
    """Create an Excel row while keeping incomplete but relevant programme records."""
    row = {column: "" for column in OUTPUT_PROGRAMME_COLUMNS}
    row["University"] = university
    row["Institution / Department"] = university
    row["Source URL"] = page.url
    row["Source Page Title"] = page.title
    row["Confidence Score"] = str(page.score)
    notes: List[str] = []
    if item:
        aliases = {
            "Programme Name": ["Programme Name", "Program Name", "name"],
            "Programme Type": ["Programme Type", "Program Type", "type"],
            "Target Students": ["Target Students", "Target Participants", "participants"],
            "Language": ["Language", "Instruction Language"],
            "Duration / Period": ["Duration / Period", "Duration", "Period"],
            "Programme Dates": ["Programme Dates", "Program Dates", "Dates"],
            "Application Deadline": ["Application Deadline", "Deadline"],
            "Programme Fee": ["Programme Fee", "Program Fee", "Fee"],
            "Housing": ["Housing", "Accommodation"],
            "Credits / Certificate": ["Credits / Certificate", "Certificate", "Credits"],
            "Main Contents": ["Main Contents", "Summary", "Contents"],
            "Eligibility": ["Eligibility"],
            "Notes": ["Notes", "Notes for Human Review"],
        }
        for target, keys in aliases.items():
            for key in keys:
                value = item.get(key) if isinstance(item, dict) else None
                if value not in (None, ""):
                    row[target] = str(value)
                    break
    if not row["Programme Name"]:
        row["Programme Name"], note = infer_programme_name(page)
        if note:
            notes.append(note)
    if not row["Programme Type"]:
        if "language" in f"{page.title} {page.url}".lower():
            row["Programme Type"] = "Japanese language / short-term"
        elif "inbound" in page.url.lower():
            row["Programme Type"] = "Inbound programme"
        else:
            row["Programme Type"] = "Short-term / non-degree"
    if not row["Language"] and re.search(r"japanese language|rsjp|rwjp|日本語", f"{page.title} {page.text}", re.I):
        row["Language"] = "Japanese"
    if page.candidate_type == "PDF":
        notes.append("PDF review needed")
    if row["Notes"]:
        notes.insert(0, row["Notes"])
    row["Notes"] = "; ".join(dict.fromkeys(notes))
    row["Missing Fields"] = calculate_missing_fields(row)
    row["Review Status"] = review_status(page.score, row["Missing Fields"])
    return row


def extract_programmes_with_openai(client: OpenAI, model: str, university: str, page: PageResult) -> Tuple[List[Dict[str, str]], Optional[str]]:
    prompt = f"""
Extract inbound short-term, non-degree, Japanese language, customizable, or inbound programme information for international students or overseas partner universities from this page.
Return valid JSON only, with this shape: {{"programmes": [{{"Programme Name":"", "Programme Type":"", "Target Students":"", "Language":"", "Duration / Period":"", "Programme Dates":"", "Application Deadline":"", "Programme Fee":"", "Housing":"", "Credits / Certificate":"", "Main Contents":"", "Eligibility":"", "Notes":""}}]}}.
Do not guess fees, dates, deadlines, housing, or eligibility. Leave missing fields blank or Unknown. Do not discard a programme just because some fields are missing. Ignore outbound programmes for Japanese/domestic students.
University: {university}
URL: {page.url}
Page title: {page.title}
Headings: {' | '.join(page.headings[:10])}
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
        rows = [extract_programme_info_from_page(university, page, item) for item in programmes if isinstance(item, dict)]
        if not rows and page.score >= RELEVANCE_THRESHOLD:
            rows = [extract_programme_info_from_page(university, page)]
        return rows, None
    except Exception as exc:
        return [], str(exc)


def fallback_extract_programmes(university: str, candidates: List[PageResult], limit: int = FALLBACK_PAGE_LIMIT) -> List[Dict[str, str]]:
    """Guarantee useful review rows when relevant candidates exist but extraction returns zero."""
    synthetic_rows: List[Dict[str, str]] = []
    for page in candidates:
        if "rsjprwjp.com" in page.url.lower():
            for name in ["RSJP", "RSJP Express", "RWJP", "RWJP Express"]:
                synthetic_page = PageResult(name, page.url, page.score, page.found_keywords, page.reason, page.pdf_links, page.notes, page.text, [name], page.candidate_type, "Yes")
                row = extract_programme_info_from_page(university, synthetic_page)
                row["Review Status"] = "Needs human review"
                row["Notes"] = "; ".join(filter(None, [row["Notes"], "Inferred from page title, heading, or URL"]))
                synthetic_rows.append(row)
            break
    html_candidates = [p for p in candidates if p.candidate_type == "HTML" and p.score >= 20]
    if not html_candidates:
        html_candidates = [p for p in candidates if p.score >= 20]
    rows: List[Dict[str, str]] = synthetic_rows
    seen_names: Set[Tuple[str, str]] = {(row["Programme Name"].lower(), row["Source URL"]) for row in synthetic_rows}
    for page in html_candidates[:limit]:
        row = extract_programme_info_from_page(university, page)
        row["Review Status"] = "Needs human review"
        key = (row["Programme Name"].lower(), row["Source URL"])
        if key not in seen_names:
            rows.append(row)
            seen_names.add(key)
    return rows


def write_excel_output(output_path: str, candidate_rows: List[Dict[str, object]], programme_rows: List[Dict[str, str]], log_rows: List[Dict[str, object]]) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(candidate_rows, columns=CANDIDATE_COLUMNS).to_excel(writer, sheet_name="Candidate Pages", index=False)
        pd.DataFrame(programme_rows, columns=OUTPUT_PROGRAMME_COLUMNS).to_excel(writer, sheet_name="Extracted Programmes", index=False)
        pd.DataFrame(log_rows, columns=RUN_LOG_COLUMNS).to_excel(writer, sheet_name="Run Log", index=False)


def run_scraper(
    input_path: str = INPUT_FILE,
    output_path: str = OUTPUT_FILE,
    api_key: Optional[str] = None,
    model: str = "gpt-4o-mini",
    progress_callback: Optional[Callable[[str], None]] = None,
) -> str:
    """Run the crawler/extractor and write the Excel workbook."""
    def report(message: str) -> None:
        print(message)
        if progress_callback:
            progress_callback(message)

    client = OpenAI(api_key=api_key) if api_key else None
    inputs = read_input_file(input_path)
    candidate_rows: List[Dict[str, object]] = []
    programme_rows: List[Dict[str, str]] = []
    log_rows: List[Dict[str, object]] = []

    report(f"Loaded {len(inputs)} university row(s) from {input_path}")
    for _, source in inputs.iterrows():
        university = str(source.get("University", "")).strip()
        start_url = str(source.get("Start URL", "")).strip()
        report(f"Starting {university or '(missing university name)'}")
        start_time = datetime.now(timezone.utc)
        status = "Completed"
        warnings: List[str] = []
        error_parts: List[str] = []
        candidates: List[PageResult] = []
        pages_visited = 0
        before_count = len(programme_rows)
        try:
            if not university or not start_url:
                raise ValueError("University and Start URL are required for each input row")
            candidates, pages_visited, crawl_error = crawl_university(university, start_url)
            if crawl_error:
                error_parts.append(crawl_error)
            for page in candidates:
                candidate_rows.append({
                    "University": university,
                    "URL": page.url,
                    "Page Title": page.title,
                    "Candidate Score": page.score,
                    "Matched Keywords": page.found_keywords,
                    "Candidate Type": page.candidate_type,
                    "Reason": page.reason,
                    "Needs Review": page.needs_review,
                    "PDF Links": page.pdf_links,
                    "Notes": page.notes,
                })
            extraction_pages = [p for p in candidates if p.candidate_type == "HTML" and p.score >= RELEVANCE_THRESHOLD]
            if not extraction_pages:
                extraction_pages = [p for p in candidates if p.candidate_type == "HTML"][:5]
            if client:
                for page in extraction_pages:
                    report(f"Extracting programme data from {page.url}")
                    rows, extract_error = extract_programmes_with_openai(client, model, university, page)
                    programme_rows.extend(rows)
                    if extract_error:
                        error_parts.append(f"{page.url}: {extract_error}")
            else:
                warnings.append("OPENAI_API_KEY not set; fallback extraction used")
            extracted_for_university = len(programme_rows) - before_count
            if candidates and extracted_for_university == 0:
                fallback_rows = fallback_extract_programmes(university, candidates)
                programme_rows.extend(fallback_rows)
                extracted_for_university = len(fallback_rows)
                warnings.append("No programmes extracted. Fallback extraction used or review needed.")
            if candidates and extracted_for_university == 0:
                status = "Completed with warnings"
            elif warnings:
                status = "Completed with warnings"
        except Exception as exc:
            status = "Failed"
            error_parts.append(str(exc))
            extracted_for_university = len(programme_rows) - before_count
        end_time = datetime.now(timezone.utc)
        log_rows.append({
            "University": university,
            "Start URL": start_url,
            "Status": status,
            "Start Time": start_time.strftime("%Y-%m-%d %H:%M:%S UTC"),
            "End Time": end_time.strftime("%Y-%m-%d %H:%M:%S UTC"),
            "Pages Visited": pages_visited,
            "Candidate Pages Found": len(candidates),
            "Programmes Extracted": extracted_for_university,
            "Warnings": " | ".join(warnings),
            "Error Details": " | ".join(error_parts),
        })
        report(f"Finished {university}: {status}; {len(candidates)} candidate page(s); {extracted_for_university} programme row(s)")

    write_excel_output(output_path, candidate_rows, programme_rows, log_rows)
    report(f"Wrote {output_path}")
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Crawl university sites and extract inbound short-term programme information.")
    parser.add_argument("--input", default=os.getenv("INPUT_FILE", INPUT_FILE), help="Path to input_urls.xlsx")
    parser.add_argument("--output", default=os.getenv("OUTPUT_FILE", OUTPUT_FILE), help="Path for output_programmes.xlsx")
    parser.add_argument("--model", default=os.getenv("OPENAI_MODEL", "gpt-4o-mini"), help="OpenAI model name")
    parser.add_argument("--api-key", default=os.getenv("OPENAI_API_KEY"), help="OpenAI API key. Defaults to OPENAI_API_KEY.")
    return parser.parse_args()


def main() -> None:
    load_dotenv()
    args = parse_args()
    run_scraper(input_path=args.input, output_path=args.output, api_key=args.api_key, model=args.model)


if __name__ == "__main__":
    main()
