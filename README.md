# Generic Programme Candidate Page Scraper

Phase 1 collects candidate pages for short-term, inbound, summer, exchange, Japanese language, and customized/customised programmes from multiple university websites. It writes results for human review to `output_programmes.xlsx`.

## Prepare `target_universities.csv`

Create a CSV named `target_universities.csv` with these columns:

| University Name | Country | Seed URL | Allowed Domain | Notes |
| --- | --- | --- | --- | --- |
| Example University | Japan | https://www.example.ac.jp/en/ | example.ac.jp | Optional notes |

- Add one university per row.
- `Seed URL` is where crawling starts.
- `Allowed Domain` keeps the crawler inside the university website, including subdomains.

## Run

```bash
pip install -r requirements.txt
python programme_scraper.py --input target_universities.csv --output output_programmes.xlsx
```

## Output workbook

`output_programmes.xlsx` contains:

1. **Candidate Pages** - candidate URL, page title, score, matched keywords, candidate type, reason, and review flag.
2. **Run Log** - start/end time, universities processed, pages visited, candidate count, status, warnings, and error details.

This phase is keyword-based candidate collection only. It does not perform final programme extraction, add a GUI, build Windows packages, configure GitHub Actions, or require the OpenAI API.
