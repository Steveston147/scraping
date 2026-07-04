# Generic Programme Candidate Page Scraper

The scraper collects candidate pages for short-term, inbound, summer, exchange, Japanese language, and customized/customised programmes from multiple university websites, then performs generic heuristic programme extraction for human review. It writes results to `output_programmes.xlsx`.

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
2. **Extracted Programmes** - one row per extracted or fallback programme candidate, including source URL, review status, confidence score, missing fields, duplicate group/status, last checked date, extraction method, and notes. Missing values are left blank or marked `Unknown`; fees, dates, deadlines, eligibility, and housing are not invented. Rows are formatted to make likely-valid, needs-review, low-confidence, and duplicate records easier to scan.
3. **Run Log** - start/end time, universities processed, universities with candidates, universities with extracted programmes, total pages visited, total candidate pages, candidate pages read, total programme rows, fallback rows, duplicate rows, status, warnings, and errors.

If a university has candidate pages but no programme rows can be extracted, the scraper writes the top 3 to 5 candidate pages into **Extracted Programmes** as fallback rows marked `Needs human review` with `Extraction Method` set to `fallback`. Heuristic extraction rows use `heuristic`; the schema also leaves room for future optional AI or manual-candidate methods without making the OpenAI API required. Duplicate rows with the same or very similar programme name and source URL are assigned a duplicate group and status. The scraper remains generic and does not require the OpenAI API.
