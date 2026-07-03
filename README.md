# Inbound Short-Term Programme Scraper MVP

This MVP reads a list of university start URLs from `input_urls.xlsx`, politely crawls public pages, scores likely inbound short-term programme pages, asks the OpenAI API to extract programme information, and writes `output_programmes.xlsx`.

The tool is intentionally simple and Windows-friendly. It does not extract PDF text. When PDF links are found, it records the PDF URL and marks the page with `PDF review needed`.

## 1. Install dependencies

Install Python 3.10 or newer. In Windows PowerShell, from this project folder, run:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

On macOS or Linux, activate the virtual environment with:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 2. Configure the OpenAI API key

Create a file named `.env` in this folder:

```text
OPENAI_API_KEY=your_api_key_here
```

Optional settings:

```text
OPENAI_MODEL=gpt-4o-mini
INPUT_FILE=input_urls.xlsx
OUTPUT_FILE=output_programmes.xlsx
```

Do not put real API keys into the code or commit them to version control.

## 3. Prepare `input_urls.xlsx`

Create an Excel file named `input_urls.xlsx` with exactly these columns:

| University | Start URL | Notes |
| --- | --- | --- |
| Example University | https://www.example.ac.jp/en/ | Optional notes |

One row equals one university crawl. Start with the university's English international office or international programmes page when possible.

## 4. Run the tool

With the virtual environment activated, run:

```bash
python programme_scraper.py
```

The crawler uses:

- depth limit `2`,
- maximum `50` pages per university,
- a polite delay between requests,
- `robots.txt` checks where possible,
- same or clearly related university domains only.

It avoids common login and contact pages and does not intentionally collect personal data.

## 5. Output

The script creates `output_programmes.xlsx` with three sheets:

1. **Candidate Pages** - likely relevant pages, relevance scores, keywords, reasons, PDF links, and review notes.
2. **Extracted Programmes** - OpenAI-extracted programme fields. Every row has `Check Status` set to `Needs human review`.
3. **Run Log** - status, pages visited, candidates found, extracted programme count, errors, and run datetime.

If one university URL fails, the error is recorded in **Run Log** and the script continues with the next university.

## 6. Known MVP limitations

- No PDF text extraction yet; PDF URLs are only recorded for human review.
- Relevance scoring is keyword-based and may miss pages with unusual wording.
- Domain matching is practical rather than perfect; complex university domain setups may need manual start URLs.
- OpenAI extraction depends on page text quality and may require human verification.
- Missing or unclear fields are left blank. The tool should not guess.
- Dynamic pages that require JavaScript may not provide enough readable HTML text.
- Login pages, forms, and personal-data collection are intentionally out of scope.
