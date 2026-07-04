# Inbound Short-Term Programme Scraper MVP

This Windows-friendly MVP reads university start URLs from `input_urls.xlsx`, politely crawls public pages, scores likely inbound short-term programme pages, asks the OpenAI API to extract programme information, and writes `output_programmes.xlsx`.

The tool supports both:

- a **Windows GUI app** (`app.py` or the packaged `InboundProgrammeScraper.exe`), and
- **command-line/script usage** (`programme_scraper.py`).

All extracted results are marked **Needs human review**. The scraper does not guess missing details, and it does **not** extract PDF text yet. If PDF links are found, the workbook records the PDF URLs for manual review.

## Prepare `input_urls.xlsx`

Create an Excel file named `input_urls.xlsx` with exactly these columns:

| University | Start URL | Notes |
| --- | --- | --- |
| Example University | https://www.example.ac.jp/en/ | Optional notes |

Tips for non-engineers:

1. Open Microsoft Excel.
2. Put `University`, `Start URL`, and `Notes` in the first row.
3. Add one university per row.
4. Use the university's English international office, inbound programmes, summer school, or international exchange page as the start URL when possible.
5. Save the file as `input_urls.xlsx`.

## Option A: Run the Windows app from GitHub Actions

1. Open this repository on GitHub.
2. Go to the **Actions** tab.
3. Open the latest successful **Build Windows executable** workflow run.
4. Download the artifact named **InboundProgrammeScraper-Windows**.
5. Unzip the downloaded artifact.
6. Double-click `InboundProgrammeScraper.exe`.
7. In the app:
   - choose your `input_urls.xlsx` file,
   - choose an output folder,
   - enter your OpenAI API key,
   - keep the default model `gpt-4o-mini` unless you know you need another model,
   - click **Run**.

When the run finishes, the app shows a completion message and saves `output_programmes.xlsx` in the output folder you selected. If an error happens, the app shows an error message and writes progress messages in the window.

## Option B: Run the GUI with Python

Install Python 3.10 or newer. In Windows PowerShell, from this project folder, run:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

## Option C: Run from the command line

Create a file named `.env` in this folder:

```text
OPENAI_API_KEY=your_api_key_here
```

Optional `.env` settings:

```text
OPENAI_MODEL=gpt-4o-mini
INPUT_FILE=input_urls.xlsx
OUTPUT_FILE=output_programmes.xlsx
```

Run with defaults:

```bash
python programme_scraper.py
```

Or pass paths and model explicitly:

```bash
python programme_scraper.py --input input_urls.xlsx --output output_programmes.xlsx --model gpt-4o-mini
```

Do not put real API keys into the code or commit them to version control.

## What the crawler does

The crawler uses:

- depth limit `2`,
- maximum `50` pages per university,
- a polite delay between requests,
- `robots.txt` checks where possible,
- same or clearly related university domains only.

It avoids common login, sign-in, contact, inquiry, and form pages and does not intentionally collect personal data.

If `robots.txt` can be read, the crawler respects it. If `robots.txt` cannot be read because of a network error, timeout, URL error, or server-side 5xx response, the crawler does not block every page automatically; it continues politely under the strict crawler limits above.

## What `output_programmes.xlsx` contains

The script creates `output_programmes.xlsx` with three sheets:

1. **Candidate Pages** - likely relevant pages with URL, page title, candidate score, matched keywords, candidate type (`HTML` or `PDF`), reason, review flag, PDF links, and notes.
2. **Extracted Programmes** - extracted or fallback programme rows with programme name, type, institution/department, target students, language, dates, fee, housing, credits/certificate, contents, eligibility, source URL/title, confidence score, missing fields, review status, and notes. Incomplete but relevant rows are kept and marked for review instead of being discarded.
3. **Run Log** - start/end time, pages visited, candidates found, programme row count, status, warnings, and error details.

If one university URL fails, the error is recorded in **Run Log** and the script continues with the next university. If candidates exist but no programme rows can be extracted, the run is marked `Completed with warnings` and fallback rows are created for human review.

## Build `InboundProgrammeScraper.exe` locally

GitHub Actions builds the Windows executable automatically for pushes to `main`, pull requests targeting `main`, and manual `workflow_dispatch` runs. To build locally on Windows:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name InboundProgrammeScraper app.py
```

The executable will be created at `dist\InboundProgrammeScraper.exe`.

## Known MVP limitations

- No PDF text extraction yet; PDF URLs are only recorded for human review.
- Every extracted result needs human review before it is used.
- Relevance scoring is keyword-based and may miss pages with unusual wording.
- Domain matching is practical rather than perfect; complex university domain setups may need manual start URLs.
- OpenAI extraction depends on page text quality and may require human verification.
- Missing or unclear fields are left blank. The tool should not guess.
- Dynamic pages that require JavaScript may not provide enough readable HTML text.
- Login pages, forms, and personal-data collection are intentionally out of scope.
