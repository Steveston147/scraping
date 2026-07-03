# Inbound Programme Scraper MVP - Windows App

This MVP helps find inbound short-term programme pages on university websites and creates an Excel file for human review.

It focuses on public pages about short-term programmes in Japan for international students or overseas partner universities. It does **not** extract PDF text yet. If it finds PDF links, it records the PDF URL and marks it as `PDF review needed`.

Both options are supported: daily Windows use through the desktop app, and command-line/script use for technical users.

## Quick start for Windows users

### 1. Download the Windows app

1. Open the GitHub repository page.
2. Go to **Actions**.
3. Open the latest successful **Build Windows executable** run.
4. Download the artifact named **InboundProgrammeScraper-Windows**.
5. Unzip it.
6. Double-click `InboundProgrammeScraper.exe`.

You do not need to install Python when using the downloaded `.exe` file.

## Prepare `input_urls.xlsx`

Create an Excel file with exactly these columns:

| University | Start URL | Notes |
| --- | --- | --- |
| Example University | https://www.example.ac.jp/en/ | Optional notes |

Tips:

- One row equals one university.
- Use a public English international office or international programmes page when possible.
- Do not use login pages.

## Run the app

1. Double-click `InboundProgrammeScraper.exe`.
2. Click **Browse...** next to the input file and select `input_urls.xlsx`.
3. Click **Browse...** next to the output folder and choose where to save the result.
4. Paste your OpenAI API key into the API key box.
5. Leave the model as `gpt-4o-mini` unless you know you want another model.
6. Click **Run**.
7. Watch the progress messages.
8. When the run finishes, the app shows a completion message.

The API key is not included in the code, repository, or Windows build. The app reads it from what you paste into the app, or from a local `.env` file if you run from source.

## Output file

The app creates this file in the output folder you selected:

```text
output_programmes.xlsx
```

The workbook contains three sheets:

1. **Candidate Pages** - pages that may be relevant, with scores, keywords, reasons, PDF links, and notes.
2. **Extracted Programmes** - programme information extracted with the OpenAI API. Every row is marked `Needs human review`.
3. **Run Log** - status, pages visited, candidates found, extracted programme count, errors, and run datetime.

If one university URL fails, the tool writes the error to **Run Log** and continues with the next row.

## Running from source instead of the Windows app

If you prefer to run from source, install Python 3.10 or newer, then run:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

The command-line script still works too:

```powershell
python programme_scraper.py
```

For command-line use, create a local `.env` file:

```text
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL=gpt-4o-mini
INPUT_FILE=input_urls.xlsx
OUTPUT_FILE=output_programmes.xlsx
```

Do not commit real API keys to version control.

## Build the Windows executable manually

GitHub Actions builds the Windows `.exe` automatically. To build it manually on Windows:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name InboundProgrammeScraper app.py
```

The executable will be created here:

```text
dist\InboundProgrammeScraper.exe
```

## What the crawler does

- Reads `input_urls.xlsx`.
- Crawls only public HTML pages on the same or clearly related university domain.
- Uses depth limit `2`.
- Uses maximum `50` pages per university.
- Adds a polite delay between requests.
- Checks `robots.txt` where possible.
- Avoids common login/contact pages and does not submit forms.
- Does not intentionally collect personal data.
- Finds candidate pages with keyword scoring.
- Sends only relevant candidate page text to the OpenAI extraction step.

## Known MVP limitations

- PDF text extraction is not implemented yet. PDF URLs are recorded for human review only.
- Relevance scoring is keyword-based and may miss pages with unusual wording.
- Dynamic pages that require JavaScript may not provide enough readable text.
- Domain matching is practical rather than perfect; some universities may require a better start URL.
- OpenAI extraction depends on source page quality and must be checked by a human.
- Missing or unclear fields are left blank. The tool should not guess.
- No web login, database, scheduling, or advanced crawling features are included.
