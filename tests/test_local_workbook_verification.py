"""Offline verification for the generic programme scraper workbook output."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

import programme_scraper as scraper


DETAIL_HTML = """
<html><head><title>Summer Programme</title></head><body>
<h1>Summer Programme</h1>
<p>Programme Dates: July 1 - July 20, 2026.</p>
<p>Application Deadline: April 30, 2026.</p>
<p>Programme Fee: JPY 120,000.</p>
<p>Housing: accommodation is available in university residence.</p>
<p>Eligibility: Undergraduate students in good academic standing.</p>
<p>Language: English.</p>
</body></html>
"""


class LocalWorkbookVerificationTest(unittest.TestCase):
    def setUp(self) -> None:
        self.target = scraper.UniversityTarget(
            name="Local Example University",
            country="Testland",
            seed_url="https://example.test/",
            allowed_domain="example.test",
            notes="offline test",
        )
        self.candidate = scraper.CandidatePage(
            university_name=self.target.name,
            country=self.target.country,
            url="https://example.test/summer",
            title="Summer Programme",
            score=75,
            matched_keywords=["summer programme", "application deadline", "programme fee", "housing", "eligibility"],
            candidate_type="Strong programme candidate",
            reason="Offline fixture",
        )

    def test_excel_output_extraction_duplicates_and_run_log(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "output_programmes.xlsx"
            input_path = Path(tmpdir) / "targets.csv"
            input_path.write_text("unused by patched read_targets\n", encoding="utf-8")

            with patch.object(scraper, "read_targets", return_value=[self.target]), \
                 patch.object(scraper, "crawl_university", return_value=([self.candidate, self.candidate], 3, ["sample warning"])), \
                 patch.object(scraper, "get_robot_parser", return_value=object()), \
                 patch.object(scraper, "fetch_html", return_value=DETAIL_HTML), \
                 patch.object(scraper.time, "sleep", return_value=None):
                scraper.run_scraper(input_path=str(input_path), output_path=str(output_path))

            workbook = pd.ExcelFile(output_path)
            self.assertEqual(["Candidate Pages", "Extracted Programmes", "Run Log"], workbook.sheet_names)

            programmes = pd.read_excel(output_path, sheet_name="Extracted Programmes")
            self.assertEqual(len(programmes), 2)
            self.assertTrue(programmes["University Name"].notna().all())
            self.assertTrue(programmes["Source URL"].notna().all())
            self.assertIn("Duplicate", set(programmes["Duplicate Status"]))
            self.assertEqual(set(programmes["Extraction Method"]), {"heuristic"})
            self.assertTrue((programmes["Review Status"].isin(["Likely valid", "Needs human review", "Low confidence"])).all())

            run_log = pd.read_excel(output_path, sheet_name="Run Log").iloc[0]
            self.assertEqual(run_log["Total Pages Visited"], 3)
            self.assertEqual(run_log["Total Candidate Pages"], 2)
            self.assertEqual(run_log["Total Programme Rows"], 2)
            self.assertEqual(run_log["Duplicate Rows"], 1)
            self.assertIn("sample warning", run_log["Warnings"])

    def test_excel_output_fallback_rows_are_marked(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "fallback_output.xlsx"
            input_path = Path(tmpdir) / "targets.csv"
            input_path.write_text("unused by patched read_targets\n", encoding="utf-8")

            with patch.object(scraper, "read_targets", return_value=[self.target]), \
                 patch.object(scraper, "crawl_university", return_value=([self.candidate], 2, [])), \
                 patch.object(scraper, "get_robot_parser", return_value=object()), \
                 patch.object(scraper, "fetch_html", return_value=None), \
                 patch.object(scraper.time, "sleep", return_value=None):
                scraper.run_scraper(input_path=str(input_path), output_path=str(output_path))

            programmes = pd.read_excel(output_path, sheet_name="Extracted Programmes")
            self.assertEqual(len(programmes), 1)
            self.assertEqual(programmes.loc[0, "Extraction Method"], "fallback")
            self.assertEqual(programmes.loc[0, "Review Status"], "Needs human review")
            self.assertIn("Programme Dates", programmes.loc[0, "Missing Fields"])


if __name__ == "__main__":
    unittest.main()
