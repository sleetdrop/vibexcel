import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "tools" / "sanitize_xlsx_package.py"

CONTENT_TYPES_CLEAN = b"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
</Types>
"""

CONTENT_TYPES_DIRTY = CONTENT_TYPES_CLEAN.replace(
    b"</Types>",
    b"""  <Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>
  <Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>
</Types>""",
)

ROOT_RELS_CLEAN = b"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

ROOT_RELS_DIRTY = ROOT_RELS_CLEAN.replace(
    b"</Relationships>",
    b"""  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="xl/webextensions/taskpanes.xml"/>
</Relationships>""",
)

WORKBOOK_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>
"""

ABS_PATH_BLOCK = b"""<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"><x15ac:absPath url="/Users/example/private/workbook/" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent>"""

WORKBOOK_XML_WITH_ABS_PATH = WORKBOOK_XML.replace(
    b"<sheets/>", b"<sheets/>" + ABS_PATH_BLOCK
)

SHEET_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><legacyDrawing r:id="rId2"/></worksheet>
"""

SHEET_RELS = b"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>
"""

COMMENTS_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Yuan Jiang</author></authors><commentList><comment ref="P88" authorId="0"><text><t>Keep this native note.</t></text></comment></commentList></comments>
"""


def write_fixture(path: Path, *, dirty: bool, abs_path: bool = False) -> None:
    members = {
        "[Content_Types].xml": CONTENT_TYPES_DIRTY if dirty else CONTENT_TYPES_CLEAN,
        "_rels/.rels": ROOT_RELS_DIRTY if dirty else ROOT_RELS_CLEAN,
        "xl/workbook.xml": WORKBOOK_XML_WITH_ABS_PATH if abs_path else WORKBOOK_XML,
        "xl/worksheets/sheet1.xml": SHEET_XML,
        "xl/worksheets/_rels/sheet1.xml.rels": SHEET_RELS,
        "xl/comments1.xml": COMMENTS_XML,
        "xl/drawings/vmlDrawing1.vml": b"<xml>native comment drawing</xml>",
    }
    if dirty:
        members.update(
            {
                "xl/webextensions/taskpanes.xml": b"<wetp:taskpanes xmlns:wetp='http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11'/>",
                "xl/webextensions/_rels/taskpanes.xml.rels": b"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'/>",
                "xl/webextensions/webextension1.xml": b"<we:webextension xmlns:we='http://schemas.microsoft.com/office/webextensions/webextension/2010/11'/>",
            }
        )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in members.items():
            archive.writestr(name, data)


def run_cli(mode: str, workbook: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(SCRIPT), mode, str(workbook)],
        text=True,
        capture_output=True,
        check=False,
    )


class SanitizeXlsxPackageTest(unittest.TestCase):
    def test_check_rejects_office_web_extension_residue(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "dirty.xlsx"
            write_fixture(workbook, dirty=True)

            result = run_cli("--check", workbook)

            self.assertEqual(result.returncode, 1)
            self.assertIn("package hygiene residue", result.stderr)
            self.assertIn("xl/webextensions/taskpanes.xml", result.stderr)

    def test_check_and_sanitize_absolute_path_metadata(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "private-path.xlsx"
            write_fixture(workbook, dirty=False, abs_path=True)

            check_result = run_cli("--check", workbook)

            self.assertEqual(check_result.returncode, 1)
            self.assertIn("absolute path metadata", check_result.stderr)

            sanitize_result = run_cli("--sanitize", workbook)

            self.assertEqual(sanitize_result.returncode, 0, sanitize_result.stderr)
            self.assertEqual(run_cli("--check", workbook).returncode, 0)
            with zipfile.ZipFile(workbook) as archive:
                self.assertEqual(archive.read("xl/workbook.xml"), WORKBOOK_XML)

    def test_check_accepts_a_clean_package(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "clean.xlsx"
            write_fixture(workbook, dirty=False)

            result = run_cli("--check", workbook)

            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn("clean", result.stdout.lower())

    def test_sanitize_removes_only_web_extension_parts_and_references(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "dirty.xlsx"
            write_fixture(workbook, dirty=True)

            result = run_cli("--sanitize", workbook)

            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(run_cli("--check", workbook).returncode, 0)
            with zipfile.ZipFile(workbook) as archive:
                names = set(archive.namelist())
                self.assertFalse(any(name.startswith("xl/webextensions/") for name in names))
                self.assertNotIn(b"webextension", archive.read("[Content_Types].xml").lower())
                self.assertNotIn(b"webextension", archive.read("_rels/.rels").lower())
                self.assertEqual(archive.read("xl/workbook.xml"), WORKBOOK_XML)
                self.assertEqual(archive.read("xl/comments1.xml"), COMMENTS_XML)
                self.assertEqual(archive.read("xl/worksheets/sheet1.xml"), SHEET_XML)
                self.assertEqual(archive.testzip(), None)


if __name__ == "__main__":
    unittest.main()
