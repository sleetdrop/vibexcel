#!/usr/bin/env python3
"""Check or remove private path and Office Web Extension XLSX residue."""

from __future__ import annotations

import argparse
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


CONTENT_TYPES = "[Content_Types].xml"
WEB_EXTENSION_PREFIX = "xl/webextensions/"
WORKBOOK_XML = "xl/workbook.xml"
ABS_PATH_ALT_CONTENT = re.compile(
    rb"<(?:[A-Za-z0-9_]+:)?AlternateContent\b[^>]*>\s*"
    rb"<(?:[A-Za-z0-9_]+:)?Choice\b[^>]*>\s*"
    rb"<(?:[A-Za-z0-9_]+:)?absPath\b[^>]*/>\s*"
    rb"</(?:[A-Za-z0-9_]+:)?Choice>\s*"
    rb"</(?:[A-Za-z0-9_]+:)?AlternateContent>",
    re.IGNORECASE | re.DOTALL,
)


class PackageError(RuntimeError):
    pass


def _lower(value: str | None) -> str:
    return (value or "").lower()


def _web_extension_relationship(element: ET.Element) -> bool:
    return "webextension" in _lower(element.get("Type")) or "webextension" in _lower(
        element.get("Target")
    )


def _web_extension_content_type(element: ET.Element) -> bool:
    return "webextension" in _lower(element.get("PartName")) or "webextension" in _lower(
        element.get("ContentType")
    )


def _filtered_xml(data: bytes, predicate) -> tuple[bytes, int]:
    root = ET.fromstring(data)
    namespace = root.tag.partition("}")[0].removeprefix("{")
    if namespace:
        ET.register_namespace("", namespace)

    removed = 0
    for child in list(root):
        if predicate(child):
            root.remove(child)
            removed += 1

    if not removed:
        return data, 0
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), removed


def _strip_absolute_path_metadata(data: bytes) -> tuple[bytes, int]:
    return ABS_PATH_ALT_CONTENT.subn(b"", data)


def find_residue(workbook: Path) -> list[str]:
    try:
        with zipfile.ZipFile(workbook) as archive:
            bad_member = archive.testzip()
            if bad_member:
                raise PackageError(f"corrupt ZIP member: {bad_member}")

            findings: list[str] = []
            for name in archive.namelist():
                if name.lower().startswith(WEB_EXTENSION_PREFIX):
                    findings.append(name)
                    continue
                if name == CONTENT_TYPES or name.lower().endswith(".rels"):
                    data = archive.read(name).lower()
                    if b"webextension" in data:
                        findings.append(f"{name} (webextension reference)")
                elif name == WORKBOOK_XML:
                    data = archive.read(name)
                    if ABS_PATH_ALT_CONTENT.search(data):
                        findings.append(f"{name} (absolute path metadata)")
            return findings
    except (OSError, zipfile.BadZipFile) as error:
        raise PackageError(str(error)) from error


def sanitize(workbook: Path) -> tuple[int, int, int]:
    findings = find_residue(workbook)
    if not findings:
        return 0, 0, 0

    removed_parts = 0
    removed_references = 0
    removed_paths = 0
    temporary_path: Path | None = None

    try:
        with zipfile.ZipFile(workbook, "r") as source:
            file_descriptor, temporary_name = tempfile.mkstemp(
                prefix=f".{workbook.name}.", suffix=".tmp", dir=workbook.parent
            )
            os.close(file_descriptor)
            temporary_path = Path(temporary_name)

            with zipfile.ZipFile(temporary_path, "w") as target:
                for info in source.infolist():
                    name = info.filename
                    if name.lower().startswith(WEB_EXTENSION_PREFIX):
                        removed_parts += 1
                        continue

                    data = source.read(name)
                    if name == CONTENT_TYPES:
                        data, removed = _filtered_xml(
                            data, _web_extension_content_type
                        )
                        removed_references += removed
                    elif name.lower().endswith(".rels"):
                        data, removed = _filtered_xml(
                            data, _web_extension_relationship
                        )
                        removed_references += removed
                    elif name == WORKBOOK_XML:
                        data, removed = _strip_absolute_path_metadata(data)
                        removed_paths += removed
                    target.writestr(info, data)

        with zipfile.ZipFile(temporary_path) as result:
            bad_member = result.testzip()
            if bad_member:
                raise PackageError(f"sanitized ZIP member is corrupt: {bad_member}")

        if find_residue(temporary_path):
            raise PackageError("sanitized package still contains package hygiene residue")

        os.replace(temporary_path, workbook)
        temporary_path = None
        return removed_parts, removed_references, removed_paths
    except (OSError, zipfile.BadZipFile, ET.ParseError) as error:
        raise PackageError(str(error)) from error
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check", action="store_true")
    mode.add_argument("--sanitize", action="store_true")
    parser.add_argument("workbook", type=Path)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    workbook = args.workbook.resolve()

    if not workbook.is_file():
        print(f"Workbook not found: {workbook}", file=sys.stderr)
        return 2

    try:
        if args.check:
            findings = find_residue(workbook)
            if findings:
                print("XLSX package hygiene residue found:", file=sys.stderr)
                for finding in findings:
                    print(f"- {finding}", file=sys.stderr)
                return 1
            print(f"Clean XLSX package: {workbook}")
            return 0

        removed_parts, removed_references, removed_paths = sanitize(workbook)
        print(
            f"Sanitized {workbook}: removed {removed_parts} parts, "
            f"{removed_references} references, and {removed_paths} absolute paths"
        )
        return 0
    except PackageError as error:
        print(f"Could not process {workbook}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
