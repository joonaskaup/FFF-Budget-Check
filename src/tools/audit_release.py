"""Inspect the final legacy XLS release without saving it or running VBA."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import struct

import olefile
from oletools.olevba import VBA_Parser, decompress_stream
import xlrd

ROOT = Path(__file__).resolve().parents[2]


def inspect(path):
    checks = []

    def check(condition, message):
        if not condition:
            raise ValueError(message)
        checks.append(message)

    with olefile.OleFileIO(path) as ole:
        streams = ole.listdir()
        allowed = {
            ("\x01CompObj",), ("\x05DocumentSummaryInformation",),
            ("\x05SummaryInformation",), ("Workbook",),
            ("_VBA_PROJECT_CUR", "PROJECT"), ("_VBA_PROJECT_CUR", "PROJECTwm"),
            ("_VBA_PROJECT_CUR", "VBA", "dir"),
            ("_VBA_PROJECT_CUR", "VBA", "_VBA_PROJECT"),
            ("_VBA_PROJECT_CUR", "VBA", "Sheet1"),
            ("_VBA_PROJECT_CUR", "VBA", "ThisWorkbook"),
        }
        check(all(tuple(name) in allowed for name in streams),
              "Only workbook, standard metadata and expected worksheet VBA streams")
        metadata = ole.get_metadata()
        personal = ("author", "last_saved_by", "company", "manager", "subject", "keywords", "comments")
        check(not any(getattr(metadata, field, None) for field in personal),
              "Personal document properties are empty")
        plain_dir = decompress_stream(ole.openstream("_VBA_PROJECT_CUR/VBA/dir").read())
        check(b"/Users/ExcelUser" in plain_dir, "Mac Forms library uses a generic username")
        stripped_dir = re.sub(rb"/Users/ExcelUser_*/", b"", plain_dir)
        check(b"/Users/" not in stripped_dir, "No personal Mac folder in VBA references")
        for name in streams:
            raw = ole.openstream(name).read()
            for encoding in ("utf-8", "utf-16-le"):
                placeholder = re.escape("/Users/ExcelUser".encode(encoding)) + b"(?:" + re.escape("_".encode(encoding)) + b")*" + re.escape("/".encode(encoding))
                raw = re.sub(placeholder, b"", raw)
                # Compressed directory references were inspected above.
                if name[-1] != "dir":
                    check("/Users/".encode(encoding) not in raw,
                          "No personal folder in " + "/".join(name) + " (" + encoding + ")")
        workbook_bytes = ole.openstream("Workbook").read()
        records = []
        offset = 0
        while offset + 4 <= len(workbook_bytes):
            record, size = struct.unpack_from("<HH", workbook_bytes, offset)
            records.append((record, workbook_bytes[offset + 4:offset + 4 + size]))
            offset += 4 + size
        check(not any(record == 0x0006 for record, _ in records), "Clean donor has no cell formulas")
        check(not any(record == 0x0018 for record, _ in records), "No inherited defined-name records")
        check(not any(record == 0x01B8 for record, _ in records), "No inherited hyperlinks")
        check(not any(record == 0x0023 for record, _ in records), "No external-name records")
        check(all(len(data) >= 4 and data[2:4] == b"\x01\x04" for record, data in records if record == 0x01AE),
              "No external workbook or add-in link records")

    book = xlrd.open_workbook(path, formatting_info=True)
    check(book.sheet_names() == ["Budget Check"], "Exactly one Budget Check worksheet")
    sheet = book.sheet_by_index(0)
    check(sheet.visibility == 0, "Checker worksheet is visible")
    check(not book.name_map, "No inherited workbook names")
    check(sheet.cell_value(3, 0) == "Ready - copy this tab into your budget to start.", "Starts ready with no prior report")
    check(sheet.cell_value(5, 1) == 0.6, "Default fringe review threshold is 60%")
    check(all(sheet.cell_value(r, c) == "" for r in range(12, sheet.nrows) for c in range(sheet.ncols)),
          "Report contains no budget findings or test data")
    check(sheet.cell_value(4, 1) == "" and sheet.cell_value(10, 0) == "", "No saved budget selection or report summary")
    check(not sheet.cell_note_map, "No cell comments")
    cells = [{"row": r + 1, "column": c + 1, "value": sheet.cell_value(r, c),
              "type": sheet.cell_type(r, c), "style": sheet.cell_xf_index(r, c)}
             for r in range(sheet.nrows) for c in range(sheet.ncols) if sheet.cell_value(r, c) != ""]

    parser = VBA_Parser(str(path))
    try:
        modules = list(parser.extract_macros())
    finally:
        parser.close()
    executable = []
    for _, stream, name, code in modules:
        cleaned = "\n".join(line for line in code.splitlines() if not line.startswith("Attribute ")).strip()
        if cleaned:
            executable.append((stream, name, cleaned))
    check(len(executable) == 1 and executable[0][1] == "Sheet1.cls", "Executable VBA lives only in the checker worksheet")
    expected = (ROOT / "src/vba/BudgetCheckWorksheet.vba").read_text().strip()
    check(executable[0][2] == expected, "Embedded worksheet VBA exactly matches published source")

    inventory = {"sheets": [{"name": sheet.name, "cells": cells,
                             "merged_ranges": sheet.merged_cells,
                             "column_widths": {str(n + 1): col.width for n, col in sheet.colinfo_map.items()}}],
                 "defined_names": [], "vba_modules": [name for _, _, name, _ in modules]}
    (ROOT / "docs/workbook-inventory.json").write_text(json.dumps(inventory, indent=2, ensure_ascii=False) + "\n")
    (ROOT / "docs/RELEASE_AUDIT.md").write_text(
        "# Release workbook inspection\n\nFile: `" + path.name + "`\n\n" +
        "\n".join("- " + message for message in checks) +
        "\n\nNative Excel execution is recorded in TEST_REPORT.md.\n")
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    manifest = {"version": "1.0.0", "asset": path.name, "sha256": digest,
                "bytes": path.stat().st_size, "source": "src/vba/BudgetCheckWorksheet.vba"}
    (ROOT / "release/manifest.json").write_text(json.dumps(manifest, indent=2) + "\n")
    (ROOT / "release/SHA256SUMS").write_text(digest + "  " + path.name + "\n")
    print(json.dumps({"checks_passed": len(checks), **manifest}, indent=2))


if __name__ == "__main__":
    arguments = argparse.ArgumentParser(description=__doc__)
    arguments.add_argument("workbook", type=Path)
    inspect(arguments.parse_args().workbook)
