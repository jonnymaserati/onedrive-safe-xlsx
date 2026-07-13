# Copyright 2026 Jonathan Corcutt
# SPDX-License-Identifier: Apache-2.0
"""
onedrive_safe_xlsx — edit .xlsx files with openpyxl WITHOUT breaking OneDrive/SharePoint.

Problem
-------
openpyxl rebuilds a workbook from its own minimal model on save. A real Excel file has
~26 OOXML parts; openpyxl writes ~9 — it silently DROPS customXml/*, docProps/custom.xml,
printerSettings, calcChain, etc. On OneDrive/SharePoint those parts carry the document's
identity and library binding, so an openpyxl-saved file can lose its version lineage
(edits look like a foreign file, not a new version of the same doc).

And even if you preserve those parts, if you copy the original `docProps/core.xml`
unchanged, its `dcterms:modified` / `cp:revision` stay the same — SharePoint reads that
as "nothing changed" and won't log a new version.

Fix
---
`save_preserving()` edits the data via openpyxl but (a) keeps every non-data part from the
ORIGINAL file (identity preserved), and (b) updates the change-tracking in core.xml
(modified -> now, lastModifiedBy, revision bump) so SharePoint records a genuine new version.

    import openpyxl
    from onedrive_safe_xlsx import save_preserving

    wb = openpyxl.load_workbook("Report.xlsx")   # MUST be an Excel/SharePoint-saved file
    wb.active["A1"] = "edited"
    save_preserving("Report.xlsx", wb, modified_by="Jane Doe")

Caveat: start from an Excel/SharePoint-saved copy (it has the parts to preserve). A file
that was only ever written by openpyxl has nothing to preserve. Intended for sequential
(not concurrent) edits.

Licensed under the Apache License, Version 2.0.
"""
import openpyxl, zipfile, re, os, tempfile
from datetime import datetime, timezone


def _now_utc() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _bump_core(xml: bytes, modified_by, now: str) -> bytes:
    s = xml.decode("utf-8")
    if "<dcterms:modified" in s:
        s = re.sub(r"(<dcterms:modified[^>]*>).*?(</dcterms:modified>)", rf"\g<1>{now}\g<2>", s)
    if modified_by is not None and "<cp:lastModifiedBy>" in s:
        s = re.sub(r"(<cp:lastModifiedBy>).*?(</cp:lastModifiedBy>)", rf"\g<1>{modified_by}\g<2>", s)
    m = re.search(r"<cp:revision>(\d+)</cp:revision>", s)
    if m:
        s = s.replace(m.group(0), f"<cp:revision>{int(m.group(1)) + 1}</cp:revision>")
    return s.encode("utf-8")


def save_preserving(original, wb, out=None, modified_by=None):
    """Save `wb` to `out` (default: overwrite `original`), preserving `original`'s
    non-data OOXML parts and updating core.xml change-tracking."""
    out = out or original
    now = _now_utc()
    fd, tmp = tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
    fd, tmp2 = tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
    try:
        wb.save(tmp)
        with zipfile.ZipFile(original) as zo, zipfile.ZipFile(tmp) as ze:
            edited = set(ze.namelist())
            with zipfile.ZipFile(tmp2, "w", zipfile.ZIP_DEFLATED) as zw:
                written = set()
                for name in zo.namelist():
                    if name == "docProps/core.xml":
                        data = _bump_core(zo.read(name), modified_by, now)
                    else:
                        data = ze.read(name) if name in edited else zo.read(name)
                    zw.writestr(name, data); written.add(name)
                for name in ze.namelist():
                    if name not in written:
                        zw.writestr(name, ze.read(name))
        os.replace(tmp2, out)
    finally:
        for p in (tmp, tmp2):
            if os.path.exists(p):
                os.remove(p)
