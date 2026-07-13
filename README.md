# onedrive_safe_xlsx

[![tests](https://github.com/jonnymaserati/onedrive-safe-xlsx/actions/workflows/test.yml/badge.svg)](https://github.com/jonnymaserati/onedrive-safe-xlsx/actions/workflows/test.yml)
[![PyPI](https://img.shields.io/pypi/v/onedrive-safe-xlsx)](https://pypi.org/project/onedrive-safe-xlsx/)

Edit `.xlsx` files with **openpyxl** without breaking **OneDrive / SharePoint** sync and version history.

## The problem

openpyxl rebuilds a workbook from its own minimal model when it saves. A real Excel file
has ~26 OOXML parts; openpyxl writes ~9 — it silently **drops** `customXml/*`,
`docProps/custom.xml`, `printerSettings`, `calcChain`, and more.

On OneDrive/SharePoint those parts carry the document's **identity and library binding**.
Strip them and your edits can:

- lose the file's **version lineage** (each save looks like a *new/foreign* file rather than a
  new version of the same document), and
- if you *do* preserve the parts but copy `docProps/core.xml` unchanged, its
  `dcterms:modified` / `cp:revision` stay the same, so SharePoint decides **nothing changed**
  and won't record a new version.

## The fix

`save_preserving()` edits the data via openpyxl but:

1. **Keeps every non-data part** from the original file (identity preserved).
2. **Updates the change-tracking** in `core.xml` — `dcterms:modified` → now,
   `cp:lastModifiedBy`, and bumps `cp:revision` — so SharePoint logs a **genuine new version**.

## Usage

```python
import openpyxl
from onedrive_safe_xlsx import save_preserving

wb = openpyxl.load_workbook("Report.xlsx")   # an Excel/SharePoint-saved file
wb.active["A1"] = "edited"
save_preserving("Report.xlsx", wb, modified_by="Jane Doe")
```

## Caveats

- **Start from an Excel/SharePoint-saved copy** — it has the parts to preserve. A file only
  ever written by openpyxl has nothing to preserve (open + save it in Excel once first).
- Uses `datetime.now(timezone.utc)` for the timestamp — cross-platform (Windows/macOS/Linux).
- Doesn't manage `rsid` co-authoring markers — intended for sequential (not concurrent) edits.

## How to cite

If it helps your work, a credit is appreciated — GitHub's "Cite this repository" button
uses `CITATION.cff`. The licence is permissive (Apache-2.0), so this is a request, not a
requirement.

## Licence

**Apache-2.0** — see `LICENSE`. Copyright 2026 Jonathan Corcutt; attribution terms in `NOTICE`.
