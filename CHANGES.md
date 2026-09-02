# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html). Dates formatted as YYYY-MM-DD as per [ISO standard](https://www.iso.org/iso-8601-date-and-time-format.html).

## 0.3.0 2026-09-2

Terminal cells - temp removed due to raising unsolved error in some cells.

### Removed

* App no longer displays terminal cells
* Terminal cell tracing in `xl-ray.extractor.extract_worksheets`

## v0.2.0 2026-06-29

Terminal cells. i.e. track cells with no forward depenendicies.  

### Change

* `schema.CellData` now has `is_terminal` flag
* `extractor.extract_worksheets` updated to perform 2nd pass on cells to identify forward references.
* `app.py` audit tab updated to include section on terminal cells.
* `xl-ray.__version__` displayed in app sidebar.

### Added

* `CHANGES.md`

## [v0.1.0 2026-06-24](https://github.com/TomMonks/xl-ray/releases/tag/v0.1.0)

Initial release.  Created `xl-ray` package and simple interface

### Added

* Read in Excel file and store in JSON format
* Basic description
* Basic audit of cells for mistakes and complexity
* Example LLM prompts.
* `streamlit` dashboard front end