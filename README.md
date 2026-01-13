# City Directory Analyzer (ELC)

## Overview

The **City Directory Analyzer** is an internal analysis and reporting tool developed for **Environmental Logistics & Consulting (ELC)** to support historical occupant research during Phase I Environmental Site Assessments.

The application processes city directory exports (CSV or Excel) and converts raw listing data into clean, structured tables that summarize historical occupants by year for both **subject properties** and **adjoining properties**. It also generates client-ready Word (DOCX) report tables that can be inserted directly into Phase I ESA reports.

The tool is built with **Python** and **Streamlit** and is intended for internal operational use.

---

## Core Capabilities

### City Directory File Ingestion

* Accepts CSV and XLSX city directory exports
* Automatically detects header rows in irregular Excel files
* Normalizes address values and forward-fills missing address rows

### Subject Property Analysis

* Groups directory listings by year for selected subject addresses
* Deduplicates occupants per year
* Compresses consecutive year ranges with identical occupants
* Displays results in clean, on-screen summary tables

### Adjoining Property Analysis

* Performs the same historical occupant analysis for adjoining properties
* Supports optional directional labeling (North, South, East, West)
* Presents results in a parallel table layout for side-by-side review

### Report-Ready Output

* Generates one-click downloadable **DOCX tables** in ELC report format
* Separate exports for:

  * Subject Property tables
  * Adjoining Property tables
* Formatting aligned with standard Phase I ESA reporting language

### Clean Presentation Layer

* Card-based address sections for readability
* Fixed-width, report-style tables
* Designed for analyst review rather than raw data inspection

---

## Application Structure

```
citydir/
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
├── venv/               # Local virtual environment 
└── README.md
```

---

## Technology Stack

* **Python 3.11+**
* **Streamlit** (UI framework)
* **Pandas** (data transformation)
* **python-docx** (Word document generation)
* **OpenPyXL** (Excel parsing)

---

## Input Data Requirements

The uploaded city directory file must contain, at minimum, the following columns:

* **ADDRESS** – Property address
* **YEAR** – Listing year
* **LISTING** – Occupant or business name

Notes:

* Column names are normalized to uppercase
* Missing address values are forward-filled automatically
* Duplicate occupant listings within the same year are deduplicated

---

## Output Behavior

### On-Screen Tables

* One table per address
* Rows represent year or compressed year ranges
* Occupants displayed as a single consolidated value per row

### DOCX Report Tables

* Client-ready formatting
* Grid-style tables suitable for direct inclusion in Phase I ESA reports
* Consistent structure across subject and adjoining properties

---

## Intended Use

This tool is designed to:

* Support historical occupant research for Phase I ESAs
* Reduce manual formatting of city directory data
* Improve consistency and defensibility of report exhibits

It is **not** intended to function as a public-facing data exploration tool.

---

## Maintenance Notes

* Changes to report formatting should be made in the DOCX helper functions
* Input schema assumptions are centralized in the file ingestion logic
* The tool intentionally avoids database storage and runs entirely in-memory

---

## Internal Status

* Actively used by ELC analysts
* Local execution only
* Version-controlled via GitHub

---

## License

Internal use only. Distribution outside of ELC requires authorization.
