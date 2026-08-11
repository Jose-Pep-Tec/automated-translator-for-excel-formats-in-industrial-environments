# Automated Translator for Excel Formats in Industrial Environments

A standalone Python GUI utility designed to translate industrial plant spreadsheets and compliance formats into technical English for international audits, preserving formulas, cell formatting, and sheet structures.

## Problem Solved

Industrial plant spreadsheets often contain critical formulas, merged cells, and conditional formatting required for audits. Manual translation breaks formulas and corrupts sheet layouts. This tool automates batch cell translation while keeping all Excel formulas, visual formatting, and sheet structures intact.

## Key Features

* **Structure & Formula Integrity:** Translates text without altering formulas, cell merges, or color formatting.
* **Custom Technical Dictionary:** Applies local glossary rules (`custom_dictionary.json`) before API calls to enforce exact plant nomenclature.
* **Dual-Language Mode:** Generates `Original / Translation` text within cells for bilingual verification.
* **Batch Request Management:** Groups cell data with retry/backoff protection to prevent rate limits on large workbooks.
* **Worksheet & File Renaming:** Automatically translates tab names, file names, and applies output prefixes.

## Repository Structure

```text
automated-translator-for-excel-formats-in-industrial-environments/
├── main_translator_gui.py     # GUI and translation engine
├── custom_dictionary.json     # Technical glossary storage
├── requirements.txt           # Python dependencies
└── README.md
