# Utility Bill Automation System

A Python desktop application that automates utility bill PDF extraction, validates meter-level billing data, and updates Excel workbook templates while preserving existing formulas and reporting logic.

## Overview

This project was built to reduce manual data entry in utility bill processing workflows. The application reads structured utility bill PDFs, extracts meter-level data, validates the extracted values, and writes the results into an Excel template used for monthly reporting.

## Key Features

- Extracts meter-level data from PDF utility bills
- Parses previous reads, current reads, usage, multipliers, billing period, and current charges
- Cleans numeric values and handles formatted readings
- Updates Excel templates using `openpyxl`
- Preserves existing workbook formulas and summary logic
- Detects missing meters, extra meters, multiplier changes, and usage mismatches
- Provides a desktop GUI for non-technical users
- Can be packaged into a standalone Windows `.exe`

## Tech Stack

- Python
- Tkinter
- pdfplumber
- openpyxl
- Regular Expressions
- PyInstaller
- Excel Automation

## Application Workflow

```text
PDF Utility Bill
      ↓
Text Extraction
      ↓
Meter Data Parsing
      ↓
Data Cleaning
      ↓
Validation Checks
      ↓
Excel Template Update
      ↓
Generated Workbook + Validation Report
