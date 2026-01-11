# Attendance Report Processor

An automated Python-based tool designed to process raw employee attendance logs and generate structured individual reports using a predefined Excel template.

## Overview

This tool was developed to address a specific workplace requirement for automating the generation of attendance reports. It is specifically designed to work with a static source data structure (`DATA.xlsx`) and map the processed information into a static predefined template (`taslak.xlsx`). By automating the transition between these two specific formats, the script ensures consistency and significantly reduces manual data entry time.

## Key Features

- **Smart Name Cleaning:** Automatically removes prefixes (e.g., `arge*`, `***`) and special characters from staff names using Regex patterns.
- **Daily Aggregation:** Merges multiple entries and exits into a single record per day, identifying the first entry, the last exit, and calculating the total net duration.
- **Template Mapping:** Intelligently matches processed data with the correct date rows in the provided static Excel template.
- **Batch Processing:** Processes all unique staff members found in the source data and generates separate, individual Excel files for each person.
- **Error Handling & Logging:** Features a professional logging system to monitor the process and handle missing or corrupted data without crashing.

## Project Structure

- `attendance_processor.py`: The main Python execution script.
- `requirements.txt`: List of required Python libraries for the environment.
- `DATA.xlsx`: The raw attendance data source (Must follow the specific static column structure).
- `taslak.xlsx`: The target report template (Must follow the specific static row/cell structure).

## Installation

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/yourusername/attendance-processor.git](https://github.com/yourusername/attendance-processor.git)
   cd attendance-processor