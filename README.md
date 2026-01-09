# Spreadsheet Combiner

A local workflow tool to verify and combine multiple spreadsheet files (.csv and .xlsx) into a single Excel workbook.

## Features
-   Scans `input/` directory for files.
-   Combines all files into a single `.xlsx` file in `output/`.
-   **CSV Support**: Adds as a single sheet named after the file.
-   **Excel Support**: Extracts **all** sheets from source workbooks and preserves their names (handling duplicates automatically).

## Installation

### Prerequisites
-   Node.js (v18 or later recommended)
-   npm

### Setup
1.  Clone the repository:
    ```bash
    git clone git@github.com:that0n3guy/spreadsheet-combiner.git
    cd spreadsheet-combiner
    ```
2.  Install dependencies:
    ```bash
    npm install
    ```

## Usage

1.  Place your `.csv` and `.xlsx` files into the `input/` folder.
2.  Run the combiner:
    ```bash
    npm start
    ```
3.  Find the result in the `output/` folder (e.g., `combined_2026-01-09....xlsx`).


## Development
-   **Run tests / Verification**:
    ```bash
    npx ts-node src/verify_output.ts
    ```
