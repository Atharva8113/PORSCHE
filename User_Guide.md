# Porsche After Sales – Invoice Extractor

Automatically extract and convert line-item data from Porsche AG After-Sales PDFs into formatted CSV files for easy reporting.

## Quick Start
*   **Requirements:** Windows OS (No Python installation required).
*   **Installation:** Simply locate the `Porsche_AfterSales_Extractor_App.exe` file in your folder.
*   **No Setup Needed:** The Nagarkot logo and all necessary logic are embedded inside the application.

## Folder Structure
To keep things organized, we recommend a structure like this:
```text
/PORSCHE_EXTRACTOR/
├── Porsche_AfterSales_Extractor_App.exe  <-- Double-click to run
├── [Input_Invoices].pdf                  <-- Place your invoices here
└── [Extracted_Data].csv                  <-- Automatically generated output
```

## How to Use (Step-by-Step)
1.  **Launch the App:** Double-click `Porsche_AfterSales_Extractor_App.exe`.
2.  **Select Invoices:** Click the **"Select PDFs"** button to choose one or multiple Porsche invoice files.
3.  **Choose Mode:**
    *   **Combined:** Merges all selected invoices into a single CSV file.
    *   **Individual:** Creates a separate CSV for every invoice (named after the Invoice Number).
4.  **Set Output Folder:** Click **"Browse..."** to choose where to save the result (defaults to the folder where the PDFs are).
5.  **Run Extraction:** Click the blue **"Extract & Generate CSV"** button in the bottom right.
6.  **Clear List:** If you want to start a new batch, click **"Clear List"** to reset the queue and output settings.

## Common Issues
*   **Windows Defender:** If a "Windows protected your PC" popup appears, click **"More info"** and then **"Run anyway"**. This is normal for internally developed tools.
*   **File in Use:** If you are overwriting an existing CSV, ensure that the file is **closed** in Excel before running the extraction, or the tool will be unable to write the data.
*   **Missing Lines:** Ensure your PDF is a readable document (not a scanned image without OCR). The tool is designed specifically for Porsche AG After-Sales invoices.

## Contact
For support, bug reports, or feature requests, please contact the IT Team.

---
*© Nagarkot Forwarders Pvt Ltd*
