# Fusion Timesheet Analyzer Chrome Extension

This Chrome Extension replaces the manual Python script workflow for analyzing Fusion timesheets. It allows you to define your planned hours once and automatically compares them against your actual hours on the Fusion website.

## Local Distribution to Colleagues

To distribute this extension to colleagues for local testing or use:

1.  **Package the extension:** Create a `.zip` file of the `fusion-extension` directory.
    *   Navigate to the directory containing `fusion-extension`.
    *   Run `zip -r fusion-extension.zip fusion-extension/` in your terminal.
2.  **Share the `.zip` file:** Provide this `fusion-extension.zip` file to your colleagues.

## Installation (for Unpacked Extension or after unzipping)

1.  Open Chrome and navigate to `chrome://extensions`.
2.  Enable **Developer mode** (toggle in the top right).
3.  Click **Load unpacked**.
4.  Select the `fusion-extension` folder (either directly if shared unpacked, or after unzipping the provided `.zip` file).

## How to Use

**Optimal Workflow:** It is highly recommended to set your planned hours BEFORE uploading your timesheet for comparison.

1.  **Set Planned Hours (First and whenever plans change)**:
    *   Click the extension icon in your Chrome toolbar.
    *   Go to the **Settings** tab.
    *   Add your projects and their planned monthly hours.
    *   Click **Save Settings**.

2.  **Upload Actual Hours (XLSX)**:
    *   **Export your timesheet from Fusion:** Go to Fusion Daily Timesheet, set the date range to "Month to Date", and click the "EXPORT" button.
    *   Click the extension icon.
    *   Go to the **Settings** tab.
    *   In the "Upload Actuals (XLSX)" section, click "Choose XLSX File" and select the timesheet file you exported from Fusion.
    *   The extension will process the file and automatically display the comparison on the Dashboard.

3.  **Analyze Timesheet**:
    *   Click the extension icon.
    *   The **Dashboard** tab will instantly show your Actual vs. Planned hours.
    *   **Unplanned Work**: This calculation now includes:
        *   Any work done on projects *not* in your planned list.
        *   Any *excess* hours worked on planned projects (e.g., if you planned 4 hours but worked 6, those extra 2 hours count as unplanned).

## Version 1.1 Updates
- **Unplanned Work Logic**: Updated to capture overage hours on planned projects.
- **Improved Use Anywhere**: The extension popup now works on any tab (including `chrome://` pages), allowing you to view and edit your Planned Hours even if the current page cannot be scraped.

## Troubleshooting

If the data isn't showing up or looks wrong:

1.  Ensure that the column names in your XLSX file for "Project" and "Hours" match what's expected by the extension. You can configure these in the **Settings** tab if your file uses different names.
2.  If you're still having issues, click the **Debug (Copy HTML)** button in the extension popup.
3.  Paste the result to the developer (you) so the scraper can be adjusted (if the issue is with scraping) or for further diagnosis.
