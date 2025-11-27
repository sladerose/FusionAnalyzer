// Content script to scrape data from Fusion

console.log("Fusion Analyzer: Content script loaded.");

// Listen for messages from the popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "scrape_hours") {
        const result = scrapeTimesheet();
        if (result.error) {
            sendResponse({ success: false, error: result.error });
        } else {
            sendResponse({ success: true, data: result });
        }
    } else if (request.action === "debug_html") {
        // Return a simplified structure of the page for debugging
        sendResponse({ success: true, html: document.body.innerHTML });
    }
    return true; // Keep channel open for async response
});

function scrapeTimesheet() {
    const projectHours = {};

    // Strategy: Find the main table. 
    // We look for a table that contains "Project Name" in a header.
    const tables = document.querySelectorAll("table");
    let targetTable = null;
    let headerRowIndex = -1;
    let projectNameColIndex = -1;

    // 1. Find the correct table and column indices
    for (const table of tables) {
        const rows = table.querySelectorAll("tr");
        for (let r = 0; r < rows.length; r++) {
            const cells = rows[r].querySelectorAll("th, td");
            for (let c = 0; c < cells.length; c++) {
                // Normalize text: remove newlines, extra spaces, to lowercase
                const text = cells[c].innerText.replace(/\s+/g, ' ').trim().toLowerCase();

                if (text.includes("project name")) {
                    targetTable = table;
                    headerRowIndex = r;
                    projectNameColIndex = c;
                    console.log("Fusion Analyzer: Found header at row", r, "col", c);
                    break;
                }
            }
            if (targetTable) break;
        }
        if (targetTable) break;
    }

    if (!targetTable) {
        console.error("Fusion Analyzer: Could not find timesheet table.");
        // Check if we're on the Dashboard page
        if (document.getElementById("NotificationPopup") || document.querySelector(".dashboardMessagesDiv")) {
            console.log("Fusion Analyzer: Detected Dashboard page.");
            return { error: "WRONG_PAGE_DASHBOARD" };
        }
        return {};
    }

    // 2. Parse the rows
    const rows = targetTable.querySelectorAll("tr");
    const headerCells = rows[headerRowIndex].querySelectorAll("th, td");
    const hourColumnIndices = [];

    for (let c = projectNameColIndex + 1; c < headerCells.length; c++) {
        const text = headerCells[c].innerText.trim();
        // Skip "Total", "Comments"
        if (text.toLowerCase().includes("comment") || text.toLowerCase().includes("total") || text.length === 0) {
            continue;
        }
        hourColumnIndices.push(c);
    }

    // Iterate data rows
    for (let r = headerRowIndex + 1; r < rows.length; r++) {
        const cells = rows[r].querySelectorAll("td");
        if (cells.length <= projectNameColIndex) continue;

        const rawProjectName = cells[projectNameColIndex].innerText.trim();
        // Ignore empty rows, Total rows, and Report Date rows
        if (!rawProjectName || rawProjectName.startsWith("Total") || rawProjectName.startsWith("Report Date")) continue;

        // Clean the name
        const cleanName = cleanProjectName(rawProjectName);

        let rowTotal = 0;
        for (const colIdx of hourColumnIndices) {
            if (colIdx < cells.length) {
                const cellText = cells[colIdx].innerText.trim();
                const val = parseFloat(cellText);
                if (!isNaN(val) && val > 0) {
                    rowTotal += val;
                }
            }
        }

        if (rowTotal > 0) {
            projectHours[cleanName] = (projectHours[cleanName] || 0) + rowTotal;
        }
    }

    return projectHours;
}
