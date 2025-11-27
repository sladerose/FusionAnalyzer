// Utility functions shared between popup and content script

/**
 * Cleans and standardizes project names to match Planned Hours keys.
 * Ported from the original Python script.
 * @param {string} name - The raw project name from the website.
 * @returns {string} - The standardized project name.
 */
function cleanProjectName(name) {
    if (!name || typeof name !== 'string') {
        return String(name || '').trim();
    }

    const originalName = name.trim();
    const nameLower = originalName.toLowerCase();

    if (nameLower.includes("doms and pos ingesting services")) {
        return "Astron DOMS & POS Ingesting";
    }
    if (nameLower.includes("glencore mobile tracking")) {
        return "Mobile Warehouse Operations";
    }
    if (nameLower.includes("psb to psd migration")) {
        return "PSB to PSD Migration";
    }
    if (nameLower.includes("stcms01-(tcms)transportation contract management system project")) {
        return "Sasol Contract Management System";
    }
    if (nameLower.includes("btt weighbridge")) {
        return "Mobile Arivals (Weighbridge Change request)";
    }
    
    let modifiedName = originalName;
    // Handle "00000001-BD-" prefix logic
    // Python: if len(name) > 7 and name[6] == '-': name = name[7:].strip()
    // JS: Check if char at index 6 is '-'
    if (modifiedName.length > 7 && modifiedName.charAt(6) === '-') {
        modifiedName = modifiedName.substring(7).trim();
    }
    
    if (modifiedName.startsWith("ges/No_image.jpg00000001-BD- ")) {
        modifiedName = "BD- Project Stronghold";
    }

    return modifiedName.trim();
}

/**
 * Formats a number to 2 decimal places.
 * @param {number} num 
 * @returns {string}
 */
function formatHours(num) {
    return (Math.round((num + Number.EPSILON) * 100) / 100).toFixed(2);
}

/**
 * Parses the complex structure of a Fusion Timesheet Export XLSX file.
 * This is a port of the original Python script's parsing logic.
 * @param {Array<Array<any>>} sheetData - The raw sheet data from XLSX.utils.sheet_to_json(..., {header: 1}).
 * @returns {Object} - An object mapping cleaned project names to their total actual hours.
 */
function parseXLSXData(sheetData) {
    const projectHours = {};
    let startDateRange = null;
    let endDateRange = null;

    // --- 1. First Pass: Find the overall report date range ---
    for (const row of sheetData) {
        if (row[0] && typeof row[0] === 'string' && row[0].startsWith("Report Date From:")) {
            const dateRangeStr = row[0].split(":")[1].trim();
            try {
                const [dateFromStr, dateToStr] = dateRangeStr.split(" to ");
                // Assuming format DD/MM/YYYY
                const fromParts = dateFromStr.split('/');
                const toParts = dateToStr.split('/');
                startDateRange = new Date(`${fromParts[2]}-${fromParts[1]}-${fromParts[0]}`);
                endDateRange = new Date(`${toParts[2]}-${toParts[1]}-${toParts[0]}`);
                // Set to midnight to ensure comparisons are date-based
                startDateRange.setHours(0, 0, 0, 0);
                endDateRange.setHours(0, 0, 0, 0);
            } catch (e) {
                console.error("Could not parse report date range:", dateRangeStr, e);
                // If we can't parse dates, we can't reliably sum hours.
                return {}; 
            }
            break; // Found it, no need to continue loop
        }
    }

    if (!startDateRange || !endDateRange) {
        console.error("Could not find 'Report Date From' row in the sheet.");
        return {};
    }

    // --- 2. Second Pass: Iterate through rows to find headers and data ---
    let currentWeekDateCols = {}; // Stores {columnIndex: DateObject}

    for (const row of sheetData) {
        if (!row || row.length === 0) continue;

        // --- A. Identify a new weekly header row ---
        if (row[0] === "Project Name" && row[1] === "Work Code") {
            currentWeekDateCols = {}; // Reset for the new week
            for (let c_idx = 2; c_idx < row.length; c_idx++) {
                const cellValue = row[c_idx];
                if (!cellValue) continue;

                let currentColDate = null;
                // The 'xlsx' library can auto-convert dates. It might be a JS Date object.
                if (cellValue instanceof Date) {
                    currentColDate = cellValue;
                } 
                // Or it might be a number (Excel's date serial number)
                else if (typeof cellValue === 'number' && cellValue > 1) {
                    // XLSX.SSF.parse_date_code is not available in the minified version, so we use a simpler conversion
                    currentColDate = new Date(Date.UTC(1899, 11, 30) + cellValue * 86400000);
                }
                // Or a string 'DD/MM/YYYY'
                else if (typeof cellValue === 'string') {
                    const parts = cellValue.split(' ')[0].split('/');
                    if (parts.length === 3) {
                       currentColDate = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
                    }
                }

                if (currentColDate) {
                    currentColDate.setHours(0, 0, 0, 0);
                    // --- B. Only include dates that are within the report's overall range ---
                    if (currentColDate >= startDateRange && currentColDate <= endDateRange) {
                        currentWeekDateCols[c_idx] = currentColDate;
                    }
                }
            }
            continue; // Move to the next row after processing the header
        }
        
        // --- C. Process data rows ---
        const projectNameRaw = row[0];
        if (projectNameRaw && typeof projectNameRaw === 'string' && 
            projectNameRaw.trim() !== '' &&
            !projectNameRaw.startsWith("Project Name") &&
            !projectNameRaw.startsWith("Total Hours") &&
            !projectNameRaw.startsWith("Report Date From:") &&
            !projectNameRaw.startsWith("Date & Time Exported:") &&
            !projectNameRaw.toLowerCase().startsWith("signature")) 
        {
            const projectNameClean = cleanProjectName(projectNameRaw);
            
            for (const c_idx in currentWeekDateCols) {
                const hoursVal = row[c_idx];
                if (hoursVal === null || hoursVal === undefined || isNaN(parseFloat(hoursVal))) {
                    continue;
                }
                
                try {
                    const hours = parseFloat(hoursVal);
                    if (hours > 0) {
                        projectHours[projectNameClean] = (projectHours[projectNameClean] || 0) + hours;
                    }
                } catch (e) {
                    // Ignore cells that can't be parsed as a number
                }
            }
        }
    }

    return projectHours;
}
