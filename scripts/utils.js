// Utility functions shared between popup and content script

/**
 * Helper to format date to YYYY-MM-DD string.
 * @param {Date} date - The date object to format.
 * @returns {string} - The date formatted as "YYYY-MM-DD".
 */
const formatDate = (date) => date.toISOString().split('T')[0];

/**
 * Checks if a given date is a business day (Monday-Friday).
 * @param {Date} date - The date to check.
 * @returns {boolean} - True if it's a business day, false otherwise.
 */
function isBusinessDay(date) {
    const dayOfWeek = date.getDay(); // 0 = Sunday, 6 = Saturday
    return dayOfWeek !== 0 && dayOfWeek !== 6;
}

/**
 * Counts the total number of business days in a given month.
 * @param {number} year - The year (e.g., 2023).
 * @param {number} month - The month (0-indexed, e.g., 0 for January).
 * @returns {number} - The total number of business days.
 */
function getBusinessDaysInMonth(year, month) {
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    let businessDays = 0;

    for (let day = 1; day <= daysInMonth; day++) {
        const date = new Date(year, month, day);
        if (isBusinessDay(date)) {
            businessDays++;
        }
    }
    return businessDays;
}

/**
 * Counts the number of elapsed business days up to a given day in a month.
 * @param {number} year - The year.
 * @param {number} month - The month (0-indexed).
 * @param {Date} today - The current date object.
 * @returns {number} - The number of business days elapsed.
 */
function getBusinessDaysElapsed(year, month, today) {
    // Ensure 'today' is within the target month for accurate elapsed calculation
    const currentMonthToday = new Date(year, month, today.getDate());

    let businessDays = 0;
    for (let day = 1; day <= currentMonthToday.getDate(); day++) {
        const date = new Date(year, month, day);
        if (isBusinessDay(date)) {
            businessDays++;
        }
    }
    return businessDays;
}

/**
 * Calculates the number of remaining business days from a given date until the end of the month.
 * @param {number} year - The year.
 * @param {number} month - The month (0-indexed).
 * @param {Date} today - The current date object.
 * @returns {number} - The number of remaining business days.
 */
function getRemainingBusinessDays(year, month, today) {
    const totalBusinessDays = getBusinessDaysInMonth(year, month);
    const elapsedBusinessDays = getBusinessDaysElapsed(year, month, today);
    return totalBusinessDays - elapsedBusinessDays;
}

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
 * Calculates the dynamic daily consumption rate for a project based on historical actual hours.
 * Uses all historical consumption data available to date.
 * @param {Object} projectActualDailyHours - An object like { "YYYY-MM-DD": hours } for a specific project.
 * @param {Date} today - The current date object.
 * @returns {number} - The calculated daily consumption rate.
 */
function calculateDailyConsumptionRate(projectActualDailyHours, today) {
    let totalConsumedHours = 0;
    let consumedBusinessDaysCount = 0;

    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth();

    for (const dateString in projectActualDailyHours) {
        const dailyDate = new Date(dateString);
        // Ensure the date is valid and within the current month/year and not in the future
        if (
            !isNaN(dailyDate) &&
            dailyDate.getFullYear() === currentYear &&
            dailyDate.getMonth() === currentMonth &&
            dailyDate <= today &&
            isBusinessDay(dailyDate)
        ) {
            totalConsumedHours += projectActualDailyHours[dateString];
            consumedBusinessDaysCount++;
        }
    }

    if (consumedBusinessDaysCount > 0) {
        return totalConsumedHours / consumedBusinessDaysCount;
    }
    return 0; // No historical data yet, or no hours logged on business days
}

/**
 * Forecasts the project status (run out of hours, have hours remaining, on track)
 * and optionally provides a projected exhaustion date.
 * @param {Object} projectActualDailyHours - An object like { "YYYY-MM-DD": hours } for a specific project.
 * @param {number} projectPlannedTotalHours - The total monthly hours allocated for this project.
 * @returns {Object} - An object containing forecast text, class, and optionally exhaustionDate.
 */
function forecastProjectStatus(projectActualDailyHours, projectPlannedTotalHours) {
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth(); // 0-indexed

    const dailyRate = calculateDailyConsumptionRate(projectActualDailyHours, now);
    const remainingBusinessDays = getRemainingBusinessDays(currentYear, currentMonth, now);

    // Calculate Total Consumed Hours so far (from projectActualDailyHours)
    let currentMonthActuals = 0;
    for (const dateString in projectActualDailyHours) {
        const dailyDate = new Date(dateString);
        if (
            !isNaN(dailyDate) &&
            dailyDate.getFullYear() === currentYear &&
            dailyDate.getMonth() === currentMonth &&
            dailyDate <= now // Only count actuals up to and including today
        ) {
            currentMonthActuals += projectActualDailyHours[dateString];
        }
    }

    if (dailyRate === 0 && currentMonthActuals === 0) {
        return { text: "No data yet", class: "forecast-neutral", exhaustionDate: null };
    }
    
    // If no planned hours, it's considered unplanned, not forecastable in this context
    if (projectPlannedTotalHours === 0) {
         return { text: "Unplanned", class: "forecast-unplanned", exhaustionDate: null };
    }


    const projectedFutureConsumption = dailyRate * remainingBusinessDays;
    const projectedRemainingHours = projectPlannedTotalHours - currentMonthActuals - projectedFutureConsumption;

    let forecastText = "On track";
    let forecastClass = "forecast-neutral";
    let exhaustionDate = null;
    const buffer = 2; // Hours buffer for "on track"

    // Helper to format date to "Mon 16th"
    const formatExhaustionDate = (dateString) => {
        const date = new Date(dateString);
        if (isNaN(date)) return dateString;
        const options = { weekday: 'short', day: 'numeric' };
        const formatted = date.toLocaleDateString('en-US', options); // e.g., "Mon, 16"
        // Add ordinal suffix (st, nd, rd, th)
        const day = date.getDate();
        let suffix = 'th';
        if (day === 1 || day === 21 || day === 31) {
            suffix = 'st';
        } else if (day === 2 || day === 22) {
            suffix = 'nd';
        } else if (day === 3 || day === 23) {
            suffix = 'rd';
        }
        return formatted.replace(/, (\d+)$/, ` ${day}${suffix}`); // Replace with day and suffix
    };

    // Check for "Ran out on Tue 12th" scenario first
    if (currentMonthActuals > projectPlannedTotalHours && dailyRate > 0) {
        // Calculate the date when hours were actually exhausted
        let consumedToDate = 0;
        let pastExhaustionDate = null;
        for (let day = 1; day <= now.getDate(); day++) {
            const date = new Date(currentYear, currentMonth, day);
            if (isBusinessDay(date)) {
                const dateString = formatDate(date);
                consumedToDate += (projectActualDailyHours[dateString] || 0);
                if (consumedToDate >= projectPlannedTotalHours) {
                    pastExhaustionDate = date;
                    break;
                }
            }
        }
        if (pastExhaustionDate) {
            forecastText = `Ran out on ${formatExhaustionDate(formatDate(pastExhaustionDate))}`;
            forecastClass = "forecast-under";
        } else {
            // This case might happen if currentMonthActuals > planned, but dailyRate is 0
            // or if the exhaustion happened due to non-business days, which isn't explicitly tracked daily.
            // For simplicity, we'll just say "Over budget"
             forecastText = `Over budget by ${formatHours(currentMonthActuals - projectPlannedTotalHours)}`;
             forecastClass = "forecast-over";
        }
    } else if (projectedRemainingHours > buffer) {
        forecastText = `Increase consumption by ${formatHours(projectedRemainingHours / remainingBusinessDays)} hrs/day`;
        forecastClass = "forecast-over"; // More hours than needed
    } else if (projectedRemainingHours < -buffer) {
        // Calculate Projected Exhaustion Date for future
        let tempCurrentActuals = 0;
        for (const dateString in projectActualDailyHours) {
            const dailyDate = new Date(dateString);
            if (
                !isNaN(dailyDate) &&
                dailyDate.getFullYear() === currentYear &&
                dailyDate.getMonth() === currentMonth &&
                dailyDate <= now
            ) {
                tempCurrentActuals += projectActualDailyHours[dateString];
            }
        }
        
        let hoursLeftToCover = projectPlannedTotalHours - tempCurrentActuals;

        if (dailyRate > 0) {
            let tempDate = new Date(now); 
            tempDate.setDate(tempDate.getDate() + 1); // Start checking from tomorrow

            while (tempDate.getMonth() === currentMonth) { // Only forecast within the current month
                if (isBusinessDay(tempDate)) {
                    hoursLeftToCover -= dailyRate;
                    if (hoursLeftToCover <= 0) {
                        exhaustionDate = formatDate(tempDate);
                        break;
                    }
                }
                tempDate.setDate(tempDate.getDate() + 1);
            }
        }
        
        if (exhaustionDate) {
            forecastText = `Run out by ${formatExhaustionDate(exhaustionDate)}`;
        } else {
            // If for some reason exhaustion date can't be calculated but hours are negative
            forecastText = `Will run out by ${formatHours(Math.abs(projectedRemainingHours))}`;
        }
        forecastClass = "forecast-under"; // Less hours than needed, will run out
    } else {
        forecastText = "On track";
        forecastClass = "forecast-neutral";
    }

    return {
        text: forecastText,
        class: forecastClass,
        remaining: projectedRemainingHours,
        exhaustionDate: exhaustionDate
    };
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
 * @returns {Object} - An object with detailed project data including daily hours and metadata.
 *   Example:
 *   {
 *       "Project Name A": {
 *           "dailyHours": {
 *               "YYYY-MM-DD": hours,
 *               // ...
 *           },
 *           "totalHours": total_sum_of_daily_hours
 *       },
 *       "meta": {
 *           "startDate": "YYYY-MM-DD",
 *           "endDate": "YYYY-MM-DD"
 *       }
 *   }
 */
function parseXLSXData(sheetData) {
    const projectData = {}; // Stores all project data with daily breakdown
    projectData.meta = {}; // Stores metadata like date range

    let startDateRange = null;
    let endDateRange = null;

    // Helper to format date to YYYY-MM-DD string
    const formatDate = (date) => date.toISOString().split('T')[0];

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
                return {};
            }
            break; // Found it, no need to continue loop
        }
    }

    if (!startDateRange || !endDateRange) {
        console.error("Could not find 'Report Date From' row in the sheet.");
        return {};
    }

    projectData.meta.startDate = formatDate(startDateRange);
    projectData.meta.endDate = formatDate(endDateRange);

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
            !projectNameRaw.startsWith("Staff Name:") &&
            !projectNameRaw.startsWith("Employee Number:") &&
            !projectNameRaw.toLowerCase().startsWith("signature"))
        {
            const projectNameClean = cleanProjectName(projectNameRaw);

            // Initialize project data if not already present
            if (!projectData[projectNameClean]) {
                projectData[projectNameClean] = {
                    dailyHours: {},
                    totalHours: 0
                };
            }

            for (const c_idx in currentWeekDateCols) {
                const hoursVal = row[c_idx];
                if (hoursVal === null || hoursVal === undefined || isNaN(parseFloat(hoursVal))) {
                    continue;
                }

                try {
                    const hours = parseFloat(hoursVal);
                    if (hours > 0) {
                        const dateString = formatDate(currentWeekDateCols[c_idx]);
                        // Add daily hours
                        projectData[projectNameClean].dailyHours[dateString] =
                            (projectData[projectNameClean].dailyHours[dateString] || 0) + hours;
                        // Keep track of total hours
                        projectData[projectNameClean].totalHours += hours;
                    }
                } catch (e) {
                    // Ignore cells that can't be parsed as a number
                }
            }
        }
    }

    return projectData;
}
