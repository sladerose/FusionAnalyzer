import { Controller } from "@hotwired/stimulus"

// --- Utility functions (copied directly from previous script.js for now) ---
// These will eventually be refactored into a separate utility file if they are used by multiple controllers.
// For now, to keep the "smallest footprint necessary" and minimal files, we'll put them here.

// Helper to format date to YYYY-MM-DD string.
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
    return 0;
}

/**
 * Forecasts the project status.
 * @param {Object} projectActualDailyHours - An object like { "YYYY-MM-DD": hours } for a specific project.
 * @param {number} projectPlannedTotalHours - The total monthly hours allocated for this project.
 * @returns {Object} - An object containing forecast text, class, and optionally exhaustionDate.
 */
function forecastProjectStatus(projectActualDailyHours, projectPlannedTotalHours) {
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth();

    const dailyRate = calculateDailyConsumptionRate(projectActualDailyHours, now);
    const remainingBusinessDays = getRemainingBusinessDays(currentYear, currentMonth, now);

    let currentMonthActuals = 0;
    for (const dateString in projectActualDailyHours) {
        const dailyDate = new Date(dateString);
        if (
            !isNaN(dailyDate) &&
            dailyDate.getFullYear() === currentYear &&
            dailyDate.getMonth() === currentMonth &&
            dailyDate <= now
        ) {
            currentMonthActuals += projectActualDailyHours[dateString];
        }
    }

    if (dailyRate === 0 && currentMonthActuals === 0) {
        return { text: "No data yet", class: "forecast-neutral", exhaustionDate: null };
    }

    if (projectPlannedTotalHours === 0) {
        return { text: "Unplanned", class: "forecast-unplanned", exhaustionDate: null };
    }

    const projectedFutureConsumption = dailyRate * remainingBusinessDays;
    const projectedRemainingHours = projectPlannedTotalHours - currentMonthActuals - projectedFutureConsumption;

    let forecastText = "On track";
    let forecastClass = "forecast-neutral";
    let exhaustionDate = null;
    const buffer = 2;

    const formatExhaustionDate = (dateString) => {
        const date = new Date(dateString);
        if (isNaN(date)) return dateString;
        const options = { weekday: 'short', day: 'numeric' };
        const formatted = date.toLocaleDateString('en-US', options);
        const day = date.getDate();
        let suffix = 'th';
        if (day === 1 || day === 21 || day === 31) {
            suffix = 'st';
        } else if (day === 2 || day === 22) {
            suffix = 'nd';
        } else if (day === 3 || day === 23) {
            suffix = 'rd';
        }
        return formatted.replace(/, (\d+)$/, ` ${day}${suffix}`);
    };

    if (currentMonthActuals > projectPlannedTotalHours && dailyRate > 0) {
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
            forecastText = `Over budget by ${formatHours(currentMonthActuals - projectPlannedTotalHours)}`;
            forecastClass = "forecast-over";
        }
    } else if (projectedRemainingHours > buffer) {
        forecastText = `Increase consumption by ${formatHours(projectedRemainingHours / remainingBusinessDays)} hrs/day`;
        forecastClass = "forecast-over";
    } else if (projectedRemainingHours < -buffer) {
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
            tempDate.setDate(tempDate.getDate() + 1);

            while (tempDate.getMonth() === currentMonth) {
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
            forecastText = `Will run out by ${formatHours(Math.abs(projectedRemainingHours))}`;
        }
        forecastClass = "forecast-under";
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
 * Gets the start and end dates of the current week (Monday to Sunday).
 * @param {Date} date
 * @returns {{start: Date, end: Date}}
 */
function getWeekRange(date) {
    const d = new Date(date);
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1);

    const start = new Date(d.setDate(diff));
    start.setHours(0, 0, 0, 0);

    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    end.setHours(23, 59, 59, 999);

    return { start, end };
}

/**
 * Checks if a date string (YYYY-MM-DD) falls within a date range.
 * @param {string} dateStr
 * @param {Date} rangeStart
 * @param {Date} rangeEnd
 * @returns {boolean}
 */
function isDateInRange(dateStr, rangeStart, rangeEnd) {
    const d = new Date(dateStr);
    d.setHours(0, 0, 0, 0);
    const start = new Date(rangeStart);
    start.setHours(0, 0, 0, 0);
    const end = new Date(rangeEnd);
    end.setHours(0, 0, 0, 0);

    return d >= start && d <= end;
}

// -----------------------------------------------------------------------------

export default class extends Controller {
    static targets = ["tableBody", "totalActual", "totalPlanned", "totalDiff", "totalActualPlanned", "unplannedWork", "viewModeToggle"];

    connect() {
        console.log("Dashboard controller connected!");
        this.loadSettings();
        this.renderDashboard();
    }

    // --- State ---
    plannedHours = {};
    actualHours = {};
    currentViewMode = 'monthly'; // Default

    loadSettings() {
        const storedPlannedHours = localStorage.getItem('plannedHours');
        const storedActualHours = localStorage.getItem('actualHours');

        if (storedPlannedHours) {
            const result = JSON.parse(storedPlannedHours);
            // Migration Logic: Convert old number format to new object format
            const migrated = {};
            for (const [key, value] of Object.entries(result)) {
                if (typeof value === 'number') {
                    migrated[key] = {
                        type: 'monthly',
                        total: value,
                        weeks: []
                    };
                } else {
                    migrated[key] = value;
                }
            }
            this.plannedHours = migrated;
        } else {
            // Default / Example data
            this.plannedHours = {
                "Example Project": { type: 'monthly', total: 40, weeks: [] }
            };
        }
        if (storedActualHours) {
            this.actualHours = JSON.parse(storedActualHours);
        }
        // Initialize view mode from toggle if available
        if (this.hasViewModeToggleTarget) {
            this.currentViewMode = this.viewModeToggleTarget.value;
        }
    }

    refreshData() {
        // Clear stored actuals if a new upload is expected, or simply re-render with current data
        // For now, we just re-render with existing data as there's no "scrape" anymore.
        this.renderDashboard();
    }

    changeViewMode() {
        this.currentViewMode = this.viewModeToggleTarget.value;
        this.renderDashboard();
    }

    renderDashboard() {
        this.tableBodyTarget.innerHTML = ''; // Clear table body

        const isWeekly = this.currentViewMode === 'weekly';
        const today = new Date();
        const { start: weekStart, end: weekEnd } = getWeekRange(today);

        const allProjects = new Set([
            ...Object.keys(this.plannedHours),
            ...Object.keys(this.actualHours || {}).filter(key => key !== 'meta')
        ]);

        const sortedProjects = Array.from(allProjects).sort();

        let totalActual = 0;
        let totalPlanned = 0;
        let totalActualPlannedWork = 0;
        let unplannedWork = 0;

        if (sortedProjects.length === 0 && Object.keys(this.actualHours).filter(key => key !== 'meta').length === 0) {
            this.tableBodyTarget.innerHTML = '<tr><td colspan="5" style="text-align: center; color: #666;">No data to display. Add projects in Settings or upload an XLSX file.</td></tr>';
        } else {
            sortedProjects.forEach(proj => {
                let actual = 0;
                let planned = 0;
                let forecast = { text: "N/A", class: "forecast-neutral" };

                // --- 1. Calculate Actuals ---
                if (this.actualHours && this.actualHours[proj]) {
                    if (isWeekly) {
                        if (this.actualHours[proj].dailyHours) {
                            for (const dateStr in this.actualHours[proj].dailyHours) {
                                if (isDateInRange(dateStr, weekStart, weekEnd)) {
                                    actual += this.actualHours[proj].dailyHours[dateStr];
                                }
                            }
                        } else {
                            actual = 0;
                        }
                    } else {
                        if (typeof this.actualHours[proj] === 'number') {
                            actual = this.actualHours[proj];
                        } else if (this.actualHours[proj].total !== undefined) {
                            actual = this.actualHours[proj].total;
                        } else if (this.actualHours[proj].totalHours !== undefined) {
                            actual = this.actualHours[proj].totalHours;
                        }
                    }
                }

                // --- 2. Calculate Planned ---
                const plannedData = this.plannedHours[proj];
                if (plannedData) {
                    if (isWeekly) {
                        if (plannedData.type === 'weekly' && plannedData.weeks) {
                            const weekEntry = plannedData.weeks.find(w => {
                                return isDateInRange(formatDate(today), new Date(w.start), new Date(w.end));
                            });
                            if (weekEntry) {
                                planned = weekEntry.hours;
                            } else {
                                planned = 0;
                            }
                        } else {
                            planned = 0;
                        }
                    } else {
                        planned = plannedData.total || 0;
                    }
                }

                const diff = actual - planned;

                totalActual += actual;
                totalPlanned += planned;

                if (actual > 0 && planned > 0) {
                    totalActualPlannedWork += actual;
                }

                if (actual > planned) {
                    unplannedWork += (actual - planned);
                }

                // Forecast logic (Only for Monthly view for now)
                if (!isWeekly) {
                    const projectActualDailyHours = (this.actualHours[proj] && this.actualHours[proj].dailyHours) ?
                        this.actualHours[proj].dailyHours : {};
                    forecast = forecastProjectStatus(projectActualDailyHours, planned);
                } else {
                    if (actual > planned) {
                        forecast = { text: "Over weekly limit", class: "forecast-over" };
                    } else if (actual < planned) {
                        forecast = { text: "Under weekly limit", class: "forecast-under" };
                    } else {
                        forecast = { text: "On track", class: "forecast-neutral" };
                    }
                }

                const tr = document.createElement('tr');
                let forecastDisplay = forecast.text;
                if (forecast.exhaustionDate) {
                    forecastDisplay += ` (${this.formatExhaustionDate(forecast.exhaustionDate)})`;
                }

                tr.innerHTML = `
                    <td>${proj}</td>
                    <td>${formatHours(actual)}</td>
                    <td>${formatHours(planned)}</td>
                    <td class="${diff > 0 ? 'diff-pos' : (diff < 0 ? 'diff-neg' : '')}">${diff > 0 ? '+' : ''}${formatHours(diff)}</td>
                    <td class="${forecast.class}">${forecastDisplay}</td>
                `;
                this.tableBodyTarget.appendChild(tr);
            });
        }

        this.totalActualTarget.textContent = formatHours(totalActual);
        this.totalPlannedTarget.textContent = formatHours(totalPlanned);

        const totalActualPlannedEl = this.totalActualPlannedTarget;
        totalActualPlannedEl.textContent = formatHours(totalActualPlannedWork);
        let actualPlannedClass = '';
        const plannedPerformanceDiff = totalActualPlannedWork - totalPlanned;
        const amberBuffer = isWeekly ? 2 : 5;

        if (plannedPerformanceDiff >= 0) {
            actualPlannedClass = 'actual-planned-green';
        } else if (plannedPerformanceDiff > -amberBuffer) {
            actualPlannedClass = 'actual-planned-amber';
        } else {
            actualPlannedClass = 'actual-planned-red';
        }
        totalActualPlannedEl.className = 'value ' + actualPlannedClass;

        const unplannedWorkEl = this.unplannedWorkTarget;
        unplannedWorkEl.textContent = formatHours(unplannedWork);
        unplannedWorkEl.className = 'value unplanned-red-text';

        const totalDiff = totalActual - totalPlanned;
        const diffEl = this.totalDiffTarget;
        diffEl.textContent = (totalDiff > 0 ? '+' : '') + formatHours(totalDiff);
        diffEl.className = 'value ' + (totalDiff > 0 ? 'diff-pos' : (totalDiff < 0 ? 'diff-neg' : ''));
    }

    // Helper to format date to "Mon 16th" - moved here as it's directly used by renderDashboard
    formatExhaustionDate(dateString) {
        const date = new Date(dateString);
        if (isNaN(date)) return dateString;
        const options = { weekday: 'short', day: 'numeric' };
        const formatted = date.toLocaleDateString('en-US', options);
        const day = date.getDate();
        let suffix = 'th';
        if (day === 1 || day === 21 || day === 31) {
            suffix = 'st';
        } else if (day === 2 || day === 22) {
            suffix = 'nd';
        } else if (day === 3 || day === 23) {
            suffix = 'rd';
        }
        return formatted.replace(/, (\d+)$/, ` ${day}${suffix}`);
    }
}