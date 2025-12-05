document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const tabDashboard = document.getElementById('tab-dashboard');
    const tabSettings = document.getElementById('tab-settings');
    const viewDashboard = document.getElementById('view-dashboard');
    const viewSettings = document.getElementById('view-settings');
    const tableBody = document.querySelector('#analysis-table tbody');
    const settingsForm = document.getElementById('settings-form');
    const projectsList = document.getElementById('projects-list');
    const btnAddProject = document.getElementById('btn-add-project');
    const btnRefresh = document.getElementById('btn-refresh');
    const btnDebug = document.getElementById('btn-debug');
    const viewModeToggle = document.getElementById('view-mode-toggle');

    // State
    let plannedHours = {};
    let actualHours = {}; // To hold data from XLSX
    let currentViewMode = 'monthly';

    // --- Initialization ---
    loadSettings().then(() => {
        renderSettings();
        // Auto-load data if on dashboard
        if (tabDashboard.classList.contains('active')) {
            refreshDashboard();
        }
    });

    // --- Event Listeners ---
    tabDashboard.addEventListener('click', () => {
        switchTab('dashboard');
        refreshDashboard(); // Refresh data when switching to the dashboard
    });
    tabSettings.addEventListener('click', () => switchTab('settings'));

    if (viewModeToggle) {
        viewModeToggle.addEventListener('change', (e) => {
            currentViewMode = e.target.value;
            renderTable(actualHours);
        });
    }

    function switchTab(tabName) {
        if (tabName === 'dashboard') {
            tabDashboard.classList.add('active');
            tabSettings.classList.remove('active');
            viewDashboard.classList.add('active');
            viewSettings.classList.remove('active');
        } else { // 'settings'
            tabDashboard.classList.remove('active');
            tabSettings.classList.add('active');
            viewDashboard.classList.remove('active');
            viewSettings.classList.add('active');
        }
    }

    // --- Add Project Logic ---
    const newProjectMode = document.getElementById('new-project-mode');
    const newProjectHours = document.getElementById('new-project-hours');

    if (newProjectMode && newProjectHours) {
        newProjectMode.addEventListener('change', (e) => {
            if (e.target.value === 'weekly') {
                newProjectHours.value = '';
                newProjectHours.disabled = true;
                newProjectHours.placeholder = 'N/A';
                newProjectHours.style.backgroundColor = '#f1f5f9';
            } else {
                newProjectHours.disabled = false;
                newProjectHours.placeholder = 'Hours';
                newProjectHours.style.backgroundColor = '';
            }
        });
    }

    btnAddProject.addEventListener('click', () => {
        const name = document.getElementById('new-project-name').value.trim();
        const mode = document.getElementById('new-project-mode').value;
        let hours = 0;

        if (mode === 'monthly') {
            hours = parseFloat(document.getElementById('new-project-hours').value);
        }

        if (name) {
            if (mode === 'monthly' && isNaN(hours)) {
                alert("Please enter valid hours for Monthly mode.");
                return;
            }

            plannedHours[name] = {
                type: mode,
                total: mode === 'monthly' ? hours : 0,
                weeks: []
            };

            document.getElementById('new-project-name').value = '';
            document.getElementById('new-project-hours').value = '';
            // Reset mode to monthly
            if (newProjectMode) {
                newProjectMode.value = 'monthly';
                newProjectMode.dispatchEvent(new Event('change'));
            }

            renderSettings();
            saveSettings();
        }
    });

    settingsForm.addEventListener('submit', (e) => {
        e.preventDefault();
        saveSettings();
        switchTab('dashboard');
        refreshDashboard();
    });

    btnRefresh.addEventListener('click', () => {
        // Clear stored actuals from XLSX before scraping anew
        actualHours = {};
        chrome.storage.local.remove('actualHours', refreshDashboard);
    });

    btnDebug.addEventListener('click', () => {
        chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
            const activeTab = tabs[0];
            if (!activeTab) return;

            // Use executeScript to get HTML directly, avoiding message passing issues
            chrome.scripting.executeScript({
                target: { tabId: activeTab.id },
                func: () => document.body.innerHTML
            }, (results) => {
                if (chrome.runtime.lastError) {
                    console.error(chrome.runtime.lastError);
                    alert("Error: Cannot access page. Make sure you are on the Fusion website.");
                    return;
                }

                if (results && results[0] && results[0].result) {
                    const html = results[0].result;
                    navigator.clipboard.writeText(html).then(() => {
                        alert("Page HTML copied to clipboard! Please paste this to the developer.");
                    });
                }
            });
        });
    });

    // --- XLSX Upload Handling (Actuals) ---
    const btnUpload = document.getElementById('btn-upload-xlsx');
    const fileInput = document.getElementById('xlsx-upload');
    const uploadStatus = document.getElementById('upload-status');

    if (btnUpload) {
        btnUpload.addEventListener('click', () => {
            fileInput.click();
        });
    }

    if (fileInput) {
        fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;

            uploadStatus.textContent = 'Reading file...';

            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    const parsedData = parseXLSXData(jsonData);

                    // Store in state and chrome.storage
                    actualHours = parsedData;
                    chrome.storage.local.set({ actualHours: actualHours }, () => {
                        const projectCount = Object.keys(actualHours).filter(key => key !== 'meta').length;
                        const timestamp = new Date().toLocaleString();
                        uploadStatus.textContent = `✓ Uploaded ${projectCount} projects at ${timestamp}`;
                        uploadStatus.style.color = '#28a745';

                        // Render the table with the data we just parsed
                        renderTable(actualHours);
                        // Switch to dashboard view to show the result
                        switchTab('dashboard');
                    });
                } catch (error) {
                    console.error('Error parsing XLSX:', error);
                    uploadStatus.textContent = '✗ Error parsing file. Please try again.';
                    uploadStatus.style.color = '#dc3545';
                }
            };

            reader.onerror = () => {
                uploadStatus.textContent = '✗ Error reading file.';
                uploadStatus.style.color = '#dc3545';
            };

            reader.readAsArrayBuffer(file);
        });
    }

    function loadSettings() {
        return new Promise((resolve) => {
            chrome.storage.local.get(['plannedHours', 'actualHours'], (result) => {
                if (result.plannedHours) {
                    // Migration Logic: Convert old number format to new object format
                    const migrated = {};
                    for (const [key, value] of Object.entries(result.plannedHours)) {
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
                    plannedHours = migrated;
                } else {
                    // Default / Example data
                    plannedHours = {
                        "Example Project": { type: 'monthly', total: 40, weeks: [] }
                    };
                }
                if (result.actualHours) {
                    actualHours = result.actualHours;
                }
                resolve();
            });
        });
    }

    function saveSettings() {
        chrome.storage.local.set({ plannedHours: plannedHours }, () => {
            console.log("Settings saved");
        });
    }

    function renderSettings() {
        projectsList.innerHTML = '';
        const sortedProjects = Object.keys(plannedHours).sort();

        sortedProjects.forEach(proj => {
            const projectData = plannedHours[proj];

            const container = document.createElement('div');
            container.className = 'project-container';
            container.style.border = '1px solid #e2e8f0';
            container.style.borderRadius = '6px';
            container.style.marginBottom = '10px';
            container.style.padding = '10px';
            container.style.background = '#fff';

            // --- Header Row ---
            const headerRow = document.createElement('div');
            headerRow.style.display = 'flex';
            headerRow.style.alignItems = 'center';
            headerRow.style.gap = '10px';
            headerRow.style.marginBottom = '8px';

            // Name Input (Read-onlyish)
            const nameInput = document.createElement('input');
            nameInput.type = 'text';
            nameInput.value = proj;
            nameInput.readOnly = true;
            nameInput.style.flex = '1';
            nameInput.style.fontWeight = '500';
            nameInput.style.border = 'none';
            nameInput.style.background = 'transparent';

            // Mode Toggle
            const modeSelect = document.createElement('select');
            modeSelect.style.padding = '4px';
            modeSelect.style.fontSize = '12px';
            modeSelect.style.borderRadius = '4px';
            modeSelect.style.border = '1px solid #ccc';

            const optMonthly = document.createElement('option');
            optMonthly.value = 'monthly';
            optMonthly.text = 'Monthly';
            const optWeekly = document.createElement('option');
            optWeekly.value = 'weekly';
            optWeekly.text = 'Weekly';

            modeSelect.add(optMonthly);
            modeSelect.add(optWeekly);
            modeSelect.value = projectData.type || 'monthly';

            modeSelect.addEventListener('change', (e) => {
                projectData.type = e.target.value;
                if (projectData.type === 'monthly') {
                    // If switching back to monthly, maybe keep total as is?
                } else {
                    // Switching to weekly, total becomes sum of weeks (initially 0 if no weeks)
                    recalculateTotal(proj);
                }
                renderSettings();
                saveSettings();
            });

            // Hours Input (Visible if Monthly)
            const hoursInput = document.createElement('input');
            hoursInput.type = 'number';
            hoursInput.step = '0.5';
            hoursInput.style.width = '70px';
            hoursInput.style.padding = '4px';
            hoursInput.style.border = '1px solid #ccc';
            hoursInput.style.borderRadius = '4px';

            if (projectData.type === 'monthly') {
                hoursInput.value = projectData.total;
                hoursInput.addEventListener('change', (e) => {
                    const val = parseFloat(e.target.value);
                    if (!isNaN(val)) {
                        projectData.total = val;
                        saveSettings();
                    }
                });
            } else {
                hoursInput.value = projectData.total;
                hoursInput.readOnly = true;
                hoursInput.style.background = '#f1f5f9';
                hoursInput.title = "Calculated from weeks";
            }

            // Remove Button
            const removeBtn = document.createElement('button');
            removeBtn.textContent = '×';
            removeBtn.style.background = 'none';
            removeBtn.style.border = 'none';
            removeBtn.style.color = '#94a3b8';
            removeBtn.style.fontSize = '18px';
            removeBtn.style.cursor = 'pointer';
            removeBtn.addEventListener('click', () => {
                if (confirm(`Remove project "${proj}"?`)) {
                    delete plannedHours[proj];
                    renderSettings();
                    saveSettings();
                }
            });

            headerRow.appendChild(nameInput);
            headerRow.appendChild(modeSelect);
            headerRow.appendChild(hoursInput);
            headerRow.appendChild(removeBtn);
            container.appendChild(headerRow);

            // --- Weekly Editor (Visible if Weekly) ---
            if (projectData.type === 'weekly') {
                const weeklyEditor = document.createElement('div');
                weeklyEditor.style.marginTop = '8px';
                weeklyEditor.style.paddingTop = '8px';
                weeklyEditor.style.borderTop = '1px solid #f1f5f9';

                // List existing weeks
                if (projectData.weeks && projectData.weeks.length > 0) {
                    projectData.weeks.forEach((week, index) => {
                        const weekRow = document.createElement('div');
                        weekRow.style.display = 'flex';
                        weekRow.style.alignItems = 'center';
                        weekRow.style.gap = '5px';
                        weekRow.style.marginBottom = '4px';
                        weekRow.style.fontSize = '12px';

                        const dateRange = document.createElement('span');
                        dateRange.textContent = `${week.start} to ${week.end}`;
                        dateRange.style.flex = '1';

                        const weekHours = document.createElement('span');
                        weekHours.textContent = `${week.hours}h`;
                        weekHours.style.fontWeight = '600';

                        const delWeekBtn = document.createElement('button');
                        delWeekBtn.textContent = '×';
                        delWeekBtn.style.border = 'none';
                        delWeekBtn.style.background = 'none';
                        delWeekBtn.style.color = '#ef4444';
                        delWeekBtn.style.cursor = 'pointer';
                        delWeekBtn.addEventListener('click', () => {
                            projectData.weeks.splice(index, 1);
                            recalculateTotal(proj);
                            renderSettings();
                            saveSettings();
                        });

                        weekRow.appendChild(dateRange);
                        weekRow.appendChild(weekHours);
                        weekRow.appendChild(delWeekBtn);
                        weeklyEditor.appendChild(weekRow);
                    });
                } else {
                    const emptyMsg = document.createElement('div');
                    emptyMsg.textContent = "No weeks added yet.";
                    emptyMsg.style.fontSize = '12px';
                    emptyMsg.style.color = '#94a3b8';
                    emptyMsg.style.fontStyle = 'italic';
                    emptyMsg.style.marginBottom = '8px';
                    weeklyEditor.appendChild(emptyMsg);
                }

                // Add New Week Form
                const addRow = document.createElement('div');
                addRow.style.display = 'flex';
                addRow.style.gap = '5px';
                addRow.style.marginTop = '8px';

                const startInput = document.createElement('input');
                startInput.type = 'date';
                startInput.style.width = '110px'; // wider for date
                startInput.style.fontSize = '11px';

                const endInput = document.createElement('input');
                endInput.type = 'date';
                endInput.style.width = '110px';
                endInput.style.fontSize = '11px';

                const wHoursInput = document.createElement('input');
                wHoursInput.type = 'number';
                wHoursInput.placeholder = 'Hrs';
                wHoursInput.step = '0.5';
                wHoursInput.style.width = '50px';
                wHoursInput.style.fontSize = '12px';

                const addWeekBtn = document.createElement('button');
                addWeekBtn.textContent = '+';
                addWeekBtn.style.padding = '2px 8px';
                addWeekBtn.style.background = '#2563eb';
                addWeekBtn.style.color = 'white';
                addWeekBtn.style.border = 'none';
                addWeekBtn.style.borderRadius = '4px';
                addWeekBtn.style.cursor = 'pointer';

                addWeekBtn.addEventListener('click', () => {
                    const s = startInput.value;
                    const e = endInput.value;
                    const h = parseFloat(wHoursInput.value);

                    if (s && e && !isNaN(h)) {
                        if (!projectData.weeks) projectData.weeks = [];
                        projectData.weeks.push({
                            start: s,
                            end: e,
                            hours: h
                        });
                        // Sort weeks by start date
                        projectData.weeks.sort((a, b) => new Date(a.start) - new Date(b.start));

                        recalculateTotal(proj);
                        renderSettings();
                        saveSettings();
                    } else {
                        alert("Please fill in Start Date, End Date, and Hours.");
                    }
                });

                addRow.appendChild(startInput);
                addRow.appendChild(endInput);
                addRow.appendChild(wHoursInput);
                addRow.appendChild(addWeekBtn);
                weeklyEditor.appendChild(addRow);

                container.appendChild(weeklyEditor);
            }

            projectsList.appendChild(container);
        });
    }

    function recalculateTotal(projName) {
        const data = plannedHours[projName];
        if (data.type === 'weekly' && data.weeks) {
            const sum = data.weeks.reduce((acc, curr) => acc + (curr.hours || 0), 0);
            data.total = sum;
        }
    }

    function refreshDashboard() {
        // 1. First, check if we have actuals from a previous XLSX upload
        if (actualHours && Object.keys(actualHours).filter(key => key !== 'meta').length > 0) {
            console.log("Rendering dashboard from stored XLSX data.");
            renderTable(actualHours);
            return; // Don't proceed to scrape
        }

        // 2. If not, try to get Actuals from Content Script
        console.log("No stored XLSX data found. Attempting to scrape from content script.");
        chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
            const activeTab = tabs[0];
            if (!activeTab || !activeTab.id) {
                renderTable({}); // Clear table if tab is not accessible
                return;
            }

            chrome.scripting.executeScript({
                target: { tabId: activeTab.id },
                files: ['scripts/utils.js', 'scripts/content.js']
            }, () => {
                if (chrome.runtime.lastError) {
                    renderTable({}); // Clear table on injection error
                    return;
                }

                chrome.tabs.sendMessage(activeTab.id, { action: "scrape_hours" }, (response) => {
                    if (chrome.runtime.lastError) {
                        renderTable({}); // Clear table if message fails
                        return;
                    }

                    if (response && response.success) {
                        actualHours = response.data; // Cache scraped data
                        renderTable(actualHours);
                    } else if (response && response.error === "WRONG_PAGE_DASHBOARD") {
                        tableBody.innerHTML = '<tr><td colspan="4" style="text-align: center; color: #d9534f; padding: 20px;"><b>Wrong Page Detected</b><br>You are on the Dashboard.<br>Please navigate to the <b>Daily Timesheet</b> page and click Refresh.</td></tr>';
                    } else {
                        renderTable({});
                    }
                });
            });
        });
    }

    function renderTable(currentActuals) {
        tableBody.innerHTML = '';

        // Determine timeframe
        const isWeekly = currentViewMode === 'weekly';
        const today = new Date();
        const { start: weekStart, end: weekEnd } = getWeekRange(today);

        // Merge keys
        const allProjects = new Set([
            ...Object.keys(plannedHours),
            ...Object.keys(currentActuals || {}).filter(key => key !== 'meta')
        ]);

        const sortedProjects = Array.from(allProjects).sort();

        let totalActual = 0;
        let totalPlanned = 0;
        let totalActualPlannedWork = 0;
        let unplannedWork = 0;

        if (sortedProjects.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="5" style="text-align: center; color: #666;">No data to display. Add projects in Settings or refresh on a Fusion page.</td></tr>';
        } else {
            sortedProjects.forEach(proj => {
                let actual = 0;
                let planned = 0;
                let forecast = { text: "N/A", class: "forecast-neutral" };

                // --- 1. Calculate Actuals ---
                if (currentActuals && currentActuals[proj]) {
                    if (isWeekly) {
                        // Calculate actuals for the current week only
                        if (currentActuals[proj].dailyHours) {
                            for (const dateStr in currentActuals[proj].dailyHours) {
                                if (isDateInRange(dateStr, weekStart, weekEnd)) {
                                    actual += currentActuals[proj].dailyHours[dateStr];
                                }
                            }
                        } else {
                            // If we only have total scraped data (no daily breakdown), we can't show weekly actuals accurately
                            actual = 0;
                        }
                    } else {
                        // Monthly / Total
                        if (typeof currentActuals[proj] === 'number') {
                            actual = currentActuals[proj];
                        } else if (currentActuals[proj].total !== undefined) {
                            actual = currentActuals[proj].total;
                        } else if (currentActuals[proj].totalHours !== undefined) {
                            actual = currentActuals[proj].totalHours;
                        }
                    }
                }

                // --- 2. Calculate Planned ---
                const plannedData = plannedHours[proj];
                if (plannedData) {
                    if (isWeekly) {
                        // Find planned hours for the current week
                        if (plannedData.type === 'weekly' && plannedData.weeks) {
                            // Find the week entry that overlaps with current week
                            const weekEntry = plannedData.weeks.find(w => {
                                // Check if today is within the range [w.start, w.end]
                                return isDateInRange(formatDate(today), new Date(w.start), new Date(w.end));
                            });

                            if (weekEntry) {
                                planned = weekEntry.hours;
                            } else {
                                planned = 0;
                            }
                        } else {
                            // Monthly mode project in Weekly View:
                            // We don't have a breakdown. Show 0? Or pro-rate?
                            // Let's show 0 and maybe a tooltip or just 0.
                            planned = 0;
                        }
                    } else {
                        // Monthly View
                        planned = plannedData.total || 0;
                    }
                }

                const diff = actual - planned;

                totalActual += actual;
                totalPlanned += planned;

                if (actual > 0 && planned > 0) {
                    totalActualPlannedWork += actual;
                }

                if (planned === 0 && actual > 0) {
                    unplannedWork += actual;
                }

                // Forecast logic (Only for Monthly view for now)
                if (!isWeekly) {
                    const projectActualDailyHours = (currentActuals[proj] && currentActuals[proj].dailyHours) ?
                        currentActuals[proj].dailyHours : {};
                    forecast = forecastProjectStatus(projectActualDailyHours, planned);
                } else {
                    // Weekly forecast
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
                    forecastDisplay += ` (${forecast.exhaustionDate})`;
                }

                tr.innerHTML = `
                    <td>${proj}</td>
                    <td>${formatHours(actual)}</td>
                    <td>${formatHours(planned)}</td>
                    <td class="${diff > 0 ? 'diff-pos' : (diff < 0 ? 'diff-neg' : '')}">${diff > 0 ? '+' : ''}${formatHours(diff)}</td>
                    <td class="${forecast.class}">${forecastDisplay}</td>
                `;
                tableBody.appendChild(tr);
            });
        }

        document.getElementById('total-actual').textContent = formatHours(totalActual);
        document.getElementById('total-planned').textContent = formatHours(totalPlanned);

        // Apply color coding to Actual Planned
        const totalActualPlannedEl = document.getElementById('total-actual-planned');
        totalActualPlannedEl.textContent = formatHours(totalActualPlannedWork);
        let actualPlannedClass = '';
        const plannedPerformanceDiff = totalActualPlannedWork - totalPlanned;
        const amberBuffer = isWeekly ? 2 : 5; // Smaller buffer for weekly view

        if (plannedPerformanceDiff >= 0) {
            actualPlannedClass = 'actual-planned-green'; // Ahead or on par
        } else if (plannedPerformanceDiff > -amberBuffer) {
            actualPlannedClass = 'actual-planned-amber'; // Slightly behind
        } else {
            actualPlannedClass = 'actual-planned-red'; // Lagging badly
        }
        totalActualPlannedEl.className = 'value ' + actualPlannedClass;

        const unplannedWorkEl = document.getElementById('unplanned-work');
        unplannedWorkEl.textContent = formatHours(unplannedWork);
        unplannedWorkEl.className = 'value unplanned-red-text'; // Always red for unplanned

        const totalDiff = totalActual - totalPlanned;
        const diffEl = document.getElementById('total-diff');
        diffEl.textContent = (totalDiff > 0 ? '+' : '') + formatHours(totalDiff);
        diffEl.className = 'value ' + (totalDiff > 0 ? 'diff-pos' : (totalDiff < 0 ? 'diff-neg' : ''));
    }
});