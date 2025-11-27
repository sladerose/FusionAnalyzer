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

    // State
    let plannedHours = {};
    let actualHours = {}; // To hold data from XLSX

    // --- Initialization ---
    loadSettings().then(() => {
        renderSettings();
        // Auto-load data if on dashboard
        refreshDashboard();
    });

    // --- Event Listeners ---
    tabDashboard.addEventListener('click', () => {
        switchTab('dashboard');
        refreshDashboard(); // Refresh data when switching to the dashboard
    });
    tabSettings.addEventListener('click', () => switchTab('settings'));

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

    btnAddProject.addEventListener('click', () => {
        const name = document.getElementById('new-project-name').value.trim();
        const hours = parseFloat(document.getElementById('new-project-hours').value);

        if (name && !isNaN(hours)) {
            plannedHours[name] = hours;
            document.getElementById('new-project-name').value = '';
            document.getElementById('new-project-hours').value = '';
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

    // XLSX Upload Handling
    const btnUpload = document.getElementById('btn-upload-xlsx');
    const fileInput = document.getElementById('xlsx-upload');
    const uploadStatus = document.getElementById('upload-status');

    btnUpload.addEventListener('click', () => {
        fileInput.click();
    });

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

                const parsedHours = parseXLSXData(jsonData);

                // Store in state and chrome.storage
                actualHours = parsedHours;
                chrome.storage.local.set({ actualHours: actualHours }, () => {
                    const projectCount = Object.keys(actualHours).length;
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
    }); // Closing for fileInput.addEventListener

    function loadSettings() {
        return new Promise((resolve) => {
            chrome.storage.local.get(['plannedHours', 'actualHours'], (result) => {
                if (result.plannedHours) {
                    plannedHours = result.plannedHours;
                } else {
                    // Default / Example data
                    plannedHours = {
                        "Example Project": 40
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
            const div = document.createElement('div');
            div.className = 'project-item';

            // Create Input for Name
            const nameInput = document.createElement('input');
            nameInput.type = 'text';
            nameInput.value = proj;
            nameInput.readOnly = true;

            // Create Input for Hours
            const hoursInput = document.createElement('input');
            hoursInput.type = 'number';
            hoursInput.value = plannedHours[proj];
            hoursInput.step = '0.5';
            hoursInput.addEventListener('change', (e) => {
                const val = parseFloat(e.target.value);
                if (!isNaN(val)) {
                    plannedHours[proj] = val;
                    saveSettings();
                }
            });

            // Create Remove Button
            const removeBtn = document.createElement('button');
            removeBtn.type = 'button';
            removeBtn.className = 'remove-btn';
            removeBtn.textContent = '×';
            removeBtn.addEventListener('click', () => {
                delete plannedHours[proj];
                renderSettings();
                saveSettings();
            });

            div.appendChild(nameInput);
            div.appendChild(hoursInput);
            div.appendChild(removeBtn);
            projectsList.appendChild(div);
        });
    }

    function refreshDashboard() {
        // 1. First, check if we have actuals from a previous XLSX upload
        if (actualHours && Object.keys(actualHours).length > 0) {
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
                        renderTable(response.data);
                    } else {
                        renderTable({});
                    }
                });
            });
        });
    }

    function renderTable(currentActuals) {
        tableBody.innerHTML = '';

        // Merge keys
        const allProjects = new Set([
            ...Object.keys(plannedHours),
            ...Object.keys(currentActuals || {})
        ]);

        const sortedProjects = Array.from(allProjects).sort();

        let totalActual = 0;
        let totalPlanned = 0;

        if (sortedProjects.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="4" style="text-align: center; color: #666;">No data to display. Upload an XLSX file or refresh on a Fusion page.</td></tr>';
        } else {
            sortedProjects.forEach(proj => {
                const actual = (currentActuals || {})[proj] || 0;
                const planned = plannedHours[proj] || 0;
                const diff = actual - planned;

                totalActual += actual;
                totalPlanned += planned;

                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${proj}</td>
                    <td>${formatHours(actual)}</td>
                    <td>${formatHours(planned)}</td>
                    <td class="${diff > 0 ? 'diff-pos' : (diff < 0 ? 'diff-neg' : '')}">${diff > 0 ? '+' : ''}${formatHours(diff)}</td>
                `;
                tableBody.appendChild(tr);
            });
        }

        document.getElementById('total-actual').textContent = formatHours(totalActual);
        document.getElementById('total-planned').textContent = formatHours(totalPlanned);

        const totalDiff = totalActual - totalPlanned;
        const diffEl = document.getElementById('total-diff');
        diffEl.textContent = (totalDiff > 0 ? '+' : '') + formatHours(totalDiff);
        diffEl.className = 'value ' + (totalDiff > 0 ? 'diff-pos' : (totalDiff < 0 ? 'diff-neg' : ''));
    }
});
