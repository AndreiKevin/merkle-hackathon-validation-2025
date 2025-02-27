// Constants and configurations
const CONFIG = {
    propertyId: "189677427",
    menuTitle: "Post Launch Validator",
    dateRange: {
        startDate: "7daysAgo",
        endDate: "yesterday"
    }
};

const HEADERS = [
    { text: "Account Name", column: 1 },
    { text: "Account ID", column: 2 },
    { text: "Property Name", column: 3 },
    { text: "Property ID", column: 4 },
    { text: "Property Create Time", column: 5 },
    { text: "Time Zone", column: 6 },
    { text: "Display Name", column: 7 },
    { text: "Data Streams", column: 8 },
];

// Helper function to get metric colors (same as used in writeReportsToSheet)
function getMetricColor(index) {
    const metricColors = [
        '#4285F4',  // Google Blue
        '#0F9D58',  // Google Green
        '#DB4437',  // Google Red
        '#F4B400',  // Google Yellow
        '#AA47BC',  // Purple
        '#00ACC1',  // Cyan
        '#FF7043',  // Deep Orange
        '#9E9D24',  // Lime
        '#5C6BC0',  // Indigo
        '#00796B',  // Teal
    ];
    return metricColors[index % metricColors.length];
}

// Utility functions
function handleError(error, context) {
    Logger.log(`Error in ${context}: ${error.toString()}`);
    if (error.error) {
        console.log(`Failed with error: ${error.error}`);
    }
    throw error;
}

function extractId(name, type) {
    return name.split("/")[1];
}

// Sheet operations
function initializeSheet(sheet) {
    if (!sheet) return;
    sheet.clear();
    HEADERS.forEach((header) => {
        sheet
            .getRange(1, header.column)
            .setValue(header.text)
            .setFontWeight("bold")
            .setHorizontalAlignment("center");
    });
}

function writeDataToSheet(sheet, rowData) {
    if (!rowData || rowData.length === 0) return;
    const columnCount = rowData[0].length;
    const range = sheet.getRange(2, 1, rowData.length, columnCount);
    range.setValues(rowData);
}

function formatSheet(sheet, columnCount) {
    // Set frozen rows (now 2 for both headers)
    sheet.setFrozenRows(2);

    // Create filter starting from second header row
    sheet.getRange(2, 1, sheet.getLastRow() - 1, columnCount).createFilter();

    // Auto-resize all columns
    for (let i = 1; i <= columnCount; i++) {
        sheet.autoResizeColumn(i);
    }
}

// Analytics data operations
function getAnalyticsAccounts() {
    try {
        const response = AnalyticsAdmin.Accounts.list();
        return response.accounts || [];
    } catch (error) {
        handleError(error, "getAnalyticsAccounts");
    }
}

function getPropertiesForAccount(accountId) {
    try {
        const response = AnalyticsAdmin.Properties.list({
            filter: `parent:accounts/${accountId}`,
        });
        return response.properties || [];
    } catch (error) {
        Logger.log(`Error getting properties for account ${accountId}: ${error}`);
        return [];
    }
}

function getDataStreamsForProperty(propertyName) {
    try {
        const response = AnalyticsAdmin.Properties.DataStreams.list({
            parent: propertyName,
        });
        return response.dataStreams || [];
    } catch (error) {
        Logger.log(`Error getting data streams for property ${propertyName}: ${error}`);
        return [];
    }
}

// Data transformation
function formatDataStreams(dataStreams) {
    return dataStreams
        .map((stream) => `${stream.displayName} (${stream.type})`)
        .join(", ");
}

function createPropertyRow(account, property, dataStreams) {
    return [
        account.displayName,
        extractId(account.name, "account"),
        property.displayName,
        extractId(property.name, "property"),
        new Date(property.createTime).toLocaleDateString(),
        property.timeZone,
        property.displayName,
        formatDataStreams(dataStreams),
    ];
}

function listAccounts() {
    try {
        const accounts = getAnalyticsAccounts();
        const sheet = SpreadsheetApp.getActiveSheet();
        
        initializeSheet(sheet);
        
        if (accounts && accounts.length > 0) {
            const accountRows = accounts.map((account, i) => [
                account.displayName,
                extractId(account.name, "account")
            ]);
            
            writeDataToSheet(sheet, accountRows);
            formatSheet(sheet, 2);
            
            sheet
                .insertRowBefore(2)
                .getRange(2, 1)
                .setValue(`Total # of Accounts = ${accounts.length}`)
                .setFontWeight("bold")
                .setHorizontalAlignment("center");
        } else {
            Logger.log("No accounts found.");
        }
    } catch (error) {
        handleError(error, "listAccounts");
    }
}

function listGA4Properties() {
    const sheet = SpreadsheetApp.getActive().getActiveSheet();
    initializeSheet(sheet);

    try {
        const accounts = getAnalyticsAccounts();
        if (!accounts || accounts.length === 0) {
            Logger.log("No accounts found.");
            return;
        }

        const rowData = accounts.reduce((acc, account) => {
            const accountId = extractId(account.name, "account");
            const properties = getPropertiesForAccount(accountId);

            if (properties && properties.length > 0) {
                const propertyRows = properties.map(property => {
                    const dataStreams = getDataStreamsForProperty(property.name);
                    return createPropertyRow(account, property, dataStreams);
                });
                return [...acc, ...propertyRows];
            }
            return acc;
        }, []);

        if (rowData.length > 0) {
            writeDataToSheet(sheet, rowData);
            formatSheet(sheet);
        }
    } catch (error) {
        handleError(error, "listGA4Properties");
    }
}

function showSidebar() {
    const html = HtmlService.createHtmlOutput(`
        <style>
            :root {
                --primary-color: #1a73e8;
                --primary-hover: #1557b0;
                --background-color: #ffffff;
                --text-color: #202124;
                --border-color: #dadce0;
                --secondary-text: #5f6368;
                --spacing-unit: 8px;
            }

            body { 
                font-family: 'Google Sans', Arial, sans-serif;
                padding: calc(var(--spacing-unit) * 2);
                margin: 0;
                background: var(--background-color);
                color: var(--text-color);
                line-height: 1.5;
            }

            .section { 
                margin-bottom: calc(var(--spacing-unit) * 3);
            }

            .section h3 {
                font-size: 18px;
                font-weight: 500;
                margin-bottom: calc(var(--spacing-unit));
                color: var(--text-color);
            }

            .property-group { 
                margin: calc(var(--spacing-unit) * 1.5) 0 0 calc(var(--spacing-unit) * 2);
            }

            .account-container { 
                background: #f8f9fa;
                border-radius: 8px;
                padding: calc(var(--spacing-unit) * 2);
                margin-bottom: calc(var(--spacing-unit) * 2);
                box-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);
            }

            .account-container strong {
                display: block;
                font-size: 16px;
                color: var(--primary-color);
                margin-bottom: var(--spacing-unit);
            }

            .property-item {
                display: flex;
                align-items: center;
                padding: calc(var(--spacing-unit) * 0.75) 0;
            }

            .property-item input[type="radio"] {
                margin-right: var(--spacing-unit);
                accent-color: var(--primary-color);
                cursor: pointer;
            }

            .property-item label {
                cursor: pointer;
                font-size: 14px;
                color: var(--secondary-text);
            }

            .property-item:hover label {
                color: var(--text-color);
            }

            .next-button {
                width: 100%;
                padding: calc(var(--spacing-unit) * 1.5);
                background-color: var(--primary-color);
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                margin-top: calc(var(--spacing-unit));
                font-family: 'Google Sans', Arial, sans-serif;
                font-size: 14px;
                font-weight: 500;
                text-transform: uppercase;
                letter-spacing: 0.25px;
                transition: background-color 0.2s ease;
            }

            .next-button:hover { 
                background-color: var(--primary-hover);
            }

            .next-button:focus {
                outline: none;
                box-shadow: 0 0 0 2px var(--primary-color), 0 0 0 4px rgba(26, 115, 232, 0.2);
            }

            .error-message {
                color: #d93025;
                font-size: 14px;
                padding: var(--spacing-unit);
                border-radius: 4px;
                background-color: #fce8e6;
            }
        </style>
        <div id="content">
            <div class="section">
                <h3>Choose Property</h3>
                <div id="properties-list">
                </div>
            </div>
            <button class="next-button" onclick="handleNext()">Continue</button>
        </div>
        <script>
            function loadProperties() {
                google.script.run
                    .withSuccessHandler(displayProperties)
                    .withFailureHandler(handleError)
                    .getPropertiesForSidebar();
            }

            function handleError(error) {
                const container = document.getElementById('properties-list');
                container.innerHTML = '<div class="error-message">Error loading properties. Please try again.</div>';
            }

            function displayProperties(data) {
                const container = document.getElementById('properties-list');
                if (!Array.isArray(data)) {
                    handleError();
                    return;
                }
                if (data.length === 0) {
                    container.innerHTML = '<div class="error-message">No properties available</div>';
                    return;
                }
                
                data.forEach(accountData => {
                    const accountDiv = document.createElement('div');
                    accountDiv.className = 'account-container';
                    accountDiv.innerHTML = \`<strong>\${accountData.name}</strong>\`;
                    
                    const propertiesDiv = document.createElement('div');
                    propertiesDiv.className = 'property-group';
                    
                    if (Array.isArray(accountData.properties)) {
                        accountData.properties.forEach(property => {
                            const propertyDiv = document.createElement('div');
                            propertyDiv.className = 'property-item';
                            propertyDiv.innerHTML = \`
                                <input type="radio" name="property" 
                                       value="\${property.id}::\${property.name}" 
                                       id="\${property.id}">
                                <label for="\${property.id}">\${property.name}</label>
                            \`;
                            propertiesDiv.appendChild(propertyDiv);
                        });
                    }
                    
                    accountDiv.appendChild(propertiesDiv);
                    container.appendChild(accountDiv);
                });
            }

            function handleNext() {
                const selected = document.querySelector('input[name="property"]:checked');
                if (selected) {
                    const [id, name] = selected.value.split('::');
                    google.script.run.handlePropertySelection(id, name);
                } else {
                    alert('Please select a property to continue');
                }
            }

            window.onload = loadProperties;
        </script>
    `)
    .setTitle('Post Launch Validator')
    .setWidth(750);

    SpreadsheetApp.getUi().showSidebar(html);
}

function getGA4Dimensions(propertyId) {
    try {
        // Make the API call to get metadata
        const response = AnalyticsData.Properties.getMetadata(`properties/${propertyId}/metadata`);
        Logger.log(response.dimensions)
        // Extract and format dimension metadata
        return response.dimensions.map(dimension => ({
            apiName: dimension.apiName,
            uiName: dimension.uiName || dimension.apiName,
            description: dimension.description || '',
            customDefinition: dimension.customDefinition || false,
            deprecatedApiNames: dimension.deprecatedApiNames || [],
            category: dimension.category || '',
            type: dimension.type || ''
        }));
    } catch (error) {
        handleError(error, 'getGA4Dimensions');
        return [];
    }
}

function getGA4Metrics(propertyId) {
    try {
        // Make the API call to get metadata
        const response = AnalyticsData.Properties.getMetadata(`properties/${propertyId}/metadata`);
        
        // Extract and format metric metadata
        return response.metrics.map(metric => ({
            apiName: metric.apiName,
            uiName: metric.uiName || metric.apiName,
            description: metric.description || '',
            customDefinition: metric.customDefinition || false,
            deprecatedApiNames: metric.deprecatedApiNames || [],
            category: metric.category || '',
            expression: metric.expression || '',
            type: metric.type || ''
        }));
    } catch (error) {
        handleError(error, 'getGA4Metrics');
        return [];
    }
}

function getPropertiesForSidebar() {
    const accounts = getAnalyticsAccounts();
    return accounts.map(account => ({
        name: account.displayName,
        properties: getPropertiesForAccount(extractId(account.name, "account"))
            .map(property => ({
                id: extractId(property.name, "property"),
                name: property.displayName
            }))
    }));
}

function handlePropertySelection(propertyId, propertyName) {
    // Store the selected property ID
    PropertiesService.getDocumentProperties().setProperty('selectedPropertyId', propertyId);
    PropertiesService.getDocumentProperties().setProperty('selectedPropertyName', propertyName);
    
    // Get dimensions and metrics for the selected property
    const dimensions = getGA4Dimensions(propertyId);
    const metrics = getGA4Metrics(propertyId);

    const html = HtmlService.createHtmlOutput(`
        <style>
            :root {
                --primary-color: #1a73e8;
                --primary-hover: #1557b0;
                --background-color: #ffffff;
                --text-color: #202124;
                --border-color: #dadce0;
                --secondary-text: #5f6368;
                --spacing-unit: 8px;
            }

            body { 
                font-family: 'Google Sans', Arial, sans-serif;
                padding: calc(var(--spacing-unit) * 2);
                background: var(--background-color);
                color: var(--text-color);
                line-height: 1.5;
            }

            h4 {
                font-size: 16px;
                font-weight: 700;
                margin-bottom: var(--spacing-unit);
                margin-top: 0;
                color: var(--text-color);
            }

            .section h3 {
                font-size: 18px;
                font-weight: 500;
                color: var(--text-color);
            }

            h3 {
                margin-bottom: 0;
            }

            .report-type-selector {
                display: block;
                gap: calc(var(--spacing-unit) * 2);
                margin-bottom: calc(var(--spacing-unit) * 2);
            }

            .report-type-option {
                display: block;
                align-items: center;
                gap: calc(var(--spacing-unit));
            }

            .report-type-option input[type="radio"] {
                accent-color: var(--primary-color);
                cursor: pointer;
            }

            .report-type-option label {
                cursor: pointer;
                font-size: 14px;
                color: var(--text-color);
            }

            .recurring-options {
                display: none; /* Initially hidden */
                flex-direction: column;
                gap: calc(var(--spacing-unit));
            }

            .recurring-option {
                display: flex;
                align-items: center;
            }

            .recurring-option input[type="radio"] {
                accent-color: var(--primary-color);
                cursor: pointer;
            }

            .recurring-option label {
                cursor: pointer;
                font-size: 14px;
                color: var(--text-color);
            }

            .date-range-container {
                display: flex;
                margin: auto;
                flex-direction: column;
                width: fit-content;
                gap: calc(var(--spacing-unit));
                margin-bottom: calc(var(--spacing-unit) * 2);
                text-align: center;
            }

            .date-input {
                width: max-content;
                padding: calc(var(--spacing-unit) * 1.5);
                background-color: #f8f9fa;
                border: 1px solid var(--border-color);
                border-radius: 4px;
                font-family: 'Google Sans', Arial, sans-serif;
                font-size: 14px;
                color: var(--text-color);
                cursor: pointer;
            }

            .compare-text {
                color: var(--secondary-text);
                font-size: 14px;
                display: inline-block;
            }

            .selected-items {
                display: flex;
                flex-wrap: wrap;
                gap: 8px;
            }

            .chip {
                display: inline-flex;
                align-items: center;
                background: var(--primary-color);
                color: white;
                padding: 4px 12px;
                border-radius: 16px;
                font-size: 12px;
                gap: 4px;
            }

            .chip .close {
                cursor: pointer;
                font-weight: bold;
                padding: 0 4px;
            }

            .dropdown-container {
                position: relative;
            }

            .dropdown-search {
                width: max-content;
                padding: 8px;
                border: 1px solid var(--border-color);
                border-radius: 4px;
                margin-bottom: 4px;
            }

            .dropdown-list {
                max-height: 200px;
                overflow-y: auto;
                border: 1px solid var(--border-color);
                border-radius: 4px;
                background: white;
            }

            .dropdown-item {
                padding: 8px;
                cursor: pointer;
                display: flex;
                align-items: center;
            }

            .dropdown-item:hover {
                background-color: #f1f3f4;
            }

            .dropdown-item input[type="checkbox"] {
                margin-right: 8px;
            }

            .date-button:hover {
                background-color: #f1f3f4;
            }

            .variance-container {
                margin: 0 auto;
                text-align: center;
                margin-bottom: calc(var(--spacing-unit) * 3);
            }

            .variance-input {
                width: 60px;
                padding: calc(var(--spacing-unit));
                border: 1px solid var(--border-color);
                border-radius: 4px;
                font-size: 14px;
                margin-right: calc(var(--spacing-unit));
            }

            .variance-label {
                font-size: 14px;
                color: var(--secondary-text);
            }

            .selection-box {
                border: 1px solid var(--border-color);
                border-radius: 4px;
                padding: calc(var(--spacing-unit) * 2);
                background-color: #f8f9fa;
                margin-bottom: calc(var(--spacing-unit) * 2);
            }

            .button-container {
                display: flex;
                gap: calc(var(--spacing-unit));
                margin-top: calc(var(--spacing-unit) * 2);
            }

            .back-button {
                width: 100%;
                padding: calc(var(--spacing-unit) * 1.5);
                background-color: #f1f3f4;
                color: var(--text-color);
                border: 1px solid var(--border-color);
                border-radius: 4px;
                cursor: pointer;
                font-family: 'Google Sans', Arial, sans-serif;
                font-size: 14px;
                font-weight: 500;
                text-transform: uppercase;
                letter-spacing: 0.25px;
                transition: background-color 0.2s ease;
            }

            .back-button:hover {
                background-color: #e8eaed;
            }

            .validate-button {
                width: 100%;
                padding: calc(var(--spacing-unit) * 1.5);
                background-color: var(--primary-color);
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-family: 'Google Sans', Arial, sans-serif;
                font-size: 14px;
                font-weight: 500;
                text-transform: uppercase;
                letter-spacing: 0.25px;
                transition: background-color 0.2s ease;
            }

            .validate-button:hover {
                background-color: var(--primary-hover);
            }

            .date-range-group {
                display: flex;
                align-items: center;
                gap: calc(var(--spacing-unit));
            }
            
            .date-range-inputs {
                display: flex;
                align-items: center;
                gap: calc(var(--spacing-unit));
            }
            
            .date-separator {
                color: var(--secondary-text);
                font-size: 14px;
            }

            .chips-container {
                margin: 8px 0;
                min-height: 32px;
                padding: 4px;
                border: 1px solid var(--border-color);
                border-radius: 4px;
                display: flex;
                flex-wrap: wrap;
                gap: 8px;
            }
            
            /* New styles for variance section */
            #varianceSection {
                margin-bottom: calc(var(--spacing-unit) * 2);
                margin-top: calc(var(--spacing-unit) * 2);
                padding: calc(var(--spacing-unit) * 2);
                background-color: #f8f9fa;
                border-radius: 4px;
                border: 1px solid var(--border-color);
            }
            
            .radio-group {
                display: flex;
                flex-direction: column;
                gap: calc(var(--spacing-unit));
                margin-top: calc(var(--spacing-unit));
            }
            
            .radio-option {
                display: flex;
                align-items: center;
                gap: calc(var(--spacing-unit));
            }
            
            .radio-option input[type="radio"] {
                accent-color: var(--primary-color);
                cursor: pointer;
            }
            
            .radio-option label {
                cursor: pointer;
                font-size: 14px;
                color: var(--text-color);
            }
            
            input[type="range"] {
                width: 200px;
                accent-color: var(--primary-color);
            }
        </style>

        <div id="content">
            <div class="section">
                <h3>Configure Detection</h3><span style="color: gray; font-size: 12px; margin-bottom: calc(var(--spacing-unit) * 3)">(PID: ${propertyId})</span>

                <div class="report-type-selector">
                    <h4>Anomaly Detection Type</h4>
                    <div class="report-type-option">
                        <input type="radio" id="reportTypeOneTime" name="reportType" value="oneTime" checked onchange="toggleReportOptions()">
                        <label for="reportTypeOneTime">One-Time</label>
                    </div>
                    <div class="report-type-option">
                        <input type="radio" id="reportTypeRecurring" name="reportType" value="recurring" onchange="toggleReportOptions()">
                        <label for="reportTypeRecurring">Recurring</label>
                    </div>
                </div>

                <div class="date-range-container" id="dateRangeContainer">
                    <div class="date-range-group">
                        <div class="date-range-inputs">
                            <input type="date" class="date-input" id="date1_start">
                            <span class="date-separator">to</span>
                            <input type="date" class="date-input" id="date1_end">
                        </div>
                    </div>
                    <span class="compare-text">Compare with</span>
                    <div class="date-range-group">
                        <div class="date-range-inputs">
                            <input type="date" class="date-input" id="date2_start">
                            <span class="date-separator">to</span>
                            <input type="date" class="date-input" id="date2_end">
                        </div>
                    </div>
                </div>

                <div class="recurring-options" id="recurringOptions">
                    <h4>Check Frequency</h4>
                    <div class="recurring-option">
                        <input type="radio" id="recurringHourly" name="recurringOption" value="hourly">
                        <label for="recurringHourly">Hourly</label>
                    </div>
                    <div class="recurring-option">
                        <input type="radio" id="recurringDaily" name="recurringOption" value="daily">
                        <label for="recurringDaily">Daily</label>
                    </div>
                </div>

                <!-- New variance section with z-score options for recurring reports -->
                <div id="varianceSection">
                    <h4>Variance Threshold</h4>
                    <div id="oneTimeVariance">
                        <p>Set the percentage threshold for flagging metric variations:</p>
                        <input type="range" id="varianceThreshold" min="1" max="50" value="10" 
                               oninput="document.getElementById('varianceValue').innerText = this.value + '%'">
                        <span id="varianceValue">10%</span>
                    </div>
                    
                    <div id="recurringVariance" style="display:none;">
                        <div class="radio-group">
                            <div class="radio-option">
                                <input type="radio" id="zScore2" name="zScoreThreshold" value="2">
                                <label for="zScore2">Strict (z-score 2) - Detects smaller anomalies</label>
                            </div>
                            <div class="radio-option">
                                <input type="radio" id="zScore3" name="zScoreThreshold" value="3" checked>
                                <label for="zScore3">Relaxed (z-score 3) - Detects larger anomalies</label>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="selection-box">
                    <h4>Dimensions</h4>
                    <div class="chips-container" id="dimensions-chips"></div>
                    <div class="dropdown-container">
                        <input type="text" class="dropdown-search" id="dimension-search"
                               onkeyup="filterItems('dimension')" placeholder="Search dimensions...">
                        <div class="dropdown-list" id="dimensions-list"></div>
                    </div>
                </div>
                <div class="selection-box">
                    <h4>Metrics</h4>
                    <div class="chips-container" id="metrics-chips"></div>
                    <div class="dropdown-container">
                        <input type="text" class="dropdown-search" id="metric-search"
                               onkeyup="filterItems('metric')" placeholder="Search metrics...">
                        <div class="dropdown-list" id="metrics-list"></div>
                    </div>
                </div>
            </div>
            <div class="button-container">
                <button class="back-button" onclick="goBack()">Back</button>
                <button class="validate-button" onclick="validateSelections()">Validate</button>
            </div>
        </div>

        <script>
            // Store the full list of dimensions and metrics
            const dimensionsData = ${JSON.stringify(dimensions)};
            const metricsData = ${JSON.stringify(metrics)};
            
            // Store the state of selections
            const state = {
                dimensions: new Set(),
                metrics: new Set()
            };

            function createItemElement(item, type) {
                const div = document.createElement('div');
                div.className = 'dropdown-item';
                
                const input = document.createElement('input');
                input.type = 'checkbox';
                input.id = \`\${type}-\${item.apiName}\`;
                input.checked = state[type + 's'].has(item.apiName);
                
                const label = document.createElement('label');
                label.htmlFor = input.id;
                label.title = item.description;
                label.textContent = item.uiName;
                
                // Add click handlers to both div and input
                div.onclick = (e) => {
                    if (e.target !== input) {
                        input.checked = !input.checked;
                        toggleSelection(item.uiName, type, input.checked);
                    }
                };
                
                input.onclick = (e) => {
                    e.stopPropagation();
                    toggleSelection(item.uiName, type, e.target.checked);
                };
                
                div.appendChild(input);
                div.appendChild(label);
                return div;
            }

            function filterItems(type) {
                const searchInput = document.getElementById(type + '-search');
                const filter = searchInput.value.toLowerCase();
                const list = document.getElementById(type + 's-list');
                
                // Clear current list
                list.innerHTML = '';
                
                // Get the correct data array
                const data = type === 'dimension' ? dimensionsData : metricsData;
                
                // Filter and create elements
                data.filter(item => 
                    item.uiName.toLowerCase().includes(filter)
                ).forEach(item => {
                    list.appendChild(createItemElement(item, type));
                });
            }

            function toggleSelection(uiName, type, isChecked) {
                // Find the corresponding data item to get its apiName
                const data = type === 'dimension' ? dimensionsData : metricsData;
                const item = data.find(i => i.uiName === uiName);
                
                if (isChecked) {
                    state[type + 's'].add(item.apiName); // Store apiName in state
                } else {
                    state[type + 's'].delete(item.apiName);
                }
                updateChips(type);
            }

            function updateChips(type) {
                const container = document.getElementById(type + 's-chips');
                container.innerHTML = '';
                
                const data = type === 'dimension' ? dimensionsData : metricsData;
                
                state[type + 's'].forEach(apiName => {
                    const item = data.find(i => i.apiName === apiName);
                    const chip = document.createElement('div');
                    chip.className = 'chip';
                    chip.innerHTML = item.uiName + 
                        \`<span class="close" onclick="removeSelection('\${apiName}', '\${type}')">&times;</span>\`;
                    container.appendChild(chip);
                });
            }

            function removeSelection(apiName, type) {
                state[type + 's'].delete(apiName);
                updateChips(type);
                filterItems(type); // Refresh the list to update checkboxes
            }

            function toggleReportOptions() {
                const reportTypeOneTime = document.getElementById('reportTypeOneTime');
                const dateRangeContainer = document.getElementById('dateRangeContainer');
                const recurringOptions = document.getElementById('recurringOptions');
                const oneTimeVariance = document.getElementById('oneTimeVariance');
                const recurringVariance = document.getElementById('recurringVariance');

                if (reportTypeOneTime.checked) {
                    dateRangeContainer.style.display = 'block';
                    recurringOptions.style.display = 'none';
                    oneTimeVariance.style.display = 'block';
                    recurringVariance.style.display = 'none';
                } else {
                    dateRangeContainer.style.display = 'none';
                    recurringOptions.style.display = 'flex';
                    oneTimeVariance.style.display = 'none';
                    recurringVariance.style.display = 'block';
                }
            }

            function validateSelections() {
                // Check for required selections
                if (state.dimensions.size === 0 || state.metrics.size === 0) {
                    alert('Please select at least one dimension and one metric');
                    return;
                }

                const selections = {
                    reportType: document.querySelector('input[name="reportType"]:checked').value,
                    dates: {
                        range1: {
                            startDate: document.getElementById('date1_start').value,
                            endDate: document.getElementById('date1_end').value
                        },
                        range2: {
                            startDate: document.getElementById('date2_start').value,
                            endDate: document.getElementById('date2_end').value
                        },
                        recurring: document.getElementById('reportTypeRecurring').checked,
                        recurringOption: document.querySelector('input[name="recurringOption"]:checked')?.value
                    },
                    variance: document.querySelector('#varianceThreshold').value,
                    dimensions: Array.from(state.dimensions),
                    metrics: Array.from(state.metrics)
                };
                
                // Add z-score threshold for recurring reports
                if (selections.reportType === 'recurring') {
                    const zScoreRadio = document.querySelector('input[name="zScoreThreshold"]:checked');
                    selections.zScoreThreshold = zScoreRadio ? zScoreRadio.value : '3';
                }

                console.log("selections from html:", selections);
                
                google.script.run
                    .withSuccessHandler(handleValidationSuccess)
                    .withFailureHandler(handleValidationError)
                    .writeReportsToSheetAndSendAnomalyEmail(selections);
            }

            function handleValidationSuccess(response) {
                if (!response.success) {
                    // Show error alert if validation failed
                    alert(response.message || 'An error occurred while generating the report');
                    return;
                }
                // Handle successful validation
                console.log('Validation successful:', response);
            }

            function handleValidationError(error) {
                // Show error alert
                alert('Failed to generate report: ' + (error.message || error.toString()));
            }

            // Initialize the lists when the page loads
            window.onload = function() {
                // Initialize date inputs
                const today = new Date();
                const lastMonth = new Date(today);
                lastMonth.setMonth(today.getMonth() - 1);
                
                // Format dates as YYYY-MM-DD
                const formatDate = (date) => date.toISOString().split('T')[0];
                
                // Set current month range
                document.getElementById('date1_start').value = formatDate(new Date(today.getFullYear(), today.getMonth(), 1));
                document.getElementById('date1_end').value = formatDate(today);
                
                // Set last month range
                document.getElementById('date2_start').value = formatDate(new Date(lastMonth.getFullYear(), lastMonth.getMonth(), 1));
                document.getElementById('date2_end').value = formatDate(new Date(lastMonth.getFullYear(), lastMonth.getMonth() + 1, 0));

                // Initialize dimensions and metrics lists
                filterItems('dimension');
                filterItems('metric');
                
                // Initialize chips containers
                updateChips('dimension');
                updateChips('metric');

                toggleReportOptions(); // Initialize UI based on default report type (One-Time)
            };

            // Add back button functionality
            function goBack() {
                google.script.run.showSidebar();
            }
        </script>
    `)
    .setTitle('Post Launch Validator')
    .setWidth(750);

    SpreadsheetApp.getUi().showSidebar(html);
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu(CONFIG.menuTitle)
        .addSubMenu(
            ui
                .createMenu("Google Analytics")
                .addItem("List Accounts", "listAccounts")
                .addItem("List Properties", "listGA4Properties")
        )
        .addItem("Launch", "showSidebar")
        .addToUi();
}

// Update the validateAndRunReport function to handle the new date range format
function validateSelectionsAndFetchData(selections) {
    try {
        const propertyId = PropertiesService.getDocumentProperties().getProperty('selectedPropertyId');
        if (!propertyId) {
            throw new Error('No property selected');
        }
        // Store the variance threshold for use in email reporting
        PropertiesService.getDocumentProperties().setProperty('varianceThreshold', selections.variance);
        
        // Store the z-score threshold if it exists
        if (selections.zScoreThreshold) {
            PropertiesService.getDocumentProperties().setProperty('zScoreThreshold', selections.zScoreThreshold);
        }

        Logger.log("Selections from validateSelectionsAndFetchData:", selections); // Log selections

        if (selections.reportType === 'recurring') {
            Logger.log("Recurring report selected");
            if (selections.dates.recurringOption === 'hourly') {
                Logger.log("Hourly recurring option selected");

                const dimensionObjects = selections.dimensions.map(d => ({ name: d }));
                const metricObjects = selections.metrics.map(m => ({ name: m }));
                const varianceThreshold = Number(selections.variance);

                // --- Hourly Recurring Logic ---
                const now = new Date();
                // Adjust for 24-hour delay
                const yesterday = new Date(now);
                yesterday.setDate(now.getDate() - 1);
                
                const currentHour = now.getHours();
                const currentDayOfWeek = yesterday.getDay(); // Use yesterday's day of week
                const timeZone = Session.getTimeZone(); // Script's timezone
                
                Logger.log(`Using 24-hour delayed data: Current hour ${currentHour} from yesterday`);

                // Function to get date range for specific hour of day, past weeks
                function getHistoricalDateRangesForHour(hour, dayOfWeek, numberOfWeeks, timeZone) {
                    const dateRanges = [];
                    for (let weekOffset = 1; weekOffset <= numberOfWeeks; weekOffset++) {
                        // Create a date for yesterday
                        const startDate = new Date(yesterday);
                        // Go back 'weekOffset' weeks from yesterday
                        startDate.setDate(startDate.getDate() - (weekOffset * 7));
                        
                        // Format dates as YYYY-MM-DD for GA4 API (without time component)
                        const formattedDate = Utilities.formatDate(startDate, timeZone, "yyyy-MM-dd");
                        
                        dateRanges.push({
                            startDate: formattedDate,
                            endDate: formattedDate // Same day for daily data
                        });
                    }
                    return dateRanges;
                }

                // Get historical date ranges for the current hour, last 4 weeks
                const historicalDateRanges = getHistoricalDateRangesForHour(currentHour, currentDayOfWeek, 4, timeZone);
                Logger.log("Historical Date Ranges for Hourly Report:", historicalDateRanges);

                // Get current date range (for yesterday, same hour)
                const currentDateRange = {
                    startDate: Utilities.formatDate(yesterday, timeZone, "yyyy-MM-dd"),
                    endDate: Utilities.formatDate(yesterday, timeZone, "yyyy-MM-dd")
                };
                Logger.log("Current Date Range for Hourly Report (24h delayed):", currentDateRange);

                // Fetch historical data (last 4 weeks, same hour)
                const historicalReports = [];
                for (const range of historicalDateRanges) {
                    Logger.log(`Fetching data for range: ${range.startDate} to ${range.endDate}, hour: ${currentHour}`);
                    const report = getGA4Report(propertyId, dimensionObjects, metricObjects, range, {
                        hourly: true,
                        hour: currentHour
                    });
                    if (report && report.rows && report.rows.length > 0) {
                        historicalReports.push(report);
                        Logger.log(`Found data for range: ${range.startDate}, rows: ${report.rows.length}`);
                    } else {
                        Logger.log(`No data for historical range: ${range.startDate} to ${range.endDate}, hour: ${currentHour}`);
                        historicalReports.push(null); // Push null to keep array length consistent
                    }
                }

                // Fetch current hour data
                Logger.log(`Fetching data for current date: ${currentDateRange.startDate}, hour: ${currentHour}`);
                let currentReport = getGA4Report(propertyId, dimensionObjects, metricObjects, currentDateRange, {
                    hourly: true,
                    hour: currentHour
                });

                if (!currentReport || !currentReport.rows || currentReport.rows.length === 0) {
                    Logger.log("No data found for current hour. Creating placeholder report.");
                    
                    // Create placeholder data instead of throwing an error
                    const placeholderRow = [
                        ...Array(dimensionObjects.length).fill("No Data"),
                        ...Array(metricObjects.length).fill("0")
                    ];
                    
                    currentReport = {
                        headers: [...dimensionObjects.map(d => d.name), ...metricObjects.map(m => m.name)],
                        dimensionHeaders: dimensionObjects.map(d => d.name),
                        rows: [placeholderRow],
                        rowCount: 1
                    };
                }

                // Process historical data to calculate mean and variance
                // Group historical data by dimension values
                const historicalDataByDimension = {};
                const dimensionCount = dimensionObjects.length;
                const metricCount = metricObjects.length;

                // save dimension count and metric count to properties for sendAnomalyEmail() to use
                PropertiesService.getDocumentProperties().setProperty('metricCount', metricCount);
                PropertiesService.getDocumentProperties().setProperty('dimensionCount', dimensionCount);

                // Initialize the structure with the dimension values from current report
                currentReport.rows.forEach(row => {
                    const dimensionKey = row.slice(0, dimensionCount).join('||');
                    historicalDataByDimension[dimensionKey] = Array(metricObjects.length).fill([]);
                });

                // Collect historical data for each dimension combination
                for (const report of historicalReports) {
                    if (report && report.rows && report.rows.length > 0) {
                        report.rows.forEach(row => {
                            const dimensionValues = row.slice(0, dimensionCount);
                            const dimensionKey = dimensionValues.join('||');
                            
                            // If this dimension combination is in our current report
                            if (historicalDataByDimension[dimensionKey]) {
                                // Add each metric value to the corresponding array
                                for (let metricIndex = 0; metricIndex < metricObjects.length; metricIndex++) {
                                    const metricValue = parseFloat(row[dimensionCount + metricIndex]) || 0;
                                    
                                    // Create a new array if needed (deep copy to avoid reference issues)
                                    if (!Array.isArray(historicalDataByDimension[dimensionKey][metricIndex])) {
                                        historicalDataByDimension[dimensionKey][metricIndex] = [];
                                    }
                                    
                                    historicalDataByDimension[dimensionKey][metricIndex].push(metricValue);
                                    Logger.log(`Historical value for dimension "${dimensionKey}", metric ${metricIndex}: ${metricValue}`);
                                }
                            }
                        });
                    }
                }

                // Calculate mean for each dimension and metric
                const calculatedVarianceData = {};

                // Process current report rows and calculate variance
                const enhancedRows = currentReport.rows.map(row => {
                    const dimensionValues = row.slice(0, dimensionCount);
                    const dimensionKey = dimensionValues.join('||');
                    const metrics = row.slice(dimensionCount);
                    
                    // Calculate historical means, standard deviations, and Z-scores for this dimension
                    const statistics = metrics.map((currentValue, metricIndex) => {
                        const currentValueNum = parseFloat(currentValue) || 0;
                        const historicalValues = historicalDataByDimension[dimensionKey][metricIndex] || [];
                        
                        // Calculate mean
                        let historicalMean = 0;
                        if (historicalValues.length > 0) {
                            const sum = historicalValues.reduce((acc, val) => acc + val, 0);
                            historicalMean = sum / historicalValues.length;
                        }
                        
                        // Calculate variance and standard deviation
                        let variance = 0;
                        let stdDev = 0;
                        if (historicalValues.length > 1) {
                            // Sum of squared differences
                            const sumOfSquaredDiffs = historicalValues.reduce((acc, val) => {
                                const diff = val - historicalMean;
                                return acc + (diff * diff);
                            }, 0);
                            
                            // Variance = sum of squared differences / (n-1)
                            variance = sumOfSquaredDiffs / (historicalValues.length - 1);
                            
                            // Standard deviation = square root of variance
                            stdDev = Math.sqrt(variance);
                        }
                        
                        // Calculate Z-score
                        let zScore = 0;
                        if (stdDev > 0) {
                            zScore = (currentValueNum - historicalMean) / stdDev;
                            // Check for invalid Z-score values
                            if (!isFinite(zScore) || isNaN(zScore)) {
                                zScore = 0; // Use 0 for any invalid calculation results
                                Logger.log(`  Warning: Invalid Z-score calculated, using 0 instead`);
                            }
                        } else if (currentValueNum !== historicalMean) {
                            // If stdDev is 0 but values differ, use a large but finite value
                            zScore = currentValueNum > historicalMean ? 10 : -10; // More reasonable than Infinity
                            Logger.log(`  Warning: Zero standard deviation with different values, using ${zScore} as Z-score`);
                        }
                        
                        // For UI display, also calculate percentage difference
                        let variancePercent = 0;
                        if (historicalMean !== 0) {
                            variancePercent = ((currentValueNum - historicalMean) / historicalMean) * 100;
                        } else if (currentValueNum !== 0) {
                            variancePercent = 100;
                        }
                        
                        Logger.log(`Statistics for dimension "${dimensionKey}", metric ${metricIndex}:`);
                        Logger.log(`  Current value: ${currentValueNum}`);
                        Logger.log(`  Historical values: ${historicalValues.join(', ')}`);
                        Logger.log(`  Historical mean: ${historicalMean}`);
                        Logger.log(`  Standard deviation: ${stdDev}`);
                        Logger.log(`  Z-score: ${zScore}`);
                        Logger.log(`  Variance percent: ${variancePercent}%`);
                        
                        return {
                            currentValue: currentValueNum,
                            historicalMean: historicalMean,
                            stdDev: stdDev,
                            zScore: zScore,
                            variancePercent: variancePercent
                        };
                    });
                    
                    calculatedVarianceData[dimensionKey] = statistics;
                    return row; // Return original row
                });

                // Create the structured variance data for the report
                const varianceData = [];
                currentReport.rows.forEach((row, rowIndex) => {
                    const dimensionValues = row.slice(0, dimensionCount);
                    const dimensionKey = dimensionValues.join('||');
                    
                    if (calculatedVarianceData[dimensionKey]) {
                        varianceData[rowIndex] = calculatedVarianceData[dimensionKey];
                    }
                });

                return {
                    report1: currentReport,
                    report2: {
                        headers: currentReport.headers,
                        dimensionHeaders: currentReport.dimensionHeaders,
                        rows: enhancedRows,
                        varianceData: varianceData
                    },
                    varianceThreshold: selections.variance,
                    zScoreThreshold: selections.zScoreThreshold,
                    dateRanges: [currentDateRange, historicalDateRanges]
                };

            } else if (selections.dates.recurringOption === 'daily') {
                Logger.log("Daily recurring option selected");
                
                const dimensionObjects = selections.dimensions.map(d => ({ name: d }));
                const metricObjects = selections.metrics.map(m => ({ name: m }));
                const varianceThreshold = Number(selections.variance);

                // --- Daily Recurring Logic ---
                const now = new Date();
                // Adjust for 1-day delay
                const twoDaysAgo = new Date(now);
                twoDaysAgo.setDate(now.getDate() - 2);
                const yesterday = new Date(now);
                yesterday.setDate(now.getDate() - 1);
                
                const timeZone = Session.getTimeZone();
                
                Logger.log(`Using 1-day delayed data: Comparing yesterday with historical data`);
                
                // Format dates for GA4 API
                const yesterdayFormatted = Utilities.formatDate(yesterday, timeZone, "yyyy-MM-dd");
                const twoDaysAgoFormatted = Utilities.formatDate(twoDaysAgo, timeZone, "yyyy-MM-dd");
                
                // Set up date ranges for current data (yesterday)
                const currentDateRange = {
                    startDate: yesterdayFormatted,
                    endDate: yesterdayFormatted
                };
                
                // Get data for yesterday (current data)
                const currentReport = getGA4Report(propertyId, dimensionObjects, metricObjects, currentDateRange);
                
                if (!currentReport || !currentReport.rows || currentReport.rows.length === 0) {
                    throw new Error('No data available for yesterday');
                }
                
                // Now gather historical data for the same day of week over the past 4 weeks
                const historicalRows = [];
                const historicalDateRanges = [];
                
                // Get data for the same day of week over the past 4 weeks
                for (let i = 1; i <= 4; i++) {
                    const historicalDate = new Date(yesterday);
                    historicalDate.setDate(historicalDate.getDate() - (7 * i)); // Same day of week, i weeks ago
                    
                    const historicalDateFormatted = Utilities.formatDate(historicalDate, timeZone, "yyyy-MM-dd");
                    Logger.log(`Historical date ${i}: ${historicalDateFormatted}`);
                    
                    const historicalDateRange = {
                        startDate: historicalDateFormatted,
                        endDate: historicalDateFormatted
                    };
                    
                    historicalDateRanges.push(historicalDateRange);
                    
                    // Get data for this historical date
                    const historicalReport = getGA4Report(propertyId, dimensionObjects, metricObjects, historicalDateRange);
                    
                    if (historicalReport && historicalReport.rows && historicalReport.rows.length > 0) {
                        historicalReport.rows.forEach(row => {
                            historicalRows.push(row);
                        });
                    }
                }
                
                if (historicalRows.length === 0) {
                    Logger.log("No historical data available, using current data only");
                }
                
                // Process the historical data to calculate baselines for each dimension combination
                const dimensionCount = dimensionObjects.length;
                const metricCount = metricObjects.length;

                // save dimension count and metric count to properties for sendAnomalyEmail() to use
                PropertiesService.getDocumentProperties().setProperty('metricCount', metricCount);
                PropertiesService.getDocumentProperties().setProperty('dimensionCount', dimensionCount);
                
                // Group historical data by dimension combinations
                const historicalData = {};
                
                historicalRows.forEach(row => {
                    const dimensionValues = row.slice(0, dimensionCount);
                    const metricValues = row.slice(dimensionCount);
                    
                    const dimensionKey = dimensionValues.join('||');
                    
                    if (!historicalData[dimensionKey]) {
                        historicalData[dimensionKey] = Array(metricCount).fill().map(() => []);
                    }
                    
                    metricValues.forEach((value, i) => {
                        historicalData[dimensionKey][i].push(parseFloat(value) || 0);
                    });
                });
                
                // Calculate mean and standard deviation for each dimension combination and metric
                const calculatedVarianceData = {};
                
                Object.keys(historicalData).forEach(dimensionKey => {
                    calculatedVarianceData[dimensionKey] = [];
                    
                    historicalData[dimensionKey].forEach((values, metricIndex) => {
                        // Calculate mean
                        const sum = values.reduce((acc, val) => acc + val, 0);
                        const mean = values.length > 0 ? sum / values.length : 0;
                        
                        // Calculate standard deviation
                        const squaredDiffs = values.map(val => Math.pow(val - mean, 2));
                        const avgSquaredDiff = squaredDiffs.length > 0 ? 
                            squaredDiffs.reduce((acc, val) => acc + val, 0) / squaredDiffs.length : 0;
                        const stdDev = Math.sqrt(avgSquaredDiff);
                        
                        calculatedVarianceData[dimensionKey][metricIndex] = {
                            historicalMean: mean,
                            stdDev: stdDev,
                            values: values
                        };
                    });
                });
                
                // Enhance current report rows with historical data
                const enhancedRows = currentReport.rows.map(row => {
                    const dimensionValues = row.slice(0, dimensionCount);
                    const metricValues = row.slice(dimensionCount);
                    const dimensionKey = dimensionValues.join('||');
                    
                    const enhanced = [...dimensionValues];
                    
                    metricValues.forEach((value, i) => {
                        const currentValueNum = parseFloat(value) || 0;
                        let historicalMean = 0;
                        let zScore = 0;
                        
                        if (calculatedVarianceData[dimensionKey] && calculatedVarianceData[dimensionKey][i]) {
                            historicalMean = calculatedVarianceData[dimensionKey][i].historicalMean;
                            const stdDev = calculatedVarianceData[dimensionKey][i].stdDev;
                            
                            // Calculate z-score if possible
                            if (stdDev > 0) {
                                zScore = (currentValueNum - historicalMean) / stdDev;
                            } else if (currentValueNum !== historicalMean) {
                                // If stdDev is 0 but values differ, assign a large but finite z-score
                                zScore = currentValueNum > historicalMean ? 10 : -10;
                            }
                            
                            // Handle invalid z-scores
                            if (isNaN(zScore) || !isFinite(zScore)) {
                                zScore = 0;
                            }
                            
                            // Store z-score in variance data
                            calculatedVarianceData[dimensionKey][i].zScore = zScore;
                        }
                        
                        enhanced.push(currentValueNum, historicalMean, zScore);
                    });
                    
                    return enhanced;
                });
                
                // Create the structured variance data for the report
                const varianceData = [];
                currentReport.rows.forEach((row, rowIndex) => {
                    const dimensionValues = row.slice(0, dimensionCount);
                    const dimensionKey = dimensionValues.join('||');
                    
                    if (calculatedVarianceData[dimensionKey]) {
                        varianceData[rowIndex] = calculatedVarianceData[dimensionKey];
                    }
                });
                
                return {
                    report1: currentReport,
                    report2: {
                        headers: currentReport.headers,
                        dimensionHeaders: currentReport.dimensionHeaders,
                        rows: enhancedRows,
                        varianceData: varianceData
                    },
                    varianceThreshold: selections.variance,
                    zScoreThreshold: selections.zScoreThreshold,
                    dateRanges: [currentDateRange, historicalDateRanges]
                };
            }

            // Default return for recurring reports (should not reach here if properly selected)
            return {
                report1: null,
                report2: null,
                varianceThreshold: selections.variance,
                zScoreThreshold: selections.zScoreThreshold,
                dateRanges: null
            };
        } else { // 'oneTime' reportType
            Logger.log("One-time report selected");
            // Format date ranges
            const dateRanges = [
                {
                    startDate: selections.dates.range1.startDate,
                    endDate: selections.dates.range1.endDate
                },
                {
                    startDate: selections.dates.range2.startDate,
                    endDate: selections.dates.range2.endDate
                }
            ];

            // Create dimension and metric objects for the API request
            const dimensionObjects = selections.dimensions.map(d => ({ name: d }));
            const metricObjects = selections.metrics.map(m => ({ name: m }));

            const dimensionCount = dimensionObjects.length;
            const metricCount = metricObjects.length;
            // save dimension count and metric count to properties for sendAnomalyEmail() to use
            PropertiesService.getDocumentProperties().setProperty('metricCount', metricCount);
            PropertiesService.getDocumentProperties().setProperty('dimensionCount', dimensionCount);
            // Get reports for both date ranges
            const report1 = getGA4Report(propertyId, dimensionObjects, metricObjects, dateRanges[0]);
            const report2 = getGA4Report(propertyId, dimensionObjects, metricObjects, dateRanges[1]);

            if (!report1 || !report2) {
                throw new Error('Failed to fetch reports');
            }

            return {
                report1: report1,
                report2: report2,
                varianceThreshold: selections.variance,
                zScoreThreshold: selections.zScoreThreshold,
                dateRanges: dateRanges
            };
        }

    } catch (error) {
        handleError(error, 'validateAndRunReport');
    }
}

function writeReportsToSheetAndSendAnomalyEmail(trigger) {
    try {
        let selections;
        // Check if the argument is trigger metadata (has day-of-month, month, year properties)
        const isTriggerMetadata = trigger && (trigger.hasOwnProperty('day-of-month') || 
                                            trigger.hasOwnProperty('triggerUid') ||
                                            trigger.hasOwnProperty('authMode'));
        
        if (isTriggerMetadata) {
            // If this is called from a trigger, ignore the passed argument
            const selectionsJson = PropertiesService.getDocumentProperties().getProperty('selections');
            if (!selectionsJson) {
                Logger.log("No saved selections found. Cannot generate report.");
                return;
            }
            selections = JSON.parse(selectionsJson);
        } else if (trigger) {
            // This is a manual call with selections, so save them
            Logger.log("Saving new selections from manual call");
            PropertiesService.getDocumentProperties().setProperty('selections', JSON.stringify(trigger));
            selections = trigger;
        }
        
        // Generate report and send email
        writeReportsToSheet(selections);
        sendAnomalyEmail();
    } catch (error) {
        Logger.log("Error in writeReportsToSheetAndSendAnomalyEmail: " + error.toString());
        if (error.stack) {
            Logger.log("Stack trace: " + error.stack);
        }
    }
}

// Update the writeReportsToSheet function to handle existing sheets
function writeReportsToSheet(selections) {
    try {
        const { report1, report2, varianceThreshold, dateRanges } = validateSelectionsAndFetchData(selections);
        
        // Create new sheet with timestamped name
        const timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        const sheetName = `Report ${timeStamp}`;
        
        // Save sheetname to properties
        PropertiesService.getDocumentProperties().setProperty('latestSheetName', sheetName);
        const spreadsheet = SpreadsheetApp.getActive();
        
        // Check if a sheet with this name already exists
        let sheet = spreadsheet.getSheetByName(sheetName);
        
        if (sheet) {
            // If sheet exists, clear it instead of creating a new one
            Logger.log(`Sheet "${sheetName}" already exists. Clearing it.`);
            sheet.clear();
        } else {
            // Create a new sheet if it doesn't exist
            Logger.log(`Creating new sheet "${sheetName}"`);
            sheet = spreadsheet.insertSheet(sheetName);
        }
        
        // Clear any existing filters
        if (sheet.getFilter()) {
            sheet.getFilter().remove();
        }

        // Parse the threshold as a number to ensure it's treated as numeric
        const threshold = Math.abs(Number(varianceThreshold));
        Logger.log(`Using variance threshold for report: ${threshold}%`);
        
        // Check if this is a recurring report
        const isRecurring = selections.reportType === 'recurring';
        
        if (isRecurring) {
            // Handle recurring report (hourly or daily)
            writeRecurringReportToSheet(sheet, report1, report2, threshold, selections);
        } else {
            // Handle one-time report (existing logic)
            writeOneTimeReportToSheet(sheet, report1, report2, threshold, dateRanges);
        }

        // Activate the new sheet for viewing
        spreadsheet.setActiveSheet(sheet);
        
        return {
            success: true,
            message: "Report generated successfully"
        };
    } catch (error) {
        Logger.log("Error in writeReportsToSheet: " + error.toString());
        return {
            success: false,
            message: error.toString(),
        };
    }
}

// Fix the writeOneTimeReportToSheet function's column indexing
function writeOneTimeReportToSheet(sheet, report1, report2, threshold, dateRanges) {
    // This is the existing logic from writeReportsToSheet for one-time reports
    const dimensionHeaders = report1.headers.slice(0, report1.dimensionHeaders.length);
    const metricHeaders = report1.headers.slice(report1.dimensionHeaders.length);
    
    // Format date ranges for display (convert dashes to slashes)
    const formatDateString = (dateStr) => dateStr.replace(/-/g, '/');
    const currentDateRange = `${formatDateString(dateRanges[0].startDate)}-${formatDateString(dateRanges[0].endDate)}`;
    const previousDateRange = `${formatDateString(dateRanges[1].startDate)}-${formatDateString(dateRanges[1].endDate)}`;

    // Write date range headers (Row 1) and style them
    const dateRangeHeaders = [
        ...Array(dimensionHeaders.length).fill('Dimensions'),
        ...metricHeaders.flatMap(() => [
            currentDateRange,
            previousDateRange,
            'Difference (%)'
        ])
    ];
    sheet.getRange(1, 1, 1, dateRangeHeaders.length).setValues([dateRangeHeaders]);

    // Define colors for metrics (Google Material Design colors)
    const metricColors = [
        '#4285F4',  // Google Blue
        '#0F9D58',  // Google Green
        '#DB4437',  // Google Red
        '#F4B400',  // Google Yellow
        '#AA47BC',  // Purple
        '#00ACC1',  // Cyan
        '#FF7043',  // Deep Orange
        '#9E9D24',  // Lime
        '#5C6BC0',  // Indigo
        '#00796B',  // Teal
    ];

    // Style headers
    if (metricHeaders.length > 0) {
        // Merge and style dimension header cells
        if (dimensionHeaders.length > 1) {
            const dimensionRange = sheet.getRange(1, 1, 1, dimensionHeaders.length);
            dimensionRange.merge()
                        .setBackground('#f3f3f3')
                        .setFontWeight('bold')
                        .setHorizontalAlignment('center')
                        .setVerticalAlignment('middle');
        } else {
            // Style single dimension header
            sheet.getRange(1, 1, 1, 1)
                .setBackground('#f3f3f3')
                .setFontWeight('bold')
                .setHorizontalAlignment('center')
                .setVerticalAlignment('middle');
        }
        
        // Style each metric group
        for (let i = 0; i < metricHeaders.length; i++) {
            // Fix: Make this consistent with the data structure (3 columns per metric)
            const startCol = dimensionHeaders.length + (i * 3) + 1;
            const metricColor = metricColors[i % metricColors.length];
            
            // Style current date range and metric
            sheet.getRange(1, startCol, 1, 1)
                .setBackground(metricColor)
                .setFontColor('white')
                .setFontWeight('bold')
                .setHorizontalAlignment('center');
            sheet.getRange(2, startCol, 1, 1)
                .setBackground(metricColor)
                .setFontColor('white')
                .setFontWeight('bold')
                .setHorizontalAlignment('center');
            
            // Style previous date range and metric
            sheet.getRange(1, startCol + 1, 1, 1)
                .setBackground(metricColor)
                .setFontColor('white')
                .setFontWeight('bold')
                .setHorizontalAlignment('center');
            sheet.getRange(2, startCol + 1, 1, 1)
                .setBackground(metricColor)
                .setFontColor('white')
                .setFontWeight('bold')
                .setHorizontalAlignment('center');
            
            // Style difference column headers
            sheet.getRange(1, startCol + 2, 1, 1)
                .setBackground('#FB8C00')  // Orange
                .setFontColor('white')
                .setFontWeight('bold')
                .setHorizontalAlignment('center');
            sheet.getRange(2, startCol + 2, 1, 1)
                .setBackground('#FB8C00')  // Orange
                .setFontColor('white')
                .setFontWeight('bold')
                .setHorizontalAlignment('center');
        }
    }

    // Write metric headers (Row 2)
    const metricGroupHeaders = [
        ...dimensionHeaders,
        ...metricHeaders.flatMap(h => [h, h, h])
    ];
    sheet.getRange(2, 1, 1, metricGroupHeaders.length).setValues([metricGroupHeaders]);

    // Combine and compare data (now starting from Row 3)
    const combinedData = report1.rows.map((row1, i) => {
        const row2 = report2.rows[i] || Array(row1.length).fill(0);
        const dimensions = row1.slice(0, report1.dimensionHeaders.length);
        const metrics1 = row1.slice(report1.dimensionHeaders.length);
        const metrics2 = row2.slice(report1.dimensionHeaders.length);
        
        // Reorganize metrics to group by metric type
        const groupedMetrics = metrics1.map((val1, idx) => {
            const val1Num = parseFloat(val1) || 0;
            const val2Num = parseFloat(metrics2[idx]) || 0;
            const variance = val2Num !== 0 ? ((val1Num - val2Num) / val2Num) * 100 : 0;
            return [val1Num, val2Num, variance];
        }).flat();

        return [...dimensions, ...groupedMetrics];
    });

    // Write data (starting from Row 3)
    if (combinedData.length > 0) {
        sheet.getRange(3, 1, combinedData.length, combinedData[0].length)
            .setValues(combinedData);
    }

    // Add number formatting to variance columns
    const metricCount = metricHeaders.length;
    const dimensionCount = dimensionHeaders.length;
    const lastRow = sheet.getLastRow();


    // When setting number format for variance columns:
    if (lastRow > 2) {
        // Format each metric value column
        for (let i = 0; i < metricCount; i++) {
            const metric1Col = dimensionCount + (i * 3) + 1;
            const metric2Col = dimensionCount + (i * 3) + 2;
            const varianceCol = dimensionCount + (i * 3) + 3;
            
            // Format metric values with thousands separators
            sheet.getRange(3, metric1Col, lastRow - 2, 1).setNumberFormat('#,##0.00');
            sheet.getRange(3, metric2Col, lastRow - 2, 1).setNumberFormat('#,##0.00');
            
            // Format variance columns consistently with two decimal places
            sheet.getRange(3, varianceCol, lastRow - 2, 1).setNumberFormat('0.00');
        }
    }

    // Update conditional formatting for variance columns with clear threshold values
    for (let i = 0; i < metricCount; i++) {
        const varianceCol = dimensionCount + (i * 3) + 3;
        const varianceRange = sheet.getRange(3, varianceCol, sheet.getLastRow() - 2, 1);
        
        Logger.log(`Applying conditional formatting for column ${varianceCol} with threshold ±${threshold}%`);
        
        // Rule for high negative variance (red)
        const highNegativeRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberLessThan(-threshold)
            .setBackground('#F4C7C3') // Red background
            .setRanges([varianceRange])
            .build();
        
        // Rule for acceptable variance (green)
        const acceptableRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberBetween(-threshold, threshold)
            .setBackground('#D9EAD3') // Green background
            .setRanges([varianceRange])
            .build();
        
        // Rule for high positive variance (red)
        const highPositiveRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberGreaterThan(threshold)
            .setBackground('#F4C7C3') // Red background
            .setRanges([varianceRange])
            .build();
        
        // Get existing rules and add new ones
        const rules = sheet.getConditionalFormatRules();
        rules.push(acceptableRule); // Add the acceptable rule first
        rules.push(highNegativeRule);
        rules.push(highPositiveRule);
        sheet.setConditionalFormatRules(rules);
    }

    // Format sheet
    formatSheet(sheet, metricGroupHeaders.length);
}

// New function to write recurring reports to sheet
function writeRecurringReportToSheet(sheet, currentReport, historicalReport, threshold, selections) {
    Logger.log("Starting writeRecurringReportToSheet");
    
    if (!currentReport || !historicalReport) {
        Logger.log("Missing report data for recurring report");
        throw new Error('Missing report data for recurring report');
    }
    
    const dimensionHeaders = currentReport.dimensionHeaders;
    const metricHeaders = currentReport.headers.slice(dimensionHeaders.length);
    const varianceData = historicalReport.varianceData || [];
    
    Logger.log(`Dimension headers: ${JSON.stringify(dimensionHeaders)}`);
    Logger.log(`Metric headers: ${JSON.stringify(metricHeaders)}`);
    Logger.log(`Variance data: ${JSON.stringify(varianceData)}`);
    
    // Format date for display
    const now = new Date();
    const yesterday = new Date(now);
    yesterday.setDate(now.getDate() - 1);

    const formattedDate = Utilities.formatDate(yesterday, Session.getTimeZone(), "yyyy/MM/dd");
    const formattedTime = Utilities.formatDate(now, Session.getTimeZone(), "HH:mm");
    const recurringType = selections.dates.recurringOption.charAt(0).toUpperCase() + selections.dates.recurringOption.slice(1);

    // Write title row with delay information
    sheet.getRange(1, 1, 1, dimensionHeaders.length + metricHeaders.length * 3)
        .merge()
        .setValue(`${recurringType} Recurring Report - ${formattedDate} ${formattedTime} (24h Delayed Data)`)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground('#4285F4')
        .setFontColor('white');
    
    // Write header rows
    const headerRow1 = [
        ...Array(dimensionHeaders.length).fill('Dimensions'),
        ...metricHeaders.flatMap(() => ['Current', 'Historical Avg', 'Z-Score'])
    ];
    
    const headerRow2 = [
        ...dimensionHeaders,
        ...metricHeaders.flatMap(header => [header, header, header])
    ];
    
    sheet.getRange(2, 1, 1, headerRow1.length).setValues([headerRow1]);
    sheet.getRange(3, 1, 1, headerRow2.length).setValues([headerRow2]);
    
    // Style headers
    // Merge and style dimension header cells
    if (dimensionHeaders.length > 1) {
        sheet.getRange(2, 1, 1, dimensionHeaders.length)
            .merge()
            .setBackground('#f3f3f3')
            .setFontWeight('bold')
            .setHorizontalAlignment('center');
    } else if (dimensionHeaders.length === 1) {
        sheet.getRange(2, 1, 1, 1)
            .setBackground('#f3f3f3')
            .setFontWeight('bold')
            .setHorizontalAlignment('center');
    }
    
    // Style metric headers
    const metricColors = [
        '#4285F4',  // Google Blue
        '#0F9D58',  // Google Green
        '#DB4437',  // Google Red
        '#F4B400',  // Google Yellow
        '#AA47BC',  // Purple
        '#00ACC1',  // Cyan
        '#FF7043',  // Deep Orange
        '#9E9D24',  // Lime
        '#5C6BC0',  // Indigo
        '#00796B',  // Teal
    ];
    
    for (let i = 0; i < metricHeaders.length; i++) {
        const startCol = dimensionHeaders.length + (i * 3) + 1;
        const metricColor = metricColors[i % metricColors.length];
        
        // Current value header
        sheet.getRange(2, startCol, 1, 1)
            .setBackground(metricColor)
            .setFontColor('white')
            .setFontWeight('bold')
            .setHorizontalAlignment('center');
        sheet.getRange(3, startCol, 1, 1)
            .setBackground(metricColor)
            .setFontColor('white')
            .setFontWeight('bold')
            .setHorizontalAlignment('center');
        
        // Historical average header
        sheet.getRange(2, startCol + 1, 1, 1)
            .setBackground(metricColor)
            .setFontColor('white')
            .setFontWeight('bold')
            .setHorizontalAlignment('center');
        sheet.getRange(3, startCol + 1, 1, 1)
            .setBackground(metricColor)
            .setFontColor('white')
            .setFontWeight('bold')
            .setHorizontalAlignment('center');
        
        // Z-score header
        sheet.getRange(2, startCol + 2, 1, 1)
            .setBackground('#FB8C00')  // Orange
            .setFontColor('white')
            .setFontWeight('bold')
            .setHorizontalAlignment('center');
        sheet.getRange(3, startCol + 2, 1, 1)
            .setBackground('#FB8C00')  // Orange
            .setFontColor('white')
            .setFontWeight('bold')
            .setHorizontalAlignment('center');
    }
    
    // Prepare data rows
    const dataRows = [];
    
    // Process each row in the current report
    for (let rowIndex = 0; rowIndex < currentReport.rows.length; rowIndex++) {
        const currentRow = currentReport.rows[rowIndex];
        const dimensions = currentRow.slice(0, dimensionHeaders.length);
        const metrics = currentRow.slice(dimensionHeaders.length);
        
        Logger.log(`Processing row ${rowIndex}:`);
        Logger.log(`  Dimensions: ${JSON.stringify(dimensions)}`);
        Logger.log(`  Metrics: ${JSON.stringify(metrics)}`);
        
        const rowData = [...dimensions];
        
        // For each metric, add current value, historical average, Z-score, and difference
        for (let metricIndex = 0; metricIndex < metrics.length; metricIndex++) {
            const currentValue = parseFloat(metrics[metricIndex]) || 0;
            
            // Get statistical data from varianceData
            let historicalMean = 0;
            let zScore = 0;
            
            if (varianceData && varianceData[rowIndex] && varianceData[rowIndex][metricIndex]) {
                historicalMean = varianceData[rowIndex][metricIndex].historicalMean || 0;
                zScore = varianceData[rowIndex][metricIndex].zScore || 0;
            }
            
            Logger.log(`  Metric ${metricIndex}:`);
            Logger.log(`    Current value: ${currentValue}`);
            Logger.log(`    Historical mean: ${historicalMean}`);
            Logger.log(`    Z-score: ${zScore}`);
            
            rowData.push(currentValue, historicalMean, zScore);
        }
        
        dataRows.push(rowData);
    }
    
    // Write data rows
    if (dataRows.length > 0) {
        sheet.getRange(4, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
    }
    
    // Get z-score threshold from selections or use default
    const zScoreThreshold = Number(selections.zScoreThreshold || 3);
    Logger.log(`Using z-score threshold for conditional formatting: ${zScoreThreshold}`);
    
    // Format numbers
    for (let i = 0; i < metricHeaders.length; i++) {
        const startCol = dimensionHeaders.length + (i * 3) + 1;
        
        // Format current and historical values
        if (dataRows.length > 0) {
            sheet.getRange(4, startCol, dataRows.length, 1).setNumberFormat('#,##0.00');
            sheet.getRange(4, startCol + 1, dataRows.length, 1).setNumberFormat('#,##0.00');
            // Change from percentage format to number format for z-scores
            sheet.getRange(4, startCol + 2, dataRows.length, 1).setNumberFormat('0.00');
        }
        
        // Add conditional formatting for Z-score thresholds
        const zScoreCol = startCol + 2;
        const zScoreRange = sheet.getRange(4, zScoreCol, Math.max(1, dataRows.length), 1);

        // Add conditional formatting for positive anomalies (z-score > threshold)
        const positiveRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberGreaterThan(zScoreThreshold)
            .setBackground('#F4C7C3') // Red background for anomalies
            .setRanges([zScoreRange])
            .build();

        // Add conditional formatting for negative anomalies (z-score < -threshold)
        const negativeRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberLessThan(-zScoreThreshold)
            .setBackground('#F4C7C3') // Red background for anomalies  
            .setRanges([zScoreRange])
            .build();
            
        // Add conditional formatting for normal values (within threshold)
        const normalRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberBetween(-zScoreThreshold, zScoreThreshold)
            .setBackground('#D9EAD3') // Green background for normal values
            .setRanges([zScoreRange])
            .build();

        // Apply the conditional formatting rules
        const rules = sheet.getConditionalFormatRules();
        rules.push(normalRule);  // Add normal rule first (will be overridden by anomaly rules)
        rules.push(positiveRule);
        rules.push(negativeRule);
        sheet.setConditionalFormatRules(rules);
    }
    
    // Add frozen rows and auto-resize columns
    sheet.setFrozenRows(3);
    formatSheet(sheet, 3);
}

// Modify getGA4Report to handle hourly filtering correctly
function getGA4Report(propertyId, dimensions, metrics, dateRanges, options = {}) {
    try {
        // Create new request objects using the Analytics Data API
        const request = AnalyticsData.newRunReportRequest();
        
        // Add dimensions
        request.dimensions = dimensions.map(d => {
            const dimension = AnalyticsData.newDimension();
            dimension.name = d.name;
            return dimension;
        });
        
        // Add metrics
        request.metrics = metrics.map(m => {
            const metric = AnalyticsData.newMetric();
            metric.name = m.name;
            return metric;
        });
        
        // Add date ranges
        request.dateRanges = Array.isArray(dateRanges) ? dateRanges.map(range => {
            const dateRange = AnalyticsData.newDateRange();
            dateRange.startDate = range.startDate;
            dateRange.endDate = range.endDate;
            return dateRange;
        }) : [(() => {
            const dateRange = AnalyticsData.newDateRange();
            dateRange.startDate = dateRanges.startDate;
            dateRange.endDate = dateRanges.endDate;
            return dateRange;
        })()];
        
        // If hourly reporting is enabled and hour is specified, add the hour dimension and filter
        if (options.hourly && options.hour !== undefined) {
            // Check if hour dimension is already included
            const hasHourDimension = dimensions.some(d => d.name === 'hour');
            
            if (!hasHourDimension) {
                // Add hour dimension
                const hourDimension = AnalyticsData.newDimension();
                hourDimension.name = 'hour';
                request.dimensions.unshift(hourDimension); // Add to beginning of array
            }
            
            // Create the filter for the specific hour - use the hour as a simple string without padding
            const hourString = options.hour.toString(); // Just use the hour as a string without padding
            
            const filterExpression = AnalyticsData.newFilterExpression();
            filterExpression.filter = AnalyticsData.newFilter();
            filterExpression.filter.fieldName = 'hour';
            filterExpression.filter.stringFilter = AnalyticsData.newStringFilter();
            filterExpression.filter.stringFilter.matchType = 'EXACT';
            filterExpression.filter.stringFilter.value = hourString;
            
            // Add the filter to the request
            request.dimensionFilter = filterExpression;
            
            Logger.log(`Filtering for hour: ${hourString}`);
        }
        
        // Make the API call
        Logger.log(`Making API call to GA4 for property: ${propertyId}`);
        const response = AnalyticsData.Properties.runReport(request, `properties/${propertyId}`);
        
        // Log response for debugging
        if (response && response.rows) {
            Logger.log(`API call successful! Received ${response.rows.length} rows of data.`);
        } else {
            Logger.log("API response has no rows");
        }
        
        return formatGA4Response(response, options);
    } catch (error) {
        Logger.log(`Error in getGA4Report: ${error.toString()}`);
        if (error.message) {
            Logger.log(`Error message: ${error.message}`);
        }
        return null;
    }
}

// Update formatGA4Response to handle hourly filtering
function formatGA4Response(response, options = {}) {
    if (!response || !response.rows) {
        return null;
    }

    // Extract headers
    const dimensionHeaders = response.dimensionHeaders.map(h => h.name);
    const metricHeaders = response.metricHeaders.map(h => h.name);
    const headers = [...dimensionHeaders, ...metricHeaders];

    // Format rows
    let rows = response.rows.map(row => {
        const dimensionValues = row.dimensionValues.map(d => d.value);
        const metricValues = row.metricValues.map(m => m.value);
        return [...dimensionValues, ...metricValues];
    });

    // If hourly filtering is enabled, remove the hour dimension from the results
    if (options.hourly && dimensionHeaders.includes('hour')) {
        const hourIndex = dimensionHeaders.indexOf('hour');
        
        // Remove hour from dimension headers
        const filteredDimensionHeaders = dimensionHeaders.filter((_, i) => i !== hourIndex);
        
        // Remove hour from rows
        rows = rows.map(row => row.filter((_, i) => i !== hourIndex));
        
        return {
            headers: [...filteredDimensionHeaders, ...metricHeaders],
            dimensionHeaders: filteredDimensionHeaders,
            rows: rows,
            rowCount: response.rowCount,
            metadata: response.metadata
        };
    }

    return {
        headers: headers,
        dimensionHeaders: dimensionHeaders,
        rows: rows,
        rowCount: response.rowCount,
        metadata: response.metadata
    };
}

// Modify createAnomalyReportHTML to return an object
function createAnomalyReportHTML() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 3) {
        return "<p>No anomaly data available.</p>";
    }
    
    // Get property ID and configuration values
    const propertyName = PropertiesService.getDocumentProperties().getProperty('selectedPropertyName') || 'Unknown Property';
    const propertyId = PropertiesService.getDocumentProperties().getProperty('selectedPropertyId') || 'Unknown Property ID';
    const zScoreThreshold = Number(PropertiesService.getDocumentProperties().getProperty('zScoreThreshold')) || 3;
    const percentThreshold = Number(PropertiesService.getDocumentProperties().getProperty('varianceThreshold')) || 10;
    
    // Detect if this is a recurring report by checking the header structure
    const isRecurring = data[1] && data[1].some(cell => cell === 'Current' || cell === 'Historical Avg' || cell === 'Z-Score');
    
    // Determine the dimension count and metric count based on the header structure
    // For recurring reports, dimensions + [Current, Historical Avg, Z-Score] for each metric
    // For one-time reports, dimensions + [Current, Previous, Difference] for each metric
    const dimensionCount = data[1].indexOf('Current') !== -1 ? data[1].indexOf('Current') :
                          data[1].indexOf(data[2][0]) !== -1 ? data[1].indexOf(data[2][0]) : 
                          parseInt(PropertiesService.getDocumentProperties().getProperty('dimensionCount')) || 1;
    
    const metricGroupSize = isRecurring ? 3 : 3; // Both use 3 columns per metric
    const metricCount = Math.floor((data[1].length - dimensionCount) / metricGroupSize);
    
    // Start building HTML table
    let html = `
    <div style="font-family: Arial, sans-serif; max-width: 800px;">
        <h2 style="color: #4285F4;">Anomaly Report: ${propertyName}</h2>
        <p>Property ID: ${propertyId}</p>
        <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%;">
    `;

    // Add header rows 
    html += '<tr>';
    // Add dimension header
    if (dimensionCount > 0) {
        html += `<th colspan="${dimensionCount}" style="background-color: #f3f3f3;">Dimensions</th>`;
    }
    
    // Add metric group headers based on report type
    if (isRecurring) {
        // For recurring reports: Current, Historical Avg, Z-Score for each metric
        for (let i = 0; i < metricCount; i++) {
            const metricColor = getMetricColor(i);
            const metricName = data[2][dimensionCount + (i * metricGroupSize)]; // Get metric name from the second header row
            
            html += `
                <th style="background-color: ${metricColor}; color: white;">Current</th>
                <th style="background-color: ${metricColor}; color: white;">Historical Avg</th>
                <th style="background-color: #FB8C00; color: white;">Z-Score</th>
            `;
        }
    } else {
        // For one-time reports: Current, Previous, Difference for each metric
        for (let i = 0; i < metricCount; i++) {
            const metricColor = getMetricColor(i);
            html += `
                <th style="background-color: ${metricColor}; color: white;">Current</th>
                <th style="background-color: ${metricColor}; color: white;">Previous</th>
                <th style="background-color: #FB8C00; color: white;">Difference (%)</th>
            `;
        }
    }
    html += '</tr>';
    
    // Add metric names as a second header row
    html += '<tr>';
    // Add empty cells for dimension columns
    for (let i = 0; i < dimensionCount; i++) {
        html += `<th style="background-color: #f3f3f3;">${data[2][i]}</th>`;
    }
    
    // Add metric names
    for (let i = 0; i < metricCount; i++) {
        const metricName = data[2][dimensionCount + (i * metricGroupSize)];
        const metricColor = getMetricColor(i);
        
        html += `
            <th style="background-color: ${metricColor}; color: white;">${metricName}</th>
            <th style="background-color: ${metricColor}; color: white;">${metricName}</th>
            <th style="background-color: #FB8C00; color: white;">${metricName}</th>
        `;
    }
    html += '</tr>';

    // Add data rows, but only include rows with anomalies
    let anomalyFound = false;
    let rowCount = 0;
    let totalAnomalyCount = 0;
    
    // Start from the first data row (after headers)
    const startRow = isRecurring ? 3 : 3;
    
    for (let i = startRow; i < data.length; i++) {
        const row = data[i];
        let hasAnomaly = false;
        
        // Check for anomalies based on report type
        if (isRecurring) {
            // For recurring reports: Check Z-score columns
            for (let j = 0; j < metricCount; j++) {
                const zScoreCol = dimensionCount + (j * metricGroupSize) + 2; // Z-Score is the 3rd column in each metric group
                const zScore = Math.abs(Number(row[zScoreCol]));
                if (zScore > zScoreThreshold) {
                    hasAnomaly = true;
                    anomalyFound = true;
                    totalAnomalyCount++;
                    break;
                }
            }
        } else {
            // For one-time reports: Check difference percentage columns
            for (let j = 0; j < metricCount; j++) {
                const diffCol = dimensionCount + (j * metricGroupSize) + 2; // Difference is the 3rd column in each metric group
                const diffPercent = Math.abs(Number(row[diffCol]));
                if (diffPercent > percentThreshold) {
                    hasAnomaly = true;
                    anomalyFound = true;
                    totalAnomalyCount++;
                    break;
                }
            }
        }
        
        // Display only rows with anomalies (limit to first 20)
        if (hasAnomaly && rowCount < 20) {
            html += '<tr>';
            
            // Add dimension values
            for (let j = 0; j < dimensionCount; j++) {
                html += `<td>${row[j]}</td>`;
            }
            
            // Add metric values with appropriate formatting
            for (let j = 0; j < metricCount; j++) {
                const baseCol = dimensionCount + (j * metricGroupSize);
                const currentValue = Number(row[baseCol]);
                const comparisonValue = Number(row[baseCol + 1]);
                const thirdValue = Number(row[baseCol + 2]); // Z-Score or Difference
                
                // Format the numbers appropriately
                const formattedCurrent = currentValue.toLocaleString(undefined, {maximumFractionDigits: 2});
                const formattedComparison = comparisonValue.toLocaleString(undefined, {maximumFractionDigits: 2});
                let formattedThird;
                
                if (isRecurring) {
                    // For Z-Score, use fixed 2 decimal places
                    formattedThird = thirdValue.toFixed(2);
                    
                    // Add background color based on Z-Score
                    const bgColor = Math.abs(thirdValue) > zScoreThreshold ? '#F4C7C3' : '#D9EAD3';
                    
                    html += `
                        <td>${formattedCurrent}</td>
                        <td>${formattedComparison}</td>
                        <td style="background-color: ${bgColor};">${formattedThird}</td>
                    `;
                } else {
                    // For percentage difference, format with % sign
                    formattedThird = thirdValue.toFixed(2) + '%';
                    
                    // Add background color based on percentage difference
                    const bgColor = Math.abs(thirdValue) > percentThreshold ? '#F4C7C3' : '#D9EAD3';
                    
                    html += `
                        <td>${formattedCurrent}</td>
                        <td>${formattedComparison}</td>
                        <td style="background-color: ${bgColor};">${formattedThird}</td>
                    `;
                }
            }
            
            html += '</tr>';
            rowCount++;
        }
    }
    
    html += `</table>`;

    // Add footer text with threshold and count information
    html += `
        <p style="font-size: 0.9em; color: #666;">
            ${isRecurring ? 
                `This report shows metrics with Z-scores exceeding ${zScoreThreshold}.` : 
                `This report shows metrics with percentage differences exceeding ${percentThreshold}%.`
            }
            ${!anomalyFound ? 
                '<br><strong>No anomalies detected based on current threshold.</strong>' : 
                `<br>Showing ${rowCount} of ${totalAnomalyCount} total anomalies${totalAnomalyCount > 20 ? 
                    '. Please check the Google Sheet for the complete list.' : 
                    '.'}`
            }
        </p>
    </div>`;
    
    // Return object with both the HTML and anomaly count
    return {
        anomalyCount: totalAnomalyCount,
        anomalyTable: html
    };
}

// Modify sendAnomalyEmail to check for anomalies before sending
function sendAnomalyEmail() {
    const propertyId = PropertiesService.getDocumentProperties().getProperty('selectedPropertyId');
    const propertyName = PropertiesService.getDocumentProperties().getProperty('selectedPropertyName');
    const sender = PropertiesService.getDocumentProperties().getProperty('sender') || 'AnomalyDetector9000@cardinalpath.com';
    const receiver = PropertiesService.getDocumentProperties().getProperty('receiver') || Session.getActiveUser().getEmail();
    const spreadsheetUrl = SpreadsheetApp.getActive().getUrl();

    // Generate HTML table and get anomaly count
    const { anomalyCount, anomalyTable } = createAnomalyReportHTML();
    
    Logger.log("Anomaly Count:");
    Logger.log(anomalyCount);
    // Only send email if there are anomalies
    if (anomalyCount === 0) {
        Logger.log("No anomalies detected - skipping email notification");
        return;
    }
    
    const message = {
      to: receiver,
      subject: `GA Anomaly Detected: ${propertyName} (${propertyId})`,
      htmlBody: `
        <p>Beep boop!</p>
        <p>View complete report in <a href="${spreadsheetUrl}" style="color: #1a73e8; text-decoration: none; font-weight: 500;">Spreadsheet</a>.</p>
        ${anomalyTable}
        <p>Please review these metrics as they exceed the configured variance threshold.</p>
        <p>This is an automated message from the Post Launch Validator tool.</p>
      `,
      name: sender
    };
    
    try {
      MailApp.sendEmail(message);
      Logger.log("Email sent successfully");
    } catch (error) {
      Logger.log("Error sending email: " + error.toString());
    }
}

/*
Sample variables structure:

combinedData = [[Chrome, 14588.0, 10561.0, 38.13085882018748], [Safari, 1139.0, 1207.0, -5.633802816901409], [Edge, 941.0, 1078.0, -12.708719851576994], [Firefox, 285.0, 438.0, -34.93150684931507], [Safari (in-app), 46.0, 303.0, -84.81848184818482], [Samsung Internet, 21.0, 54.0, -61.111111111111114], [Android Webview, 19.0, 36.0, -47.22222222222222], [Opera, 12.0, 31.0, -61.29032258064516], [Internet Explorer, 4.0, 27.0, -85.18518518518519], [PaleMoon, 4.0, 18.0, -77.77777777777779], [(not set), 2.0, 5.0, -60.0]]
*/

/* Tests */
function testGetCommonGA4DimensionsAndMetrics() {
    var propertyId = '189677427';
    
    // Get dimensions and metrics for the selected property
    const dimensions = getGA4Dimensions(propertyId);
    const metrics = getGA4Metrics(propertyId);
    
    Logger.log("dimensions:", dimensions);
    Logger.log("metrics:", metrics);
}

function testOneTimeReport() {
    var propertyId = '189677427';
    PropertiesService.getDocumentProperties().setProperty('selectedPropertyId', propertyId);

    var selections = {
        "dates": {
            "range1": {
                "startDate": "2024-11-01",
                "endDate": "2024-11-30"
            },
            "range2": {
                "startDate": "2024-12-01",
                "endDate": "2024-12-31"
            }
        },
        "variance": "10",
        "dimensions": [
            "browser",
            "city",
            "hostname",
        ],
        "metrics": [
            "activeUsers",
            "eventCount",
            "newUsers",
        ]
    };

    writeReportsToSheet(selections);
}

function testSendEmail() {
    // Example usage
    const sampleTable = `
    <table border="1">
    <tr>
        <th>Metric</th>
        <th>Value</th>
        <th>Expected Range</th>
    </tr>
    <tr>
        <td>Temperature</td>
        <td>35°C</td>
        <td>20-30°C</td>
    </tr>
    </table>
    `;

    sendAnomalyEmail(sampleTable);
}

// Test function that simulates a user selection for hourly recurring reports
function testHourlyRecurringReport() {
  try {
    // Set a property ID if not already set
    const propertyId = PropertiesService.getDocumentProperties().getProperty('selectedPropertyId') || '189677427';
    PropertiesService.getDocumentProperties().setProperty('selectedPropertyId', propertyId);
    
    Logger.log(`Using property ID: ${propertyId}`);
    
    const dummySelections = {
        "reportType": "recurring",
        "dates": {
            "recurring": true,
            "recurringOption": "hourly"
        },
        "variance": "10",
        "zScoreThreshold": "3",
        "dimensions": [
            "eventName",
        ],
        "metrics": [
            "eventCount",
        ]
    };
    
    Logger.log("Testing hourly report with the following selections:");
    Logger.log(JSON.stringify(dummySelections, null, 2));
    
    // Call the actual report generation function
    const result = writeReportsToSheet(dummySelections);
    
    Logger.log("Test completed with result:");
    Logger.log(JSON.stringify(result));
    
    return "Test completed. Check the newly created sheet and logs for details.";
  } catch (error) {
    Logger.log(`Error in testHourlyReportWithSelections: ${error.toString()}`);
    if (error.message) {
      Logger.log(`Error message: ${error.message}`);
    }
    if (error.stack) {
      Logger.log(`Stack trace: ${error.stack}`);
    }
    return `Test failed with error: ${error.toString()}`;
  }
}

// Test function to debug hourly report calculations
function testHourlyReportCalculations() {
    try {
        // Set a property ID if not already set
        const propertyId = PropertiesService.getDocumentProperties().getProperty('selectedPropertyId') || '189677427';
        PropertiesService.getDocumentProperties().setProperty('selectedPropertyId', propertyId);
        
        Logger.log(`Using property ID: ${propertyId}`);

        const dummySelections = {
            "reportType": "recurring",
            "dates": {
                "recurring": true,
                "recurringOption": "hourly"
            },
            "variance": "10",
            "zScoreThreshold": "3",
            "dimensions": [
                "eventName",
            ],
            "metrics": [
                "eventCount",
            ]
        };
        
        Logger.log("Testing hourly report calculations with the following selections:");
        Logger.log(JSON.stringify(dummySelections, null, 2));
        
        // Call the validation function directly to see the calculated values
        const result = validateSelectionsAndFetchData(dummySelections);
        
        Logger.log("Validation result:");
        Logger.log(JSON.stringify(result, null, 2));
        
        return "Test completed. Check logs for details.";
    } catch (error) {
        Logger.log(`Error in testHourlyReportCalculations: ${error.toString()}`);
        if (error.message) {
            Logger.log(`Error message: ${error.message}`);
        }
        if (error.stack) {
            Logger.log(`Stack trace: ${error.stack}`);
        }
        return `Test failed with error: ${error.toString()}`;
    }
}

// Test function for the recurring daily option
function testDailyRecurringReport() {
    try {
        // Set a property ID if not already set
        const propertyId = PropertiesService.getDocumentProperties().getProperty('selectedPropertyId') || '189677427';
        PropertiesService.getDocumentProperties().setProperty('selectedPropertyId', propertyId);
        
        Logger.log(`Using property ID: ${propertyId}`);
        
        // Create a dummy selections object that mimics what would come from the UI
        const dummySelections = {
            "reportType": "recurring",
            "dates": {
                "recurring": true,
                "recurringOption": "daily"
            },
            "variance": "10",
            "zScoreThreshold": "3",
            "dimensions": [
                "eventName",
            ],
            "metrics": [
                "eventCount",
            ]
        };
        
        Logger.log("Testing daily recurring report with the following selections:");
        Logger.log(JSON.stringify(dummySelections, null, 2));
        
        // Call the actual report generation function
        const result = writeReportsToSheet(dummySelections);
        
        Logger.log("Test completed with result:");
        Logger.log(JSON.stringify(result));
        
        return "Test completed. Check the newly created sheet and logs for details.";
    } catch (error) {
        Logger.log(`Error in testDailyRecurringReport: ${error.toString()}`);
        if (error.message) {
            Logger.log(`Error message: ${error.message}`);
        }
        if (error.stack) {
            Logger.log(`Stack trace: ${error.stack}`);
        }
        return `Test failed with error: ${error.toString()}`;
    }
}