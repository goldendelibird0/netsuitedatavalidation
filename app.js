// ==========================================================================
// CONFIGURATION & GLOBAL STATE & ERROR HANDLING
// ==========================================================================
window.addEventListener('error', function(event) {
    alert("Unhandled Error: " + event.message + "\\nSource: " + event.filename + ":" + event.lineno);
});
window.addEventListener('unhandledrejection', function(event) {
    alert("Unhandled Promise Rejection: " + event.reason);
});

const SPREADSHEETS = {
    coa: {
        id: "1ZqWqTC0p7ODVoa_8volrX0ZO57KNWh4nZWUuC72kNPE",
        name: "Chart of Accounts (COA)",
        reportName: "coa_validation_report.xlsx"
    },
    employee: {
        id: "10XGfKsnJtdZc6FNTnOj1YW-SGVOC6v5ubPF0ASMeWiY",
        name: "Employees",
        reportName: "employee_validation_report.xlsx"
    },
    inventory: {
        id: "1hbIYhEIYjYh7tPK2G2KlDQ5EWnPsKGhVMrXFWmX9B-g",
        name: "Inventory Items",
        reportName: "inventory_validation_report.xlsx"
    },
    customer: {
        id: "1_PlTBeT5IA8RVvvJhu-ojQkSktFITAik",
        name: "Customers",
        reportName: "customer_validation_report.xlsx"
    },
    vendor: {
        id: "1eA6JjN_2TlFx8fQ5oQyEyJAfmN9i6Zm4",
        name: "Vendors",
        reportName: "vendor_validation_report.xlsx"
    }
};

let currentWorkbook = null;
let currentReportBlob = null;
let currentReportName = "";

// ==========================================================================
// SHEET URL HELPER
// ==========================================================================
function buildSheetUrl(sheetId, sheetName) {
    return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
}

// ==========================================================================
// HELPER FUNCTIONS
// ==========================================================================
function cleanValue(val) {
    if (val === null || val === undefined) return "";
    return String(val).trim();
}

function isBlank(val) {
    return cleanValue(val) === "";
}

function normalizeColumnName(name) {
    return String(name)
        .trim()
        .toLowerCase()
        .replace(/\s+/g, "");
}

// ==========================================================================
// APP INITIALIZATION & DOM ROUTING
// ==========================================================================
document.addEventListener("DOMContentLoaded", () => {
    const templateSelect = document.getElementById("template-select");
    const schemaLink = document.getElementById("schema-link");
    const dropzone = document.getElementById("dropzone");
    const fileInput = document.getElementById("file-input");
    const downloadBtn = document.getElementById("download-btn");
    
    // Welcome Screen Transition
    const startAppBtn = document.getElementById("start-app-btn");
    const welcomeScreen = document.getElementById("welcome-screen");
    const appContainer = document.getElementById("app-container");
    
    startAppBtn.addEventListener("click", () => {
        welcomeScreen.classList.add("fade-out");
        welcomeScreen.addEventListener("transitionend", () => {
            welcomeScreen.style.display = "none";
            appContainer.classList.remove("hidden");
        }, { once: true });
    });
    
    // Set initial schema link
    updateSchemaLink();

    templateSelect.addEventListener("change", () => {
        updateSchemaLink();
        resetStatus();
    });

    function updateSchemaLink() {
        const selected = templateSelect.value;
        const sheetId = SPREADSHEETS[selected].id;
        schemaLink.href = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
    }

    // Drag and Drop Logic
    dropzone.addEventListener("click", () => fileInput.click());
    
    dropzone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropzone.classList.add("dragover");
    });

    dropzone.addEventListener("dragleave", () => {
        dropzone.classList.remove("dragover");
    });

    dropzone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropzone.classList.remove("dragover");
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener("change", () => {
        if (fileInput.files.length > 0) {
            handleFile(fileInput.files[0]);
        }
    });

    // Tab Switching Logic
    const tabs = document.querySelectorAll(".tab-btn");
    tabs.forEach(tab => {
        tab.addEventListener("click", () => {
            tabs.forEach(t => t.classList.remove("active"));
            tab.classList.add("active");
            
            const targetId = tab.dataset.tab;
            const contents = document.querySelectorAll(".tab-content");
            contents.forEach(content => content.classList.remove("active"));
            document.getElementById(targetId).classList.add("active");
        });
    });

    // Download Button Click Handler
    downloadBtn.addEventListener("click", () => {
        if (currentReportBlob) {
            const link = document.createElement("a");
            link.href = URL.createObjectURL(currentReportBlob);
            link.download = currentReportName;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
});

function resetStatus() {
    document.getElementById("status-card").classList.add("hidden");
    document.getElementById("results-panel").classList.add("hidden");
    document.getElementById("download-btn").disabled = true;
    currentWorkbook = null;
    currentReportBlob = null;
}

// ==========================================================================
// CORE DATA LOADING & RUNNER ROUTINE
// ==========================================================================
async function handleFile(file) {
    if (!file.name.endsWith(".xlsx")) {
        alert("Invalid file format. Please upload a .xlsx Excel spreadsheet.");
        return;
    }

    const templateType = document.getElementById("template-select").value;
    const spreadsheetConfig = SPREADSHEETS[templateType];
    
    // UI Loading State
    const statusCard = document.getElementById("status-card");
    const statusTitle = document.getElementById("status-title");
    const statusFileName = document.getElementById("status-file-name");
    const statusSpinner = document.getElementById("status-spinner");
    const statusIcon = document.getElementById("status-icon");
    const statusDetails = document.getElementById("status-details");
    
    statusCard.classList.remove("hidden");
    statusCard.className = "status-card"; // reset classes
    statusDetails.classList.add("hidden");
    statusSpinner.style.display = "block";
    statusIcon.style.display = "none";
    statusTitle.textContent = "Loading configuration and parsing schema...";
    statusFileName.textContent = file.name;
    document.getElementById("download-btn").disabled = true;
    document.getElementById("results-panel").classList.add("hidden");

    try {
        // 1. Fetch Schema from Google Sheets
        const schemaUrl = buildSheetUrl(spreadsheetConfig.id, "schema");
        const schemaResponse = await fetch(schemaUrl);
        if (!schemaResponse.ok) throw new Error("Failed to download schema configuration sheet.");
        const schemaCsvText = await schemaResponse.text();
        
        const parsedSchema = Papa.parse(schemaCsvText, { header: true, skipEmptyLines: true });
        const schemaRows = parsedSchema.data.filter(row => cleanValue(row.column_name) !== "");

        // 2. Fetch Validation Lists
        statusTitle.textContent = "Loading validation lists...";
        const validationLists = {};
        const validationListNames = [...new Set(
            schemaRows
                .map(r => cleanValue(r.validation_list))
                .filter(val => val !== "" && val !== "nan")
                .map(val => normalizeColumnName(val))
        )];

        for (const listName of validationListNames) {
            const listUrl = buildSheetUrl(spreadsheetConfig.id, listName);
            const listResponse = await fetch(listUrl);
            if (listResponse.ok) {
                const listCsvText = await listResponse.text();
                const parsedList = Papa.parse(listCsvText, { header: false, skipEmptyLines: true });
                validationLists[listName] = parsedList.data
                    .map(r => cleanValue(r[0]))
                    .filter(val => val !== "");
            }
        }

        // 3. Load Excel File using ExcelJS
        statusTitle.textContent = "Reading uploaded Excel spreadsheet...";
        const arrayBuffer = await file.arrayBuffer();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);
        const worksheet = workbook.worksheets[0];
        
        // 4. Validate and Execute
        statusTitle.textContent = "Validating dataset...";
        const validationResults = runValidation(worksheet, schemaRows, validationLists, templateType);

        // 5. Generate and highlight Report Workbook
        statusTitle.textContent = "Generating validation report...";
        const reportBlob = await generateReportBlob(worksheet, schemaRows, validationResults, templateType);
        
        // 6. Update UI with Results
        currentReportBlob = reportBlob;
        currentReportName = spreadsheetConfig.reportName;

        statusSpinner.style.display = "none";
        statusDetails.classList.remove("hidden");
        
        document.getElementById("stat-rows-checked").textContent = validationResults.rowsChecked;
        document.getElementById("stat-rows-issues").textContent = validationResults.rowsWithIssuesCount;
        document.getElementById("stat-core-issues").textContent = validationResults.coreIssues.length;

        const resultsPanel = document.getElementById("results-panel");
        resultsPanel.classList.remove("hidden");

        // Set Tab Headers & Counts
        document.querySelector("[data-tab='row-issues-tab']").textContent = `Row Issues (${validationResults.rowsWithIssuesCount})`;
        document.querySelector("[data-tab='core-issues-tab']").textContent = `Core Issues (${validationResults.coreIssues.length})`;

        // Populate Table Contents
        populateRowIssuesTable(validationResults.rowIssuesDetailed);
        populateCoreIssuesTable(validationResults.coreIssues);
        populateRulesRefTable(validationResults.ruleReference);

        if (validationResults.coreIssues.length > 0) {
            statusCard.classList.add("error");
            statusIcon.className = "fa-solid fa-triangle-exclamation status-large-icon";
            statusTitle.textContent = "Errors found: Check structural validation";
        } else if (validationResults.rowsWithIssuesCount > 0) {
            statusCard.classList.add("warning");
            statusIcon.className = "fa-solid fa-circle-exclamation status-large-icon";
            statusTitle.textContent = "Validation complete: Row issues detected";
        } else {
            statusCard.classList.add("success");
            statusIcon.className = "fa-solid fa-circle-check status-large-icon";
            statusTitle.textContent = "File is clean! Ready for NetSuite import";
        }
        document.getElementById("download-btn").disabled = false;

    } catch (err) {
        console.error(err);
        statusSpinner.style.display = "none";
        statusIcon.className = "fa-solid fa-circle-xmark status-large-icon";
        statusIcon.style.display = "block";
        statusCard.classList.add("error");
        statusTitle.textContent = "An error occurred during validation";
        statusFileName.textContent = err.message;
    }
}

// ==========================================================================
// VALIDATION ENGINE (PORTED FROM PYTHON)
// ==========================================================================
function runValidation(worksheet, schemaRows, validationLists, templateType) {
    const coreIssues = [];
    const rowIssues = {}; // excel_row -> array of {column, rule, message}

    // 1. Gather original headers
    const excelHeaders = [];
    const headerRow = worksheet.getRow(1);
    for (let c = 1; c <= worksheet.columnCount; c++) {
        excelHeaders.push(cleanValue(headerRow.getCell(c).value));
    }

    // Trim trailing empty cells from headers
    while (excelHeaders.length > 0 && excelHeaders[excelHeaders.length - 1] === "") {
        excelHeaders.pop();
    }

    // Helper to normalize strings for comparison (removes all non-alphanumeric characters and lowercase)
    function normalizeHeader(str) {
        if (!str) return "";
        return String(str).toLowerCase().replace(/[^a-z0-9]/g, "");
    }

    const COLUMN_ALIASES = {
        "externalid": ["extid", "ext", "external", "externalid", "extidno"],
        "employeeid": ["empid", "employeeid", "empno", "employeeno"],
        "vendorid": ["vendorid", "vendorno", "vendorcode", "vendid"],
        "itemnamenumber": ["itemnamenumber", "itemname", "itemnumber", "itemcode", "partnumber", "partno"],
        "entityid": ["entityid", "customerid", "customercode", "custid", "entityno"],
        "unitstype": ["unitstype", "units", "unitstype", "unitofmeasure", "uom"],
        "stockdescrption": ["stockdescrption", "stockdescription", "stockdesc"], // handle schema typo
        "purchasedescription": ["purchasedescription", "purchasedesc"],
        "salesdescription": ["salesdescription", "salesdesc"],
        "vatregnumber": ["vatregnumber", "vatnumber", "taxid", "tin", "taxidtin"]
    };

    const mappedColumns = {}; // schemaRow.column_name -> 1-based excel column index
    const usedExcelIndices = new Set();

    // Map each schema column to an Excel column index
    schemaRows.forEach(schemaRow => {
        const expectedName = schemaRow.column_name;
        const normExpected = normalizeHeader(expectedName);

        // 1. Try direct exact match (case-insensitive)
        let matchedIndex = -1;
        for (let i = 0; i < excelHeaders.length; i++) {
            if (usedExcelIndices.has(i)) continue;
            if (excelHeaders[i].toLowerCase() === expectedName.toLowerCase()) {
                matchedIndex = i;
                break;
            }
        }

        // 2. Try normalized match
        if (matchedIndex === -1) {
            for (let i = 0; i < excelHeaders.length; i++) {
                if (usedExcelIndices.has(i)) continue;
                if (normalizeHeader(excelHeaders[i]) === normExpected) {
                    matchedIndex = i;
                    break;
                }
            }
        }

        // 3. Try alias match
        if (matchedIndex === -1) {
            const aliases = COLUMN_ALIASES[normExpected] || [];
            for (let i = 0; i < excelHeaders.length; i++) {
                if (usedExcelIndices.has(i)) continue;
                const normActual = normalizeHeader(excelHeaders[i]);
                if (aliases.includes(normActual)) {
                    matchedIndex = i;
                    break;
                }
            }
        }

        if (matchedIndex !== -1) {
            mappedColumns[expectedName] = matchedIndex + 1; // 1-based index
            usedExcelIndices.add(matchedIndex);

            // Report deviation if the name doesn't match expectedName
            const actualName = excelHeaders[matchedIndex];
            if (actualName !== expectedName) {
                coreIssues.push({
                    issue: `Column header deviation: Found '${actualName}' at column ${matchedIndex + 1} instead of '${expectedName}'`,
                    fix: `We will automatically rename this header to '${expectedName}' in your downloaded report.`
                });
            }
        } else {
            // Missing column
            if (cleanValue(schemaRow.required).toUpperCase() === "TRUE") {
                coreIssues.push({
                    issue: `Missing required column: '${expectedName}'`,
                    fix: `Insert a column named '${expectedName}' in your sheet.`
                });
            }
        }
    });

    // Report unrecognized extra columns in the Excel file
    for (let i = 0; i < excelHeaders.length; i++) {
        if (!usedExcelIndices.has(i)) {
            const actualName = excelHeaders[i];
            if (actualName !== "") {
                coreIssues.push({
                    issue: `Unrecognized/Extra column at column ${i + 1}: '${actualName}'`,
                    fix: "This column will be kept in the downloaded report but is not part of the NetSuite schema."
                });
            }
        }
    }

    // 4. Build Rules Config
    const rulesConfig = {};
    const ruleReference = {
        required: "Required field is missing",
        max_length: "Field exceeds allowed character limit",
        validation_list: "Field value does not exist in approved validation list",
        duplicate: "Field value must be unique"
    };

    // Add format rules references if applicable
    if (templateType === "customer" || templateType === "vendor") {
        ruleReference.email_format = "Email format is invalid";
        ruleReference.phone_format = "Phone number format is invalid";
    }
    if (templateType === "employee") {
        ruleReference.email = "Email format is invalid";
        ruleReference.phone = "Phone format is invalid";
        ruleReference.date = "Date format is invalid";
        ruleReference.conditional_required = "Field required based on another field";
        ruleReference.password_match = "Passwords do not match";
    }
    if (templateType === "customer") {
        ruleReference.numeric_format = "Value must be numeric";
        ruleReference.customer_type_logic = "Individual/company customer fields are inconsistent";
    }
    if (templateType === "vendor") {
        ruleReference.boolean_format = "Value must be TRUE or FALSE";
        ruleReference.vendor_type_logic = "Individual/company vendor fields are inconsistent";
    }
    if (templateType === "coa") {
        ruleReference.conditional_required = "Field required based on another field";
    }

    schemaRows.forEach(row => {
        const column = normalizeColumnName(row.column_name);
        rulesConfig[column] = {
            required: cleanValue(row.required).toUpperCase() === "TRUE",
            unique: cleanValue(row.unique).toUpperCase() === "TRUE",
            max_length: !isBlank(row.max_length) ? parseInt(row.max_length, 10) : null,
            validation_list: !isBlank(row.validation_list) ? normalizeColumnName(row.validation_list) : null
        };
        
        // Formats check
        if (row.format_type) {
            rulesConfig[column].format = cleanValue(row.format_type).toLowerCase();
        } else if (row.format) { // customer/vendor sheets use format instead of format_type
            rulesConfig[column].format = cleanValue(row.format).toLowerCase();
        }
    });

    const rowCount = worksheet.rowCount;
    if (rowCount <= 1) {
        coreIssues.push({
            issue: "File contains no rows",
            fix: "Populate the template with records"
        });
    }

    // 6. Map excel rows into an array of objects
    const dataRows = []; // index 0 matches excel row 2
    for (let r = 2; r <= rowCount; r++) {
        const row = worksheet.getRow(r);
        const rowData = {};
        let isRowBlank = true;
        
        schemaRows.forEach(schemaRow => {
            const expectedNameNormalized = normalizeColumnName(schemaRow.column_name);
            const colIndex = mappedColumns[schemaRow.column_name];
            if (colIndex) {
                const val = cleanValue(row.getCell(colIndex).value);
                if (val !== "") {
                    isRowBlank = false;
                }
                rowData[expectedNameNormalized] = val;
            } else {
                rowData[expectedNameNormalized] = "";
            }
        });
        
        if (!isRowBlank) {
            dataRows.push({
                excelRow: r,
                data: rowData
            });
        }
    }

    // Helper to log row issues
    function addRowIssue(excelRow, column, rule, message) {
        if (!rowIssues[excelRow]) {
            rowIssues[excelRow] = [];
        }
        rowIssues[excelRow].push({ column, rule, message });
    }

    // 7. Row Validation Loop
    dataRows.forEach(rowObj => {
        const excelRow = rowObj.excelRow;
        const rowData = rowObj.data;

        Object.keys(rulesConfig).forEach(column => {
            const schemaRow = schemaRows.find(row => normalizeColumnName(row.column_name) === column);
            if (!schemaRow) return;
            const colIndex = mappedColumns[schemaRow.column_name];
            if (!colIndex) return; // skip if column was not present in the Excel sheet

            const value = rowData[column];
            const rules = rulesConfig[column];

            // Required Check
            if (rules.required) {
                if (isBlank(value)) {
                    addRowIssue(excelRow, column, "required", `Blank ${column}`);
                    return; // skip subsequent validation if blank
                }
            }

            // Skip optional empty fields
            if (isBlank(value)) return;

            // Character Length Check
            if (rules.max_length !== null) {
                if (value.length > rules.max_length) {
                    addRowIssue(excelRow, column, "max_length", `${column} exceeds ${rules.max_length} characters`);
                }
            }

            // Validation List Check
            if (rules.validation_list) {
                const validOptions = validationLists[rules.validation_list] || [];
                const normVal = value.toLowerCase();
                const normOptions = validOptions.map(opt => opt.toLowerCase());
                if (!normOptions.includes(normVal)) {
                    addRowIssue(excelRow, column, "validation_list", `Invalid ${column}`);
                }
            }

            // Email/Phone Formats Checks
            if (rules.format === "email") {
                const emailPattern = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
                if (!emailPattern.test(value)) {
                    if (templateType === "employee") {
                        addRowIssue(excelRow, column, "email", `Invalid ${column}`);
                    } else {
                        addRowIssue(excelRow, column, "email_format", `Invalid ${column} format`);
                    }
                }
            }

            if (rules.format === "phone") {
                if (templateType === "vendor") {
                    const cleanedPhone = value.replace(/[\s\-\(\)]/g, "");
                    const phMobilePattern = /^(09\d{9}|\+639\d{9})$/;
                    const phLandlinePattern = /^(02\d{8}|\+632\d{8})$/;
                    if (!phMobilePattern.test(cleanedPhone) && !phLandlinePattern.test(cleanedPhone)) {
                        addRowIssue(excelRow, column, "phone_format", `Invalid ${column} format`);
                    }
                } else if (templateType === "employee") {
                    const phonePattern = /^[0-9\-\+\(\)\s\.extEXT]+$/i;
                    if (!phonePattern.test(value)) {
                        addRowIssue(excelRow, column, "phone", `Invalid ${column}`);
                    }
                } else { // customer
                    const phonePattern = /^(\d{3}-\d{3}-\d{4}|\(\d{3}\)\s\d{3}-\d{4}|1-\d{3}-\d{3}-\d{4}|1\s\(\d{3}\)\s\d{3}-\d{4}|\d{3}-\d{3}-\d{4}\sext\s\d+|\+\d{1,3}\s\(\d+\)\s[\d-]+)$/i;
                    if (!phonePattern.test(value)) {
                        addRowIssue(excelRow, column, "phone_format", `Invalid ${column} format`);
                    }
                }
            }

            // Date Format check for employees
            if (rules.format === "date") {
                // simple date validation
                const dateVal = Date.parse(value);
                if (isNaN(dateVal)) {
                    if (templateType === "employee") {
                        addRowIssue(excelRow, column, "date", `Invalid ${column}`);
                    } else {
                        addRowIssue(excelRow, column, "date_format", `Invalid ${column} date`);
                    }
                }
            }

            // Numeric check for customer
            if (rules.format === "numeric") {
                if (isNaN(Number(value))) {
                    addRowIssue(excelRow, column, "numeric_format", `${column} must be numeric`);
                }
            }

            // Boolean check for vendor
            if (rules.format === "boolean") {
                const upperVal = value.toUpperCase();
                if (upperVal !== "TRUE" && upperVal !== "FALSE") {
                    addRowIssue(excelRow, column, "boolean_format", `${column} must be TRUE or FALSE`);
                }
            }
        });

        // 8. Custom Conditional Logics

        // COA Business Rules
        if (templateType === "coa") {
            const accType = String(rowData["type"] || "").trim().toLowerCase();
            const currency = String(rowData["currency"] || "").trim();
            if (accType === "bank" && currency === "") {
                addRowIssue(excelRow, "currency", "conditional_required", "Currency required for bank account");
            }
        }

        // Employee Conditional Validation
        if (templateType === "employee") {
            const giveAccess = String(rowData["giveaccess"] || "").trim().toUpperCase();
            const password = String(rowData["password"] || "").trim();
            const confirmPassword = String(rowData["confirmpassword"] || "").trim();
            const role = String(rowData["role"] || "").trim();
            const email = String(rowData["email"] || "").trim();

            if (giveAccess === "TRUE") {
                if (role === "") addRowIssue(excelRow, "role", "conditional_required", "Blank role");
                if (password === "") addRowIssue(excelRow, "password", "conditional_required", "Blank password");
                if (confirmPassword === "") addRowIssue(excelRow, "confirmpassword", "conditional_required", "Blank confirmpassword");
                if (email === "") addRowIssue(excelRow, "email", "conditional_required", "Blank email");
            }

            if (password !== "" && confirmPassword !== "" && password !== confirmPassword) {
                addRowIssue(excelRow, "confirmpassword", "password_match", "Passwords do not match");
            }
        }

        // Customer Conditional Validation
        if (templateType === "customer") {
            const isPerson = String(rowData["isperson"] || "").trim().toUpperCase();
            const firstName = String(rowData["firstname"] || "").trim();
            const lastName = String(rowData["lastname"] || "").trim();
            const companyName = String(rowData["companyname"] || "").trim();

            if (isPerson === "TRUE") {
                if (firstName === "") addRowIssue(excelRow, "firstname", "customer_type_logic", "firstName is required when isPerson is TRUE");
                if (lastName === "") addRowIssue(excelRow, "lastname", "customer_type_logic", "lastName is required when isPerson is TRUE");
                if (companyName !== "") addRowIssue(excelRow, "companyname", "customer_type_logic", "companyName should be blank when isPerson is TRUE");
            }
            if (isPerson === "FALSE") {
                if (companyName === "") addRowIssue(excelRow, "companyname", "customer_type_logic", "companyName is required when isPerson is FALSE");
            }
        }

        // Vendor Conditional Validation
        if (templateType === "vendor") {
            const individual = String(rowData["individual"] || "").trim().toUpperCase();
            const firstName = String(rowData["firstname"] || "").trim();
            const lastName = String(rowData["lastname"] || "").trim();
            const companyName = String(rowData["companyname"] || "").trim();

            if (individual === "TRUE") {
                if (firstName === "") addRowIssue(excelRow, "firstname", "vendor_type_logic", "First Name is required when Individual is TRUE");
                if (lastName === "") addRowIssue(excelRow, "lastname", "vendor_type_logic", "Last Name is required when Individual is TRUE");
                if (companyName !== "") addRowIssue(excelRow, "companyname", "vendor_type_logic", "Company Name should be blank when Individual is TRUE");
            }
            if (individual === "FALSE") {
                if (companyName === "") addRowIssue(excelRow, "companyname", "vendor_type_logic", "Company Name is required when Individual is FALSE");
            }
        }
    });

    // 9. Unique / Duplicate check loop
    Object.keys(rulesConfig).forEach(column => {
        if (!rulesConfig[column].unique) return;
        const schemaRow = schemaRows.find(row => normalizeColumnName(row.column_name) === column);
        if (!schemaRow) return;
        const colIndex = mappedColumns[schemaRow.column_name];
        if (!colIndex) return; // skip if column was not present

        // Group values and detect duplicates
        const valMap = {}; // value -> array of excelRow numbers
        dataRows.forEach(rowObj => {
            const val = String(rowObj.data[column]).trim();
            if (val !== "") {
                if (!valMap[val]) valMap[val] = [];
                valMap[val].push(rowObj.excelRow);
            }
        });

        Object.keys(valMap).forEach(val => {
            if (valMap[val].length > 1) {
                valMap[val].forEach(excelRow => {
                    addRowIssue(excelRow, column, "duplicate", `Duplicate ${column}`);
                });
            }
        });
    });

    // 10. Consolidate detailed findings for GUI display
    const rowIssuesDetailed = [];
    Object.keys(rowIssues).sort((a, b) => parseInt(a, 10) - parseInt(b, 10)).forEach(rowNum => {
        const issues = rowIssues[rowNum];
        const affectedCols = [...new Set(issues.map(i => i.column))];
        const messages = issues.map(i => i.message);
        
        rowIssuesDetailed.push({
            row: parseInt(rowNum, 10),
            issueCount: issues.length,
            columnsAffected: affectedCols.join(", "),
            findings: messages.join(" • ")
        });
    });

    return {
        rowsChecked: dataRows.length,
        rowsWithIssuesCount: Object.keys(rowIssues).length,
        coreIssues,
        rowIssues, // raw issues dictionary
        rowIssuesDetailed,
        ruleReference,
        mappedColumns
    };
}

// ==========================================================================
// EXCEL GENERATOR & CELL HIGHLIGHTER
// ==========================================================================
async function generateReportBlob(originalWorksheet, schemaRows, validationResults, templateType) {
    const reportWorkbook = new ExcelJS.Workbook();
    
    // 1. Core Issues Tab
    const coreSheet = reportWorkbook.addWorksheet("Core Issues");
    coreSheet.columns = [
        { header: "Core Issue", key: "issue", width: 45 },
        { header: "Suggested Fix", key: "fix", width: 45 }
    ];
    validationResults.coreIssues.forEach(ci => {
        coreSheet.addRow(ci);
    });

    // 2. Row Issues Tab
    const rowIssuesSheet = reportWorkbook.addWorksheet("Row Issues");
    rowIssuesSheet.columns = [
        { header: "Row", key: "row", width: 10 },
        { header: "Issue Count", key: "issueCount", width: 15 },
        { header: "Columns Affected", key: "columnsAffected", width: 30 },
        { header: "Validation Findings", key: "findings", width: 60 }
    ];
    validationResults.rowIssuesDetailed.forEach(ri => {
        rowIssuesSheet.addRow(ri);
    });

    // 3. Rule Reference Tab
    const ruleRefSheet = reportWorkbook.addWorksheet("Rule Reference");
    ruleRefSheet.columns = [
        { header: "Rule", key: "rule", width: 25 },
        { header: "Description", key: "desc", width: 60 }
    ];
    Object.keys(validationResults.ruleReference).forEach(rule => {
        ruleRefSheet.addRow({
            rule: rule,
            desc: validationResults.ruleReference[rule]
        });
    });

    // 4. Highlighted Data Tab
    const highlightedSheet = reportWorkbook.addWorksheet("Highlighted Data");

    // Copy entire original sheets row-by-row
    originalWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const vals = [];
        for (let c = 1; c <= originalWorksheet.columnCount; c++) {
            vals.push(row.getCell(c).value);
        }
        highlightedSheet.addRow(vals);
    });

    // Overwrite headers of mapped columns with exact expected schema headers
    const headerRow = highlightedSheet.getRow(1);
    schemaRows.forEach(schemaRow => {
        const colIndex = validationResults.mappedColumns[schemaRow.column_name];
        if (colIndex) {
            headerRow.getCell(colIndex).value = schemaRow.column_name;
        }
    });

    const originalHeaders = [];
    for (let c = 1; c <= highlightedSheet.columnCount; c++) {
        originalHeaders.push(cleanValue(headerRow.getCell(c).value));
    }

    // Highlight Style fills
    const redFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFC7CE' } // openpyxl standard red fill
    };

    const lightRedFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFF2F2' } // openpyxl light row highlight
    };

    const highlightedRows = new Set();
    const invalidCells = [];

    // Parse invalid cells list from raw validation outputs
    Object.keys(validationResults.rowIssues).forEach(rowNum => {
        const issues = validationResults.rowIssues[rowNum];
        issues.forEach(issue => {
            invalidCells.push({
                rowNum: parseInt(rowNum, 10),
                column: issue.column
            });
        });
    });

    // Highlight Invalid Cells positionally based on mapped columns
    invalidCells.forEach(cellObj => {
        const rowNum = cellObj.rowNum;
        const normCol = cellObj.column;
        const schemaRow = schemaRows.find(row => normalizeColumnName(row.column_name) === normCol);

        if (schemaRow) {
            const colIndex = validationResults.mappedColumns[schemaRow.column_name];
            if (colIndex && colIndex <= originalHeaders.length) {
                const excelRow = highlightedSheet.getRow(rowNum);
                const cell = excelRow.getCell(colIndex);
                cell.fill = redFill;
                highlightedRows.add(rowNum);
            }
        }
    });

    // Highlight entire invalid rows in light red
    highlightedRows.forEach(rowNum => {
        const excelRow = highlightedSheet.getRow(rowNum);
        for (let c = 1; c <= originalHeaders.length; c++) {
            const cell = excelRow.getCell(c);
            // Apply light red to non-deep red styled cells
            if (!cell.fill || cell.fill.fgColor?.argb !== 'FFFFC7CE') {
                cell.fill = lightRedFill;
            }
        }
    });

    // Auto Column Width fitting
    const colWidths = [];
    highlightedSheet.eachRow({ includeEmpty: true }, (row) => {
        row.eachCell({ includeEmpty: true }, (cell, colNum) => {
            const val = cell.value;
            const length = val ? String(val).length : 0;
            if (!colWidths[colNum]) colWidths[colNum] = 0;
            if (length > colWidths[colNum]) {
                colWidths[colNum] = length;
            }
        });
    });
    for (let c = 1; c <= originalHeaders.length; c++) {
        const col = highlightedSheet.getColumn(c);
        if (col) {
            col.width = Math.min((colWidths[c] || 0) + 5, 40);
        }
    }

    // Write workbook to Buffer and return Blob
    const buffer = await reportWorkbook.xlsx.writeBuffer();
    return new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

// ==========================================================================
// RESULT TABLE POPULATION
// ==========================================================================
function populateRowIssuesTable(issues) {
    const tbody = document.getElementById("row-issues-body");
    const emptyMsg = document.getElementById("no-row-issues");
    tbody.innerHTML = "";

    if (issues.length === 0) {
        emptyMsg.style.display = "flex";
        document.getElementById("row-issues-table").style.display = "none";
        return;
    }

    emptyMsg.style.display = "none";
    document.getElementById("row-issues-table").style.display = "table";

    issues.forEach(ri => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>Row ${ri.row}</td>
            <td><span class="badge-count">${ri.issueCount}</span></td>
            <td><span class="affected-cols-list">${ri.columnsAffected}</span></td>
            <td><span class="findings-list">${ri.findings}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

function populateCoreIssuesTable(issues) {
    const tbody = document.getElementById("core-issues-body");
    const emptyMsg = document.getElementById("no-core-issues");
    tbody.innerHTML = "";

    if (issues.length === 0) {
        emptyMsg.style.display = "flex";
        document.getElementById("core-issues-table").style.display = "none";
        return;
    }

    emptyMsg.style.display = "none";
    document.getElementById("core-issues-table").style.display = "table";

    issues.forEach(ci => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td><span class="findings-list">${ci.issue}</span></td>
            <td><span class="affected-cols-list">${ci.fix}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

function populateRulesRefTable(rules) {
    const tbody = document.getElementById("rules-ref-body");
    tbody.innerHTML = "";

    Object.keys(rules).forEach(rule => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td><code>${rule}</code></td>
            <td>${rules[rule]}</td>
        `;
        tbody.appendChild(tr);
    });
}
