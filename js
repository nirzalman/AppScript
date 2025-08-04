// =================================================================
// GLOBAL CONFIGURATION
// הגדרות גלובליות
// =================================================================
const SS = SpreadsheetApp.getActiveSpreadsheet();
const DATA_SHEET = SS.getSheetByName("Data");
const ROLES_SHEET = SS.getSheetByName("Roles");
const CONFIG_SHEET = SS.getSheetByName("Config");
const LOG_SHEET = SS.getSheetByName("HttpRequestLogs");

// הגדרת עמודות בגיליון הנתונים
const COL = {
    ID: 1,
    OWNER: 2,
    BRAND: 3,
    DOMAIN: 4,
    MAIN_CATEGORY: 5,
    SUB_CATEGORY: 6,
    CUSTOM_TEXT_1: 7,
    CUSTOM_TEXT_2: 8,
    MULTI_LEVEL_CATEGORY: 9,
    MULTI_LEVEL_TAG: 10,
    CREATED_AT: 11
};

/**
 * Sets up the required headers in the sheets if they are empty.
 * This makes the script more robust for first-time use.
 */
function setupSheetHeaders() {
    // Setup for Data Sheet
    if (DATA_SHEET.getLastRow() === 0) {
        const dataHeaders = [
            'ID', 'Owner', 'Brand', 'Domain', 'Main Category Name',
            'Sub-Category Name', 'Custom Text 1', 'Custom Text 2',
            'Multi-Level Category', 'Multi-Level Tag', 'Created At'
        ];
        DATA_SHEET.appendRow(dataHeaders);
    }

    // Setup for Roles Sheet
    if (ROLES_SHEET.getLastRow() === 0) {
        const rolesHeaders = ['Email', 'Role'];
        ROLES_SHEET.appendRow(rolesHeaders);
        // Add the current user as the first Admin as an example
        ROLES_SHEET.appendRow([Session.getEffectiveUser().getEmail(), 'Admin']);
    }
}


// =================================================================
// SERVING THE WEB APP
// הגשת האפליקציה
// =================================================================
function doGet(e) {
  try {
    setupSheetHeaders(); // Run setup on each load to ensure sheets are ready
    return HtmlService.createTemplateFromFile('WebApp')
        .evaluate()
        .setTitle('Zendesk Field Management')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
    return HtmlService.createHtmlOutput('<h1>An error occurred. Please contact support.</h1><p>' + error.toString() + '</p>');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// =================================================================
// USER & PERMISSIONS
// ניהול משתמשים והרשאות
// =================================================================
function getUserData() {
    const email = Session.getActiveUser().getEmail();
    const rolesData = ROLES_SHEET.getDataRange().getValues();
    let userRole = 'User'; // Default role for any user

    // Find the user in the Roles sheet to check for elevated permissions
    for (let i = 1; i < rolesData.length; i++) {
        if (rolesData[i][0] && typeof rolesData[i][0] === 'string') {
            if (rolesData[i][0].toLowerCase().trim() === email.toLowerCase().trim()) {
                userRole = (rolesData[i][1] && typeof rolesData[i][1] === 'string') ? rolesData[i][1].trim() : 'User';
                break;
            }
        }
    }

    return {
        email: email,
        role: userRole
    };
}


// =================================================================
// DATA FETCHING
// שליפת נתונים
// =================================================================
function getDashboardData(brandFilter) {
    // 1. Get User Info
    const user = getUserData();
    const userEmail = <URL>.toLowerCase().trim();
    const isAdmin = user.role.toLowerCase() === 'admin';

    // 2. Get All Data from Sheet
    const dataRange = DATA_SHEET.getDataRange();
    const allValues = dataRange.getValues();
    const headers = allValues.shift(); // Get and remove header row

    // 3. Filter Rows based on Permissions
    let permittedRows = [];
    if (isAdmin) {
        // Admin sees everything
        permittedRows = allValues;
    } else {
        // User sees only their own records
        for (let i = 0; i < allValues.length; i++) {
            const row = allValues[i];
            const ownerEmailInSheet = row[COL.OWNER - 1]; // Get email from the 'Owner' column

            // Robust check: ensure ownerEmailInSheet is a non-empty string before comparing
            if (ownerEmailInSheet && typeof ownerEmailInSheet === 'string') {
                if (ownerEmailInSheet.toLowerCase().trim() === userEmail) {
                    permittedRows.push(row);
                }
            }
        }
    }

    // 4. Map rows to objects
    let records = <URL>(row => {
        const record = {};
        headers.forEach((header, i) => {
            record[header] = row[i];
        });
        return record;
    }).filter(record => <URL>); // Ensure it's a valid record with an ID

    // 5. Filter by Brand
    if (brandFilter && brandFilter !== 'All') {
        records = records.filter(record => record.Brand === brandFilter);
    }

    // 6. Get unique values for comboboxes (from all data, not just filtered)
    const allRecordsForCombobox = DATA_SHEET.getDataRange().getValues();
    allRecordsForCombobox.shift();
    const uniqueMainCategories = [...new Set(<URL>(r => r[COL.MAIN_CATEGORY - 1]))].filter(Boolean);
    const uniqueSubCategories = [...new Set(<URL>(r => r[COL.SUB_CATEGORY - 1]))].filter(Boolean);

    // 7. Return final data structure
    return {
        user: user,
        records: records,
        uniqueValues: {
            mainCategories: uniqueMainCategories,
            subCategories: uniqueSubCategories
        }
    };
}


// =================================================================
// CRUD OPERATIONS
// פעולות על נתונים
// =================================================================
function addOrUpdateRecord(recordObject) {
    try {
        const user = getUserData();

        // חישוב השדות המחושבים
        const multiLevelCategory = `${recordObject.domain}:: ${recordObject.main_category}:: ${recordObject.sub_category}`;
        const rawTag = `${recordObject.brand}_${recordObject.domain}______${recordObject.main_category}_____${recordObject.sub_category}`;
        const multiLevelTag = rawTag.toLowerCase().replace(/[\(\)\[\]\{\}\'\?\,]/g, '').replace(/[\s&:]/g, '_');

        const id = new Date().getTime();
        const createdAt = new Date();

        const newRow = [
            id,
            <URL>, // Always use the currently logged-in user's email
            recordObject.brand,
            recordObject.domain,
            recordObject.main_category,
            recordObject.sub_category,
            recordObject.custom_text_1,
            recordObject.custom_text_2,
            multiLevelCategory,
            multiLevelTag,
            createdAt
        ];

        DATA_SHEET.appendRow(newRow);
        // **CRITICAL FIX**: Force the spreadsheet to save all pending changes immediately.
        // This prevents a race condition where the client re-fetches data before the new row is saved.
        SpreadsheetApp.flush();

        syncToExternalSheets('ADD', newRow);

        // The client will re-fetch the data, so we don't need to return the new record object.
        return { success: true, message: 'הרשומה נוספה בהצלחה' };
    } catch (e) {
        Logger.log(e); // Log the full error for debugging
        return { success: false, message: 'אירעה שגיאה בעת הוספת הרשומה: ' + e.message };
    }
}

function deleteRecord(recordId) {
    try {
        const user = getUserData();
        const data = DATA_SHEET.getDataRange().getValues();

        for (let i = 1; i < data.length; i++) {
            if (data[i][<URL> - 1] == recordId) {
                const recordOwner = data[i][COL.OWNER - 1];
                // בדיקת הרשאות
                if (user.role.toLowerCase() !== 'admin' && <URL> !== recordOwner) {
                    return { success: false, message: 'אין לך הרשאה למחוק רשומה זו.' };
                }

                const rowToDelete = data[i];
                DATA_SHEET.deleteRow(i + 1);
                SpreadsheetApp.flush(); // Ensure deletion is committed
                syncToExternalSheets('DELETE', rowToDelete);

                return { success: true, message: 'הרשומה נמחקה בהצלחה.' };
            }
        }
        return { success: false, message: 'הרשומה לא נמצאה.' };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

/**
 * Deletes all records owned by the current active user.
 * @returns {object} A result object with success status and a message.
 */
function deleteAllUserRecords() {
    try {
        const user = getUserData();
        const userEmail = <URL>.toLowerCase().trim();
        if (!userEmail) {
            return { success: false, message: 'לא ניתן לזהות את המשתמש.' };
        }

        const data = DATA_SHEET.getDataRange().getValues();
        let deletedCount = 0;

        // Iterate backwards to avoid index shifting issues on deletion
        for (let i = data.length - 1; i >= 1; i--) {
            const ownerEmail = data[i][COL.OWNER - 1];
            if (ownerEmail && typeof ownerEmail === 'string' && ownerEmail.toLowerCase().trim() === userEmail) {
                const rowToDelete = data[i];
                DATA_SHEET.deleteRow(i + 1);
                syncToExternalSheets('DELETE', rowToDelete);
                deletedCount++;
            }
        }

        if (deletedCount > 0) {
            SpreadsheetApp.flush(); // Ensure all deletions are committed
            return { success: true, message: `${deletedCount} רשומות נמחקו בהצלחה.` };
        } else {
            return { success: false, message: 'לא נמצאו רשומות למחיקה עבור המשתמש שלך.' };
        }

    } catch (e) {
        return { success: false, message: 'אירעה שגיאה: ' + e.message };
    }
}


// =================================================================
// ADMIN MODULES
// מודולים לאדמין
// =================================================================
function getAdminSettings() {
    const user = getUserData();
    if (user.role.toLowerCase() !== 'admin') return null;

    const configData = CONFIG_SHEET.getRange("A1:B2").getValues();
    const templatesData = LOG_SHEET.getSheetValues(1, 1, LOG_SHEET.getLastRow(), 5); // דוגמה

    return {
        sheets: {
            sheet1: configData[0][1],
            sheet2: configData[1][1]
        },
        templates: templatesData.slice(1) // דוגמה
    };
}

function saveSheetsUrls(urls) {
    const user = getUserData();
    if (user.role.toLowerCase() !== 'admin') return { success: false, message: 'אין הרשאה' };

    CONFIG_SHEET.getRange("B1").setValue(urls.sheet1);
    CONFIG_SHEET.getRange("B2").setValue(urls.sheet2);
    return { success: true, message: 'הגדרות נשמרו' };
}

function sendHttpRequest(templateName, recordId) {
    // לוגיקה לשליפת תבנית, החלפת משתנים ושליחת הבקשה
    // ...
    // לאחר השליחה, כתיבת לוג
    LOG_SHEET.appendRow([new Date(), templateName, 'POST', 'https://...', 200, 'OK']);
    return { success: true, message: 'הבקשה נשלחה' };
}


// =================================================================
// EXTERNAL SYNC
// סנכרון חיצוני
// =================================================================
function syncToExternalSheets(action, rowData) {
    const configData = CONFIG_SHEET.getRange("A1:B2").getValues();
    const url1 = configData[0][1];
    const url2 = configData[1][1];

    if (url1) {
        try {
            const sheet1 = SpreadsheetApp.openByUrl(url1).getSheets()[0];
            handleSync(sheet1, action, rowData);
        } catch(e) {
            Logger.log(`Error syncing to sheet 1: ${e.message}`);
        }
    }
    if (url2) {
        try {
            const sheet2 = SpreadsheetApp.openByUrl(url2).getSheets()[0];
            handleSync(sheet2, action, rowData);
        } catch(e) {
            Logger.log(`Error syncing to sheet 2: ${e.message}`);
        }
    }
}

function handleSync(sheet, action, rowData) {
    if (action === 'ADD') {
        sheet.appendRow(rowData);
    } else if (action === 'DELETE') {
        const recordId = rowData[<URL> - 1];
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][<URL> - 1] == recordId) {
                sheet.deleteRow(i + 1);
                return;
            }
        }
    }
}
