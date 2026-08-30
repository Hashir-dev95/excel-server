console.log("***** USING CURRENT excel.service.js *****");
const ExcelJS = require("exceljs");
const path = require("path");

/*
===========================================================
NAME ALIAS MAP
  key   → normalized alias  (lowercase, no dots, no spaces)
  value → official Excel column name (exact match in header row)
===========================================================
*/

const NAME_ALIAS_MAP = {

   // ── M.Yaseen aliases ─────────────────────────────
    "myaseen": "M.Yaseen",
    "myassen": "M.Yaseen",
    "myasin": "M.Yaseen",
    "yaseen": "M.Yaseen",
    "yassen": "M.Yaseen",
    "yasin": "M.Yaseen",
    // ── Muhammad Ali ─────────────────────────────────────
    "Ali": "Muhammad Ali",
    "muhammadali": "Muhammad Ali",
    "mali": "Muhammad Ali",

    // ── Mujeebullah ───────────────────────────────────
    "Mujeebullah": "Mujeebullah",
    "mujeebullah": "Mujeebullah",
    "Mujeeb": "Mujeebullah",
    "Mujeeb": "Mujeebullah",

 // ── Akhtar / Akhter ─────────────────────────────
"akhtar": "AKHTAR",
"akhter": "AKHTAR",


    // ── Aman Wahid ──────────────────────────────────────
    "aman": "Aman Wahid",
    "wahid": "Aman Wahid",
    "waheed": "Aman Wahid",
    "amanwahid": "Aman Wahid",
    "amanwaheed": "Aman Wahid",
    "awahid": "Aman Wahid",
    "awaheed": "Aman Wahid",
};

/*
===========================================================
console.log("***** USING CURRENT excel.service.js *****");

const ExcelJS = require("exceljs");
const path = require("path");

/*
===========================================================
NAME ALIAS MAP

key   → normalized alias
        lowercase, no dots, no spaces

value → official Excel column name
        must match the Excel header
===========================================================
*/

const NAME_ALIAS_MAP = {

    // ── M.Yaseen / M.Yassen ─────────────────────────────
    "myaseen": "M.Yaseen",
    "myassen": "M.Yaseen",
    "myasin": "M.Yaseen",
    "yaseen": "M.Yaseen",
    "yassen": "M.Yaseen",
    "yasin": "M.Yaseen",

    // ── Muhammad Ali ─────────────────────────────────────
    "ali": "Muhammad Ali",
    "muhammadali": "Muhammad Ali",
    "mali": "Muhammad Ali",

    // ── Mujeebullah ──────────────────────────────────────
    "mujeebullah": "Mujeebullah",
    "mujeeb": "Mujeebullah",

    // ── Akhtar / Akhter ──────────────────────────────────
    "akhtar": "AKHTAR",
    "akhter": "AKHTAR",

    // ── Aman Wahid ───────────────────────────────────────
    "aman": "Aman Wahid",
    "wahid": "Aman Wahid",
    "waheed": "Aman Wahid",
    "amanwahid": "Aman Wahid",
    "amanwaheed": "Aman Wahid",
    "awahid": "Aman Wahid",
    "awaheed": "Aman Wahid",

};

/*
===========================================================
Resolve Official Name

Converts any alias → official Excel name.

Returns null when no alias matches.
===========================================================
*/

function resolveOfficialName(rawName) {

    if (!rawName) {
        return null;
    }

    const key = String(rawName)
        .toLowerCase()
        .replace(/\./g, "")
        .replace(/\s+/g, "")
        .trim();

    const official = NAME_ALIAS_MAP[key] || null;

    if (official) {

        console.log(
            `[AliasMap] "${rawName}" → "${official}"`
        );

    } else {

        console.warn(
            `[AliasMap] No match for "${rawName}" → Unknown Salesman`
        );

    }

    return official;

}

/*
===========================================================
Normalize Date

Supported:

16.07.2026
16-07-2026
16/07/2026
Excel Date object
===========================================================
*/

function normalizeDate(value) {

    if (!value) {
        return "";
    }

    // Excel Date object
    if (value instanceof Date) {

        const day = String(
            value.getDate()
        ).padStart(2, "0");

        const month = String(
            value.getMonth() + 1
        ).padStart(2, "0");

        const year = value.getFullYear();

        return `${day}-${month}-${year}`;
    }

    let text = String(value).trim();

    text = text.replace(/[./]/g, "-");

    const parts = text.split("-");

    if (parts.length === 3) {

        let [day, month, year] = parts;

        day = day.padStart(2, "0");
        month = month.padStart(2, "0");

        if (year.length === 2) {
            year = "20" + year;
        }

        return `${day}-${month}-${year}`;
    }

    return text;

}

/*
===========================================================
Normalize Text
===========================================================
*/

function normalizeText(value) {

    if (!value) {
        return "";
    }

    return String(value)
        .toLowerCase()
        .replace(/\./g, "")
        .replace(/\s+/g, "")
        .trim();

}

/*
===========================================================
Find Date Row

Date is stored in Column A.
===========================================================
*/

function findDateRow(sheet, targetDate) {

    console.log("Searching Date:", targetDate);

    const searchDate = normalizeDate(targetDate);

    let rowFound = -1;

    sheet.eachRow((row, rowNumber) => {

        const rawValue = row.getCell(1).value;

        const excelDate = normalizeDate(rawValue);

        console.log(
            "Row:",
            rowNumber,
            "| Raw:",
            rawValue,
            "| Normalized:",
            excelDate
        );

        if (excelDate === searchDate) {

            console.log(
                "MATCH FOUND:",
                rowNumber
            );

            rowFound = rowNumber;

        }

    });

    return rowFound;

}

/*
===========================================================
Find DSF Column

Header Row = 6

Salesman columns:
B = 2
C = 3
D = 4
E = 5
F = 6

Total:
G = 7

We intentionally search only B:F so that
the Total column can never be selected as a salesman.
===========================================================
*/

function findDSFColumn(sheet, dsfName) {

    const headerRow = sheet.getRow(6);

    const search = normalizeText(dsfName);

    let columnFound = -1;

    console.log(
        "[findDSFColumn] Searching for:",
        dsfName,
        "→ normalized:",
        search
    );

    // Search only salesman columns B:F
    for (let columnNumber = 2; columnNumber <= 6; columnNumber++) {

        const cell = headerRow.getCell(columnNumber);

        const raw = cell.value
            ? String(cell.value)
            : "";

        const value = normalizeText(raw);

        console.log(
            `  Col ${columnNumber}: raw="${raw}" normalized="${value}"`
        );

        if (value === search) {

            columnFound = columnNumber;

            break;

        }

    }

    /*
    -------------------------------------------------------
    Fuzzy fallback

    Example:

    M.Yaseen
    myaseen
    yaseen

    This is only used if exact matching fails.
    -------------------------------------------------------
    */

    if (columnFound === -1) {

        console.warn(
            "[findDSFColumn] Exact match failed, trying fuzzy match..."
        );

        const searchCore = search.replace(
            /^[a-z]/,
            ""
        );

        for (
            let columnNumber = 2;
            columnNumber <= 6;
            columnNumber++
        ) {

            const cell = headerRow.getCell(
                columnNumber
            );

            const value = normalizeText(
                cell.value
                    ? String(cell.value)
                    : ""
            );

            const valueCore = value.replace(
                /^[a-z]/,
                ""
            );

            if (
                valueCore === searchCore &&
                searchCore.length > 3
            ) {

                console.log(
                    `[findDSFColumn] Fuzzy match at col ${columnNumber}: "${value}" ~ "${search}"`
                );

                columnFound = columnNumber;

                break;

            }

        }

    }

    return columnFound;

}

/*
===========================================================
Load Workbook
===========================================================
*/

async function loadWorkbook(filePath) {

    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.readFile(filePath);

    console.log(
        "Workbook Loaded:",
        filePath
    );

    return workbook;

}

/*
===========================================================
Get Numeric Cell Value
===========================================================
*/

function getNumericValue(cell) {

    const rawValue = cell.value;

    if (typeof rawValue === "number") {

        return rawValue;

    }

    if (
        rawValue &&
        typeof rawValue === "object" &&
        rawValue.result != null
    ) {

        return Number(
            rawValue.result
        ) || 0;

    }

    return Number(
        String(rawValue || "")
            .replace(/,/g, "")
    ) || 0;

}

/*
===========================================================
Recalculate Workbook

Structure:

B:F = 5 Salesmen
G   = Total

Rows:

2 = Target
3 = Achieved
4 = Percentage
6 = Headers
7:37 = Daily data
===========================================================
*/

function recalculateWorkbook(sheet) {

    console.log(
        "===== RECALCULATING WORKBOOK ====="
    );

    /*
    =======================================================
    STEP 1
    Daily Total

    B:F = 5 salesman columns
    G   = Total
    =======================================================
    */

    for (
        let row = 7;
        row <= 37;
        row++
    ) {

        let total = 0;

        // B:F = 5 salesman columns
        for (
            let col = 2;
            col <= 6;
            col++
        ) {

            const cell = sheet
                .getRow(row)
                .getCell(col);

            const value = getNumericValue(cell);

            total += value;

        }

        // G = Total
        sheet
            .getRow(row)
            .getCell(7)
            .value = total;

        console.log(
            `Row ${row} Total = ${total}`
        );

    }

    /*
    =======================================================
    STEP 2
    Salesman Achieved Totals

    B3:F3
    =======================================================
    */

    for (
        let col = 2;
        col <= 6;
        col++
    ) {

        let total = 0;

        for (
            let row = 7;
            row <= 37;
            row++
        ) {

            total += getNumericValue(
                sheet
                    .getRow(row)
                    .getCell(col)
            );

        }

        const columnLetter =
            sheet.getColumn(col).letter;

        sheet
            .getRow(3)
            .getCell(col)
            .value = total;

        console.log(
            `${columnLetter}3 Updated:`,
            total
        );

    }

    /*
    =======================================================
    STEP 3
    Grand Total

    G3 = sum of daily totals
    =======================================================
    */

    let grandTotal = 0;

    for (
        let row = 7;
        row <= 37;
        row++
    ) {

        grandTotal += getNumericValue(
            sheet
                .getRow(row)
                .getCell(7)
        );

    }

    sheet
        .getRow(3)
        .getCell(7)
        .value = grandTotal;

    console.log(
        "G3 Updated:",
        grandTotal
    );

    /*
    =======================================================
    STEP 4
    Targets / Percentages

    B:G
    =======================================================
    */

    for (
        let col = 2;
        col <= 7;
        col++
    ) {

        const target = getNumericValue(
            sheet
                .getRow(2)
                .getCell(col)
        );

        const achieved = getNumericValue(
            sheet
                .getRow(3)
                .getCell(col)
        );

        let percentage = 0;

        if (target > 0) {

            percentage =
                achieved / target;

        }

        sheet
            .getRow(4)
            .getCell(col)
            .value = percentage;

        const columnLetter =
            sheet.getColumn(col).letter;

        console.log(
            `${columnLetter}4 Percentage = ${percentage}`
        );

    }

    console.log(
        "===== RECALCULATION COMPLETE ====="
    );

}

/*
===========================================================
Save Workbook
===========================================================
*/

async function saveWorkbook(
    workbook,
    originalPath
) {

    const outputPath = path.join(
        path.dirname(originalPath),
        "saved_test.xlsx"
    );

    await workbook.xlsx.writeFile(
        outputPath
    );

    console.log(
        "Workbook Saved:",
        outputPath
    );

    return outputPath;

}

/*
===========================================================
Update Workbook
===========================================================
*/

async function updateWorkbook(
    filePath,
    parsedData
) {

    const workbook =
        await loadWorkbook(filePath);

    console.log(
        "\n===== WORKSHEETS ====="
    );

    workbook.worksheets.forEach(
        (ws, index) => {

            console.log(
                index + 1,
                ":",
                ws.name
            );

        }
    );

    console.log(
        "======================\n"
    );

    /*
    =======================================================
    Sheet
    =======================================================
    */

    const sheet =
        workbook.getWorksheet("Sheet1");

    if (!sheet) {

        throw new Error(
            "Sheet1 not found."
        );

    }

    /*
    =======================================================
    Find Date
    =======================================================
    */

    const rowNumber =
        findDateRow(
            sheet,
            parsedData.date
        );

    if (rowNumber === -1) {

        throw new Error(
            `Date not found : ${parsedData.date}`
        );

    }

    console.log(
        "Date Row :",
        rowNumber
    );

    /*
    =======================================================
    Resolve Alias
    =======================================================
    */

    const officialName =
        resolveOfficialName(
            parsedData.dsf
        );

    if (!officialName) {

        throw new Error(
            `Unknown Salesman: "${parsedData.dsf}" did not match any known alias. Excel was NOT updated.`
        );

    }

    /*
    =======================================================
    Find DSF Column
    =======================================================
    */

    const columnNumber =
        findDSFColumn(
            sheet,
            officialName
        );

    if (columnNumber === -1) {

        throw new Error(
            `DSF column not found in Excel for official name: "${officialName}"`
        );

    }

    console.log(
        "DSF Column :",
        columnNumber
    );

    /*
    =======================================================
    Update Target Cell
    =======================================================
    */

    const cell =
        sheet
            .getRow(rowNumber)
            .getCell(columnNumber);

    console.log(
        "Target Cell:",
        cell.address
    );

    console.log(
        "Old Value:",
        cell.value
    );

    const numericValue =
        Number(
            String(parsedData.todayValue)
                .replace(/,/g, "")
        );

    if (!Number.isFinite(numericValue)) {

        throw new Error(
            `Invalid sales value: "${parsedData.todayValue}"`
        );

    }

    cell.value = numericValue;

    console.log(
        "New Value:",
        cell.value
    );

    /*
    =======================================================
    Recalculate
    =======================================================
    */

    recalculateWorkbook(sheet);

    /*
    =======================================================
    Save
    =======================================================
    */

    const savedFile =
        await saveWorkbook(
            workbook,
            filePath
        );

    return {

        success: true,

        output: savedFile,

        sheet: sheet.name,

        row: rowNumber,

        column: columnNumber,

        cell: cell.address,

        value: cell.value

    };

}

/*
===========================================================
Test Workbook
===========================================================
*/

async function testWorkbook(filePath) {

    const parsedData = {

        date: "16-07-2026",

        dsf: "M.Yassen",

        todayValue: 80480

    };

    const result =
        await updateWorkbook(
            filePath,
            parsedData
        );

    console.log(result);

    return result;

}

/*
===========================================================
Exports
===========================================================
*/

module.exports = {

    testWorkbook,

    updateWorkbook,

    loadWorkbook,

    saveWorkbook,

    findDateRow,

    findDSFColumn

};
