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

    // ── M Yaseen / M Yassen (both spellings accepted) ─────
    "Yaseen": "M.Yaseen",
    "Yasin": "M.Yaseen",
    "myaseen": "M.Yaseen",
    "myassen": "M.Yaseen",
    "yassen": "M.Yaseen",
    "myasin": "M.Yaseen",

    // ── Muhammad Ali ─────────────────────────────────────
    "ali": "Muhammad Ali",
    "muhammadali": "Muhammad Ali",
    "mali": "Muhammad Ali",

    // ── Akhtar Hussain ───────────────────────────────────
    "hussain": "Akhtar Hussain",
    "husain": "Akhtar Hussain",
    "akhtar": "Akhtar Hussain",
    "akhtarhussain": "Akhtar Hussain",
    "akhtarhusain": "Akhtar Hussain",

    // ── Aman Waheed ──────────────────────────────────────
    "waheed": "Aman Waheed",
    "wahid": "Aman Waheed",
    "aman": "Aman Waheed",
    "amanwaheed": "Aman Waheed",
    "amanwahid": "Aman Waheed",

};

/*
===========================================================
Resolve Official Name
  Converts any alias → official Excel name.
  Returns null when no alias matches (→ Unknown Salesman).
===========================================================
*/

function resolveOfficialName(rawName) {

    if (!rawName) return null;

    // Normalize: lowercase, strip dots and extra spaces
    const key = String(rawName)
        .toLowerCase()
        .replace(/\./g, "")
        .replace(/\s+/g, "")
        .trim();

    const official = NAME_ALIAS_MAP[key] || null;

    if (official) {
        console.log(`[AliasMap] "${rawName}" → "${official}"`);
    } else {
        console.warn(`[AliasMap] No match for "${rawName}" → Unknown Salesman`);
    }

    return official;

}

/*
===========================================================
Normalize Date
16.07.2026
16-07-2026
16/07/2026
===========================================================
*/

function normalizeDate(value) {

    if (!value) return "";

    // Excel Date object
    if (value instanceof Date) {

        const day = String(value.getDate()).padStart(2, "0");
        const month = String(value.getMonth() + 1).padStart(2, "0");
        const year = value.getFullYear();

        return `${day}-${month}-${year}`;
    }

    // String date
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

    if (!value) return "";

    return String(value)
        .toLowerCase()
        .replace(/\./g, "")
        .replace(/\s+/g, "")
        .trim();

}

/*
===========================================================
Find Date Row
Column A
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

            console.log("MATCH FOUND:", rowNumber);

            rowFound = rowNumber;

        }

    });

    return rowFound;

}

/*
===========================================================
Find DSF Column
Header Row = 6
===========================================================
*/

function findDSFColumn(sheet, dsfName) {

    const headerRow = sheet.getRow(6);

    // dsfName is already the official name resolved by resolveOfficialName()
    const search = normalizeText(dsfName);

    let columnFound = -1;

    // Log all header values for debugging
    console.log("[findDSFColumn] Searching for:", dsfName, "→ normalized:", search);

    headerRow.eachCell((cell, columnNumber) => {

        const raw = cell.value ? String(cell.value) : "";
        const value = normalizeText(raw);

        console.log(`  Col ${columnNumber}: raw="${raw}" normalized="${value}"`);

        if (value === search) {

            columnFound = columnNumber;

        }

    });

    // Fuzzy fallback: if not found by exact normalized match,
    // try matching just the core name part (strip leading initials like M)
    if (columnFound === -1) {

        console.warn("[findDSFColumn] Exact match failed, trying fuzzy match...");

        // e.g. "myassen" → strip leading single char "m" → "yassen"
        const searchCore = search.replace(/^[a-z]/, "");

        headerRow.eachCell((cell, columnNumber) => {

            const value = normalizeText(cell.value ? String(cell.value) : "");
            const valueCore = value.replace(/^[a-z]/, "");

            if (valueCore === searchCore && searchCore.length > 3) {

                console.log(`[findDSFColumn] Fuzzy match at col ${columnNumber}: "${value}" ~ "${search}"`);
                columnFound = columnNumber;

            }

        });

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

    console.log("Workbook Loaded");

    return workbook;

}

/*
===========================================================
Save Workbook
===========================================================
*/
function recalculateWorkbook(sheet) {

    console.log("===== RECALCULATING WORKBOOK =====");

    // STEP 1 - Update Daily Total (Column F)

    for (let row = 7; row <= 37; row++) {

        let total = 0;

        // B to E
        for (let col = 2; col <= 5; col++) {

            const rawValue = sheet.getRow(row).getCell(col).value;

            let value = 0;

            if (typeof rawValue === "number") {

                value = rawValue;

            }
            else if (
                rawValue &&
                typeof rawValue === "object" &&
                rawValue.result != null
            ) {

                value = Number(rawValue.result) || 0;

            }
            else {

                value = Number(rawValue) || 0;

            }

            total += value;

        }

        sheet.getRow(row).getCell(6).value = total;

        console.log(
            `Row ${row} Total = ${total}`
        );

    }
    // STEP 2  ← YE LOOP KE BAHAR HOGA

    let totalB = 0;

    for (let row = 7; row <= 37; row++) {

        const value = Number(
            sheet.getRow(row).getCell(2).value || 0
        );

        totalB += value;
    }

    sheet.getCell("B3").value = totalB;

    console.log("B3 Updated:", totalB);

    // STEP 3 - Update C3

    let totalC = 0;

    for (let row = 7; row <= 37; row++) {

        const value = Number(
            sheet.getRow(row).getCell(3).value || 0
        );

        totalC += value;
    }

    sheet.getCell("C3").value = totalC;

    console.log("C3 Updated:", totalC);

    // D3
    let totalD = 0;

    for (let row = 7; row <= 37; row++) {
        totalD += Number(sheet.getRow(row).getCell(4).value || 0);
    }

    sheet.getCell("D3").value = totalD;

    console.log("D3 Updated:", totalD);

    // E3
    let totalE = 0;

    for (let row = 7; row <= 37; row++) {
        totalE += Number(sheet.getRow(row).getCell(5).value || 0);
    }

    sheet.getCell("E3").value = totalE;

    console.log("E3 Updated:", totalE);

    // STEP 2.5 - Update F3 (Grand Total)

    let totalF = 0;

    for (let row = 7; row <= 37; row++) {

        totalF += Number(
            sheet.getRow(row).getCell(6).value || 0
        );

    }

    sheet.getCell("F3").value = totalF;

    console.log("F3 Updated:", totalF);

    console.log("F2 Target =", sheet.getCell("F2").value);
    console.log("F3 Achieved =", sheet.getCell("F3").value);

    // STEP 3 - Update Percentage Row (Row 4)

    for (let col = 2; col <= 6; col++) {

        const targetCell = sheet.getRow(2).getCell(col).value;

        const target =
            targetCell &&
                typeof targetCell === "object" &&
                targetCell.result != null
                ? Number(targetCell.result)
                : Number(targetCell || 0);

        const achievedCell = sheet.getRow(3).getCell(col).value;

        const achieved =
            achievedCell &&
                typeof achievedCell === "object" &&
                achievedCell.result != null
                ? Number(achievedCell.result)
                : Number(achievedCell || 0);

        let percentage = 0;

        if (target > 0) {
            percentage = achieved / target;
        }

        sheet.getRow(4).getCell(col).value = percentage;

        console.log(
            `Column ${col} Percentage = ${percentage}`
        );

    }
}
async function saveWorkbook(workbook, originalPath) {

    const outputPath = path.join(
        path.dirname(originalPath),
        "saved_test.xlsx"
    );

    await workbook.xlsx.writeFile(outputPath);

    console.log("Workbook Saved");

    return outputPath;

}
/*
===========================================================
Update Workbook
===========================================================
*/

async function updateWorkbook(filePath, parsedData) {

    const workbook = await loadWorkbook(filePath);
    console.log("\n===== WORKSHEETS =====");

    workbook.worksheets.forEach((ws, index) => {
        console.log(index + 1, ":", ws.name);
    });

    console.log("======================\n");

    const sheet = workbook.getWorksheet("Sheet1");

    if (!sheet) {
        throw new Error("Sheet1 not found.");
    }

    /*
    ---------------------------------------
    Find Date Row
    ---------------------------------------
    */
    const rowNumber = findDateRow(
        sheet,
        parsedData.date
    );

    if (rowNumber === -1) {
        throw new Error(
            `Date not found : ${parsedData.date}`
        );
    }

    console.log("Date Row :", rowNumber);
    /*
    ---------------------------------------
    Resolve Alias → Official Name
    ---------------------------------------
    */

    const officialName = resolveOfficialName(parsedData.dsf);

    if (!officialName) {
        throw new Error(
            `Unknown Salesman: "${parsedData.dsf}" did not match any known alias. Excel was NOT updated.`
        );
    }

    /*
    ---------------------------------------
    Find DSF Column
    ---------------------------------------
    */

    const columnNumber = findDSFColumn(
        sheet,
        officialName
    );

    if (columnNumber === -1) {

        throw new Error(
            `DSF column not found in Excel for official name: "${officialName}"`
        );

    }

    console.log("DSF Column :", columnNumber);

    /*
    ---------------------------------------
    Update Target Cell
    ---------------------------------------
    */

    const cell = sheet.getRow(rowNumber).getCell(columnNumber);

    console.log("Target Cell :", cell.address);

    console.log("Old Value :", cell.value);

    cell.value = Number(
        String(parsedData.todayValue)
            .replace(/,/g, "")
    );
    recalculateWorkbook(sheet);

    console.log("New Value :", cell.value);

    /*
    ---------------------------------------
    Save Workbook
    ---------------------------------------
    */

    const savedFile = await saveWorkbook(
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

    const result = await updateWorkbook(
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