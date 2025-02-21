import * as xlsx from "xlsx";
import fs from "fs";
import path from "path";

export async function processExcel(fileBuffer) {
  try {
    // Load JSON files correctly
    const objectCodes = await fetchJson("objectCodes.json");
    const objectCodeCategories = await fetchJson("objectCodeCategories.json");

    // Read the uploaded Excel file from buffer
    const workbook = xlsx.read(fileBuffer, { type: "buffer" });

    if (workbook.SheetNames.length === 0) {
      throw new Error("Excel file is empty.");
    }

    const sheetName = workbook.SheetNames[0];
    const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], {
      raw: false,
    });

    if (sheetData.length === 0) {
      throw new Error("No data found in Excel sheet.");
    }

    // Define the column mapping manually
    const mappedData = sheetData.map((row) => {
      const objectCode = row["Object"]
        ? row["Object"].toString().replace(/'/g, "").trim()
        : "";
      if (!objectCode) {
        return { ...row, ObjectDescription: "N/A", ObjectCodeCategory: "N/A" };
      }

      const objectInfo = objectCodes.find(
        (obj) => obj.OBJECT === Number(objectCode)
      );
      const objectCodePrefix = objectCode.substring(0, 2);
      const objectCategory = objectCodeCategories.find((cat) =>
        cat.OBJECT.startsWith(objectCodePrefix)
      );

      return {
        UmbrellaPoolee: row["Umbrella/Poolee (U/P)"] || "",
        GLAccountNumber: row["GL Account Number"] || "",
        GLAccountDescription: row["GL Account Description"] || "",
        TransactionType: row["Transaction Type"] || "",
        Document: row["Document"] || "",
        Date: row["Date"] ? formatDate(row["Date"]) : "",
        Description: row["Description"] || "",
        Budget: row["Budget"] || "",
        Actuals: row["Actuals"] || "",
        Requisitions: row["Requisitions"] || "",
        Encumbrances: row["Encumbrances"] || "",
        Location: row["Location"] || "",
        Unit: row["Unit"] || "",
        Object: row["Object"] || "",
        ObjectDescription: objectInfo ? objectInfo.DESCRIPTION : "N/A",
        ObjectCodeCategory: objectCategory ? objectCategory.DESCRIPTION : "N/A",
        Fund: row["Fund"] || "",
        Function: row["Function"] || "",
      };
    });

    // Create new workbook
    const newWorkbook = xlsx.utils.book_new();
    const newWorksheet = xlsx.utils.json_to_sheet(mappedData);
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "Cleaned Data");

    // Write to buffer
    const outputBuffer = xlsx.write(newWorkbook, {
      type: "buffer",
      bookType: "xlsx",
    });

    return outputBuffer;
  } catch (error) {
    console.error("Error processing Excel:", error);
    throw new Error("Failed to process the Excel file.");
  }
}

/**
 * Reads a JSON file from the filesystem synchronously.
 * @param {string} filename - The JSON file name.
 * @returns {Promise<Object>} - Parsed JSON object.
 */
async function fetchJson(filename) {
  try {
    const filePath = path.join(process.cwd(), filename);
    const fileContent = fs.readFileSync(filePath, "utf8");
    return JSON.parse(fileContent);
  } catch (error) {
    console.error(`Error loading JSON file: ${filename}`, error);
    throw new Error(`Failed to load JSON file: ${filename}`);
  }
}

/**
 * Formats the date to 'YYYY-MM-DD' format.
 * @param {string} excelDate - The raw date from the Excel sheet.
 * @returns {string} - Formatted date string.
 */
function formatDate(excelDate) {
  const parsedDate = new Date(excelDate);

  if (isNaN(parsedDate.getTime())) {
    return excelDate; // Return original value if it's not a valid date
  }

  const month = parsedDate.getMonth() + 1; // getMonth() returns 0-11
  const day = parsedDate.getDate();
  const year = parsedDate.getFullYear();

  return `${month}/${day}/${year}`;
}
