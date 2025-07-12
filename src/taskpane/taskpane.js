/* global console, document, Excel, Office, FileReader */
import { GoogleGenerativeAI } from "@google/generative-ai";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const sideloadMsg = document.getElementById("sideload-msg");
    if (sideloadMsg) {
      sideloadMsg.style.display = "none";
    }
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("verify-button").onclick = verifyData;
  }
});

// --- Core Function ---
async function verifyData() {
  const status = document.getElementById("status");
  status.textContent = "Starting verification...";

  try {
    // --- 1. Get all inputs from the user ---
    const apiKey = document.getElementById("api-key").value;
    let parsePrompt = document.getElementById("parse-prompt").value;
    let comparePrompt = document.getElementById("compare-prompt").value;
    const pdfFile = document.getElementById("pdf-upload").files[0];

    if (!apiKey || !pdfFile) {
      status.textContent = "Please provide a Gemini API key and select a PDF file.";
      return;
    }

    // --- Use default prompts if user leaves them blank ---
    if (!parsePrompt.trim()) {
      parsePrompt = `Analyze this document and extract all structured data into a JSON format. 
  
Please follow these guidelines:
1. Identify the document type (invoice, receipt, form, table, report, etc.)
2. Extract all text content, preserving relationships between data points
3. For tabular data: organize into arrays of objects with consistent field names
4. For forms: extract field labels and their corresponding values
5. For invoices/receipts: extract items, quantities, prices, totals, dates, vendor info
6. Use descriptive field names (e.g., "customerName", "invoiceDate", "lineItems")
7. Preserve numerical values as numbers, not strings
8. Include dates in ISO format (YYYY-MM-DD) when possible
9. Group related information into nested objects
10. If the document contains multiple tables or sections, organize them separately

Return a well-structured JSON object that captures all the meaningful data from the document.
The JSON should be easily comparable with Excel spreadsheet data.`;
    }
    if (!comparePrompt.trim()) {
      comparePrompt =
        "Compare the data from the PDF (Dataset 1) with the data from Excel (Dataset 2). Match the keys from the PDF data to the header columns in the Excel data, ignoring case and special characters. Identify any cells in the Excel data that do not match the corresponding PDF data.";
    }

    const genAI = new GoogleGenerativeAI(apiKey);

    // --- 2. Parse the PDF with Gemini Vision ---
    status.textContent = "Step 1: Parsing PDF with Gemini...";
    const pdfData = await callGeminiParse(genAI, pdfFile, parsePrompt);
    status.textContent = "Step 1: PDF Parsed successfully.";

    await Excel.run(async (context) => {
      // --- 3. Store parsed data in a new sheet ---
      status.textContent = "Step 2: Storing parsed data...";
      await writeDataToNewSheet(context, pdfFile.name, pdfData);
      status.textContent = "Step 2: Parsed data stored.";

      // --- 4. Get selected data from Excel ---
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["values", "address"]);
      await context.sync();
      const excelData = selectedRange.values;

      // --- 5. Compare PDF data with Excel data with Gemini ---
      status.textContent = "Step 3: Comparing data with Gemini...";
      const mismatches = await callGeminiCompare(genAI, pdfData, excelData, comparePrompt);
      status.textContent = "Step 3: Comparison complete.";

      // --- 6. Highlight mismatched cells ---
      if (mismatches && mismatches.length > 0) {
        status.textContent = `Step 4: Highlighting ${mismatches.length} mismatch(es)...`;
        selectedRange.format.fill.clear();
        await context.sync();

        mismatches.forEach((mismatch) => {
          if (mismatch.row < excelData.length && mismatch.col < excelData[0].length) {
            const cell = selectedRange.getCell(mismatch.row, mismatch.col);
            cell.format.fill.color = "red";
          }
        });
        await context.sync();
        status.textContent = `Verification complete. Found ${mismatches.length} mismatch(es).`;
      } else {
        status.textContent = "Verification complete. No mismatches found.";
      }
    });
  } catch (error) {
    console.error("Verification failed:", error);
    status.textContent = `Error: ${error.message}`;
  }
}

// --- Helper Functions ---

/**
 * Writes a JSON object to a new or existing Excel worksheet.
 * @param {Excel.RequestContext} context - The request context.
 * @param {string} pdfFileName - The name of the original PDF file.
 * @param {object} data - The JSON object to write.
 */
async function writeDataToNewSheet(context, pdfFileName, data) {
  let sheetName = `PDF_${pdfFileName.replace(/.pdf$/i, "")}`.substring(0, 31);
  let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  await context.sync();

  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(sheetName);
  } else {
    sheet.getUsedRange().clear();
  }

  const headers = [["Key", "Value"]];
  // Convert the data object to an array of [key, value] pairs,
  // ensuring that any nested objects are stringified to prevent Excel errors.
  const dataRows = Object.entries(data).map(([key, value]) => {
    if (typeof value === "object" && value !== null) {
      // Stringify nested objects/arrays for clean insertion into a cell
      return [key, JSON.stringify(value, null, 2)];
    }
    return [key, value];
  });

  const headerRange = sheet.getRange("A1:B1");
  headerRange.values = headers;
  headerRange.format.font.bold = true;

  if (dataRows.length > 0) {
    const dataRange = sheet.getRangeByIndexes(1, 0, dataRows.length, 2);
    dataRange.values = dataRows;
  }

  sheet.getUsedRange().format.autofitColumns();
  await context.sync();
}

async function fileToGenerativePart(file) {
  const base64EncodedData = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result.split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
  return {
    inlineData: { data: base64EncodedData, mimeType: file.type },
  };
}

async function callGeminiParse(genAI, pdfFile, prompt) {
  const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });
  const imagePart = await fileToGenerativePart(pdfFile);
  const fullPrompt = `${prompt}. Respond with only the JSON object, without any markdown formatting.`;

  const result = await model.generateContent([fullPrompt, imagePart]);
  const response = await result.response;
  const text = response.text();

  try {
    const cleanedText = text
      .replace(/```json/g, "")
      .replace(/```/g, "")
      .trim();
    return JSON.parse(cleanedText);
  } catch (error) {
    console.error("Failed to parse JSON from Gemini response:", error, text);
    throw new Error(
      "Could not parse structured data from the PDF. The LLM returned an invalid format."
    );
  }
}

async function callGeminiCompare(genAI, pdfData, excelData, prompt) {
  const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });
  const fullPrompt = `
    You are a data comparison assistant. Your task is to compare two datasets and identify mismatches in the second dataset.
    **INSTRUCTIONS:**
    1.  The user has provided a custom prompt: "${prompt}".
    2.  Dataset 1 is a JSON object extracted from a source document.
    3.  Dataset 2 is a 2D array of data from an Excel sheet, where the first row contains headers.
    4.  Analyze the user's prompt to understand how to map fields from Dataset 1 to columns in Dataset 2.
    5.  Compare the values for each corresponding row.
    6.  Identify any cells in Dataset 2 that do not match the corresponding value in Dataset 1.
    7.  Return a JSON array of objects, where each object represents a single mismatched cell in Dataset 2.
    8.  Each object in the array must have two keys: "row" (the 0-based row index of the mismatch in Dataset 2) and "col" (the 0-based column index).
    9.  If there are no mismatches, return an empty array [].
    10. Respond with ONLY the JSON array, without any explanation or markdown formatting.
    **Dataset 1 (from PDF):**
    ${JSON.stringify(pdfData, null, 2)}
    **Dataset 2 (from Excel):**
    ${JSON.stringify(excelData, null, 2)}
    **JSON Array of Mismatches:**
  `;

  const result = await model.generateContent(fullPrompt);
  const response = await result.response;
  const text = response.text();

  try {
    const cleanedText = text
      .replace(/```json/g, "")
      .replace(/```/g, "")
      .trim();
    return JSON.parse(cleanedText);
  } catch (error) {
    console.error("Failed to parse JSON from Gemini comparison response:", error, text);
    throw new Error("Could not parse comparison results. The LLM returned an invalid format.");
  }
}
