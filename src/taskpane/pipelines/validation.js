/* global console, Excel */
/**
 * Data Validation Pipeline
 * Handles the orchestration of validating Excel data against PDF documents
 */

import { GoogleGenerativeAI } from "@google/generative-ai";
import { parseDocument, compareData } from "../services/gemini.service.js";
import {
  writeDataToNewSheet,
  getSelectedRangeDataSmart,
  highlightMismatches,
} from "../services/excel.service.js";
import { updateState } from "../state.js";
import { updateStatus } from "../components/ui.js";

/**
 * Main orchestration function for data validation
 * 1. Parses the PDF with Gemini Vision
 * 2. Stores parsed data in a new sheet
 * 3. Gets selected data from Excel
 * 4. Compares PDF data with Excel data
 * 5. Highlights mismatched cells and adds comments
 * @param {string} apiKey - Gemini API key
 * @param {File} pdfFile - PDF file to validate against
 * @param {string} parsePrompt - Custom parsing prompt
 * @param {string} comparePrompt - Custom comparison prompt
 */
export async function executeValidation(apiKey, pdfFile, parsePrompt, comparePrompt) {
  try {
    const genAI = new GoogleGenerativeAI(apiKey);

    // Set default prompts if not provided
    const defaultParsePrompt = `Analyze this document and extract all structured data into a JSON format. 

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

    const defaultComparePrompt =
      "Compare the data from the PDF (Dataset 1) with the data from Excel (Dataset 2). Match the keys from the PDF data to the header columns in the Excel data, ignoring case and special characters. Identify any cells in the Excel data that do not match the corresponding PDF data.";

    const finalParsePrompt = parsePrompt.trim() ? parsePrompt : defaultParsePrompt;
    const finalComparePrompt = comparePrompt.trim() ? comparePrompt : defaultComparePrompt;

    // --- 1. Parse the PDF with Gemini Vision ---
    updateStatus("Step 1: Parsing PDF with Gemini...");
    const pdfData = await parseDocument(genAI, pdfFile, finalParsePrompt);
    updateState({ pdfData });
    updateStatus("Step 1: PDF Parsed successfully.");

    await Excel.run(async (context) => {
      // --- 2. Store parsed data in a new sheet ---
      updateStatus("Step 2: Storing parsed data...");
      await writeDataToNewSheet(context, pdfFile.name, pdfData);
      updateStatus("Step 2: Parsed data stored.");

      // --- 3. Get selected data from Excel ---
      updateStatus("Step 3: Getting selected data from Excel...");
      const {
        values: excelData,
        address: selectedRangeAddress,
        wasExpanded,
        originalAddress,
      } = await getSelectedRangeDataSmart(context);

      // Provide user feedback if range was expanded
      if (wasExpanded) {
        updateStatus(
          `Step 3: Single cell selection (${originalAddress}) expanded to data range (${selectedRangeAddress})`
        );
        console.log(`Smart range expansion: ${originalAddress} → ${selectedRangeAddress}`);
      }

      updateState({
        excelData,
        selectedRangeAddress,
        wasRangeExpanded: wasExpanded,
        originalRangeAddress: originalAddress,
      });

      console.log("Excel data retrieved:", {
        dataShape: `${excelData.length} rows × ${excelData[0]?.length || 0} columns`,
        range: selectedRangeAddress,
        wasExpanded: wasExpanded,
        originalRange: originalAddress,
      });

      updateStatus(
        `Step 3: Excel data retrieved (${excelData.length} rows × ${excelData[0]?.length || 0} columns)${wasExpanded ? " (auto-expanded from single cell)" : ""}`
      );

      // --- 4. Compare PDF data with Excel data with Gemini ---
      updateStatus("Step 4: Comparing data with Gemini...");
      const mismatches = await compareData(genAI, pdfData, excelData, finalComparePrompt);
      updateState({ mismatches });
      updateStatus("Step 4: Comparison complete.");

      // --- 5. Highlight mismatched cells and add comments ---
      if (mismatches && mismatches.length > 0) {
        updateStatus(`Step 5: Highlighting ${mismatches.length} mismatch(es)...`);
        await highlightMismatches(context, mismatches, selectedRangeAddress);
        updateStatus(
          `Validation complete. Found ${mismatches.length} mismatch(es). Hover over red cells to see expected values.`
        );
      } else {
        updateStatus("Validation complete. No mismatches found.");
      }
    });
  } catch (error) {
    console.error("Validation failed:", error);
    updateStatus(`Error: ${error.message}`);
    throw error;
  }
}
