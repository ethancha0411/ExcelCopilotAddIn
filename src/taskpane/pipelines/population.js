/* global console, Excel */
/**
 * Template Population Pipeline
 * Handles the orchestration of extracting data from documents and populating Excel templates
 */

import { GoogleGenerativeAI } from "@google/generative-ai";
import { extractDataFromDocument, mapDataToTemplate } from "../services/gemini.service.js";
import { analyzeTemplate, populateTemplate } from "../services/excel.service.js";
import { updateState, getState } from "../state.js";
import { updateStatus } from "../components/ui.js";

/**
 * Main orchestration function for template population
 * 1. Analyzes the selected Excel template using LLM
 * 2. Extracts data from the uploaded document
 * 3. Creates a UI for the user to map the data
 * @param {string} apiKey - Gemini API key
 * @param {File} documentFile - Document file to extract data from
 * @param {string} userPrompt - Custom extraction prompt
 */
export async function executePopulation(apiKey, documentFile, userPrompt) {
  try {
    // Initialize Gemini AI for template analysis
    const genAI = new GoogleGenerativeAI(apiKey);

    // --- 1. Analyze Excel Template with LLM ---
    updateStatus("Step 1: Analyzing Excel template structure with AI...");
    await Excel.run(async (context) => {
      const { headers, templateRange, templateStructure } = await analyzeTemplate(context, genAI);
      updateState({
        templateHeaders: headers,
        templateRangeAddress: templateRange.address,
        templateStructure: templateStructure,
      });

      console.log("Template analysis complete:", {
        orientation: templateStructure.orientation,
        fieldCount: templateStructure.fields.length,
        headers: headers,
      });

      updateStatus(
        `Template analyzed: ${templateStructure.orientation} format with ${templateStructure.fields.length} fields`
      );
    });

    // --- 2. Extract Data from Document ---
    updateStatus("Step 2: Extracting data from document with Gemini...");
    const defaultPrompt = `Analyze this document and extract all structured data into a comprehensive JSON object. 

Focus on extracting:
- Property information (name, address, city, state, zip, type)
- Tenant/Customer information (name, unit, contact details)
- Financial data (amounts, balances, charges, due dates)
- Account information (account numbers, statement periods, dates)
- Payment information (payment methods, remittance details)
- Any other structured data fields

Use descriptive, clear field names that would be easily understood (e.g., "propertyName", "tenantName", "totalDue", "statementDate", "propertyAddress", etc.).

If the document contains multiple records, return an array of objects. Otherwise, return a single object.
Ensure all values are properly typed (strings, numbers, dates, booleans).`;
    const finalPrompt = userPrompt.trim() ? userPrompt : defaultPrompt;

    const extractedData = await extractDataFromDocument(genAI, documentFile, finalPrompt);
    updateState({ extractedData });
    updateStatus("Step 2: Data extracted successfully.");

    // --- 3. Map data to template using LLM ---
    updateStatus("Step 3: Mapping data to template...");
    const state = getState();

    console.log("Template headers:", state.templateHeaders);
    console.log("Template structure:", state.templateStructure);
    console.log("Extracted data:", extractedData);

    const mappedData = await mapDataToTemplate(
      genAI,
      extractedData,
      state.templateHeaders,
      state.templateStructure
    );

    console.log("Mapped data:", mappedData);

    updateState({ mappedData });
    updateStatus("Step 3: Data mapped successfully.");

    // --- 4. Populate the template ---
    updateStatus("Step 4: Populating template...");
    await Excel.run(async (context) => {
      const templateRange = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange(state.templateRangeAddress);
      await populateTemplate(context, mappedData, templateRange, state.templateStructure);
      updateStatus("Population complete!");
    });
  } catch (error) {
    console.error("Template population failed:", error);
    updateStatus(`Error: ${error.message}`);
    throw error;
  }
}
