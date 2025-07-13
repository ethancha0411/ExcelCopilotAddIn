/* global console */
/**
 * Analyzes the user-selected range in Excel using LLM to understand template structure.
 * Works with both horizontal and vertical (transposed) templates.
 * @param {Excel.RequestContext} context - The request context.
 * @param {GoogleGenerativeAI} genAI - The GoogleGenerativeAI instance for analysis.
 * @returns {Promise<{headers: string[], templateRange: Excel.Range, templateStructure: object}>} Template analysis results.
 */
export async function analyzeTemplate(context, genAI) {
  const templateRange = context.workbook.getSelectedRange();
  templateRange.load(["values", "address", "rowCount", "columnCount"]);
  await context.sync();

  if (!templateRange.values || templateRange.rowCount === 0 || templateRange.columnCount === 0) {
    throw new Error(
      "The selected range is empty. Please select a valid template area with headers."
    );
  }

  // Use LLM to analyze the template structure
  const { analyzeTemplateStructure } = await import("./gemini.service.js");
  const templateStructure = await analyzeTemplateStructure(
    genAI,
    templateRange.values,
    templateRange.address
  );

  // Extract field names for backward compatibility
  const headers = templateStructure.fields.map((field) => field.fieldName);

  console.log("Template structure analysis:", templateStructure);
  console.log("Extracted headers:", headers);

  return { headers, templateRange, templateStructure };
}

/**
 * Sanitizes a value for Excel cell insertion
 * @param {any} value - The value to sanitize
 * @returns {string|number|boolean} Sanitized value safe for Excel
 */
function sanitizeValueForExcel(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "object") {
    // Convert objects/arrays to readable strings
    return JSON.stringify(value);
  }

  if (typeof value === "string") {
    // Ensure string doesn't contain problematic characters
    return value.toString();
  }

  if (typeof value === "number" || typeof value === "boolean") {
    return value;
  }

  return String(value);
}

/**
 * Validates that mapped data respects the template structure
 * @param {Array} mappedData - Array of mapped data objects
 * @param {string[]} templateHeaders - Original template headers
 * @returns {Array} Validated and corrected mapped data
 */
function validateMappedData(mappedData, templateHeaders) {
  if (!mappedData || mappedData.length === 0) {
    return [];
  }

  return mappedData.map((dataRow) => {
    const validatedRow = {};

    // Ensure all template headers are present
    templateHeaders.forEach((header) => {
      if (Object.prototype.hasOwnProperty.call(dataRow, header)) {
        validatedRow[header] = dataRow[header];
      } else {
        validatedRow[header] = ""; // Default empty value
      }
    });

    return validatedRow;
  });
}

/**
 * Populates the Excel template with LLM-mapped data using structure analysis
 * @param {Excel.RequestContext} context - The request context
 * @param {Array} mappedData - Array of objects where keys are template headers and values are the data
 * @param {Excel.Range} templateRange - The user-selected range for the template
 * @param {object} templateStructure - The LLM-analyzed template structure
 */
export async function populateTemplate(context, mappedData, templateRange, templateStructure) {
  if (!mappedData || mappedData.length === 0) {
    return; // Nothing to populate
  }

  if (!templateStructure || !templateStructure.fields) {
    throw new Error("Template structure information is missing");
  }

  // Load template range properties - including address for debugging
  templateRange.load(["values", "rowIndex", "columnIndex", "address"]);
  await context.sync();

  // Validate that mapped data respects template structure
  const headers = templateStructure.fields.map((field) => field.fieldName);
  const validatedData = validateMappedData(mappedData, headers);

  if (validatedData.length === 0) {
    throw new Error("No valid data to populate after validation");
  }

  // Handle different template orientations
  if (templateStructure.orientation === "vertical") {
    await populateVerticalTemplate(context, validatedData, templateRange, templateStructure);
  } else {
    await populateHorizontalTemplate(context, validatedData, templateRange, templateStructure);
  }
}

/**
 * Populates a vertical (transposed) template where fields are in rows
 * @param {Excel.RequestContext} context - The request context
 * @param {Array} validatedData - Validated mapped data
 * @param {Excel.Range} templateRange - The template range
 * @param {object} templateStructure - The template structure
 */
async function populateVerticalTemplate(context, validatedData, templateRange, templateStructure) {
  // For vertical templates, we populate each field's value location directly
  // We'll use the first data row since vertical templates typically show one record
  const dataRow = validatedData[0] || {};

  console.log("🔍 PopulateVerticalTemplate Debug Info:");
  console.log("Template Range Position:", {
    rowIndex: templateRange.rowIndex,
    columnIndex: templateRange.columnIndex,
    address: templateRange.address,
  });
  console.log("Data to populate:", dataRow);

  // Create an array to batch all the updates
  const cellUpdates = [];

  templateStructure.fields.forEach((field, index) => {
    const value = dataRow[field.fieldName];
    if (value !== undefined && value !== null) {
      // Calculate absolute positions by adding template range offset
      const absoluteRow = templateRange.rowIndex + field.valueLocation.row;
      const absoluteCol = templateRange.columnIndex + field.valueLocation.col;

      console.log(`Field ${index}: "${field.fieldName}"`);
      console.log(
        `  - Relative position: row ${field.valueLocation.row}, col ${field.valueLocation.col}`
      );
      console.log(`  - Absolute position: row ${absoluteRow}, col ${absoluteCol}`);
      console.log(`  - Value: "${value}"`);

      // Validate that the absolute position is reasonable
      if (absoluteRow < 0 || absoluteCol < 0) {
        throw new Error(
          `Invalid absolute position for field "${field.fieldName}": row ${absoluteRow}, col ${absoluteCol}. Check template analysis.`
        );
      }

      cellUpdates.push({
        fieldName: field.fieldName,
        row: absoluteRow,
        col: absoluteCol,
        value: sanitizeValueForExcel(value),
        relativePosition: { row: field.valueLocation.row, col: field.valueLocation.col },
      });
    } else {
      console.log(`Field "${field.fieldName}" skipped - no value found in data`);
    }
  });

  console.log("📝 Cell Updates to Apply:", cellUpdates.length);

  // Apply all updates in batch
  for (const update of cellUpdates) {
    try {
      const cell = templateRange.worksheet.getCell(update.row, update.col);
      cell.values = [[update.value]];
      console.log(`✅ Updated cell (${update.row}, ${update.col}) with: "${update.value}"`);
    } catch (error) {
      console.error(
        `❌ Failed to update cell (${update.row}, ${update.col}) for field "${update.fieldName}":`,
        error
      );
      throw new Error(
        `Failed to populate field "${update.fieldName}" at position (${update.row}, ${update.col}): ${error.message}`
      );
    }
  }

  await context.sync();

  console.log("✅ Vertical template population completed successfully");
}

/**
 * Populates a horizontal template where fields are column headers
 * @param {Excel.RequestContext} context - The request context
 * @param {Array} validatedData - Validated mapped data
 * @param {Excel.Range} templateRange - The template range
 * @param {object} templateStructure - The template structure
 */
async function populateHorizontalTemplate(
  context,
  validatedData,
  templateRange,
  templateStructure
) {
  // For horizontal templates, we add rows of data below the headers
  const headers = templateStructure.fields.map((field) => field.fieldName);

  // Find the first empty row to start writing data
  const startRow = templateRange.rowIndex + 1; // Start populating below the header row

  // Construct the 2D array for the new rows
  const newRowsData = validatedData.map((dataObject) => {
    const newRow = [];
    headers.forEach((header) => {
      const value = dataObject[header] !== undefined ? dataObject[header] : "";
      newRow.push(sanitizeValueForExcel(value));
    });
    return newRow;
  });

  // Get the range to write the new data to
  const populateRange = templateRange.worksheet.getRangeByIndexes(
    startRow,
    templateRange.columnIndex,
    newRowsData.length,
    headers.length
  );

  populateRange.values = newRowsData;
  await context.sync();
}

/**
 * Writes a JSON object to a new or existing Excel worksheet.
 * Useful for debugging or showing the raw extracted data.
 * @param {Excel.RequestContext} context - The request context.
 * @param {string} sheetNamePrefix - A prefix for the new sheet name.
 * @param {object} data - The JSON object to write.
 */
export async function writeDataToNewSheet(context, sheetNamePrefix, data) {
  let sheetName = `${sheetNamePrefix}`.substring(0, 31);
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

/**
 * Gets the selected range data from Excel
 * @param {Excel.RequestContext} context - The request context
 * @returns {Promise<{values: Array, address: string}>} Selected range data and address
 */
export async function getSelectedRangeData(context) {
  const selectedRange = context.workbook.getSelectedRange();
  selectedRange.load(["values", "address"]);
  await context.sync();

  return {
    values: selectedRange.values,
    address: selectedRange.address,
  };
}

/**
 * Highlights mismatched cells and adds comments
 * @param {Excel.RequestContext} context - The request context
 * @param {Array} mismatches - Array of mismatch objects
 * @param {string} rangeAddress - Address of the range to highlight
 */
export async function highlightMismatches(context, mismatches, rangeAddress) {
  if (!mismatches || mismatches.length === 0) {
    return;
  }

  const selectedRange = context.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);
  selectedRange.load("values");
  await context.sync();

  const excelData = selectedRange.values;

  // Clear existing formatting
  selectedRange.format.fill.clear();
  await context.sync();

  // Highlight all mismatched cells
  mismatches.forEach((mismatch) => {
    if (mismatch.row < excelData.length && mismatch.col < excelData[0].length) {
      const cell = selectedRange.getCell(mismatch.row, mismatch.col);
      cell.format.fill.color = "red";
    }
  });
  await context.sync();

  // Add comments to each mismatched cell
  const validMismatches = mismatches.filter(
    (mismatch) =>
      mismatch.row < excelData.length &&
      mismatch.col < excelData[0].length &&
      mismatch.expectedValue !== undefined &&
      mismatch.actualValue !== undefined
  );

  for (const mismatch of validMismatches) {
    try {
      const commentText = `Mismatch Found:\nExpected: ${mismatch.expectedValue}\nActual: ${mismatch.actualValue}`;
      const cell = selectedRange.getCell(mismatch.row, mismatch.col);
      context.workbook.comments.add(cell, commentText);
    } catch (commentError) {
      console.warn(
        `Failed to add comment to cell (${mismatch.row}, ${mismatch.col}):`,
        commentError.message
      );
    }
  }

  await context.sync();
}
