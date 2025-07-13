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
  templateRange.load(["values", "rowIndex", "columnIndex", "address", "rowCount", "columnCount"]);
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
 * Debug function to log template structure details
 * @param {object} templateStructure - The template structure
 * @param {Excel.Range} templateRange - The template range
 */
function debugTemplateStructure(templateStructure, templateRange) {
  console.log("🔍 TEMPLATE STRUCTURE DEBUG:");
  console.log("Template Orientation:", templateStructure.orientation);
  console.log("Template Range:", templateRange.address);
  console.log("Template Dimensions:", `${templateRange.rowCount}×${templateRange.columnCount}`);
  console.log(
    "Template Position:",
    `Row ${templateRange.rowIndex}, Col ${templateRange.columnIndex}`
  );

  console.log("\nFields Analysis:");
  templateStructure.fields.forEach((field, index) => {
    console.log(`Field ${index}: "${field.fieldName}"`);
    console.log(
      `  - Field Location: {row:${field.fieldLocation.row}, col:${field.fieldLocation.col}}`
    );
    console.log(
      `  - Value Location: {row:${field.valueLocation.row}, col:${field.valueLocation.col}}`
    );
    console.log(`  - Description: ${field.description}`);
    console.log(`  - Data Type: ${field.dataType}`);
  });

  console.log("\nTemplate Data Preview:");
  if (templateRange.values && templateRange.values.length > 0) {
    templateRange.values
      .slice(0, Math.min(5, templateRange.values.length))
      .forEach((row, rowIndex) => {
        console.log(`Row ${rowIndex}:`, row);
      });
  }
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
  // Debug the template structure first
  debugTemplateStructure(templateStructure, templateRange);

  console.log("🔍 PopulateHorizontalTemplate Debug Info:");
  console.log("Data to populate:", validatedData);

  // For horizontal templates, we add rows of data below the headers
  // PROGRAMMATICALLY determine where values should go (ignore LLM's valueLocation)

  // Find the header row (assume it's the first row with field names)
  let headerRowIndex = 0;

  // Look for the row that contains the most field names
  for (let row = 0; row < templateRange.rowCount; row++) {
    const rowValues = templateRange.values[row] || [];
    const matchingFields = templateStructure.fields.filter((field) => {
      return rowValues.some(
        (cellValue) =>
          cellValue && cellValue.toString().toLowerCase().includes(field.fieldName.toLowerCase())
      );
    });

    if (matchingFields.length > headerRowIndex) {
      headerRowIndex = row;
    }
  }

  console.log(`Header row identified at template row index: ${headerRowIndex}`);

  // The data should start in the row immediately BELOW the header row
  const dataStartRow = templateRange.rowIndex + headerRowIndex + 1;

  console.log(`Data will be populated starting at absolute row: ${dataStartRow}`);

  // Create a mapping of field names to their column positions in the header row
  const fieldToColumnMap = new Map();

  templateStructure.fields.forEach((field) => {
    // Find which column this field is in based on the header row
    const headerRow = templateRange.values[headerRowIndex] || [];

    for (let col = 0; col < headerRow.length; col++) {
      const cellValue = headerRow[col];
      if (cellValue && cellValue.toString().toLowerCase().includes(field.fieldName.toLowerCase())) {
        const absoluteCol = templateRange.columnIndex + col;
        fieldToColumnMap.set(field.fieldName, absoluteCol);
        console.log(
          `Field "${field.fieldName}" mapped to column ${absoluteCol} (Excel col ${String.fromCharCode(65 + absoluteCol)})`
        );
        break;
      }
    }
  });

  console.log("Field to column mapping:", Array.from(fieldToColumnMap.entries()));

  // Now populate each data row
  for (let dataRowIndex = 0; dataRowIndex < validatedData.length; dataRowIndex++) {
    const dataObject = validatedData[dataRowIndex];
    const targetRowIndex = dataStartRow + dataRowIndex;

    console.log(`\nPopulating data row ${dataRowIndex} at absolute row ${targetRowIndex}:`);

    // Populate each field in this row
    for (const [fieldName, columnIndex] of fieldToColumnMap) {
      const value = dataObject[fieldName];
      const sanitizedValue = sanitizeValueForExcel(value);

      try {
        const cell = templateRange.worksheet.getCell(targetRowIndex, columnIndex);
        cell.values = [[sanitizedValue]];

        const excelAddress = `${String.fromCharCode(65 + columnIndex)}${targetRowIndex + 1}`;
        console.log(`  ✅ ${fieldName}: "${sanitizedValue}" → ${excelAddress}`);
      } catch (error) {
        console.error(
          `  ❌ Failed to populate ${fieldName} at (${targetRowIndex}, ${columnIndex}):`,
          error
        );
        throw new Error(`Failed to populate field "${fieldName}": ${error.message}`);
      }
    }
  }

  await context.sync();
  console.log("✅ Horizontal template population completed successfully");
}

/**
 * Populates a vertical (transposed) template where fields are in rows
 * @param {Excel.RequestContext} context - The request context
 * @param {Array} validatedData - Validated mapped data
 * @param {Excel.Range} templateRange - The template range
 * @param {object} templateStructure - The template structure
 */
async function populateVerticalTemplate(context, validatedData, templateRange, templateStructure) {
  // Debug the template structure first
  debugTemplateStructure(templateStructure, templateRange);

  // For vertical templates, we populate each field's value location directly
  // We'll use the first data row since vertical templates typically show one record
  const dataRow = validatedData[0] || {};

  console.log("🔍 PopulateVerticalTemplate Debug Info:");
  console.log("Data to populate:", dataRow);

  // PROGRAMMATICALLY determine where values should go (ignore LLM's valueLocation)
  // For vertical templates, values typically go to the RIGHT of field names

  const cellUpdates = [];

  templateStructure.fields.forEach((field, index) => {
    const value = dataRow[field.fieldName];
    if (value !== undefined && value !== null) {
      // Use field location and ALWAYS place value to the right (+1 column)
      const fieldRow = field.fieldLocation.row;
      const fieldCol = field.fieldLocation.col;

      // PROGRAMMATIC APPROACH: Always place values one column to the right of field names
      const targetRow = templateRange.rowIndex + fieldRow;
      const targetCol = templateRange.columnIndex + fieldCol + 1;

      console.log(`Field ${index}: "${field.fieldName}"`);
      console.log(`  - Field location: template row ${fieldRow}, col ${fieldCol}`);
      console.log(`  - Value placement: absolute position (${targetRow}, ${targetCol})`);
      console.log(`  - Excel address: ${String.fromCharCode(65 + targetCol)}${targetRow + 1}`);
      console.log(`  - Value: "${value}"`);

      // Validate that the absolute position is reasonable
      if (targetRow < 0 || targetCol < 0) {
        throw new Error(
          `Invalid absolute position for field "${field.fieldName}": row ${targetRow}, col ${targetCol}.`
        );
      }

      cellUpdates.push({
        fieldName: field.fieldName,
        row: targetRow,
        col: targetCol,
        value: sanitizeValueForExcel(value),
        excelAddress: `${String.fromCharCode(65 + targetCol)}${targetRow + 1}`,
      });
    } else {
      console.log(`Field "${field.fieldName}" skipped - no value found in data`);
    }
  });

  console.log("📝 Cell Updates to Apply:", cellUpdates.length);
  cellUpdates.forEach((update, index) => {
    console.log(
      `Update ${index}: ${update.fieldName} → ${update.excelAddress} = "${update.value}"`
    );
  });

  // Apply all updates in batch
  for (const update of cellUpdates) {
    try {
      const cell = templateRange.worksheet.getCell(update.row, update.col);
      cell.values = [[update.value]];
      console.log(
        `✅ Updated ${update.excelAddress} with: "${update.value}" (field: ${update.fieldName})`
      );
    } catch (error) {
      console.error(
        `❌ Failed to update ${update.excelAddress} for field "${update.fieldName}":`,
        error
      );
      throw new Error(
        `Failed to populate field "${update.fieldName}" at ${update.excelAddress}: ${error.message}`
      );
    }
  }

  await context.sync();

  console.log("✅ Vertical template population completed successfully");
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
 * Detects if the user has selected only a single cell and automatically expands
 * to include all contiguous data in the current region
 * @param {Excel.RequestContext} context - The request context
 * @returns {Promise<Excel.Range>} The original range if multi-cell, or expanded range if single cell
 */
/**
 * Adds intelligent padding around a detected range to ensure we capture complete template structures
 * @param {Excel.RequestContext} context - The request context
 * @param {Excel.Range} detectedRange - The initially detected range
 * @param {number} rowPadding - Number of rows to add above and below
 * @param {number} colPadding - Number of columns to add left and right
 * @returns {Promise<Excel.Range>} Padded range
 */
async function addRangePadding(context, detectedRange, rowPadding = 3, colPadding = 2) {
  const worksheet = detectedRange.worksheet;

  console.log("📝 Adding Range Padding:");
  console.log(
    `Original range: ${detectedRange.address} at (${detectedRange.rowIndex}, ${detectedRange.columnIndex})`
  );
  console.log(`Range size: ${detectedRange.rowCount}×${detectedRange.columnCount}`);
  console.log(`Padding: ${rowPadding} rows, ${colPadding} columns`);

  // Calculate padded boundaries
  const startRow = Math.max(0, detectedRange.rowIndex - rowPadding);
  const startCol = Math.max(0, detectedRange.columnIndex - colPadding);
  const endRow = detectedRange.rowIndex + detectedRange.rowCount + rowPadding - 1;
  const endCol = detectedRange.columnIndex + detectedRange.columnCount + colPadding - 1;

  console.log(
    `Calculated boundaries: start(${startRow}, ${startCol}) to end(${endRow}, ${endCol})`
  );

  // Get the padded range
  const paddedRange = worksheet.getRangeByIndexes(
    startRow,
    startCol,
    endRow - startRow + 1,
    endCol - startCol + 1
  );

  paddedRange.load(["values", "address", "rowCount", "columnCount"]);
  await context.sync();

  console.log(
    `✅ Padded range created: ${paddedRange.address} (${paddedRange.rowCount}×${paddedRange.columnCount})`
  );

  return paddedRange;
}

/**
 * Validates if a detected range is suitable for template analysis
 * @param {Excel.Range} range - The range to validate
 * @returns {boolean} True if range seems suitable for template analysis
 */
function isValidTemplateRange(range) {
  // Template should have at least 2 rows and 2 columns for meaningful analysis
  if (range.rowCount < 2 || range.columnCount < 2) {
    return false;
  }

  // Avoid extremely wide ranges that are likely just headers
  if (range.rowCount === 1 && range.columnCount > 10) {
    return false;
  }

  // Avoid extremely tall ranges that might be just a single column
  if (range.columnCount === 1 && range.rowCount > 20) {
    return false;
  }

  return true;
}

export async function getSmartRange(context) {
  const selectedRange = context.workbook.getSelectedRange();
  selectedRange.load(["values", "address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
  await context.sync();

  // Check if only one cell is selected
  if (selectedRange.rowCount === 1 && selectedRange.columnCount === 1) {
    console.log("Single cell detected:", selectedRange.address);
    console.log("Expanding to include all contiguous data...");

    let bestRange = selectedRange;
    let rangeSource = "original";

    // Try to get the current region (contiguous data around the selected cell)
    try {
      const currentRegion = selectedRange.getCurrentRegion();
      currentRegion.load([
        "values",
        "address",
        "rowCount",
        "columnCount",
        "rowIndex",
        "columnIndex",
      ]);
      await context.sync();

      console.log(
        "Current region detected:",
        currentRegion.address,
        `(${currentRegion.rowCount}×${currentRegion.columnCount})`
      );

      // If the current region seems valid, use it
      if (isValidTemplateRange(currentRegion)) {
        bestRange = currentRegion;
        rangeSource = "currentRegion";
      } else if (currentRegion.rowCount > 1 || currentRegion.columnCount > 1) {
        // Current region exists but might be incomplete, try adding padding
        console.log("Current region seems incomplete, adding padding...");
        const paddedRange = await addRangePadding(context, currentRegion);
        console.log(
          "Padded range:",
          paddedRange.address,
          `(${paddedRange.rowCount}×${paddedRange.columnCount})`
        );

        if (isValidTemplateRange(paddedRange)) {
          bestRange = paddedRange;
          rangeSource = "paddedCurrentRegion";
        } else {
          bestRange = currentRegion;
          rangeSource = "currentRegion";
        }
      }
    } catch (regionError) {
      console.log("Error getting current region:", regionError.message);
    }

    // If we still don't have a good range, try worksheet's used range
    if (rangeSource === "original") {
      console.log("No suitable current region found, trying worksheet used range...");
      const worksheet = selectedRange.worksheet;

      try {
        const usedRange = worksheet.getUsedRange();
        usedRange.load(["values", "address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
        await context.sync();

        if (usedRange.rowCount > 0 && usedRange.columnCount > 0) {
          console.log(
            "Used range detected:",
            usedRange.address,
            `(${usedRange.rowCount}×${usedRange.columnCount})`
          );
          bestRange = usedRange;
          rangeSource = "usedRange";
        }
      } catch (usedRangeError) {
        console.log("No used range found:", usedRangeError.message);
      }
    }

    // Final validation and fallback
    if (rangeSource === "original") {
      console.log("Using original single cell as fallback");
    } else {
      console.log(`Using ${rangeSource} as expanded range:`, bestRange.address);
    }

    // Ensure the range has all required properties loaded
    bestRange.load(["values", "address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    await context.sync();

    return bestRange;
  }

  // Multi-cell selection - return as is
  console.log("Multi-cell selection detected:", selectedRange.address);
  return selectedRange;
}

/**
 * Enhanced version of getSelectedRangeData that automatically handles single cell selection
 * @param {Excel.RequestContext} context - The request context
 * @returns {Promise<{values: Array, address: string, wasExpanded: boolean}>} Enhanced range data
 */
export async function getSelectedRangeDataSmart(context) {
  const originalRange = context.workbook.getSelectedRange();
  originalRange.load(["rowCount", "columnCount", "address"]);
  await context.sync();

  const wasSingleCell = originalRange.rowCount === 1 && originalRange.columnCount === 1;
  const smartRange = await getSmartRange(context);

  return {
    values: smartRange.values,
    address: smartRange.address,
    wasExpanded: wasSingleCell,
    originalAddress: wasSingleCell ? originalRange.address : smartRange.address,
  };
}

/**
 * Original getSelectedRangeData function for backward compatibility
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
 * Enhanced template range detection specifically for template analysis
 * Ensures we capture complete template structures, not just headers
 * @param {Excel.RequestContext} context - The request context
 * @returns {Promise<Excel.Range>} Enhanced template range
 */
async function getSmartTemplateRange(context) {
  const smartRange = await getSmartRange(context);

  console.log("🔍 Smart Template Range Analysis:");
  console.log(
    `Initial range: ${smartRange.address} (${smartRange.rowCount}×${smartRange.columnCount})`
  );

  // ALWAYS expand single-row ranges for templates since they're likely just headers
  if (smartRange.rowCount === 1) {
    console.log("🚨 DETECTED SINGLE ROW RANGE - This is likely headers only!");
    console.log("Applying mandatory expansion for template analysis...");

    try {
      // Add significant padding below for templates since they need space for data
      const expandedRange = await addRangePadding(context, smartRange, 8, 2);
      console.log(
        "✅ Expanded single-row template range:",
        expandedRange.address,
        `(${expandedRange.rowCount}×${expandedRange.columnCount})`
      );
      return expandedRange;
    } catch (error) {
      console.log("❌ Failed to expand single-row template range:", error.message);
      return smartRange;
    }
  }

  // If we have a very wide, shallow range (likely just headers), try to expand it
  if (smartRange.rowCount <= 2 && smartRange.columnCount > 5) {
    console.log("Detected potential header-only range, expanding for template analysis...");

    try {
      // Add more padding below for templates since they usually have empty rows for data
      const expandedRange = await addRangePadding(context, smartRange, 5, 1);
      console.log(
        "Expanded template range:",
        expandedRange.address,
        `(${expandedRange.rowCount}×${expandedRange.columnCount})`
      );
      return expandedRange;
    } catch (error) {
      console.log("Failed to expand template range, using original:", error.message);
      return smartRange;
    }
  }

  // For vertical templates, ensure we have enough columns
  if (smartRange.rowCount > 3 && smartRange.columnCount < 3) {
    console.log("Detected potential vertical template, expanding horizontally...");

    try {
      const expandedRange = await addRangePadding(context, smartRange, 1, 3);
      console.log(
        "Expanded vertical template range:",
        expandedRange.address,
        `(${expandedRange.rowCount}×${expandedRange.columnCount})`
      );
      return expandedRange;
    } catch (error) {
      console.log("Failed to expand vertical template range, using original:", error.message);
      return smartRange;
    }
  }

  console.log("✅ Using original smart range (already suitable):", smartRange.address);
  return smartRange;
}

/**
 * Enhanced template analysis that automatically handles single cell selection
 * Works with both horizontal and vertical (transposed) templates.
 * @param {Excel.RequestContext} context - The request context.
 * @param {GoogleGenerativeAI} genAI - The GoogleGenerativeAI instance for analysis.
 * @returns {Promise<{headers: string[], templateRange: Excel.Range, templateStructure: object, wasExpanded: boolean}>} Enhanced template analysis results.
 */
export async function analyzeTemplateSmart(context, genAI) {
  const originalRange = context.workbook.getSelectedRange();
  originalRange.load(["rowCount", "columnCount", "address"]);
  await context.sync();

  const wasSingleCell = originalRange.rowCount === 1 && originalRange.columnCount === 1;
  const templateRange = await getSmartTemplateRange(context);

  if (!templateRange.values || templateRange.rowCount === 0 || templateRange.columnCount === 0) {
    throw new Error(
      "The selected range is empty. Please select a valid template area with headers."
    );
  }

  // Log expansion information with detailed debugging
  if (wasSingleCell) {
    console.log(
      `Single cell selection (${originalRange.address}) expanded to template range: ${templateRange.address}`
    );
    console.log(
      `Template range details: ${templateRange.rowCount} rows × ${templateRange.columnCount} columns`
    );
    console.log("Template data preview:", templateRange.values.slice(0, 3)); // Show first 3 rows for debugging
  }

  // CRITICAL: Final safety check - never send single-row ranges to LLM
  if (templateRange.rowCount === 1 && templateRange.columnCount > 3) {
    console.log("🚨 CRITICAL: Single-row range detected at final stage!");
    console.log("This will cause bounds validation errors. Applying emergency expansion...");

    try {
      // Emergency expansion - add 10 rows below the header
      const worksheet = templateRange.worksheet;
      const emergencyRange = worksheet.getRangeByIndexes(
        templateRange.rowIndex,
        templateRange.columnIndex,
        Math.max(10, templateRange.rowCount),
        templateRange.columnCount
      );
      emergencyRange.load(["values", "address", "rowCount", "columnCount"]);
      await context.sync();

      console.log("✅ Emergency expansion applied:", emergencyRange.address);

      // Use the emergency range for analysis
      const { analyzeTemplateStructure } = await import("./gemini.service.js");
      const templateStructure = await analyzeTemplateStructure(
        genAI,
        emergencyRange.values,
        emergencyRange.address
      );

      // Update return values to use emergency range
      const headers = templateStructure.fields.map((field) => field.fieldName);
      console.log("Emergency template structure analysis:", templateStructure);

      return {
        headers,
        templateRange: emergencyRange,
        templateStructure,
        wasExpanded: true,
        originalAddress: wasSingleCell ? originalRange.address : emergencyRange.address,
      };
    } catch (emergencyError) {
      console.error("❌ Emergency expansion failed:", emergencyError);
      // Continue with original range and hope for the best
    }
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

  return {
    headers,
    templateRange,
    templateStructure,
    wasExpanded: wasSingleCell,
    originalAddress: wasSingleCell ? originalRange.address : templateRange.address,
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
