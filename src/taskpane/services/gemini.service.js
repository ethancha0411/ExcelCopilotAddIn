/* global console, FileReader */
import { SchemaType } from "@google/generative-ai";

/**
 * Converts a File object to a GoogleGenerativeAI.Part object.
 * Supports various document and image MIME types.
 * @param {File} file The file to convert.
 * @returns {Promise<object>} A promise that resolves with the generative part object.
 */
async function fileToGenerativePart(file) {
  const base64EncodedData = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result.split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });

  return {
    inlineData: {
      data: base64EncodedData,
      mimeType: file.type,
    },
  };
}

/**
 * Parses JSON response from Gemini, handling common formatting issues
 * @param {string} text - Raw text response from Gemini
 * @returns {object} Parsed JSON object
 */
function parseGeminiResponse(text) {
  try {
    const cleanedText = text
      .replace(/```json/g, "")
      .replace(/```/g, "")
      .trim();
    return JSON.parse(cleanedText);
  } catch (error) {
    console.error("Failed to parse JSON from Gemini response:", error, text);
    throw new Error(
      "Could not parse structured data from the document. The LLM returned an invalid format."
    );
  }
}

/**
 * Calls the Gemini API to extract structured data from a document.
 * @param {GoogleGenerativeAI} genAI - The GoogleGenerativeAI instance.
 * @param {File} file - The document file (PDF, DOCX, PNG, JPG).
 * @param {string} prompt - The prompt to guide the extraction.
 * @returns {Promise<object>} A promise that resolves with the extracted JSON object.
 */
export async function extractDataFromDocument(genAI, file, prompt) {
  // The new Gemini 1.5 models can handle various document types directly.
  const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
  const imagePart = await fileToGenerativePart(file);

  const fullPrompt = `${prompt}. Respond with only the JSON object, without any markdown formatting.`;

  const result = await model.generateContent([fullPrompt, imagePart]);
  const response = await result.response;
  const text = response.text();

  return parseGeminiResponse(text);
}

/**
 * Parses a document using Gemini with custom prompt
 * @param {GoogleGenerativeAI} genAI - The GoogleGenerativeAI instance
 * @param {File} file - The document file
 * @param {string} prompt - The parsing prompt
 * @param {object} schema - Optional schema for structured output
 * @returns {Promise<object>} Parsed data object
 */
export async function parseDocument(genAI, file, prompt, schema = null) {
  const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });
  const imagePart = await fileToGenerativePart(file);
  const fullPrompt = `${prompt}. ${schema ? "Extract the data according to the specified structure." : "Respond with only the JSON object, without any markdown formatting."}`;

  let result;
  if (schema) {
    result = await model.generateContent({
      contents: [{ parts: [{ text: fullPrompt }, imagePart] }],
      generationConfig: {
        responseMimeType: "application/json",
        responseSchema: schema,
      },
    });
  } else {
    result = await model.generateContent([fullPrompt, imagePart]);
  }

  const response = await result.response;
  const text = response.text();

  if (schema) {
    console.log("Structured document parsing response:", text);
    return JSON.parse(text);
  } else {
    return parseGeminiResponse(text);
  }
}

/**
 * Compares two datasets using Gemini with structured output
 * @param {GoogleGenerativeAI} genAI - The GoogleGenerativeAI instance
 * @param {object} pdfData - Data extracted from PDF
 * @param {Array} excelData - Data from Excel sheet
 * @param {string} prompt - Comparison prompt
 * @returns {Promise<Array>} Array of mismatch objects
 */
export async function compareData(genAI, pdfData, excelData, prompt) {
  const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });

  // Define schema for mismatch objects
  const mismatchSchema = {
    type: SchemaType.OBJECT,
    properties: {
      row: { type: SchemaType.INTEGER },
      col: { type: SchemaType.INTEGER },
      expectedValue: { type: SchemaType.STRING },
      actualValue: { type: SchemaType.STRING },
    },
    propertyOrdering: ["row", "col", "expectedValue", "actualValue"],
  };

  const responseSchema = {
    type: SchemaType.ARRAY,
    items: mismatchSchema,
  };

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
    8.  Each object in the array must have four keys: 
        - "row" (the 0-based row index of the mismatch in Dataset 2)
        - "col" (the 0-based column index)
        - "expectedValue" (the correct value from Dataset 1)
        - "actualValue" (the incorrect value found in Dataset 2)
    9.  If there are no mismatches, return an empty array [].
    **Dataset 1 (from PDF):**
    ${JSON.stringify(pdfData, null, 2)}
    **Dataset 2 (from Excel):**
    ${JSON.stringify(excelData, null, 2)}
    
    Compare the datasets and identify mismatches.
  `;

  const result = await model.generateContent({
    contents: [{ parts: [{ text: fullPrompt }] }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: responseSchema,
    },
  });

  const response = await result.response;
  const text = response.text();

  console.log("Structured comparison response:", text);
  return JSON.parse(text);
}

/**
 * Uses Gemini to automatically map extracted data to Excel template headers with structured output
 * @param {GoogleGenerativeAI} genAI - The GoogleGenerativeAI instance
 * @param {object|Array} extractedData - Data extracted from document
 * @param {string[]} templateHeaders - Headers from Excel template
 * @param {object} templateStructure - Optional template structure for enhanced mapping
 * @returns {Promise<Array>} Array of mapped row objects ready for Excel
 */
export async function mapDataToTemplate(
  genAI,
  extractedData,
  templateHeaders,
  templateStructure = null
) {
  const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });

  // Create a dynamic schema based on the template headers
  const rowSchema = {
    type: SchemaType.OBJECT,
    properties: {},
    propertyOrdering: templateHeaders,
  };

  // Add each template header as a required string property
  templateHeaders.forEach((header) => {
    rowSchema.properties[header] = { type: SchemaType.STRING };
  });

  // Define the response schema as an array of row objects
  const responseSchema = {
    type: SchemaType.ARRAY,
    items: rowSchema,
  };

  console.log("Generated schema for template mapping:", JSON.stringify(responseSchema, null, 2));

  // Build enhanced template context if structure is available
  let templateContext = "";
  if (templateStructure && templateStructure.fields) {
    templateContext = `
**TEMPLATE STRUCTURE CONTEXT:**
- Template Orientation: ${templateStructure.orientation}
- Field Details:
${templateStructure.fields
  .map((field) => `  • ${field.fieldName}: ${field.description} (${field.dataType})`)
  .join("\n")}
`;
  }

  const prompt = `
You are an Excel template data mapping specialist. Your task is to populate an existing Excel template with extracted document data.

**CRITICAL INSTRUCTIONS:**
1. The Excel template has predefined headers that MUST NOT be changed or overwritten.
2. You must respect the exact template structure and only populate VALUES under the existing headers.
3. Your job is to find the most appropriate data from the extracted document for each template field.
4. Use semantic understanding to match document data to template fields (e.g., "Property Name" could match "propertyName", "building_name", "property_title", etc.).
5. If the extracted data is a single object, create one row. If it's an array, create multiple rows.
6. For each row, create an object where keys are EXACTLY the template headers and values are the corresponding data from the document.
7. If no suitable data exists for a header, use an empty string "".
8. Ensure all values are simple types (string, number, boolean) - convert objects/arrays to readable strings.
9. Be intelligent about data types - dates should be formatted properly, numbers should be numbers, etc.
10. Pay attention to field descriptions and data types to make more accurate mappings.

**TEMPLATE HEADERS (these are the ONLY allowed keys):**
${JSON.stringify(templateHeaders)}

${templateContext}

**DOCUMENT DATA TO MAP:**
${JSON.stringify(extractedData, null, 2)}

**ENHANCED MAPPING EXAMPLES:**
- "Property Name" header might map to extractedData.propertyName, extractedData.building_name, or extractedData.property_title
- "Property Address" header might map to extractedData.address, extractedData.property_address, or extractedData.street_address
- "Tenant Name" header might map to extractedData.tenantName, extractedData.tenant_info.name, or extractedData.customer_name
- "Statement Date" header might map to extractedData.statementDate, extractedData.date, or extractedData.issued_date
- "Total Due" header might map to extractedData.totalDue, extractedData.amount_due, or extractedData.balance_due
- "Account Number" header might map to extractedData.accountNumber, extractedData.account_id, or extractedData.customer_account
- "Property City" header might map to extractedData.city, extractedData.property_city, or extractedData.location_city
- "Property State" header might map to extractedData.state, extractedData.property_state, or extractedData.location_state
- "Property Zip" header might map to extractedData.zip, extractedData.zipcode, or extractedData.postal_code

**IMPORTANT NOTES:**
- Use the field descriptions and data types from the template structure to guide your mapping decisions
- For date fields, format dates consistently (e.g., "MM/DD/YYYY" or "YYYY-MM-DD")
- For numeric fields, ensure proper number formatting without currency symbols unless specified
- For text fields, keep values concise and readable

Map the document data to create rows that match the template structure exactly.
`;

  const result = await model.generateContent({
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: responseSchema,
    },
  });

  const response = await result.response;
  const text = response.text();

  console.log("Structured output response:", text);

  const parsedResponse = JSON.parse(text);

  // Validate the response structure
  if (!Array.isArray(parsedResponse)) {
    throw new Error("Expected array response from structured output");
  }

  // Validate each row has only the expected headers
  parsedResponse.forEach((row, index) => {
    const rowKeys = Object.keys(row);
    const unexpectedKeys = rowKeys.filter((key) => !templateHeaders.includes(key));
    if (unexpectedKeys.length > 0) {
      console.warn(`Row ${index} contains unexpected keys: ${unexpectedKeys.join(", ")}`);
    }

    const missingKeys = templateHeaders.filter((header) => !(header in row));
    if (missingKeys.length > 0) {
      console.warn(`Row ${index} missing expected keys: ${missingKeys.join(", ")}`);
    }
  });

  return parsedResponse;
}

/**
 * Uses Gemini to analyze Excel template structure and extract field information
 * @param {GoogleGenerativeAI} genAI - The GoogleGenerativeAI instance
 * @param {Array} templateData - 2D array of template data from Excel
 * @param {string} templateAddress - Excel range address for reference
 * @returns {Promise<object>} Template analysis with field details
 */
export async function analyzeTemplateStructure(genAI, templateData, templateAddress) {
  const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });

  // Define schema for template field analysis
  const fieldSchema = {
    type: SchemaType.OBJECT,
    properties: {
      fieldName: { type: SchemaType.STRING },
      fieldLocation: {
        type: SchemaType.OBJECT,
        properties: {
          row: { type: SchemaType.INTEGER },
          col: { type: SchemaType.INTEGER },
        },
      },
      valueLocation: {
        type: SchemaType.OBJECT,
        properties: {
          row: { type: SchemaType.INTEGER },
          col: { type: SchemaType.INTEGER },
        },
      },
      description: { type: SchemaType.STRING },
      dataType: { type: SchemaType.STRING },
    },
    propertyOrdering: ["fieldName", "fieldLocation", "valueLocation", "description", "dataType"],
  };

  const responseSchema = {
    type: SchemaType.OBJECT,
    properties: {
      orientation: { type: SchemaType.STRING },
      fields: {
        type: SchemaType.ARRAY,
        items: fieldSchema,
      },
    },
    propertyOrdering: ["orientation", "fields"],
  };

  const prompt = `
You are an Excel template structure analyzer. Your task is to analyze the given Excel template data and extract detailed information about each field and its structure.

**TEMPLATE DATA (2D Array):**
${JSON.stringify(templateData, null, 2)}

**TEMPLATE RANGE ADDRESS:** ${templateAddress}

🔴 **CRITICAL RULE: VALUE LOCATION CANNOT BE IDENTICAL TO FIELD LOCATION** 🔴
- The valueLocation MUST be different from fieldLocation
- Values should be placed in ADJACENT cells, NOT on top of field names
- Field names must be preserved and never overwritten
- This is the most important rule to prevent data corruption

**CRITICAL COORDINATE SYSTEM RULES:**
🔴 **IMPORTANT**: All row and column indices must be 0-based and RELATIVE to the template data array provided above.
- The template data array starts at index [0][0] (first row, first column)
- DO NOT use Excel cell addresses (like A1, B2) or absolute sheet positions
- Use only array indices relative to the template data (0, 1, 2, etc.)

**ANALYSIS INSTRUCTIONS:**
1. **Orientation Detection**: Determine if this is a "horizontal" template (fields are column headers) or "vertical" template (fields are in rows).

2. **Field Analysis**: For each data field in the template, extract:
   - **fieldName**: The actual field name (e.g., "Property Name", "Tenant Name", "Total Due")
   - **fieldLocation**: The 0-based row and column indices where the field name appears IN THE TEMPLATE DATA ARRAY
   - **valueLocation**: The 0-based row and column indices where the value should be populated IN THE TEMPLATE DATA ARRAY
   - **description**: Brief description of what this field represents
   - **dataType**: Expected data type ("string", "number", "date", "boolean")

**VALUE PLACEMENT LOGIC:**
🔴 **MANDATORY**: valueLocation MUST be different from fieldLocation
- **Vertical Templates**: Values go in the next column (col + 1) from the field name
- **Horizontal Templates**: Values go in the next row (row + 1) from the field header
- **NEVER place values on the same position as field names**

**COORDINATE SYSTEM EXAMPLES:**

**Vertical Template Example:**
If the template data array is:
\`\`\`
[
  ["Field", "Value"],           <- Array row 0
  ["Property Name", ""],        <- Array row 1  
  ["Property Address", ""],     <- Array row 2
  ["Total Due", ""]             <- Array row 3
]
\`\`\`

Then the correct analysis would be:
- "Property Name": fieldLocation={row:1,col:0}, valueLocation={row:1,col:1} ✅ DIFFERENT
- "Property Address": fieldLocation={row:2,col:0}, valueLocation={row:2,col:1} ✅ DIFFERENT
- "Total Due": fieldLocation={row:3,col:0}, valueLocation={row:3,col:1} ✅ DIFFERENT

**INCORRECT Example (DO NOT DO THIS):**
- "Property Name": fieldLocation={row:1,col:0}, valueLocation={row:1,col:0} ❌ SAME LOCATION - FORBIDDEN

**Horizontal Template Example:**
If the template data array is:
\`\`\`
[
  ["Property Name", "Property Address", "Total Due"],  <- Array row 0 (headers)
  ["", "", ""],                                        <- Array row 1 (values go here)
  ["", "", ""]                                         <- Array row 2 (additional rows)
]
\`\`\`

Then the correct analysis would be:
- "Property Name": fieldLocation={row:0,col:0}, valueLocation={row:1,col:0} ✅ DIFFERENT (next row)
- "Property Address": fieldLocation={row:0,col:1}, valueLocation={row:1,col:1} ✅ DIFFERENT (next row)
- "Total Due": fieldLocation={row:0,col:2}, valueLocation={row:1,col:2} ✅ DIFFERENT (next row)

**INCORRECT Example (DO NOT DO THIS):**
- "Property Name": fieldLocation={row:0,col:0}, valueLocation={row:0,col:0} ❌ SAME LOCATION - FORBIDDEN

**ANALYSIS RULES:**
- Skip header rows that contain generic labels like "Field", "Value", "Header", etc.
- Look for actual field names that describe data (Property Name, Tenant Name, etc.)
- For vertical templates, values typically go in the column next to the field name (same row, next column)
- For horizontal templates, values typically go in rows below the field headers (next row, same column)
- Ignore empty cells or cells with just whitespace
- Infer data types from field names (e.g., "Date" = date, "Amount" = number, "Name" = string)
- ALL POSITIONS MUST BE RELATIVE TO THE TEMPLATE DATA ARRAY (0-based indices)
- **CRITICAL**: ALWAYS ensure valueLocation != fieldLocation to prevent overwriting field names

**VALIDATION CHECKS:**
- Ensure all row indices are between 0 and ${templateData.length - 1}
- Ensure all column indices are between 0 and ${templateData[0]?.length - 1 || 0}
- fieldLocation and valueLocation must be valid array indices
- **MANDATORY**: fieldLocation and valueLocation must be DIFFERENT positions
- Values should be placed logically adjacent to field names, not on top of them

**EXPECTED OUTPUT:**
Return a JSON object with:
- "orientation": "vertical" or "horizontal"
- "fields": Array of field objects with the structure defined above

Analyze the template carefully and provide comprehensive field information with correct relative positioning, ensuring values never overwrite field names.
`;

  const result = await model.generateContent({
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: responseSchema,
    },
  });

  const response = await result.response;
  const text = response.text();

  console.log("Template structure analysis response:", text);

  const parsedResponse = JSON.parse(text);

  // Validate the response structure
  if (!parsedResponse.orientation || !Array.isArray(parsedResponse.fields)) {
    throw new Error("Invalid template analysis response structure");
  }

  // Validate that all positions are within bounds of the template data
  const maxRow = templateData.length - 1;
  const maxCol = templateData[0]?.length - 1 || 0;

  console.log("Template bounds validation:");
  console.log(
    `Template data dimensions: ${templateData.length} rows × ${templateData[0]?.length || 0} columns`
  );
  console.log(`Valid bounds: [0-${maxRow}, 0-${maxCol}]`);
  console.log("Template data preview:", templateData.slice(0, 3));

  parsedResponse.fields.forEach((field, index) => {
    const { fieldLocation, valueLocation } = field;

    console.log(`Validating field ${index} (${field.fieldName}):`);
    console.log(`  fieldLocation: {row:${fieldLocation.row}, col:${fieldLocation.col}}`);
    console.log(`  valueLocation: {row:${valueLocation.row}, col:${valueLocation.col}}`);

    // 🔴 CRITICAL: Validate that valueLocation is different from fieldLocation
    if (fieldLocation.row === valueLocation.row && fieldLocation.col === valueLocation.col) {
      console.error(
        `❌ CRITICAL ERROR: Field "${field.fieldName}" has identical fieldLocation and valueLocation!`
      );
      console.error(
        `This would overwrite the field name with data. fieldLocation and valueLocation must be different.`
      );
      throw new Error(
        `Field "${field.fieldName}" has invalid configuration: valueLocation {row:${valueLocation.row}, col:${valueLocation.col}} is identical to fieldLocation {row:${fieldLocation.row}, col:${fieldLocation.col}}. Values cannot be placed on top of field names. This would overwrite the field name and corrupt the template.`
      );
    }

    // Validate field location
    if (
      fieldLocation.row < 0 ||
      fieldLocation.row > maxRow ||
      fieldLocation.col < 0 ||
      fieldLocation.col > maxCol
    ) {
      console.error(`Field location validation failed for field ${index} (${field.fieldName})`);
      console.error(`Template data:`, templateData);
      console.error(`Template address: ${templateAddress}`);
      throw new Error(
        `Field ${index} (${field.fieldName}) has invalid fieldLocation: {row:${fieldLocation.row}, col:${fieldLocation.col}}. Must be within bounds [0-${maxRow}, 0-${maxCol}]. Template data has ${templateData.length} rows and ${templateData[0]?.length || 0} columns.`
      );
    }

    // Validate value location
    if (
      valueLocation.row < 0 ||
      valueLocation.row > maxRow ||
      valueLocation.col < 0 ||
      valueLocation.col > maxCol
    ) {
      console.error(`Value location validation failed for field ${index} (${field.fieldName})`);
      console.error(`Template data:`, templateData);
      console.error(`Template address: ${templateAddress}`);
      throw new Error(
        `Field ${index} (${field.fieldName}) has invalid valueLocation: {row:${valueLocation.row}, col:${valueLocation.col}}. Must be within bounds [0-${maxRow}, 0-${maxCol}]. Template data has ${templateData.length} rows and ${templateData[0]?.length || 0} columns.`
      );
    }

    console.log(`  ✅ Field validation passed - locations are different and within bounds`);
  });

  console.log("✅ All fields validated successfully - no field/value location conflicts detected");
  return parsedResponse;
}
