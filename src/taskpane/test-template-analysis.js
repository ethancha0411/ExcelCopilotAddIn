/* global console */
/**
 * Test demonstration for LLM-based template analysis
 * This file shows how the enhanced system works with different template formats
 */

import { GoogleGenerativeAI } from "@google/generative-ai";
import { analyzeTemplateStructure } from "./services/gemini.service.js";

/**
 * Demo function to test template analysis with sample data
 */
export async function demoTemplateAnalysis() {
  const genAI = new GoogleGenerativeAI("YOUR_API_KEY_HERE");

  // Test Case 1: Vertical Template (like the user's transposed format)
  console.log("=== Testing Vertical Template Analysis ===");

  const verticalTemplate = [
    ["Field", "Value"],
    ["Property Name", ""],
    ["Property Address", ""],
    ["Property City", ""],
    ["Property State", ""],
    ["Property Zip", ""],
    ["Tenant Name", ""],
    ["Tenant Unit", ""],
    ["Account Number", ""],
    ["Statement Period", ""],
    ["Statement Date", ""],
    ["Date Due", ""],
    ["Previous Balance", ""],
    ["Current Charges", ""],
    ["Total Due", ""],
    ["Make Checks Payable To", ""],
    ["Remit To", ""],
    ["Remit To Address", ""],
  ];

  try {
    const verticalAnalysis = await analyzeTemplateStructure(genAI, verticalTemplate, "A1:B18");
    console.log("Vertical Template Analysis Result:", JSON.stringify(verticalAnalysis, null, 2));

    // Expected output for vertical template:
    // {
    //   "orientation": "vertical",
    //   "fields": [
    //     {
    //       "fieldName": "Property Name",
    //       "fieldLocation": { "row": 1, "col": 0 },
    //       "valueLocation": { "row": 1, "col": 1 },
    //       "description": "Name of the property",
    //       "dataType": "string"
    //     },
    //     ... more fields
    //   ]
    // }
  } catch (error) {
    console.error("Vertical template analysis failed:", error);
  }

  // Test Case 2: Horizontal Template (traditional format)
  console.log("\n=== Testing Horizontal Template Analysis ===");

  const horizontalTemplate = [
    ["Property Name", "Tenant Name", "Account Number", "Total Due", "Date Due"],
    ["", "", "", "", ""],
    ["", "", "", "", ""],
  ];

  try {
    const horizontalAnalysis = await analyzeTemplateStructure(genAI, horizontalTemplate, "A1:E3");
    console.log(
      "Horizontal Template Analysis Result:",
      JSON.stringify(horizontalAnalysis, null, 2)
    );

    // Expected output for horizontal template:
    // {
    //   "orientation": "horizontal",
    //   "fields": [
    //     {
    //       "fieldName": "Property Name",
    //       "fieldLocation": { "row": 0, "col": 0 },
    //       "valueLocation": { "row": 1, "col": 0 },
    //       "description": "Name of the property",
    //       "dataType": "string"
    //     },
    //     ... more fields
    //   ]
    // }
  } catch (error) {
    console.error("Horizontal template analysis failed:", error);
  }
}

/**
 * Demo function showing how the enhanced data mapping works
 */
export async function demoDataMapping() {
  console.log("\n=== Testing Enhanced Data Mapping ===");

  // Sample extracted data from a document
  const extractedData = {
    propertyName: "The Forge Apartments",
    propertyAddress: "123 Main Street",
    city: "Long Island City",
    state: "NY",
    zipCode: "11101",
    tenantName: "JUNG MIN CHOI",
    unit: "44-28",
    accountNumber: "Purves 5/1/24",
    statementPeriod: "May 2024",
    statementDate: "5/1/24",
    dueDate: "5/1/24",
    previousBalance: "3304.29",
    currentCharges: "3304.29",
    totalDue: "3304.29",
  };

  // Sample template structure (vertical format)
  const templateStructure = {
    orientation: "vertical",
    fields: [
      {
        fieldName: "Property Name",
        fieldLocation: { row: 1, col: 0 },
        valueLocation: { row: 1, col: 1 },
        description: "Name of the property",
        dataType: "string",
      },
      {
        fieldName: "Property Address",
        fieldLocation: { row: 2, col: 0 },
        valueLocation: { row: 2, col: 1 },
        description: "Physical address of the property",
        dataType: "string",
      },
      {
        fieldName: "Property City",
        fieldLocation: { row: 3, col: 0 },
        valueLocation: { row: 3, col: 1 },
        description: "City where the property is located",
        dataType: "string",
      },
      {
        fieldName: "Total Due",
        fieldLocation: { row: 14, col: 0 },
        valueLocation: { row: 14, col: 1 },
        description: "Total amount due for payment",
        dataType: "number",
      },
    ],
  };

  console.log("Sample extracted data:", extractedData);
  console.log("Sample template structure:", templateStructure);

  // The enhanced mapping would use this context to create more accurate mappings
  console.log("\nExpected mapping result:");
  console.log([
    {
      "Property Name": "The Forge Apartments",
      "Property Address": "123 Main Street",
      "Property City": "Long Island City",
      "Total Due": "3304.29",
    },
  ]);
}

/**
 * Demo function to test template range offset handling
 */
export async function demoTemplateOffsetHandling() {
  console.log("\n=== Testing Template Range Offset Handling ===");

  // Test Case 1: Template starting at A1 (row=0, col=0)
  console.log("\n📍 Test Case 1: Template at A1 (row=0, col=0)");
  const templateAtA1 = [
    ["Field", "Value"],
    ["Property Name", ""],
    ["Total Due", ""],
  ];

  const mockTemplateRangeA1 = {
    rowIndex: 0, // Template starts at row 1 (0-based)
    columnIndex: 0, // Template starts at column A (0-based)
    address: "A1:B3",
  };

  console.log("Template data:", templateAtA1);
  console.log("Template range:", mockTemplateRangeA1);

  // Simulate LLM analysis result
  const analysisA1 = {
    orientation: "vertical",
    fields: [
      {
        fieldName: "Property Name",
        fieldLocation: { row: 1, col: 0 }, // Relative to template data
        valueLocation: { row: 1, col: 1 }, // Relative to template data
        description: "Name of the property",
        dataType: "string",
      },
      {
        fieldName: "Total Due",
        fieldLocation: { row: 2, col: 0 }, // Relative to template data
        valueLocation: { row: 2, col: 1 }, // Relative to template data
        description: "Total amount due",
        dataType: "number",
      },
    ],
  };

  console.log("LLM Analysis result:", analysisA1);

  // Calculate absolute positions
  console.log("\n🧮 Absolute Position Calculations:");
  analysisA1.fields.forEach((field) => {
    const absoluteRow = mockTemplateRangeA1.rowIndex + field.valueLocation.row;
    const absoluteCol = mockTemplateRangeA1.columnIndex + field.valueLocation.col;
    console.log(`${field.fieldName}:`);
    console.log(`  - Relative: (${field.valueLocation.row}, ${field.valueLocation.col})`);
    console.log(
      `  - Absolute: (${absoluteRow}, ${absoluteCol}) -> Excel cell ${String.fromCharCode(65 + absoluteCol)}${absoluteRow + 1}`
    );
  });

  // Test Case 2: Template starting at D5 (row=4, col=3)
  console.log("\n📍 Test Case 2: Template at D5 (row=4, col=3)");
  const templateAtD5 = [
    ["Field", "Value"],
    ["Property Name", ""],
    ["Total Due", ""],
  ];

  const mockTemplateRangeD5 = {
    rowIndex: 4, // Template starts at row 5 (0-based)
    columnIndex: 3, // Template starts at column D (0-based)
    address: "D5:E7",
  };

  console.log("Template data:", templateAtD5);
  console.log("Template range:", mockTemplateRangeD5);

  // Same LLM analysis result (relative positions don't change)
  const analysisD5 = {
    orientation: "vertical",
    fields: [
      {
        fieldName: "Property Name",
        fieldLocation: { row: 1, col: 0 }, // Still relative to template data
        valueLocation: { row: 1, col: 1 }, // Still relative to template data
        description: "Name of the property",
        dataType: "string",
      },
      {
        fieldName: "Total Due",
        fieldLocation: { row: 2, col: 0 }, // Still relative to template data
        valueLocation: { row: 2, col: 1 }, // Still relative to template data
        description: "Total amount due",
        dataType: "number",
      },
    ],
  };

  console.log("LLM Analysis result:", analysisD5);

  // Calculate absolute positions with offset
  console.log("\n🧮 Absolute Position Calculations with Offset:");
  analysisD5.fields.forEach((field) => {
    const absoluteRow = mockTemplateRangeD5.rowIndex + field.valueLocation.row;
    const absoluteCol = mockTemplateRangeD5.columnIndex + field.valueLocation.col;
    console.log(`${field.fieldName}:`);
    console.log(`  - Relative: (${field.valueLocation.row}, ${field.valueLocation.col})`);
    console.log(
      `  - Absolute: (${absoluteRow}, ${absoluteCol}) -> Excel cell ${String.fromCharCode(65 + absoluteCol)}${absoluteRow + 1}`
    );
  });

  // Test Case 3: Edge case - Template at Z100
  console.log("\n📍 Test Case 3: Edge case - Template at Z100");
  const mockTemplateRangeZ100 = {
    rowIndex: 99, // Template starts at row 100 (0-based)
    columnIndex: 25, // Template starts at column Z (0-based)
    address: "Z100:AA102",
  };

  console.log("Template range:", mockTemplateRangeZ100);

  // Calculate absolute positions for extreme case
  console.log("\n🧮 Absolute Position Calculations for Edge Case:");
  analysisD5.fields.forEach((field) => {
    const absoluteRow = mockTemplateRangeZ100.rowIndex + field.valueLocation.row;
    const absoluteCol = mockTemplateRangeZ100.columnIndex + field.valueLocation.col;
    const colName =
      absoluteCol < 26
        ? String.fromCharCode(65 + absoluteCol)
        : "A" + String.fromCharCode(65 + absoluteCol - 26);
    console.log(`${field.fieldName}:`);
    console.log(`  - Relative: (${field.valueLocation.row}, ${field.valueLocation.col})`);
    console.log(
      `  - Absolute: (${absoluteRow}, ${absoluteCol}) -> Excel cell ${colName}${absoluteRow + 1}`
    );
  });

  console.log("\n✅ Template offset handling test completed!");
  console.log("\n🔑 Key Insights:");
  console.log("1. ✅ LLM returns positions RELATIVE to template data array");
  console.log(
    "2. ✅ We add templateRange.rowIndex + templateRange.columnIndex for absolute positioning"
  );
  console.log("3. ✅ This works regardless of where the template is located in the sheet");
  console.log("4. ✅ Template at A1 vs D5 vs Z100 all work the same way");
}

/**
 * Demo function showing potential issues and solutions
 */
export async function demoCommonOffsetIssues() {
  console.log("\n=== Common Template Offset Issues and Solutions ===");

  console.log("\n❌ ISSUE 1: LLM returns Excel cell addresses instead of array indices");
  console.log("Bad LLM response: valueLocation = 'B2' (Excel address)");
  console.log("✅ Solution: Enhanced prompt explicitly requires array indices");

  console.log("\n❌ ISSUE 2: LLM returns absolute sheet positions");
  console.log("Bad LLM response: valueLocation = {row: 6, col: 4} (absolute position)");
  console.log("✅ Solution: Validation checks ensure positions are within template bounds");

  console.log("\n❌ ISSUE 3: Forgetting to add template range offset");
  console.log(
    "Bad code: cell = worksheet.getCell(field.valueLocation.row, field.valueLocation.col)"
  );
  console.log(
    "✅ Solution: Always add offset - getCell(templateRange.rowIndex + field.valueLocation.row, ...)"
  );

  console.log("\n❌ ISSUE 4: Template range not loaded properly");
  console.log("Bad code: Using templateRange without loading rowIndex/columnIndex");
  console.log("✅ Solution: templateRange.load(['values', 'rowIndex', 'columnIndex'])");

  console.log("\n🔧 Best Practices:");
  console.log("1. Always validate LLM positions are within template bounds");
  console.log("2. Use detailed logging to debug position calculations");
  console.log("3. Test with templates in different sheet locations");
  console.log("4. Handle edge cases (empty cells, merged cells, etc.)");
}

/**
 * Enhanced main demo function
 */
export async function runTemplateAnalysisDemo() {
  console.log("🚀 Starting Enhanced Template Analysis Demo\n");

  await demoTemplateAnalysis();
  await demoDataMapping();
  await demoTemplateOffsetHandling();
  await demoCommonOffsetIssues();

  console.log("\n✅ Complete demo finished!");
  console.log("\n🎯 Summary of Template Range Offset Handling:");
  console.log("✅ LLM analyzes template structure with relative positions");
  console.log("✅ System adds template range offset for absolute positioning");
  console.log("✅ Works with templates anywhere in the Excel sheet");
  console.log("✅ Includes validation and detailed debugging");
  console.log("✅ Handles both vertical and horizontal template formats");
}

// Uncomment to run the demo
// runTemplateAnalysisDemo();
