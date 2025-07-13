/**
 * Centralized application state management
 * Handles shared state between Template Population and Data Validation pipelines
 */

// Application state object
const appState = {
  // Common state
  currentMode: "population", // 'population' or 'validation'
  apiKey: "",
  status: "",

  // Template Population state
  extractedData: null,
  templateHeaders: null,
  templateRangeAddress: null,
  templateStructure: null,
  mappedData: null,
  wasRangeExpanded: false,
  originalRangeAddress: null,

  // Data Validation state
  pdfData: null,
  excelData: null,
  mismatches: null,
  selectedRangeAddress: null,
};

/**
 * Get the current application state
 * @returns {object} Current state object
 */
export function getState() {
  return appState;
}

/**
 * Update specific state properties
 * @param {object} updates - Object containing state updates
 */
export function updateState(updates) {
  Object.assign(appState, updates);
}

/**
 * Reset state for a specific pipeline
 * @param {string} pipeline - 'population' or 'validation'
 */
export function resetPipelineState(pipeline) {
  if (pipeline === "population") {
    appState.extractedData = null;
    appState.templateHeaders = null;
    appState.templateRangeAddress = null;
    appState.mappedData = null;
  } else if (pipeline === "validation") {
    appState.pdfData = null;
    appState.excelData = null;
    appState.mismatches = null;
    appState.selectedRangeAddress = null;
  }
}

/**
 * Get state specific to a pipeline
 * @param {string} pipeline - 'population' or 'validation'
 * @returns {object} Pipeline-specific state
 */
export function getPipelineState(pipeline) {
  if (pipeline === "population") {
    return {
      extractedData: appState.extractedData,
      templateHeaders: appState.templateHeaders,
      templateRangeAddress: appState.templateRangeAddress,
      templateStructure: appState.templateStructure,
      mappedData: appState.mappedData,
      wasRangeExpanded: appState.wasRangeExpanded,
      originalRangeAddress: appState.originalRangeAddress,
    };
  } else if (pipeline === "validation") {
    return {
      pdfData: appState.pdfData,
      excelData: appState.excelData,
      mismatches: appState.mismatches,
      selectedRangeAddress: appState.selectedRangeAddress,
      wasRangeExpanded: appState.wasRangeExpanded,
      originalRangeAddress: appState.originalRangeAddress,
    };
  }
  return {};
}

/**
 * Clear all state
 */
export function clearState() {
  appState.extractedData = null;
  appState.templateHeaders = null;
  appState.templateRangeAddress = null;
  appState.mappedData = null;
  appState.pdfData = null;
  appState.excelData = null;
  appState.mismatches = null;
  appState.selectedRangeAddress = null;
  appState.status = "";
}
