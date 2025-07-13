/* global document */
/**
 * UI Component Module
 * Handles all DOM manipulations and UI state management
 */

import { getState, updateState } from "../state.js";

/**
 * Initialize the UI components and event listeners
 */
export function initializeUI() {
  // Set up mode selector event listeners
  const modeSelectors = document.querySelectorAll('input[name="mode"]');
  modeSelectors.forEach((selector) => {
    selector.addEventListener("change", handleModeChange);
  });

  // Initialize with default mode
  showModeSection(getState().currentMode);
}

/**
 * Handle mode change between population and validation
 * @param {Event} event - Change event from radio button
 */
function handleModeChange(event) {
  const selectedMode = event.target.value;
  updateState({ currentMode: selectedMode });
  showModeSection(selectedMode);
  clearStatus();
}

/**
 * Show the appropriate UI section based on selected mode
 * @param {string} mode - 'population' or 'validation'
 */
function showModeSection(mode) {
  const populationSection = document.getElementById("population-section");
  const validationSection = document.getElementById("validation-section");
  const populationInstructions = document.getElementById("population-instructions");
  const validationInstructions = document.getElementById("validation-instructions");

  if (mode === "population") {
    populationSection.style.display = "block";
    validationSection.style.display = "none";
    populationInstructions.style.display = "block";
    validationInstructions.style.display = "none";
  } else {
    populationSection.style.display = "none";
    validationSection.style.display = "block";
    populationInstructions.style.display = "none";
    validationInstructions.style.display = "block";
  }
}

/**
 * Update the status message
 * @param {string} message - Status message to display
 */
export function updateStatus(message) {
  const statusElement = document.getElementById("status");
  if (statusElement) {
    statusElement.textContent = message;
  }
  updateState({ status: message });
}

/**
 * Clear the status message
 */
export function clearStatus() {
  updateStatus("");
}

/**
 * Get the current API key from the input field
 * @returns {string} API key value
 */
export function getApiKey() {
  const apiKeyInput = document.getElementById("api-key");
  return apiKeyInput ? apiKeyInput.value.trim() : "";
}

/**
 * Get the uploaded file from the appropriate input based on current mode
 * @returns {File|null} Selected file or null
 */
export function getUploadedFile() {
  const mode = getState().currentMode;
  const fileInputId = mode === "population" ? "document-upload" : "pdf-upload";
  const fileInput = document.getElementById(fileInputId);
  return fileInput && fileInput.files.length > 0 ? fileInput.files[0] : null;
}

/**
 * Get the prompt text based on current mode
 * @returns {string} Prompt text
 */
export function getPrompt() {
  const mode = getState().currentMode;
  const promptInputId = mode === "population" ? "extract-prompt" : "parse-prompt";
  const promptInput = document.getElementById(promptInputId);
  return promptInput ? promptInput.value.trim() : "";
}

/**
 * Get the comparison prompt for validation mode
 * @returns {string} Comparison prompt text
 */
export function getComparisonPrompt() {
  const comparePromptInput = document.getElementById("compare-prompt");
  return comparePromptInput ? comparePromptInput.value.trim() : "";
}

/**
 * Validate required inputs based on current mode
 * @returns {object} Validation result with isValid boolean and message
 */
export function validateInputs() {
  const apiKey = getApiKey();
  const file = getUploadedFile();
  const mode = getState().currentMode;

  const missingItems = [];
  if (!apiKey) missingItems.push("Gemini API key");
  if (!file) {
    const fileType = mode === "population" ? "document" : "PDF file";
    missingItems.push(fileType);
  }

  if (missingItems.length > 0) {
    return {
      isValid: false,
      message: `Please provide: ${missingItems.join(" and ")}.`,
    };
  }

  return { isValid: true, message: "" };
}

/**
 * Show an error message
 * @param {string} message - Error message to display
 */
export function showError(message) {
  updateStatus(`Error: ${message}`);
}

/**
 * Get the current mode from the UI
 * @returns {string} Current mode ('population' or 'validation')
 */
export function getCurrentMode() {
  const modeSelector = document.querySelector('input[name="mode"]:checked');
  return modeSelector ? modeSelector.value : "population";
}
