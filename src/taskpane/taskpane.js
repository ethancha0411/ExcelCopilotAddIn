/* global console, document, Office */
import {
  initializeUI,
  validateInputs,
  getApiKey,
  getUploadedFile,
  getPrompt,
  getComparisonPrompt,
  showError,
  getCurrentMode,
} from "./components/ui.js";
import { updateState } from "./state.js";
import { executePopulation } from "./pipelines/population.js";
import { executeValidation } from "./pipelines/validation.js";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const sideloadMsg = document.getElementById("sideload-msg");
    if (sideloadMsg) {
      sideloadMsg.style.display = "none";
    }
    document.getElementById("app-body").style.display = "flex";

    // Initialize UI components
    initializeUI();

    // Set up event listeners
    document.getElementById("run-button").onclick = handleRunButton;
  }
});

/**
 * Handle the main run button click
 * Routes to appropriate pipeline based on current mode
 */
async function handleRunButton() {
  try {
    // Validate inputs
    const validation = validateInputs();
    if (!validation.isValid) {
      showError(validation.message);
      return;
    }

    // Get inputs
    const apiKey = getApiKey();
    const file = getUploadedFile();
    const mode = getCurrentMode();

    // Store API key in state
    updateState({ apiKey });

    // Route to appropriate pipeline
    if (mode === "population") {
      const prompt = getPrompt();
      await executePopulation(apiKey, file, prompt);
    } else if (mode === "validation") {
      const parsePrompt = getPrompt();
      const comparePrompt = getComparisonPrompt();
      await executeValidation(apiKey, file, parsePrompt, comparePrompt);
    }
  } catch (error) {
    console.error("Pipeline execution failed:", error);
    showError(error.message);
  }
}
