# Engineering Design: Unifying Taskpane Features

**Author:** AI Assistant
**Date:** 2025-07-13
**Status:** Recommended Path Approved

## 1. Overview

The Excel Add-in currently contains two distinct, isolated features within the taskpane: **Template Population** and **Data Validation**. The user interface (`taskpane.html`) is hardcoded for the "Template Population" feature, and switching to the "Data Validation" feature requires manual code changes. This design prevents users from accessing both features, creates code duplication, and hinders future development.

This document proposes a refactoring of the taskpane codebase to create a unified, modular, and scalable architecture. The goal is to allow users to select their desired operation from a single, cohesive UI.

## 2. Current Architecture Analysis

The current implementation suffers from several key issues:

-   **Isolated Features**: `template_populator.js` (Template Population) and `taskpane.js` (Data Validation) operate independently. The UI in `taskpane.html` only serves the population feature, making the validation feature inaccessible without code modifications.
-   **Code Duplication**: Both scripts contain redundant logic for:
    -   Handling API key and file inputs.
    -   Interacting with the Gemini API.
    -   Interacting with the Excel JavaScript API.
-   **Lack of Modularity**: Business logic is tightly coupled with UI manipulation and API calls. For example, `taskpane.js` contains its own data parsing and Excel writing functions instead of leveraging the existing `gemini.service.js` and `excel.service.js`.
-   **Inconsistent State Management**: `template_populator.js` uses a global `appState` object, which is not a robust solution for managing state in a more complex application.

## 3. Proposed Architecture

We will adopt a modular, single-page application (SPA) architecture for the taskpane. This involves a unified UI and a clear separation of concerns in the JavaScript codebase.

### 3.1. UI/UX Design

The `taskpane.html` will be updated to include a mode selector, allowing users to switch between "Populate Template" and "Validate Data" modes.

-   **Mode Selector**: Radio buttons at the top of the taskpane to select the active pipeline.
-   **Dynamic UI Sections**: Each pipeline will have a dedicated `<div>` container for its specific inputs (e.g., prompts, file uploads). These sections will be shown or hidden based on the selected mode.
-   **Common Elements**: The API key input, the main action button (e.g., "Run"), and the status display will be shared across both modes.

### 3.2. Code Structure Refactoring

The JavaScript code will be reorganized into a more modular structure to promote reusability and maintainability.

```
src/taskpane/
├── components/
│   └── ui.js             # Handles all DOM manipulations and UI state
├── pipelines/
│   ├── population.js     # Orchestration logic for the "Populate Template" pipeline
│   └── validation.js     # Orchestration logic for the "Validate Data" pipeline
├── services/
│   ├── excel.service.js  # Unified service for all Excel interactions
│   └── gemini.service.js # Unified service for all Gemini API calls
├── taskpane.css
├── taskpane.html         # The single, unified HTML file
├── taskpane.js           # Main controller/entry point
└── state.js              # Centralized application state management
```

**Key Responsibilities:**

-   **`taskpane.js` (Controller)**: The main entry point. It initializes the add-in, handles UI events (like mode switching), and calls the appropriate pipeline orchestrator based on the current mode.
-   **`state.js`**: A simple module to hold and manage shared application state (e.g., API key, extracted data, current mode).
-   **`components/ui.js`**: A module dedicated to all DOM manipulations. It will contain functions to build UI elements, show/hide sections, and update status messages, completely decoupling the business logic from the view.
-   **`services/`**:
    -   `gemini.service.js`: Consolidates all calls to the Google Generative AI API. It will provide generic functions like `extractData(file, prompt)` and `compareData(data1, data2, prompt)`.
    -   `excel.service.js`: Consolidates all interactions with the Excel API. It will provide functions like `getSelectedData()`, `highlightMismatches()`, and `populateTemplate()`.
-   **`pipelines/`**:
    -   `population.js`: Contains the high-level logic for the template population flow, adapted from `template_populator.js`.
    -   `validation.js`: Contains the high-level logic for the data validation flow, adapted from `taskpane.js`.

## 4. Design Trade-offs

### Option A: Single-Page Application (Recommended)

The proposed architecture is a single-page application (SPA) where UI changes are handled dynamically with JavaScript.

-   **Pros**:
    -   **Seamless User Experience**: No page reloads when switching between modes.
    -   **Efficient State Management**: A single, shared state is easy to manage across features.
    -   **High Code Reusability**: Services and components are shared efficiently.
-   **Cons**:
    -   **Initial Complexity**: The main controller (`taskpane.js`) will be more complex as it needs to manage different UI states.
    -   **Potentially Larger Bundle**: All code is loaded at once, though this is negligible for an application of this scale.

### Option B: Multiple HTML Files

An alternative would be to use separate HTML files for each feature and navigate between them.

-   **Pros**:
    -   **Simpler Logic Per Page**: Each page and its corresponding script are self-contained.
-   **Cons**:
    -   **Poor User Experience**: Requires page navigation within the taskpane, which can be slow and jarring.
    -   **Complex State Sharing**: Sharing state (like the API key) between pages is more difficult.
    -   **Lower Cohesion**: More difficult to maintain a consistent look and feel.

Given the context of an Office Add-in, the SPA approach is strongly recommended for its superior user experience and more maintainable structure.

## 5. Implementation Plan

1.  **Refactor UI (`taskpane.html`)**:
    -   Add radio buttons for mode selection.
    -   Restructure the body into sections for each pipeline and for common elements.

2.  **Create New Modules**:
    -   Create the proposed directory structure.
    -   Create empty files for `state.js`, `ui.js`, `pipelines/population.js`, and `pipelines/validation.js`.

3.  **Consolidate Services**:
    -   Move Gemini-related functions from `taskpane.js` into `services/gemini.service.js` and merge with existing logic.
    -   Move Excel-related functions from `taskpane.js` into `services/excel.service.js` and merge.

4.  **Isolate Pipeline Logic**:
    -   Move the orchestration logic from `template_populator.js` into `pipelines/population.js`.
    -   Move the orchestration logic from the old `taskpane.js` into `pipelines/validation.js`.
    -   Refactor these files to import dependencies from the new `services`, `ui`, and `state` modules.

5.  **Implement Main Controller (`taskpane.js`)**:
    -   Delete the old `template_populator.js`.
    -   Rewrite `taskpane.js` to act as the main controller.
    -   Implement `Office.onReady` to initialize the UI, set default state, and attach event listeners to the mode selector and "Run" button.
    -   Write the event handler logic to call the correct pipeline from the `pipelines/` directory based on the selected mode.

## 6. Future Considerations

-   **Scalability**: This modular design makes it easy to add new pipelines in the future by simply adding a new file to the `pipelines/` directory, updating the UI, and adding a case in the main controller.
-   **Testing**: The separation of concerns allows for easier unit testing. Services and UI components can be tested in isolation.
-   **Error Handling**: A centralized error handling strategy can be implemented in the main controller or as a dedicated utility to provide consistent feedback to the user. 