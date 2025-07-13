# Engineering Design Document: Universal Document-to-Excel Template Populator (MVP)

## 1. Overview

This document outlines the engineering strategy, architecture, and implementation plan for the Minimum Viable Product (MVP) of the **Universal Document-to-Excel Template Populator**. The goal of this Excel add-in is to automate the process of populating existing Excel templates with structured data extracted from various document formats, leveraging multimodal Large Language Models (LLMs).

This plan is based on the provided Product Requirements Document (PRD) and an analysis of the existing "Data Verifier" add-in, focusing on the Phase 1 (MVP) scope.

---

## 2. Current State Analysis

The current codebase is an Excel add-in that functions as a **Data Verifier**. Its primary workflow is:

1.  User provides a PDF and selects a data range in Excel.
2.  The add-in uses the Gemini API to parse the PDF into a structured JSON object.
3.  It then compares this JSON object with the data in the selected Excel range.
4.  Discrepancies are highlighted in the Excel sheet with explanatory comments.

This existing foundation provides:
- A working Excel Add-in structure with a taskpane.
- Established integration with the Google Generative AI SDK.
- A functional example of file handling (`FileReader`) and Office.js API usage for reading from and writing to a worksheet.

While the core business logic will shift from **verification** to **population**, the current setup is a valuable starting point, significantly de-risking the initial project setup and integration with key APIs.

---

## 3. Proposed MVP Architecture

The MVP will be architected as a client-side application running entirely within the Excel taskpane. This aligns with the requirement to keep the frontend and backend within a single application for the initial phase.

### 3.1. High-Level Data Flow

The user workflow will be as follows:

```mermaid
graph TD
    A[Start: User selects Excel template area] --> B{Upload Document <br/> (PDF, DOCX, PNG, JPG)};
    B --> C[1. Document Extraction <br/> (Client-side via Gemini API)];
    C --> D{Extracted Data (JSON)};
    A --> E[2. Template Analysis <br/> (via Office.js)];
    E --> F{Template Structure <br/> (Headers, Fillable Area)};
    D & F --> G[3. Mapping UI <br/> (User confirms/adjusts mapping)];
    G --> H[4. Population Engine <br/> (Writes data to sheet via Office.js)];
    H --> I[End: Template Populated];
```

### 3.2. Component Breakdown

| Component | Technology/API | Responsibility (MVP) | Exists? |
| :--- | :--- | :--- | :--- |
| **Taskpane UI** | HTML, CSS | Rework UI for document upload, mapping, and status updates. | Yes (needs rework) |
| **Orchestrator** | `taskpane.js` | Manage the overall workflow from upload to population. | Yes (needs rework) |
| **Document Extractor** | `@google/generative-ai` | Send document to a multimodal LLM and receive a structured JSON object. | Yes (needs extension) |
| **Template Analyzer** | `Office.js` | Identify header row and the data entry area below it. | No (New) |
| **Mapping Engine** | JS, HTML | Display extracted keys and template headers; allow user to map them. | No (New) |
| **Population Engine**| `Office.js` | Write mapped data into the appropriate cells of the template. | No (New) |

---

## 4. Detailed Design and Implementation Plan

### 4.1. Project Structure (Trade-off)

The existing structure places all logic into `src/taskpane/taskpane.js`. For the MVP's complexity, this will become unmanageable.

- **Option A: Monolith (`taskpane.js`)**
  - **Pros:** Fastest to start.
  - **Cons:** Poor maintainability, difficult to test, high coupling.
- **Option B: Modular Structure**
  - Refactor `src/taskpane` into subdirectories:
    - `components/`: Self-contained UI modules (e.g., `mapping-component.js`).
    - `services/`: Business logic (e.g., `gemini.service.js`, `excel.service.js`).
    - `utils/`: Reusable helper functions.
  - **Pros:** Scalable, maintainable, clear separation of concerns.
  - **Cons:** Requires initial refactoring effort.

**Recommendation:** Pursue **Option B**. The upfront investment in a modular structure will accelerate development and improve quality as features are added.

### 4.2. Document Extraction Service

This service will adapt the existing `callGeminiParse` function.

- **File Support:** The `fileToGenerativePart` function will be updated to handle MIME types for PDF (`application/pdf`), DOCX (`application/vnd.openxmlformats-officedocument.wordprocessingml.document`), and images (`image/jpeg`, `image/png`). The Gemini `1.5-flash` or `1.5-pro` models can process these formats directly.
- **Prompt Engineering:** A new, robust prompt is required. It will instruct the LLM to analyze the document (regardless of type) and return a **single, flat JSON object**. For tabular data within the document, it should be instructed to return an array of objects.
- **Implementation:**
    1. Create `src/taskpane/services/gemini.service.js`.
    2. Move and generalize the PDF parsing logic into an `extractDataFromDocument(file, prompt)` function.
    3. Ensure the function robustly cleans and parses the LLM's string response into a valid JSON object, with enhanced error handling for malformed responses.

### 4.3. Template Analysis Service (MVP)

This new service will use Office.js to understand the user's template.

- **Assumption:** For the MVP, we will assume the user's template is a simple table with headers in the first row of a selected range.
- **Implementation:**
    1. Create `src/taskpane/services/excel.service.js`.
    2. Implement `analyzeTemplate(context)`:
        - Get the user's selected range: `context.workbook.getSelectedRange()`.
        - Load its `values`.
        - Extract the first row as the `headers` array.
        - Identify the "fillable area" as the part of the range below the header row.

### 4.4. Mapping Engine & UI (MVP)

This is a critical new UI component.

- **Automatic Mapping:**
    1. Once data is extracted (JSON keys) and the template is analyzed (headers), perform a simple automatic mapping.
    2. A utility function `mapKeys(jsonKeys, headers)` will compare strings after converting them to lowercase and removing non-alphanumeric characters.
- **UI for Confirmation:**
    1. Create a new component in `src/taskpane/components/`.
    2. For each key from the extracted JSON, display the key and a dropdown list of all available Excel headers.
    3. The dropdown should be pre-set with the best guess from the automatic mapping step.
    4. A "Confirm & Populate" button triggers the final step.

### 4.5. Population Engine

This service will execute the final step of writing data.

- **Implementation (in `excel.service.js`):**
    1. Implement `populateTemplate(context, mapping, extractedData, templateRange)`.
    2. Identify the first empty row within the `templateRange`'s fillable area.
    3. Construct a 2D array representing the new row of data, ordered correctly based on the user-confirmed `mapping`.
    4. Write the entire row in a single operation using `range.values = [newRowData]` for optimal performance.
    5. **Handling tabular data**: If `extractedData` contains an array of objects, the function will loop through it, populating multiple new rows in the sheet.

---

## 5. Risks and Mitigation Strategies

| Risk | Probability | Impact | Mitigation |
| :--- | :--- | :--- | :--- |
| **Inaccurate LLM Extraction** | Medium | High | - **Prompt Engineering:** Iteratively refine the extraction prompt with few-shot examples. <br/>- **Manual Override:** Allow the user to view and edit the extracted JSON in a `<textarea>` before mapping. |
| **Flaky Template Analysis** | High | Medium | - **Clear Documentation:** Advise users that the MVP works best on simple tables with a single header row. <br/>- **User Selection:** Rely on the user selecting the precise table range, including headers. |
| **Client-Side Performance** | Medium | Medium | - **Status Updates:** Implement a clear, multi-step status indicator in the UI. <br/>- **Acknowledge Limits:** Recognize that the PRD's performance goals may not be fully met with a purely client-side MVP. Defer heavy processing to a backend in a future phase. |
| **API Costs** | Medium | Medium | Implement logic to show a confirmation dialog before making the call to the Gemini API, estimating the potential cost if possible. Cache extraction results for a session. |

---

## 6. Out of Scope for MVP

To ensure focus and timely delivery, the following features from the PRD are explicitly **out of scope** for the MVP:

- **Backend Architecture:** A secure backend for API key management and processing is a critical next step but will be deferred. The API key will remain a user-input field in the taskpane.
- **Advanced Template Analysis:** Detecting formulas, merged cells, or non-standard table structures.
- **Multi-Document Processing:** The MVP workflow will handle one document at a time.
- **Data Validation & Normalization:** We will trust the data as it comes from the LLM.
- **"Undo" Feature:** Users will be advised to save their work before running the populator.

This design provides a clear and actionable path to developing a valuable MVP that delivers on the core user story while laying a strong, modular foundation for future enhancements. 