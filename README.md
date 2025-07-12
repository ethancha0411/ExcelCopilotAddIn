# Excel Copilot Add-in 

This Excel add-in uses the power of Google's Gemini models to automate the tedious process of data verification. It can parse structured data from a PDF document and compare it against a selected data range within your Excel worksheet, highlighting any discrepancies it finds.

## Features

- **Intelligent PDF Parsing**: Uses the `gemini-2.5-flash` model to extract structured data from PDF documents.
- **Automated Data Comparison**: Leverages a Gemini text model to intelligently compare the extracted PDF data against your selected Excel data.
- **Mismatch Highlighting**: Automatically colors the cells in your selection that contain data inconsistent with the PDF.
- **Enhanced Comment System**: Adds detailed comments to mismatched cells showing both expected and actual values for easy comparison.
- **Environment Variable Support**: Securely store your Gemini API key in a `.env` file to avoid entering it manually each time.
- **Flexible API Key Management**: Choose between using environment variables or manual entry for your API key.
- **Parsed Data Output**: Creates a new worksheet containing the structured data extracted from the PDF for your review.
- **Theme-Aware UI**: The user interface adapts to your Office theme, including full support for dark mode.
- **Customizable Prompts**: While the add-in provides effective default prompts, you can provide your own instructions to guide the parsing and comparison logic.

## Prerequisites

- [Node.js](https://nodejs.org/) and npm
- Microsoft Excel (for Windows or macOS, compatible with Office Add-ins)

## Installation

1.  **Clone the repository or download the source code.**
2.  **Navigate to the project directory in your terminal:**
    ```bash
    cd /path/to/excel-copilot-addin
    ```
3.  **Install the required dependencies:**
    ```bash
    npm install
    ```
4.  **Set up environment variables (recommended):**
    - Copy `.env.example` to `.env`:
      ```bash
      cp .env.example .env
      ```
    - Open `.env` and replace `your_gemini_api_key_here` with your actual Gemini API key
    - Get your API key from [Google AI Studio](https://aistudio.google.com/app/apikey)

## Running the Add-in

1.  **Start the development server:**
    ```bash
    npm start
    ```
    This command will build the project, start a local server, and attempt to open and sideload the add-in in Excel.

2.  **Sideload the Add-in (if needed):**
    :warning: Caution: This is still being tested.
    If Excel doesn't open automatically, or if the add-in doesn't appear, you can load it manually:
    - Go to **Insert** > **My Add-ins**.
    - Click the **...** menu in the top right, select **Upload My Add-in**, and choose the `manifest.xml` file from the project's root directory.

## How to Use

1.  **Prepare Your Excel Data**: Create a table in your worksheet that corresponds to the data in your PDF document.
2.  **Open the Add-in**: Go to the **Home** tab in the Excel ribbon and click the **Data Verifier** button to open the task pane.
3.  **API Key Setup**: 
    - **Option 1 (Recommended)**: Set up your API key in the `.env` file during installation - the add-in will automatically use it
    - **Option 2**: Manually enter your API key in the task pane input field each time
    - You can get a key from [Google AI Studio](https://aistudio.google.com/app/apikey)
4.  **Provide Prompts (Optional)**:
    - **PDF Parsing Prompt**: You can provide specific instructions for extracting data from the PDF. If you leave this blank, a general-purpose extraction prompt will be used.
    - **Comparison Prompt**: You can specify how the PDF data should be compared to the Excel columns. If left blank, a default comparison logic will be used.
5.  **Select Your Data**: In your worksheet, select the entire data range you want to verify, including the headers.
6.  **Select a PDF**: In the task pane, click to upload the PDF document you want to compare against.
7.  **Run Verification**: Click the **Verify Selected Data** button.

The add-in will now perform the verification. Upon completion:
- A new worksheet (e.g., `PDF_YourFileName`) will be created containing the data extracted from the PDF.
- Any cells in your original selection that do not match the PDF data will be highlighted in red.
- The status panel at the bottom will show the final result.

## Customizing the Default Prompts

You can change the default prompts used by the add-in by editing the `src/taskpane/taskpane.js` file. Find the `verifyData` function and modify the strings for `parsePrompt` and `comparePrompt` within this block:

```javascript
// src/taskpane/taskpane.js

// ... inside the verifyData function
    if (!parsePrompt.trim()) {
      parsePrompt = "Extract all key-value pairs from the document. The keys should be in camelCase. Return the result as a single, flat JSON object.";
    }
    if (!comparePrompt.trim()) {
      comparePrompt = "Compare the data from the PDF (Dataset 1) with the data from Excel (Dataset 2). Match the keys from the PDF data to the header columns in the Excel data, ignoring case and special characters. Identify any cells in the Excel data that do not match the corresponding PDF data.";
    }
// ...
```

## ⚠️ Security Notice

For simplicity and development purposes, this add-in handles the Gemini API key on the client-side (in the browser). **This is not a secure practice for a production environment.** If you plan to deploy this application, you must implement a secure backend service to manage and use the API key. The Excel add-in should make requests to your secure backend, not directly to the Google API.

## Troubleshooting

- **Add-in title or icon not updating**: If you change the `manifest.xml` file, you may need to clear Excel's cache. On macOS, you can do this by closing Excel and running the command: `rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/com.microsoft.Osf.loader/Wef/`
- **Viewing Logs**: To see `console.log` messages for debugging, you need to use the browser's developer tools. On macOS, this is the Safari Web Inspector.
