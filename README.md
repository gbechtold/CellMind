# CellMindAI

**CellMindAI** is an advanced Google Sheets integration that allows you to process and analyze your spreadsheet data using Claude AI. Turn your cells into intelligent insights with prompt-based analysis and multi-step prompt chains.


[How to Loom Video of this CellMindAI](https://script.google.com/home](https://www.loom.com/share/b47788ef647044e8babb68e37bf1fae6?sid=8e5efe9b-8008-4199-9585-9001b6f26c89)


## Features

- **Seamless Google Sheets Integration**: Add AI capabilities directly in your spreadsheets
- **Smart Data Analysis**: Process your spreadsheet data with natural language prompts
- **Prompt Chains**: Build multi-step analyses where each step builds on previous results
- **Flexible Data Referencing**: Reference data across different sheets and ranges
- **No-Code Solution**: Use without any programming knowledge
- **Fallback Mode**: Works even if the library connection fails

## Installation

### Option 1: Library Setup (Recommended)

#### Step 1: Set up the Library

1. Go to [Google Apps Script](https://script.google.com/home)
2. Create a new project and name it "CellMindAILib"
3. Delete the default code in `Code.gs` and paste the entire library code from `CellMindAILib.js`
4. Save the project
5. Go to "Deploy" > "New deployment"
6. Select "Library" as deployment type
7. Enter a description (e.g., "Initial version") and click "Deploy"
8. Copy the Script ID that appears

#### Step 2: Add the Client Code to Your Spreadsheet

1. Open your Google Sheet
2. Go to Extensions > Apps Script
3. Paste the code from `CellMindAIClient.js`
4. In the left sidebar, click on "Libraries" (+)
5. Enter the Script ID you copied in Step 1
6. Set the Identifier to exactly "CellMindAILib" (case sensitive)
7. Select the latest version
8. Click "Add"
9. Save and refresh your spreadsheet

### Option 2: Standalone Setup

If you prefer a simpler setup without library dependencies:

1. Open your Google Sheet
2. Go to Extensions > Apps Script
3. Copy all the code from `CellMindAIStandalone.js` (not included in this repository)
4. Save and refresh your spreadsheet

## Usage

### Basic Usage

1. After installation, you'll see a "CellMindAI" menu in your Google Sheet
2. Click "CellMindAI" > "Configure API Key" to set up your Anthropic API key
3. Click "CellMindAI" > "Process with CellMindAI" to analyze your data
4. Enter your prompt, specify the data range, and choose whether to include headers
5. The results will appear in a new sheet

### Prompt Chains

For complex, multi-step analyses:

1. Click "CellMindAI" > "Execute Prompt Chain" > "YES" to create a template
2. Fill out the template with:
   - Prompts for each step of your analysis
   - Data ranges for each step (can be different sheets/ranges)
   - Whether to include previous results in each step
3. Click "CellMindAI" > "Execute Prompt Chain" > "NO" to run your chain
4. View the complete results in a new sheet

### Data Referencing Examples

You can reference your data in several ways:

- Simple range: `A1:D10`
- Another sheet: `Sales!A1:F20`
- Named range: `MonthlyStats`

## Troubleshooting

If you encounter issues:

1. Click "CellMindAI" > "Run Diagnostics" to check the integration status
2. Ensure your API key is correctly configured
3. Check that data ranges are valid
4. If using library mode, verify the library is correctly linked

## Requirements

- A Google account with access to Google Sheets
- An Anthropic API key (Claude AI)
- Chrome browser (recommended)

## License

MIT License

Copyright (c) 2025 CellMindAI

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
