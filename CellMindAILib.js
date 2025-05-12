/**
 * CellMindAI Integration for Google Sheets
 * This library enables integration of Claude AI with Google Sheets.
 * 
 * Library Identifier: CellMindAILib
 * 
 * Features:
 * - API key management
 * - Prompt processing with reference to sheet data
 * - Chain processing for multi-step requests
 * - Simple integration into existing sheets
 */

// Anthropic API Constants
const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';
const CLAUDE_MODEL = 'claude-3-5-sonnet-20240620'; // Current version - update as needed

/**
 * CellMindAI class for main functionality
 */
class CellMindAI {
  constructor() {
    this.apiKey = null;
    this.userProperties = PropertiesService.getUserProperties();
  }
  
  /**
   * Stores the API key in user settings
   * @param {string} apiKey - The Anthropic API key
   */
  setApiKey(apiKey) {
    if (!apiKey || apiKey.trim() === '') {
      throw new Error('API key cannot be empty');
    }
    
    this.apiKey = apiKey;
    this.userProperties.setProperty('CELLMINDAI_API_KEY', apiKey);
    return 'API key successfully saved';
  }
  
  /**
   * Retrieves the stored API key
   * @return {string} The stored API key or null
   */
  getApiKey() {
    if (this.apiKey) {
      return this.apiKey;
    }
    
    const savedKey = this.userProperties.getProperty('CELLMINDAI_API_KEY');
    if (savedKey) {
      this.apiKey = savedKey;
      return this.apiKey;
    }
    
    return null;
  }
  
  /**
   * Checks if an API key is set
   * @return {boolean} True if API key exists
   */
  hasApiKey() {
    return this.getApiKey() !== null;
  }
  
  /**
   * Deletes the stored API key
   */
  clearApiKey() {
    this.apiKey = null;
    this.userProperties.deleteProperty('CELLMINDAI_API_KEY');
    return 'API key successfully deleted';
  }
  
  /**
   * Sends a request to the Claude AI API
   * @param {string} prompt - The prompt for Claude
   * @param {Array} data - The data from the sheet
   * @param {Object} options - Additional options (max_tokens, temperature, etc.)
   * @return {Object} The response from Claude
   */
  sendPrompt(prompt, data, options = {}) {
    if (!this.hasApiKey()) {
      throw new Error('No API key found. Please set an API key first with setApiKey()');
    }
    
    // Format data as a table
    const dataTable = this._formatDataAsTable(data);
    
    // Create complete prompt with data
    const fullPrompt = `${prompt}\n\nHere is the data from the table:\n\n${dataTable}`;
    
    // API request parameters
    const requestOptions = {
      model: options.model || CLAUDE_MODEL,
      max_tokens: options.max_tokens || 4000,
      temperature: options.temperature || 0.7,
      messages: [{ role: 'user', content: fullPrompt }]
    };
    
    // Send HTTP request
    const response = this._sendRequest(requestOptions);
    
    return {
      response: response.content[0].text,
      rawResponse: response
    };
  }
  
  /**
   * Executes a chain of prompts in sequence
   * @param {Array} promptChain - Array of prompt objects with data and options
   * @return {Array} Array of Claude responses
   */
  executePromptChain(promptChain) {
    if (!Array.isArray(promptChain) || promptChain.length === 0) {
      throw new Error('Prompt chain must be a non-empty array');
    }
    
    const results = [];
    let previousResponse = null;
    
    for (let i = 0; i < promptChain.length; i++) {
      const currentStep = promptChain[i];
      
      // Enhance the current prompt with the previous response, if available
      let enhancedPrompt = currentStep.prompt;
      if (previousResponse && currentStep.includeLastResult) {
        enhancedPrompt = `${enhancedPrompt}\n\nResult of the previous step:\n${previousResponse}`;
      }
      
      // Execute the current prompt
      const result = this.sendPrompt(enhancedPrompt, currentStep.data || [], currentStep.options || {});
      results.push(result);
      
      // Save the response for the next step
      previousResponse = result.response;
    }
    
    return results;
  }
  
  /**
   * Processes the current data from the active sheet
   * @param {string} prompt - The prompt for Claude
   * @param {Object} options - Additional options (range, includeHeaders, etc.)
   * @return {Object} The response from Claude
   */
  processCurrentSheet(prompt, options = {}) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = options.range ? sheet.getRange(options.range) : sheet.getDataRange();
    const data = range.getValues();
    
    // Remove headers if not desired
    const processedData = options.includeHeaders === false ? data.slice(1) : data;
    
    return this.sendPrompt(prompt, processedData, options);
  }
  
  /**
   * Writes the result back to the spreadsheet
   * @param {string} result - The result to be written to the sheet
   * @param {Object} options - Options for writing (sheetName, cell, createNewSheet)
   */
  writeResultToSheet(result, options = {}) {
    let sheet;
    
    // Determine the target sheet
    if (options.createNewSheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(options.sheetName || 'CellMindAI Result');
    } else if (options.sheetName) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(options.sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(options.sheetName);
      }
    } else {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    }
    
    // Determine the target cell
    const cell = options.cell || 'A1';
    sheet.getRange(cell).setValue(result);
    
    return {
      success: true,
      message: `Result written to ${sheet.getName()}!${cell}`
    };
  }
  
  /**
   * Formats 2D array data as a Markdown table
   * @private
   * @param {Array} data - 2D array of data
   * @return {string} Formatted table as string
   */
  _formatDataAsTable(data) {
    if (!data || data.length === 0) {
      return 'No data available';
    }
    
    // Create Markdown table
    let tableStr = '';
    
    // Add all rows
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      tableStr += row.join('\t') + '\n';
    }
    
    return tableStr;
  }
  
  /**
   * Sends HTTP request to the Claude API
   * @private
   * @param {Object} requestData - Request data
   * @return {Object} API response as JSON
   */
  _sendRequest(requestData) {
    const options = {
      method: 'post',
      headers: {
        'x-api-key': this.getApiKey(),
        'content-type': 'application/json',
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(requestData),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(ANTHROPIC_API_URL, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode !== 200) {
        throw new Error(`API error (${responseCode}): ${responseText}`);
      }
      
      return JSON.parse(responseText);
    } catch (error) {
      throw new Error(`Error during API request: ${error.message}`);
    }
  }
}

// Global instance
let cellMindInstance = null;

/**
 * Initializes the CellMindAI instance
 * @return {CellMindAI} The CellMindAI instance
 */
function initCellMindAI() {
  if (!cellMindInstance) {
    cellMindInstance = new CellMindAI();
  }
  return cellMindInstance;
}

/**
 * Sets the API key for CellMindAI
 * @param {string} apiKey - The Anthropic API key
 * @return {string} Confirmation message
 */
function setApiKey(apiKey) {
  return initCellMindAI().setApiKey(apiKey);
}

/**
 * Checks if an API key is set
 * @return {boolean} True if API key exists
 */
function hasApiKey() {
  return initCellMindAI().hasApiKey();
}

/**
 * Deletes the stored API key
 * @return {string} Confirmation message
 */
function clearApiKey() {
  return initCellMindAI().clearApiKey();
}

/**
 * Processes data from the current sheet with CellMindAI
 * @param {string} prompt - The prompt for CellMindAI
 * @param {Object} options - Additional options
 * @return {Object} CellMindAI response
 */
function processSheet(prompt, options = {}) {
  return initCellMindAI().processCurrentSheet(prompt, options);
}

/**
 * Sends a prompt with custom data to CellMindAI
 * @param {string} prompt - The prompt for CellMindAI
 * @param {Array} data - The data (2D array)
 * @param {Object} options - Additional options
 * @return {Object} CellMindAI response
 */
function sendCustomData(prompt, data, options = {}) {
  return initCellMindAI().sendPrompt(prompt, data, options);
}

/**
 * Executes a chain of prompts
 * @param {Array} promptChain - Array of prompt objects
 * @return {Array} Array of CellMindAI responses
 */
function executePromptChain(promptChain) {
  return initCellMindAI().executePromptChain(promptChain);
}

/**
 * Writes a result to the spreadsheet
 * @param {string} result - The result to be written
 * @param {Object} options - Options for writing
 * @return {Object} Result of the write operation
 */
function writeResult(result, options = {}) {
  return initCellMindAI().writeResultToSheet(result, options);
}

/**
 * Creates a menu in the spreadsheet for using the library
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CellMindAI')
    .addItem('Configure API Key', 'showApiKeyDialog')
    .addSeparator()
    .addItem('Process Data with CellMindAI', 'showPromptDialog')
    .addItem('Execute Prompt Chain', 'showChainDialog')
    .addToUi();
}

/**
 * Shows a dialog for configuring the API key
 */
function showApiKeyDialog() {
  const ui = SpreadsheetApp.getUi();
  const hasKey = hasApiKey();
  
  const promptMsg = hasKey ? 
    'An API key is already configured. Do you want to replace it?' :
    'Please enter your Anthropic API key:';
  
  const result = ui.prompt(
    'Configure CellMindAI API Key',
    promptMsg,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const key = result.getResponseText();
    if (key && key.trim() !== '') {
      setApiKey(key);
      ui.alert('Success', 'API key has been successfully saved.', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'Empty API key. No changes made.', ui.ButtonSet.OK);
    }
  }
}

/**
 * Shows a dialog for entering a prompt
 */
function showPromptDialog() {
  const ui = SpreadsheetApp.getUi();
  
  if (!hasApiKey()) {
    ui.alert('Error', 'No API key configured. Please configure an API key first.', ui.ButtonSet.OK);
    return;
  }
  
  // Simple prompt dialog
  const promptResult = ui.prompt(
    'CellMindAI Prompt',
    'Enter your prompt for CellMindAI:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (promptResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const prompt = promptResult.getResponseText();
  if (!prompt || prompt.trim() === '') {
    ui.alert('Error', 'Empty prompt. Please enter a prompt.', ui.ButtonSet.OK);
    return;
  }
  
  // Range dialog
  const rangeResult = ui.prompt(
    'Data Range',
    'Enter the data range (e.g., A1:D10) or leave empty to use the entire sheet:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (rangeResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const range = rangeResult.getResponseText();
  
  // Headers dialog
  const includeHeadersResult = ui.alert(
    'Include Headers',
    'Do you want to include column headers?',
    ui.ButtonSet.YES_NO
  );
  
  const includeHeaders = (includeHeadersResult === ui.Button.YES);
  
  // Process the prompt
  try {
    const options = {
      includeHeaders: includeHeaders
    };
    
    if (range && range.trim() !== '') {
      options.range = range;
    }
    
    const result = processSheet(prompt, options);
    
    // Write the result to a new sheet
    writeResult(result.response, {
      createNewSheet: true,
      sheetName: 'CellMindAI Result ' + new Date().toLocaleString()
    });
    
    ui.alert('Success', 'The request was successfully processed.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Error during processing: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Shows a dialog for executing a prompt chain
 */
function showChainDialog() {
  const ui = SpreadsheetApp.getUi();
  
  if (!hasApiKey()) {
    ui.alert('Error', 'No API key configured. Please configure an API key first.', ui.ButtonSet.OK);
    return;
  }
  
  // Inform the user about the format for prompt chains
  ui.alert(
    'Prompt Chain Instructions',
    'To execute a prompt chain, your current sheet must have the following structure:\n\n' +
    '1. Column A: Prompts (one per row)\n' +
    '2. Column B: Data ranges (optional, one per row)\n' +
    '3. Column C: Include previous result (true/false, optional)\n\n' +
    'The first row should contain headers.\n\n' +
    'Click OK to process the prompt chain from the current sheet.',
    ui.ButtonSet.OK_CANCEL
  );
  
  // Get the current sheet data
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      ui.alert('Error', 'The sheet must contain at least a header row and one data row.', ui.ButtonSet.OK);
      return;
    }
    
    // Build the prompt chain
    const promptChain = [];
    
    for (let i = 1; i < data.length; i++) {
      const prompt = data[i][0];
      const rangeStr = data[i][1];
      const includeLastResult = data[i][2] === true || data[i][2] === 'true';
      
      if (!prompt || prompt.trim() === '') continue;
      
      let rangeData = [];
      
      if (rangeStr && rangeStr.trim() !== '') {
        try {
          rangeData = sheet.getRange(rangeStr).getValues();
        } catch (error) {
          ui.alert('Error', `Invalid range "${rangeStr}" in row ${i+1}.`, ui.ButtonSet.OK);
          return;
        }
      }
      
      promptChain.push({
        prompt: prompt,
        data: rangeData,
        includeLastResult: includeLastResult,
        options: {}
      });
    }
    
    if (promptChain.length === 0) {
      ui.alert('Error', 'No valid prompts found in the sheet.', ui.ButtonSet.OK);
      return;
    }
    
    // Execute the prompt chain
    const results = executePromptChain(promptChain);
    
    // Create a new sheet for the results
    const resultSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Chain Results ' + new Date().toLocaleString());
    
    // Write the results
    resultSheet.getRange('A1').setValue('Prompt Chain Results');
    resultSheet.getRange('A1').setFontWeight('bold');
    
    for (let i = 0; i < results.length; i++) {
      resultSheet.getRange(`A${i*4+3}`).setValue(`Step ${i+1}:`);
      resultSheet.getRange(`A${i*4+3}`).setFontWeight('bold');
      
      resultSheet.getRange(`A${i*4+4}`).setValue(results[i].response);
      resultSheet.getRange(`A${i*4+4}:E${i*4+4}`).merge();
      
      // Empty line between results
      if (i < results.length - 1) {
        resultSheet.getRange(`A${i*4+6}`).setValue('');
      }
    }
    
    // Adjust column widths
    resultSheet.autoResizeColumn(1);
    
    ui.alert('Success', 'The prompt chain was successfully executed.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Error executing the prompt chain: ' + error.message, ui.ButtonSet.OK);
  }
}
