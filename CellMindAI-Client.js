/**
 * CellMindAI Integration Client
 * 
 * This client code connects to the CellMindAILib library
 * and provides enhanced functionality for prompt chains
 */

// Global reference to the library
var CellMindLib;
var libraryAvailable = false;

/**
 * Initialize when the spreadsheet opens
 */
function onOpen() {
  try {
    // Initialize library reference
    initLibrary();
    
    // Create menu
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('CellMindAI');
    
    // Add menu items
    menu.addItem('Configure API Key', 'configureApiKey')
      .addSeparator()
      .addItem('Process with CellMindAI', 'processWithCellMind')
      .addItem('Execute Prompt Chain', 'executePromptChain')
      .addItem('Run Diagnostics', 'runDiagnostics');
    
    // Add menu to UI
    menu.addToUi();
  } catch (e) {
    console.error("Error in onOpen:", e);
    
    // Create a minimal error menu
    SpreadsheetApp.getUi()
      .createMenu('CellMindAI (Error)')
      .addItem('Run Diagnostics', 'runDiagnostics')
      .addToUi();
  }
}

/**
 * Initialize the library reference
 */
function initLibrary() {
  try {
    // Check if the library is available
    if (typeof CellMindAILib !== 'undefined') {
      CellMindLib = CellMindAILib;
      
      // Test a function to confirm it's working
      if (typeof CellMindLib.hasApiKey === 'function') {
        libraryAvailable = true;
        console.log("Library initialized successfully");
      } else {
        libraryAvailable = false;
        console.log("Library object exists but functions not accessible");
      }
    } else {
      libraryAvailable = false;
      console.log("Library not available");
    }
  } catch (e) {
    libraryAvailable = false;
    console.error("Error initializing library:", e);
  }
  
  return libraryAvailable;
}

/**
 * Run diagnostics to help identify issues
 */
function runDiagnostics() {
  const ui = SpreadsheetApp.getUi();
  let report = "=== CellMindAI Integration Diagnostics ===\n\n";
  
  // Check library accessibility
  report += "1. Library accessibility check:\n";
  try {
    if (typeof CellMindAILib !== 'undefined') {
      report += "   ✓ Library is defined in global scope\n";
      
      // Check for specific functions
      try {
        if (typeof CellMindAILib.hasApiKey === 'function') {
          report += "   ✓ Library functions are accessible\n";
          
          // Try actual function call
          try {
            const hasKey = CellMindAILib.hasApiKey();
            report += "   ✓ Function call successful: hasApiKey() returned " + hasKey + "\n";
          } catch (e) {
            report += "   ✗ Function call failed: " + e.message + "\n";
          }
        } else {
          report += "   ✗ Library functions not accessible\n";
        }
      } catch (e) {
        report += "   ✗ Error checking functions: " + e.message + "\n";
      }
    } else {
      report += "   ✗ Library not defined in global scope\n";
    }
  } catch (e) {
    report += "   ✗ Error checking library: " + e.message + "\n";
  }
  
  // Check API key
  report += "\n2. API Key check:\n";
  try {
    const hasApiKey = getApiKey() !== null;
    report += "   API Key status: " + (hasApiKey ? "API key is configured" : "No API key configured") + "\n";
  } catch (e) {
    report += "   ✗ Error checking API key: " + e.message + "\n";
  }
  
  // Library status
  report += "\n3. Current operation mode:\n";
  report += "   Using " + (libraryAvailable ? "library mode" : "fallback mode") + "\n";
  
  // Display report
  ui.alert("Diagnostics Results", report, ui.ButtonSet.OK);
}

/**
 * Configure the API key
 */
function configureApiKey() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const response = ui.prompt(
      'Configure CellMindAI API Key',
      'Enter your Anthropic API key:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const key = response.getResponseText().trim();
      if (key) {
        // Try library method first
        if (libraryAvailable) {
          CellMindLib.setApiKey(key);
        } else {
          // Fallback to direct property storage
          PropertiesService.getUserProperties().setProperty('CELLMINDAI_API_KEY', key);
        }
        ui.alert('Success', 'API key saved successfully', ui.ButtonSet.OK);
      } else {
        ui.alert('Error', 'API key cannot be empty', ui.ButtonSet.OK);
      }
    }
  } catch (error) {
    ui.alert('Error', 'Error saving API key: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Get the API key
 */
function getApiKey() {
  try {
    if (libraryAvailable) {
      return CellMindLib.getApiKey();
    } else {
      return PropertiesService.getUserProperties().getProperty('CELLMINDAI_API_KEY');
    }
  } catch (e) {
    // Fallback if library method fails
    return PropertiesService.getUserProperties().getProperty('CELLMINDAI_API_KEY');
  }
}

/**
 * Process data with CellMindAI
 */
function processWithCellMind() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Check for API key
    const apiKey = getApiKey();
    if (!apiKey) {
      ui.alert('Error', 'Please configure your API key first', ui.ButtonSet.OK);
      return;
    }
    
    // Get prompt
    const promptResponse = ui.prompt(
      'Process with CellMindAI',
      'Enter your prompt:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (promptResponse.getSelectedButton() !== ui.Button.OK) return;
    
    const prompt = promptResponse.getResponseText().trim();
    if (!prompt) {
      ui.alert('Error', 'Prompt cannot be empty', ui.ButtonSet.OK);
      return;
    }
    
    // Get range
    const rangeResponse = ui.prompt(
      'Data Range',
      'Enter the range to process (e.g., A1:D20) or leave empty for all data:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (rangeResponse.getSelectedButton() !== ui.Button.OK) return;
    
    const rangeStr = rangeResponse.getResponseText().trim();
    
    // Get headers preference
    const includeHeadersResponse = ui.alert(
      'Include Headers',
      'Include column headers in the data?',
      ui.ButtonSet.YES_NO
    );
    
    const includeHeaders = (includeHeadersResponse === ui.Button.YES);
    
    // Process data
    if (libraryAvailable) {
      // Use library method
      try {
        const options = {
          includeHeaders: includeHeaders,
          range: rangeStr || undefined
        };
        
        const result = CellMindLib.processSheet(prompt, options);
        
        // Write result to a new sheet
        CellMindLib.writeResult(result.response, {
          createNewSheet: true,
          sheetName: 'CellMindAI Result ' + new Date().toLocaleString()
        });
        
        ui.alert('Success', 'Processing complete (library mode)', ui.ButtonSet.OK);
      } catch (e) {
        ui.alert('Library Error', 'Error using library: ' + e.message + '\nSwitching to fallback mode...', ui.ButtonSet.OK);
        processDirectly(prompt, rangeStr, includeHeaders);
      }
    } else {
      // Use direct method
      processDirectly(prompt, rangeStr, includeHeaders);
    }
  } catch (error) {
    ui.alert('Error', 'Processing error: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Process directly without library
 */
function processDirectly(prompt, rangeStr, includeHeaders) {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Get data
    const sheet = SpreadsheetApp.getActiveSheet();
    let data;
    
    if (rangeStr) {
      data = sheet.getRange(rangeStr).getValues();
    } else {
      data = sheet.getDataRange().getValues();
    }
    
    // Remove headers if not desired
    const processedData = includeHeaders ? data : data.slice(1);
    
    // Format data as table
    let dataTable = '';
    for (let i = 0; i < processedData.length; i++) {
      dataTable += processedData[i].join('\t') + '\n';
    }
    
    // Create full prompt
    const fullPrompt = `${prompt}\n\nHere is the data from the table:\n\n${dataTable}`;
    
    // Send to Claude API
    const result = callClaudeAPI(fullPrompt);
    
    // Create result sheet
    const resultSheet = SpreadsheetApp.getActiveSpreadsheet()
      .insertSheet('CellMindAI Result ' + new Date().toLocaleString());
    
    resultSheet.getRange('A1').setValue(result);
    resultSheet.setColumnWidth(1, 800);
    
    ui.alert('Success', 'Processing complete (fallback mode)', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error processing data: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Enhanced prompt chain execution function
 */
function executePromptChain() {
  const ui = SpreadsheetApp.getUi();
  
  // Check for API key
  if (!getApiKey()) {
    ui.alert(
      'API Key Missing', 
      'Please configure your API key first via the "CellMindAI" > "Configure API Key" menu.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Clear options
  const result = ui.alert(
    'Execute Prompt Chain',
    'What would you like to do?\n\n' +
    '- YES: Create a new template for prompt chains\n' +
    '- NO: Execute an existing prompt chain from the current sheet\n' +
    '- CANCEL: Return to spreadsheet',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (result === ui.Button.YES) {
    createPromptChainTemplate();
  } else if (result === ui.Button.NO) {
    executePromptChainFromCurrentSheet();
  }
}

/**
 * Creates a template for prompt chains
 */
function createPromptChainTemplate() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if template already exists
  let sheet = ss.getSheetByName('Prompt Chain Template');
  if (sheet) {
    const renameResult = ui.alert(
      'Existing Template Found',
      'A "Prompt Chain Template" already exists. What would you like to do?\n\n' +
      '- YES: Overwrite the existing template\n' +
      '- NO: Rename the existing template and create a new one\n' +
      '- CANCEL: Abort operation',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (renameResult === ui.Button.CANCEL) {
      return;
    } else if (renameResult === ui.Button.NO) {
      // Rename
      sheet.setName('Prompt Chain Template (old)');
      sheet = null;
    }
    // For YES: Overwrite, so use existing sheet
  }
  
  // Create new template if needed
  if (!sheet) {
    sheet = ss.insertSheet('Prompt Chain Template');
  } else {
    // Clear the sheet if we're overwriting
    sheet.clear();
  }
  
  // Set headers
  sheet.getRange('A1:D1').setValues([['Prompt', 'Data Range', 'Include Previous Result', 'Notes']]);
  sheet.getRange('A1:D1').setFontWeight('bold');
  
  // Example data
  sheet.getRange('A2:D5').setValues([
    ['Analyze this data and identify the top 3 trends', 'Sheet1!A1:F20', 'NO', 'Initial analysis of raw data'],
    ['Based on the above analysis, explain possible causes for these trends', '', 'YES', 'Cause analysis based on step 1'],
    ['Create recommendations based on the identified trends and causes', '', 'YES', 'Derive recommendations from steps 1+2'],
    ['Summarize all findings and recommendations in a concise executive summary', '', 'YES', 'Final conclusion for management']
  ]);
  
  // Formatting
  sheet.setColumnWidth(1, 400); // Prompt
  sheet.setColumnWidth(2, 200); // Data range
  sheet.setColumnWidth(3, 200); // Include previous
  sheet.setColumnWidth(4, 250); // Notes
  
  // Add help and description
  sheet.getRange('A7:D7').merge();
  sheet.getRange('A7').setValue('HELP: How to use this template');
  sheet.getRange('A7').setFontWeight('bold');
  
  sheet.getRange('A8:D16').merge();
  sheet.getRange('A8').setValue(
    'DATA RANGE REFERENCES:\n' +
    '- Single sheet: Sheet1!A1:F20\n' +
    '- Named range: NamedRange\n' +
    '- Current sheet: A1:D10\n' +
    '- Leave empty for no additional data\n\n' +
    'INCLUDE PREVIOUS RESULT:\n' +
    '- YES: Result from previous step will be included in this prompt\n' +
    '- NO: Prompt will be executed without previous result\n\n' +
    'After filling in, execute the chain via: CellMindAI > Execute Prompt Chain > NO (Execute existing)'
  );
  
  // Cell formatting - "Include Previous Result" column
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['YES', 'NO'], true)
    .build();
  sheet.getRange('C2:C100').setDataValidation(validationRule);
  
  // Examples for data range references
  sheet.getRange('A18:D18').merge();
  sheet.getRange('A18').setValue('EXAMPLES OF DATA RANGE REFERENCES:');
  sheet.getRange('A18').setFontWeight('bold');
  
  sheet.getRange('A19:D22').setValues([
    ['Type', 'Example', 'Description', 'Usage'],
    ['Simple Range', 'A1:D10', 'Range in current sheet', 'For data in current sheet'],
    ['Sheet Reference', 'Sheet1!A1:F20', 'Range in another sheet', 'For data in other sheets'],
    ['Named Range', 'MyRange', 'Named range in spreadsheet', 'For predefined data ranges']
  ]);
  
  // Success notification
  ui.alert(
    'Template Created',
    'The prompt chain template has been successfully created. Fill it with your prompts and then execute the chain with "Execute Prompt Chain" > NO.',
    ui.ButtonSet.OK
  );
  
  // Switch to new sheet
  sheet.activate();
  
  return sheet;
}

/**
 * Executes a prompt chain from the current sheet
 * Enhanced version with better data referencing
 */
function executePromptChainFromCurrentSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    ui.alert(
      'Insufficient Data',
      'The sheet must contain at least a header row and one data row.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Identify columns flexibly (regardless of position and exact wording)
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  
  // Find prompt column (required)
  const promptKeywords = ['prompt', 'query', 'question', 'instruction'];
  const promptCol = findColumnIndex(headers, promptKeywords);
  
  if (promptCol === -1) {
    ui.alert(
      'Prompt Column Missing',
      'Could not find a column with "Prompt" (or similar term). ' +
      'Please ensure your sheet has such a column.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Find data range column (optional)
  const rangeKeywords = ['range', 'data', 'area', 'selection'];
  const rangeCol = findColumnIndex(headers, rangeKeywords);
  
  // Find previous result column (optional)
  const includeKeywords = ['previous', 'include', 'prior', 'last'];
  const includeCol = findColumnIndex(headers, includeKeywords);
  
  // Build prompt chain
  const promptChain = [];
  let hasValidPrompts = false;
  
  for (let i = 1; i < data.length; i++) {
    const promptText = data[i][promptCol];
    
    // Skip empty rows
    if (!promptText || promptText.trim() === '') continue;
    
    hasValidPrompts = true;
    
    // Process data range if specified
    let rangeData = [];
    if (rangeCol !== -1 && data[i][rangeCol]) {
      const rangeReference = data[i][rangeCol].toString().trim();
      
      if (rangeReference) {
        try {
          // Process different reference types
          if (rangeReference.includes('!')) {
            // Sheet reference (e.g. "Sheet1!A1:B10")
            const [sheetName, rangeAddress] = rangeReference.split('!');
            const targetSheet = ss.getSheetByName(sheetName);
            
            if (!targetSheet) {
              throw new Error(`Sheet "${sheetName}" not found`);
            }
            
            rangeData = targetSheet.getRange(rangeAddress).getValues();
          } else {
            try {
              // Try as named range
              const namedRange = ss.getRangeByName(rangeReference);
              
              if (namedRange) {
                rangeData = namedRange.getValues();
              } else {
                // Try as normal range in current sheet
                rangeData = sheet.getRange(rangeReference).getValues();
              }
            } catch (e) {
              // Try as normal range in current sheet
              rangeData = sheet.getRange(rangeReference).getValues();
            }
          }
        } catch (e) {
          // Show warning for invalid range
          const continueResult = ui.alert(
            'Data Range Issue',
            `There was a problem with the data range "${rangeReference}" in row ${i+1}: ${e.message}\n\n` +
            'Would you like to continue with an empty data set for this step?',
            ui.ButtonSet.YES_NO
          );
          
          if (continueResult === ui.Button.NO) {
            return;
          }
          
          // For YES: Use empty data set (already initialized as [])
        }
      }
    }
    
    // Check if previous result should be included
    let includeLastResult = false;
    if (includeCol !== -1) {
      const includeValue = String(data[i][includeCol]).toLowerCase().trim();
      includeLastResult = 
        includeValue === 'yes' || 
        includeValue === 'true' || 
        includeValue === '1';
    }
    
    // Add prompt object to chain
    promptChain.push({
      prompt: promptText,
      data: rangeData,
      includeLastResult: includeLastResult,
      options: {}
    });
  }
  
  if (!hasValidPrompts) {
    ui.alert(
      'No Valid Prompts',
      'No valid prompts were found in the sheet. Please add at least one prompt.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Get confirmation
  const confirmResult = ui.alert(
    'Execute Prompt Chain',
    `Ready to execute ${promptChain.length} prompts in sequence.\n\n` +
    'Results will be saved in a new sheet.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResult !== ui.Button.YES) return;
  
  // Show progress indicator
  const statusDialog = ui.alert(
    'Execution Started',
    'The prompt chain is now being executed. This may take some time depending on the number and complexity of prompts.\n\n' +
    'Click OK to continue. You will be notified when execution is complete.',
    ui.ButtonSet.OK
  );
  
  // Execute prompt chain (either with library or directly)
  try {
    let results;
    
    // Execute chain (library or directly)
    if (libraryAvailable && typeof CellMindLib.executePromptChain === 'function') {
      // Use library method
      results = CellMindLib.executePromptChain(promptChain);
    } else {
      // Use direct method
      results = executeChainDirectly(promptChain);
    }
    
    // Write results to new sheet
    const resultSheet = ss.insertSheet(`Prompt Chain Results ${new Date().toLocaleString()}`);
    
    // Heading
    resultSheet.getRange('A1').setValue('Prompt Chain Results');
    resultSheet.getRange('A1').setFontWeight('bold');
    resultSheet.getRange('A1:E1').merge();
    
    // Write results
    for (let i = 0; i < results.length; i++) {
      // Step heading
      resultSheet.getRange(`A${i*5+3}`).setValue(`Step ${i+1}:`);
      resultSheet.getRange(`A${i*5+3}`).setFontWeight('bold');
      resultSheet.getRange(`B${i*5+3}`).setValue(promptChain[i].prompt);
      resultSheet.getRange(`B${i*5+3}:E${i*5+3}`).merge();
      
      // Result
      resultSheet.getRange(`A${i*5+4}`).setValue(results[i].response || results[i]);
      resultSheet.getRange(`A${i*5+4}:E${i*5+4}`).merge();
      resultSheet.getRange(`A${i*5+4}`).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      
      // Spacing between results
      if (i < results.length - 1) {
        resultSheet.getRange(`A${i*5+6}`).setValue('');
      }
    }
    
    // Formatting
    resultSheet.setColumnWidth(1, 200);
    resultSheet.setColumnWidths(2, 4, 150);
    
    // Success notification
    ui.alert(
      'Execution Complete',
      'The prompt chain was successfully executed. Results have been saved to a new sheet.',
      ui.ButtonSet.OK
    );
    
    // Switch to results sheet
    resultSheet.activate();
  } catch (error) {
    ui.alert(
      'Execution Error',
      'An error occurred while executing the prompt chain:\n\n' + error.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Execute a prompt chain directly (without library)
 * @param {Array} promptChain - Array of prompt objects
 * @return {Array} Array of results
 */
function executeChainDirectly(promptChain) {
  const results = [];
  let previousResult = null;
  
  for (let i = 0; i < promptChain.length; i++) {
    // Format data as table
    let dataTable = '';
    if (promptChain[i].data && promptChain[i].data.length > 0) {
      for (let j = 0; j < promptChain[i].data.length; j++) {
        dataTable += promptChain[i].data[j].join('\t') + '\n';
      }
    }
    
    // Create full prompt
    let fullPrompt = promptChain[i].prompt;
    
    if (previousResult && promptChain[i].includeLastResult) {
      fullPrompt += '\n\nResult from previous step:\n' + previousResult;
    }
    
    if (dataTable) {
      fullPrompt += '\n\nHere is the data from the table:\n\n' + dataTable;
    }
    
    // Call Claude API
    const result = callClaudeAPI(fullPrompt);
    previousResult = result;
    results.push(result);
  }
  
  return results;
}

/**
 * Call the Claude API directly
 * @param {string} prompt - The prompt for Claude
 * @return {string} Claude's response
 */
function callClaudeAPI(prompt) {
  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error('API key not configured');
  }
  
  const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';
  const CLAUDE_MODEL = 'claude-3-5-sonnet-20240620';
  
  const requestOptions = {
    model: CLAUDE_MODEL,
    max_tokens: 4000,
    temperature: 0.7,
    messages: [{ role: 'user', content: prompt }]
  };
  
  const options = {
    method: 'post',
    headers: {
      'x-api-key': apiKey,
      'content-type': 'application/json',
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(requestOptions),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(ANTHROPIC_API_URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (responseCode !== 200) {
    throw new Error(`API error (${responseCode}): ${responseText}`);
  }
  
  const jsonResponse = JSON.parse(responseText);
  return jsonResponse.content[0].text;
}

/**
 * Helper function to find a column based on keywords
 * @param {Array} headers - Array of header strings
 * @param {Array} keywords - Array of keywords to match
 * @return {number} Index of found column or -1
 */
function findColumnIndex(headers, keywords) {
  for (let i = 0; i < headers.length; i++) {
    for (let j = 0; j < keywords.length; j++) {
      if (headers[i].includes(keywords[j])) {
        return i;
      }
    }
  }
  return -1;
}
