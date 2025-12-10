// Google Apps Script for APP Form and Data Team Rating Form
// This script handles form submission and saves data to Google Sheets

const SHEET_ID = '1O_-A3yIEpyh-3211I4SbYaWNzHuVM-3QPtyjQjJGl2Q';
const TAB_NAME = 'DTF';
const APP_TAB_NAME = 'APP';

/**
 * doGet function to serve the HTML forms
 * Handles routing for both APP form and Data Team Rating Form
 */
function doGet(e) {
  // Check if this is a request for a scoring form
  if (e.parameter.page) {
    const page = e.parameter.page;
    let filename = '';
    let title = '';
    
    switch(page) {
      case 'collaboration':
        filename = 'Collaboration_Scoring';
        title = 'Score Collaboration';
        break;
      case 'consulting':
        filename = 'Consulting_Scoring';
        title = 'Score Consulting';
        break;
      case 'coaching':
        filename = 'Coaching_Scoring';
        title = 'Score General Coaching';
        break;
      case 'meeting':
        filename = 'Meeting_Scoring';
        title = 'Score Meeting Participation';
        break;
      case 'datateam':
        filename = 'DataTeamRatingForm';
        title = 'Score Data Team';
        break;
      case 'coachdatacollection':
        filename = 'CoachDataCollectionRubric';
        title = 'Coach Data Collection Rubric';
        break;
      case 'teacherdataanalysis':
        filename = 'TeacherDataAnalysisRubric';
        title = 'Teacher Data Analysis Rubric';
        break;
      case 'meetingstructures':
        filename = 'MeetingStructuresRubric';
        title = 'Meeting Structures Rubric';
        break;
      default:
        // If unknown page parameter, fall through to main APP form
        break;
    }
    
    if (filename) {
      return HtmlService.createHtmlOutputFromFile(filename)
        .setTitle(title)
        .setWidth(900)
        .setHeight(800)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }
  
  // Default: Show main APP form
  const template = HtmlService.createTemplateFromFile('APPNotes');
  
  let pageTitle = 'APP Form';
  
  // Set the script URL as a template variable
  try {
    template.scriptUrl = ScriptApp.getService().getUrl() || '';
  } catch (error) {
    Logger.log('Error getting script URL: ' + error);
    template.scriptUrl = '';
  }
  
  // Get row number from URL parameter
  if (e.parameter.row) {
    template.rowNumber = e.parameter.row;
    
    // Get the event title from the row to use as title
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(APP_TAB_NAME);
      // Event Title is in column 7 (Column G)
      const eventTitle = sheet.getRange(parseInt(e.parameter.row), 7).getValue();
      
      if (eventTitle) {
        pageTitle = eventTitle + ' - APP';
      }
    } catch (error) {
      Logger.log('Error getting event title for page title: ' + error);
    }
  } else {
    template.rowNumber = 'null';
  }
  
  return template.evaluate()
    .setTitle(pageTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://www.gstatic.com/images/branding/product/1x/apps_script_48dp.png');
}

// ============================================================================
// DATA TEAM RATING FORM FUNCTIONS
// ============================================================================

/**
 * Submit form data to Google Sheet
 * @param {Object} formData - The form data object containing all responses
 * @returns {Object} Success or error result
 */
function submitFormData(formData) {
  try {
    // Open the spreadsheet and get or create the tab
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(TAB_NAME);
    
    // If tab doesn't exist, create it
    if (!sheet) {
      sheet = ss.insertSheet(TAB_NAME);
    }
    
    // Check if this is a new sheet (no data at all)
    const lastRow = sheet.getLastRow();
    
    // Only create headers if sheet is completely empty
    if (lastRow === 0) {
      // Add header row
      const headers = [
        'Date',
        'Data Team',
        'Total Score',
        'Percentage',
        'Average Score',
        'P1: All team members present & on time, actively engaged in the process',
        'P2: Administrator is present and engaged, provides input, and supports intervention adjustments',
        'P3: Related professionals are invited & participate as appropriate',
        'C1: Members really embrace team roles (recorder, focus monitor, time keeper, etc)',
        'C2: Teachers willingly share student information & instructional details',
        'C3: All members express an interest in other teachers\' data – ask questions, volunteer ideas',
        'C4: Teachers are open to consideration of intervention changes & suggestions from others',
        'D1: Members focus on observable data rather than "feelings, thoughts, impressions"',
        'D2: Teachers set aggressive/ambitious goals for their students',
        'D3: Teachers indicate that they are analyzing graphs frequently, both in between & prior to data team meetings',
        'D4: Teachers can independently identify when a student needs a change in instruction/intervention based on the data',
        'D5: Teachers conduct weekly progress monitoring',
        'D6: Teachers share data/show graphs to students',
        'D7: Members make connections between FASTBridge data & other sources of academic data (i-Ready, SIPPS, Corrective Reading, Focus Math, etc.)',
        'D8: Members recognize the need and appropriateness for more diagnostic assessments to determine areas of weakness and strength',
        'I1: Teachers brainstorm possible intervention ideas prior to the data team',
        'I2: Teachers can articulate instructional/intervention details (beyond just naming a program or strategy)',
        'I3: Team conversation is focused on things within our sphere of control (instruction, curriculum, environment)',
        'I4: Team members equally offer intervention ideas to their peers',
        'I5: Team members exhibit flexibility in scheduling, planning, providing services for students',
        'I6: Team considers research behind interventions (discusses whether current instructional practices are research-based)',
        'I7: Team considers fidelity of instruction when analyzing data, and brainstorms strategies for improving fidelity',
        'I8: For students requiring a change, teachers leave meeting with a plan',
        'I9: If time does not allow for resolution of a student discussion, team tables discussion with a specific plan/steps for follow up'
      ];
      
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header row (only on creation)
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#8d6e63');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setFontFamily('McLaren');
    }
    
    // Prepare the data row
    const rowData = [
      formData.date,
      formData.dataTeam,
      formData.totalScore,
      formData.percentage,
      formData.averageScore,
      formData.p1,
      formData.p2,
      formData.p3,
      formData.c1,
      formData.c2,
      formData.c3,
      formData.c4,
      formData.d1,
      formData.d2,
      formData.d3,
      formData.d4,
      formData.d5,
      formData.d6,
      formData.d7,
      formData.d8,
      formData.i1,
      formData.i2,
      formData.i3,
      formData.i4,
      formData.i5,
      formData.i6,
      formData.i7,
      formData.i8,
      formData.i9
    ];
    
    // Add the data to the next available row
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Format the data row
    const dataRange = sheet.getRange(nextRow, 1, 1, rowData.length);
    dataRange.setFontFamily('McLaren');
    
    return {
      success: true,
      message: 'Form data saved successfully!',
      row: nextRow
    };
    
  } catch (error) {
    Logger.log('Error in submitFormData: ' + error.toString());
    return {
      success: false,
      message: 'Error saving form data: ' + error.toString()
    };
  }
}

/**
 * Calculate total score from form responses
 * @param {Object} responses - Object containing all form responses
 * @returns {Number} Total score
 */
function calculateTotalScore(responses) {
  const scores = [
    responses.p1, responses.p2, responses.p3,
    responses.c1, responses.c2, responses.c3, responses.c4,
    responses.d1, responses.d2, responses.d3, responses.d4, 
    responses.d5, responses.d6, responses.d7, responses.d8,
    responses.i1, responses.i2, responses.i3, responses.i4, 
    responses.i5, responses.i6, responses.i7, responses.i8, responses.i9
  ];
  
  return scores.reduce((total, score) => total + score, 0);
}

/**
 * Utility function to refresh/fix headers in existing sheet
 * Run this if your sheet has blank headers or needs the Average Score column added
 */
function refreshHeaders() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(TAB_NAME);
    
    if (!sheet) {
      Logger.log('Tab "' + TAB_NAME + '" does not exist.');
      return false;
    }
    
    // Define headers
    const headers = [
      'Date',
      'Data Team',
      'Total Score',
      'Percentage',
      'Average Score',
      'P1: All team members present & on time, actively engaged in the process',
      'P2: Administrator is present and engaged, provides input, and supports intervention adjustments',
      'P3: Related professionals are invited & participate as appropriate',
      'C1: Members really embrace team roles (recorder, focus monitor, time keeper, etc)',
      'C2: Teachers willingly share student information & instructional details',
      'C3: All members express an interest in other teachers\' data – ask questions, volunteer ideas',
      'C4: Teachers are open to consideration of intervention changes & suggestions from others',
      'D1: Members focus on observable data rather than "feelings, thoughts, impressions"',
      'D2: Teachers set aggressive/ambitious goals for their students',
      'D3: Teachers indicate that they are analyzing graphs frequently, both in between & prior to data team meetings',
      'D4: Teachers can independently identify when a student needs a change in instruction/intervention based on the data',
      'D5: Teachers conduct weekly progress monitoring',
      'D6: Teachers share data/show graphs to students',
      'D7: Members make connections between FASTBridge data & other sources of academic data (i-Ready, SIPPS, Corrective Reading, Focus Math, etc.)',
      'D8: Members recognize the need and appropriateness for more diagnostic assessments to determine areas of weakness and strength',
      'I1: Teachers brainstorm possible intervention ideas prior to the data team',
      'I2: Teachers can articulate instructional/intervention details (beyond just naming a program or strategy)',
      'I3: Team conversation is focused on things within our sphere of control (instruction, curriculum, environment)',
      'I4: Team members equally offer intervention ideas to their peers',
      'I5: Team members exhibit flexibility in scheduling, planning, providing services for students',
      'I6: Team considers research behind interventions (discusses whether current instructional practices are research-based)',
      'I7: Team considers fidelity of instruction when analyzing data, and brainstorms strategies for improving fidelity',
      'I8: For students requiring a change, teachers leave meeting with a plan',
      'I9: If time does not allow for resolution of a student discussion, team tables discussion with a specific plan/steps for follow up'
    ];
    
    // Write headers to row 1
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#8d6e63');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontFamily('McLaren');
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log('Headers refreshed successfully!');
    return true;
    
  } catch (error) {
    Logger.log('Error refreshing headers: ' + error.toString());
    return false;
  }
}

// ============================================================================
// APP FORM FUNCTIONS
// ============================================================================

/**
 * Get data from a specific row in the APP sheet
 * @param {number} rowNumber - The row number to retrieve
 * @returns {Object} Row data or error
 */
function getRowData(rowNumber) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(APP_TAB_NAME);
    
    if (!sheet) {
      return { error: 'APP sheet not found' };
    }
    
    if (!rowNumber || rowNumber < 2) {
      return { error: 'Invalid row number' };
    }
    
    // Get all data from the row (now including columns through AC at index 28)
    const data = sheet.getRange(rowNumber, 1, 1, 29).getValues()[0];
    
    Logger.log('Raw data from row ' + rowNumber + ': ' + JSON.stringify(data));
    
    // Handle time formatting
    let timeValue = '';
    if (data[1]) {
      const timeCell = sheet.getRange(rowNumber, 2);
      const displayValue = timeCell.getDisplayValue();
      Logger.log('Time display value: ' + displayValue);
      
      if (displayValue && displayValue.includes(':')) {
        const match = displayValue.match(/(\d{1,2}):(\d{2})/);
        if (match) {
          let hours = parseInt(match[1]);
          const minutes = match[2];
          
          if (displayValue.toLowerCase().includes('pm') && hours < 12) {
            hours += 12;
          } else if (displayValue.toLowerCase().includes('am') && hours === 12) {
            hours = 0;
          }
          
          timeValue = String(hours).padStart(2, '0') + ':' + minutes;
        }
      }
    }
    
    // Handle date formatting
    let dateValue = '';
    if (data[0] instanceof Date) {
      const year = data[0].getFullYear();
      const month = String(data[0].getMonth() + 1).padStart(2, '0');
      const day = String(data[0].getDate()).padStart(2, '0');
      dateValue = year + '-' + month + '-' + day;
    }
    
    // Get formatted notes from the cell
    const notesHtml = getFormattedNotesAsHtml(sheet, rowNumber, 14);

    // Convert agenda and taskField to HTML (they contain checkbox symbols)
    const agendaText = sheet.getRange(rowNumber, 22).getValue() || '';
    const agendaHtml = convertBulletsToHtml(agendaText.toString());

    const taskFieldText = sheet.getRange(rowNumber, 11).getValue() || '';
    const taskFieldHtml = convertBulletsToHtml(taskFieldText.toString());

    // Column mapping:
    // A=Date(0), B=Time(1), C=Location(2), D=Owner(3), E=Name(4), F=HIDDEN(5), G=Title(6),
    // H=Type(7), I=Purpose(8), J=HIDDEN(9), K=Task(10), L=Status(11), M=Strategy(12), N=Notes(13), O=Minutes(14),
    // V=Agenda(21), W=Thoughts(22), Y=Consciousness(24), Z=Craftsmanship(25), AA=Efficacy(26), AB=Flexibility(27), AC=Interdependence(28)
    const rowData = {
      row: rowNumber,
      date: dateValue,
      time: timeValue,
      location: data[2] || '',
      owner: data[3] || '',
      name: data[4] || '',
      eventTitle: data[6] || '',        // Column G (index 6)
      task: data[7] || '',              // Column H (index 7)
      purpose: data[8] || '',           // Column I (index 8)
      taskField: taskFieldHtml,          // Column K (index 10) - converted to HTML
      status: data[11] || '',           // Column L (index 11)
      strategy: data[12] || '',         // Column M (index 12)
      notes: notesHtml,                 // Column N (index 13)
      minutes: data[14] || 0,           // Column O (index 14)
      agenda: agendaHtml,               // Column V (index 21) - converted to HTML
      thoughts: data[22] || '',         // Column W (index 22)
      consciousness: data[24] || 0,     // Column Y (index 24)
      craftsmanship: data[25] || 0,     // Column Z (index 25)
      efficacy: data[26] || 0,          // Column AA (index 26)
      flexibility: data[27] || 0,       // Column AB (index 27)
      interdependence: data[28] || 0    // Column AC (index 28)
    };
    
    Logger.log('Returning rowData: ' + JSON.stringify(rowData));
    return rowData;
    
  } catch (error) {
    Logger.log('Error in getRowData: ' + error.toString());
    return { error: 'Error loading row data: ' + error.toString() };
  }
}

/**
 * Save or update row data in the APP sheet
 * @param {Object} data - The form data to save
 * @returns {Object} Success result with row number
 */
function saveRowData(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(APP_TAB_NAME);
    
    if (!sheet) {
      return { success: false, message: 'APP sheet not found' };
    }
    
    const row = data.row;
    
    if (!row) {
      throw new Error('No row specified');
    }
    
    // Handle date
    if (data.date) {
      const dateParts = data.date.split('-');
      const dateString = dateParts[1] + '/' + dateParts[2] + '/' + dateParts[0];
      sheet.getRange(row, 1).setValue(dateString);
    }
    
    // Handle time
    if (data.time) {
      const timeParts = data.time.split(':');
      const timeString = timeParts[0] + ':' + timeParts[1] + ':00';
      sheet.getRange(row, 2).setValue(timeString);
      sheet.getRange(row, 2).setNumberFormat('h:mm AM/PM');
    }
    
    // Save all other fields
    sheet.getRange(row, 3).setValue(data.location || '');      // Column C
    sheet.getRange(row, 4).setValue(data.owner || '');         // Column D
    sheet.getRange(row, 5).setValue(data.name || '');          // Column E
    sheet.getRange(row, 7).setValue(data.eventTitle || '');    // Column G
    sheet.getRange(row, 8).setValue(data.task || '');          // Column H
    sheet.getRange(row, 9).setValue(data.purpose || '');       // Column I
    sheet.getRange(row, 11).setValue(data.taskField || '');    // Column K
    sheet.getRange(row, 12).setValue(data.status || '');       // Column L
    sheet.getRange(row, 13).setValue(data.strategy || '');     // Column M
    
    // Save formatted notes
    if (data.notes) {
      saveFormattedNotes(sheet, row, 14, data.notes);
    } else {
      sheet.getRange(row, 14).setValue('');
    }
    
    // Handle taskField formatting (Column K) - with error handling
    try {
      const columnKCell = sheet.getRange(row, 11);
      if (data.taskField && data.taskField.trim() !== '') {
        // Try to remove conditional formatting rules for this cell
        try {
          const rules = sheet.getConditionalFormatRules();
          const newRules = rules.filter(function(rule) {
            const ranges = rule.getRanges();
            const affectsCell = ranges.some(function(range) {
              const startRow = range.getRow();
              const endRow = range.getLastRow();
              const startCol = range.getColumn();
              const endCol = range.getLastColumn();
              return row >= startRow && row <= endRow && 11 >= startCol && 11 <= endCol;
            });
            return !affectsCell;
          });
          sheet.setConditionalFormatRules(newRules);
        } catch (formatError) {
          Logger.log('Warning: Could not remove conditional formatting: ' + formatError.toString());
        }
        
        columnKCell.setBackgroundRGB(255, 0, 203);
        columnKCell.setFontColor('#ffffff');
      } else {
        columnKCell.setBackground(null);
        columnKCell.setFontColor(null);
      }
    } catch (taskFieldFormatError) {
      Logger.log('Warning: Error formatting task field: ' + taskFieldFormatError.toString());
    }
    
    sheet.getRange(row, 15).setValue(data.minutes || 0);       // Column O
    sheet.getRange(row, 22).setValue(data.agenda || '');       // Column V
    
    // Handle agenda field formatting (Column V) - with error handling
    try {
      const columnVCell = sheet.getRange(row, 22);
      if (data.agenda && data.agenda.trim() !== '') {
        try {
          const rules = sheet.getConditionalFormatRules();
          const newRules = rules.filter(function(rule) {
            const ranges = rule.getRanges();
            const affectsCell = ranges.some(function(range) {
              const startRow = range.getRow();
              const endRow = range.getLastRow();
              const startCol = range.getColumn();
              const endCol = range.getLastColumn();
              return row >= startRow && row <= endRow && 22 >= startCol && 22 <= endCol;
            });
            return !affectsCell;
          });
          sheet.setConditionalFormatRules(newRules);
        } catch (formatError) {
          Logger.log('Warning: Could not remove conditional formatting: ' + formatError.toString());
        }
        
        columnVCell.setBackgroundRGB(250, 248, 132);
        columnVCell.setFontColor('#000000');
      } else {
        columnVCell.setBackground(null);
        columnVCell.setFontColor(null);
      }
    } catch (agendaFormatError) {
      Logger.log('Warning: Error formatting agenda field: ' + agendaFormatError.toString());
    }
    
    // Save Thoughts field (Column W)
    sheet.getRange(row, 23).setValue(data.thoughts || '');     // Column W
    
    // Handle thoughts field formatting (Column W) - with error handling
    try {
      const columnWCell = sheet.getRange(row, 23);
      if (data.thoughts && data.thoughts.trim() !== '') {
        try {
          const rules = sheet.getConditionalFormatRules();
          const newRules = rules.filter(function(rule) {
            const ranges = rule.getRanges();
            const affectsCell = ranges.some(function(range) {
              const startRow = range.getRow();
              const endRow = range.getLastRow();
              const startCol = range.getColumn();
              const endCol = range.getLastColumn();
              return row >= startRow && row <= endRow && 23 >= startCol && 23 <= endCol;
            });
            return !affectsCell;
          });
          sheet.setConditionalFormatRules(newRules);
        } catch (formatError) {
          Logger.log('Warning: Could not remove conditional formatting: ' + formatError.toString());
        }
        
        columnWCell.setBackground('#82CAFF');
        columnWCell.setFontColor('#000000');
      } else {
        columnWCell.setBackground(null);
        columnWCell.setFontColor(null);
      }
    } catch (thoughtsFormatError) {
      Logger.log('Warning: Error formatting thoughts field: ' + thoughtsFormatError.toString());
    }
    
    // Save States of Mind ratings (Columns Y, Z, AA, AB, AC)
    sheet.getRange(row, 25).setValue(data.consciousness || 0);    // Column Y
    sheet.getRange(row, 26).setValue(data.craftsmanship || 0);    // Column Z
    sheet.getRange(row, 27).setValue(data.efficacy || 0);         // Column AA
    sheet.getRange(row, 28).setValue(data.flexibility || 0);      // Column AB
    sheet.getRange(row, 29).setValue(data.interdependence || 0);  // Column AC
    
    return {success: true, row: row};
      
  } catch (error) {
    Logger.log('Error in saveRowData: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    throw new Error('Save failed: ' + error.toString());
  }
}

/**
 * Save effect size data to a separate tab
 * @param {Object} effectSizeData - Object containing pre, post, name, date, and purpose
 * @returns {Object} Success or error result
 */
function saveEffectSize(effectSizeData) {
  try {
    Logger.log('Saving effect size data: ' + JSON.stringify(effectSizeData));
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('CoachESData');
    
    // If the sheet doesn't exist, create it
    if (!sheet) {
      sheet = ss.insertSheet('CoachESData');
      // Add headers
      sheet.getRange(1, 1, 1, 4).setValues([['Name', 'Purpose', 'PRE', 'POST']]);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    }
    
    // Find the next available row (first row where column A is empty)
    let nextRow = 2; // Start at row 2 (after header)
    const lastRow = sheet.getLastRow();
    
    // Check each row starting from row 2 to find the first empty cell in column A
    for (let i = 2; i <= lastRow + 1; i++) {
      const cellValue = sheet.getRange(i, 1).getValue();
      if (!cellValue || cellValue === '') {
        nextRow = i;
        break;
      }
    }
    
    // Write the data to the next available row
    sheet.getRange(nextRow, 1).setValue(effectSizeData.name || '');
    sheet.getRange(nextRow, 2).setValue(effectSizeData.purpose || '');
    sheet.getRange(nextRow, 3).setValue(effectSizeData.pre || '');
    sheet.getRange(nextRow, 4).setValue(effectSizeData.post || '');
    
    Logger.log('Effect size data saved to row ' + nextRow);
    
    return {
      success: true,
      row: nextRow,
      message: 'Effect size data saved successfully to row ' + nextRow
    };
    
  } catch (error) {
    Logger.log('Error in saveEffectSize: ' + error.toString());
    throw new Error('Failed to save effect size: ' + error.toString());
  }
}

/**
 * Save action plan data to the ActPlan tab
 * @param {Object} actionPlanData - Object containing action plan details
 * @returns {Object} Success or error result
 */
function saveActionPlan(actionPlanData) {
  try {
    Logger.log('Saving action plan data: ' + JSON.stringify(actionPlanData));
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('ActPlan');
    
    if (!sheet) {
      return { 
        success: false, 
        message: 'ActPlan sheet not found. Please create the ActPlan tab first.' 
      };
    }
    
    const teacherName = actionPlanData.name;
    
    if (!teacherName || teacherName.trim() === '') {
      return { 
        success: false, 
        message: 'No teacher name provided' 
      };
    }
    
    // Find the teacher's row
    const lastRow = sheet.getLastRow();
    const namesColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    let teacherRow = -1;
    for (let i = 0; i < namesColumn.length; i++) {
      if (namesColumn[i][0] === teacherName) {
        teacherRow = i + 2;
        break;
      }
    }
    
    if (teacherRow === -1) {
      return { 
        success: false, 
        message: 'Teacher "' + teacherName + '" not found in ActPlan tab' 
      };
    }
    
    Logger.log('Found teacher at row: ' + teacherRow);
    
    // Find the first empty group of 4 columns starting at column D (column 4)
    const lastCol = sheet.getLastColumn();
    let targetCol = -1;
    
    // Start at column D (index 4) and check every group of 4 columns
    for (let col = 4; col <= lastCol; col += 4) {
      const groupData = sheet.getRange(teacherRow, col, 1, 4).getValues()[0];
      
      if (!groupData[0] || groupData[0] === '') {
        targetCol = col;
        Logger.log('Found empty group at column: ' + targetCol);
        break;
      }
    }
    
    // If no empty group found in existing columns, use the next group after last column
    if (targetCol === -1) {
      if (lastCol < 4) {
        targetCol = 4;
      } else {
        targetCol = Math.ceil((lastCol + 1) / 4) * 4;
        if (targetCol < 4) targetCol = 4;
      }
      Logger.log('No empty group found, using new group at column: ' + targetCol);
    }
    
    // Save the action plan data to the target group
    sheet.getRange(teacherRow, targetCol).setValue(actionPlanData.actionItems || '');
    sheet.getRange(teacherRow, targetCol + 1).setValue(actionPlanData.dueDate || '');
    sheet.getRange(teacherRow, targetCol + 2).setValue(actionPlanData.actionStatus || '');
    sheet.getRange(teacherRow, targetCol + 3).setValue(actionPlanData.actionDetails || '');
    
    // Format the due date cell as a date if there's a value
    if (actionPlanData.dueDate && actionPlanData.dueDate !== '') {
      sheet.getRange(teacherRow, targetCol + 1).setNumberFormat('m/d/yyyy');
    }
    
    // Apply text wrapping to the Details/Notes cell
    sheet.getRange(teacherRow, targetCol + 3).setWrap(true);
    sheet.getRange(teacherRow, targetCol + 3).setVerticalAlignment('top');
    
    Logger.log('Action plan data saved successfully');
    
    // Convert column number to letter for the message
    function columnToLetter(column) {
      let temp, letter = '';
      while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
      }
      return letter;
    }
    
    return {
      success: true,
      row: teacherRow,
      column: targetCol,
      message: 'Action plan saved successfully at row ' + teacherRow + ', columns ' + 
               columnToLetter(targetCol) + '-' + columnToLetter(targetCol + 3)
    };
    
  } catch (error) {
    Logger.log('Error in saveActionPlan: ' + error.toString());
    return { 
      success: false, 
      message: 'Failed to save action plan: ' + error.toString() 
    };
  }
}

/**
 * Get the most recent action plan for a teacher
 * @param {string} teacherName - The teacher's name
 * @returns {Object} Most recent action plan data or null
 */
function getCurrentActionPlan(teacherName) {
  try {
    if (!teacherName || teacherName.trim() === '') {
      return { success: false, message: 'No teacher name provided' };
    }
    
    Logger.log('Getting current action plan for: ' + teacherName);
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('ActPlan');
    
    if (!sheet) {
      return { success: false, message: 'ActPlan sheet not found' };
    }
    
    // Find the teacher's row
    const lastRow = sheet.getLastRow();
    const namesColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    let teacherRow = -1;
    for (let i = 0; i < namesColumn.length; i++) {
      if (namesColumn[i][0] === teacherName) {
        teacherRow = i + 2;
        break;
      }
    }
    
    if (teacherRow === -1) {
      return { success: false, message: 'Teacher not found in ActPlan tab' };
    }
    
    Logger.log('Found teacher at row: ' + teacherRow);
    
    // Search from right to left to find the most recent action plan
    const lastCol = sheet.getLastColumn();
    let mostRecentCol = -1;
    
    let startCol = Math.floor((lastCol - 3) / 4) * 4 + 4;
    if (startCol < 4) startCol = 4;
    
    // Search from right to left
    for (let col = startCol; col >= 4; col -= 4) {
      const actionItemValue = sheet.getRange(teacherRow, col).getValue();
      
      if (actionItemValue && actionItemValue !== '') {
        mostRecentCol = col;
        Logger.log('Found most recent action plan at column: ' + mostRecentCol);
        break;
      }
    }
    
    if (mostRecentCol === -1) {
      return { success: false, message: 'No action plans found for ' + teacherName };
    }
    
    // Get the data from the most recent action plan
    const actionItems = sheet.getRange(teacherRow, mostRecentCol).getValue() || '';
    const dueDate = sheet.getRange(teacherRow, mostRecentCol + 1).getValue() || '';
    const status = sheet.getRange(teacherRow, mostRecentCol + 2).getValue() || '';
    const details = sheet.getRange(teacherRow, mostRecentCol + 3).getValue() || '';
    
    // Format the due date if it's a Date object
    let formattedDueDate = '';
    if (dueDate instanceof Date) {
      formattedDueDate = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    } else if (dueDate) {
      formattedDueDate = String(dueDate);
    }
    
    return {
      success: true,
      actionItems: actionItems,
      dueDate: formattedDueDate,
      status: status,
      details: details,
      teacher: teacherName
    };
    
  } catch (error) {
    Logger.log('Error in getCurrentActionPlan: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

/**
 * Get conversation questions based on strategy type
 * @param {string} conversationType - The strategy type selected
 * @param {string} currentNotes - Current notes content to preserve
 * @returns {Object} HTML content with questions or error
 */
function getConversationQuestions(conversationType, currentNotes) {
  try {
    Logger.log('Getting questions for: ' + conversationType);
    
    // Preserve existing notes if they exist
    var existingContent = '';
    if (currentNotes && currentNotes.trim() !== '') {
      existingContent = currentNotes + '<br><br>';
    }
    
    // Check if strategy contains "meeting" (case insensitive)
    if (conversationType.toLowerCase().indexOf('meeting') !== -1) {
      const meetingQuestions = [
        'Notes'
      ];
      
      let html = '';
      for (var i = 0; i < meetingQuestions.length; i++) {
        html += '<b>' + meetingQuestions[i] + '</b><br>';
        html += '<ul><li></li></ul>';
        if (i < meetingQuestions.length - 1) {
          html += '<br>';
        }
      }
      
      return { success: true, html: existingContent + html };
    }
    
    const templates = {
      'Planning Conversation': [
        'What goal are you working on? What is the outcome?',
        'What will it look like or sound like when you reach that goal?',
        'What are some of the things you\'ve already tried?',
        'What might be some strategies you\'re considering?',
        'What could be some of the ways I can support you?'
      ],
      'Reflecting Conversation': [
        'When you think about your professional practice, what might be some of the things you have been focused on?  To what extent have you met your goal?',
        'What data or information do we have that measures the progress?  When you think about these data, what are you noticing?  On what aspects of the data would you like to focus?',
        'What might need to be changed or adjusted?  What goals or outcomes do you have in mind?  What could it look like or sound like when you reach the goal?',
        'What ideas for adjustments do you have moving forward?  What might be some strategies to use?',
        'What might be some support you need from me?'
      ],
      'Problem Resolving Conversation': [
        'What would you like to talk about?  What is the problem?',
        'What would it look like or sound like if the problem is resolved or gets better?',
        'What ideas do you have for how to address the problem?',
        'Moving forward, what might be some actions steps you are planning to take?'
      ],
      'General Coaching': [
        'What\'s going well?',
        'What are some of your current challenges?',
        'What is your plan moving forward?',
        'What might be some ways I could support you?'
      ],
      'Observation': [
        'Areas of Strength',
        'Opportunities for Improvement'
      ],
      'PDSA': [
        'Plan',
        '',
        'Do',
        '',
        'Study',
        'Plus',
        '•',
        'Delta',
        '•',
        'Act',
        ''
      ],
      'AM Tasks': [
        'Inbox / Email',
        'Update Calendar',
        'Pull Appointments from Calendar',
        'Calculate Mileage',
        'Feedback Emails',
        'Appointment Reminders',
        'Prep for Appointments',
        'Import Data Team Data',
        'Prep for Data Teams',
        'Data Team Reminders',
        'KickUp - Create Logs',
        'Student Specific - Prep'
      ],
      'PM Tasks': [
        'Inbox / Email',
        'Update Calendar',
        'Submit Mileage',
        'KickUp - Submit Logs',
        'Update Tracker',
        'Tasks from Appointments',
        'Student Specific Reports',
        'Observation Reports',
        'IEP Alignment Form',
        'Coaching Reports',
        'Update Teacher Data Sheets',
        'Move Teacher and Student Files',
        'Push Data Team Rows to Teacher Tabs',
        'Archive Appointments',
        'Push Appointments to Calendar',
        'Pull Appointments for Tomorrow'
      ],
      'Other Tasks': [
        'Code - Debug, Adjust, Create',
        'Memorandum of Conversation',
        'Other Task (Division, Professional Responsibility, Etc.)'
      ],
      'Weekend Task': [
        'Data Entry',
        'Data Analysis',
        'Division Task',
        'HSD Task',
        'Coding',
        'Review Upcoming Action Plans',
        'Review Previous Notes for Upcoming Appointments'
      ],
      'Data Team': [
        'Reading Goal',
        'NOTELINE',
        'Math Goal',
        'NOTELINE',
        'Social Emotional Goal',
        'NOTELINE',
        'Written Expression Goal',
        'NOTELINE',
        'Notes:'
      ],
      'Peer Observation': [
        'Host Teacher:',
        'Focus:',
        'Next Steps:'
      ],
      'Impact Cycle': [
        'IDENTIFY',
        'On a scale of 1-10, with 1 being the worst lesson you\'ve ever taught and 10 being your ideal or best lesson, how close was that lesson to your ideal?',
        'What pleased you about this lesson?',
        'What would you have to change to make the lesson closer to a 10?',
        'What would you see you students doing differently if your class was a 10?',
        'Tell me more about what that would look like.',
        'What teaching strategy can you use to hit your goal?',
        'What are your next steps?',
        'When should we meet again?',
        'What tasks have to be done before we meet?',
        'When will those tasks be done?',
        'Who will do them?',
        'IMPROVE',
        'What has gone well?',
        'What are you seeing that shows this strategy is successful?',
        'What progress has been made toward the goal?',
        'What did you learn?',
        'What surprised you?',
        'When should we meet again?',
        'What roadblocks are you running into?'
      ],
      'Student Specific Observation': [
        'Name of Student:',
        'What is the purpose for the request or the concern that is being addressed?',
        'To what extent are the Universal Classroom Supports in place?',
        'What IEP accommodations or behavior intervention plan strategies are expected?',
        'To what extent is the IEP implemented with fidelity?',
        'Notes:'
      ],
      'Appointment Reminders': [
        'Reminder sent'
      ],
      'Data Team Reminders': [
        'Reminder sent'
      ],
      'Explicit Instruction - Part 1': [
        'Part 1 How do you know you have created the right objectives and provided effective modeling?',
        'The methods used to create objectives should:',
        'Choose objectives based on student performance relative to goals.',
        '☐ Select a goal from IEP or standards',
        'NOTELINE',
        '☐ Choose an objective that is the next step toward the goal',
        'NOTELINE',
        'Write focused objectives that describe the specific learning outcome.',
        '☐ Limit the objective to one singular next step toward the goal',
        'NOTELINE',
        '☐ Describe a learning outcome in behavioral terms that assesses mastery of the objective',
        'NOTELINE',
        'The methods used to provide modeling should:',
        'Give clear explanations',
        '• Match the explanation to the learning outcome',
        'NOTELINE',
        '• Design the explanation so that it is correct, clear and concise',
        'NOTELINE',
        '• Use the explanation consistently',
        'NOTELINE',
        'Model multiple planned examples',
        '• Show all steps or provide unique examples',
        'NOTELINE',
        '• Verbalize your thinking',
        'NOTELINE',
        '• Have students observe',
        'NOTELINE',
        'Use supporting practices'
      ],
      'Explicit Instruction - Part 2': [
        'Part 2 How do you know you have effectively implemented practice?',
        'The methods used to provide guided practice should:',
        'Lead students in steps toward the learning outcome',
        '• Procedural task: execute each step separately',
        'NOTELINE',
        '• Knowledge task: address each unit (e.g. vocabulary) one at a time',
        'NOTELINE',
        'Provide appropriate prompts',
        '• Design a variety of prompt types linked to task and likely student need',
        'NOTELINE',
        '• Change level of prompting in response to student\'s progress',
        'NOTELINE',
        'Observe and provide immediate feedback',
        '• Watch students carefully',
        'NOTELINE',
        '• Interpret the meaning of errors',
        'NOTELINE',
        '• Provide feedback that aligns with the type of error',
        'NOTELINE',
        'The methods used to provide independent practice should:',
        'Review expectations and resources for meeting the learning outcome',
        '• Select objective-linked practice items that lead to 90-95% accuracy',
        'NOTELINE',
        '• Review expectations for meeting outcomes',
        'NOTELINE',
        '• Remind students how they self-prompt',
        'NOTELINE',
        'Allow student to work without support',
        '• Limit guidance from teacher',
        'NOTELINE',
        '• Monitor throughout independent practice',
        'NOTELINE',
        '• Give mini-reminders and record points of confusion',
        'NOTELINE',
        'Observe and provide immediate and delayed feedback',
        '• Check for understanding throughout',
        'NOTELINE',
        '• Provide feedback following completion of task',
        'NOTELINE'
      ],
      'Explicit Instruction - Part 3': [
        'Part 3 How do you know you have effectively used the supporting practices?',
        'The methods used to elicit a response should:',
        '• Maintain or check accuracy of processing',
        'NOTELINE',
        '• Match the learning outcome',
        'NOTELINE',
        '• Match student abilities',
        'NOTELINE',
        '• Match the desired response format',
        'NOTELINE',
        '• Maximize student involvement',
        'NOTELINE',
        'The methods used to provide feedback should be:',
        '• Immediate: delivered as soon as possible after response',
        'NOTELINE',
        '• Specific: tied directly to students\' actions',
        'NOTELINE',
        'The methods used to maintain a brisk pace should:',
        '• Move on when students are ready',
        'NOTELINE',
        '• Use the other supporting practices',
        'NOTELINE',
        'The methods used to encourage student reflection and data collection should:',
        '• Encourage students to self-assess their use of the skill',
        'NOTELINE',
        '• Provide relevant assessment data directly aligned with the lesson outcome',
        'NOTELINE'
      ],
      'HLP 7': [
        'Posts 3-5 classroom rules that are positively stated and referenced when applicable.',
        'NOTELINE',
        'Classroom routines for all members of the classroom (students and adults) are evident, with a minimum of entering and exiting procedures.',
        'NOTELINE',
        'Implements classroom procedures in a positive manner in which the teacher is always the exemplar in treating everyone with dignity.',
        'NOTELINE',
        'The allocated time is structured and focused on instruction.',
        'NOTELINE',
        'Gives specific, clear positive feedback at a ratio at least 4:1 when compared to corrective feedback that includes academic, behavior, and social-emotional.',
        'NOTELINE',
        'The classroom and other high-traffic areas are designed in a way that will meet the needs of the students.',
        'NOTELINE'
      ],
      'Universal Checklist': [
        '1. Effectively design the classroom environment that supports the learning outcomes (Centers, arranging desks, lighting….)',
        '• The classroom has clearly defined learning spaces',
        'NOTELINE',
        '• A clear path for mobility for all students',
        'NOTELINE',
        '• All materials are accessible to all students',
        'NOTELINE',
        '• Classroom is clear of clutter',
        'NOTELINE',
        '2. The classroom environment is culturally relevant, enriched with materials that are culturally diverse.',
        '• Classrooms have diverse representation of cultures and abilities, through materials, props, clothing, and languages on visuals',
        'NOTELINE',
        '• The classroom has readily available culturally relevant materials including books, pictures, visuals etc.',
        'NOTELINE',
        '3. There are opportunities for movement within the room.',
        '• Students have built in brain breaks and opportunities for movement during whole and small group times, transitions, or a specific classroom space is dedicated to large muscle movement.',
        'NOTELINE',
        '• The teacher provides intentional opportunities for movement that are based on learning goals.',
        'NOTELINE',
        '4. Classroom Routines:',
        '• Developed for each part of the day',
        'NOTELINE',
        '• Connected to school-wide expectations',
        'NOTELINE',
        '• Include student voice',
        'NOTELINE',
        '• Adult (TAs, SLPs, classroom teachers…) routines exist',
        'NOTELINE',
        '• Taught and practiced',
        'NOTELINE',
        '• A variety of routines are posted.',
        'NOTELINE',
        '• Classroom staff reference the visual routines as a prompt for students as needed.',
        'NOTELINE',
        '• There is evidence of teaching as students follow routines for various times of the day.',
        'NOTELINE',
        '• Transitions are not disruptive to the learning environment',
        'NOTELINE',
        '• Instruction occurs with minimal downtime',
        'NOTELINE',
        '6. Class Schedule:',
        '• A classroom schedule is accessible in all environments.',
        'NOTELINE',
        '• Students are allowed to manipulate the schedule as activities are completed.',
        'NOTELINE',
        '• The teacher has a schedule within a schedule or activity routines visually represented.',
        'NOTELINE',
        '• Schedules use modes that meet student needs',
        'NOTELINE',
        '• Schedule varies preferred with non-preferred activities',
        'NOTELINE',
        '7. Classroom rules:',
        '• Three to five positively stated expectations/rules posted in a variety of locations/settings.',
        'NOTELINE',
        '• Rules have been taught, practiced and referenced.',
        'NOTELINE',
        '• Students are able to demonstrate understanding of rules verbally/gesturally/behaviorally with minimal re-directs.',
        'NOTELINE',
        '• Teaching of rules is expanded and accessible throughout the day as needed.',
        'NOTELINE',
        '• Expectations/rules posted in mode matched to student need',
        'NOTELINE',
        '• Reviewed prior to instruction, and are referenced throughout instruction',
        'NOTELINE'
      ],
      'Classroom Learning System': [
        '1. The teacher sets and communicates direction for students and their families',
        'Learning Requirements',
        '• Posted',
        'NOTELINE',
        '• Kid Friendly Language',
        'NOTELINE',
        '• Teacher mentions learning requirement in lesson',
        'NOTELINE',
        'Class Goal',
        '• Posted',
        'NOTELINE',
        '• Long term',
        'NOTELINE',
        '• Student friendly language/representations',
        'NOTELINE',
        '• Teacher refers to class goal',
        'NOTELINE',
        '• Goal is measurable',
        'NOTELINE',
        'Learning Results',
        '• Chart/Visual for hard data',
        'NOTELINE',
        '• Teacher refers to learning results',
        'NOTELINE',
        '• Teacher refers to analyzing the learning results',
        'NOTELINE',
        'Mission Statement',
        '• Posted',
        'NOTELINE',
        '• Adult Language/generic Student Language',
        'NOTELINE',
        '• Consensus is visually shown (signed, consensogram)',
        'NOTELINE',
        '• Teacher references mission statement',
        'NOTELINE',
        '• Students communicate mission statement',
        'NOTELINE',
        '2. The teacher engages students in regular and frequent evaluation and improvement of classroom learning processes.',
        'Plan',
        '• Learning target posted',
        'NOTELINE',
        '• Measure of proficiency included',
        'NOTELINE',
        '• Clearly leads to achievement of the class goal',
        'NOTELINE',
        '• Teacher shares learning target with students',
        'NOTELINE',
        'Do',
        '• High yield strategies posted',
        'NOTELINE',
        '• Teacher is using the high yield strategies',
        'NOTELINE',
        '• Strategies include the role of the teacher and student',
        'NOTELINE',
        '• Students are included in picking the strategies used',
        'NOTELINE',
        '• A strategy bank is posted/used',
        'NOTELINE',
        '• Teacher refers to/names the high yield strategies and/or why the class is using them',
        'NOTELINE',
        'Study',
        '• Chart for posting results of learning target proficiency',
        'NOTELINE',
        '• Evidence of student feedback is posted',
        'NOTELINE',
        '• Students participate in a plus/delta or other form of feedback',
        'NOTELINE',
        '• Students analyze the data',
        'NOTELINE',
        '• Teacher refers to the learning results',
        'NOTELINE',
        '• Teacher refers to plus/delta',
        'NOTELINE',
        'Act',
        '• Documentation of adjustment in the learning cycle',
        'NOTELINE',
        '• Based on feedback from the previous cycle results',
        'NOTELINE',
        '• Kid language/input',
        'NOTELINE',
        '• Teacher refers to adjustments',
        'NOTELINE',
        '• Teacher reminds students of their/the class adjustments',
        'NOTELINE'
      ]
    };
    
    const questions = templates[conversationType];
    
    if (!questions) {
      return { success: false, message: 'No questions found for: ' + conversationType };
    }
    
    let html = '';
    
    // Special handling for PDSA
    if (conversationType === 'PDSA') {
      html += '<b>Study</b><br>';
      html += '<b>Plus</b><br>';
      html += '<ul><li></li></ul>';
      html += '<b>Delta</b><br>';
      html += '<ul><li></li></ul>';
      html += '<b>Act</b><br><br><br>';
      html += '<b>Plan</b><br><br><br>';
      html += '<b>Do</b><br><br>';
    } else if (conversationType === 'Observation') {
      // Special handling for Observation - agenda checkboxes go to agenda field
      var agendaHtml = '<div class="tasks-checklist">';
      var agendaItems = ['Data Collection', 'Report', 'Feedback'];
      for (var a = 0; a < agendaItems.length; a++) {
        agendaHtml += '<div class="checkbox-item">';
        agendaHtml += '<input type="checkbox" id="agenda_task_' + a + '">';
        agendaHtml += '<label contenteditable="true">' + agendaItems[a] + '</label>';
        agendaHtml += '</div>';
      }
      agendaHtml += '</div>';
      
      // Questions go to notes field
      for (var i = 0; i < questions.length; i++) {
        html += '<b>' + questions[i] + '</b><br>';
        html += '<ul><li></li></ul>';
      }
      
      return { success: true, html: existingContent + html, agendaHtml: agendaHtml };
    } else if (conversationType === 'General Coaching') {
      // Special handling for General Coaching - agenda checkboxes go to agenda field, task items go to task field
      
      // Agenda items
      var agendaHtml = '<div class="tasks-checklist">';
      var agendaItems = ['Coaching', 'Schedule Next Appointment', 'Report'];
      for (var a = 0; a < agendaItems.length; a++) {
        agendaHtml += '<div class="checkbox-item">';
        agendaHtml += '<input type="checkbox" id="agenda_task_' + a + '">';
        agendaHtml += '<label contenteditable="true">' + agendaItems[a] + '</label>';
        agendaHtml += '</div>';
      }
      agendaHtml += '</div>';
      
      // Task field items
      var taskFieldHtml = '<div class="tasks-checklist">';
      var taskItems = ['KickUp', 'Action Plan'];
      for (var t = 0; t < taskItems.length; t++) {
        taskFieldHtml += '<div class="checkbox-item">';
        taskFieldHtml += '<input type="checkbox" id="taskField_task_' + t + '">';
        taskFieldHtml += '<label contenteditable="true">' + taskItems[t] + '</label>';
        taskFieldHtml += '</div>';
      }
      taskFieldHtml += '</div>';
      
      // Questions go to notes field
      for (var i = 0; i < questions.length; i++) {
        html += '<b>' + questions[i] + '</b><br>';
        html += '<ul><li></li></ul>';
        if (i < questions.length - 1) {
          html += '<br><br>';
        }
      }
      
      return { success: true, html: existingContent + html, agendaHtml: agendaHtml, taskFieldHtml: taskFieldHtml };
    } else if (conversationType === 'AM Tasks' || conversationType === 'PM Tasks' || conversationType === 'Other Tasks' || conversationType === 'Weekend Task') {
      html += '<div class="tasks-checklist">';
      for (var i = 0; i < questions.length; i++) {
        html += '<div class="checkbox-item">';
        html += '<input type="checkbox" id="task_' + i + '" data-task="' + questions[i] + '">';
        html += '<label contenteditable="true">' + questions[i] + '</label>';
        html += '</div>';
      }
      html += '</div>';
      html += '<b>Notes:</b><br><br>';
    } else if (conversationType === 'Data Team') {
      // Special handling for Data Team - agenda checkboxes go to agenda field, task items go to task field, goals go to notes
      
      // Agenda items
      var agendaHtml = '<div class="tasks-checklist">';
      var agendaItems = ['Celebrations', 'Academic Data', 'Action Plan', 'Student Specific', 'Meeting Evaluation'];
      for (var a = 0; a < agendaItems.length; a++) {
        agendaHtml += '<div class="checkbox-item">';
        agendaHtml += '<input type="checkbox" id="agenda_task_' + a + '">';
        agendaHtml += '<label contenteditable="true">' + agendaItems[a] + '</label>';
        agendaHtml += '</div>';
      }
      agendaHtml += '</div>';
      
      // Task field items
      var taskFieldHtml = '<div class="tasks-checklist">';
      var taskItems = ['Import Data', 'Prepare Data Team Notes', 'Data Team Reporting Form', 'Data Team Fidelity Assessment', 'Push Data Team to Teacher Tabs'];
      for (var t = 0; t < taskItems.length; t++) {
        taskFieldHtml += '<div class="checkbox-item">';
        taskFieldHtml += '<input type="checkbox" id="taskField_task_' + t + '">';
        taskFieldHtml += '<label contenteditable="true">' + taskItems[t] + '</label>';
        taskFieldHtml += '</div>';
      }
      taskFieldHtml += '</div>';
      
      // Goals and notes go to notes field
      for (var i = 0; i < questions.length; i++) {
        var line = questions[i];
        
        if (line === 'NOTELINE') {
          html += '<ul><li></li></ul>';
        } else {
          html += '<b>' + line + '</b><br>';
        }
      }
      
      return { success: true, html: existingContent + html, agendaHtml: agendaHtml, taskFieldHtml: taskFieldHtml };
    } else if (conversationType === 'Peer Observation') {
      for (var i = 0; i < questions.length; i++) {
        html += '<b>' + questions[i] + '</b><br>';
        html += '<ul><li></li></ul>';
      }
    } else if (conversationType === 'Impact Cycle') {
      for (var i = 0; i < questions.length; i++) {
        if (questions[i] === 'IDENTIFY' || questions[i] === 'IMPROVE') {
          if (i > 0) {
            html += '<br>';
          }
          html += '<b>' + questions[i] + '</b><br><br>';
        } else {
          html += '<b>' + questions[i] + '</b><br>';
          html += '<ul><li></li></ul>';
        }
      }
    } else if (conversationType === 'Student Specific Observation') {
      // Special handling for Student Specific Observation - agenda checkboxes go to agenda field
      var agendaHtml = '<div class="tasks-checklist">';
      var agendaItems = ['Data Collection', 'Records Review', 'Report'];
      for (var a = 0; a < agendaItems.length; a++) {
        agendaHtml += '<div class="checkbox-item">';
        agendaHtml += '<input type="checkbox" id="agenda_task_' + a + '">';
        agendaHtml += '<label contenteditable="true">' + agendaItems[a] + '</label>';
        agendaHtml += '</div>';
      }
      agendaHtml += '</div>';
      
      // Questions go to notes field
      for (var i = 0; i < questions.length; i++) {
        html += '<b>' + questions[i] + '</b><br>';
        html += '<ul><li></li></ul>';
      }
      
      return { success: true, html: existingContent + html, agendaHtml: agendaHtml };
    } else if (conversationType === 'Appointment Reminders' || conversationType === 'Data Team Reminders') {
      for (var i = 0; i < questions.length; i++) {
        html += '<b>' + questions[i] + '</b><br>';
        html += '<ul><li></li></ul>';
      }
    } else if (conversationType === 'Explicit Instruction - Part 1' || conversationType === 'Explicit Instruction - Part 2' || conversationType === 'Explicit Instruction - Part 3') {
      for (var i = 0; i < questions.length; i++) {
        var line = questions[i];
        
        if (line === 'NOTELINE') {
          html += '<ul><li></li></ul>';
        } else if (line.indexOf('☐') === 0 || line.indexOf('•') === 0) {
          html += '<b>' + line + '</b><br>';
        } else {
          html += '<b>' + line + '</b><br>';
        }
      }
    } else if (conversationType === 'HLP 7') {
      for (var i = 0; i < questions.length; i++) {
        var line = questions[i];
        
        if (line === 'NOTELINE') {
          html += '<ul><li></li></ul>';
        } else {
          html += '<b>' + line + '</b><br>';
        }
      }
    } else if (conversationType === 'Universal Checklist') {
      for (var i = 0; i < questions.length; i++) {
        var line = questions[i];
        
        if (line === 'NOTELINE') {
          html += '<ul><li></li></ul>';
        } else if (line.indexOf('•') === 0) {
          html += '<b>' + line + '</b><br>';
        } else {
          html += '<b>' + line + '</b><br>';
        }
      }
    } else if (conversationType === 'Classroom Learning System') {
      for (var i = 0; i < questions.length; i++) {
        var line = questions[i];
        
        if (line === 'NOTELINE') {
          html += '<ul><li></li></ul>';
        } else if (line.indexOf('•') === 0) {
          html += '<b>' + line + '</b><br>';
        } else {
          html += '<b>' + line + '</b><br>';
        }
      }
    } else {
      for (var i = 0; i < questions.length; i++) {
        html += '<b>' + questions[i] + '</b><br>';
        html += '<ul><li></li></ul>';
        if (i < questions.length - 1) {
          html += '<br><br>';
        }
      }
    }
    
    return { success: true, html: existingContent + html };
    
  } catch (error) {
    Logger.log('Error in getConversationQuestions: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

/**
 * Get previous appointment notes for a specific person, optionally filtered by type
 * @param {string} teacherName - The name to search for
 * @param {string} taskType - The task/type to match (optional)
 * @returns {Object} Previous notes HTML or message
 */
function getPreviousAppointmentNotes(teacherName, taskType) {
  try {
    if (!teacherName || teacherName.trim() === '') {
      return { success: false, message: 'No teacher name provided' };
    }
    
    Logger.log('Getting previous notes for: ' + teacherName + ', Type filter: ' + (taskType || 'none'));
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const teacherSheet = ss.getSheetByName(teacherName);
    
    if (!teacherSheet) {
      return { success: false, message: 'No tab found for ' + teacherName };
    }
    
    const lastRow = teacherSheet.getLastRow();
    
    if (lastRow < 2) {
      return { success: false, message: 'No previous appointments found for ' + teacherName };
    }
    
    let targetRow = -1;
    let matchedType = '';
    
    if (taskType && taskType.trim() !== '') {
      const dataRange = teacherSheet.getRange(2, 1, lastRow - 1, 8).getValues();
      
      for (let i = dataRange.length - 1; i >= 0; i--) {
        const rowType = dataRange[i][1];
        const rowNotes = dataRange[i][7];
        
        if (rowType === taskType) {
          if (rowNotes && String(rowNotes).trim() !== '') {
            targetRow = i + 2;
            matchedType = taskType;
            Logger.log('Found matching type "' + taskType + '" with notes at row ' + targetRow);
            break;
          } else {
            Logger.log('Found matching type "' + taskType + '" at row ' + (i + 2) + ' but no notes, continuing search...');
          }
        }
      }
      
      if (targetRow === -1) {
        Logger.log('No matching type with notes found, searching for any row with notes...');
        for (let i = dataRange.length - 1; i >= 0; i--) {
          const rowNotes = dataRange[i][7];
          if (rowNotes && String(rowNotes).trim() !== '') {
            targetRow = i + 2;
            const rowType = dataRange[i][1];
            matchedType = rowType || '';
            Logger.log('Using last row with notes at row ' + targetRow + ', type: ' + matchedType);
            break;
          }
        }
      }
      
      if (targetRow === -1) {
        return { 
          success: false, 
          message: 'Found appointments with type "' + taskType + '" but no notes available' 
        };
      }
    } else {
      const dataRange = teacherSheet.getRange(2, 1, lastRow - 1, 8).getValues();
      
      for (let i = dataRange.length - 1; i >= 0; i--) {
        const rowNotes = dataRange[i][7];
        if (rowNotes && String(rowNotes).trim() !== '') {
          targetRow = i + 2;
          Logger.log('Found last row with notes at row ' + targetRow);
          break;
        }
      }
      
      if (targetRow === -1) {
        return { success: false, message: 'No appointments with notes found for ' + teacherName };
      }
    }
    
    const notesCell = teacherSheet.getRange(targetRow, 8);
    const richTextValue = notesCell.getRichTextValue();
    
    let notesHtml = '';
    
    if (richTextValue) {
      const text = richTextValue.getText();
      if (text) {
        let html = '';
        const runs = richTextValue.getRuns();
        
        for (let i = 0; i < runs.length; i++) {
          const run = runs[i];
          const runText = run.getText();
          const textStyle = run.getTextStyle();
          
          let styledText = runText;
          
          if (textStyle.isBold()) {
            styledText = '<b>' + styledText + '</b>';
          }
          if (textStyle.isItalic()) {
            styledText = '<i>' + styledText + '</i>';
          }
          if (textStyle.isUnderline()) {
            styledText = '<u>' + styledText + '</u>';
          }
          if (textStyle.isStrikethrough()) {
            styledText = '<s>' + styledText + '</s>';
          }
          
          html += styledText;
        }
        
        html = html.replace(/\n/g, '<br>');
        html = convertBulletsToHtml(html);
        notesHtml = html;
      }
    } else {
      const plainValue = notesCell.getValue();
      if (plainValue) {
        notesHtml = String(plainValue).replace(/\n/g, '<br>');
      }
    }
    
    if (!notesHtml || notesHtml.trim() === '') {
      return { success: false, message: 'No notes found in appointment for ' + teacherName };
    }
    
    const dateCell = teacherSheet.getRange(targetRow, 1);
    let dateString = '';
    if (dateCell.getValue()) {
      const dateValue = dateCell.getValue();
      if (dateValue instanceof Date) {
        dateString = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      } else {
        dateString = String(dateValue);
      }
    }
    
    if (matchedType === '') {
      const typeCell = teacherSheet.getRange(targetRow, 2);
      matchedType = typeCell.getValue() || '';
    }
    
    return { 
      success: true, 
      html: notesHtml,
      date: dateString,
      type: matchedType,
      teacher: teacherName
    };
    
  } catch (error) {
    Logger.log('Error in getPreviousAppointmentNotes: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

/**
 * Get the previous POST score for a given purpose
 * @param {string} name - The name (not used, kept for compatibility)
 * @param {string} purpose - The purpose/strategy to match
 * @returns {Object} Previous POST score or null
 */
function getPreviousPostScore(name, purpose) {
  try {
    Logger.log('Getting previous POST score for purpose: ' + purpose);
    
    if (!purpose || purpose.trim() === '') {
      return { success: false, score: null };
    }
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('CoachESData');
    
    if (!sheet) {
      Logger.log('CoachESData sheet not found');
      return { success: false, score: null };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      Logger.log('No data rows found');
      return { success: false, score: null };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    
    let foundScore = null;
    for (let i = data.length - 1; i >= 0; i--) {
      const rowPurpose = data[i][1];
      const rowPost = data[i][3];
      
      if (rowPurpose === purpose && rowPost !== '') {
        foundScore = rowPost;
        Logger.log('Found matching POST score: ' + foundScore + ' for purpose "' + purpose + '" at row ' + (i + 2));
        break;
      }
    }
    
    if (foundScore !== null) {
      return { success: true, score: foundScore };
    } else {
      Logger.log('No matching record found for purpose: ' + purpose);
      return { success: false, score: null };
    }
    
  } catch (error) {
    Logger.log('Error in getPreviousPostScore: ' + error.toString());
    return { success: false, score: null };
  }
}

// ============================================================================
// HELPER FUNCTIONS FOR FORMATTING
// ============================================================================

/**
 * Get formatted notes as HTML from a cell
 */
function getFormattedNotesAsHtml(sheet, row, col) {
  const notesCell = sheet.getRange(row, col);
  const richTextValue = notesCell.getRichTextValue();

  if (!richTextValue) {
    return notesCell.getValue() || '';
  }

  const text = richTextValue.getText();
  if (!text) {
    return '';
  }

  let html = '';
  const runs = richTextValue.getRuns();

  for (let i = 0; i < runs.length; i++) {
    const run = runs[i];
    const runText = run.getText();
    const textStyle = run.getTextStyle();

    // Escape HTML special characters in the text
    let styledText = runText
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');

    // Check for foreground color
    const fgColor = textStyle.getForegroundColor();
    if (fgColor && fgColor !== '#000000') {
      styledText = '<span style="color:' + fgColor + ';">' + styledText + '</span>';
    }

    if (textStyle.isBold()) {
      styledText = '<b>' + styledText + '</b>';
    }
    if (textStyle.isItalic()) {
      styledText = '<i>' + styledText + '</i>';
    }
    if (textStyle.isUnderline()) {
      styledText = '<u>' + styledText + '</u>';
    }
    if (textStyle.isStrikethrough()) {
      styledText = '<s>' + styledText + '</s>';
    }

    html += styledText;
  }

  html = html.replace(/\n/g, '<br>');
  html = convertBulletsToHtml(html);

  return html;
}

/**
 * Convert checkbox characters to HTML checkbox inputs
 * ☑ becomes checked checkbox, ☐ becomes unchecked checkbox
 */
function convertCheckboxesToHtml(text) {
  if (!text) return '';

  var result = text;

  // First, handle checkboxes that might be wrapped in formatting tags like <b>☑ text</b>
  // Strip the formatting tags from checkbox lines and convert to checkbox HTML
  result = result.replace(/<(b|i|u|s)>([☑☐])\s*([^<]+)<\/\1>/g, function(match, tag, checkSymbol, labelText) {
    var isChecked = checkSymbol === '☑';
    var checkedAttr = isChecked ? ' checked' : '';
    return '<div class="checkbox-item"><input type="checkbox"' + checkedAttr + '><label>' + labelText.trim() + '</label></div>';
  });

  // Convert checked checkbox character to HTML (for unformatted checkboxes)
  // ☑ text -> <div class="checkbox-item"><input type="checkbox" checked><label>text</label></div>
  result = result.replace(/☑\s*([^\n<]+)/g, function(match, labelText) {
    return '<div class="checkbox-item"><input type="checkbox" checked><label>' + labelText.trim() + '</label></div>';
  });

  // Convert unchecked checkbox character to HTML (for unformatted checkboxes)
  // ☐ text -> <div class="checkbox-item"><input type="checkbox"><label>text</label></div>
  result = result.replace(/☐\s*([^\n<]+)/g, function(match, labelText) {
    return '<div class="checkbox-item"><input type="checkbox"><label>' + labelText.trim() + '</label></div>';
  });

  return result;
}

/**
 * Convert bullet characters to HTML list format
 * Handles multiple bullet levels: • (level 1), ○ (level 2), ■ (level 3)
 * Also handles checkboxes: ☑ (checked), ☐ (unchecked)
 */
function convertBulletsToHtml(text) {
  if (!text) return '';

  // First, convert checkbox characters to HTML checkboxes
  text = convertCheckboxesToHtml(text);

  // Strip formatting tags from bullet lines before processing
  // Handle patterns like <b>• text</b> or <b>•</b> text
  text = text.replace(/<(b|i|u|s)>([•○■▪])\s*/g, '$2 ');
  text = text.replace(/([•○■▪])\s*<\/(b|i|u|s)>/g, '$1 ');

  var lines = text.split('<br>');
  var result = [];
  var currentLevel = 0; // Track current nesting depth

  // Bullet symbols for each level
  var bulletSymbols = {
    '•': 1,  // Level 1: filled circle
    '○': 2,  // Level 2: open circle
    '■': 3,  // Level 3: square
    '▪': 3   // Alternative square
  };

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var trimmedLine = line.trim();

    // Strip leading formatting tags to check for bullet
    var strippedLine = trimmedLine.replace(/^(<(b|i|u|s)>)+/, '');

    // Check if line starts with any bullet symbol
    var bulletLevel = 0;
    var bulletChar = '';

    for (var symbol in bulletSymbols) {
      if (strippedLine.indexOf(symbol) === 0) {
        bulletLevel = bulletSymbols[symbol];
        bulletChar = symbol;
        break;
      }
    }

    if (bulletLevel > 0) {
      // This is a bullet item - extract the text after the bullet
      // Remove bullet and any surrounding formatting tags
      var itemText = strippedLine.replace(new RegExp('^[' + bulletChar + ']\\s*'), '');
      // Clean up any trailing closing tags from the bullet stripping
      itemText = itemText.replace(/^(<\/(b|i|u|s)>)+/, '').trim();

      // Adjust list nesting
      while (currentLevel < bulletLevel) {
        result.push('<ul>');
        currentLevel++;
      }
      while (currentLevel > bulletLevel) {
        result.push('</ul>');
        currentLevel--;
      }

      result.push('<li>' + itemText + '</li>');
    } else {
      // Not a bullet item - close all open lists first
      while (currentLevel > 0) {
        result.push('</ul>');
        currentLevel--;
      }
      
      // Add non-bullet content
      if (trimmedLine !== '') {
        // Add line break before content if there's already content
        if (result.length > 0) {
          var lastItem = result[result.length - 1];
          // Don't add <br> right after closing </ul>
          if (lastItem !== '</ul>') {
            result.push('<br>');
          }
        }
        result.push(line);
      } else if (result.length > 0) {
        // Empty line - only add <br> if not right after a list
        var lastItem = result[result.length - 1];
        if (lastItem !== '</ul>' && lastItem !== '<br>') {
          result.push('<br>');
        }
      }
    }
  }
  
  // Close any remaining open lists
  while (currentLevel > 0) {
    result.push('</ul>');
    currentLevel--;
  }
  
  return result.join('');
}

/**
 * Save formatted notes to a cell
 */
function saveFormattedNotes(sheet, row, col, notesText) {
  const notesCell = sheet.getRange(row, col);
  
  // Handle empty notes
  if (!notesText || notesText.trim() === '') {
    notesCell.setValue('');
    return;
  }
  
  // Handle checklists - these are already plain text with checkbox symbols
  if (notesText && (notesText.includes('☐') || notesText.includes('☑') || notesText.includes('□') || notesText.includes('☒'))) {
    notesCell.setValue(notesText);
    notesCell.setWrap(true);
    notesCell.setVerticalAlignment('top');
    return;
  }
  
  var plainText = convertHtmlToPlainText(notesText);
  
  // If plain text is empty after conversion, just return
  if (!plainText || plainText.trim() === '') {
    notesCell.setValue('');
    return;
  }
  
  // ============================================================================
  // ALL QUESTIONS ARRAY - Questions that should be bolded when saved
  // ============================================================================
  const allQuestions = [
    // Basic labels
    'Agenda Items', 'Notes',
    
    // ========================================
    // PLANNING CONVERSATION
    // ========================================
    'What goal are you working on? What is the outcome?',
    'What will it look like or sound like when you reach that goal?',
    'What are some of the things you\'ve already tried?',
    'What might be some strategies you\'re considering?',
    'What could be some of the ways I can support you?',
    
    // ========================================
    // REFLECTING CONVERSATION
    // ========================================
    'When you think about your professional practice, what might be some of the things you have been focused on?  To what extent have you met your goal?',
    'What data or information do we have that measures the progress?  When you think about these data, what are you noticing?  On what aspects of the data would you like to focus?',
    'What might need to be changed or adjusted?  What goals or outcomes do you have in mind?  What could it look like or sound like when you reach the goal?',
    'What ideas for adjustments do you have moving forward?  What might be some strategies to use?',
    'What might be some support you need from me?',
    
    // ========================================
    // PROBLEM RESOLVING CONVERSATION
    // ========================================
    'What would you like to talk about?  What is the problem?',
    'What would it look like or sound like if the problem is resolved or gets better?',
    'What ideas do you have for how to address the problem?',
    'Moving forward, what might be some actions steps you are planning to take?',
    
    // ========================================
    // GENERAL COACHING
    // ========================================
    'What\'s going well?',
    'What are some of your current challenges?',
    'What is your plan moving forward?',
    'What might be some ways I could support you?',
    
    // ========================================
    // OBSERVATION
    // ========================================
    'Areas of Strength',
    'Opportunities for Improvement',
    
    // ========================================
    // PDSA
    // ========================================
    'Plan', 'Do', 'Study', 'Plus', 'Delta', 'Act',
    
    // ========================================
    // AM TASKS
    // ========================================
    'Inbox / Email', 'Update Calendar', 'Pull Appointments from Calendar',
    'Calculate Mileage', 'Feedback Emails', 'Appointment Reminders',
    'Prep for Appointments', 'Data Team Reminder Emails', 'Import Data Team Data',
    'Prep for Data Teams', 'Data Team Reminders', 'KickUp - Create Logs',
    'Student Specific - Prep',
    
    // ========================================
    // PM TASKS
    // ========================================
    'Submit Mileage', 'KickUp - Submit Logs',
    'Update Tracker', 'Tasks from Appointments', 'Student Specific Reports',
    'Observation Reports', 'IEP Alignment Form', 'Coaching Reports',
    'Update Teacher Data Sheets', 'Move Teacher and Student Files',
    'Push Data Team Rows to Teacher Tabs', 'Archive Appointments',
    'Push Appointments to Calendar', 'Pull Appointments for Tomorrow',
    
    // ========================================
    // OTHER TASKS
    // ========================================
    'Code - Debug, Adjust, Create', 'Memorandum of Conversation',
    'Other Task (Division, Professional Responsibility, Etc.)',
    
    // ========================================
    // WEEKEND TASKS
    // ========================================
    'Data Entry', 'Data Analysis', 'Division Task', 'HSD Task', 'Coding',
    'Review Upcoming Action Plans', 'Review Previous Notes for Upcoming Appointments',
    
    // ========================================
    // DATA TEAM
    // ========================================
    'Data Import', 'Data Team Notes Prepared', 'Data Team Reporting Form',
    'Data Team Fidelity Assessment', 'Push Data Team to Teacher Tabs',
    'Reading Goal', 'Math Goal', 'Social Emotional Goal', 'Written Expression Goal',
    'Notes:',
    
    // ========================================
    // PEER OBSERVATION
    // ========================================
    'Host Teacher:', 'Focus:', 'Next Steps:',
    
    // ========================================
    // IMPACT CYCLE
    // ========================================
    'IDENTIFY', 'IMPROVE',
    'On a scale of 1-10, with 1 being the worst lesson you\'ve ever taught and 10 being your ideal or best lesson, how close was that lesson to your ideal?',
    'What pleased you about this lesson?',
    'What would you have to change to make the lesson closer to a 10?',
    'What would you see you students doing differently if your class was a 10?',
    'Tell me more about what that would look like.',
    'What teaching strategy can you use to hit your goal?',
    'What are your next steps?',
    'When should we meet again?',
    'What tasks have to be done before we meet?',
    'When will those tasks be done?',
    'Who will do them?',
    'What has gone well?',
    'What are you seeing that shows this strategy is successful?',
    'What progress has been made toward the goal?',
    'What did you learn?',
    'What surprised you?',
    'What roadblocks are you running into?',
    
    // ========================================
    // STUDENT SPECIFIC OBSERVATION
    // ========================================
    'Name of Student:',
    'What is the purpose for the request or the concern that is being addressed?',
    'To what extent are the Universal Classroom Supports in place?',
    'What IEP accommodations or behavior intervention plan strategies are expected?',
    'To what extent is the IEP implemented with fidelity?',
    
    // ========================================
    // REMINDERS
    // ========================================
    'Reminder sent',
    
    // ========================================
    // QUESTION BANK - Setting Direction
    // ========================================
    'What are some of the things that are going well?',
    'What might be some opportunities for improvement?',
    'What content or skills are students currently learning?',
    'What data are you using to measure their achievement?',
    'What might be some goals you have for your students?',
    'What kind of support would you and your students benefit from?',
    'What is the goal you\'re working toward with your students?',
    'What does success look like for this lesson/unit?',
    'What specific student outcomes are you hoping to see?',
    'How will you know when students have mastered this skill?',
    
    // ========================================
    // QUESTION BANK - Planning
    // ========================================
    'What are your goals or outcomes?',
    'How will you know you are successful? What might success look or sound like?',
    'What evidence will you collect?',
    'What data could I collect for you?',
    'If you could videotape this lesson, what would you want to see or hear in yourself?',
    'Given the data, what content, skill, or IEP goals might be your focus?',
    'What do the data tell you about students\' current achievement in that area?',
    'What changes would you like to see in the next month? What might be a reasonable goal?',
    'In thinking about that goal, what could be some strategies you are considering?',
    'What evidence will you collect to measure students\' achievement?',
    'What support might be helpful for you and your students?',
    'How will the learning targets be introduced to the students?',
    'How will we record what we notice about student learning (during the lesson and after)?',
    'How will feedback be provided to students?',
    'How will the students self-assess?',
    
    // ========================================
    // QUESTION BANK - Reflecting
    // ========================================
    'How do you think it went? How would you rate the lesson on a scale of 1-10?',
    'How does the lesson compare with what you had planned or envisioned?',
    'What do you want to stay mindful of from now on?',
    'Given the data, what might be your next steps?',
    'During our last conversation, you selected _____ as a goal for you and your students. How did it go?',
    'What does the evidence tell you about students\' current achievement? What do they know now that they didn\'t know before?',
    'What strategies did you use to help your students get to this point? What strategies weren\'t as effective?',
    'Given the evidence, what might be your next steps?',
    'What worked? What would we tweak? Where do we go from here?',
    'What did we notice about student learning?',
    'What did we notice about instructional practices as connected to student learning?',
    'How does this inform our next steps?',
    
    // ========================================
    // QUESTION BANK - Problem-Resolving
    // ========================================
    'What can you tell me about what\'s going on?',
    'What is it that you want? How do you want things to be?',
    'What are some ideas that you have?',
    'How has this conversation helped your thinking?',
    
    // ========================================
    // QUESTION BANK - Data
    // ========================================
    'What are you noticing about the data?',
    'What can we celebrate?',
    'What are some opportunities for improvement?',
    'What might be a reasonable goal?',
    'What strategies will we use?',
    'What does the data tell us about student learning?',
    'What patterns or trends do you notice?',
    'Which students are making progress? Which need additional support?',
    'How does this data compare to previous assessments?',
    'What instructional changes might address the gaps?',
    'What additional data might be helpful to collect?',
    
    // ========================================
    // QUESTION BANK - HLP 7
    // ========================================
    'What will respect, responsibility, best effort, etc., mean or look like in your classroom, and what informed your responses?',
    'In the classroom and school in which you work, do the expectations, rules, and procedures reflect the cultures, values, and beliefs of the students and families you serve?',
    'Why should school staff provide students with a rationale for the importance of expectations, rules, and procedures?',
    'Why should educators treat behavioral challenges the same way we treat academic challenges?',
    'In what ways are addressing behavioral challenges and academic challenges similar?',
    'Posts 3-5 classroom rules that are positively stated and referenced when applicable.',
    'Classroom routines for all members of the classroom (students and adults) are evident, with a minimum of entering and exiting procedures.',
    'Implements classroom procedures in a positive manner in which the teacher is always the exemplar in treating everyone with dignity.',
    'The allocated time is structured and focused on instruction.',
    'Gives specific, clear positive feedback at a ratio at least 4:1 when compared to corrective feedback that includes academic, behavior, and social-emotional.',
    'The classroom and other high-traffic areas are designed in a way that will meet the needs of the students.',
    
    // ========================================
    // QUESTION BANK - HLP 16
    // ========================================
    'As you think about this lesson, what would you like students to know and be able to do at the end? How will you share this with students?',
    'What might be the skills students already know that will be needed for this lesson? In what ways will they briefly practice these skills?',
    'When you envision the model, what do you see yourself doing or hear yourself saying? How will students know you\'re "thinking aloud"?',
    'What might be some activities you are considering to help students practice the new skill? What level of mastery are you looking for?',
    'In what way will you review the lesson goal, and how will you help students self-assess their understanding of the new skill?',
    'As you think about the goal for this lesson, to what extent do you think students met the expectation? How do you know?',
    'During the model, what did you do or say that made the steps of the new skill evident for students? How do you know they understand the skill?',
    'How did your students self-assess their learning? Was their thinking aligned with yours in relation to their skill level?',
    'Why is providing students with a clear goal and statement of expectations, purpose, and rationale for each lesson important?',
    'What are your strengths and areas you feel you need to improve when implementing explicit instruction? How do you plan to address your areas of improvement?',
    'Why is keeping a brisk pace considered a key element of explicit instruction?',
    'Why are frequent opportunities for student response so critical to an explicit lesson? What strategies can be used to elicit student responses?',
    
    // ========================================
    // QUESTION BANK - Universal Checklist
    // ========================================
    'What plans do you have for effectively designing the classroom environment to support the learning outcomes (centers, arranging desks, lighting…)?',
    'How will you make sure the classroom environment is culturally relevant, enriched with materials that are inclusive of all students?',
    'What ideas do you have for creating brief opportunities for movement within the room?',
    'What might be some of the rules you are considering?',
    'How will you make sure the rules are posted and represent the language and communication level of the students?',
    'In what ways will the rules connect to school-wide expectations?',
    'What are some of the routines you are planning to put in place?',
    'How will you make sure schedules are posted in various methods that represent the language and communication level of the students?',
    'Are there any individual student schedules that you will need to develop?',
    'What plan do you have for reinforcing appropriate behavior?',
    
    // Universal Checklist template items
    '1. Effectively design the classroom environment that supports the learning outcomes (Centers, arranging desks, lighting….)',
    '• The classroom has clearly defined learning spaces',
    '• A clear path for mobility for all students',
    '• All materials are accessible to all students',
    '• Classroom is clear of clutter',
    '2. The classroom environment is culturally relevant, enriched with materials that are culturally diverse.',
    '• Classrooms have diverse representation of cultures and abilities, through materials, props, clothing, and languages on visuals',
    '• The classroom has readily available culturally relevant materials including books, pictures, visuals etc.',
    '3. There are opportunities for movement within the room.',
    '• Students have built in brain breaks and opportunities for movement during whole and small group times, transitions, or a specific classroom space is dedicated to large muscle movement.',
    '• The teacher provides intentional opportunities for movement that are based on learning goals.',
    '4. Classroom Routines:',
    '• Developed for each part of the day',
    '• Connected to school-wide expectations',
    '• Include student voice',
    '• Adult (TAs, SLPs, classroom teachers…) routines exist',
    '• Taught and practiced',
    '• A variety of routines are posted.',
    '• Classroom staff reference the visual routines as a prompt for students as needed.',
    '• There is evidence of teaching as students follow routines for various times of the day.',
    '• Transitions are not disruptive to the learning environment',
    '• Instruction occurs with minimal downtime',
    '6. Class Schedule:',
    '• A classroom schedule is accessible in all environments.',
    '• Students are allowed to manipulate the schedule as activities are completed.',
    '• The teacher has a schedule within a schedule or activity routines visually represented.',
    '• Schedules use modes that meet student needs',
    '• Schedule varies preferred with non-preferred activities',
    '7. Classroom rules:',
    '• Three to five positively stated expectations/rules posted in a variety of locations/settings.',
    '• Rules have been taught, practiced and referenced.',
    '• Students are able to demonstrate understanding of rules verbally/gesturally/behaviorally with minimal re-directs.',
    '• Teaching of rules is expanded and accessible throughout the day as needed.',
    '• Expectations/rules posted in mode matched to student need',
    '• Reviewed prior to instruction, and are referenced throughout instruction',
    
    // ========================================
    // QUESTION BANK - Classroom Learning System (CCI)
    // ========================================
    // CCI PLANNING
    'What might be the learning requirements for this class?',
    'In what ways can you make sure the learning goal is related to the learning requirements?',
    'What data might you use to determine the current level of the class?',
    'How do you plan to engage your students in creating a mission statement?',
    'How will you choose a learning target for this cycle?',
    'What might be some of the strategies you are considering adding to the strategy bank?',
    'How do you plan to involve students when analyzing the data?',
    'How will you help students to decide on adjustments?',
    // CCI REFLECTING
    'How did you decide on the learning requirements for the class?',
    'In what ways does your class goal relate to the learning requirements?',
    'How did you share the learning results with the class?',
    'What activity did you choose to facilitate the creation of the mission statement?',
    'How did you decide on the target for this learning cycle?',
    'Which strategies did the class choose?',
    'What process did you use to analyze the data with students?',
    'What adjustments did the students want to make for the next learning cycle?',
    
    // CCI Template items
    '1. The teacher sets and communicates direction for students and their families',
    'Learning Requirements',
    '• Posted',
    '• Kid Friendly Language',
    '• Teacher mentions learning requirement in lesson',
    'Class Goal',
    '• Long term',
    '• Student friendly language/representations',
    '• Teacher refers to class goal',
    '• Goal is measurable',
    'Learning Results',
    '• Chart/Visual for hard data',
    '• Teacher refers to learning results',
    '• Teacher refers to analyzing the learning results',
    'Mission Statement',
    '• Adult Language/generic Student Language',
    '• Consensus is visually shown (signed, consensogram)',
    '• Teacher references mission statement',
    '• Students communicate mission statement',
    '2. The teacher engages students in regular and frequent evaluation and improvement of classroom learning processes.',
    '• Learning target posted',
    '• Measure of proficiency included',
    '• Clearly leads to achievement of the class goal',
    '• Teacher shares learning target with students',
    '• High yield strategies posted',
    '• Teacher is using the high yield strategies',
    '• Strategies include the role of the teacher and student',
    '• Students are included in picking the strategies used',
    '• A strategy bank is posted/used',
    '• Teacher refers to/names the high yield strategies and/or why the class is using them',
    '• Chart for posting results of learning target proficiency',
    '• Evidence of student feedback is posted',
    '• Students participate in a plus/delta or other form of feedback',
    '• Students analyze the data',
    '• Teacher refers to the learning results',
    '• Teacher refers to plus/delta',
    '• Documentation of adjustment in the learning cycle',
    '• Based on feedback from the previous cycle results',
    '• Kid language/input',
    '• Teacher refers to adjustments',
    '• Teacher reminds students of their/the class adjustments',
    
    // ========================================
    // QUESTION BANK - Student
    // ========================================
    'What does the student like to do? What are they good at?',
    'What specific behavior will we focus on first? Be as detailed and observable as possible.',
    'Why do we think this behavior is occurring? What purpose does it serve for the student?',
    'What strategies are we currently using that are working well?',
    'What specific behavior do we want to see in the next 4 weeks? Make it measurable and achievable.',
    'How will we teach the desired replacement behavior?',
    'What strategies will prevent the problem behavior from occurring?',
    'How will we reinforce and maintain the new positive behavior?',
    'What will we do when the target behavior occurs?',
    'What needs to happen next? What follow-up actions or additional supports are needed?',
    'What data will we collect to monitor progress on this plan?',
    
    // ========================================
    // QUESTION BANK - Meeting
    // ========================================
    'What are our objectives for today\'s meeting?',
    'What items need to be addressed on our agenda?',
    'What decisions need to be made today?',
    'Who needs to be involved in this discussion?',
    'What information do we need before making this decision?',
    'Are there any concerns or barriers we should discuss?',
    'How can we ensure everyone\'s voice is heard?',
    'What are the next steps and who is responsible?',
    'When do we need to follow up on this action item?',
    'Is there anything we need to table for the next meeting?',
    'What was most valuable about today\'s meeting?',
    'Do we need to schedule a follow-up meeting?',
    'Are there any resources or supports needed to move forward?',
    'How will we communicate this information to others who need to know?',
    
    // ========================================
    // QUESTION BANK - Organize and Integrate (Debrief)
    // ========================================
    'What are some things that have gone well?',
    'What might be some of the things you have done that led to the outcomes?',
    'Moving forward, what are some things you will pursue?',
    'In what ways has our collaboration supported your practice?',
    'Where are you now in your practice compared to where you were when we started?',
    'How did your students benefit from this coaching cycle?',
    'How did you benefit from this coaching cycle?',
    'Which coaching practices were most useful / least useful to you?',
    'What did the assessment data reveal at the end of the coaching cycle?',
    'Do we need to adjust and continue our coaching cycle?',
    'What are your next steps for student learning in this area?',
    'When can we meet to touch base about how this work has continued for you?',
    'What other feedback or questions do you have?',
    
    // ========================================
    // EXPLICIT INSTRUCTION TEMPLATES
    // ========================================
    'Part 1 How do you know you have created the right objectives and provided effective modeling?',
    'The methods used to create objectives should:',
    'Choose objectives based on student performance relative to goals.',
    '☐ Select a goal from IEP or standards',
    '☐ Choose an objective that is the next step toward the goal',
    'Write focused objectives that describe the specific learning outcome.',
    '☐ Limit the objective to one singular next step toward the goal',
    '☐ Describe a learning outcome in behavioral terms that assesses mastery of the objective',
    'The methods used to provide modeling should:',
    'Give clear explanations',
    '• Match the explanation to the learning outcome',
    '• Design the explanation so that it is correct, clear and concise',
    '• Use the explanation consistently',
    'Model multiple planned examples',
    '• Show all steps or provide unique examples',
    '• Verbalize your thinking',
    '• Have students observe',
    'Use supporting practices',
    
    'Part 2 How do you know you have effectively implemented practice?',
    'The methods used to provide guided practice should:',
    'Lead students in steps toward the learning outcome',
    '• Procedural task: execute each step separately',
    '• Knowledge task: address each unit (e.g. vocabulary) one at a time',
    'Provide appropriate prompts',
    '• Design a variety of prompt types linked to task and likely student need',
    '• Change level of prompting in response to student\'s progress',
    'Observe and provide immediate feedback',
    '• Watch students carefully',
    '• Interpret the meaning of errors',
    '• Provide feedback that aligns with the type of error',
    'The methods used to provide independent practice should:',
    'Review expectations and resources for meeting the learning outcome',
    '• Select objective-linked practice items that lead to 90-95% accuracy',
    '• Review expectations for meeting outcomes',
    '• Remind students how they self-prompt',
    'Allow student to work without support',
    '• Limit guidance from teacher',
    '• Monitor throughout independent practice',
    '• Give mini-reminders and record points of confusion',
    'Observe and provide immediate and delayed feedback',
    '• Check for understanding throughout',
    '• Provide feedback following completion of task',
    
    'Part 3 How do you know you have effectively used the supporting practices?',
    'The methods used to elicit a response should:',
    '• Maintain or check accuracy of processing',
    '• Match the learning outcome',
    '• Match student abilities',
    '• Match the desired response format',
    '• Maximize student involvement',
    'The methods used to provide feedback should be:',
    '• Immediate: delivered as soon as possible after response',
    '• Specific: tied directly to students\' actions',
    'The methods used to maintain a brisk pace should:',
    '• Move on when students are ready',
    '• Use the other supporting practices',
    'The methods used to encourage student reflection and data collection should:',
    '• Encourage students to self-assess their use of the skill',
    '• Provide relevant assessment data directly aligned with the lesson outcome',

    // ========================================
    // PROFESSIONAL LEARNING
    // ========================================
    // Domain 1: Needs Assessment & Goal Setting
    'Data Analysis: What data sources are being used to identify professional learning needs?',
    'Contextual Factors: What contextual factors (school culture, resources, time) have been considered?',
    'SMART Goals: Are the professional learning goals Specific, Measurable, Achievable, Relevant, and Time-bound?',
    'Alignment: How does this professional learning align with school/district priorities and teacher needs?',
    
    // Domain 2: Designing Learning Experiences
    'Learning Objectives: What specific learning objectives will guide this professional learning experience?',
    'Learning Design: How is the professional learning designed to be engaging, relevant, and applicable?',
    'Resources: What resources (materials, technology, facilitators) are needed to support this learning?',
    'Timeline: What is the timeline for implementing this professional learning, including follow-up?',
    'Evaluation: How will we evaluate the effectiveness of the professional learning design?',
    
    // Domain 3: Implementation & Support
    'Communication: How has the purpose and expectations of this professional learning been communicated?',
    'Follow-up: What follow-up support will be provided after the initial learning session?',
    'Collaboration: How are opportunities for collaboration and peer learning being incorporated?',
    
    // Domain 4: Evaluation & Reflection
    'Data Collection: What data will be collected to measure the impact of this professional learning?',
    'Reflection: How are participants being given opportunities to reflect on their learning and practice?',
    'Review & Revision: How will the professional learning be reviewed and revised based on feedback and outcomes?',
    
    // Domain 5: Key Considerations
    'Focus on Student Learning: How does this professional learning directly connect to improving student outcomes?',
    'Evidence-Based Practices: Is this professional learning grounded in research and evidence-based practices?',
    'Continuous Improvement: How does this fit into a cycle of continuous improvement for professional practice?',
    
    // ========================================
    // TASKS COMPLETION CHECKLIST
    // ========================================
    // Domain 1: Scope & Requirements
    'Does the final output directly address the core problem or goal defined at the start of the task?',
    'Have I created, checked, and included all components (reports, code, designs, documentation, etc.) specified in the original request?',
    'Did I avoid significant "scope creep" (doing extra, unrequested work) or "scope shrinkage" (missing required elements)?',
    'Has the person who requested or will use this task\'s output formally approved the result?',
    
    // Domain 2: Quality & Validation
    'Does the deliverable work correctly and accurately under all specified conditions (e.g., is the data correct, does the code run without errors)?',
    'Does the output meet all relevant internal standards (e.g., branding, coding styles, safety protocols)?',
    'Have I performed sufficient testing or proofreading, and are there zero critical or major known errors remaining?',
    'Is the result clear, intuitive, and easy to use or understand for the intended audience?',
    
    // Domain 3: Efficiency & Documentation
    'Was the task completed within the estimated time or budget, and if not, have I documented the reasons?',
    'Did I use resources (e.g., human hours, materials, software licenses) efficiently, without waste?',
    'Is there clear documentation (e.g., a "how-to" guide, final meeting notes, a process diagram) so someone else could repeat or update this task?',
    'Have I identified at least one thing that went well and one thing that could be improved for future similar tasks?'
  ];
  
  // Set the plain text value first
  notesCell.setValue(plainText);

  try {
    const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(plainText);
    const textLength = plainText.length;

    // Parse HTML formatting with bounds checking
    var formattingData = parseHtmlFormatting(notesText);

    for (var i = 0; i < formattingData.length; i++) {
      var format = formattingData[i];

      // Bounds checking - skip invalid ranges
      if (format.start < 0 || format.end < 0 ||
          format.start >= textLength || format.end > textLength ||
          format.start >= format.end) {
        continue;
      }

      var textStyle = SpreadsheetApp.newTextStyle();

      if (format.bold) textStyle.setBold(true);
      if (format.italic) textStyle.setItalic(true);
      if (format.underline) textStyle.setUnderline(true);
      if (format.strikethrough) textStyle.setStrikethrough(true);

      // Handle text colors
      if (format.color) {
        var colorValue = parseColorToHex(format.color);
        if (colorValue) {
          textStyle.setForegroundColor(colorValue);
        }
      }

      // For background colors (highlights), use a colored text as workaround
      // since Google Sheets RichText doesn't support per-character backgrounds
      // We'll convert highlights to colored text markers
      if (format.backgroundColor && !format.color) {
        // Convert background highlight to a visible foreground color indicator
        var bgColorValue = parseColorToHex(format.backgroundColor);
        // Use dark colors for text that had light highlights
        if (bgColorValue) {
          // Keep the text color as-is but we'll mark it
          // Note: True highlight preservation requires storing HTML separately
        }
      }

      richTextBuilder.setTextStyle(format.start, format.end, textStyle.build());
    }

    // Apply bold formatting to known questions
    for (var i = 0; i < allQuestions.length; i++) {
      var question = allQuestions[i];
      var searchIndex = 0;
      var foundIndex = plainText.indexOf(question, searchIndex);

      while (foundIndex !== -1) {
        var endIndex = foundIndex + question.length;

        // Bounds checking
        if (foundIndex >= 0 && endIndex <= textLength && foundIndex < endIndex) {
          richTextBuilder.setTextStyle(
            foundIndex,
            endIndex,
            SpreadsheetApp.newTextStyle().setBold(true).build()
          );
        }

        searchIndex = endIndex;
        foundIndex = plainText.indexOf(question, searchIndex);
      }
    }

    notesCell.setRichTextValue(richTextBuilder.build());

  } catch (formatError) {
    // If rich text formatting fails, just keep the plain text that was already set
    Logger.log('Warning: Rich text formatting failed, using plain text: ' + formatError.toString());
  }

  notesCell.setWrap(true);
  notesCell.setVerticalAlignment('top');
}

/**
 * Parse color string (rgb, hex, or named) to hex format
 */
function parseColorToHex(colorStr) {
  if (!colorStr) return null;

  colorStr = colorStr.trim().toLowerCase();

  // Already hex format
  if (colorStr.charAt(0) === '#') {
    return colorStr;
  }

  // RGB format: rgb(r, g, b)
  var rgbMatch = colorStr.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
  if (rgbMatch) {
    var r = parseInt(rgbMatch[1]).toString(16).padStart(2, '0');
    var g = parseInt(rgbMatch[2]).toString(16).padStart(2, '0');
    var b = parseInt(rgbMatch[3]).toString(16).padStart(2, '0');
    return '#' + r + g + b;
  }

  // Named colors (common ones)
  var namedColors = {
    'red': '#ff0000',
    'green': '#00ff00',
    'blue': '#0000ff',
    'yellow': '#ffff00',
    'orange': '#ffa500',
    'purple': '#800080',
    'pink': '#ffc0cb',
    'black': '#000000',
    'white': '#ffffff',
    'gray': '#808080',
    'grey': '#808080',
    'cyan': '#00ffff',
    'magenta': '#ff00ff'
  };

  return namedColors[colorStr] || null;
}

/**
 * Convert HTML to plain text with nested list support
 * Level 1: • (filled circle)
 * Level 2: ○ (open circle)  
 * Level 3: ■ (square)
 */
function convertHtmlToPlainText(html) {
  if (!html) return '';
  
  var text = html;
  
  // First, normalize br tags to newlines
  text = text.replace(/<br\s*\/?>/gi, '\n');
  
  // Remove other common HTML entities
  text = text.replace(/&nbsp;/gi, ' ');
  text = text.replace(/&amp;/gi, '&');
  text = text.replace(/&lt;/gi, '<');
  text = text.replace(/&gt;/gi, '>');
  text = text.replace(/&quot;/gi, '"');
  
  // Track nesting level for lists
  var nestingLevel = 0;
  var result = '';
  var i = 0;
  var inListItem = false;
  
  while (i < text.length) {
    // Check for opening ul tag
    if (text.substr(i, 4).toLowerCase() === '<ul>') {
      nestingLevel++;
      i += 4;
      continue;
    }

    // Check for closing ul tag
    if (text.substr(i, 5).toLowerCase() === '</ul>') {
      nestingLevel--;
      if (nestingLevel < 0) nestingLevel = 0;
      i += 5;
      continue;
    }
        
    // Check for opening li tag
    if (text.substr(i, 4).toLowerCase() === '<li>') {
      // Add newline before bullet if there's content and we're not at the start
      if (result.length > 0 && result[result.length - 1] !== '\n') {
        result += '\n';
      }
      
      // Add appropriate bullet based on nesting level
      var bullet = '• ';  // Level 1: filled circle
      if (nestingLevel === 2) {
        bullet = '○ ';    // Level 2: open circle
      } else if (nestingLevel >= 3) {
        bullet = '■ ';    // Level 3+: square
      }
      
      result += bullet;
      inListItem = true;
      i += 4;
      continue;
    }
    
    // Check for closing li tag
    if (text.substr(i, 5).toLowerCase() === '</li>') {
      inListItem = false;
      // Add newline after list item
      if (result.length > 0 && result[result.length - 1] !== '\n') {
        result += '\n';
      }
      i += 5;
      continue;
    }
    
    // Check for other HTML tags
    if (text[i] === '<') {
      var tagEnd = text.indexOf('>', i);
      if (tagEnd !== -1) {
        var tagContent = text.substring(i + 1, tagEnd);
        var tagContentLower = tagContent.toLowerCase();

        // Handle checkbox inputs
        if (tagContentLower.indexOf('input') !== -1 && tagContentLower.indexOf('checkbox') !== -1) {
          // Check if the checkbox is checked
          if (tagContentLower.indexOf('checked') !== -1) {
            result += '☑ ';  // Checked checkbox
          } else {
            result += '☐ ';  // Unchecked checkbox
          }
          i = tagEnd + 1;
          continue;
        }

        // Handle block-level tags that should add newlines
        if (tagContentLower === '/p' || tagContentLower === '/div') {
          if (result.length > 0 && result[result.length - 1] !== '\n') {
            result += '\n';
          }
        }

        // Handle label tags - often wrap checkbox text
        if (tagContentLower === '/label') {
          // Add newline after label (checkbox item complete)
          if (result.length > 0 && result[result.length - 1] !== '\n' && result[result.length - 1] !== ' ') {
            result += '\n';
          }
        }

        // Skip past the tag
        i = tagEnd + 1;
        continue;
      }
    }
    
    // Handle newline characters
    if (text[i] === '\n') {
      // Only add newline if we're not in a list item (list items handle their own newlines)
      // and if the last character isn't already a newline
      if (!inListItem && (result.length === 0 || result[result.length - 1] !== '\n')) {
        result += '\n';
      } else if (inListItem) {
        // Inside a list item, convert newline to space to keep content together
        if (result.length > 0 && result[result.length - 1] !== ' ') {
          result += ' ';
        }
      }
      i++;
      continue;
    }
    
    // Regular character - add it
    result += text[i];
    i++;
  }
  
  // Clean up: remove excessive newlines (max 2 in a row)
  result = result.replace(/\n{3,}/g, '\n\n');
  
  // Remove leading/trailing whitespace but preserve internal structure
  result = result.replace(/^\n+/, '');
  result = result.replace(/\n+$/, '');
  
  return result;
}

/**
 * Parse HTML formatting including colors and highlights
 */
function parseHtmlFormatting(html) {
  if (!html) return [];

  var formattingRanges = [];
  var plainText = convertHtmlToPlainText(html);

  var stack = [];
  var currentPlainPos = 0;
  var currentHtmlPos = 0;

  // Track list nesting for accurate position calculation
  var listNestingLevel = 0;

  while (currentHtmlPos < html.length) {
    var char = html.charAt(currentHtmlPos);

    if (char === '<') {
      var tagEnd = html.indexOf('>', currentHtmlPos);
      if (tagEnd === -1) break;

      var fullTag = html.substring(currentHtmlPos + 1, tagEnd);
      var isClosing = fullTag.charAt(0) === '/';

      // Extract tag name (first word before space or end)
      var tagContent = isClosing ? fullTag.substring(1) : fullTag;
      var spaceIndex = tagContent.indexOf(' ');
      var tagName = (spaceIndex > 0 ? tagContent.substring(0, spaceIndex) : tagContent).toLowerCase();

      if (tagName === 'strong') tagName = 'b';
      if (tagName === 'em') tagName = 'i';

      // Handle span tags with style attributes (for colors/highlights)
      if (tagName === 'span' && !isClosing) {
        var styleMatch = fullTag.match(/style\s*=\s*["']([^"']*)["']/i);
        var spanInfo = { tag: 'span', startPos: currentPlainPos };

        if (styleMatch) {
          var styleStr = styleMatch[1];

          // Check for text color
          var colorMatch = styleStr.match(/(?:^|;)\s*color\s*:\s*([^;]+)/i);
          if (colorMatch) {
            spanInfo.color = colorMatch[1].trim();
          }

          // Check for background color (highlight)
          var bgMatch = styleStr.match(/background(?:-color)?\s*:\s*([^;]+)/i);
          if (bgMatch) {
            spanInfo.backgroundColor = bgMatch[1].trim();
          }
        }

        stack.push(spanInfo);
      } else if (tagName === 'span' && isClosing) {
        // Find matching span in stack
        for (var i = stack.length - 1; i >= 0; i--) {
          if (stack[i].tag === 'span') {
            var format = {
              start: stack[i].startPos,
              end: currentPlainPos
            };

            if (stack[i].color) {
              format.color = stack[i].color;
            }
            if (stack[i].backgroundColor) {
              format.backgroundColor = stack[i].backgroundColor;
            }

            if (format.end > format.start && (format.color || format.backgroundColor)) {
              formattingRanges.push(format);
            }
            stack.splice(i, 1);
            break;
          }
        }
      } else if (!isClosing && (tagName === 'b' || tagName === 'i' || tagName === 'u' || tagName === 's')) {
        stack.push({
          tag: tagName,
          startPos: currentPlainPos
        });
      } else if (isClosing && (tagName === 'b' || tagName === 'i' || tagName === 'u' || tagName === 's')) {
        for (var i = stack.length - 1; i >= 0; i--) {
          if (stack[i].tag === tagName) {
            var format = {
              start: stack[i].startPos,
              end: currentPlainPos,
              bold: (tagName === 'b'),
              italic: (tagName === 'i'),
              underline: (tagName === 'u'),
              strikethrough: (tagName === 's')
            };
            if (format.end > format.start) {
              formattingRanges.push(format);
            }
            stack.splice(i, 1);
            break;
          }
        }
      } else if (tagName === 'ul') {
        listNestingLevel++;
      } else if (tagName === '/ul') {
        listNestingLevel--;
        if (listNestingLevel < 0) listNestingLevel = 0;
      } else if (tagName === 'br' || tagName === 'br/') {
        currentPlainPos++;
      } else if (tagName === 'li') {
        // Account for indentation and bullet character
        var indentSpaces = (listNestingLevel - 1) * 4;
        currentPlainPos += indentSpaces + 2; // indent + bullet + space
      } else if (tagName === '/li') {
        currentPlainPos++;
      }

      currentHtmlPos = tagEnd + 1;
    } else {
      currentPlainPos++;
      currentHtmlPos++;
    }
  }

  return formattingRanges;
}

// ============================================================================
// CALENDAR SEARCH FUNCTION
// ============================================================================

/**
 * Search the 'ETA' calendar for future appointments matching the given name
 * @param {string} searchName - The name to search for in calendar events
 * @returns {Object} Object containing success status and array of appointments
 */
function getFutureAppointmentsFromCalendar(searchName) {
  try {
    if (!searchName || searchName.trim() === '') {
      return { success: false, message: 'No name provided', appointments: [] };
    }
    
    Logger.log('Searching ETA calendar for: ' + searchName);
    
    // Get the ETA calendar by name
    const calendars = CalendarApp.getCalendarsByName('ETA');
    
    if (!calendars || calendars.length === 0) {
      Logger.log('ETA calendar not found');
      return { success: false, message: 'ETA calendar not found', appointments: [] };
    }
    
    const etaCalendar = calendars[0];
    Logger.log('Found ETA calendar: ' + etaCalendar.getName());
    
    // Define the search range: from tomorrow to June 1, 2026
    const now = new Date();
    const tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
    tomorrow.setHours(0, 0, 0, 0); // Start at midnight tomorrow
    const futureDate = new Date(2026, 5, 1); // June 1, 2026 (months are 0-indexed, so 5 = June)
    
    Logger.log('Searching from: ' + tomorrow.toString() + ' to: ' + futureDate.toString());
    
    // Get all events in the date range
    const events = etaCalendar.getEvents(tomorrow, futureDate);
    Logger.log('Found ' + events.length + ' total events in date range');
    
    // Filter events that contain the search name (case-insensitive)
    const searchNameLower = searchName.toLowerCase().trim();
    
    // Get the first word (most unique identifier - e.g., "Barrington" from "Barrington Data Team" or "Amanda" from "Amanda Barton")
    const nameParts = searchName.trim().split(' ');
    const firstWord = nameParts[0] ? nameParts[0].toLowerCase() : '';
    
    // For 2-word names (like "Amanda Barton"), also get the second word as a fallback
    // But skip generic words like "Data", "Team", "Meeting" for matching
    const genericWords = ['data', 'team', 'meeting', 'the', 'a', 'an'];
    let secondWord = '';
    if (nameParts.length === 2 && !genericWords.includes(nameParts[1].toLowerCase())) {
      secondWord = nameParts[1].toLowerCase();
    }
    
    Logger.log('Searching for full name: "' + searchNameLower + '", first word: "' + firstWord + '", second word: "' + secondWord + '"');
    
    const matchingAppointments = [];
    
    // Log the first 10 event titles for debugging
    Logger.log('=== First 10 upcoming event titles ===');
    for (let i = 0; i < Math.min(10, events.length); i++) {
      Logger.log('Event ' + i + ': "' + events[i].getTitle() + '"');
    }
    Logger.log('=== End of event titles ===');
    
    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      const title = event.getTitle() || '';
      const description = event.getDescription() || '';
      const titleLower = title.toLowerCase();
      const descLower = description.toLowerCase();
      
      // Check if the name appears in the title or description
      // Priority: full name match, then first word match (required for partial matches)
      const matchesFull = titleLower.includes(searchNameLower) || descLower.includes(searchNameLower);
      const matchesFirstWord = firstWord && (titleLower.includes(firstWord) || descLower.includes(firstWord));
      const matchesSecondWord = secondWord && (titleLower.includes(secondWord) || descLower.includes(secondWord));
      
      // Match if: full name matches, OR first word matches (the unique identifier like "Barrington" or "Amanda")
      // For 2-word person names, also match if second word (last name) matches
      if (matchesFull || matchesFirstWord || matchesSecondWord) {
        const startTime = event.getStartTime();
        const endTime = event.getEndTime();
        
        // Format the date and time
        const dateFormatted = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'EEE, MMM d');
        
        let timeFormatted = '';
        if (event.isAllDayEvent()) {
          timeFormatted = 'All Day';
        } else {
          const startFormatted = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'h:mm a');
          const endFormatted = Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'h:mm a');
          timeFormatted = startFormatted + ' - ' + endFormatted;
        }
        
        matchingAppointments.push({
          date: dateFormatted,
          time: timeFormatted,
          title: title,
          description: description.substring(0, 100) + (description.length > 100 ? '...' : ''),
          startDateTime: startTime.getTime() // For sorting
        });
        
        Logger.log('Found matching event: ' + title + ' on ' + dateFormatted);
      }
    }
    
    // Sort by date (earliest first)
    matchingAppointments.sort(function(a, b) {
      return a.startDateTime - b.startDateTime;
    });
    
    // Limit to first 5 appointments to keep the display manageable
    const limitedAppointments = matchingAppointments.slice(0, 5);
    
    Logger.log('Returning ' + limitedAppointments.length + ' matching appointments');
    
    return {
      success: true,
      appointments: limitedAppointments,
      totalFound: matchingAppointments.length,
      searchName: searchName
    };
    
  } catch (error) {
    Logger.log('Error in getFutureAppointmentsFromCalendar: ' + error.toString());
    return { 
      success: false, 
      message: 'Error searching calendar: ' + error.toString(),
      appointments: []
    };
  }
}

/**
 * Test function to verify calendar access
 * Run this from the Apps Script editor to test the calendar search
 */
function testCalendarSearch() {
  // Test with a sample name - change this to a name that should exist in your calendar
  const result = getFutureAppointmentsFromCalendar('Amanda Barton');
  Logger.log('Test result: ' + JSON.stringify(result, null, 2));
  
  // Also test listing all calendars the user has access to
  const allCalendars = CalendarApp.getAllCalendars();
  Logger.log('=== Available calendars ===');
  for (let i = 0; i < allCalendars.length; i++) {
    Logger.log('  - "' + allCalendars[i].getName() + '"');
  }
  Logger.log('=== End of calendars ===');
}

/**
 * Debug function to list all upcoming events in ETA calendar
 * Run this to see what events exist and how they're titled
 */
function debugListAllETAEvents() {
  const calendars = CalendarApp.getCalendarsByName('ETA');
  
  if (!calendars || calendars.length === 0) {
    Logger.log('ETA calendar not found!');
    
    // List all calendars to help find the right name
    const allCalendars = CalendarApp.getAllCalendars();
    Logger.log('Available calendars:');
    for (let i = 0; i < allCalendars.length; i++) {
      Logger.log('  - "' + allCalendars[i].getName() + '"');
    }
    return;
  }
  
  const etaCalendar = calendars[0];
  const now = new Date();
  const futureDate = new Date(2026, 5, 1);
  
  const events = etaCalendar.getEvents(now, futureDate);
  
  Logger.log('=== All ' + events.length + ' events in ETA calendar ===');
  for (let i = 0; i < events.length; i++) {
    const event = events[i];
    const startTime = event.getStartTime();
    const dateStr = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    Logger.log(dateStr + ' - "' + event.getTitle() + '"');
  }
  Logger.log('=== End of events ===');
}


/**
 * Get teacher observation data from OBSSummary tab and previous appointment notes
 * @param {string} teacherName - The name to search for
 * @param {string} taskType - The task/type to match (optional)
 * @returns {Object} Combined observation data and previous notes HTML
 */
function getTeacherDataWithObservations(teacherName, taskType) {
  try {
    if (!teacherName || teacherName.trim() === '') {
      return { success: false, message: 'No teacher name provided' };
    }
    
    Logger.log('Getting observation data and previous notes for: ' + teacherName + ', Type: ' + taskType);
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    let observationHtml = '';
    const obsSummarySheet = ss.getSheetByName('OBSSummary');
    
    if (obsSummarySheet) {
      const lastRow = obsSummarySheet.getLastRow();
      const namesColumn = obsSummarySheet.getRange(2, 1, lastRow - 1, 1).getValues();
      
      let teacherRow = -1;
      for (let i = 0; i < namesColumn.length; i++) {
        if (namesColumn[i][0] === teacherName) {
          teacherRow = i + 2;
          break;
        }
      }
      
      if (teacherRow > -1) {
        const lastCol = obsSummarySheet.getLastColumn();
        const headers = obsSummarySheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const rowData = obsSummarySheet.getRange(teacherRow, 1, 1, lastCol).getValues()[0];
        
        const observationsData = [];
        
        for (let col = 1; col < lastCol; col += 11) {
          const dateValue = rowData[col];
          
          if (dateValue && dateValue !== '') {
            const observation = {
              date: dateValue instanceof Date ? 
                    Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'MM/dd/yyyy') : 
                    String(dateValue),
              data: []
            };
            
            for (let i = 1; i <= 10 && (col + i) < lastCol; i++) {
              const header = headers[col + i];
              const cellValue = rowData[col + i];
              const cell = obsSummarySheet.getRange(teacherRow, col + i + 1);
              
              if (header && header !== '') {
                const bgColor = cell.getBackground();
                
                let formattedValue = 'N/A';
                if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
                  const numberFormat = cell.getNumberFormat();
                  
                  if (numberFormat && numberFormat.includes('%')) {
                    formattedValue = (cellValue * 100).toFixed(0) + '%';
                  } else if (typeof cellValue === 'number') {
                    if (cellValue % 1 === 0) {
                      formattedValue = cellValue.toString();
                    } else {
                      formattedValue = cellValue.toFixed(2);
                    }
                  } else {
                    formattedValue = String(cellValue);
                  }
                }
                
                observation.data.push({
                  label: header,
                  value: formattedValue,
                  bgColor: bgColor
                });
              }
            }
            
            if (observation.data.length > 0) {
              observationsData.push(observation);
            }
          }
        }
        
        if (observationsData.length > 0) {
          observationHtml = '<div style="background-color: #f0f0f0; padding: 10px; border-radius: 5px; margin-bottom: 15px;">';
          observationHtml += '<h3 style="margin-top: 0; color: #8d6e63;">Observation Summary (' + observationsData.length + ' observation' + (observationsData.length > 1 ? 's' : '') + ')</h3>';
          
          for (let i = 0; i < observationsData.length; i++) {
            const obs = observationsData[i];
            observationHtml += '<div style="background-color: white; padding: 8px; margin-bottom: 10px; border-radius: 3px; border-left: 4px solid #8d6e63;">';
            observationHtml += '<b style="color: #8d6e63;">Observation ' + (i + 1) + ' - ' + obs.date + '</b><br>';
            observationHtml += '<table style="width: 100%; margin-top: 5px; font-size: 0.9em; border-collapse: collapse;">';
            
            for (let j = 0; j < obs.data.length; j++) {
              const item = obs.data[j];
              const cellStyle = 'padding: 4px 6px; border: 1px solid #ddd; text-align: center;' + 
                                (item.bgColor !== '#ffffff' && item.bgColor !== '#fff' ? 
                                 ' background-color: ' + item.bgColor + ';' : '');
              
              observationHtml += '<tr>';
              observationHtml += '<td style="padding: 4px 6px; width: 60%; border: 1px solid #ddd;"><i>' + item.label + ':</i></td>';
              observationHtml += '<td style="' + cellStyle + '"><b>' + item.value + '</b></td>';
              observationHtml += '</tr>';
            }
            
            observationHtml += '</table>';
            observationHtml += '</div>';
          }
          
          observationHtml += '</div>';
        }
      } else {
        Logger.log('Teacher not found in OBSSummary tab');
      }
    } else {
      Logger.log('OBSSummary tab not found');
    }
    
    const previousNotesResult = getPreviousAppointmentNotes(teacherName, taskType);
    
    let combinedHtml = '';
    
    if (observationHtml !== '') {
      combinedHtml += observationHtml;
    }
    
    if (previousNotesResult.success) {
      if (combinedHtml !== '') {
        combinedHtml += '<hr style="margin: 15px 0;">';
      }
      combinedHtml += '<h3 style="color: #8d6e63; margin-top: 0;">Previous Appointment Notes</h3>';
      if (previousNotesResult.date) {
        combinedHtml += '<p style="font-style: italic; color: #666;">Date: ' + previousNotesResult.date;
        if (previousNotesResult.type) {
          combinedHtml += ' | Type: ' + previousNotesResult.type;
        }
        combinedHtml += '</p>';
      }
      combinedHtml += previousNotesResult.html;
    }
    
    if (combinedHtml !== '') {
      return { 
        success: true, 
        html: combinedHtml,
        teacher: teacherName,
        hasObservations: observationHtml !== '',
        hasPreviousNotes: previousNotesResult.success
      };
    } else {
      return { 
        success: false, 
        message: 'No observation data or previous notes found for ' + teacherName 
      };
    }
    
  } catch (error) {
    Logger.log('Error in getTeacherDataWithObservations: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

/**
 * Test function to verify sheet access
 */
function testSheetAccess() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('Successfully accessed spreadsheet: ' + ss.getName());
    
    let appSheet = ss.getSheetByName(APP_TAB_NAME);
    if (!appSheet) {
      Logger.log('WARNING: APP sheet "' + APP_TAB_NAME + '" not found!');
    } else {
      Logger.log('Successfully accessed APP sheet: ' + APP_TAB_NAME);
      Logger.log('APP sheet has ' + appSheet.getLastRow() + ' rows');
    }
    
    let dtfSheet = ss.getSheetByName(TAB_NAME);
    if (!dtfSheet) {
      Logger.log('Tab "' + TAB_NAME + '" does not exist. It will be created on first form submission.');
    } else {
      Logger.log('Successfully accessed tab: ' + TAB_NAME);
      Logger.log('Current number of rows: ' + dtfSheet.getLastRow());
    }
    
    return true;
  } catch (error) {
    Logger.log('Error accessing sheet: ' + error.toString());
    return false;
  }
}

/**
 * Debug function to log the structure of the APP sheet
 */
function debugAppSheetStructure() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(APP_TAB_NAME);
    
    if (!sheet) {
      Logger.log('APP sheet not found!');
      return;
    }
    
    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    
    Logger.log('APP Sheet Structure:');
    Logger.log('Total columns: ' + lastColumn);
    Logger.log('Headers:');
    for (let i = 0; i < headers.length; i++) {
      Logger.log('Column ' + (i + 1) + ': ' + headers[i]);
    }
    
    if (sheet.getLastRow() >= 2) {
      Logger.log('\nSample data from row 2:');
      const sampleData = sheet.getRange(2, 1, 1, lastColumn).getValues()[0];
      for (let i = 0; i < sampleData.length; i++) {
        Logger.log('Column ' + (i + 1) + ' (' + headers[i] + '): ' + sampleData[i]);
      }
    }
    
  } catch (error) {
    Logger.log('Error debugging sheet structure: ' + error.toString());
  }
}