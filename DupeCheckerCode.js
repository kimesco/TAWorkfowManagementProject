/*DISCLAIMER: AFTER TESTING AND REVIEWING THE CODE IT DOES NOT INDICATE ANY FORM
 OF SECURITY RISK. HOWEVER, ANY IMPLEMENTATION OF THIS ACROSS THE COLLEGE SHOULD
  ABIDE BY ANY CYBERSECURITY POLICY IMPLEMENTED BY THE COLLEGE AND UNIVERSITY.
 */



/*
 A function where it provides the Dupe Check "button" on Google Sheets
with the viablew options that it can contain
Button Name: Dupe Checker
Once clicked you have the options avaialable:
1. Highlight Duplicate
2. Remove Single Colum Duplicates
3. Highlight Multi-Column Duplicates
4. Remove Multi-Column Duplicates
*/
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Dupe Checker')
      .addItem('Highlight Duplicate', 'highlightDuplicates')
      .addItem('Remove Single Column Duplicates', 'removeDuplicates')
      .addItem('Highlight Multi-Column Duplicates', 'highlightMultiColumnDuplicates')
      .addItem('Remove Multi-Column Duplicates', 'removeMultiColumnDuplicates')
      .addToUi();
  }
  
  /*
  Function will highlight the duplicates in the selected column. 
  The highlight function will be changing the background of the cells 
  to yellow when the second occurance or more happens.
   */
  function highlightDuplicates() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var columnLetter = getColumnInput('Highlight Duplicates');
    if (!columnLetter) return;
    
    var columnNumber = columnLetter.charCodeAt(0) - 64;
    var data = sheet.getDataRange().getValues();
    var seenValues = {};
    var duplicateDetails = {};
    
    for (var i = 1; i < data.length; i++) {
      var cellValue = data[i][columnNumber - 1];
      if (seenValues[cellValue]) {
        sheet.getRange(i + 1, columnNumber).setBackground('yellow');
        duplicateDetails[cellValue] = (duplicateDetails[cellValue] || 1) + 1;
      } else {
        seenValues[cellValue] = true;
      }
    }
    
    var message = Object.keys(duplicateDetails).map(function(key) {
      return 'Value "' + key + '" has ' + (duplicateDetails[key] - 1) + ' additional occurrences.';
    }).join('\n');
  
    SpreadsheetApp.getUi().alert(message || 'No duplicates found.');
  }
  
  /*
  A function that removes the duplicates from a single column selected. Then 
  condensed the sheet so there won't be white spaces in between.
   */
  function removeDuplicates() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var columnLetter = getColumnInput('Remove Duplicates');
    if (!columnLetter) return;
  
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).setBackground(null);
    
    var columnNumber = columnLetter.charCodeAt(0) - 64;
    var data = sheet.getDataRange().getValues();
    var uniqueRows = [data[0]];
    var seenValues = {};
  
    for (var i = 1; i < data.length; i++) {
      var cellValue = data[i][columnNumber - 1];
      if (!seenValues[cellValue]) {
        uniqueRows.push(data[i]);
        seenValues[cellValue] = true;
      }
    }
  
    sheet.clearContents();
    sheet.getRange(1, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
  }
  
  /*
  Function will highlight the duplicates in the multiple columns selected. 
  The highlight function will be changing the background of the cells 
  to yellow when the second occurance or more happens.
  */
  function highlightMultiColumnDuplicates() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var columnLetters = getColumnInput('Highlight Multi-Column Duplicates', true);
    if (!columnLetters) return;
    
    var columnNumbers = columnLetters.split(',').map(function(letter) {
      return letter.trim().toUpperCase().charCodeAt(0) - 64;
    });
    
    var data = sheet.getDataRange().getValues();
    var seenCombinations = {};
    var duplicateDetails = {};
  
    for (var i = 1; i < data.length; i++) {
      var combination = columnNumbers.map(function(number) {
        return data[i][number - 1];
      }).join('|');
      
      if (seenCombinations[combination]) {
        columnNumbers.forEach(function(number) {
          sheet.getRange(i + 1, number).setBackground('yellow');
        });
        duplicateDetails[combination] = (duplicateDetails[combination] || 1) + 1;
      } else {
        seenCombinations[combination] = true;
      }
    }
    
    var message = Object.keys(duplicateDetails).map(function(key) {
      return 'Combination "' + key.replace(/\|/g, ', ') + '" has ' + (duplicateDetails[key] - 1) + ' additional occurrences.';
    }).join('\n');
  
    SpreadsheetApp.getUi().alert(message || 'No duplicates found.');
  }
  
  /*
  A function that removes the duplicates from the multiple columns selected. Then 
  condensed the sheet so there won't be white spaces in between.
  */
  function removeMultiColumnDuplicates() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var columnLetters = getColumnInput('Remove Multi-Column Duplicates', true);
    if (!columnLetters) return;
  
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).setBackground(null);
    
    var columnNumbers = columnLetters.split(',').map(function(letter) {
      return letter.trim().toUpperCase().charCodeAt(0) - 64;
    });
    
    var data = sheet.getDataRange().getValues();
    var uniqueRows = [data[0]];
    var seenCombinations = {};
  
    for (var i = 1; i < data.length; i++) {
      var combination = columnNumbers.map(function(number) {
        return data[i][number - 1];
      }).join('|');
  
      if (!seenCombinations[combination]) {
        uniqueRows.push(data[i]);
        seenCombinations[combination] = true;
      }
    }
  
    sheet.clearContents();
    sheet.getRange(1, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
  } 
  
  /*
  Helper method that tells the main function if the entered input from the user
  was one or more column
   */
  function getColumnInput(actionName, isMultiColumn = false) {
    var message = isMultiColumn ? 'Enter the column letters separated by comma (e.g., A,B):' : 'Enter the column letter:';
    var response = SpreadsheetApp.getUi().prompt(actionName, message, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
      return null;
    }
    return response.getResponseText();
  }