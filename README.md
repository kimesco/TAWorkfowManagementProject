# TAWorkfowManagementProject
## AppScript Duplication Checker for Google Sheets

*Disclaimer: After testing and reviewing the code, it did not indicate any form of security risk. However, implementation of this across the INFO College should abide by any cybersecurity policy implemented by the college and university*
 
## How the Code works
The code does the following:
- Highlights duplicate values from one or more columns.
- Deletes duplicate values from one or more columns.
 
onOpen() function:
This function is a Google function that allows the user to create a “button” in the Menu option of Google Sheets.
 
highlightDuplicates() function:
This function will grab the active sheet and go to the chosen column the user inputs in the text box. It will check if the column exists. The confirmation of the column will run a for loop to traverse the current active sheet cells of the chosen column. As it traverses through the column, it will highlight the second encounter of the value they found and highlight that duplicate cell with a yellow background.
 
removeDuplicates() function:
This function will grab the current sheet and go to the chosen column from the user input. After confirming if the column exists, the duplicate, second encounter or more of the same cell value, will be deleted. As those duplicates are deleted, the row it belongs to also gets deleted. Then it condenses the whitespaces, empty rows, of the current active sheet to be formatted.
 
highlightMultiDuplicates() function:
Like the highlightDuplicates function, it does the same function however, it involves more than one column and will highlight the duplicates of multiple columns chosen. Then highlight the background of those duplicates to yellow.
removeMultiDuplicates() function:
The removeMultiDuplicates function does the same as removeDuplicates function but with multiple columns. Chosen. Goes through the columns cells to find the duplicates and deletes those duplicates from the multiple chosen columns.
getColumnInput() function:
This prompts the pop-up box for grabbing the user input. It records the user’s response to determine what the user has chosen. If the user does not press “okay” it does nothing.


Azhar, Aris. Effortlessly Remove Duplicates in Google Sheets Using Apps Script - Aris Azhar, Aris Azhar - Trusted Digital Concierge, 5 Feb. 2024, arisazhar.com/effortlessly-remove-duplicates-in-google-sheets-using-apps-script/
