
function main(workbook: ExcelScript.Workbook) {
    // Set up the list to run and filter through
    let firstWorkSheet = workbook.getActiveWorksheet();
    let i = 1;
    let FirstLibrarianList: string[] = [];
    let SecondLibrarinList: string[] = [];
    let ThirdLibrarianList: string[] = [];
    let UnableToPlaceBookList: string[] = [];
    // Loop through values of instances.created_at because it is a gauranteed field that will always be filled
    while (firstWorkSheet.getCell(i, 38).getValue() != null && firstWorkSheet.getCell(i, 38).getValue() !== '') {
        // Logic to determine which liaison area it should go to
        // First see if the library is WAWL
        if (firstWorkSheet.getCell(i, 2).getValue() == 'Library') {
            //Prefilter location to know liason first
            let itemLocation = firstWorkSheet.getCell(i, 3).getValue().toString();
            if (itemLocation === '' || itemLocation === '' || itemLocation === '') {
                ThirdLibrarianList.push('Call Number: ' + firstWorkSheet.getCell(i, 9).getValue().toString() + ' Title: ' + firstWorkSheet.getCell(i, 56).getValue().toString() + ' Barcode: ' + firstWorkSheet.getCell(i, 70).getValue().toString());
            } else if (itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '') {
                FirstLibrarianList.push('Call Number: ' + firstWorkSheet.getCell(i, 9).getValue().toString() + ' Title: ' + firstWorkSheet.getCell(i, 56).getValue().toString() + ' Barcode: ' + firstWorkSheet.getCell(i, 70).getValue().toString());
            } else if (itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '' || itemLocation === '') {
                SecondLibrarinList.push('Call Number: ' + firstWorkSheet.getCell(i, 9).getValue().toString() + ' Title: ' + firstWorkSheet.getCell(i, 56).getValue().toString() + ' Barcode: ' + firstWorkSheet.getCell(i, 70).getValue().toString());
            } else {
                // This should be everything except refrence and oversized
                
                let callNumber = firstWorkSheet.getCell(i, 9).getValue().toString();
                // Sanitize the string first
                if (callNumber.indexOf('Q') === 0) {
                    callNumber = callNumber.slice(2);
                }
                if (callNumber.indexOf('REF') === 0) {
                    callNumber = callNumber.slice(4);
                }
                if (callNumber.indexOf('R') === 0) {
                    callNumber = callNumber.slice(2);
                }
                // Split the call number in to number and cutter
                let callNumberValue = Number(callNumber.split(' ')[0]);
                // Add based on DDC location
                if (callNumberValue < 10 || (callNumberValue >= 330 && callNumberValue < 340) || (callNumberValue >= 500 && callNumberValue < 610) || (callNumberValue >= 620 && callNumberValue < 700)) {
                    SecondLibrarinList.push('Call Number: ' + firstWorkSheet.getCell(i, 9).getValue().toString() + ' Title: ' + firstWorkSheet.getCell(i, 56).getValue().toString() + ' Barcode: ' + firstWorkSheet.getCell(i, 70).getValue().toString());
                } else if ((callNumberValue >= 370 && callNumberValue < 380) || (callNumberValue >= 610 && callNumberValue < 620)) {
                    ThirdLibrarianList.push('Call Number: ' + firstWorkSheet.getCell(i, 9).getValue().toString() + ' Title: ' + firstWorkSheet.getCell(i, 56).getValue().toString() + ' Barcode: ' + firstWorkSheet.getCell(i, 70).getValue().toString());
                } else if (callNumberValue >= 700  || (callNumberValue >= 10 && callNumberValue < 330) || (callNumberValue >= 340 && callNumberValue < 360) || (callNumberValue >= 380 && callNumberValue < 500)) {
                    FirstLibrarianList.push('Call Number: ' + firstWorkSheet.getCell(i, 9).getValue().toString() + ' Title: ' + firstWorkSheet.getCell(i, 56).getValue().toString() + ' Barcode: ' + firstWorkSheet.getCell(i, 70).getValue().toString());
                } else {
                    UnableToPlaceBookList.push('Call Number: ' + firstWorkSheet.getCell(i, 9).getValue().toString() + ' Title: ' + firstWorkSheet.getCell(i, 56).getValue().toString() + ' Barcode: ' + firstWorkSheet.getCell(i, 70).getValue().toString());
                }
            }
            
        }
        i++;
    }
    // Create a new sheet for the emmail templates
    workbook.addWorksheet('EmailBody');
    let emailWorksheet = workbook.getWorksheet("EmailBody");
    let bookListString = '';
    ThirdLibrarianList.forEach(bookRow => bookListString = bookListString + bookRow + '\n');
    emailWorksheet.getCell(0, 0).setValue('ThirdLibrarian');
    emailWorksheet.getCell(0, 1).setValue('Hi ThirdLibrarian, \n Here are the books that were listed as lost or missing in the past month. Let me know if you have any questions. \n' + bookListString);
    // Reset string
    bookListString = '';
    FirstLibrarianList.forEach(bookRow => bookListString = bookListString + bookRow + '\n');
    emailWorksheet.getCell(1, 0).setValue('FirstLibrarian');
    emailWorksheet.getCell(1, 1).setValue('Hi FirstLibrarian, \n Here are the books that were listed as lost or missing in the past month. Let me know if you have any questions. \n' + bookListString);
    // Reset string
    bookListString = '';
    SecondLibrarinList.forEach(bookRow => bookListString = bookListString + bookRow + '\n');
    emailWorksheet.getCell(2, 0).setValue('SecondLibrarin Email');
    emailWorksheet.getCell(2, 1).setValue('Hi SecondLibrarin, \n Here are the books that were listed as lost or missing in the past month. Let me know if you have any questions. \n' + bookListString);
    // Reset string
    bookListString = '';
    UnableToPlaceBookList.forEach(bookRow => bookListString = bookListString + bookRow + '\n');
    emailWorksheet.getCell(3, 0).setValue('Unable to Sort');
    emailWorksheet.getCell(3, 1).setValue(bookListString);
}
