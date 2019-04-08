var Excel = require('exceljs');

var workbook = new Excel.Workbook();

var worksheet = workbook.addWorksheet('My Sheet');

var dobCol1 = worksheet.getColumn(1);
var dobCol2 = worksheet.getColumn(3);
var dobCol3 = worksheet.getColumn(6)

dobCol1.width = 30;
dobCol2.width = 25;
dobCol3.width = 25;

worksheet.getColumn(1).alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getColumn(3).alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getColumn(6).alignment = { vertical: 'middle', horizontal: 'center' };


worksheet.mergeCells('A1:A9');
worksheet.mergeCells('C1:G1');

worksheet.getCell('A1').value  = 'HEADER LEVEL INFORMATION';
worksheet.getCell('C1').value  = 'Supply of Gaskets';
worksheet.getCell('C2').value = 'Technical Bid Due Date';
worksheet.getCell('C4').value = 'Offer Refernce No.';
worksheet.getCell('C5').value = 'Currency';
worksheet.getCell('C6').value = 'Offer Validity';
worksheet.getCell('C7').value = 'Delivery Period';
worksheet.getCell('C8').value = 'Payment Terms';

worksheet.getCell('F2').value = 'Commercial Bid Due Date';
worksheet.getCell('F4').value = 'Inco Term 1';
worksheet.getCell('F5').value = 'Inco Term 2 - Location';
worksheet.getCell('F6').value = 'HSN Code';
worksheet.getCell('F7').value = 'GS code';
worksheet.getCell('F8').value = 'Certification Charges';

workbook.xlsx.writeFile('testfile.xlsx').then(function() {
    // done
    console.log('file is written');
});