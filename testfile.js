// var Excel = require('exceljs');

// var workbook = new Excel.Workbook();

// var worksheet = workbook.addWorksheet('My Sheet');

// var dobCol1 = worksheet.getColumn(1);
// var dobCol2 = worksheet.getColumn(3);
// var dobCol3 = worksheet.getColumn(6)

// dobCol1.width = 30;
// dobCol2.width = 25;
// dobCol3.width = 25;

// worksheet.getColumn(1).alignment = { vertical: 'middle', horizontal: 'center' };
// worksheet.getColumn(3).alignment = { vertical: 'middle', horizontal: 'center' };
// worksheet.getColumn(6).alignment = { vertical: 'middle', horizontal: 'center' };


// worksheet.mergeCells('A1:A9');
// worksheet.mergeCells('C1:G1');

// worksheet.getCell('A1').value  = 'HEADER LEVEL INFORMATION';
// worksheet.getCell('C1').value  = 'Supply of Gaskets';
// worksheet.getCell('C2').value = 'Technical Bid Due Date';
// worksheet.getCell('C4').value = 'Offer Refernce No.';
// worksheet.getCell('C5').value = 'Currency';
// worksheet.getCell('C6').value = 'Offer Validity';
// worksheet.getCell('C7').value = 'Delivery Period';
// worksheet.getCell('C8').value = 'Payment Terms';

// worksheet.getCell('F2').value = 'Commercial Bid Due Date';
// worksheet.getCell('F4').value = 'Inco Term 1';
// worksheet.getCell('F5').value = 'Inco Term 2 - Location';
// worksheet.getCell('F6').value = 'HSN Code';
// worksheet.getCell('F7').value = 'GS code';
// worksheet.getCell('F8').value = 'Certification Charges';

// workbook.xlsx.writeFile('testfile.xlsx').then(function() {
//     // done
//     console.log('file is written');
// });


///////////////////////////////////////////////////////////////
var Excel = require('exceljs');


module.exports.generateExcel = async (reqData) => {
    let rfqNo;
  //  let rfqSno = reqData.length;
    // console.log(reqData.length,'length')
    // for(let i = 1; i< reqData.length; i++){ 
    //     console.log(reqData[i],'-------')
    // }
     
       for(let result of reqData){
         rfqNo = reqData.rfqNo;
        let 
       }

    var workbook = new Excel.Workbook();

    var worksheet = workbook.addWorksheet('My Sheet');

    var dobCol1 = worksheet.getColumn(1);
    var dobCol2 = worksheet.getColumn(3);
    var dobCol3 = worksheet.getColumn(6)
    var dobCol4 = worksheet.getColumn(2);

    dobCol1.width = 15;
    dobCol2.width = 25;
    dobCol3.width = 35;
    dobCol4.width = 35;


    worksheet.mergeCells('A1:A9');
    worksheet.mergeCells('C1:G1');
    worksheet.mergeCells('A11:A13');
    worksheet.mergeCells('A15:A19');
    worksheet.mergeCells('B11:C13');
    worksheet.mergeCells('D11:D13');

    var cell = worksheet.getCell('B11');
    var html = '<div>' + cell.html + '</div>';

    worksheet.getCell('A1').value = 'HEADER LEVEL INFORMATION';
    worksheet.getCell('B1').value = 'RFX No.: RFQ '+ rfqNo;
    worksheet.getCell('B2').value = 'Rev 1';
    worksheet.getCell('A11').value = 'Bill To Ship To level Information';
    worksheet.getCell('B11').value = "Bill To:  RELIANCE INDUSTRIES LIMITED Village : Meghpar / Padana District : Jamnagar 361280 Gujarat"
    worksheet.getCell('C1').value = 'Supply of Gaskets';
    worksheet.getCell('C2').value = 'Technical Bid Due Date';
    worksheet.getCell('C4').value = 'Offer Refernce No.';
    worksheet.getCell('C5').value = 'Currency';
    worksheet.getCell('C6').value = 'Offer Validity';
    worksheet.getCell('C7').value = 'Delivery Period';
    worksheet.getCell('C8').value = 'Payment Terms';

    worksheet.getCell('D11').value = 'Ship To: DTA Jamnagar';

    worksheet.getCell('F2').value = 'Commercial Bid Due Date';
    worksheet.getCell('F4').value = 'Inco Term 1';
    worksheet.getCell('F5').value = 'Inco Term 2 - Location';
    worksheet.getCell('F6').value = 'HSN Code';
    worksheet.getCell('F7').value = 'GS code';
    worksheet.getCell('F8').value = 'Certification Charges';
    worksheet.getCell('F12').value = 'Packaging & Forwarding Charges';
    worksheet.getCell('F13').value = 'Freight Charges';
    worksheet.getCell('H12').value = 'Currency';
    worksheet.getCell('H13').value = 'Inco Terms';

    worksheet.getCell('I4').value = 'Inspection Charges';
    worksheet.getCell('I5').value = 'Miscellaneous Charges';
    worksheet.getCell('I6').value = 'CGST';
    worksheet.getCell('I7').value = 'SGST';
    worksheet.getCell('I8').value = 'IGST';


    worksheet.getRow(15).values = ['Item Level Information', 'S.No.', 'Item Code', 'Item Description', 'UOM', 'RIL Delivery Date', 'RIL Qty', 'Qty', 'Delivery Period', 'Unit Rate', 'HSN Code', 'GS Code', 'Offer Validity', 'Remarks', 'Inspection Charges', 'Certificate charges', 'Miscellaneous Charges', 'CGST', 'SGST', 'IGST']
    worksheet.eachRow(function (row, rowID) {
        worksheet.getRow(rowID).font = { bold: true }
        worksheet.getRow(rowID).alignment = { vertical: 'middle', horizontal: 'center' };
    });
    worksheet.getColumn(1).alignment = { vertical: 'middle', horizontal: 'center', textRotation: 90 };
    await workbook.xlsx.writeFile('testfile.xlsx');
    return true;
}



