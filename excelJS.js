var dataToConvert = [
    { ID: 10011, NAME: "A", DEPARTMENT: "Sales", MONTH: "Jan", YEAR: 2020, SALES: 132412, CHANGE: 12, LEADS: 35 },
    { ID: 10012, NAME: "A", DEPARTMENT: "Sales", MONTH: "Feb", YEAR: 2020, SALES: 232324, CHANGE: 2, LEADS: 443 },
    { ID: 10013, NAME: "A", DEPARTMENT: "Sales", MONTH: "Mar", YEAR: 2020, SALES: 542234, CHANGE: 45, LEADS: 345 },
    { ID: 10014, NAME: "A", DEPARTMENT: "Sales", MONTH: "Apr", YEAR: 2020, SALES: 223335, CHANGE: 32, LEADS: 234 },
    { ID: 10015, NAME: "A", DEPARTMENT: "Sales", MONTH: "May", YEAR: 2020, SALES: 455535, CHANGE: 21, LEADS: 12 },
  ];

function generateExcel(){
		var filename = 'test-excel.xlsx';
		var worksheetName = 'sheet1';
		var workbook = new ExcelJS.Workbook();
		
		workbook.creator = 'Me';
		workbook.lastModifiedBy = 'Her';
		workbook.created = new Date(1985, 8, 30);
		workbook.modified = new Date();
		workbook.lastPrinted = new Date(2016, 9, 27);
				
		var worksheet = workbook.addWorksheet(worksheetName, {
										properties:{tabColor:{argb:'FFC0000'}},
										pageSetup:{paperSize: 9, orientation:'landscape'}
									});
		// Header
		worksheet.addRow(['ID', 'NAME','DEPARTMENT']);
    
        // Add rows to excel
        dataToConvert.forEach(element => 		worksheet.addRow([element.ID, element.NAME, element.DEPARTMENT]));
    
        worksheet.getColumn(3).font = {
            name: 'Comic Sans MS',
            color: { argb: 'FFFF0000' },
            family: 4,
            size: 16,
            underline: true,
            bold: true
        };
    
		var mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
		workbook.xlsx.writeBuffer()
			.then(function(data) {
				console.log('Binary Buffer Opened');
				console.log(data);

				console.log('Creating blob');
				var blob = new Blob([data], { type : mimeType });

				console.log('Writing file to output/test.xlsx');

				saveAs(blob, filename);
				console.log('File written!');
			});
	
}
