function generateInvoiceCsv() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('當前沒有選定工作表。');
    return;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  let serial_number_increment = 0;
  
  // 表頭記錄
  let csvContent = 'H,INVO,C11231776956,332637694,' + sheetName + ',,,,,,,,,,,\n';
  let hasError = false;
  let errorMessage = '';
  
  try {
    data.slice(1).forEach(row => {
      const serial_number = sheetName + ('00' + (serial_number_increment + 1)).slice(-3);
      serial_number_increment++;
      
      const sale_type = row[0];
      let sale_row;
      if (sale_type === 'B2B') {
        sale_row = ['S', serial_number, sale_type, row[1], row[2], row[4], '', '', '', '', 'Y', '1', '5'];
      } else if (sale_type === 'B2C') {
        sale_row = ['S', serial_number, sale_type, '', row[3], row[4], '', '', '', '', 'Y', '1', '5'];
      } else {
        throw new Error("Unknown sale type: " + sale_type);
      }
      
      const product_price = row[5];
      const sale_amount = Math.round(product_price / 1.05);
      const tax_amount = product_price - sale_amount;
      sale_row = sale_row.concat([sale_amount, tax_amount, product_price]);
      csvContent += sale_row.join(',') + '\n';
      
      const product_name = row[6];
      const quantity = row[7];
      const unit = row[8];
      const item_total = quantity * product_price;
      const item_row = ['I', serial_number, product_name, quantity, unit, product_price, item_total];
      csvContent += item_row.join(',') + '\n';
    });
  } catch (error) {
    hasError = true;
    errorMessage = error.message;
  }
  
  const fileName = '332637694_' + sheetName + '.csv';
  const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
  
  // 指定文件夾 ID
  const folderId = '1hwnEhYitCMIdaWUkzr6uFrj89eDCm2ev';
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);
  const fileUrl = file.getUrl();
  
  const lastColumn = sheet.getLastColumn();
  sheet.getRange(1, lastColumn + 1).setValue('執行時間');
  sheet.getRange(2, lastColumn + 1, data.length - 1, 1).setValue(timestamp);
  sheet.getRange(1, lastColumn + 2).setValue('檔案位置');
  sheet.getRange(2, lastColumn + 2).setValue(fileUrl);
  
  if (hasError) {
    sheet.getRange(1, lastColumn + 3).setValue('錯誤說明');
    sheet.getRange(2, lastColumn + 3, data.length - 1, 1).setValue(errorMessage);
  }
  
  SpreadsheetApp.getUi().alert('CSV file created: ' + fileUrl);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Generate Invoice CSV', 'generateInvoiceCsv')
    .addToUi();
}
