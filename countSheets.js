// Подсчет листов в таблице по токину

function countSheets() {
    var spreadsheetId = '1jPEzFRdpyVEtzQPjqQMoj4Nt7bd3q1xwM9tdVtZLY38';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheets = spreadsheet.getSheets();
    var sheetCount = sheets.length;
    
    Logger.log("Количество листов в таблице: " + sheetCount);
    return sheetCount;
  }
  