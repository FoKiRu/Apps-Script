// Подсчет защищенных листов
function countProtectedSheets() {
    var spreadsheerId = "1jPEzFRdpyVEtzQPjqQMoj4Nt7bd3q1xwM9tdVtZLY38";
    var spreasheet = SpreadsheetApp.openById(spreadsheerId);
    var sheets = spreadsheet.getSheets();
    var protectedSheetCount = 0;

    // Проходим по каждому листу и проверяем, есть ли защита
    sheets.forEach(function(shhet) {
        var protection = sheet.getProtection(SpreadsheetApp.ProtectionType.SHEET);

        // Если хотя бы одна защита найдена, счтаем лист защищенным
        if (protections.length > 0) {
            protectedSheetCount++;
        }
    });

    Logger.log("Количество листов в таблице: " + sheets.length);
    Logger.log("Количество защищенных листов в таблице: " + protectedSheetCount);
    
    return protectedSheetCount;
}