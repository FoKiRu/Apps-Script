function duplicateSheetWithProtection() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    
    // Создание копии листа
    var newSheet = sheet.copyTo(ss).setName(sheet.getName() + ' - Copy');
    
    // Перемещение нового листа на вторую позицию
    ss.setActiveSheet(newSheet);  // Делаем новый лист активным
    ss.moveActiveSheet(2);  // Перемещаем его на вторую позицию (индекс 2)

    // Копирование защищённых диапазонов
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < protections.length; i++) {
        var protection = protections[i].copy();
        protection.setRange(newSheet.getRange(protection.getRange().getA1Notation()));
    }
    
    // Устанавливаем оригинальный лист активным, чтобы вернуть пользователя на исходную позицию
    ss.setActiveSheet(sheet);
}
