function duplicateSheetWithProtection() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    
    // Создание копии листа
    var newSheet = sheet.copyTo(ss).setName(sheet.getName() + ' - Copy');
    
    // Перемещение нового листа на вторую позицию
    ss.setActiveSheet(newSheet);  // Делаем новый лист активным
    ss.moveActiveSheet(2);  // Перемещаем его на вторую позицию (индекс 2)

    // Копирование защиты листа (если есть)
    var sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (sheetProtections.length > 0) {
        var sheetProtection = sheetProtection[0]; // Защита листа (ожидается одна)

        // Создание новой защиты на новом листе
        var newSheetProtection = newSheet.protect();
        newSheetProtection.setDescription(sheetProtection.getDescription());
        newSheetProtection.setWarningOnly(sheetProtection.isWarningOnly());

        // Копирование исключений (диапазонов с исключениями)
        var unprotectedRanges = sheetProtection.getUnprotectedRanges();
        var newUnprotectedRanges = [];
        for (var j = 0; j < unprotectedRanges.length; j++) {
            var range = unprotectedRanges[j];
            var newRange = newSheet.getRange(range.getA1Notation());
            newUnprotectedRanges.push(newRange);
        }
        newSheetProtection.setUnprotectedRanges(newUnprotectedRanges);

        // Копирование редакторов
        if (!sheetProtection.isWarningOnly()) {
            newSheetProtection.removeEditors(newSheetProtection.getEditors());
            newSheetProtection.addEditors(sheetProtection.getEditors());

            if (sheetProtection.canDomainEdit()) {
                newSheetProtection.setDomainEdit(true);
            }
        }
    }
    
    // Устанавливаем оригинальный лист активным, чтобы вернуть пользователя на исходную позицию
    ss.setActiveSheet(sheet);
}
