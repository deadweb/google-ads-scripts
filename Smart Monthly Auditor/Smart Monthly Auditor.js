/**
 * Name: Smart Monthly Auditor
 * Version: 1.0.0
 * Description: Виводить кількість конверсій і витрати (без ПДВ) для вказаних акаунтів в таблицю Google Sheets.
 */

function main() {
  // 1. НАЛАШТУВАННЯ
  var SPREADSHEET_URL = 'ВАША_ТАБЛИЦЯ';
  var LABEL_NAME = 'Smart_Monthly_Auditor'; // Мітка для вибору акаунтів
  
  // ОПЦІЯ ПДВ: true — рахувати з ПДВ (х1.2), false — без ПДВ
  var INCLUDE_VAT = true; 
  var VAT_RATE = 1.2;

  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var accountIterator = AdsManagerApp.accounts()
      .withCondition('LabelNames CONTAINS "' + LABEL_NAME + '"')
      .get();

  while (accountIterator.hasNext()) {
    var account = accountIterator.next();
    var accountName = account.getName();
    var accountId = account.getCustomerId(); // Скрипт бере ID про всяк випадок
    
    // АВТОМАТИЧНЕ ПРИСВОЄННЯ НАЗВИ: 
    // Якщо у акаунта немає імені, скрипт назве аркуш його цифровим ID
    var sheetName = accountName ? accountName : accountId;
    
    // Очищення від символів, які заборонені в назвах аркушів Google Sheets (\ / ? * : [ ])
    sheetName = sheetName.replace(/[\*＼？：\/\\\[\]]/g, '').trim();
    
    // Обрізання довжини до 31 символу (максимальний ліміт Google Sheets для назви аркуша)
    if (sheetName.length > 31) {
      sheetName = sheetName.substring(0, 31);
    }
    
    AdsManagerApp.select(account);
    
    // Скрипт шукає або автоматично створює аркуш із згенерованим ім'ям
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    processAccountSheet(sheet, INCLUDE_VAT, VAT_RATE);
  }
}

function processAccountSheet(sheet, includeVat, vatRate) {
  sheet.clear();
  
  // Заголовки
  var costHeader = includeVat ? "Витрати (з ПДВ)" : "Витрати (без ПДВ)";
  sheet.getRange("A1").setValue("Місяць");
  sheet.getRange("B1").setValue("Конверсії");
  sheet.getRange("C1").setValue("Цінність конв.");
  sheet.getRange("D1").setValue(costHeader);
  sheet.getRange("A1:D1").setFontWeight("bold");

  // Запит даних
  var report = AdsApp.report(
    'SELECT segments.month, metrics.conversions, metrics.conversions_value, metrics.cost_micros ' +
    'FROM customer ' +
    'WHERE segments.date ≥ "2000-01-01" AND segments.date ≤ "3000-12-31" ' +
    'ORDER BY segments.month ASC'
  );

  var rows = report.rows();
  var output = [];

  while (rows.hasNext()) {
    var row = rows.next();
    var parts = row['segments.month'].split('-');
    var displayMonth = parts[1] + "." + parts[0]; 
    
    var conversions = parseFloat(row['metrics.conversions']);
    var convValue = parseFloat(row['metrics.conversions_value']);
    var rawCost = parseFloat(row['metrics.cost_micros']) / 1000000;
    
    // Розрахунок витрат з ПДВ або без
    var finalCost = includeVat ? (rawCost * vatRate) : rawCost;

    output.push([displayMonth, conversions, convValue, finalCost]);
  }

  // Запис та форматування
  if (output.length > 0) {
    var range = sheet.getRange(2, 1, output.length, 4);
    range.setValues(output);
    
    // Формат з комою для стовпців B, C, D
    sheet.getRange(2, 2, output.length, 3).setNumberFormat("#,##0.00");
  }
  
  sheet.getRange("A:D").setHorizontalAlignment("left");
}
