/**
 * Name: Google Ads Shopping Zombie Products Collector
 * Version: 1.0.0
 * Description: Знаходить Shopping-товари зі статусом ELIGIBLE, які не мали показів за останні N днів,
 *              та вивантажує їхні ID у Google Таблицю для присвоєння Custom Label (zombie).
 */

function main() {
  const SPREADSHEET_URL = 'ВАША_ПОСИЛАННЯ_НА_ТАБЛИЦЮ';
  const SHEET_NAME = 'zombie';

  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'custom_label_2']);
  }

  // Очищення таблиці (крім заголовків)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow(), 2).clearContent();
  }

  // Крок 1: Отримуємо ID товарів з показами > 0
  const activeIds = new Set();
  const performanceQuery =
      "SELECT segments.product_item_id " +
      "FROM shopping_performance_view " +
      "WHERE metrics.impressions > 0 " +
      "AND segments.date ≥ '" + getDateNDaysAgo(30) + "' " +
      "AND segments.date ≤ '" + getDateNDaysAgo(1) + "'";

  const perfResults = AdsApp.search(performanceQuery);
  while (perfResults.hasNext()) {
    const row = perfResults.next();
    // Безпечна перевірка наявності segments та productItemId
    if (row.segments && row.segments.productItemId) {
      activeIds.add(row.segments.productItemId);
    }
  }

  // Крок 2: Отримуємо всі товари зі статусом ELIGIBLE
  const allProductsQuery =
      "SELECT shopping_product.item_id " +
      "FROM shopping_product " +
      "WHERE shopping_product.status = 'ELIGIBLE'";

  const allResults = AdsApp.search(allProductsQuery);
  const zombieRows = [];

  while (allResults.hasNext()) {
    const row = allResults.next();
    if (row.shoppingProduct && row.shoppingProduct.itemId) {
      const productId = row.shoppingProduct.itemId;

      if (!activeIds.has(productId)) {
        zombieRows.push([productId, 'zombie']);
      }
    }
  }

  // Крок 3: Запис результатів
  if (zombieRows.length > 0) {
    sheet.getRange(2, 1, zombieRows.length, 2).setValues(zombieRows);
    console.log("Успішно знайдено 'зомбі': " + zombieRows.length);
  } else {
    console.log("Товарів без показів не знайдено.");
  }
}

function getDateNDaysAgo(n) {
  const date = new Date();
  date.setDate(date.getDate() - n);
  const year = date.getFullYear();
  const month = ("0" + (date.getMonth() + 1)).slice(-2);
  const day = ("0" + date.getDate()).slice(-2);
  return year + "-" + month + "-" + day;
}
