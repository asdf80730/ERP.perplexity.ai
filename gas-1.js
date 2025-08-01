/**
 * 簡易庫存管理系統 - Google Apps Script 後端 (完整修正版)
 * 解決 Timestamp 序列化問題，包含完整功能和測試函數
 */

// ========== 主要入口函數 ==========

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('簡易庫存管理系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ========== 測試和除錯函數 ==========

function debugConnection() {
  try {
    console.log('=== 開始除錯 ===');
    
    // 1. 測試基本連接
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('試算表名稱:', ss.getName());
    console.log('試算表 ID:', ss.getId());
    
    // 2. 檢查工作表
    const sheets = ss.getSheets();
    console.log('現有工作表:', sheets.map(s => s.getName()));
    
    // 3. 測試 getAllInitialData 函數
    const result = getAllInitialData();
    console.log('getAllInitialData 執行結果:', result);
    
    return '除錯完成';
    
  } catch (error) {
    console.error('除錯過程發生錯誤:', error);
    console.error('錯誤詳細訊息:', error.message);
    console.error('錯誤堆疊:', error.stack);
    throw error;
  }
}

function simpleTest() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('試算表名稱:', ss.getName());
    return {
      success: true,
      name: ss.getName(),
      id: ss.getId()
    };
  } catch (error) {
    console.error('簡單測試失敗:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function testConnection() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('Google Sheet 連接正常');
    console.log('試算表名稱:', ss.getName());
    console.log('試算表ID:', ss.getId());
    return 'Google Apps Script 連接成功！';
  } catch (error) {
    console.error('連接測試失敗:', error);
    throw new Error('連接測試失敗: ' + error.message);
  }
}

function testDataFormat() {
  try {
    const result = getAllInitialData();
    
    console.log('=== 資料格式測試 ===');
    console.log('locations 類型:', typeof result.locations);
    console.log('locations 是否為陣列:', Array.isArray(result.locations));
    console.log('locations 長度:', result.locations.length);
    
    console.log('transactions 類型:', typeof result.transactions);
    console.log('transactions 是否為陣列:', Array.isArray(result.transactions));
    console.log('transactions 長度:', result.transactions.length);
    
    if (result.transactions.length > 0) {
      console.log('第一筆交易範例:', JSON.stringify(result.transactions[0]));
    }
    
    return result;
    
  } catch (error) {
    console.error('測試失敗:', error);
    throw error;
  }
}

function resetAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ['Locations', 'Products', 'Inventory', 'Transactions'];
  
  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.clear();
    }
  });
  
  console.log('所有數據已重置');
}

function insertTestData() {
  try {
    addLocations(['倉庫A', '倉庫B', '店面']);
    
    addProducts([
      { ProductName: '筆記型電腦', Brand: 'Dell', Category: '電子產品' },
      { ProductName: '無線滑鼠', Brand: 'Logitech', Category: '配件' },
      { ProductName: '機械鍵盤', Brand: 'Razer', Category: '配件' }
    ]);
    
    console.log('測試數據已插入');
  } catch (error) {
    console.error('插入測試數據失敗:', error);
    throw error;
  }
}

// ========== 系統初始化 ==========

function getAllInitialData() {
  try {
    console.log('=== getAllInitialData 開始執行 ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('成功取得試算表:', ss.getName());
    
    // 確保所有必要的工作表存在
    ensureSheetExists(ss, 'Locations', ['LocationID', 'LocationName']);
    ensureSheetExists(ss, 'Products', ['ProductID', 'ProductName', 'Brand', 'Category']);
    ensureSheetExists(ss, 'Inventory', ['InventoryID', 'ProductID', 'LocationID', 'Quantity']);
    ensureSheetExists(ss, 'Transactions', ['TransactionID', 'Timestamp', 'Type', 'ProductID', 'LocationID', 'Quantity', 'Notes']);
    
    // 讀取各工作表數據
    const locations = getSheetDataAsObjects(ss, 'Locations');
    const products = getSheetDataAsObjects(ss, 'Products');
    const inventory = getSheetDataAsObjects(ss, 'Inventory');
    const transactions = getSheetDataAsObjects(ss, 'Transactions');
    
    // 計算每個地點的商品總件數
    const locationCounts = calculateLocationCounts(inventory);
    const locationsWithCounts = locations.map(location => ({
      LocationID: String(location.LocationID),
      LocationName: String(location.LocationName),
      Count: locationCounts[location.LocationID] || 0
    }));
    
    // 確保所有資料都是可序列化的格式
    const cleanProducts = products.map(product => ({
      ProductID: String(product.ProductID),
      ProductName: String(product.ProductName),
      Brand: String(product.Brand),
      Category: String(product.Category)
    }));
    
    const cleanInventory = inventory.map(item => ({
      InventoryID: String(item.InventoryID),
      ProductID: String(item.ProductID),
      LocationID: String(item.LocationID),
      Quantity: parseInt(item.Quantity) || 0
    }));
    
    const cleanTransactions = transactions.map(transaction => ({
      TransactionID: String(transaction.TransactionID),
      Timestamp: String(transaction.Timestamp),
      Type: String(transaction.Type),
      ProductID: String(transaction.ProductID),
      LocationID: String(transaction.LocationID),
      Quantity: parseInt(transaction.Quantity) || 0,
      Notes: String(transaction.Notes || '')
    }));
    
    // 建構最終回傳物件
    const result = {
      locations: locationsWithCounts,
      products: cleanProducts,
      inventory: cleanInventory,
      transactions: cleanTransactions
    };
    
    console.log('準備回傳的資料:', {
      locationsCount: result.locations.length,
      productsCount: result.products.length,
      inventoryCount: result.inventory.length,
      transactionsCount: result.transactions.length
    });
    
    return result;
    
  } catch (error) {
    console.error('getAllInitialData 發生錯誤:', error);
    console.error('錯誤訊息:', error.message);
    throw new Error('讀取數據失敗: ' + error.message);
  }
}

function ensureSheetExists(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
    }
  }
  return sheet;
}

function getSheetDataAsObjects(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    console.log(`工作表 ${sheetName} 不存在，回傳空陣列`);
    return [];
  }
  
  const dataRange = sheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    console.log(`工作表 ${sheetName} 只有標題行或為空，回傳空陣列`);
    return [];
  }
  
  const data = dataRange.getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  const result = [];
  for (let i = 0; i < rows.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      const header = headers[j];
      const value = rows[i][j];
      
      if (header === 'Quantity' && value !== '') {
        obj[header] = parseInt(value) || 0;
      } else if (header === 'Timestamp' && value instanceof Date) {
        obj[header] = value.toLocaleString('zh-TW');
      } else if (header === 'Timestamp' && value) {
        obj[header] = new Date(value).toLocaleString('zh-TW');
      } else {
        obj[header] = value || '';
      }
    }
    result.push(obj);
  }
  
  console.log(`從 ${sheetName} 讀取到 ${result.length} 筆資料`);
  return result;
}

function calculateLocationCounts(inventory) {
  const counts = {};
  
  if (Array.isArray(inventory)) {
    inventory.forEach(item => {
      if (item.LocationID) {
        if (!counts[item.LocationID]) {
          counts[item.LocationID] = 0;
        }
        const quantity = parseInt(item.Quantity) || 0;
        counts[item.LocationID] += quantity;
      }
    });
  }
  
  return counts;
}

// ========== 地點管理功能 ==========

function addLocations(locationNames) {
  try {
    if (!locationNames || locationNames.length === 0) {
      throw new Error('地點名稱不能為空');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Locations');
    
    const existingData = getSheetDataAsObjects(ss, 'Locations');
    const existingNames = existingData.map(item => item.LocationName);
    
    for (let name of locationNames) {
      if (existingNames.includes(name)) {
        throw new Error(`地點名稱 "${name}" 已存在`);
      }
    }
    
    const newRows = locationNames.map(name => [
      generateId('L'),
      name
    ]);
    
    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 2).setValues(newRows);
    }
    
    console.log(`成功新增 ${newRows.length} 個地點`);
    return true;
    
  } catch (error) {
    console.error('addLocations 錯誤:', error);
    throw error;
  }
}

function updateLocationName(locationId, newName) {
  try {
    if (!locationId || !newName) {
      throw new Error('地點ID和新名稱不能為空');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Locations');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === newName && data[i][0] !== locationId) {
        throw new Error(`地點名稱 "${newName}" 已存在`);
      }
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === locationId) {
        sheet.getRange(i + 1, 2).setValue(newName);
        console.log(`成功更新地點 ${locationId} 的名稱為 "${newName}"`);
        return true;
      }
    }
    
    throw new Error(`找不到地點ID: ${locationId}`);
    
  } catch (error) {
    console.error('updateLocationName 錯誤:', error);
    throw error;
  }
}

function deleteLocationsAndTransferInventory(locationIds, transferToLocationId) {
  try {
    if (!locationIds || locationIds.length === 0) {
      throw new Error('請選擇要刪除的地點');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (transferToLocationId) {
      transferInventoryBetweenLocations(locationIds, transferToLocationId);
      console.log(`庫存已轉移至地點 ${transferToLocationId}`);
    } else {
      deleteInventoryByLocations(locationIds);
      console.log('相關庫存已刪除');
    }
    
    deleteRowsByIds(ss, 'Locations', locationIds, 0);
    console.log(`成功刪除 ${locationIds.length} 個地點`);
    
    return true;
    
  } catch (error) {
    console.error('deleteLocationsAndTransferInventory 錯誤:', error);
    throw error;
  }
}

// ========== 產品管理功能 ==========

function addProducts(products) {
  try {
    if (!products || products.length === 0) {
      throw new Error('產品資料不能為空');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Products');
    
    const existingData = getSheetDataAsObjects(ss, 'Products');
    const existingProducts = existingData.map(item => 
      `${item.ProductName}-${item.Brand}-${item.Category}`
    );
    
    const newRows = products.map(product => {
      if (!product.ProductName || !product.Brand || !product.Category) {
        throw new Error('產品名稱、品牌和類別都不能為空');
      }
      
      const productKey = `${product.ProductName}-${product.Brand}-${product.Category}`;
      if (existingProducts.includes(productKey)) {
        throw new Error(`產品 "${product.ProductName}" (${product.Brand} - ${product.Category}) 已存在`);
      }
      
      return [
        generateId('P'),
        product.ProductName,
        product.Brand,
        product.Category
      ];
    });
    
    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
    }
    
    console.log(`成功新增 ${newRows.length} 個產品`);
    return true;
    
  } catch (error) {
    console.error('addProducts 錯誤:', error);
    throw error;
  }
}

function updateProduct(productId, newProductName, newBrand, newCategory) {
  try {
    if (!productId || !newProductName || !newBrand || !newCategory) {
      throw new Error('所有欄位都不能為空');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Products');
    const data = sheet.getDataRange().getValues();
    
    const newProductKey = `${newProductName}-${newBrand}-${newCategory}`;
    for (let i = 1; i < data.length; i++) {
      const existingKey = `${data[i][1]}-${data[i][2]}-${data[i][3]}`;
      if (existingKey === newProductKey && data[i][0] !== productId) {
        throw new Error(`產品 "${newProductName}" (${newBrand} - ${newCategory}) 已存在`);
      }
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productId) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[newProductName, newBrand, newCategory]]);
        console.log(`成功更新產品 ${productId}`);
        return true;
      }
    }
    
    throw new Error(`找不到產品ID: ${productId}`);
    
  } catch (error) {
    console.error('updateProduct 錯誤:', error);
    throw error;
  }
}

function deleteProducts(productIds) {
  try {
    if (!productIds || productIds.length === 0) {
      throw new Error('請選擇要刪除的產品');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    deleteInventoryByProducts(productIds);
    console.log('相關庫存已刪除');
    
    deleteRowsByIds(ss, 'Products', productIds, 0);
    console.log(`成功刪除 ${productIds.length} 個產品`);
    
    return true;
    
  } catch (error) {
    console.error('deleteProducts 錯誤:', error);
    throw error;
  }
}

// ========== 交易管理功能 ==========

function addBulkTransactionsAndUpdateInventory(transactions) {
  try {
    if (!transactions || transactions.length === 0) {
      throw new Error('交易資料不能為空');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transactionSheet = ss.getSheetByName('Transactions');
    
    const timestamp = new Date().toLocaleString('zh-TW');
    
    for (let transaction of transactions) {
      if (!transaction.ProductID || !transaction.LocationID || !transaction.Quantity || !transaction.Type) {
        throw new Error('交易資料不完整');
      }
      
      if (transaction.Type === 'Out') {
        const currentStock = getCurrentStock(transaction.ProductID, transaction.LocationID);
        if (currentStock < transaction.Quantity) {
          throw new Error(`庫存不足：當前庫存 ${currentStock}，要求出庫 ${transaction.Quantity}`);
        }
      }
    }
    
    const transactionRows = transactions.map(t => [
      generateId('T'),
      timestamp,
      t.Type,
      t.ProductID,
      t.LocationID,
      parseInt(t.Quantity),
      t.Notes || ''
    ]);
    
    if (transactionRows.length > 0) {
      transactionSheet.getRange(transactionSheet.getLastRow() + 1, 1, transactionRows.length, 7).setValues(transactionRows);
    }
    
    transactions.forEach(transaction => {
      const quantityChange = transaction.Type === 'In' ? parseInt(transaction.Quantity) : -parseInt(transaction.Quantity);
      updateInventoryQuantity(transaction.ProductID, transaction.LocationID, quantityChange);
    });
    
    console.log(`成功新增 ${transactions.length} 筆交易`);
    return true;
    
  } catch (error) {
    console.error('addBulkTransactionsAndUpdateInventory 錯誤:', error);
    throw error;
  }
}

function updateTransactionAndInventory(transactionId, newNotes, newQuantity) {
  try {
    if (!transactionId || newQuantity <= 0) {
      throw new Error('交易ID不能為空，數量必須大於0');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transactionSheet = ss.getSheetByName('Transactions');
    const data = transactionSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === transactionId) {
        const oldQuantity = parseInt(data[i][5]);
        const productId = data[i][3];
        const locationId = data[i][4];
        const type = data[i][2];
        
        if (type === 'Out') {
          const currentStock = getCurrentStock(productId, locationId);
          const availableStock = currentStock + oldQuantity;
          if (availableStock < newQuantity) {
            throw new Error(`庫存不足：可用庫存 ${availableStock}，要求出庫 ${newQuantity}`);
          }
        }
        
        transactionSheet.getRange(i + 1, 6, 1, 2).setValues([[newQuantity, newNotes || '']]);
        
        const oldEffect = type === 'In' ? -oldQuantity : oldQuantity;
        const newEffect = type === 'In' ? newQuantity : -newQuantity;
        updateInventoryQuantity(productId, locationId, oldEffect + newEffect);
        
        console.log(`成功更新交易 ${transactionId}`);
        return true;
      }
    }
    
    throw new Error(`找不到交易ID: ${transactionId}`);
    
  } catch (error) {
    console.error('updateTransactionAndInventory 錯誤:', error);
    throw error;
  }
}

function deleteTransactionsAndUpdateInventory(transactionIds) {
  try {
    if (!transactionIds || transactionIds.length === 0) {
      throw new Error('請選擇要刪除的交易');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transactionSheet = ss.getSheetByName('Transactions');
    const data = transactionSheet.getDataRange().getValues();
    
    transactionIds.forEach(transactionId => {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === transactionId) {
          const quantity = parseInt(data[i][5]);
          const productId = data[i][3];
          const locationId = data[i][4];
          const type = data[i][2];
          
          const effect = type === 'In' ? -quantity : quantity;
          updateInventoryQuantity(productId, locationId, effect);
          break;
        }
      }
    });
    
    deleteRowsByIds(ss, 'Transactions', transactionIds, 0);
    console.log(`成功刪除 ${transactionIds.length} 筆交易`);
    
    return true;
    
  } catch (error) {
    console.error('deleteTransactionsAndUpdateInventory 錯誤:', error);
    throw error;
  }
}

// ========== 庫存管理輔助函數 ==========

function getCurrentStock(productId, locationId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Inventory');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === productId && data[i][2] === locationId) {
      return parseInt(data[i][3]) || 0;
    }
  }
  
  return 0;
}

function updateInventoryQuantity(productId, locationId, quantityChange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Inventory');
  const data = sheet.getDataRange().getValues();
  
  let found = false;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === productId && data[i][2] === locationId) {
      const newQuantity = parseInt(data[i][3]) + quantityChange;
      
      if (newQuantity <= 0) {
        sheet.deleteRow(i + 1);
      } else {
        sheet.getRange(i + 1, 4).setValue(newQuantity);
      }
      
      found = true;
      break;
    }
  }
  
  if (!found && quantityChange > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 4).setValues([[
      generateId('I'),
      productId,
      locationId,
      quantityChange
    ]]);
  }
}

function transferInventoryBetweenLocations(fromLocationIds, toLocationId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Inventory');
  const data = sheet.getDataRange().getValues();
  
  const transferData = {};
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (fromLocationIds.includes(data[i][2])) {
      const productId = data[i][1];
      const quantity = parseInt(data[i][3]);
      
      if (!transferData[productId]) {
        transferData[productId] = 0;
      }
      transferData[productId] += quantity;
      
      sheet.deleteRow(i + 1);
    }
  }
  
  Object.keys(transferData).forEach(productId => {
    updateInventoryQuantity(productId, toLocationId, transferData[productId]);
  });
}

function deleteInventoryByLocations(locationIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Inventory');
  const data = sheet.getDataRange().getValues();
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (locationIds.includes(data[i][2])) {
      sheet.deleteRow(i + 1);
    }
  }
}

function deleteInventoryByProducts(productIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Inventory');
  const data = sheet.getDataRange().getValues();
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (productIds.includes(data[i][1])) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ========== 通用輔助函數 ==========

function generateId(prefix) {
  const timestamp = Date.now();
  const random = Math.floor(Math.random() * 1000);
  return `${prefix}${timestamp}${random}`;
}

function deleteRowsByIds(spreadsheet, sheetName, ids, idColumnIndex) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (ids.includes(data[i][idColumnIndex])) {
      sheet.deleteRow(i + 1);
    }
  }
}
