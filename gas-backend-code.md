# Google Apps Script 後端完整程式碼

## 1. Code.gs - 主要後端邏輯

```javascript
// 全域變數設定
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

// 主要的 doGet 函數 - 處理 GET 請求
function doGet(e) {
  try {
    const action = e.parameter.action;
    
    switch(action) {
      case 'getProducts':
        return handleGetProducts();
      case 'getLocations':
        return handleGetLocations();
      case 'getInventory':
        return handleGetInventory();
      case 'getRecords':
        return handleGetRecords();
      case 'getAllData':
        return handleGetAllData();
      default:
        return createJsonResponse({ success: false, message: '未知的操作' });
    }
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 主要的 doPost 函數 - 處理 POST 請求
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    switch(action) {
      case 'saveProduct':
        return handleSaveProduct(data);
      case 'deleteProduct':
        return handleDeleteProduct(data);
      case 'saveLocation':
        return handleSaveLocation(data);
      case 'deleteLocation':
        return handleDeleteLocation(data);
      case 'saveStockTransaction':
        return handleSaveStockTransaction(data);
      case 'updateInventory':
        return handleUpdateInventory(data);
      case 'testConnection':
        return handleTestConnection();
      default:
        return createJsonResponse({ success: false, message: '未知的操作' });
    }
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 工具函數 - 創建 JSON 回應
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 工具函數 - 取得試算表
function getSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('請先設定試算表 ID');
  }
  return SpreadsheetApp.openById(id);
}

// 工具函數 - 取得工作表
function getSheet(sheetName) {
  const spreadsheet = getSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    initializeSheet(sheet, sheetName);
  }
  
  return sheet;
}

// 初始化工作表表頭
function initializeSheet(sheet, sheetName) {
  let headers = [];
  
  switch(sheetName) {
    case 'products':
      headers = ['id', 'name', 'description', 'unit', 'created_at', 'updated_at'];
      break;
    case 'locations':
      headers = ['id', 'name', 'address', 'description', 'created_at', 'updated_at'];
      break;
    case 'inventory':
      headers = ['location_id', 'product_id', 'quantity', 'updated_at'];
      break;
    case 'records':
      headers = ['id', 'type', 'location_id', 'product_id', 'quantity', 'operator', 'timestamp', 'notes'];
      break;
  }
  
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

// 處理取得商品資料
function handleGetProducts() {
  try {
    const sheet = getSheet('products');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return createJsonResponse({ success: true, data: [] });
    }
    
    const headers = data[0];
    const products = data.slice(1).map(row => {
      const product = {};
      headers.forEach((header, index) => {
        product[header] = row[index];
      });
      return product;
    });
    
    return createJsonResponse({ success: true, data: products });
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 處理取得地點資料
function handleGetLocations() {
  try {
    const sheet = getSheet('locations');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return createJsonResponse({ success: true, data: [] });
    }
    
    const headers = data[0];
    const locations = data.slice(1).map(row => {
      const location = {};
      headers.forEach((header, index) => {
        location[header] = row[index];
      });
      return location;
    });
    
    return createJsonResponse({ success: true, data: locations });
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 處理取得庫存資料
function handleGetInventory() {
  try {
    const sheet = getSheet('inventory');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return createJsonResponse({ success: true, data: [] });
    }
    
    const headers = data[0];
    const inventory = data.slice(1).map(row => {
      const item = {};
      headers.forEach((header, index) => {
        item[header] = row[index];
      });
      return item;
    });
    
    return createJsonResponse({ success: true, data: inventory });
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 處理取得異動記錄
function handleGetRecords() {
  try {
    const sheet = getSheet('records');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return createJsonResponse({ success: true, data: [] });
    }
    
    const headers = data[0];
    const records = data.slice(1).map(row => {
      const record = {};
      headers.forEach((header, index) => {
        record[header] = row[index];
      });
      return record;
    }).sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    return createJsonResponse({ success: true, data: records });
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 處理取得所有資料
function handleGetAllData() {
  try {
    const products = JSON.parse(handleGetProducts().getContent()).data;
    const locations = JSON.parse(handleGetLocations().getContent()).data;
    const inventory = JSON.parse(handleGetInventory().getContent()).data;
    const records = JSON.parse(handleGetRecords().getContent()).data;
    
    return createJsonResponse({
      success: true,
      data: {
        products: products,
        locations: locations,
        inventory: inventory,
        records: records
      }
    });
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 處理儲存商品
function handleSaveProduct(data) {
  try {
    const sheet = getSheet('products');
    const product = data.product;
    const timestamp = new Date().toISOString();
    
    if (product.id) {
      // 更新現有商品
      const existingData = sheet.getDataRange().getValues();
      const headers = existingData[0];
      const rowIndex = existingData.findIndex(row => row[0] == product.id);
      
      if (rowIndex > 0) {
        product.updated_at = timestamp;
        const values = headers.map(header => product[header] || '');
        sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([values]);
      }
    } else {
      // 新增商品
      product.id = generateId();
      product.created_at = timestamp;
      product.updated_at = timestamp;
      
      const values = [product.id, product.name, product.description, product.unit, product.created_at, product.updated_at];
      sheet.appendRow(values);
    }
    
    return createJsonResponse({ success: true, message: '商品已儲存' });
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 處理刪除商品
function handleDeleteProduct(data) {
  try {
    const sheet = getSheet('products');
    const productId = data.productId;
    const existingData = sheet.getDataRange().getValues();
    const rowIndex = existingData.findIndex(row => row[0] == productId);
    
    if (rowIndex > 0) {
      sheet.deleteRow(rowIndex + 1);
      
      // 同時清理相關的庫存資料
      const inventorySheet = getSheet('inventory');
      const inventoryData = inventorySheet.getDataRange().getValues();
      for (let i = inventoryData.length - 1; i >= 1; i--) {
        if (inventoryData[i][1] == productId) {
          inventorySheet.deleteRow(i + 1);
        }
      }
    }
    
    return createJsonResponse({ success: true, message: '商品已刪除' });
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 處理儲存地點
function handleSaveLocation(data) {
  try {
    const sheet = getSheet('locations');
    const location = data.location;
    const timestamp = new Date().toISOString();
    
    if (location.id) {
      // 更新現有地點
      const existingData = sheet.getDataRange().getValues();
      const headers = existingData[0];
      const rowIndex = existingData.findIndex(row => row[0] == location.id);
      
      if (rowIndex > 0) {
        location.updated_at = timestamp;
        const values = headers.map(header => location[header] || '');
        sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([values]);
      }
    } else {
      // 新增地點
      location.id = generateId();
      location.created_at = timestamp;
      location.updated_at = timestamp;
      
      const values = [location.id, location.name, location.address, location.description, location.created_at, location.updated_at];
      sheet.appendRow(values);
    }
    
    return createJsonResponse({ success: true, message: '地點已儲存' });
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 處理刪除地點
function handleDeleteLocation(data) {
  try {
    const sheet = getSheet('locations');
    const locationId = data.locationId;
    const existingData = sheet.getDataRange().getValues();
    const rowIndex = existingData.findIndex(row => row[0] == locationId);
    
    if (rowIndex > 0) {
      sheet.deleteRow(rowIndex + 1);
      
      // 同時清理相關的庫存資料
      const inventorySheet = getSheet('inventory');
      const inventoryData = inventorySheet.getDataRange().getValues();
      for (let i = inventoryData.length - 1; i >= 1; i--) {
        if (inventoryData[i][0] == locationId) {
          inventorySheet.deleteRow(i + 1);
        }
      }
    }
    
    return createJsonResponse({ success: true, message: '地點已刪除' });
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 處理進出貨異動
function handleSaveStockTransaction(data) {
  try {
    const { type, location_id, product_id, quantity, operator, notes } = data;
    
    // 記錄異動
    const recordsSheet = getSheet('records');
    const recordId = generateId();
    const timestamp = new Date().toISOString();
    
    recordsSheet.appendRow([
      recordId,
      type,
      location_id,
      product_id,
      quantity,
      operator,
      timestamp,
      notes || ''
    ]);
    
    // 更新庫存
    updateInventoryQuantity(location_id, product_id, type === '進貨' ? quantity : -quantity);
    
    return createJsonResponse({ success: true, message: '異動記錄已儲存' });
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 更新庫存數量
function updateInventoryQuantity(locationId, productId, quantityChange) {
  const sheet = getSheet('inventory');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 查找現有庫存記錄
  let existingIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == locationId && data[i][1] == productId) {
      existingIndex = i;
      break;
    }
  }
  
  const timestamp = new Date().toISOString();
  
  if (existingIndex > -1) {
    // 更新現有記錄
    const currentQuantity = parseFloat(data[existingIndex][2]) || 0;
    const newQuantity = Math.max(0, currentQuantity + quantityChange);
    
    sheet.getRange(existingIndex + 1, 3).setValue(newQuantity);
    sheet.getRange(existingIndex + 1, 4).setValue(timestamp);
  } else {
    // 創建新記錄
    const newQuantity = Math.max(0, quantityChange);
    sheet.appendRow([locationId, productId, newQuantity, timestamp]);
  }
}

// 處理更新庫存
function handleUpdateInventory(data) {
  try {
    const { location_id, product_id, quantity } = data;
    
    const sheet = getSheet('inventory');
    const existingData = sheet.getDataRange().getValues();
    
    let updated = false;
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] == location_id && existingData[i][1] == product_id) {
        sheet.getRange(i + 1, 3).setValue(quantity);
        sheet.getRange(i + 1, 4).setValue(new Date().toISOString());
        updated = true;
        break;
      }
    }
    
    if (!updated) {
      sheet.appendRow([location_id, product_id, quantity, new Date().toISOString()]);
    }
    
    return createJsonResponse({ success: true, message: '庫存已更新' });
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 處理連接測試
function handleTestConnection() {
  try {
    const spreadsheet = getSpreadsheet();
    const sheetCount = spreadsheet.getSheets().length;
    
    return createJsonResponse({
      success: true,
      message: '連接成功',
      data: {
        spreadsheetName: spreadsheet.getName(),
        sheetCount: sheetCount,
        timestamp: new Date().toISOString()
      }
    });
  } catch (error) {
    return createJsonResponse({ success: false, message: error.toString() });
  }
}

// 生成唯一 ID
function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

// 設定試算表 ID（一次性執行）
function setSpreadsheetId() {
  const spreadsheetId = '請輸入你的 Google Sheets ID';
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
  console.log('試算表 ID 已設定:', spreadsheetId);
}

// 初始化試算表結構（一次性執行）
function initializeSpreadsheet() {
  try {
    const spreadsheet = getSpreadsheet();
    
    // 創建必要的工作表
    const sheetNames = ['products', 'locations', 'inventory', 'records'];
    
    sheetNames.forEach(sheetName => {
      let sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
        initializeSheet(sheet, sheetName);
        console.log('已創建工作表:', sheetName);
      }
    });
    
    console.log('試算表初始化完成');
  } catch (error) {
    console.error('初始化失敗:', error.toString());
  }
}
```

## 2. 部署與設定步驟

### 步驟 1：建立 Google Sheets
1. 開啟 Google Sheets，建立新的試算表
2. 複製試算表的 ID（從網址中取得）
3. 設定試算表權限為「知道連結的任何人都可編輯」

### 步驟 2：建立 Google Apps Script 專案
1. 在 Google Apps Script (script.google.com) 建立新專案
2. 將上面的程式碼貼到 `Code.gs` 檔案中
3. 修改 `setSpreadsheetId()` 函數中的試算表 ID
4. 執行 `setSpreadsheetId()` 函數設定試算表 ID
5. 執行 `initializeSpreadsheet()` 函數初始化試算表結構

### 步驟 3：部署為網頁應用程式
1. 點選「部署」→「新增部署作業」
2. 選擇類型為「網頁應用程式」
3. 說明填寫「進銷存系統 API」
4. 執行身分選擇「我」
5. 存取權限選擇「任何人」
6. 點選「部署」並複製網頁應用程式 URL

### 步驟 4：整合前端系統
1. 開啟前端應用程式
2. 進入「設定」頁面
3. 將 Google Apps Script 的網頁應用程式 URL 貼到「GAS API URL」欄位
4. 點選「測試連接」確認連接成功
5. 啟用「自動同步」功能

## 3. API 介面說明

### GET 請求格式
```
GET {GAS_URL}?action={action_name}
```

支援的 action：
- `getProducts` - 取得所有商品
- `getLocations` - 取得所有地點
- `getInventory` - 取得庫存資料
- `getRecords` - 取得異動記錄
- `getAllData` - 取得所有資料

### POST 請求格式
```
POST {GAS_URL}
Content-Type: application/json

{
  "action": "action_name",
  "data": { ... }
}
```

支援的 action：
- `saveProduct` - 儲存商品
- `deleteProduct` - 刪除商品
- `saveLocation` - 儲存地點
- `deleteLocation` - 刪除地點
- `saveStockTransaction` - 儲存進出貨異動
- `updateInventory` - 更新庫存
- `testConnection` - 測試連接

## 4. 資料表結構

### products 工作表
| 欄位 | 說明 |
|------|------|
| id | 商品ID |
| name | 商品名稱 |
| description | 商品描述 |
| unit | 單位 |
| created_at | 建立時間 |
| updated_at | 更新時間 |

### locations 工作表
| 欄位 | 說明 |
|------|------|
| id | 地點ID |
| name | 地點名稱 |
| address | 地址 |
| description | 描述 |
| created_at | 建立時間 |
| updated_at | 更新時間 |

### inventory 工作表
| 欄位 | 說明 |
|------|------|
| location_id | 地點ID |
| product_id | 商品ID |
| quantity | 數量 |
| updated_at | 更新時間 |

### records 工作表
| 欄位 | 說明 |
|------|------|
| id | 記錄ID |
| type | 異動類型（進貨/出貨） |
| location_id | 地點ID |
| product_id | 商品ID |
| quantity | 數量 |
| operator | 操作人員 |
| timestamp | 時間戳記 |
| notes | 備註 |

## 5. 注意事項

1. **權限設定**：確保 Google Sheets 和 Apps Script 都設定為公開存取
2. **資料備份**：建議定期備份 Google Sheets 資料
3. **效能考量**：Google Apps Script 有執行時間限制，大量資料時可能需要分批處理
4. **錯誤處理**：系統已包含基本的錯誤處理機制
5. **安全性**：由於是公開存取，請勿儲存敏感資料

## 6. 疑難排解

### 常見問題：
1. **403 錯誤**：檢查部署權限設定
2. **找不到試算表**：確認試算表 ID 是否正確
3. **資料同步失敗**：檢查網路連接和 API URL
4. **授權問題**：重新授權 Apps Script 權限

這個整合方案提供了完整的後端 API 功能，搭配前端應用程式可以實現完整的進銷存管理系統。