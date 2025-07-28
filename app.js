// 全域變數
let currentPage = 'dashboard';
let products = [];
let locations = [];
let inventory = [];
let records = [];
let lowStockThreshold = 5;

// GAS 整合變數
let gasUrl = '';
let autoSyncEnabled = false;
let syncInterval = 5;
let lastSyncTime = null;
let syncIntervalId = null;
let isOnline = navigator.onLine;

// 初始化應用程式
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
    setupNetworkListeners();
});

function initializeApp() {
    loadDataFromStorage();
    loadGasSettings();
    initializeSampleData();
    setupEventListeners();
    updateDashboard();
    renderProducts();
    renderLocations();
    updateSelectOptions();
    renderInventory();
    renderRecords();
    updateLastSyncTime();
    showPage('dashboard');
    
    // 啟動自動同步（如果已啟用且在線）
    if (autoSyncEnabled && gasUrl && isOnline) {
        startAutoSync();
    }
}

// 網路狀態監聽
function setupNetworkListeners() {
    window.addEventListener('online', () => {
        isOnline = true;
        updateGasStatus();
        showNotification('網路連線已恢復');
        if (autoSyncEnabled && gasUrl) {
            startAutoSync();
        }
    });
    
    window.addEventListener('offline', () => {
        isOnline = false;
        updateGasStatus();
        showNotification('網路連線中斷，切換到離線模式', 'warning');
        stopAutoSync();
    });
}

// 載入資料從 localStorage
function loadDataFromStorage() {
    try {
        products = JSON.parse(localStorage.getItem('products')) || [];
        locations = JSON.parse(localStorage.getItem('locations')) || [];
        inventory = JSON.parse(localStorage.getItem('inventory')) || [];
        records = JSON.parse(localStorage.getItem('records')) || [];
        lowStockThreshold = parseInt(localStorage.getItem('lowStockThreshold')) || 5;
        
        // 設定庫存警告閾值
        setTimeout(() => {
            const thresholdInput = document.getElementById('low-stock-threshold');
            if (thresholdInput) {
                thresholdInput.value = lowStockThreshold;
            }
        }, 100);
    } catch (error) {
        console.error('載入本地資料失敗:', error);
        showNotification('載入本地資料失敗', 'error');
    }
}

// 載入 GAS 設定
function loadGasSettings() {
    gasUrl = localStorage.getItem('gasUrl') || '';
    autoSyncEnabled = localStorage.getItem('autoSyncEnabled') === 'true';
    syncInterval = parseInt(localStorage.getItem('syncInterval')) || 5;
    lastSyncTime = localStorage.getItem('lastSyncTime');
    
    setTimeout(() => {
        const gasUrlInput = document.getElementById('gas-url');
        const autoSyncCheckbox = document.getElementById('auto-sync-enabled');
        const syncIntervalInput = document.getElementById('sync-interval');
        
        if (gasUrlInput) gasUrlInput.value = gasUrl;
        if (autoSyncCheckbox) autoSyncCheckbox.checked = autoSyncEnabled;
        if (syncIntervalInput) syncIntervalInput.value = syncInterval;
        
        updateGasStatus();
    }, 100);
}

// 儲存 GAS 設定
function saveGasSettings() {
    localStorage.setItem('gasUrl', gasUrl);
    localStorage.setItem('autoSyncEnabled', autoSyncEnabled.toString());
    localStorage.setItem('syncInterval', syncInterval.toString());
    if (lastSyncTime) {
        localStorage.setItem('lastSyncTime', lastSyncTime);
    }
}

// 儲存資料到 localStorage
function saveDataToStorage() {
    try {
        localStorage.setItem('products', JSON.stringify(products));
        localStorage.setItem('locations', JSON.stringify(locations));
        localStorage.setItem('inventory', JSON.stringify(inventory));
        localStorage.setItem('records', JSON.stringify(records));
        localStorage.setItem('lowStockThreshold', lowStockThreshold.toString());
    } catch (error) {
        console.error('儲存本地資料失敗:', error);
        showNotification('儲存本地資料失敗', 'error');
    }
}

// 初始化範例資料
function initializeSampleData() {
    if (products.length === 0 && locations.length === 0) {
        // 使用提供的範例資料
        locations = [
            {id: "L001", name: "倉庫1", address: "台北市信義區", description: "主要倉庫"},
            {id: "L002", name: "倉庫2", address: "新北市板橋區", description: "次要倉庫"}
        ];
        
        products = [
            {id: "P001", name: "商品A", description: "商品A描述", unit: "個"},
            {id: "P002", name: "商品B", description: "商品B描述", unit: "箱"}
        ];
        
        inventory = [
            {locationId: "L001", productId: "P001", quantity: 100, lastUpdated: new Date().toISOString()},
            {locationId: "L001", productId: "P002", quantity: 200, lastUpdated: new Date().toISOString()},
            {locationId: "L002", productId: "P001", quantity: 50, lastUpdated: new Date().toISOString()},
            {locationId: "L002", productId: "P002", quantity: 0, lastUpdated: new Date().toISOString()}
        ];
        
        records = [
            {
                id: generateId(),
                type: 'in',
                locationId: "L001",
                productId: "P001",
                quantity: 50,
                operator: "張三",
                timestamp: new Date().toISOString(),
                note: "定期補貨"
            },
            {
                id: generateId(),
                type: 'out',
                locationId: "L001",
                productId: "P001",
                quantity: 20,
                operator: "李四",
                timestamp: new Date().toISOString(),
                note: "客戶訂單"
            }
        ];
        
        saveDataToStorage();
    }
}

// Google Apps Script 整合功能 - 使用最佳實踐避免CORS
async function testGasConnection() {
    const gasUrlInput = document.getElementById('gas-url');
    const statusDiv = document.getElementById('gas-connection-status');
    
    if (!gasUrlInput || !statusDiv) return;
    
    const url = gasUrlInput.value.trim();
    if (!url) {
        statusDiv.innerHTML = '<div class="connection-status error">請輸入 GAS URL</div>';
        return;
    }
    
    if (!isOnline) {
        statusDiv.innerHTML = '<div class="connection-status error">請檢查網路連線</div>';
        return;
    }
    
    statusDiv.innerHTML = '<div class="connection-status testing">正在測試連接...</div>';
    
    try {
        // 使用 FormData 避免觸發預檢請求
        const formData = new FormData();
        formData.append('action', 'test');
        
        const response = await fetch(url, {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            const text = await response.text();
            let data;
            try {
                data = JSON.parse(text);
            } catch {
                // 如果不是JSON，可能是HTML錯誤頁面
                throw new Error('GAS回應格式錯誤');
            }
            
            if (data.success) {
                statusDiv.innerHTML = '<div class="connection-status success">✓ 連接成功！GAS 整合已準備就緒</div>';
                gasUrl = url;
                updateGasStatus();
            } else {
                statusDiv.innerHTML = '<div class="connection-status error">✗ GAS 回應錯誤：' + (data.error || '未知錯誤') + '</div>';
            }
        } else {
            statusDiv.innerHTML = '<div class="connection-status error">✗ 連接失敗：HTTP ' + response.status + '</div>';
        }
    } catch (error) {
        statusDiv.innerHTML = '<div class="connection-status error">✗ 連接失敗：' + error.message + '</div>';
    }
}

function saveGasSettingsHandler() {
    const gasUrlInput = document.getElementById('gas-url');
    const autoSyncCheckbox = document.getElementById('auto-sync-enabled');
    const syncIntervalInput = document.getElementById('sync-interval');
    
    if (!gasUrlInput || !autoSyncCheckbox || !syncIntervalInput) return;
    
    gasUrl = gasUrlInput.value.trim();
    autoSyncEnabled = autoSyncCheckbox.checked;
    syncInterval = parseInt(syncIntervalInput.value) || 5;
    
    saveGasSettings();
    updateGasStatus();
    
    if (autoSyncEnabled && gasUrl && isOnline) {
        startAutoSync();
    } else {
        stopAutoSync();
    }
    
    showNotification('GAS 設定已儲存');
}

function updateGasStatus() {
    const statusElement = document.getElementById('gas-status');
    if (!statusElement) return;
    
    if (!isOnline) {
        statusElement.textContent = '離線模式';
        statusElement.className = 'gas-status offline';
    } else if (gasUrl) {
        statusElement.textContent = '已連接';
        statusElement.className = 'gas-status online';
    } else {
        statusElement.textContent = '未設定';
        statusElement.className = 'gas-status offline';
    }
}

// 改進的同步函數，使用FormData避免CORS問題
async function syncDataToGas(dataType, data, retryCount = 3) {
    if (!gasUrl || !isOnline) return false;
    
    for (let i = 0; i < retryCount; i++) {
        try {
            const formData = new FormData();
            formData.append('action', 'sync');
            formData.append('dataType', dataType);
            formData.append('data', JSON.stringify(data));
            
            const response = await fetch(gasUrl, {
                method: 'POST',
                body: formData
            });
            
            if (response.ok) {
                const text = await response.text();
                const result = JSON.parse(text);
                return result.success;
            }
        } catch (error) {
            console.error(`同步嘗試 ${i + 1} 失敗:`, error);
            if (i === retryCount - 1) {
                throw error;
            }
            // 等待一段時間後重試
            await new Promise(resolve => setTimeout(resolve, 1000 * (i + 1)));
        }
    }
    return false;
}

async function syncAllDataToGas() {
    if (!gasUrl) {
        showNotification('請先設定 GAS URL', 'error');
        return;
    }
    
    if (!isOnline) {
        showNotification('請檢查網路連線', 'error');
        return;
    }
    
    showLoading('正在同步所有資料...');
    
    try {
        const allData = {
            products,
            locations,
            inventory,
            records,
            lastSync: new Date().toISOString()
        };
        
        const success = await syncDataToGas('all', allData);
        
        if (success) {
            lastSyncTime = new Date().toISOString();
            localStorage.setItem('lastSyncTime', lastSyncTime);
            updateLastSyncTime();
            showNotification('所有資料同步成功');
        } else {
            showNotification('資料同步失敗', 'error');
        }
    } catch (error) {
        console.error('同步錯誤:', error);
        showNotification('同步過程發生錯誤：' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

async function pullDataFromGas() {
    if (!gasUrl) {
        showNotification('請先設定 GAS URL', 'error');
        return;
    }
    
    if (!isOnline) {
        showNotification('請檢查網路連線', 'error');
        return;
    }
    
    showLoading('正在從雲端拉取資料...');
    
    try {
        const formData = new FormData();
        formData.append('action', 'pull');
        
        const response = await fetch(gasUrl, {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            const text = await response.text();
            const result = JSON.parse(text);
            
            if (result.success && result.data) {
                products = result.data.products || products;
                locations = result.data.locations || locations;
                inventory = result.data.inventory || inventory;
                records = result.data.records || records;
                
                saveDataToStorage();
                initializeApp();
                showNotification('資料拉取成功');
            } else {
                showNotification('資料拉取失敗：' + (result.error || '未知錯誤'), 'error');
            }
        } else {
            showNotification('連接失敗：HTTP ' + response.status, 'error');
        }
    } catch (error) {
        console.error('拉取錯誤:', error);
        showNotification('拉取過程發生錯誤：' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

function startAutoSync() {
    stopAutoSync(); // 先停止現有的同步
    
    if (gasUrl && autoSyncEnabled && isOnline) {
        syncIntervalId = setInterval(() => {
            syncAllDataToGas();
        }, syncInterval * 60 * 1000); // 轉換為毫秒
    }
}

function stopAutoSync() {
    if (syncIntervalId) {
        clearInterval(syncIntervalId);
        syncIntervalId = null;
    }
}

function updateLastSyncTime() {
    const syncStatusElement = document.getElementById('last-sync-time');
    if (!syncStatusElement) return;
    
    if (lastSyncTime) {
        const syncDate = new Date(lastSyncTime);
        syncStatusElement.innerHTML = `<div class="sync-status success">上次同步：${syncDate.toLocaleString('zh-TW')}</div>`;
    } else {
        syncStatusElement.innerHTML = '<div class="sync-status">尚未同步</div>';
    }
}

function showGasCode() {
    const gasCode = `/**
 * Google Apps Script 進銷存管理系統後端
 * 使用最佳實踐避免CORS問題
 * 支援FormData請求，不會觸發OPTIONS預檢請求
 */

function doPost(e) {
  try {
    // 處理FormData請求
    const params = e.parameter;
    const action = params.action;
    
    // 設定CORS標頭
    const output = ContentService.createTextOutput();
    output.setMimeType(ContentService.MimeType.JSON);
    
    switch (action) {
      case 'test':
        return output.setContent(JSON.stringify({ 
          success: true, 
          message: '連接成功',
          timestamp: new Date().toISOString()
        }));
        
      case 'sync':
        return handleSync(params, output);
        
      case 'pull':
        return handlePull(output);
        
      default:
        return output.setContent(JSON.stringify({ 
          success: false, 
          error: '未知操作：' + action 
        }));
    }
  } catch (error) {
    Logger.log('錯誤：' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: false, 
        error: error.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleSync(params, output) {
  try {
    const dataType = params.dataType;
    const data = JSON.parse(params.data);
    
    const spreadsheet = getOrCreateSpreadsheet();
    
    if (dataType === 'all') {
      // 同步所有資料
      if (data.products) updateSheet(spreadsheet, '商品', data.products, ['id', 'name', 'description', 'unit']);
      if (data.locations) updateSheet(spreadsheet, '地點', data.locations, ['id', 'name', 'address', 'description']);
      if (data.inventory) updateSheet(spreadsheet, '庫存', data.inventory, ['locationId', 'productId', 'quantity', 'lastUpdated']);
      if (data.records) updateSheet(spreadsheet, '異動記錄', data.records, ['id', 'type', 'locationId', 'productId', 'quantity', 'operator', 'timestamp', 'note']);
    } else {
      // 同步特定類型資料
      const sheetNames = {
        'products': '商品',
        'locations': '地點',
        'inventory': '庫存',
        'records': '異動記錄'
      };
      
      const headers = {
        'products': ['id', 'name', 'description', 'unit'],
        'locations': ['id', 'name', 'address', 'description'],
        'inventory': ['locationId', 'productId', 'quantity', 'lastUpdated'],
        'records': ['id', 'type', 'locationId', 'productId', 'quantity', 'operator', 'timestamp', 'note']
      };
      
      if (sheetNames[dataType] && headers[dataType]) {
        updateSheet(spreadsheet, sheetNames[dataType], data, headers[dataType]);
      }
    }
    
    return output.setContent(JSON.stringify({ 
      success: true, 
      message: '同步完成',
      timestamp: new Date().toISOString()
    }));
    
  } catch (error) {
    Logger.log('同步錯誤：' + error.toString());
    return output.setContent(JSON.stringify({ 
      success: false, 
      error: '同步失敗：' + error.toString() 
    }));
  }
}

function handlePull(output) {
  try {
    const spreadsheet = getOrCreateSpreadsheet();
    
    const data = {
      products: getSheetData(spreadsheet, '商品', ['id', 'name', 'description', 'unit']),
      locations: getSheetData(spreadsheet, '地點', ['id', 'name', 'address', 'description']),
      inventory: getSheetData(spreadsheet, '庫存', ['locationId', 'productId', 'quantity', 'lastUpdated']),
      records: getSheetData(spreadsheet, '異動記錄', ['id', 'type', 'locationId', 'productId', 'quantity', 'operator', 'timestamp', 'note']),
      lastPull: new Date().toISOString()
    };
    
    return output.setContent(JSON.stringify({ 
      success: true, 
      data: data,
      timestamp: new Date().toISOString()
    }));
    
  } catch (error) {
    Logger.log('拉取錯誤：' + error.toString());
    return output.setContent(JSON.stringify({ 
      success: false, 
      error: '拉取失敗：' + error.toString() 
    }));
  }
}

function getOrCreateSpreadsheet() {
  // 嘗試獲取現有的試算表或建立新的
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (error) {
    // 如果沒有活動試算表，建立一個新的
    const spreadsheet = SpreadsheetApp.create('進銷存管理系統資料');
    Logger.log('建立新試算表：' + spreadsheet.getUrl());
    return spreadsheet;
  }
}

function updateSheet(spreadsheet, sheetName, data, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    // 設定標題列
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  // 清除現有資料（除標題列外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  
  // 寫入新資料
  if (data && data.length > 0) {
    const values = data.map(item => 
      headers.map(header => {
        const value = item[header];
        return value !== undefined && value !== null ? value.toString() : '';
      })
    );
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }
  
  // 自動調整欄寬
  sheet.autoResizeColumns(1, headers.length);
}

function getSheetData(spreadsheet, sheetName, headers) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  
  return values.map(row => {
    const item = {};
    headers.forEach((header, index) => {
      const value = row[index];
      item[header] = value === '' ? null : value;
    });
    return item;
  }).filter(item => {
    // 過濾掉空白行
    return Object.values(item).some(value => value !== null && value !== '');
  });
}`;

    showModal('Google Apps Script 程式碼', `
        <div class="code-header">
            <h4>GAS 程式碼</h4>
            <button class="copy-code-btn" onclick="copyToClipboard(\`${gasCode.replace(/`/g, '\\`').replace(/\$/g, '\\$')}\`)">複製程式碼</button>
        </div>
        <div class="code-block">${gasCode}</div>
        <p style="margin-top: 16px; font-size: 14px; color: var(--color-text-secondary);">
            請將此程式碼貼到您的 Google Apps Script 專案中，然後部署為 Web App。<br>
            <strong>重要：</strong>此版本使用FormData避免CORS問題，不會觸發OPTIONS預檢請求。
        </p>
    `);
}

function copyToClipboard(text) {
    navigator.clipboard.writeText(text).then(() => {
        showNotification('程式碼已複製到剪貼簿');
    }).catch(() => {
        // 降級處理
        const textArea = document.createElement('textarea');
        textArea.value = text;
        document.body.appendChild(textArea);
        textArea.select();
        try {
            document.execCommand('copy');
            showNotification('程式碼已複製到剪貼簿');
        } catch (err) {
            showNotification('複製失敗，請手動選取複製', 'error');
        }
        document.body.removeChild(textArea);
    });
}

// 設定事件監聽器
function setupEventListeners() {
    // 導航選單
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => {
        item.addEventListener('click', function(e) {
            e.preventDefault();
            const page = this.getAttribute('data-page');
            if (page) {
                showPage(page);
            }
        });
    });
    
    // GAS 相關事件
    setTimeout(() => {
        const testGasBtn = document.getElementById('test-gas-connection');
        const saveGasBtn = document.getElementById('save-gas-settings');
        const syncAllBtn = document.getElementById('sync-all-data');
        const manualSyncBtn = document.getElementById('manual-sync-all');
        const pullFromGasBtn = document.getElementById('pull-from-gas');
        const showGasCodeBtn = document.getElementById('show-gas-code');
        
        if (testGasBtn) testGasBtn.addEventListener('click', testGasConnection);
        if (saveGasBtn) saveGasBtn.addEventListener('click', saveGasSettingsHandler);
        if (syncAllBtn) syncAllBtn.addEventListener('click', syncAllDataToGas);
        if (manualSyncBtn) manualSyncBtn.addEventListener('click', syncAllDataToGas);
        if (pullFromGasBtn) pullFromGasBtn.addEventListener('click', pullDataFromGas);
        if (showGasCodeBtn) showGasCodeBtn.addEventListener('click', showGasCode);
        
        // 單獨同步按鈕
        const syncProductsBtn = document.getElementById('sync-products');
        const syncLocationsBtn = document.getElementById('sync-locations');
        const syncInventoryBtn = document.getElementById('sync-inventory');
        const syncRecordsBtn = document.getElementById('sync-records');
        
        if (syncProductsBtn) syncProductsBtn.addEventListener('click', () => syncIndividualData('products'));
        if (syncLocationsBtn) syncLocationsBtn.addEventListener('click', () => syncIndividualData('locations'));
        if (syncInventoryBtn) syncInventoryBtn.addEventListener('click', () => syncIndividualData('inventory'));
        if (syncRecordsBtn) syncRecordsBtn.addEventListener('click', () => syncIndividualData('records'));
    }, 100);
    
    // 商品管理
    const addProductBtn = document.getElementById('add-product-btn');
    if (addProductBtn) {
        addProductBtn.addEventListener('click', () => showProductModal());
    }
    
    const productSearch = document.getElementById('product-search');
    if (productSearch) {
        productSearch.addEventListener('input', filterProducts);
    }
    
    // 地點管理
    const addLocationBtn = document.getElementById('add-location-btn');
    if (addLocationBtn) {
        addLocationBtn.addEventListener('click', () => showLocationModal());
    }
    
    const locationSearch = document.getElementById('location-search');
    if (locationSearch) {
        locationSearch.addEventListener('input', filterLocations);
    }
    
    // 進出貨管理
    const stockInForm = document.getElementById('quick-stock-in-form');
    if (stockInForm) {
        stockInForm.addEventListener('submit', handleStockIn);
    }
    
    const stockOutForm = document.getElementById('quick-stock-out-form');
    if (stockOutForm) {
        stockOutForm.addEventListener('submit', handleStockOut);
    }
    
    // 添加延遲設置事件監聽器，確保DOM元素已經存在
    setTimeout(() => {
        const stockOutProduct = document.getElementById('stock-out-product');
        const stockOutLocation = document.getElementById('stock-out-location');
        if (stockOutProduct) {
            stockOutProduct.addEventListener('change', updateStockAvailability);
        }
        if (stockOutLocation) {
            stockOutLocation.addEventListener('change', updateStockAvailability);
        }
    }, 100);
    
    // 庫存查詢
    const applyInventoryFilterBtn = document.getElementById('apply-inventory-filter');
    if (applyInventoryFilterBtn) {
        applyInventoryFilterBtn.addEventListener('click', applyInventoryFilter);
    }
    
    const clearInventoryFilterBtn = document.getElementById('clear-inventory-filter');
    if (clearInventoryFilterBtn) {
        clearInventoryFilterBtn.addEventListener('click', clearInventoryFilter);
    }
    
    const exportInventoryBtn = document.getElementById('export-inventory-btn');
    if (exportInventoryBtn) {
        exportInventoryBtn.addEventListener('click', exportInventory);
    }
    
    const printInventoryBtn = document.getElementById('print-inventory-btn');
    if (printInventoryBtn) {
        printInventoryBtn.addEventListener('click', printInventory);
    }
    
    // 異動記錄
    const applyRecordFilterBtn = document.getElementById('apply-record-filter');
    if (applyRecordFilterBtn) {
        applyRecordFilterBtn.addEventListener('click', applyRecordFilter);
    }
    
    const clearRecordFilterBtn = document.getElementById('clear-record-filter');
    if (clearRecordFilterBtn) {
        clearRecordFilterBtn.addEventListener('click', clearRecordFilter);
    }
    
    const exportRecordsBtn = document.getElementById('export-records-btn');
    if (exportRecordsBtn) {
        exportRecordsBtn.addEventListener('click', exportRecords);
    }
    
    // 設定
    const backupDataBtn = document.getElementById('backup-data-btn');
    if (backupDataBtn) {
        backupDataBtn.addEventListener('click', backupData);
    }
    
    const restoreDataBtn = document.getElementById('restore-data-btn');
    if (restoreDataBtn) {
        restoreDataBtn.addEventListener('click', restoreData);
    }
    
    const clearDataBtn = document.getElementById('clear-data-btn');
    if (clearDataBtn) {
        clearDataBtn.addEventListener('click', clearAllData);
    }
    
    const lowStockThresholdInput = document.getElementById('low-stock-threshold');
    if (lowStockThresholdInput) {
        lowStockThresholdInput.addEventListener('change', updateLowStockThreshold);
    }
    
    // 模態對話框
    const modalClose = document.getElementById('modal-close');
    if (modalClose) {
        modalClose.addEventListener('click', hideModal);
    }
    
    const modalCancel = document.getElementById('modal-cancel');
    if (modalCancel) {
        modalCancel.addEventListener('click', hideModal);
    }
    
    const notificationClose = document.getElementById('notification-close');
    if (notificationClose) {
        notificationClose.addEventListener('click', hideNotification);
    }
    
    // 點擊模態背景關閉
    const modal = document.getElementById('modal');
    if (modal) {
        modal.addEventListener('click', (e) => {
            if (e.target.id === 'modal') hideModal();
        });
    }
}

async function syncIndividualData(dataType) {
    if (!gasUrl) {
        showNotification('請先設定 GAS URL', 'error');
        return;
    }
    
    if (!isOnline) {
        showNotification('請檢查網路連線', 'error');
        return;
    }
    
    const statusElement = document.getElementById('gas-status');
    const originalText = statusElement.textContent;
    statusElement.textContent = '同步中...';
    statusElement.className = 'gas-status syncing';
    
    try {
        let data;
        switch (dataType) {
            case 'products':
                data = products;
                break;
            case 'locations':
                data = locations;
                break;
            case 'inventory':
                data = inventory;
                break;
            case 'records':
                data = records;
                break;
        }
        
        const success = await syncDataToGas(dataType, data);
        
        if (success) {
            showNotification(`${getDataTypeName(dataType)}同步成功`);
        } else {
            showNotification(`${getDataTypeName(dataType)}同步失敗`, 'error');
        }
    } catch (error) {
        showNotification('同步過程發生錯誤：' + error.message, 'error');
    } finally {
        statusElement.textContent = originalText;
        updateGasStatus();
    }
}

function getDataTypeName(dataType) {
    const names = {
        'products': '商品資料',
        'locations': '地點資料',
        'inventory': '庫存資料',
        'records': '異動記錄'
    };
    return names[dataType] || dataType;
}

// 頁面導航
function showPage(pageName) {
    // 隱藏所有頁面
    const pages = document.querySelectorAll('.page');
    pages.forEach(page => {
        page.classList.remove('active');
    });
    
    // 移除所有導航項目的活動狀態
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => {
        item.classList.remove('active');
    });
    
    // 顯示選定頁面
    const targetPage = document.getElementById(pageName);
    if (targetPage) {
        targetPage.classList.add('active');
    }
    
    // 設定對應導航項目為活動狀態
    const targetNavItem = document.querySelector(`[data-page="${pageName}"]`);
    if (targetNavItem) {
        targetNavItem.classList.add('active');
    }
    
    currentPage = pageName;
    
    // 更新頁面資料和選單選項
    if (pageName === 'dashboard') {
        updateDashboard();
    } else if (pageName === 'stock-in' || pageName === 'stock-out') {
        // 確保選單選項已更新
        updateSelectOptions();
    }
}

// 更新儀表板
function updateDashboard() {
    // 統計資料
    const totalProductsEl = document.getElementById('total-products');
    if (totalProductsEl) {
        totalProductsEl.textContent = products.length;
    }
    
    const totalLocationsEl = document.getElementById('total-locations');
    if (totalLocationsEl) {
        totalLocationsEl.textContent = locations.length;
    }
    
    const totalInventoryEl = document.getElementById('total-inventory');
    if (totalInventoryEl) {
        totalInventoryEl.textContent = inventory.reduce((sum, item) => sum + item.quantity, 0);
    }
    
    // 今日異動
    const today = new Date().toDateString();
    const todayRecords = records.filter(record => new Date(record.timestamp).toDateString() === today);
    const todayTransactionsEl = document.getElementById('today-transactions');
    if (todayTransactionsEl) {
        todayTransactionsEl.textContent = todayRecords.length;
    }
    
    // 庫存不足警告
    updateLowStockAlerts();
    
    // 最近異動記錄
    updateRecentRecords();
}

// 庫存不足警告
function updateLowStockAlerts() {
    const alertsContainer = document.getElementById('low-stock-alerts');
    if (!alertsContainer) return;
    
    const lowStockItems = inventory.filter(item => item.quantity <= lowStockThreshold);
    
    if (lowStockItems.length === 0) {
        alertsContainer.innerHTML = '<p>目前沒有庫存不足的商品</p>';
        return;
    }
    
    const alertsHtml = lowStockItems.map(item => {
        const product = products.find(p => p.id === item.productId);
        const location = locations.find(l => l.id === item.locationId);
        return `
            <div class="alert-item">
                <strong>${product?.name || '未知商品'}</strong> 在 
                <strong>${location?.name || '未知地點'}</strong> 
                庫存不足 (剩餘: ${item.quantity} ${product?.unit || ''})
            </div>
        `;
    }).join('');
    
    alertsContainer.innerHTML = alertsHtml;
}

// 最近異動記錄
function updateRecentRecords() {
    const recentContainer = document.getElementById('recent-records');
    if (!recentContainer) return;
    
    const recentRecords = records.slice(-5).reverse();
    
    if (recentRecords.length === 0) {
        recentContainer.innerHTML = '<p>暫無異動記錄</p>';
        return;
    }
    
    const recordsHtml = recentRecords.map(record => {
        const product = products.find(p => p.id === record.productId);
        const location = locations.find(l => l.id === record.locationId);
        const date = new Date(record.timestamp).toLocaleString('zh-TW');
        
        return `
            <div class="record-item">
                <div class="record-info">
                    <div>
                        <span class="record-type ${record.type}">${record.type === 'in' ? '進貨' : '出貨'}</span>
                        <strong>${product?.name || '未知商品'}</strong>
                    </div>
                    <small>${location?.name || '未知地點'} • ${record.quantity} ${product?.unit || ''} • ${date}</small>
                </div>
            </div>
        `;
    }).join('');
    
    recentContainer.innerHTML = recordsHtml;
}

// 商品管理
function renderProducts() {
    const tbody = document.querySelector('#products-table tbody');
    if (!tbody) return;
    
    if (products.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align: center;">暫無商品資料</td></tr>';
        return;
    }
    
    const html = products.map(product => `
        <tr>
            <td>${product.id}</td>
            <td>${product.name}</td>
            <td>${product.description}</td>
            <td>${product.unit}</td>
            <td>
                <div class="table-actions">
                    <button class="btn-edit" onclick="editProduct('${product.id}')">編輯</button>
                    <button class="btn-delete" onclick="deleteProduct('${product.id}')">刪除</button>
                </div>
            </td>
        </tr>
    `).join('');
    
    tbody.innerHTML = html;
}

function filterProducts() {
    const searchInput = document.getElementById('product-search');
    if (!searchInput) return;
    
    const searchTerm = searchInput.value.toLowerCase();
    const filteredProducts = products.filter(product => 
        product.name.toLowerCase().includes(searchTerm) || 
        product.id.toLowerCase().includes(searchTerm)
    );
    
    const tbody = document.querySelector('#products-table tbody');
    if (!tbody) return;
    
    if (filteredProducts.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align: center;">找不到符合條件的商品</td></tr>';
        return;
    }
    
    const html = filteredProducts.map(product => `
        <tr>
            <td>${product.id}</td>
            <td>${product.name}</td>
            <td>${product.description}</td>
            <td>${product.unit}</td>
            <td>
                <div class="table-actions">
                    <button class="btn-edit" onclick="editProduct('${product.id}')">編輯</button>
                    <button class="btn-delete" onclick="deleteProduct('${product.id}')">刪除</button>
                </div>
            </td>
        </tr>
    `).join('');
    
    tbody.innerHTML = html;
}

function showProductModal(productId = null) {
    const isEdit = productId !== null;
    const product = isEdit ? products.find(p => p.id === productId) : null;
    
    showModal(isEdit ? '編輯商品' : '新增商品', `
        <form id="product-form">
            <div class="form-group">
                <label class="form-label">商品編號</label>
                <input type="text" class="form-control" id="product-id" value="${product?.id || ''}" ${isEdit ? 'readonly' : ''} required>
            </div>
            <div class="form-group">
                <label class="form-label">商品名稱</label>
                <input type="text" class="form-control" id="product-name" value="${product?.name || ''}" required>
            </div>
            <div class="form-group">
                <label class="form-label">商品描述</label>
                <textarea class="form-control" id="product-description" rows="3">${product?.description || ''}</textarea>
            </div>
            <div class="form-group">
                <label class="form-label">單位</label>
                <input type="text" class="form-control" id="product-unit" value="${product?.unit || ''}" required>
            </div>
        </form>
    `, () => saveProduct(isEdit));
}

async function saveProduct(isEdit) {
    const id = document.getElementById('product-id').value.trim();
    const name = document.getElementById('product-name').value.trim();
    const description = document.getElementById('product-description').value.trim();
    const unit = document.getElementById('product-unit').value.trim();
    
    if (!id || !name || !unit) {
        showNotification('請填寫所有必填欄位', 'error');
        return;
    }
    
    if (!isEdit && products.some(p => p.id === id)) {
        showNotification('商品編號已存在', 'error');
        return;
    }
    
    if (isEdit) {
        const productIndex = products.findIndex(p => p.id === id);
        products[productIndex] = { id, name, description, unit };
    } else {
        products.push({ id, name, description, unit });
        // 為新商品在所有地點初始化庫存
        locations.forEach(location => {
            inventory.push({
                productId: id,
                locationId: location.id,
                quantity: 0,
                lastUpdated: new Date().toISOString()
            });
        });
    }
    
    saveDataToStorage();
    renderProducts();
    renderInventory();
    // 重要：立即更新選單選項
    updateSelectOptions();
    hideModal();
    showNotification(isEdit ? '商品已更新' : '商品已新增');
    
    // 自動同步到 GAS (如果啟用)
    if (autoSyncEnabled && gasUrl && isOnline) {
        await syncDataToGas('products', products);
        if (!isEdit) {
            await syncDataToGas('inventory', inventory);
        }
    }
}

function editProduct(productId) {
    showProductModal(productId);
}

async function deleteProduct(productId) {
    const product = products.find(p => p.id === productId);
    if (!product) return;
    
    if (confirm(`確定要刪除商品「${product.name}」嗎？這將同時刪除所有相關的庫存和異動記錄。`)) {
        products = products.filter(p => p.id !== productId);
        inventory = inventory.filter(i => i.productId !== productId);
        records = records.filter(r => r.productId !== productId);
        
        saveDataToStorage();
        renderProducts();
        updateSelectOptions();
        renderInventory();
        updateDashboard();
        showNotification('商品已刪除');
        
        // 自動同步到 GAS (如果啟用)
        if (autoSyncEnabled && gasUrl && isOnline) {
            await syncAllDataToGas();
        }
    }
}

// 地點管理
function renderLocations() {
    const tbody = document.querySelector('#locations-table tbody');
    if (!tbody) return;
    
    if (locations.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align: center;">暫無地點資料</td></tr>';
        return;
    }
    
    const html = locations.map(location => `
        <tr>
            <td>${location.id}</td>
            <td>${location.name}</td>
            <td>${location.address}</td>
            <td>${location.description}</td>
            <td>
                <div class="table-actions">
                    <button class="btn-edit" onclick="editLocation('${location.id}')">編輯</button>
                    <button class="btn-delete" onclick="deleteLocation('${location.id}')">刪除</button>
                </div>
            </td>
        </tr>
    `).join('');
    
    tbody.innerHTML = html;
}

function filterLocations() {
    const searchInput = document.getElementById('location-search');
    if (!searchInput) return;
    
    const searchTerm = searchInput.value.toLowerCase();
    const filteredLocations = locations.filter(location => 
        location.name.toLowerCase().includes(searchTerm) || 
        location.address.toLowerCase().includes(searchTerm)
    );
    
    const tbody = document.querySelector('#locations-table tbody');
    if (!tbody) return;
    
    if (filteredLocations.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align: center;">找不到符合條件的地點</td></tr>';
        return;
    }
    
    const html = filteredLocations.map(location => `
        <tr>
            <td>${location.id}</td>
            <td>${location.name}</td>
            <td>${location.address}</td>
            <td>${location.description}</td>
            <td>
                <div class="table-actions">
                    <button class="btn-edit" onclick="editLocation('${location.id}')">編輯</button>
                    <button class="btn-delete" onclick="deleteLocation('${location.id}')">刪除</button>
                </div>
            </td>
        </tr>
    `).join('');
    
    tbody.innerHTML = html;
}

function showLocationModal(locationId = null) {
    const isEdit = locationId !== null;
    const location = isEdit ? locations.find(l => l.id === locationId) : null;
    
    showModal(isEdit ? '編輯地點' : '新增地點', `
        <form id="location-form">
            <div class="form-group">
                <label class="form-label">地點編號</label>
                <input type="text" class="form-control" id="location-id" value="${location?.id || ''}" ${isEdit ? 'readonly' : ''} required>
            </div>
            <div class="form-group">
                <label class="form-label">地點名稱</label>
                <input type="text" class="form-control" id="location-name" value="${location?.name || ''}" required>
            </div>
            <div class="form-group">
                <label class="form-label">地址</label>
                <input type="text" class="form-control" id="location-address" value="${location?.address || ''}" required>
            </div>
            <div class="form-group">
                <label class="form-label">描述</label>
                <textarea class="form-control" id="location-description" rows="3">${location?.description || ''}</textarea>
            </div>
        </form>
    `, () => saveLocation(isEdit));
}

async function saveLocation(isEdit) {
    const id = document.getElementById('location-id').value.trim();
    const name = document.getElementById('location-name').value.trim();
    const address = document.getElementById('location-address').value.trim();
    const description = document.getElementById('location-description').value.trim();
    
    if (!id || !name || !address) {
        showNotification('請填寫所有必填欄位', 'error');
        return;
    }
    
    if (!isEdit && locations.some(l => l.id === id)) {
        showNotification('地點編號已存在', 'error');
        return;
    }
    
    if (isEdit) {
        const locationIndex = locations.findIndex(l => l.id === id);
        locations[locationIndex] = { id, name, address, description };
    } else {
        locations.push({ id, name, address, description });
        // 為新地點在所有商品初始化庫存
        products.forEach(product => {
            inventory.push({
                productId: product.id,
                locationId: id,
                quantity: 0,
                lastUpdated: new Date().toISOString()
            });
        });
    }
    
    saveDataToStorage();
    renderLocations();
    renderInventory();
    // 重要：立即更新選單選項
    updateSelectOptions();
    hideModal();
    showNotification(isEdit ? '地點已更新' : '地點已新增');
    
    // 自動同步到 GAS (如果啟用)
    if (autoSyncEnabled && gasUrl && isOnline) {
        await syncDataToGas('locations', locations);
        if (!isEdit) {
            await syncDataToGas('inventory', inventory);
        }
    }
}

function editLocation(locationId) {
    showLocationModal(locationId);
}

async function deleteLocation(locationId) {
    const location = locations.find(l => l.id === locationId);
    if (!location) return;
    
    if (confirm(`確定要刪除地點「${location.name}」嗎？這將同時刪除所有相關的庫存和異動記錄。`)) {
        locations = locations.filter(l => l.id !== locationId);
        inventory = inventory.filter(i => i.locationId !== locationId);
        records = records.filter(r => r.locationId !== locationId);
        
        saveDataToStorage();
        renderLocations();
        updateSelectOptions();
        renderInventory();
        updateDashboard();
        showNotification('地點已刪除');
        
        // 自動同步到 GAS (如果啟用)
        if (autoSyncEnabled && gasUrl && isOnline) {
            await syncAllDataToGas();
        }
    }
}

// 更新下拉選單選項 - 修正選項保留問題
function updateSelectOptions() {
    // 地點選項
    const locationOptions = locations.map(location => 
        `<option value="${location.id}">${location.name}</option>`
    ).join('');
    
    // 商品選項
    const productOptions = products.map(product => 
        `<option value="${product.id}">${product.name}</option>`
    ).join('');
    
    // 更新進出貨選單
    const selectorsToUpdate = [
        'stock-in-location',
        'stock-in-product', 
        'stock-out-location',
        'stock-out-product',
        'inventory-location-filter',
        'inventory-product-filter'
    ];
    
    selectorsToUpdate.forEach(selectorId => {
        const element = document.getElementById(selectorId);
        if (element) {
            // 不保存當前值，讓使用者重新選擇
            if (selectorId.includes('location')) {
                if (selectorId.includes('filter')) {
                    element.innerHTML = '<option value="">所有地點</option>' + locationOptions;
                } else {
                    element.innerHTML = '<option value="">請選擇...</option>' + locationOptions;
                }
            } else if (selectorId.includes('product')) {
                if (selectorId.includes('filter')) {
                    element.innerHTML = '<option value="">所有商品</option>' + productOptions;
                } else {
                    element.innerHTML = '<option value="">請選擇...</option>' + productOptions;
                }
            }
        }
    });
}

// 進貨管理
async function handleStockIn(e) {
    e.preventDefault();
    
    const locationId = document.getElementById('stock-in-location').value;
    const productId = document.getElementById('stock-in-product').value;
    const quantity = parseInt(document.getElementById('stock-in-quantity').value);
    const note = document.getElementById('stock-in-note').value.trim();
    const operator = document.getElementById('stock-in-operator').value.trim() || '系統';
    
    if (!locationId || !productId || !quantity || quantity <= 0) {
        showNotification('請填寫所有必填欄位', 'error');
        return;
    }
    
    // 更新庫存
    const inventoryItem = inventory.find(i => i.productId === productId && i.locationId === locationId);
    if (inventoryItem) {
        inventoryItem.quantity += quantity;
        inventoryItem.lastUpdated = new Date().toISOString();
    } else {
        inventory.push({ 
            productId, 
            locationId, 
            quantity,
            lastUpdated: new Date().toISOString()
        });
    }
    
    // 記錄異動
    const record = {
        id: generateId(),
        type: 'in',
        productId,
        locationId,
        quantity,
        timestamp: new Date().toISOString(),
        operator,
        note
    };
    records.push(record);
    
    saveDataToStorage();
    updateDashboard();
    renderInventory();
    renderRecords();
    
    // 清空表單
    document.getElementById('quick-stock-in-form').reset();
    showNotification('進貨記錄已新增');
    
    // 自動同步到 GAS (如果啟用)
    if (autoSyncEnabled && gasUrl && isOnline) {
        await syncDataToGas('inventory', inventory);
        await syncDataToGas('records', records);
    }
}

// 出貨管理
async function handleStockOut(e) {
    e.preventDefault();
    
    const locationId = document.getElementById('stock-out-location').value;
    const productId = document.getElementById('stock-out-product').value;
    const quantity = parseInt(document.getElementById('stock-out-quantity').value);
    const note = document.getElementById('stock-out-note').value.trim();
    const operator = document.getElementById('stock-out-operator').value.trim() || '系統';
    
    if (!locationId || !productId || !quantity || quantity <= 0) {
        showNotification('請填寫所有必填欄位', 'error');
        return;
    }
    
    // 檢查庫存是否足夠
    const inventoryItem = inventory.find(i => i.productId === productId && i.locationId === locationId);
    const currentStock = inventoryItem ? inventoryItem.quantity : 0;
    
    if (currentStock < quantity) {
        showNotification(`庫存不足！目前庫存：${currentStock}`, 'error');
        return;
    }
    
    // 更新庫存
    inventoryItem.quantity -= quantity;
    inventoryItem.lastUpdated = new Date().toISOString();
    
    // 記錄異動
    const record = {
        id: generateId(),
        type: 'out',
        productId,
        locationId,
        quantity,
        timestamp: new Date().toISOString(),
        operator,
        note
    };
    records.push(record);
    
    saveDataToStorage();
    updateDashboard();
    renderInventory();
    renderRecords();
    
    // 清空表單和庫存資訊
    document.getElementById('quick-stock-out-form').reset();
    const stockAvailability = document.getElementById('stock-availability');
    if (stockAvailability) {
        stockAvailability.innerHTML = '';
    }
    showNotification('出貨記錄已新增');
    
    // 自動同步到 GAS (如果啟用)
    if (autoSyncEnabled && gasUrl && isOnline) {
        await syncDataToGas('inventory', inventory);
        await syncDataToGas('records', records);
    }
}

function updateStockAvailability() {
    const locationId = document.getElementById('stock-out-location')?.value;
    const productId = document.getElementById('stock-out-product')?.value;
    const availabilityDiv = document.getElementById('stock-availability');
    
    if (!availabilityDiv || !locationId || !productId) {
        if (availabilityDiv) availabilityDiv.innerHTML = '';
        return;
    }
    
    const product = products.find(p => p.id === productId);
    const location = locations.find(l => l.id === locationId);
    const inventoryItem = inventory.find(i => i.productId === productId && i.locationId === locationId);
    const currentStock = inventoryItem ? inventoryItem.quantity : 0;
    
    let statusClass = 'stock-info';
    if (currentStock <= 0) {
        statusClass += ' error';
    } else if (currentStock <= lowStockThreshold) {
        statusClass += ' warning';
    }
    
    availabilityDiv.className = statusClass;
    availabilityDiv.innerHTML = `
        <strong>${location?.name}</strong> 的 <strong>${product?.name}</strong> 
        現有庫存：${currentStock} ${product?.unit}
        ${currentStock <= lowStockThreshold ? ' (庫存不足)' : ''}
    `;
}

// 庫存查詢
function renderInventory() {
    const tbody = document.querySelector('#inventory-table tbody');
    if (!tbody) return;
    
    if (inventory.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align: center;">暫無庫存資料</td></tr>';
        return;
    }
    
    const html = inventory.map(item => {
        const product = products.find(p => p.id === item.productId);
        const location = locations.find(l => l.id === item.locationId);
        
        let status = 'normal';
        let statusText = '正常';
        
        if (item.quantity === 0) {
            status = 'empty';
            statusText = '缺貨';
        } else if (item.quantity <= lowStockThreshold) {
            status = 'low';
            statusText = '庫存不足';
        }
        
        const lastUpdated = item.lastUpdated ? 
            new Date(item.lastUpdated).toLocaleString('zh-TW') : 
            '未知';
        
        return `
            <tr>
                <td>${location?.name || '未知地點'}</td>
                <td>${product?.name || '未知商品'}</td>
                <td>${product?.id || '未知'}</td>
                <td>${item.quantity}</td>
                <td>${product?.unit || ''}</td>
                <td><span class="status-tag ${status}">${statusText}</span></td>
                <td>${lastUpdated}</td>
            </tr>
        `;
    }).join('');
    
    tbody.innerHTML = html;
}

function applyInventoryFilter() {
    const locationFilter = document.getElementById('inventory-location-filter');
    const productFilter = document.getElementById('inventory-product-filter');
    
    if (!locationFilter || !productFilter) return;
    
    const locationValue = locationFilter.value;
    const productValue = productFilter.value;
    
    let filteredInventory = inventory;
    
    if (locationValue) {
        filteredInventory = filteredInventory.filter(item => item.locationId === locationValue);
    }
    
    if (productValue) {
        filteredInventory = filteredInventory.filter(item => item.productId === productValue);
    }
    
    const tbody = document.querySelector('#inventory-table tbody');
    if (!tbody) return;
    
    if (filteredInventory.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align: center;">找不到符合條件的庫存</td></tr>';
        return;
    }
    
    const html = filteredInventory.map(item => {
        const product = products.find(p => p.id === item.productId);
        const location = locations.find(l => l.id === item.locationId);
        
        let status = 'normal';
        let statusText = '正常';
        
        if (item.quantity === 0) {
            status = 'empty';
            statusText = '缺貨';
        } else if (item.quantity <= lowStockThreshold) {
            status = 'low';
            statusText = '庫存不足';
        }
        
        const lastUpdated = item.lastUpdated ? 
            new Date(item.lastUpdated).toLocaleString('zh-TW') : 
            '未知';
        
        return `
            <tr>
                <td>${location?.name || '未知地點'}</td>
                <td>${product?.name || '未知商品'}</td>
                <td>${product?.id || '未知'}</td>
                <td>${item.quantity}</td>
                <td>${product?.unit || ''}</td>
                <td><span class="status-tag ${status}">${statusText}</span></td>
                <td>${lastUpdated}</td>
            </tr>
        `;
    }).join('');
    
    tbody.innerHTML = html;
}

function clearInventoryFilter() {
    const locationFilter = document.getElementById('inventory-location-filter');
    const productFilter = document.getElementById('inventory-product-filter');
    
    if (locationFilter) locationFilter.value = '';
    if (productFilter) productFilter.value = '';
    renderInventory();
}

// 異動記錄
function renderRecords() {
    const tbody = document.querySelector('#records-table tbody');
    if (!tbody) return;
    
    if (records.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align: center;">暫無異動記錄</td></tr>';
        return;
    }
    
    const sortedRecords = records.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    const html = sortedRecords.map(record => {
        const product = products.find(p => p.id === record.productId);
        const location = locations.find(l => l.id === record.locationId);
        const date = new Date(record.timestamp).toLocaleString('zh-TW');
        
        return `
            <tr>
                <td>${date}</td>
                <td><span class="record-type ${record.type}">${record.type === 'in' ? '進貨' : '出貨'}</span></td>
                <td>${location?.name || '未知地點'}</td>
                <td>${product?.name || '未知商品'}</td>
                <td>${record.quantity} ${product?.unit || ''}</td>
                <td>${record.operator}</td>
                <td>${record.note || ''}</td>
            </tr>
        `;
    }).join('');
    
    tbody.innerHTML = html;
}

function applyRecordFilter() {
    const typeFilter = document.getElementById('record-type-filter');
    const dateFrom = document.getElementById('record-date-from');
    const dateTo = document.getElementById('record-date-to');
    
    if (!typeFilter || !dateFrom || !dateTo) return;
    
    const typeValue = typeFilter.value;
    const dateFromValue = dateFrom.value;
    const dateToValue = dateTo.value;
    
    let filteredRecords = records;
    
    if (typeValue) {
        filteredRecords = filteredRecords.filter(record => record.type === typeValue);
    }
    
    if (dateFromValue) {
        const fromDate = new Date(dateFromValue);
        filteredRecords = filteredRecords.filter(record => new Date(record.timestamp) >= fromDate);
    }
    
    if (dateToValue) {
        const toDate = new Date(dateToValue);
        toDate.setHours(23, 59, 59, 999);
        filteredRecords = filteredRecords.filter(record => new Date(record.timestamp) <= toDate);
    }
    
    const tbody = document.querySelector('#records-table tbody');
    if (!tbody) return;
    
    if (filteredRecords.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align: center;">找不到符合條件的記錄</td></tr>';
        return;
    }
    
    const sortedRecords = filteredRecords.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    const html = sortedRecords.map(record => {
        const product = products.find(p => p.id === record.productId);
        const location = locations.find(l => l.id === record.locationId);
        const date = new Date(record.timestamp).toLocaleString('zh-TW');
        
        return `
            <tr>
                <td>${date}</td>
                <td><span class="record-type ${record.type}">${record.type === 'in' ? '進貨' : '出貨'}</span></td>
                <td>${location?.name || '未知地點'}</td>
                <td>${product?.name || '未知商品'}</td>
                <td>${record.quantity} ${product?.unit || ''}</td>
                <td>${record.operator}</td>
                <td>${record.note || ''}</td>
            </tr>
        `;
    }).join('');
    
    tbody.innerHTML = html;
}

function clearRecordFilter() {
    const typeFilter = document.getElementById('record-type-filter');
    const dateFrom = document.getElementById('record-date-from');
    const dateTo = document.getElementById('record-date-to');
    
    if (typeFilter) typeFilter.value = '';
    if (dateFrom) dateFrom.value = '';
    if (dateTo) dateTo.value = '';
    renderRecords();
}

// 設定功能
function updateLowStockThreshold() {
    const thresholdInput = document.getElementById('low-stock-threshold');
    if (!thresholdInput) return;
    
    const threshold = parseInt(thresholdInput.value);
    if (threshold >= 0) {
        lowStockThreshold = threshold;
        saveDataToStorage();
        updateDashboard();
        renderInventory();
        showNotification('庫存警告設定已更新');
    }
}

function backupData() {
    const data = {
        products,
        locations,
        inventory,
        records,
        lowStockThreshold,
        gasUrl,
        autoSyncEnabled,
        syncInterval,
        timestamp: new Date().toISOString()
    };
    
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `inventory-backup-${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
    
    showNotification('資料備份已下載');
}

function restoreData() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.onchange = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = JSON.parse(e.target.result);
                
                if (data.products && data.locations && data.inventory && data.records) {
                    products = data.products;
                    locations = data.locations;
                    inventory = data.inventory;
                    records = data.records;
                    lowStockThreshold = data.lowStockThreshold || 5;
                    
                    // 還原 GAS 設定
                    if (data.gasUrl) gasUrl = data.gasUrl;
                    if (data.autoSyncEnabled !== undefined) autoSyncEnabled = data.autoSyncEnabled;
                    if (data.syncInterval) syncInterval = data.syncInterval;
                    
                    saveDataToStorage();
                    saveGasSettings();
                    initializeApp();
                    showNotification('資料還原成功');
                } else {
                    showNotification('無效的備份檔案格式', 'error');
                }
            } catch (error) {
                showNotification('備份檔案解析失敗', 'error');
            }
        };
        reader.readAsText(file);
    };
    input.click();
}

function clearAllData() {
    if (confirm('確定要清除所有資料嗎？此操作無法復原！')) {
        products = [];
        locations = [];
        inventory = [];
        records = [];
        lowStockThreshold = 5;
        gasUrl = '';
        autoSyncEnabled = false;
        syncInterval = 5;
        lastSyncTime = null;
        
        // 清除 localStorage
        const keysToKeep = [];  // 保留某些設定
        Object.keys(localStorage).forEach(key => {
            if (!keysToKeep.includes(key)) {
                localStorage.removeItem(key);
            }
        });
        
        stopAutoSync();
        initializeApp();
        showNotification('所有資料已清除');
    }
}

// 匯出功能
function exportInventory() {
    const data = inventory.map(item => {
        const product = products.find(p => p.id === item.productId);
        const location = locations.find(l => l.id === item.locationId);
        
        let status = '正常';
        if (item.quantity === 0) status = '缺貨';
        else if (item.quantity <= lowStockThreshold) status = '庫存不足';
        
        return {
            '地點': location?.name || '未知地點',
            '商品名稱': product?.name || '未知商品',
            '商品編號': product?.id || '未知',
            '現有庫存': item.quantity,
            '單位': product?.unit || '',
            '狀態': status,
            '最後更新': item.lastUpdated ? new Date(item.lastUpdated).toLocaleString('zh-TW') : '未知'
        };
    });
    
    exportToCSV(data, '庫存清單');
}

function exportRecords() {
    const data = records.map(record => {
        const product = products.find(p => p.id === record.productId);
        const location = locations.find(l => l.id === record.locationId);
        
        return {
            '時間': new Date(record.timestamp).toLocaleString('zh-TW'),
            '類型': record.type === 'in' ? '進貨' : '出貨',
            '地點': location?.name || '未知地點',
            '商品': product?.name || '未知商品',
            '數量': record.quantity,
            '單位': product?.unit || '',
            '操作員': record.operator,
            '備註': record.note || ''
        };
    });
    
    exportToCSV(data, '異動記錄');
}

function exportToCSV(data, filename) {
    if (data.length === 0) {
        showNotification('沒有資料可匯出', 'warning');
        return;
    }
    
    const headers = Object.keys(data[0]);
    const csvContent = [
        headers.join(','),
        ...data.map(row => headers.map(header => `"${row[header]}"`).join(','))
    ].join('\n');
    
    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${filename}-${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    
    showNotification(`${filename}已匯出`);
}

function printInventory() {
    window.print();
}

// 工具函數
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

function showModal(title, body, onConfirm = null) {
    const modal = document.getElementById('modal');
    const modalTitle = document.getElementById('modal-title');
    const modalBody = document.getElementById('modal-body');
    const confirmBtn = document.getElementById('modal-confirm');
    
    if (!modal || !modalTitle || !modalBody || !confirmBtn) return;
    
    modalTitle.textContent = title;
    modalBody.innerHTML = body;
    modal.classList.remove('hidden');
    
    // 重新綁定確認按鈕事件
    confirmBtn.onclick = onConfirm;
    
    if (!onConfirm) {
        confirmBtn.style.display = 'none';
    } else {
        confirmBtn.style.display = 'inline-flex';
    }
}

function hideModal() {
    const modal = document.getElementById('modal');
    if (modal) {
        modal.classList.add('hidden');
    }
}

function showNotification(message, type = 'success') {
    const notification = document.getElementById('notification');
    const messageElement = document.getElementById('notification-message');
    
    if (!notification || !messageElement) return;
    
    messageElement.textContent = message;
    notification.className = `notification ${type}`;
    notification.classList.remove('hidden');
    
    setTimeout(() => {
        hideNotification();
    }, 5000);
}

function hideNotification() {
    const notification = document.getElementById('notification');
    if (notification) {
        notification.classList.add('hidden');
    }
}

function showLoading(message = '載入中...') {
    const loading = document.getElementById('loading');
    const loadingText = document.querySelector('.loading-text');
    
    if (loading) {
        loading.classList.remove('hidden');
    }
    if (loadingText) {
        loadingText.textContent = message;
    }
}

function hideLoading() {
    const loading = document.getElementById('loading');
    if (loading) {
        loading.classList.add('hidden');
    }
}

// 鍵盤快捷鍵支援
document.addEventListener('keydown', (e) => {
    // Ctrl/Cmd + 數字鍵切換頁面
    if ((e.ctrlKey || e.metaKey) && e.key >= '1' && e.key <= '8') {
        e.preventDefault();
        const pages = ['dashboard', 'products', 'locations', 'stock-in', 'stock-out', 'inventory', 'records', 'settings'];
        const pageIndex = parseInt(e.key) - 1;
        if (pages[pageIndex]) {
            showPage(pages[pageIndex]);
        }
    }
    
    // ESC 關閉模態對話框
    if (e.key === 'Escape') {
        hideModal();
        hideNotification();
        hideLoading();
    }
});

// 頁面關閉前停止自動同步
window.addEventListener('beforeunload', () => {
    stopAutoSync();
});