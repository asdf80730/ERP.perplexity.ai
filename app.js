// 全域變數
let currentPage = 'home';
let products = [];
let locations = [];
let inventory = [];
let records = [];
let lowStockThreshold = 5;

// GAS 整合變數
let gasUrl = '';
let autoSyncEnabled = false;
let isOnline = navigator.onLine;
let lastSyncTime = null;

// 應用程式初始化
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
    setupEventListeners();
    setupNetworkListeners();
    
    // 載入完成後顯示應用程式
    setTimeout(() => {
        document.body.style.opacity = '1';
    }, 100);
});

function initializeApp() {
    console.log('初始化應用程式...');
    loadDataFromStorage();
    loadGasSettings();
    initializeSampleData();
    updateConnectionStatus();
    updateDashboard();
    updateSelectOptions();
    renderProducts();
    renderInventory();
    renderRecords();
    showPage('home');
    console.log('應用程式初始化完成');
}

// 載入本地資料
function loadDataFromStorage() {
    try {
        const storedData = {
            products: localStorage.getItem('mobile_products'),
            locations: localStorage.getItem('mobile_locations'),
            inventory: localStorage.getItem('mobile_inventory'),
            records: localStorage.getItem('mobile_records'),
            lowStockThreshold: localStorage.getItem('mobile_lowStockThreshold')
        };

        products = storedData.products ? JSON.parse(storedData.products) : [];
        locations = storedData.locations ? JSON.parse(storedData.locations) : [];
        inventory = storedData.inventory ? JSON.parse(storedData.inventory) : [];
        records = storedData.records ? JSON.parse(storedData.records) : [];
        lowStockThreshold = storedData.lowStockThreshold ? parseInt(storedData.lowStockThreshold) : 5;
        
        console.log('本地資料載入完成:', { products: products.length, locations: locations.length, inventory: inventory.length, records: records.length });
    } catch (error) {
        console.error('載入本地資料失敗:', error);
        showNotification('載入本地資料失敗', 'error');
    }
}

// 儲存本地資料
function saveDataToStorage() {
    try {
        localStorage.setItem('mobile_products', JSON.stringify(products));
        localStorage.setItem('mobile_locations', JSON.stringify(locations));
        localStorage.setItem('mobile_inventory', JSON.stringify(inventory));
        localStorage.setItem('mobile_records', JSON.stringify(records));
        localStorage.setItem('mobile_lowStockThreshold', lowStockThreshold.toString());
        console.log('本地資料儲存完成');
    } catch (error) {
        console.error('儲存本地資料失敗:', error);
        showNotification('儲存本地資料失敗', 'error');
    }
}

// 載入 GAS 設定
function loadGasSettings() {
    gasUrl = localStorage.getItem('mobile_gasUrl') || '';
    autoSyncEnabled = localStorage.getItem('mobile_autoSyncEnabled') === 'true';
    lastSyncTime = localStorage.getItem('mobile_lastSyncTime');
    
    setTimeout(() => {
        const gasUrlInput = document.getElementById('gas-url');
        const autoSyncCheckbox = document.getElementById('auto-sync-enabled');
        
        if (gasUrlInput) gasUrlInput.value = gasUrl;
        if (autoSyncCheckbox) autoSyncCheckbox.checked = autoSyncEnabled;
        
        updateSyncInfo();
    }, 100);
}

// 儲存 GAS 設定
function saveGasSettingsToStorage() {
    localStorage.setItem('mobile_gasUrl', gasUrl);
    localStorage.setItem('mobile_autoSyncEnabled', autoSyncEnabled.toString());
    if (lastSyncTime) {
        localStorage.setItem('mobile_lastSyncTime', lastSyncTime);
    }
}

// 初始化範例資料（使用提供的資料）
function initializeSampleData() {
    if (products.length === 0 && locations.length === 0) {
        console.log('初始化範例資料...');
        // 使用提供的應用程式資料
        locations = [
            {id: "loc1", name: "主倉庫", address: "台北市信義區", description: "主要存放地點"},
            {id: "loc2", name: "門市A", address: "台北市大安區", description: "零售門市"},
            {id: "loc3", name: "門市B", address: "台中市西區", description: "分店庫存"}
        ];
        
        products = [
            {id: "prod1", name: "商品A", unit: "個", description: "熱銷商品", category: "電子產品"},
            {id: "prod2", name: "商品B", unit: "盒", description: "季節商品", category: "服飾配件"},
            {id: "prod3", name: "商品C", unit: "件", description: "基本款", category: "日用品"}
        ];
        
        inventory = [
            {productId: "prod1", locationId: "loc1", quantity: 150, lastUpdated: "2025-01-28"},
            {productId: "prod1", locationId: "loc2", quantity: 25, lastUpdated: "2025-01-28"},
            {productId: "prod2", locationId: "loc1", quantity: 8, lastUpdated: "2025-01-27"},
            {productId: "prod2", locationId: "loc3", quantity: 45, lastUpdated: "2025-01-28"},
            {productId: "prod3", locationId: "loc1", quantity: 2, lastUpdated: "2025-01-26"}
        ];
        
        records = [
            {id: generateId(), type: "進貨", productId: "prod1", locationId: "loc1", quantity: 50, operator: "管理員", timestamp: "2025-01-28 14:30"},
            {id: generateId(), type: "出貨", productId: "prod2", locationId: "loc2", quantity: 12, operator: "員工A", timestamp: "2025-01-28 11:15"},
            {id: generateId(), type: "進貨", productId: "prod3", locationId: "loc3", quantity: 30, operator: "管理員", timestamp: "2025-01-27 16:45"}
        ];
        
        saveDataToStorage();
        console.log('範例資料初始化完成');
    }
}

// 設定事件監聽器
function setupEventListeners() {
    console.log('設定事件監聽器...');
    
    // 底部導航
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => {
        item.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            const page = this.getAttribute('data-page');
            console.log('點擊導航項目:', page);
            if (page) {
                showPage(page);
                // 觸覺回饋
                if (navigator.vibrate) {
                    navigator.vibrate(50);
                }
            }
        });
    });

    // 篩選標籤
    document.addEventListener('click', function(e) {
        if (e.target.classList.contains('filter-tab')) {
            const tabs = e.target.parentElement.querySelectorAll('.filter-tab');
            tabs.forEach(tab => tab.classList.remove('active'));
            e.target.classList.add('active');
            
            const filter = e.target.getAttribute('data-filter');
            if (currentPage === 'inventory') {
                applyInventoryFilter(filter);
            } else if (currentPage === 'records') {
                applyRecordsFilter(filter);
            }
        }
    });

    // 模態框背景點擊關閉
    document.addEventListener('click', function(e) {
        if (e.target.classList.contains('modal')) {
            hideModal(e.target.id);
        }
    });

    // ESC 鍵關閉模態框
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            const modals = document.querySelectorAll('.modal:not(.hidden)');
            modals.forEach(modal => hideModal(modal.id));
        }
    });

    // 設定頁面的庫存警告閾值
    setTimeout(() => {
        const thresholdInput = document.getElementById('low-stock-threshold');
        if (thresholdInput) {
            thresholdInput.value = lowStockThreshold;
        }
    }, 100);
    
    console.log('事件監聽器設定完成');
}

// 網路狀態監聽
function setupNetworkListeners() {
    window.addEventListener('online', () => {
        isOnline = true;
        updateConnectionStatus();
        showNotification('網路連線已恢復');
    });
    
    window.addEventListener('offline', () => {
        isOnline = false;
        updateConnectionStatus();
        showNotification('網路連線中斷，切換到離線模式', 'warning');
    });
}

// 更新連接狀態
function updateConnectionStatus() {
    const statusDot = document.getElementById('status-dot');
    const statusText = document.getElementById('status-text');
    
    if (!statusDot || !statusText) return;
    
    if (!isOnline) {
        statusDot.className = 'status-dot';
        statusText.textContent = '離線模式';
    } else if (gasUrl) {
        statusDot.className = 'status-dot online';
        statusText.textContent = '已連接雲端';
    } else {
        statusDot.className = 'status-dot';
        statusText.textContent = '未設定雲端';
    }
}

// 頁面導航
function showPage(pageName) {
    console.log('切換到頁面:', pageName);
    
    // 隱藏所有頁面
    const pages = document.querySelectorAll('.page');
    pages.forEach(page => page.classList.remove('active'));
    
    // 移除所有導航項目的活動狀態
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => item.classList.remove('active'));
    
    // 顯示選定頁面
    const targetPage = document.getElementById(pageName);
    if (targetPage) {
        targetPage.classList.add('active');
        console.log('頁面顯示成功:', pageName);
    } else {
        console.error('找不到頁面:', pageName);
        return;
    }
    
    // 設定對應導航項目為活動狀態
    const targetNavItem = document.querySelector(`[data-page="${pageName}"]`);
    if (targetNavItem) {
        targetNavItem.classList.add('active');
    }
    
    currentPage = pageName;
    
    // 頁面特定的更新
    if (pageName === 'home') {
        updateDashboard();
    } else if (pageName === 'inventory') {
        renderInventory();
    } else if (pageName === 'records') {
        renderRecords();
    } else if (pageName === 'products') {
        renderProducts();
    }
}

// 更新儀表板
function updateDashboard() {
    console.log('更新儀表板...');
    
    // 統計資料
    const totalProductsEl = document.getElementById('total-products');
    const totalLocationsEl = document.getElementById('total-locations');
    const totalInventoryEl = document.getElementById('total-inventory');
    const todayTransactionsEl = document.getElementById('today-transactions');
    
    if (totalProductsEl) totalProductsEl.textContent = products.length;
    if (totalLocationsEl) totalLocationsEl.textContent = locations.length;
    if (totalInventoryEl) {
        const total = inventory.reduce((sum, item) => sum + item.quantity, 0);
        totalInventoryEl.textContent = total;
    }
    
    // 今日異動
    const today = new Date().toDateString();
    const todayRecords = records.filter(record => {
        const recordDate = new Date(record.timestamp).toDateString();
        return recordDate === today;
    });
    if (todayTransactionsEl) todayTransactionsEl.textContent = todayRecords.length;
    
    // 庫存警告
    updateLowStockAlerts();
    
    // 最近異動
    updateRecentRecords();
}

// 庫存警告
function updateLowStockAlerts() {
    const alertsContainer = document.getElementById('low-stock-alerts');
    if (!alertsContainer) return;
    
    const lowStockItems = inventory.filter(item => item.quantity <= lowStockThreshold);
    
    if (lowStockItems.length === 0) {
        alertsContainer.innerHTML = '<div class="no-data">目前沒有庫存不足的商品</div>';
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
        recentContainer.innerHTML = '<div class="no-data">暫無異動記錄</div>';
        return;
    }
    
    const recordsHtml = recentRecords.map(record => {
        const product = products.find(p => p.id === record.productId);
        const location = locations.find(l => l.id === record.locationId);
        const date = new Date(record.timestamp).toLocaleString('zh-TW');
        
        return `
            <div class="record-item">
                <div class="record-info">
                    <div class="record-title">
                        <span class="record-type ${record.type === '進貨' ? 'in' : 'out'}">${record.type}</span>
                        ${product?.name || '未知商品'}
                    </div>
                    <div class="record-details">
                        ${location?.name || '未知地點'} • ${record.quantity} ${product?.unit || ''} • ${date}
                    </div>
                </div>
            </div>
        `;
    }).join('');
    
    recentContainer.innerHTML = recordsHtml;
}

// 快速操作函數 - 修復全域函數問題
window.showQuickStockIn = function() {
    console.log('顯示快速進貨模態框');
    updateSelectOptions();
    showModal('stock-in-modal');
};

window.showQuickStockOut = function() {
    console.log('顯示快速出貨模態框');
    updateSelectOptions();
    showModal('stock-out-modal');
};

window.showProductModal = function(productId = null) {
    console.log('顯示商品模態框:', productId);
    const isEdit = productId !== null;
    const product = isEdit ? products.find(p => p.id === productId) : null;
    
    const modal = document.getElementById('product-modal');
    const title = document.getElementById('product-modal-title');
    const form = document.getElementById('product-form');
    
    if (!modal || !title || !form) return;
    
    title.textContent = isEdit ? '編輯商品' : '新增商品';
    
    // 填入表單資料
    document.getElementById('product-id').value = product?.id || '';
    document.getElementById('product-name').value = product?.name || '';
    document.getElementById('product-unit').value = product?.unit || '';
    document.getElementById('product-category').value = product?.category || '';
    document.getElementById('product-description').value = product?.description || '';
    
    // 設定編輯模式
    document.getElementById('product-id').readOnly = isEdit;
    
    // 儲存編輯狀態到表單
    form.dataset.isEdit = isEdit;
    form.dataset.productId = productId || '';
    
    showModal('product-modal');
};

window.editProduct = function(productId) {
    showProductModal(productId);
};

window.deleteProduct = function(productId) {
    const product = products.find(p => p.id === productId);
    if (!product) return;
    
    if (confirm(`確定要刪除商品「${product.name}」嗎？這將同時刪除所有相關的庫存和異動記錄。`)) {
        products = products.filter(p => p.id !== productId);
        inventory = inventory.filter(i => i.productId !== productId);
        records = records.filter(r => r.productId !== productId);
        
        saveDataToStorage();
        renderProducts();
        updateSelectOptions();
        updateDashboard();
        showNotification('商品已刪除');
        
        // 觸覺回饋
        if (navigator.vibrate) {
            navigator.vibrate(200);
        }
        
        // 自動同步
        if (autoSyncEnabled && gasUrl && isOnline) {
            syncToGas();
        }
    }
};

window.refreshInventory = function() {
    console.log('刷新庫存');
    // 觸覺回饋
    if (navigator.vibrate) {
        navigator.vibrate(50);
    }
    
    renderInventory();
    showNotification('庫存資料已刷新');
};

window.refreshRecords = function() {
    console.log('刷新記錄');
    // 觸覺回饋
    if (navigator.vibrate) {
        navigator.vibrate(50);
    }
    
    renderRecords();
    showNotification('記錄資料已刷新');
};

// 商品管理
function renderProducts() {
    const container = document.getElementById('products-list');
    if (!container) return;
    
    if (products.length === 0) {
        container.innerHTML = '<div class="no-data">暫無商品資料<br><button class="btn btn--primary" onclick="showProductModal()">新增第一個商品</button></div>';
        return;
    }
    
    const html = products.map(product => `
        <div class="list-item">
            <div class="item-info">
                <div class="item-title">${product.name}</div>
                <div class="item-subtitle">
                    編號: ${product.id} • 單位: ${product.unit}
                    ${product.category ? ` • ${product.category}` : ''}
                </div>
            </div>
            <div class="item-actions">
                <button class="btn btn--small btn--outline" onclick="editProduct('${product.id}')">編輯</button>
                <button class="btn btn--small btn--secondary" onclick="deleteProduct('${product.id}')">刪除</button>
            </div>
        </div>
    `).join('');
    
    container.innerHTML = html;
}

window.filterProducts = function() {
    const searchInput = document.getElementById('product-search');
    const searchTerm = searchInput ? searchInput.value.toLowerCase() : '';
    
    const filteredProducts = products.filter(product => 
        product.name.toLowerCase().includes(searchTerm) || 
        product.id.toLowerCase().includes(searchTerm) ||
        (product.category && product.category.toLowerCase().includes(searchTerm))
    );
    
    const container = document.getElementById('products-list');
    if (!container) return;
    
    if (filteredProducts.length === 0) {
        container.innerHTML = '<div class="no-data">找不到符合条件的商品</div>';
        return;
    }
    
    const html = filteredProducts.map(product => `
        <div class="list-item">
            <div class="item-info">
                <div class="item-title">${product.name}</div>
                <div class="item-subtitle">
                    編號: ${product.id} • 單位: ${product.unit}
                    ${product.category ? ` • ${product.category}` : ''}
                </div>
            </div>
            <div class="item-actions">
                <button class="btn btn--small btn--outline" onclick="editProduct('${product.id}')">編輯</button>
                <button class="btn btn--small btn--secondary" onclick="deleteProduct('${product.id}')">刪除</button>
            </div>
        </div>
    `).join('');
    
    container.innerHTML = html;
};

window.saveProduct = function(event) {
    event.preventDefault();
    
    const form = document.getElementById('product-form');
    const isEdit = form.dataset.isEdit === 'true';
    const originalId = form.dataset.productId;
    
    const id = document.getElementById('product-id').value.trim();
    const name = document.getElementById('product-name').value.trim();
    const unit = document.getElementById('product-unit').value.trim();
    const category = document.getElementById('product-category').value.trim();
    const description = document.getElementById('product-description').value.trim();
    
    if (!id || !name || !unit) {
        showNotification('請填寫所有必填欄位', 'error');
        return;
    }
    
    if (!isEdit && products.some(p => p.id === id)) {
        showNotification('商品編號已存在', 'error');
        return;
    }
    
    if (isEdit) {
        const productIndex = products.findIndex(p => p.id === originalId);
        if (productIndex !== -1) {
            products[productIndex] = { id, name, unit, category, description };
        }
    } else {
        products.push({ id, name, unit, category, description });
        // 為新商品在所有地點初始化庫存
        locations.forEach(location => {
            inventory.push({
                productId: id,
                locationId: location.id,
                quantity: 0,
                lastUpdated: new Date().toISOString().split('T')[0]
            });
        });
    }
    
    saveDataToStorage();
    renderProducts();
    updateSelectOptions();
    updateDashboard();
    hideModal('product-modal');
    showNotification(isEdit ? '商品已更新' : '商品已新增');
    
    // 觸覺回饋
    if (navigator.vibrate) {
        navigator.vibrate(100);
    }
    
    // 自動同步
    if (autoSyncEnabled && gasUrl && isOnline) {
        syncToGas();
    }
};

// 庫存管理
function renderInventory() {
    const container = document.getElementById('inventory-list');
    if (!container) return;
    
    if (inventory.length === 0) {
        container.innerHTML = '<div class="no-data">暫無庫存資料</div>';
        return;
    }
    
    const html = inventory.map(item => {
        const product = products.find(p => p.id === item.productId);
        const location = locations.find(l => l.id === item.locationId);
        
        let statusClass = 'normal';
        let statusText = '正常';
        
        if (item.quantity === 0) {
            statusClass = 'empty';
            statusText = '缺貨';
        } else if (item.quantity <= lowStockThreshold) {
            statusClass = 'low';
            statusText = '庫存不足';
        }
        
        return `
            <div class="list-item">
                <div class="item-info">
                    <div class="item-title">${product?.name || '未知商品'}</div>
                    <div class="item-subtitle">
                        ${location?.name || '未知地點'} • ${item.quantity} ${product?.unit || ''} • ${statusText}
                    </div>
                </div>
                <div class="item-badge ${statusClass}">${item.quantity}</div>
            </div>
        `;
    }).join('');
    
    container.innerHTML = html;
}

window.filterInventory = function() {
    const searchInput = document.getElementById('inventory-search');
    const searchTerm = searchInput ? searchInput.value.toLowerCase() : '';
    
    let filteredInventory = inventory.filter(item => {
        const product = products.find(p => p.id === item.productId);
        const location = locations.find(l => l.id === item.locationId);
        
        return (product?.name.toLowerCase().includes(searchTerm)) ||
               (location?.name.toLowerCase().includes(searchTerm)) ||
               (product?.id.toLowerCase().includes(searchTerm));
    });
    
    renderFilteredInventory(filteredInventory);
};

function applyInventoryFilter(filter) {
    let filteredInventory = inventory;
    
    if (filter === 'low') {
        filteredInventory = inventory.filter(item => item.quantity > 0 && item.quantity <= lowStockThreshold);
    } else if (filter === 'empty') {
        filteredInventory = inventory.filter(item => item.quantity === 0);
    }
    
    renderFilteredInventory(filteredInventory);
}

function renderFilteredInventory(filteredInventory) {
    const container = document.getElementById('inventory-list');
    if (!container) return;
    
    if (filteredInventory.length === 0) {
        container.innerHTML = '<div class="no-data">找不到符合條件的庫存</div>';
        return;
    }
    
    const html = filteredInventory.map(item => {
        const product = products.find(p => p.id === item.productId);
        const location = locations.find(l => l.id === item.locationId);
        
        let statusClass = 'normal';
        let statusText = '正常';
        
        if (item.quantity === 0) {
            statusClass = 'empty';
            statusText = '缺貨';
        } else if (item.quantity <= lowStockThreshold) {
            statusClass = 'low';
            statusText = '庫存不足';
        }
        
        return `
            <div class="list-item">
                <div class="item-info">
                    <div class="item-title">${product?.name || '未知商品'}</div>
                    <div class="item-subtitle">
                        ${location?.name || '未知地點'} • ${item.quantity} ${product?.unit || ''} • ${statusText}
                    </div>
                </div>
                <div class="item-badge ${statusClass}">${item.quantity}</div>
            </div>
        `;
    }).join('');
    
    container.innerHTML = html;
}

// 記錄管理
function renderRecords() {
    const container = document.getElementById('records-list');
    if (!container) return;
    
    if (records.length === 0) {
        container.innerHTML = '<div class="no-data">暫無異動記錄</div>';
        return;
    }
    
    const sortedRecords = [...records].sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    const html = sortedRecords.map(record => {
        const product = products.find(p => p.id === record.productId);
        const location = locations.find(l => l.id === record.locationId);
        const date = new Date(record.timestamp).toLocaleString('zh-TW');
        
        return `
            <div class="list-item">
                <div class="item-info">
                    <div class="item-title">
                        <span class="record-type ${record.type === '進貨' ? 'in' : 'out'}">${record.type}</span>
                        ${product?.name || '未知商品'}
                    </div>
                    <div class="item-subtitle">
                        ${location?.name || '未知地點'} • ${record.quantity} ${product?.unit || ''}<br>
                        ${record.operator} • ${date}
                    </div>
                </div>
            </div>
        `;
    }).join('');
    
    container.innerHTML = html;
}

function applyRecordsFilter(filter) {
    let filteredRecords = records;
    
    if (filter === 'in') {
        filteredRecords = records.filter(record => record.type === '進貨');
    } else if (filter === 'out') {
        filteredRecords = records.filter(record => record.type === '出貨');
    }
    
    renderFilteredRecords(filteredRecords);
}

function renderFilteredRecords(filteredRecords) {
    const container = document.getElementById('records-list');
    if (!container) return;
    
    if (filteredRecords.length === 0) {
        container.innerHTML = '<div class="no-data">找不到符合條件的記錄</div>';
        return;
    }
    
    const sortedRecords = [...filteredRecords].sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    const html = sortedRecords.map(record => {
        const product = products.find(p => p.id === record.productId);
        const location = locations.find(l => l.id === record.locationId);
        const date = new Date(record.timestamp).toLocaleString('zh-TW');
        
        return `
            <div class="list-item">
                <div class="item-info">
                    <div class="item-title">
                        <span class="record-type ${record.type === '進貨' ? 'in' : 'out'}">${record.type}</span>
                        ${product?.name || '未知商品'}
                    </div>
                    <div class="item-subtitle">
                        ${location?.name || '未知地點'} • ${record.quantity} ${product?.unit || ''}<br>
                        ${record.operator} • ${date}
                    </div>
                </div>
            </div>
        `;
    }).join('');
    
    container.innerHTML = html;
}

function updateSelectOptions() {
    console.log('更新選單選項...');
    
    // 更新地點選項
    const locationSelects = ['stock-in-location', 'stock-out-location'];
    locationSelects.forEach(selectId => {
        const select = document.getElementById(selectId);
        if (select) {
            const currentValue = select.value;
            select.innerHTML = '<option value="">請選擇地點</option>' + 
                locations.map(location => 
                    `<option value="${location.id}" ${currentValue === location.id ? 'selected' : ''}>${location.name}</option>`
                ).join('');
        }
    });
    
    // 更新商品選項
    const productSelects = ['stock-in-product', 'stock-out-product'];
    productSelects.forEach(selectId => {
        const select = document.getElementById(selectId);
        if (select) {
            const currentValue = select.value;
            select.innerHTML = '<option value="">請選擇商品</option>' + 
                products.map(product => 
                    `<option value="${product.id}" ${currentValue === product.id ? 'selected' : ''}>${product.name}</option>`
                ).join('');
        }
    });
}

window.adjustQuantity = function(inputId, delta) {
    const input = document.getElementById(inputId);
    if (!input) return;
    
    const currentValue = parseInt(input.value) || 1;
    const newValue = Math.max(1, currentValue + delta);
    input.value = newValue;
    
    // 觸覺回饋
    if (navigator.vibrate) {
        navigator.vibrate(30);
    }
};

window.updateStockAvailability = function() {
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
    
    let className = 'stock-info';
    if (currentStock <= 0) {
        className += ' error';
    } else if (currentStock <= lowStockThreshold) {
        className += ' warning';
    }
    
    availabilityDiv.className = className;
    availabilityDiv.innerHTML = `
        <strong>${location?.name}</strong> 的 <strong>${product?.name}</strong> 
        現有庫存：${currentStock} ${product?.unit}
        ${currentStock <= lowStockThreshold ? ' (庫存不足)' : ''}
    `;
};

window.handleStockIn = function(event) {
    event.preventDefault();
    console.log('處理進貨...');
    
    const locationId = document.getElementById('stock-in-location').value;
    const productId = document.getElementById('stock-in-product').value;
    const quantity = parseInt(document.getElementById('stock-in-quantity').value);
    const operator = document.getElementById('stock-in-operator').value.trim() || '系統';
    const note = document.getElementById('stock-in-note').value.trim();
    
    if (!locationId || !productId || !quantity || quantity <= 0) {
        showNotification('請填寫所有必填欄位', 'error');
        return;
    }
    
    // 更新庫存
    const inventoryItem = inventory.find(i => i.productId === productId && i.locationId === locationId);
    if (inventoryItem) {
        inventoryItem.quantity += quantity;
        inventoryItem.lastUpdated = new Date().toISOString().split('T')[0];
    } else {
        inventory.push({
            productId,
            locationId,
            quantity,
            lastUpdated: new Date().toISOString().split('T')[0]
        });
    }
    
    // 記錄異動
    const record = {
        id: generateId(),
        type: '進貨',
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
    document.getElementById('stock-in-form').reset();
    hideModal('stock-in-modal');
    showNotification('進貨記錄已新增');
    
    // 觸覺回饋
    if (navigator.vibrate) {
        navigator.vibrate([100, 50, 100]);
    }
    
    // 自動同步
    if (autoSyncEnabled && gasUrl && isOnline) {
        syncToGas();
    }
};

window.handleStockOut = function(event) {
    event.preventDefault();
    console.log('處理出貨...');
    
    const locationId = document.getElementById('stock-out-location').value;
    const productId = document.getElementById('stock-out-product').value;
    const quantity = parseInt(document.getElementById('stock-out-quantity').value);
    const operator = document.getElementById('stock-out-operator').value.trim() || '系統';
    const note = document.getElementById('stock-out-note').value.trim();
    
    if (!locationId || !productId || !quantity || quantity <= 0) {
        showNotification('請填寫所有必填欄位', 'error');
        return;
    }
    
    // 檢查庫存是否足夠
    const inventoryItem = inventory.find(i => i.productId === productId && i.locationId === locationId);
    const currentStock = inventoryItem ? inventoryItem.quantity : 0;
    
    if (currentStock < quantity) {
        showNotification(`庫存不足！目前庫存：${currentStock}`, 'error');
        if (navigator.vibrate) {
            navigator.vibrate([200, 100, 200]);
        }
        return;
    }
    
    // 更新庫存
    inventoryItem.quantity -= quantity;
    inventoryItem.lastUpdated = new Date().toISOString().split('T')[0];
    
    // 記錄異動
    const record = {
        id: generateId(),
        type: '出貨',
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
    document.getElementById('stock-out-form').reset();
    const stockAvailability = document.getElementById('stock-availability');
    if (stockAvailability) {
        stockAvailability.innerHTML = '';
    }
    hideModal('stock-out-modal');
    showNotification('出貨記錄已新增');
    
    // 觸覺回饋
    if (navigator.vibrate) {
        navigator.vibrate([100, 50, 100]);
    }
    
    // 自動同步
    if (autoSyncEnabled && gasUrl && isOnline) {
        syncToGas();
    }
};

// Google Apps Script 整合
window.testGasConnection = async function() {
    const gasUrlInput = document.getElementById('gas-url');
    const statusDiv = document.getElementById('gas-connection-status');
    
    if (!gasUrlInput || !statusDiv) return;
    
    const url = gasUrlInput.value.trim();
    if (!url) {
        statusDiv.innerHTML = '<div class="connection-result error">請輸入 GAS URL</div>';
        return;
    }
    
    if (!isOnline) {
        statusDiv.innerHTML = '<div class="connection-result error">請檢查網路連線</div>';
        return;
    }
    
    statusDiv.innerHTML = '<div class="connection-result testing">正在測試連接...</div>';
    
    try {
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
                throw new Error('GAS回應格式錯誤');
            }
            
            if (data.success) {
                statusDiv.innerHTML = '<div class="connection-result success">✓ 連接成功！</div>';
            } else {
                statusDiv.innerHTML = '<div class="connection-result error">✗ GAS 回應錯誤：' + (data.error || '未知錯誤') + '</div>';
            }
        } else {
            statusDiv.innerHTML = '<div class="connection-result error">✗ 連接失敗：HTTP ' + response.status + '</div>';
        }
    } catch (error) {
        statusDiv.innerHTML = '<div class="connection-result error">✗ 連接失敗：' + error.message + '</div>';
    }
};

window.saveGasSettings = function() {
    const gasUrlInput = document.getElementById('gas-url');
    const autoSyncCheckbox = document.getElementById('auto-sync-enabled');
    
    if (!gasUrlInput || !autoSyncCheckbox) return;
    
    gasUrl = gasUrlInput.value.trim();
    autoSyncEnabled = autoSyncCheckbox.checked;
    
    saveGasSettingsToStorage();
    updateConnectionStatus();
    updateSyncInfo();
    
    showNotification('GAS 設定已儲存');
    
    // 觸覺回饋
    if (navigator.vibrate) {
        navigator.vibrate(100);
    }
};

async function syncToGas() {
    if (!gasUrl || !isOnline) return false;
    
    const statusDot = document.getElementById('status-dot');
    if (statusDot) {
        statusDot.className = 'status-dot syncing';
    }
    
    try {
        const allData = {
            products,
            locations,
            inventory,
            records,
            lastSync: new Date().toISOString()
        };
        
        const formData = new FormData();
        formData.append('action', 'sync');
        formData.append('dataType', 'all');
        formData.append('data', JSON.stringify(allData));
        
        const response = await fetch(gasUrl, {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            const text = await response.text();
            const result = JSON.parse(text);
            
            if (result.success) {
                lastSyncTime = new Date().toISOString();
                localStorage.setItem('mobile_lastSyncTime', lastSyncTime);
                updateConnectionStatus();
                updateSyncInfo();
                return true;
            }
        }
    } catch (error) {
        console.error('同步錯誤:', error);
    }
    
    updateConnectionStatus();
    return false;
}

window.syncAllData = async function() {
    if (!gasUrl) {
        showNotification('請先設定 GAS URL', 'error');
        return;
    }
    
    if (!isOnline) {
        showNotification('請檢查網路連線', 'error');
        return;
    }
    
    showLoading('正在同步資料...');
    
    try {
        const success = await syncToGas();
        
        if (success) {
            showNotification('資料同步成功');
            // 觸覺回饋
            if (navigator.vibrate) {
                navigator.vibrate([100, 50, 100, 50, 100]);
            }
        } else {
            showNotification('資料同步失敗', 'error');
        }
    } catch (error) {
        showNotification('同步過程發生錯誤', 'error');
    } finally {
        hideLoading();
    }
};

window.pullFromGas = async function() {
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
                updateDashboard();
                renderProducts();
                renderInventory();
                renderRecords();
                updateSelectOptions();
                showNotification('資料拉取成功');
                
                // 觸覺回饋
                if (navigator.vibrate) {
                    navigator.vibrate([100, 50, 100, 50, 100]);
                }
            } else {
                showNotification('資料拉取失敗', 'error');
            }
        } else {
            showNotification('連接失敗', 'error');
        }
    } catch (error) {
        showNotification('拉取過程發生錯誤', 'error');
    } finally {
        hideLoading();
    }
};

function updateSyncInfo() {
    const syncInfoElement = document.getElementById('last-sync-time');
    if (!syncInfoElement) return;
    
    if (lastSyncTime) {
        const syncDate = new Date(lastSyncTime);
        syncInfoElement.textContent = `上次同步：${syncDate.toLocaleString('zh-TW')}`;
    } else {
        syncInfoElement.textContent = '尚未同步';
    }
}

// 設定功能
window.updateLowStockThreshold = function() {
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
};

window.backupData = function() {
    const data = {
        products,
        locations,
        inventory,
        records,
        lowStockThreshold,
        gasUrl,
        autoSyncEnabled,
        timestamp: new Date().toISOString()
    };
    
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `mobile-inventory-backup-${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
    
    showNotification('資料備份已下載');
};

window.restoreData = function() {
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
                    
                    if (data.gasUrl) gasUrl = data.gasUrl;
                    if (data.autoSyncEnabled !== undefined) autoSyncEnabled = data.autoSyncEnabled;
                    
                    saveDataToStorage();
                    saveGasSettingsToStorage();
                    initializeApp();
                    showNotification('資料還原成功');
                    
                    // 觸覺回饋
                    if (navigator.vibrate) {
                        navigator.vibrate([200, 100, 200]);
                    }
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
};

window.clearAllData = function() {
    if (confirm('確定要清除所有資料嗎？此操作無法復原！')) {
        products = [];
        locations = [];
        inventory = [];
        records = [];
        lowStockThreshold = 5;
        gasUrl = '';
        autoSyncEnabled = false;
        lastSyncTime = null;
        
        // 清除 localStorage
        const keys = Object.keys(localStorage);
        keys.forEach(key => {
            if (key.startsWith('mobile_')) {
                localStorage.removeItem(key);
            }
        });
        
        initializeApp();
        showNotification('所有資料已清除');
        
        // 觸覺回饋
        if (navigator.vibrate) {
            navigator.vibrate([300, 200, 300]);
        }
    }
};

// 工具函數
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

function showModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
        modal.classList.remove('hidden');
        console.log('模態框已顯示:', modalId);
        // 觸覺回饋
        if (navigator.vibrate) {
            navigator.vibrate(50);
        }
    } else {
        console.error('找不到模態框:', modalId);
    }
}

window.hideModal = function(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
        modal.classList.add('hidden');
        console.log('模態框已隱藏:', modalId);
    }
};

function showNotification(message, type = 'success') {
    const notification = document.getElementById('notification');
    const messageElement = document.getElementById('notification-message');
    
    if (!notification || !messageElement) return;
    
    messageElement.textContent = message;
    notification.className = `notification ${type}`;
    notification.classList.remove('hidden');
    
    console.log('顯示通知:', message, type);
    
    setTimeout(() => {
        hideNotification();
    }, 4000);
}

window.hideNotification = function() {
    const notification = document.getElementById('notification');
    if (notification) {
        notification.classList.add('hidden');
    }
};

function showLoading(message = '處理中...') {
    const loading = document.getElementById('loading');
    const loadingText = document.querySelector('.loading-text');
    
    if (loading) {
        loading.classList.remove('hidden');
    }
    if (loadingText) {
        loadingText.textContent = message;
    }
    
    console.log('顯示載入:', message);
}

function hideLoading() {
    const loading = document.getElementById('loading');
    if (loading) {
        loading.classList.add('hidden');
    }
    
    console.log('隱藏載入');
}

// 防止意外離開頁面時丟失未儲存的資料
window.addEventListener('beforeunload', function(e) {
    // 如果有未儲存的資料變更，可以在這裡提醒用戶
    // e.preventDefault();
    // e.returnValue = '';
});

// 頁面可見性變化時同步資料
document.addEventListener('visibilitychange', function() {
    if (!document.hidden && autoSyncEnabled && gasUrl && isOnline) {
        syncToGas();
    }
});

// 支援手勢操作（可選功能）
let touchStartX = 0;
let touchStartY = 0;

document.addEventListener('touchstart', function(e) {
    touchStartX = e.touches[0].clientX;
    touchStartY = e.touches[0].clientY;
}, false);

document.addEventListener('touchend', function(e) {
    if (!touchStartX || !touchStartY) return;
    
    const touchEndX = e.changedTouches[0].clientX;
    const touchEndY = e.changedTouches[0].clientY;
    
    const diffX = touchStartX - touchEndX;
    const diffY = touchStartY - touchEndY;
    
    // 只在主要內容區域檢測手勢
    if (e.target.closest('.modal') || e.target.closest('.bottom-nav')) {
        return;
    }
    
    // 水平滑動手勢
    if (Math.abs(diffX) > Math.abs(diffY) && Math.abs(diffX) > 100) {
        const pages = ['home', 'products', 'inventory', 'records', 'settings'];
        const currentIndex = pages.indexOf(currentPage);
        
        if (diffX > 0 && currentIndex < pages.length - 1) {
            // 向左滑動，顯示下一頁
            showPage(pages[currentIndex + 1]);
        } else if (diffX < 0 && currentIndex > 0) {
            // 向右滑動，顯示上一頁
            showPage(pages[currentIndex - 1]);
        }
    }
    
    touchStartX = 0;
    touchStartY = 0;
}, false);