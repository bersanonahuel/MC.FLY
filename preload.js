const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    // --- PRODUCTOS ---
    getAllData: () => ipcRenderer.invoke('db-get-all'),
    saveProduct: (p) => ipcRenderer.invoke('db-save-product', p),
    deleteProduct: (code) => ipcRenderer.invoke('db-delete-product', code),
    bulkUpdate: (data) => ipcRenderer.invoke('db-bulk-update', data),
    bulkUpdateSelection: (data) => ipcRenderer.invoke('db-bulk-update-selection', data),

    // Búsqueda rápida (Minimercado)
    searchProduct: (query) => ipcRenderer.invoke('db-search-product', query),

    // --- VENTAS Y CLIENTES ---
    saveSale: (s) => ipcRenderer.invoke('db-save-sale', s),
    saveClient: (c) => ipcRenderer.invoke('db-save-client', c),
    payDebt: (data) => ipcRenderer.invoke('db-pay-debt', data),
    deleteClient: (id) => ipcRenderer.invoke('db-delete-client', id), // <--- Ya la tenías, perfecto.

    // --- CAJA Y GASTOS (PROVEEDORES) ---
    saveClosure: (c) => ipcRenderer.invoke('db-save-closure', c),

    // --- TURNOS (FONDO DE CAJA) ---
    getActiveShift: () => ipcRenderer.invoke('db-get-active-shift'),
    openShift: (data) => ipcRenderer.invoke('db-open-shift', data),
    closeShift: (data) => ipcRenderer.invoke('db-close-shift', data),
    updateShiftFondo: (data) => ipcRenderer.invoke('db-update-shift-fondo', data),
    getLastClosure: () => ipcRenderer.invoke('db-get-last-closure'),

    // 🔥 ESTA ES LA NUEVA (Para pagar a proveedores):
    saveExpense: (ex) => ipcRenderer.invoke('db-save-expense', ex),

    exportarExcel: (d) => ipcRenderer.invoke('exportar-excel', d),
    focusWindow: () => ipcRenderer.invoke('focus-window'),
    generateWordLabels: (products) => ipcRenderer.invoke('generate-word-labels', products),
    addUser: (u) => ipcRenderer.invoke('db-add-user', u),
    logAudit: (data) => ipcRenderer.invoke('db-log-audit', data),

    // --- MANTENIMIENTO ---
    factoryReset: () => ipcRenderer.invoke('db-factory-reset'),
    restoreBackup: (data) => ipcRenderer.invoke('db-restore-backup', data),

    // --- EXTRAS (Fiscal e Impresión) ---
    facturar: (v) => ipcRenderer.invoke('facturar-afip', v),
    imprimirTicket: (html) => ipcRenderer.invoke('print-ticket', html),
});