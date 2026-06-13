// ─── AFIP TLS Fix: Debe ir ANTES de cualquier require ─────────────────────
// Los servidores de AFIP usan claves DH de 1024 bits que Node.js v17+ rechaza.
// Bajamos el nivel de seguridad TLS para permitir la conexión.
const tls = require('tls');
const _origCreateSecureContext = tls.createSecureContext;
tls.createSecureContext = (opts = {}) => {
    opts.ciphers = opts.ciphers || 'DEFAULT@SECLEVEL=0';
    if (!opts.ciphers.includes('SECLEVEL')) opts.ciphers += ':DEFAULT@SECLEVEL=0';
    return _origCreateSecureContext(opts);
};
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';
// ─────────────────────────────────────────────────────────────────────────────

const { app, BrowserWindow, ipcMain, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const sqlite3 = require('sqlite3').verbose();
const xlsx = require('xlsx');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle } = require('docx');

let dbPath;
let db;

app.disableHardwareAcceleration();

function initDB() {
    dbPath = path.join(app.getPath('userData'), 'MCFLY.db');
    console.log("Ruta de la Base de Datos:", dbPath);
    db = new sqlite3.Database(dbPath);

    db.serialize(() => {
        // 1. Crear tablas base si no existen
        db.run(`CREATE TABLE IF NOT EXISTS products (code TEXT PRIMARY KEY, name TEXT, supplier TEXT, brand TEXT, category TEXT, subcategory TEXT, size TEXT, color TEXT, cost REAL, margin REAL, price REAL, qty INTEGER)`);
        db.run(`CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, pin TEXT)`);
        db.run(`CREATE TABLE IF NOT EXISTS closures (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, user TEXT, sys_cash REAL, real_cash REAL, diff REAL, status TEXT, details_json TEXT)`);
        db.run(`CREATE TABLE IF NOT EXISTS sales (id INTEGER PRIMARY KEY, date TEXT, seller TEXT, total REAL, payment TEXT, card_name TEXT, surcharge REAL, discount REAL, client_id INTEGER, items_json TEXT, cuit TEXT, fiscal TEXT, observations TEXT)`);
        db.run(`CREATE TABLE IF NOT EXISTS clients (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, phone TEXT, address TEXT, balance REAL DEFAULT 0, dni TEXT, credit_limit REAL DEFAULT 0)`);
        db.run(`CREATE TABLE IF NOT EXISTS client_movements (id INTEGER PRIMARY KEY AUTOINCREMENT, client_id INTEGER, date TEXT, type TEXT, amount REAL, note TEXT, ts INTEGER DEFAULT 0)`);
        db.run(`CREATE TABLE IF NOT EXISTS expenses (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, description TEXT, amount REAL, method TEXT, user TEXT, ts INTEGER DEFAULT 0)`);

        db.run(`CREATE TABLE IF NOT EXISTS shifts (id INTEGER PRIMARY KEY AUTOINCREMENT, date_open TEXT, date_close TEXT, user_open TEXT, user_close TEXT, monto_inicial REAL DEFAULT 0, status TEXT DEFAULT 'open', ts_open INTEGER DEFAULT 0)`);
        // Recuperación Admin
        db.get("SELECT * FROM users WHERE name = 'Admin'", (err, row) => {
            if (!row) db.run("INSERT INTO users (name, pin) VALUES ('Admin', '1111')");
        });
    });

    // 3. MIGRADOR DE COLUMNAS FALTANTES (¡ESTO SOLUCIONA TU ERROR!)
    // Función auxiliar para agregar columnas si faltan
    function addCol(table, col, type) {
        db.all(`PRAGMA table_info(${table})`, (err, cols) => {
            if (!err && cols) {
                // Verificamos si la columna ya existe
                const exists = cols.some(c => c.name === col);
                if (!exists) {
                    console.log(`[Migración] Agregando columna ${col} a ${table}...`);
                    db.run(`ALTER TABLE ${table} ADD COLUMN ${col} ${type}`, (err) => {
                        if (err) console.log(`Error migrando ${col}:`, err.message);
                    });
                }
            }
        });
    }

    // Lista de columnas nuevas que quizás tu DB vieja no tiene
    addCol('clients', 'dni', 'TEXT');
    addCol('clients', 'credit_limit', 'REAL DEFAULT 0'); // <--- AQUÍ ESTÁ EL ARREGLO
    addCol('products', 'brand', 'TEXT');
    addCol('products', 'category', 'TEXT');
    addCol('products', 'subcategory', 'TEXT');
    addCol('sales', 'cuit', 'TEXT');
    addCol('sales', 'fiscal', 'TEXT');
    addCol('sales', 'observations', 'TEXT');
    addCol('sales', 'cae', 'TEXT');
    addCol('sales', 'invoice_number', 'TEXT');
    addCol('shifts', 'ts_open', 'INTEGER DEFAULT 0');
    addCol('expenses', 'ts', 'INTEGER DEFAULT 0');
    addCol('client_movements', 'ts', 'INTEGER DEFAULT 0');

    // --- NUEVO: ROLES Y AUDITORÍA ---
    addCol('users', 'role', "TEXT DEFAULT 'Admin'");
    db.run(`CREATE TABLE IF NOT EXISTS audit_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, ts INTEGER, username TEXT, action TEXT, details TEXT)`);
}

// --- API ---

// 1. OBTENER TODO
ipcMain.handle('db-get-all', async () => {
    return new Promise((resolve) => {
        setTimeout(() => {
            // AGREGAMOS expenses: [] AQUI
            const data = { stock: [], users: [], sales: [], closures: [], clients: [], movements: [], expenses: [] };
            db.serialize(() => {
                db.all("SELECT * FROM products", (err, rows) => data.stock = rows || []);
                db.all("SELECT * FROM sales", (err, rows) => data.sales = (rows || []).map(r => ({ ...r, items: JSON.parse(r.items_json) })));
                db.all("SELECT * FROM closures", (err, rows) => data.closures = (rows || []).map(r => ({ ...r, details: JSON.parse(r.details_json) })));
                db.all("SELECT * FROM users", (err, rows) => data.users = rows || []);
                db.all("SELECT * FROM clients ORDER BY name ASC", (err, rows) => data.clients = rows || []);
                db.all("SELECT * FROM client_movements ORDER BY id DESC LIMIT 500", (err, rows) => data.movements = rows || []);

                // --- NUEVO: CARGAR GASTOS Y AUDITORÍA ---
                db.all("SELECT * FROM expenses ORDER BY id DESC LIMIT 100", (err, rows) => {
                    data.expenses = rows || [];
                    db.all("SELECT * FROM audit_logs ORDER BY id DESC", (errLogs, rowLogs) => {
                        data.audit_logs = rowLogs || [];
                        resolve(data);
                    });
                });
            });
        }, 500);
    });
});

// 2. GUARDAR PRODUCTO (CORREGIDO: AHORA GUARDA MARCA Y CATEGORIA)
ipcMain.handle('db-save-product', async (e, p) => {
    return new Promise((resolve) => {
        db.get("SELECT code FROM products WHERE code = ?", [p.code], (err, row) => {
            // ERROR ESTABA AQUI: Faltaban brand, category y subcategory en este array
            const params = [
                p.name,
                p.supplier,
                p.brand,       // Agregado
                p.category,    // Agregado
                p.subcategory, // Agregado
                p.size,
                p.color,
                p.cost,
                p.margin,
                p.price,
                p.qty
            ];

            if (row) {
                // UPDATE
                db.run(`UPDATE products SET name=?, supplier=?, brand=?, category=?, subcategory=?, size=?, color=?, cost=?, margin=?, price=?, qty=? WHERE code=?`,
                    [...params, p.code], (err) => {
                        if (err) console.log("Error Update:", err); // Para ver errores
                        resolve(!err);
                    });
            } else {
                // INSERT
                db.run(`INSERT INTO products VALUES (?,?,?,?,?,?,?,?,?,?,?,?)`,
                    [p.code, ...params], (err) => {
                        if (err) console.log("Error Insert:", err); // Para ver errores
                        resolve(!err);
                    });
            }
        });
    });
});

// 3. Guardar Venta
ipcMain.handle('db-save-sale', async (e, s) => {
    return new Promise((resolve) => {
        db.serialize(() => {
            db.run("BEGIN TRANSACTION");
            try {
                db.run("INSERT INTO sales (id, date, seller, total, payment, card_name, surcharge, discount, client_id, items_json, observations) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                    [s.id, s.date, s.seller, s.total, s.payment, s.card_name, s.surcharge, s.discount, s.client_id, JSON.stringify(s.items), s.observations || ''],
                    function (insertErr) {
                        if (insertErr) {
                            // Fallback: intentar sin la columna observations
                            console.log("INSERT con observations falló, reintentando sin ella:", insertErr.message);
                            db.run("INSERT INTO sales (id, date, seller, total, payment, card_name, surcharge, discount, client_id, items_json) VALUES (?,?,?,?,?,?,?,?,?,?)",
                                [s.id, s.date, s.seller, s.total, s.payment, s.card_name, s.surcharge, s.discount, s.client_id, JSON.stringify(s.items)]);
                        }
                    });

                s.items.forEach(item => {
                    if (!item.isManual) db.run("UPDATE products SET qty = qty - ? WHERE code = ?", [item.qty, item.code]);
                });

                if (s.payment === 'Cuenta Corriente' && s.client_id) {
                    db.run("UPDATE clients SET balance = balance + ? WHERE id = ?", [s.total, s.client_id]);
                    const type = s.total < 0 ? 'DEVOLUCION' : 'DEUDA';
                    const amount = Math.abs(s.total);
                    db.run("INSERT INTO client_movements (client_id, date, type, amount, note) VALUES (?,?,?,?,?)",
                        [s.client_id, s.date, type, amount, `Venta #${s.id}${s.total < 0 ? ' (Devolución)' : ''}`]);
                }
                db.run("COMMIT", (err) => resolve({ success: !err }));
            } catch (error) {
                db.run("ROLLBACK");
                resolve({ success: false, error: error.message });
            }
        });
    });
});

// 4. Excel
// REEMPLAZAR EL BLOQUE ipcMain.handle('exportar-excel'...) EN main.js

ipcMain.handle('exportar-excel', async (event, payload) => {
    try {
        const { ventas, cierres, audit } = payload;
        const reportDir = path.join(app.getPath('documents'), 'Sistema_Reportes');
        if (!fs.existsSync(reportDir)) fs.mkdirSync(reportDir, { recursive: true });

        const today = new Date();
        const fileName = `Cierre_${String(today.getDate()).padStart(2, '0')}-${String(today.getMonth() + 1).padStart(2, '0')}-${today.getFullYear()}.xlsx`;
        const filePath = path.join(reportDir, fileName);
        const wb = xlsx.utils.book_new();

        // 1. Cálculos Generales
        let totalEfectivo = 0, totalDigital = 0, totalCtaCte = 0, totalGeneral = 0;
        ventas.forEach(v => {
            totalGeneral += v.total;
            if (v.payment === 'Efectivo') totalEfectivo += v.total;
            else if (v.payment === 'Cuenta Corriente') totalCtaCte += v.total;
            else totalDigital += v.total;
        });

        // 2. Armado de Hoja Resumen
        // Agregamos la columna "Detalle Físico" al final
        const resumenData = [
            ["REPORTE DE CAJA"], ["Fecha:", new Date().toLocaleDateString()], [" "],
            ["RESUMEN FINANCIERO", "MONTO"],
            ["Total Efectivo", totalEfectivo], ["Total Digital", totalDigital], ["Total Fiado", totalCtaCte],
            ["TOTAL VENDIDO", totalGeneral], [" "],
            ["HISTORIAL DE CIERRES (ARQUEOS)"],
            ["Hora", "Usuario", "Sistema (Debería)", "Real (Contado)", "Diferencia", "Estado", "Detalle Físico (Billetes)"]
        ];

        cierres.forEach(c => {
            // Convertimos el JSON de detalles a texto legible para el Excel
            // Ej: "5 x $1000 | 2 x $500"
            let detalleTexto = "";
            try {
                // Si viene como string, lo parseamos, si ya es objeto lo usamos
                const det = (typeof c.details === 'string') ? JSON.parse(c.details) : c.details;

                if (det) {
                    // Recorremos las claves (ej: $1000) y valores (ej: 5)
                    detalleTexto = Object.entries(det)
                        .map(([billete, cant]) => `${cant} x ${billete}`)
                        .join(" | ");
                }
            } catch (e) { detalleTexto = "-"; }

            resumenData.push([
                c.date.split(',')[1] || c.date, // Hora
                c.user,
                c.sys_cash,
                c.real_cash,
                c.diff,
                c.status,
                detalleTexto // <--- Aquí va el desglose de billetes
            ]);
        });

        xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(resumenData), "Resumen");

        // 3. Armado de Hoja Detalle Ventas
        const detalleData = [["ID", "Hora", "Vendedor", "Cliente", "Items", "Pago", "Total"]];
        ventas.forEach(v => {
            let itemsStr = "";
            try {
                const items = (typeof v.items === 'string') ? JSON.parse(v.items) : v.items;
                itemsStr = items.map(i => `${i.qty}x ${i.name}`).join(" | ");
            } catch (e) { itemsStr = "Error datos"; }

            detalleData.push([v.id, v.date.split(',')[1], v.seller, v.client_id ? 'Cliente' : 'Final', itemsStr, v.payment, v.total]);
        });
        xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(detalleData), "Detalle Ventas");

        // 4. Armado de Hoja Auditoría
        if (audit && audit.length > 0) {
            const auditData = [["Fecha y Hora", "Usuario", "Acción Crítica", "Detalles"]];
            audit.forEach(a => {
                const dateStr = new Date(a.ts).toLocaleString('es-AR');
                auditData.push([dateStr, a.username, a.action, a.details]);
            });
            xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(auditData), "Auditoría");
        }

        xlsx.writeFile(wb, filePath);
        shell.showItemInFolder(filePath);
        return { success: true };
    } catch (e) { return { success: false, error: e.message }; }
});

// --- FUNCION AYUDANTE DE FECHA ---
function getServerDate() {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, '0');
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const y = now.getFullYear();
    const hh = String(now.getHours()).padStart(2, '0');
    const mm = String(now.getMinutes()).padStart(2, '0');
    const ss = String(now.getSeconds()).padStart(2, '0');
    return `${d}/${m}/${y}, ${hh}:${mm}:${ss}`;
}

// Otros Handlers
// REEMPLAZAR EN main.js
// --- MEJORA 1: VALIDACIÓN DE CLIENTES DUPLICADOS ---
ipcMain.handle('db-save-client', async (e, c) => {
    return new Promise(resolve => {
        // 1. Si es un cliente NUEVO (no tiene ID), verificamos duplicados
        if (!c.id) {
            // Buscamos si existe alguien con ese DNI (si no está vacío) O con ese Nombre exacto
            const checkSql = "SELECT id FROM clients WHERE (dni = ? AND dni != '') OR name = ?";
            db.get(checkSql, [c.dni, c.name], (err, row) => {
                if (row) {
                    // Si encontramos una fila, devolvemos error
                    resolve({ success: false, error: "Ya existe un cliente con ese Nombre o DNI." });
                } else {
                    // Si no existe, procedemos al INSERT
                    db.run("INSERT INTO clients (name, dni, phone, address, balance) VALUES (?,?,?,?,?)",
                        [c.name, c.dni, c.phone, c.address, c.balance || 0],
                        (err) => resolve({ success: !err, error: err ? err.message : null }));
                }
            });
        } else {
            // 2. Si es UPDATE (tiene ID), actualizamos normal y registramos el movimiento si cambió la deuda
            db.get("SELECT balance FROM clients WHERE id = ?", [c.id], (err, row) => {
                if (err) return resolve({ success: false, error: err.message });

                const oldBalance = row ? parseFloat(row.balance) || 0 : 0;
                const newBalance = parseFloat(c.balance) || 0;
                const diff = newBalance - oldBalance;

                db.run("UPDATE clients SET name=?, dni=?, phone=?, address=?, balance=? WHERE id=?",
                    [c.name, c.dni, c.phone, c.address, newBalance, c.id], (errUpdate) => {
                        if (errUpdate) return resolve({ success: false, error: errUpdate.message });

                        if (Math.abs(diff) > 0.01) {
                            // Si aumentó la deuda, el tipo es DEUDA. Si disminuyó (ej: perdón de deuda), lo tomamos como PAGO.
                            const type = diff > 0 ? 'DEUDA' : 'PAGO';
                            db.run("INSERT INTO client_movements (client_id, date, type, amount, note, ts) VALUES (?,?,?,?,?,?)",
                                [c.id, getServerDate(), type, Math.abs(diff), 'Ajuste manual de saldo', Date.now()],
                                (errInsert) => resolve({ success: !errInsert, error: errInsert ? errInsert.message : null })
                            );
                        } else {
                            resolve({ success: true, error: null });
                        }
                    }
                );
            });
        }
    });
});
ipcMain.handle('db-pay-debt', async (e, { clientId, amount, note }) => {
    return new Promise(resolve => {
        db.serialize(() => {
            db.run("BEGIN");
            db.run("UPDATE clients SET balance = balance - ? WHERE id = ?", [amount, clientId]);
            db.run("INSERT INTO client_movements (client_id, date, type, amount, note, ts) VALUES (?,?,?,?,?,?)", [clientId, getServerDate(), 'PAGO', amount, note || 'Pago', Date.now()]);
            db.run("COMMIT", (err) => resolve({ success: !err }));
        });
    });
});
// REEMPLAZAR EN main.js
// EN MAIN.JS: Reemplaza todo el ipcMain.handle('db-bulk-update'...) por esto:

ipcMain.handle('db-bulk-update', async (e, { criteria, value, pct }) => {
    const factor = 1 + (pct / 100);
    let sql = "";
    let params = [];

    // LÓGICA CORREGIDA:
    // Actualizamos Costo Y Precio multiplicando cada uno por el porcentaje.
    // Así, si el costo es 0 pero el precio es 100, el precio sube correctamente.

    if (criteria === 'supplier') {
        sql = `UPDATE products SET cost = ROUND(cost * ?, 2), price = ROUND(price * ?, 2) WHERE supplier = ?`;
        params = [factor, factor, value];
    }
    else if (criteria === 'category') {
        sql = `UPDATE products SET cost = ROUND(cost * ?, 2), price = ROUND(price * ?, 2) WHERE category = ?`;
        params = [factor, factor, value];
    }

    return new Promise(r => db.run(sql, params, function (err) {
        if (err) console.log(err);
        r({ success: !err, count: this.changes });
    }));
});
ipcMain.handle('db-save-closure', async (e, c) => {
    return new Promise(resolve => db.run("INSERT INTO closures (date, user, sys_cash, real_cash, diff, status, details_json) VALUES (?,?,?,?,?,?,?)", [c.date, c.user, c.sysCash, c.realCash, c.diff, c.status, JSON.stringify(c.details)], (err) => resolve(!err)));
});
ipcMain.handle('db-add-user', async (e, u) => { db.run("INSERT INTO users (name, pin, role) VALUES (?,?,?)", [u.name, u.pin, u.role || 'Vendedor']); return true; });

// --- REGISTRO DE AUDITORÍA ---
ipcMain.handle('db-log-audit', async (e, { username, action, details }) => {
    return new Promise(resolve => db.run("INSERT INTO audit_logs (ts, username, action, details) VALUES (?,?,?,?)", [Date.now(), username, action, details || ''], (err) => resolve(!err)));
});
// REEMPLAZAR EN main.js
ipcMain.handle('db-factory-reset', async () => {
    return new Promise(r => db.serialize(() => {
        // 1. Borramos todas las tablas
        db.run("DELETE FROM products");
        db.run("DELETE FROM sales");
        db.run("DELETE FROM closures");
        db.run("DELETE FROM clients");
        db.run("DELETE FROM client_movements");

        // --- LAS QUE FALTABAN ---
        db.run("DELETE FROM expenses");
        db.run("DELETE FROM shifts");
        db.run("DELETE FROM users");

        // 2. IMPORTANTE: Recrear al usuario Admin por defecto
        // Si no hacemos esto, el sistema queda vacío y nadie puede entrar.
        db.run("INSERT INTO users (name, pin) VALUES ('Admin', '1111')");

        r({ success: true });
    }));
});
// --- RESTAURACIÓN INTELIGENTE (ADMITE BACKUPS VIEJOS) ---
// --- RESTAURACIÓN INTELIGENTE (ADMITE BACKUPS VIEJOS) ---
// --- NUEVO: BORRAR CLIENTE ---
ipcMain.handle('db-delete-client', async (e, id) => {
    return new Promise(resolve => {
        // Primero verificamos si tiene deuda
        db.get("SELECT balance FROM clients WHERE id = ?", [id], (err, row) => {
            if (row && row.balance > 0.5) { // Si debe más de 50 centavos
                resolve({ success: false, error: "El cliente tiene deuda pendiente. No se puede borrar." });
            } else {
                // Si no debe nada, lo borramos
                db.run("DELETE FROM clients WHERE id = ?", [id], (err) => {
                    if (!err) {
                        // Opcional: Borrar historial de movimientos de ese cliente para limpiar DB
                        db.run("DELETE FROM client_movements WHERE client_id = ?", [id]);
                    }
                    resolve({ success: !err, error: err ? err.message : null });
                });
            }
        });
    });
});

ipcMain.handle('db-restore-backup', async (e, data) => {
    return new Promise((resolve) => {
        db.serialize(() => {
            db.run("BEGIN TRANSACTION");
            try {
                // 1. LIMPIEZA TOTAL (Borramos datos actuales para evitar duplicados)
                db.run("DELETE FROM products");
                db.run("DELETE FROM clients");
                db.run("DELETE FROM sales");
                db.run("DELETE FROM closures");
                db.run("DELETE FROM client_movements");
                db.run("DELETE FROM users");

                // 2. RESTAURAR PRODUCTOS (Adaptando estructura vieja a nueva)
                if (data.stock && data.stock.length > 0) {
                    const stmt = db.prepare(`INSERT INTO products (
                        code, name, supplier, brand, category, subcategory, 
                        size, color, cost, margin, price, qty
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)`);

                    data.stock.forEach(p => {
                        stmt.run(
                            p.code,
                            p.name,
                            p.supplier || 'General',   // Si no tiene proveedor, pone 'General'
                            p.brand || '',             // Rellena huecos nuevos con vacío
                            p.category || '',
                            p.subcategory || '',
                            p.size || '',
                            p.color || '',
                            p.cost || 0,
                            p.margin || 0,
                            p.price || 0,
                            p.qty || 0
                        );
                    });
                    stmt.finalize();
                }

                // 3. RESTAURAR CLIENTES (Adaptando estructura vieja)
                if (data.clients && data.clients.length > 0) {
                    const stmt = db.prepare("INSERT INTO clients (id, name, phone, address, balance, dni, credit_limit) VALUES (?,?,?,?,?,?,?)");
                    data.clients.forEach(c => {
                        stmt.run(
                            c.id,
                            c.name,
                            c.phone || '',
                            c.address || '',
                            c.balance || 0,
                            c.dni || '',           // Campo nuevo
                            c.credit_limit || 0    // Campo nuevo
                        );
                    });
                    stmt.finalize();
                }

                // 4. RESTAURAR VENTAS (Sales)
                if (data.sales && data.sales.length > 0) {
                    // Nota: items_json es vital. Si el backup viejo lo tiene, genial.
                    const stmt = db.prepare("INSERT INTO sales (id, date, seller, total, payment, card_name, surcharge, discount, client_id, items_json) VALUES (?,?,?,?,?,?,?,?,?,?)");
                    data.sales.forEach(s => {
                        // Aseguramos que items_json sea un string válido
                        let itemsStr = typeof s.items_json === 'string' ? s.items_json : JSON.stringify(s.items || []);

                        stmt.run(
                            s.id, s.date, s.seller, s.total, s.payment,
                            s.card_name || '',
                            s.surcharge || 0,
                            s.discount || 0,
                            s.client_id,
                            itemsStr
                        );
                    });
                    stmt.finalize();
                }

                // 5. RESTAURAR EL RESTO (Cierres, Movimientos, Usuarios)
                // Estos suelen cambiar menos, pero aplicamos la misma lógica si hiciera falta.
                if (data.closures) {
                    const stmt = db.prepare("INSERT INTO closures VALUES (?,?,?,?,?,?,?,?)");
                    data.closures.forEach(c => stmt.run(c.id, c.date, c.user, c.sys_cash, c.real_cash, c.diff, c.status, (typeof c.details_json === 'string' ? c.details_json : JSON.stringify(c.details))));
                    stmt.finalize();
                }

                if (data.movements) {
                    const stmt = db.prepare("INSERT INTO client_movements VALUES (?,?,?,?,?,?,?)");
                    data.movements.forEach(m => stmt.run(m.id, m.client_id, m.date, m.type, m.amount, m.note, m.ts || 0));
                    stmt.finalize();
                }

                if (data.users) {
                    const stmt = db.prepare("INSERT INTO users VALUES (?,?,?,?)");
                    data.users.forEach(u => {
                        // Handle old backups that didn't have role
                        stmt.run(u.id, u.name, u.pin, u.role || 'Admin');
                    });
                    stmt.finalize();
                }

                db.run("COMMIT", () => resolve({ success: true }));

            } catch (err) {
                db.run("ROLLBACK"); // Si algo falla, deshace todo para no romper la DB
                console.error("Error Restore:", err);
                resolve({ success: false, error: err.message });
            }
        });
    });
});
// Integración AFIP directa (WSAA + WSFE via SOAP)
const afipService = require('./afip-service');

ipcMain.handle('facturar-afip', async (e, reqData) => {
    const saleId = typeof reqData === 'object' ? reqData.saleId : reqData;
    const forceCbte = typeof reqData === 'object' ? reqData.cbteTipo : 'auto';
    const docStrFromUI = typeof reqData === 'object' ? reqData.docStr : null;

    return new Promise((resolve) => {
        db.get("SELECT * FROM sales WHERE id = ?", [saleId], async (err, sale) => {
            if (err || !sale) {
                return resolve({ success: false, error: "Venta no encontrada en la base de datos." });
            }

            try {
                const total = parseFloat(sale.total);
                if (isNaN(total) || total <= 0) {
                    return resolve({ success: false, error: "El total de la venta debe ser mayor a cero para facturar." });
                }

                // Punto de venta 2. Tipo 6 = Factura B (R.Inscripto → Consumidor Final)
                const ptoVta   = 2;
                let   cbteTipo = 6;
                let   docTipo  = 99; // Consumidor Final por defecto
                let   docNro   = 0;

                // Usamos el doc ingresado en el modal, o si está vacío, el guardado en la venta
                const docStr = (docStrFromUI && docStrFromUI.trim() !== '') ? docStrFromUI.trim() : (sale.cuit ? sale.cuit.trim() : '');

                // Si se ingresó un CUIT/DNI en la venta lo enviamos a AFIP
                if (docStr !== '') {
                    const docClean = docStr.replace(/\D/g, '');
                    if (docClean.length === 11) {
                        docTipo  = 80; // CUIT
                        docNro   = parseInt(docClean);
                        cbteTipo = 1;  // Factura A sugerida
                    } else if (docClean.length >= 7 && docClean.length <= 8) {
                        docTipo = 96; // DNI
                        docNro  = parseInt(docClean);
                    }
                }

                // Aplicar selección forzada de factura desde la UI
                if (forceCbte === '1') {
                    cbteTipo = 1; // Forzar Factura A
                } else if (forceCbte === '6') {
                    cbteTipo = 6; // Forzar Factura B
                }

                // AFIP Validación extra: Factura A requiere CUIT
                if (cbteTipo === 1 && docTipo !== 80) {
                    return resolve({ success: false, error: "Para emitir una Factura A debes ingresar un CUIT válido (11 dígitos)." });
                }

                const resultado = await afipService.crearComprobante({
                    total, ptoVta, cbteTipo, docTipo, docNro
                });

                // Guardar CAE y número de factura real en la DB
                db.run(
                    "UPDATE sales SET cae = ?, invoice_number = ?, fiscal = ? WHERE id = ?",
                    [resultado.cae, resultado.invoiceNumber, 'A', saleId],
                    (updateErr) => {
                        if (updateErr) {
                            resolve({ success: false, error: "Factura emitida en AFIP pero falló al guardar en tu sistema: " + updateErr.message });
                        } else {
                            resolve({
                                success:        true,
                                cae:            resultado.cae,
                                caeVto:         resultado.caeVto,
                                invoice_number: resultado.invoiceNumber
                            });
                        }
                    }
                );

            } catch (afipErr) {
                console.error('[AFIP] Error:', afipErr.message);
                resolve({ success: false, error: 'Error de AFIP: ' + afipErr.message });
            }
        });
    });
});
// En main.js
ipcMain.handle('db-delete-product', async (e, code) => {
    return new Promise(resolve => {
        db.run("DELETE FROM products WHERE code = ?", [code], (err) => {
            resolve({ success: !err });
        });
    });
});
let mainWindow;

// --- NUEVO: GUARDAR GASTO ---
ipcMain.handle('db-save-expense', async (e, ex) => {
    return new Promise(resolve => {
        db.run("INSERT INTO expenses (date, description, amount, method, user, ts) VALUES (?,?,?,?,?,?)",
            [ex.date, ex.description, ex.amount, ex.method, ex.user, Date.now()],
            (err) => resolve({ success: !err, error: err ? err.message : null })
        );
    });
});

function createWindow() {
    mainWindow = new BrowserWindow({ width: 1200, height: 800, webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false } });
    mainWindow.loadFile('index.html');
}

ipcMain.handle('print-ticket', async (event, content) => {
    const workerWindow = new BrowserWindow({ show: false });

    return new Promise((resolve) => {
        // IMPORTANTE: Registrar el listener ANTES de cargar la URL
        workerWindow.webContents.once('did-finish-load', () => {
            workerWindow.webContents.print({
                silent: false,
                printBackground: true
            }, (success, errorType) => {
                if (!success) console.log("Error impresión:", errorType);
                workerWindow.close();
                if (mainWindow && !mainWindow.isDestroyed()) {
                    mainWindow.focus();
                }
                resolve({ success });
            });
        });

        // Cargar DESPUÉS de registrar el listener
        workerWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(content)}`);
    });
});

app.whenReady().then(() => {
    initDB();
    createWindow();
});

// --- FORZAR FOCO ---
ipcMain.handle('focus-window', async () => {
    if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.show();
        mainWindow.focus();
    }
});

app.on('activate', () => {
    if (mainWindow && !mainWindow.isDestroyed()) mainWindow.focus();
});

// --- TURNOS (SHIFTS) ---
ipcMain.handle('db-get-active-shift', async () => {
    return new Promise(resolve => {
        db.get("SELECT * FROM shifts WHERE status = 'open' ORDER BY id DESC LIMIT 1", (err, row) => {
            resolve(row || null);
        });
    });
});

ipcMain.handle('db-open-shift', async (e, data) => {
    return new Promise(resolve => {
        db.run("INSERT INTO shifts (date_open, user_open, monto_inicial, status, ts_open) VALUES (?,?,?,?,?)",
            [data.date_open, data.user_open, data.monto_inicial, 'open', data.ts_open || Date.now()],
            function (err) {
                resolve({ success: !err, id: this ? this.lastID : null, error: err ? err.message : null });
            }
        );
    });
});

ipcMain.handle('db-close-shift', async (e, data) => {
    return new Promise(resolve => {
        db.run("UPDATE shifts SET status = 'closed', date_close = ?, user_close = ? WHERE id = ?",
            [data.date_close, data.user_close, data.id],
            (err) => resolve({ success: !err, error: err ? err.message : null })
        );
    });
});

ipcMain.handle('db-update-shift-fondo', async (e, data) => {
    return new Promise(resolve => {
        db.run("UPDATE shifts SET monto_inicial = ? WHERE id = ?",
            [data.monto_inicial, data.id],
            (err) => resolve({ success: !err })
        );
    });
});

// --- OBTENER ÚLTIMO CIERRE (para fondo inicial automático) ---
ipcMain.handle('db-get-last-closure', async () => {
    return new Promise(resolve => {
        db.get("SELECT sys_cash FROM closures ORDER BY id DESC LIMIT 1", (err, row) => {
            resolve(row ? row.sys_cash : 0);
        });
    });
});

// --- GENERAR DOCX ETIQUETAS PARA HOJAS ADHESIVAS ---
// Formato: A4 con 24 etiquetas (3 columnas x 8 filas)
// Compatible con hojas adhesivas estándar (~66mm x 33mm por etiqueta)
ipcMain.handle('generate-word-labels', async (event, products) => {
    try {
        const { createCanvas } = require('canvas');
        const JsBarcode = require('jsbarcode');
        const { ImageRun, AlignmentType, convertInchesToTwip, VerticalAlign, HeightRule } = require('docx');

        const docDir = path.join(app.getPath('documents'), 'Sistema_Etiquetas');
        if (!fs.existsSync(docDir)) fs.mkdirSync(docDir, { recursive: true });

        const today = new Date();
        const dateStr = `${String(today.getDate()).padStart(2, '0')}-${String(today.getMonth() + 1).padStart(2, '0')}-${today.getFullYear()}`;
        const fileName = `Etiquetas_${dateStr}.docx`;
        const filePath = path.join(docDir, fileName);

        // --- Función para generar barcode compacto para etiquetas adhesivas ---
        function generateBarcodeImage(code) {
            const canvas = createCanvas(200, 70);
            JsBarcode(canvas, String(code), {
                format: "CODE128",
                width: 1.5,
                height: 40,
                displayValue: true,
                fontSize: 11,
                font: "monospace",
                fontOptions: "bold",
                textMargin: 2,
                margin: 3,
                background: "#ffffff",
                lineColor: "#000000"
            });
            return canvas.toBuffer('image/png');
        }

        // --- Expandir productos según cantidad solicitada ---
        const expandedProducts = [];
        for (const p of products) {
            const qty = p.labelQty || 1;
            for (let q = 0; q < qty; q++) {
                expandedProducts.push(p);
            }
        }

        // --- Configuración de la grilla: 3 columnas x 10 filas = 30 por página ---
        const COLS = 3;
        const ROWS_PER_PAGE = 8;
        const LABELS_PER_PAGE = COLS * ROWS_PER_PAGE;

        // Medidas en DXA (1 inch = 1440 DXA, 1mm ≈ 56.7 DXA)
        // Etiqueta: ~66mm ancho x 33mm alto (más holgada para que no quede apretado)
        const LABEL_WIDTH = 3780;  // ~66.7mm en DXA
        const LABEL_HEIGHT = 1870; // ~33mm en DXA

        // Márgenes de página para centrar las 3 columnas en A4 (210mm)
        // 3 etiquetas x 66.7mm = 200mm, sobran 10mm → 5mm cada lado
        const PAGE_MARGIN_LEFT = 284;   // ~5mm
        const PAGE_MARGIN_RIGHT = 284;  // ~5mm
        // A4 alto = 297mm, 8 filas x 33mm = 264mm, sobran 33mm → ~12mm arriba/abajo
        const PAGE_MARGIN_TOP = 680;    // ~12mm
        const PAGE_MARGIN_BOTTOM = 567; // ~10mm

        // Bordes finos punteados (guía de corte, no se ven en la adhesiva)
        const thinBorder = {
            style: BorderStyle.DOTTED,
            size: 1,
            color: "CCCCCC"
        };
        const cellBorders = {
            top: thinBorder,
            bottom: thinBorder,
            left: thinBorder,
            right: thinBorder
        };

        // --- Generar páginas ---
        const sections = [];
        const totalPages = Math.ceil(expandedProducts.length / LABELS_PER_PAGE);

        for (let page = 0; page < totalPages; page++) {
            const pageProducts = expandedProducts.slice(
                page * LABELS_PER_PAGE,
                (page + 1) * LABELS_PER_PAGE
            );

            // Agrupar en filas de 3
            const rows = [];
            for (let i = 0; i < pageProducts.length; i += COLS) {
                rows.push(pageProducts.slice(i, i + COLS));
            }

            // Completar hasta 10 filas si la última página no está llena
            while (rows.length < ROWS_PER_PAGE) {
                rows.push([]);
            }

            // Crear filas de tabla
            const tableRows = [];

            for (const rowProds of rows) {
                const cells = [];

                for (let c = 0; c < COLS; c++) {
                    const prod = rowProds[c];

                    if (prod) {
                        const barcodeBuffer = generateBarcodeImage(prod.code);

                        cells.push(
                            new TableCell({
                                borders: cellBorders,
                                width: { size: LABEL_WIDTH, type: WidthType.DXA },
                                verticalAlign: VerticalAlign.CENTER,
                                margins: {
                                    top: 80,
                                    bottom: 60,
                                    left: 57,   // ~1mm
                                    right: 57
                                },
                                children: [
                                    // Nombre del producto (compacto)
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        spacing: { before: 0, after: 20, line: 200 },
                                        children: [
                                            new TextRun({
                                                text: (prod.name || '').substring(0, 28),
                                                bold: true,
                                                size: 14,  // 7pt - compacto
                                                font: "Arial"
                                            })
                                        ]
                                    }),
                                    // Código de barras
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        spacing: { before: 0, after: 0 },
                                        children: [
                                            new ImageRun({
                                                data: barcodeBuffer,
                                                transformation: {
                                                    width: 150,
                                                    height: 52
                                                },
                                                type: 'png'
                                            })
                                        ]
                                    })
                                ]
                            })
                        );
                    } else {
                        // Celda vacía
                        cells.push(
                            new TableCell({
                                borders: cellBorders,
                                width: { size: LABEL_WIDTH, type: WidthType.DXA },
                                children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [] })]
                            })
                        );
                    }
                }

                tableRows.push(
                    new TableRow({
                        height: { value: LABEL_HEIGHT, rule: HeightRule.EXACT },
                        children: cells
                    })
                );
            }

            sections.push({
                properties: {
                    page: {
                        size: {
                            width: convertInchesToTwip(8.27),   // A4 ancho: 210mm
                            height: convertInchesToTwip(11.69)  // A4 alto: 297mm
                        },
                        margin: {
                            top: PAGE_MARGIN_TOP,
                            bottom: PAGE_MARGIN_BOTTOM,
                            left: PAGE_MARGIN_LEFT,
                            right: PAGE_MARGIN_RIGHT
                        }
                    }
                },
                children: [
                    new Table({
                        rows: tableRows,
                        width: { size: 100, type: WidthType.PERCENTAGE }
                    })
                ]
            });
        }

        const doc = new Document({ sections });

        const buffer = await Packer.toBuffer(doc);
        fs.writeFileSync(filePath, buffer);
        shell.showItemInFolder(filePath);
        return { success: true, count: expandedProducts.length, pages: totalPages };
    } catch (e) {
        console.error("Error Export Word Etiquetas:", e);
        return { success: false, error: e.message };
    }
});