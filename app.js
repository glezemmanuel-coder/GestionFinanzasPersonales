const $ = (selector) => document.querySelector(selector);
const STORAGE_KEY = 'saldo-expenses-v1';
const SETTINGS_KEY = 'saldo-budgets-v1';
const PEOPLE_KEY = 'saldo-people-v1';
const PAYMENTS_KEY = 'saldo-payments-v1';
const CARDS_KEY = 'saldo-cards-v1';
const MSI_KEY = 'saldo-installments-v1';
const PROFILE_KEY = 'saldo-profile-v1';
const CARD_PAYMENTS_KEY = 'saldo-card-payments-v1';
const INCOMES_KEY = 'saldo-incomes-v1';
const DATABASE_NAME = 'saldo-database';
const DEFAULT_CARDS = [{ id: '2now', label: '2Now', cutoff: 9, paymentDueDay: 0, paymentDueOffset: 1, color: '#8270e4', active: true }, { id: 'air', label: 'Air', cutoff: 15, paymentDueDay: 0, paymentDueOffset: 1, color: '#6296d5', active: true }];
const DEFAULT_PEOPLE = [{ id: 'yo', name: 'Yo', tracksDebt: false, active: true }, { id: 'esposa', name: 'Mi esposa', tracksDebt: true, active: true }];
let cards = readJSON('saldo-cards-v1', DEFAULT_CARDS);
let people = readJSON(PEOPLE_KEY, DEFAULT_PEOPLE);
const cardPalette = ['#8270e4', '#6296d5', '#55aa88', '#d28a54', '#d36f94', '#4aa5a3'];
const formatter = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', maximumFractionDigits: 2 });
const dateFormatter = new Intl.DateTimeFormat('es-MX', { day: 'numeric', month: 'short' });
const monthFormatter = new Intl.DateTimeFormat('es-MX', { month: 'long', year: 'numeric' });
const today = new Date();
today.setHours(0, 0, 0, 0);
let expenses = readJSON(STORAGE_KEY, []);
let payments = readJSON(PAYMENTS_KEY, []);
let installmentPlans = readJSON(MSI_KEY, []);
let userProfile = readJSON(PROFILE_KEY, { name: '' });
let cardPayments = readJSON(CARD_PAYMENTS_KEY, []);
let incomes = readJSON(INCOMES_KEY, []);
let budgets = { total: 0, now: 0, air: 0, ...readJSON(SETTINGS_KEY, {}) };
let database = null;
let excelHandle = null;
let excelSaveTimer = null;
let toastTimer;
let recordActionTarget = null;
const EXCEL_HEADERS = {
  expenses: ['id', 'fecha', 'descripcion', 'monto_mxn', 'tarjeta_id', 'tarjeta', 'persona_id', 'persona', 'por_cobrar', 'categoria', 'nota', 'pagado', 'creado_en'],
  payments: ['id', 'fecha', 'persona_id', 'persona', 'tarjeta_id', 'tarjeta', 'monto_mxn', 'nota', 'creado_en'],
  people: ['id', 'nombre', 'me_debe_nuevos_gastos', 'activa'],
  cards: ['id', 'nombre', 'dia_de_corte', 'dia_limite_pago', 'mes_limite_pago_offset', 'tope_por_ciclo_mxn', 'color', 'activa'],
  cardPayments: ['id', 'fecha_pago', 'tarjeta_id', 'tarjeta', 'corte_pagado', 'monto_mxn', 'nota', 'creado_en'],
  incomes: ['id', 'fecha', 'fuente', 'monto_mxn', 'nota', 'creado_en'],
  budgets: ['concepto', 'monto_mxn'],
  installments: ['id', 'fecha_compra', 'descripcion', 'monto_total_mxn', 'tarjeta_id', 'tarjeta', 'persona_id', 'persona', 'plazo_meses', 'mensualidades_pagadas', 'nota', 'creado_en']
};

function readJSON(key, fallback) { try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; } }
function writeData(key, value, localKey) {
  localStorage.setItem(localKey, JSON.stringify(value));
  if (!database) return;
  try {
    const transaction = database.transaction('data', 'readwrite');
    transaction.objectStore('data').put({ key, value });
  } catch (error) { console.error('No se pudo escribir en IndexedDB', error); }
  if (excelHandle) scheduleExcelSave();
}
function persist() { writeData('expenses', expenses, STORAGE_KEY); }
function persistPayments() { writeData('payments', payments, PAYMENTS_KEY); }
function persistPeople() { writeData('people', people, PEOPLE_KEY); }
function persistCards() { writeData('cards', cards, CARDS_KEY); }
function persistBudgets() { writeData('budgets', budgets, SETTINGS_KEY); }
function persistInstallments() { writeData('installments', installmentPlans, MSI_KEY); }
function persistProfile() { writeData('profile', userProfile, PROFILE_KEY); }
function persistCardPayments() { writeData('cardPayments', cardPayments, CARD_PAYMENTS_KEY); }
function persistIncomes() { writeData('incomes', incomes, INCOMES_KEY); }
function normalizeState() {
  cards = (Array.isArray(cards) ? cards : DEFAULT_CARDS).map((card) => ({ ...card, cycleBudget: card.cycleBudget ?? (card.id === '2now' ? budgets.now : card.id === 'air' ? budgets.air : 0), paymentDueDay: Number(card.paymentDueDay) || 0, paymentDueOffset: Number(card.paymentDueOffset) === 0 ? 0 : 1, active: card.active !== false }));
  people = (Array.isArray(people) ? people : DEFAULT_PEOPLE).map((person) => ({ ...person, active: person.active !== false }));
  expenses = (Array.isArray(expenses) ? expenses : []).map((expense) => {
    const personId = expense.personId || legacyPersonId(expense.person);
    return { ...expense, personId, receivable: expense.receivable ?? personById(personId).tracksDebt };
  });
  payments = Array.isArray(payments) ? payments : [];
  cardPayments = (Array.isArray(cardPayments) ? cardPayments : []).map((payment) => ({ ...payment, amount: Math.max(0, Number(payment.amount) || 0), cardId: String(payment.cardId || '2now'), statementEnd: String(payment.statementEnd || ''), date: String(payment.date || isoDate(today)) }));
  incomes = (Array.isArray(incomes) ? incomes : []).map((income) => ({ ...income, amount: Math.max(0, Number(income.amount) || 0), date: String(income.date || isoDate(today)), source: String(income.source || 'Ingreso') }));
  installmentPlans = (Array.isArray(installmentPlans) ? installmentPlans : []).map((plan) => ({
    ...plan,
    amount: Math.max(0, Number(plan.amount) || 0),
    months: Math.max(1, Math.floor(Number(plan.months) || 1)),
    installmentsPaid: Math.min(Math.max(0, Math.floor(Number(plan.installmentsPaid) || 0)), Math.max(1, Math.floor(Number(plan.months) || 1)))
  }));
  userProfile = { name: String(userProfile?.name || '').trim().slice(0, 35) };
}
function openDatabase() {
  return new Promise((resolve, reject) => {
    if (!('indexedDB' in window)) { reject(new Error('IndexedDB no está disponible')); return; }
    const request = indexedDB.open(DATABASE_NAME, 1);
    request.onupgradeneeded = () => { if (!request.result.objectStoreNames.contains('data')) request.result.createObjectStore('data', { keyPath: 'key' }); };
    request.onsuccess = () => { request.result.onversionchange = () => request.result.close(); resolve(request.result); };
    request.onerror = () => reject(request.error || new Error('No se pudo abrir la base local'));
  });
}
function readDatabaseValue(key) {
  return new Promise((resolve, reject) => {
    const request = database.transaction('data', 'readonly').objectStore('data').get(key);
    request.onsuccess = () => resolve(request.result?.value);
    request.onerror = () => reject(request.error);
  });
}
function putDatabaseValue(key, value) {
  return new Promise((resolve, reject) => {
    const transaction = database.transaction('data', 'readwrite');
    transaction.objectStore('data').put({ key, value });
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
    transaction.onabort = () => reject(transaction.error || new Error('No se pudo guardar el registro'));
  });
}
async function initDatabase() {
  database = await openDatabase();
  const collections = [
    { key: 'expenses', localKey: STORAGE_KEY, fallback: expenses },
    { key: 'payments', localKey: PAYMENTS_KEY, fallback: payments },
    { key: 'people', localKey: PEOPLE_KEY, fallback: people },
    { key: 'cards', localKey: CARDS_KEY, fallback: cards },
    { key: 'budgets', localKey: SETTINGS_KEY, fallback: budgets },
    { key: 'installments', localKey: MSI_KEY, fallback: installmentPlans },
    { key: 'profile', localKey: PROFILE_KEY, fallback: userProfile },
    { key: 'cardPayments', localKey: CARD_PAYMENTS_KEY, fallback: cardPayments },
    { key: 'incomes', localKey: INCOMES_KEY, fallback: incomes }
  ];
  const loaded = {};
  for (const item of collections) {
    const saved = await readDatabaseValue(item.key);
    loaded[item.key] = saved === undefined ? readJSON(item.localKey, item.fallback) : saved;
  }
  expenses = loaded.expenses;
  payments = loaded.payments;
  people = loaded.people;
  cards = loaded.cards;
  budgets = { total: 0, now: 0, air: 0, ...loaded.budgets };
  installmentPlans = loaded.installments;
  userProfile = loaded.profile;
  cardPayments = loaded.cardPayments;
  incomes = loaded.incomes;
  normalizeState();
  for (const item of collections) {
    const value = ({ expenses, payments, people, cards, budgets, installments: installmentPlans, profile: userProfile, cardPayments, incomes })[item.key];
    localStorage.setItem(item.localKey, JSON.stringify(value));
    await putDatabaseValue(item.key, value);
  }
  const savedHandle = await readDatabaseValue('excelFileHandle');
  if (savedHandle) {
    try {
      if (await savedHandle.queryPermission({ mode: 'readwrite' }) === 'granted') {
        excelHandle = savedHandle;
        const imported = await readExcelDatabase(await excelHandle.getFile());
        if (imported) {
          applyExcelData(imported);
          await saveAllCollections();
          setExcelStatus(`Excel conectado: ${excelHandle.name} · guardado automático`);
        } else { excelHandle = null; setExcelStatus('El libro vinculado ya no tiene las hojas de GFP · vuelve a conectarlo'); }
      } else setExcelStatus('Base local activa · vuelve a conectar el Excel para sincronizarlo');
    } catch (error) { excelHandle = null; console.info('Excel no disponible en esta sesión', error); setExcelStatus('Base local activa · no se pudo abrir el Excel vinculado'); }
  } else setExcelStatus('Datos guardados en este dispositivo · conecta un Excel para sincronizar el archivo');
  if (navigator.storage?.persist) navigator.storage.persist().catch(() => {});
}
function setExcelStatus(message) { const status = $('#excelStatus'); if (status) status.textContent = message; }
function parseExcelBoolean(value) { return value === true || value === 1 || ['true', '1', 'sí', 'si', 'yes'].includes(String(value).trim().toLocaleLowerCase('es-MX')); }
function excelDate(value) {
  if (value instanceof Date) return isoDate(value);
  if (typeof value === 'number' && window.XLSX?.SSF) {
    const parts = XLSX.SSF.parse_date_code(value);
    if (parts) return isoDate(new Date(parts.y, parts.m - 1, parts.d));
  }
  const text = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? isoDate(today) : isoDate(parsed);
}
function workbookRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  return sheet ? XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true }) : [];
}
async function readExcelDatabase(file) {
  if (!window.XLSX) throw new Error('No está cargado el lector de Excel');
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
  const expected = ['Gastos', 'Abonos', 'Personas', 'Tarjetas', 'Presupuesto'];
  if (!expected.every((name) => workbook.SheetNames.includes(name))) return null;
  const rows = Object.fromEntries([...expected, 'Meses sin intereses', 'Pagos a tarjeta', 'Ingresos'].map((name) => [name, workbookRows(workbook, name)]));
  return {
    expenses: rows.Gastos.filter((row) => row.id).map((row) => ({
      id: String(row.id), date: excelDate(row.fecha), description: String(row.descripcion || ''), amount: Number(row.monto_mxn) || 0,
      card: String(row.tarjeta_id || '2now'), personId: String(row.persona_id || 'yo'), receivable: parseExcelBoolean(row.por_cobrar),
      category: String(row.categoria || 'Otro'), note: String(row.nota || ''), paid: parseExcelBoolean(row.pagado), createdAt: String(row.creado_en || '')
    })),
    payments: rows.Abonos.filter((row) => row.id).map((row) => ({
      id: String(row.id), date: excelDate(row.fecha), personId: String(row.persona_id || 'yo'), cardId: String(row.tarjeta_id || '2now'),
      amount: Number(row.monto_mxn) || 0, note: String(row.nota || ''), createdAt: String(row.creado_en || '')
    })),
    people: rows.Personas.filter((row) => row.id && row.nombre).map((row) => ({ id: String(row.id), name: String(row.nombre), tracksDebt: parseExcelBoolean(row.me_debe_nuevos_gastos), active: parseExcelBoolean(row.activa) })),
    cards: rows.Tarjetas.filter((row) => row.id && row.nombre).map((row) => ({ id: String(row.id), label: String(row.nombre), cutoff: Number(row.dia_de_corte) || 1, paymentDueDay: Number(row.dia_limite_pago) || 0, paymentDueOffset: String(row.mes_limite_pago_offset).trim() === '0' ? 0 : 1, cycleBudget: Number(row.tope_por_ciclo_mxn) || 0, color: String(row.color || '#8270e4'), active: parseExcelBoolean(row.activa) })),
    budgets: { total: Number(rows.Presupuesto.find((row) => row.concepto === 'Mensual total')?.monto_mxn) || 0 },
    installments: workbook.Sheets['Meses sin intereses'] ? rows['Meses sin intereses'].filter((row) => row.id).map((row) => ({
      id: String(row.id), date: excelDate(row.fecha_compra), description: String(row.descripcion || ''), amount: Number(row.monto_total_mxn) || 0,
      cardId: String(row.tarjeta_id || '2now'), personId: String(row.persona_id || 'yo'), months: Number(row.plazo_meses) || 1,
      installmentsPaid: Number(row.mensualidades_pagadas) || 0, note: String(row.nota || ''), createdAt: String(row.creado_en || '')
    })) : null,
    cardPayments: workbook.Sheets['Pagos a tarjeta'] ? rows['Pagos a tarjeta'].filter((row) => row.id).map((row) => ({ id: String(row.id), date: excelDate(row.fecha_pago), cardId: String(row.tarjeta_id || ''), statementEnd: excelDate(row.corte_pagado), amount: Number(row.monto_mxn) || 0, note: String(row.nota || ''), createdAt: String(row.creado_en || '') })) : null,
    incomes: workbook.Sheets['Ingresos'] ? rows.Ingresos.filter((row) => row.id).map((row) => ({ id: String(row.id), date: excelDate(row.fecha), source: String(row.fuente || 'Ingreso'), amount: Number(row.monto_mxn) || 0, note: String(row.nota || ''), createdAt: String(row.creado_en || '') })) : null
  };
}
function sheetForRows(headers, rows, widths) {
  const sheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  sheet['!cols'] = widths.map((wch) => ({ wch }));
  if (rows.length) sheet['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: headers.length - 1 } }) };
  return sheet;
}
function createExcelWorkbook() {
  if (!window.XLSX) throw new Error('No está cargada la biblioteca de Excel');
  const workbook = XLSX.utils.book_new();
  const guide = XLSX.utils.aoa_to_sheet([
    ['GFP | Base de datos'],
    ['El aplicativo guarda gastos, abonos, pagos a tarjetas e ingresos en sus hojas correspondientes.'],
    ['No cambies los nombres de las hojas ni los encabezados. Las tarjetas y personas archivadas conservan el historial.'],
    ['El archivo conectado se actualiza automáticamente al registrar cambios en GFP.']
  ]);
  guide['!cols'] = [{ wch: 100 }];
  XLSX.utils.book_append_sheet(workbook, guide, 'Inicio');
  const expenseRows = expenses.map((item) => [item.id, item.date, item.description, Number(item.amount), item.card, cardById(item.card).label, item.personId, personById(item.personId).name, Boolean(item.receivable), item.category, item.note || '', Boolean(item.paid), item.createdAt || '']);
  const paymentRows = payments.map((item) => [item.id, item.date, item.personId, personById(item.personId).name, item.cardId, cardById(item.cardId).label, Number(item.amount), item.note || '', item.createdAt || '']);
  const peopleRows = people.map((item) => [item.id, item.name, Boolean(item.tracksDebt), Boolean(item.active)]);
  const cardRows = cards.map((item) => [item.id, item.label, Number(item.cutoff), Number(item.paymentDueDay) || 0, Number(item.paymentDueOffset) || 0, Number(item.cycleBudget) || 0, item.color, Boolean(item.active)]);
  const cardPaymentRows = cardPayments.map((item) => [item.id, item.date, item.cardId, cardById(item.cardId).label, item.statementEnd, Number(item.amount), item.note || '', item.createdAt || '']);
  const incomeRows = incomes.map((item) => [item.id, item.date, item.source, Number(item.amount), item.note || '', item.createdAt || '']);
  const budgetRows = [['Mensual total', Number(budgets.total) || 0]];
  const installmentRows = installmentPlans.map((item) => [item.id, item.date, item.description, Number(item.amount), item.cardId, cardById(item.cardId).label, item.personId, personById(item.personId).name, Number(item.months), Number(item.installmentsPaid), item.note || '', item.createdAt || '']);
  const sheets = [
    ['Gastos', EXCEL_HEADERS.expenses, expenseRows, [38, 14, 30, 16, 24, 20, 24, 22, 14, 22, 36, 12, 25]],
    ['Abonos', EXCEL_HEADERS.payments, paymentRows, [38, 14, 24, 22, 24, 20, 16, 36, 25]],
    ['Personas', EXCEL_HEADERS.people, peopleRows, [38, 28, 29, 12]],
    ['Tarjetas', EXCEL_HEADERS.cards, cardRows, [38, 28, 16, 20, 24, 24, 14, 12]],
    ['Pagos a tarjeta', EXCEL_HEADERS.cardPayments, cardPaymentRows, [38, 16, 24, 24, 20, 16, 36, 25]],
    ['Ingresos', EXCEL_HEADERS.incomes, incomeRows, [38, 16, 28, 16, 36, 25]],
    ['Presupuesto', EXCEL_HEADERS.budgets, budgetRows, [34, 20]],
    ['Meses sin intereses', EXCEL_HEADERS.installments, installmentRows, [38, 14, 34, 20, 24, 22, 24, 22, 14, 22, 36, 25]]
  ];
  for (const [name, headers, rows, widths] of sheets) XLSX.utils.book_append_sheet(workbook, sheetForRows(headers, rows, widths), name);
  return workbook;
}
function applyExcelData(data) {
  expenses = data.expenses;
  payments = data.payments;
  if (Array.isArray(data.installments)) installmentPlans = data.installments;
  if (Array.isArray(data.cardPayments)) cardPayments = data.cardPayments;
  if (Array.isArray(data.incomes)) incomes = data.incomes;
  people = data.people.length ? data.people : DEFAULT_PEOPLE.map((person) => ({ ...person }));
  cards = data.cards.length ? data.cards : DEFAULT_CARDS.map((card) => ({ ...card, cycleBudget: 0 }));
  budgets = { total: 0, ...data.budgets };
  normalizeState();
}
async function saveAllCollections() {
  persist(); persistPayments(); persistPeople(); persistCards(); persistBudgets(); persistInstallments(); persistProfile(); persistCardPayments(); persistIncomes();
  await Promise.all([
    putDatabaseValue('expenses', expenses), putDatabaseValue('payments', payments), putDatabaseValue('people', people),
    putDatabaseValue('cards', cards), putDatabaseValue('budgets', budgets), putDatabaseValue('installments', installmentPlans), putDatabaseValue('cardPayments', cardPayments), putDatabaseValue('incomes', incomes)
  ]);
}
function downloadExcel() {
  const data = XLSX.write(createExcelWorkbook(), { bookType: 'xlsx', type: 'array', compression: true });
  const url = URL.createObjectURL(new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  const anchor = document.createElement('a'); anchor.href = url; anchor.download = 'gfp-base-datos.xlsx'; anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
async function saveExcelHandle() {
  if (!excelHandle) return;
  const data = XLSX.write(createExcelWorkbook(), { bookType: 'xlsx', type: 'array', compression: true });
  const writable = await excelHandle.createWritable();
  await writable.write(data);
  await writable.close();
  setExcelStatus(`Excel guardado · ${new Date().toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' })}`);
}
function scheduleExcelSave() {
  clearTimeout(excelSaveTimer);
  setExcelStatus('Guardando cambios en el Excel…');
  excelSaveTimer = setTimeout(() => saveExcelHandle().catch((error) => { console.error(error); setExcelStatus('No se pudo guardar el Excel · vuelve a conectarlo'); }), 350);
}
function legacyPersonId(value) { const text = String(value || 'Yo'); return text === 'Yo' ? 'yo' : text === 'Mi esposa' ? 'esposa' : people.find((person) => person.name === text)?.id || text; }
function personById(id) { return people.find((person) => person.id === id) || { id, name: id === 'yo' ? 'Yo' : id === 'esposa' ? 'Mi esposa' : 'Persona archivada', tracksDebt: true, active: false }; }
function personDebt(personId, cardId = null) {
  const owed = expenses.filter((item) => item.personId === personId && item.receivable && (!cardId || item.card === cardId));
  const allocations = new Map([...new Set(owed.map((item) => item.card))].flatMap((id) => [...allocationsForCard(id)]));
  return Math.max(0, total(owed) - total(owed.map((item) => ({ amount: allocations.get(item.id) || 0 }))));
}
function personCredit(personId, cardId = null) {
  const cardIds = cardId ? [cardId] : [...new Set(payments.filter((item) => item.personId === personId).map((item) => item.cardId))];
  return cardIds.reduce((sum, id) => {
    const expenseTotal = total(expenses.filter((item) => item.personId === personId && item.card === id));
    const paymentTotal = total(payments.filter((item) => item.personId === personId && item.cardId === id));
    return sum + Math.max(0, paymentTotal - expenseTotal);
  }, 0);
}
function dateFromISO(value) { const [year, month, day] = value.split('-').map(Number); return new Date(year, month - 1, day); }
function isoDate(date) { return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`; }
function dateISO(expense) { return expense.date; }
function cardById(id) { return cards.find((card) => card.id === id) || { id, label: id === '2now' ? '2Now' : id === 'air' ? 'Air' : 'Tarjeta archivada', cutoff: 1, color: '#8270e4', active: false }; }
function money(value) { return formatter.format(value || 0); }
function monthKey(date) { return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`; }
function inCurrentMonth(expense) { return monthKey(dateFromISO(dateISO(expense))) === monthKey(today); }
function atMonth(year, month, day) { return new Date(year, month, day); }
function cycleFor(card, date = today) {
  const cutoff = cardById(card).cutoff;
  const thisCutoff = atMonth(date.getFullYear(), date.getMonth(), cutoff);
  const end = date <= thisCutoff ? thisCutoff : atMonth(date.getFullYear(), date.getMonth() + 1, cutoff);
  const start = atMonth(end.getFullYear(), end.getMonth() - 1, cutoff + 1);
  return { start, end };
}
function inCycle(expense, card) {
  if (expense.card !== card) return false;
  const date = dateFromISO(expense.date);
  const { start, end } = cycleFor(card);
  return date >= start && date <= end;
}
function cardExpenses(card) { const id = typeof card === 'string' ? card : card.id; return expenses.filter((expense) => inCycle(expense, id)); }
function paymentsInCycle(cardId) { const { start, end } = cycleFor(cardId); return payments.filter((payment) => payment.cardId === cardId && dateFromISO(payment.date) >= start && dateFromISO(payment.date) <= end); }
function allocationsForCard(cardId) {
  const allocations = new Map();
  const peopleWithExpenses = new Set(expenses.filter((expense) => expense.card === cardId).map((expense) => expense.personId));
  for (const personId of peopleWithExpenses) {
    const debts = expenses.filter((expense) => expense.card === cardId && expense.personId === personId).sort((a, b) => a.date.localeCompare(b.date) || (a.createdAt || '').localeCompare(b.createdAt || ''));
    let remainingPayment = total(payments.filter((payment) => payment.cardId === cardId && payment.personId === personId));
    for (const debt of debts) {
      const applied = Math.min(Number(debt.amount), remainingPayment);
      allocations.set(debt.id, applied);
      remainingPayment -= applied;
      if (remainingPayment <= 0) break;
    }
  }
  return allocations;
}
function collectedForCycle(cardId) { const allocations = allocationsForCard(cardId); return total(cardExpenses(cardId).map((expense) => ({ amount: allocations.get(expense.id) || 0 }))); }
function total(items) { return items.reduce((sum, expense) => sum + Number(expense.amount || 0), 0); }
function installmentAmount(plan) { return Math.round((Number(plan.amount) / plan.months) * 100) / 100; }
function installmentPaidAmount(plan) { return plan.installmentsPaid >= plan.months ? Number(plan.amount) : Math.min(Number(plan.amount), installmentAmount(plan) * Number(plan.installmentsPaid || 0)); }
function installmentRemaining(plan) { return Math.max(0, Number(plan.amount) - installmentPaidAmount(plan)); }
function installmentsDueForCycle(cardId, statementEnd = null) {
  const targetEnd = statementEnd ? dateFromISO(statementEnd) : cycleFor(cardId).end;
  const currentCycleIndex = targetEnd.getFullYear() * 12 + targetEnd.getMonth();
  return installmentPlans.filter((plan) => plan.cardId === cardId && plan.installmentsPaid < plan.months).reduce((sum, plan) => {
    const purchaseCycle = cycleFor(cardId, dateFromISO(plan.date));
    const purchaseCycleIndex = purchaseCycle.end.getFullYear() * 12 + purchaseCycle.end.getMonth();
    const elapsedCycles = currentCycleIndex - purchaseCycleIndex;
    const scheduled = Math.min(plan.months, Math.max(0, elapsedCycles));
    const overdueCount = Math.max(0, scheduled - plan.installmentsPaid);
    return sum + Math.min(installmentRemaining(plan), installmentAmount(plan) * overdueCount);
  }, 0);
}
function statementCycle(cardId, statementEnd) {
  const end = dateFromISO(statementEnd);
  const start = atMonth(end.getFullYear(), end.getMonth() - 1, cardById(cardId).cutoff + 1);
  return { start, end };
}
function lastClosedStatementEnd(cardId, reference = today) {
  const card = cardById(cardId);
  const thisCutoff = atMonth(reference.getFullYear(), reference.getMonth(), card.cutoff);
  return isoDate(reference >= thisCutoff ? thisCutoff : atMonth(reference.getFullYear(), reference.getMonth() - 1, card.cutoff));
}
function statementExpenses(cardId, statementEnd) {
  const { start, end } = statementCycle(cardId, statementEnd);
  return expenses.filter((item) => item.card === cardId && dateFromISO(item.date) >= start && dateFromISO(item.date) <= end);
}
function statementTotal(cardId, statementEnd) {
  return total(statementExpenses(cardId, statementEnd)) + installmentsDueForCycle(cardId, statementEnd);
}
function statementPayments(cardId, statementEnd) {
  return total(cardPayments.filter((item) => item.cardId === cardId && item.statementEnd === statementEnd));
}
function paymentDueDate(cardId, statementEnd) {
  const card = cardById(cardId);
  if (!card.paymentDueDay) return null;
  const end = dateFromISO(statementEnd);
  const month = end.getMonth() + (Number(card.paymentDueOffset) || 0);
  const lastDay = new Date(end.getFullYear(), month + 1, 0).getDate();
  return atMonth(end.getFullYear(), month, Math.min(card.paymentDueDay, lastDay));
}
function dueStatus(cardId, statementEnd, remaining) {
  if (remaining <= 0) return 'Liquidado';
  const due = paymentDueDate(cardId, statementEnd);
  if (!due) return 'Configura fecha límite';
  const days = Math.ceil((due - today) / 86400000);
  return days < 0 ? `Venció hace ${Math.abs(days)} ${Math.abs(days) === 1 ? 'día' : 'días'}` : days === 0 ? 'Vence hoy' : days === 1 ? 'Vence mañana' : `Vence en ${days} días`;
}
function renderCardPayments() {
  const list = cards.filter((card) => card.active || cardPayments.some((payment) => payment.cardId === card.id) || expenses.some((expense) => expense.card === card.id) || installmentPlans.some((plan) => plan.cardId === card.id && plan.installmentsPaid < plan.months)).map((card) => {
    const end = lastClosedStatementEnd(card.id);
    const cycle = statementCycle(card.id, end);
    const due = statementTotal(card.id, end);
    const paid = statementPayments(card.id, end);
    const remaining = Math.max(0, due - paid);
    const dueDate = paymentDueDate(card.id, end);
    const dateLabel = dueDate ? new Intl.DateTimeFormat('es-MX', { day: 'numeric', month: 'short' }).format(dueDate) : 'Fecha límite sin definir';
    const history = cardPayments.filter((payment) => payment.cardId === card.id).sort((a, b) => b.date.localeCompare(a.date)).slice(0, 3);
    return `<article class="bank-card-payment" style="--card-accent:${card.color}"><div class="bank-card-heading"><div><span class="bank-card-label">${escapeHTML(card.label)}</span><small>Estado de cuenta ${dateFormatter.format(cycle.start)} – ${dateFormatter.format(cycle.end)}</small></div><span class="bank-due-status ${remaining <= 0 ? 'is-paid' : ''}">${dueStatus(card.id, end, remaining)}</span></div><div class="bank-card-totals"><div><small>Total del estado</small><b>${money(due)}</b></div><div><small>Pagado al banco</small><b>${money(paid)}</b></div><div><small>Por pagar</small><b>${money(remaining)}</b></div></div><div class="bank-payment-footer"><small>Fecha límite: ${dateLabel}</small><button type="button" class="text-button add-bank-payment" data-card-id="${escapeHTML(card.id)}" data-statement-end="${end}">＋ Registrar pago</button></div>${history.length ? `<div class="bank-payment-history">${history.map((payment) => `<div><span>${dateFormatter.format(dateFromISO(payment.date))} · ${payment.statementEnd === end ? 'este estado' : `corte ${dateFormatter.format(dateFromISO(payment.statementEnd))}`}</span><b>${money(payment.amount)}</b><button type="button" class="bank-payment-delete" data-bank-payment-id="${escapeHTML(payment.id)}" aria-label="Borrar pago">×</button></div>`).join('')}</div>` : ''}</article>`;
  }).join('');
  $('#bankPaymentCards').innerHTML = list || '<p class="empty-bank-payments">Agrega una tarjeta para comenzar a controlar sus pagos.</p>';
}
function updateBankStatementOptions(cardId, selectedEnd = null) {
  const select = $('#bankStatementEnd');
  const latest = lastClosedStatementEnd(cardId);
  const endDate = dateFromISO(latest);
  const options = Array.from({ length: 12 }, (_, index) => {
    const end = isoDate(atMonth(endDate.getFullYear(), endDate.getMonth() - index, cardById(cardId).cutoff));
    const { start } = statementCycle(cardId, end);
    const amountLeft = Math.max(0, statementTotal(cardId, end) - statementPayments(cardId, end));
    return `<option value="${end}">${dateFormatter.format(start)} – ${dateFormatter.format(dateFromISO(end))} · pendiente ${money(amountLeft)}</option>`;
  });
  select.innerHTML = options.join('');
  select.value = options.some((option) => option.includes(`value="${selectedEnd}"`)) ? selectedEnd : latest;
  const remaining = Math.max(0, statementTotal(cardId, select.value) - statementPayments(cardId, select.value));
  $('#bankPaymentAmount').value = remaining > 0 ? remaining.toFixed(2) : '';
}
function openBankPayment(cardId = cards.find((card) => card.active)?.id, statementEnd = null) {
  if (!cardId) { showToast('Agrega una tarjeta antes de registrar su pago'); return; }
  $('#bankPaymentCard').innerHTML = cards.map((card) => `<option value="${escapeHTML(card.id)}">${escapeHTML(card.label)}${card.active ? '' : ' · baja'}</option>`).join('');
  $('#bankPaymentCard').value = cardId;
  $('#bankPaymentDate').value = isoDate(today);
  $('#bankPaymentNote').value = '';
  updateBankStatementOptions(cardId, statementEnd);
  $('#bankPaymentDialog').showModal();
}
function totalInstallmentCommitment() { return installmentPlans.filter((plan) => plan.installmentsPaid < plan.months).reduce((sum, plan) => sum + installmentAmount(plan), 0); }
function setTheme(theme) {
  if (theme === 'system') document.documentElement.removeAttribute('data-theme');
  else document.documentElement.dataset.theme = theme;
  localStorage.setItem('saldo-theme', theme);
  $('#themeToggle').textContent = theme === 'dark' ? '☼' : '◐';
  drawChart();
}

function renderDebtBreakdown() {
  const debtors = people.map((person) => {
    const assigned = expenses.filter((expense) => expense.personId === person.id && expense.receivable);
    if (!assigned.length) return null;
    const gross = total(assigned);
    const pending = personDebt(person.id);
    const cardsForPerson = [...new Set(assigned.map((expense) => expense.card))].map((cardId) => {
      const cardExpensesForPerson = assigned.filter((expense) => expense.card === cardId);
      const cardGross = total(cardExpensesForPerson);
      const cardPending = personDebt(person.id, cardId);
      return { card: cardById(cardId), gross: cardGross, pending: cardPending, applied: Math.max(0, cardGross - cardPending) };
    });
    return { person, gross, pending, applied: Math.max(0, gross - pending), credit: personCredit(person.id), cards: cardsForPerson };
  }).filter(Boolean).sort((a, b) => b.pending - a.pending || a.person.name.localeCompare(b.person.name, 'es'));
  const pendingTotal = debtors.reduce((sum, item) => sum + item.pending, 0);
  const payables = people.map((person) => {
    const assigned = expenses.filter((expense) => expense.personId === person.id && !expense.receivable);
    if (!assigned.length) return null;
    const gross = total(assigned);
    const cardsForPerson = [...new Set(assigned.map((expense) => expense.card))].map((cardId) => {
      const cardExpensesForPerson = assigned.filter((expense) => expense.card === cardId);
      const cardGross = total(cardExpensesForPerson);
      const allocations = allocationsForCard(cardId);
      const applied = total(cardExpensesForPerson.map((expense) => ({ amount: allocations.get(expense.id) || 0 })));
      return { card: cardById(cardId), gross: cardGross, applied, pending: Math.max(0, cardGross - applied) };
    });
    return { person, gross, applied: cardsForPerson.reduce((sum, card) => sum + card.applied, 0), pending: cardsForPerson.reduce((sum, card) => sum + card.pending, 0), cards: cardsForPerson };
  }).filter(Boolean).sort((a, b) => b.pending - a.pending || a.person.name.localeCompare(b.person.name, 'es'));
  const payableTotal = payables.reduce((sum, item) => sum + item.pending, 0);
  $('#debtTotal').textContent = `${money(pendingTotal)} por cobrar`;
  $('#payableTotal').textContent = `${money(payableTotal)} por pagar`;
  $('#debtEmpty').hidden = debtors.length + payables.length > 0;
  const receivableMarkup = debtors.map(({ person, gross, applied, pending, credit, cards: cardDetails }) => `
    <article class="debt-person">
      <div class="debt-person-top"><div class="debt-person-name"><b>${escapeHTML(person.name)}</b><small>${cardDetails.length} ${cardDetails.length === 1 ? 'tarjeta' : 'tarjetas'} · ${expenses.filter((expense) => expense.personId === person.id && expense.receivable).length} gastos</small></div><span class="debt-balance ${pending === 0 ? 'settled' : ''}">${money(pending)}<small>${pending === 0 ? 'liquidado' : 'por cobrar'}</small></span></div>
      <div class="debt-stats"><div><small>Asignado</small><b>${money(gross)}</b></div><div><small>Abonos aplicados</small><b>${money(applied)}</b></div><div><small>Saldo pendiente</small><b>${money(pending)}</b></div></div>
      <div class="debt-card-list">${cardDetails.map(({ card, gross: cardGross, applied: cardApplied, pending: cardPending }) => `<div class="debt-card-row"><span class="debt-card-label"><i style="--card-color:${escapeHTML(card.color)}"></i>${escapeHTML(card.label)}</span><span>${money(cardGross)} asignado</span><span>${money(cardApplied)} abonado</span><b>${money(cardPending)}</b></div>`).join('')}</div>
      ${credit > 0 ? `<p class="debt-credit">Saldo a favor disponible: <b>${money(credit)}</b></p>` : ''}
    </article>`).join('');
  const payableMarkup = payables.map(({ person, gross, applied, pending, cards: cardDetails }) => `
    <article class="debt-person payable-person">
      <div class="debt-person-top"><div class="debt-person-name"><b>${escapeHTML(person.name)} · gasto propio</b><small>${cardDetails.length} ${cardDetails.length === 1 ? 'tarjeta' : 'tarjetas'} · lo cubres tú</small></div><span class="debt-balance payable-balance ${pending === 0 ? 'settled' : ''}">${money(pending)}<small>${pending === 0 ? 'cubierto' : 'por pagar'}</small></span></div>
      <div class="debt-stats"><div><small>Gasto propio</small><b>${money(gross)}</b></div><div><small>Abonos aplicados</small><b>${money(applied)}</b></div><div><small>Falta cubrir</small><b>${money(pending)}</b></div></div>
      <div class="debt-card-list">${cardDetails.map(({ card, gross: cardGross, applied: cardApplied, pending: cardPending }) => `<div class="debt-card-row"><span class="debt-card-label"><i style="--card-color:${escapeHTML(card.color)}"></i>${escapeHTML(card.label)}</span><span>${money(cardGross)} propio</span><span>${money(cardApplied)} abonado</span><b>${money(cardPending)}</b></div>`).join('')}</div>
    </article>`).join('');
  $('#debtBreakdown').innerHTML = `${receivableMarkup}${payableMarkup}`;
}

function renderInstallments() {
  const balance = installmentPlans.reduce((sum, plan) => sum + installmentRemaining(plan), 0);
  const paid = installmentPlans.reduce((sum, plan) => sum + installmentPaidAmount(plan), 0);
  const active = installmentPlans.filter((plan) => plan.installmentsPaid < plan.months);
  $('#installmentCount').textContent = installmentPlans.length;
  $('#installmentBalance').textContent = money(balance);
  $('#installmentPaidTotal').textContent = money(paid);
  $('#installmentMonthlyTotal').textContent = money(totalInstallmentCommitment());
  $('#installmentEmpty').hidden = installmentPlans.length > 0;
  $('#installmentList').innerHTML = [...installmentPlans].sort((a, b) => b.date.localeCompare(a.date)).map((plan) => {
    const person = personById(plan.personId);
    const card = cardById(plan.cardId);
    const complete = plan.installmentsPaid >= plan.months;
    const monthly = installmentAmount(plan);
    const paidPercent = Math.min(100, plan.installmentsPaid / plan.months * 100);
    return `<article class="installment-card ${complete ? 'complete' : ''}"><div class="installment-card-head"><div class="installment-title"><span class="installment-icon">◷</span><div><b>${escapeHTML(plan.description)}</b><small>${escapeHTML(person.name)} · ${escapeHTML(card.label)} · compra ${dateFormatter.format(dateFromISO(plan.date))}</small></div></div><div class="installment-actions"><button class="installment-action" type="button" data-msi-action="edit" data-msi-id="${escapeHTML(plan.id)}">Editar</button><button class="installment-action danger" type="button" data-msi-action="delete" data-msi-id="${escapeHTML(plan.id)}" aria-label="Eliminar ${escapeHTML(plan.description)}">×</button></div></div><div class="installment-progress-label"><span>${plan.installmentsPaid} de ${plan.months} mensualidades ${complete ? 'pagadas' : 'pagadas'}</span><b>${money(installmentRemaining(plan))} restante</b></div><div class="installment-progress"><i style="width:${paidPercent}%"></i></div><div class="installment-stats"><div><small>Compra</small><b>${money(plan.amount)}</b></div><div><small>Mensualidad</small><b>${money(monthly)}</b></div><div><small>Pagado</small><b>${money(installmentPaidAmount(plan))}</b></div></div>${plan.note ? `<p class="installment-note">${escapeHTML(plan.note)}</p>` : ''}<div class="installment-footer"><span class="installment-status ${complete ? 'done' : ''}">${complete ? '✓ Plan liquidado' : `Faltan ${plan.months - plan.installmentsPaid} mensualidades`}</span>${complete ? '' : `<button class="secondary-button" type="button" data-msi-action="pay" data-msi-id="${escapeHTML(plan.id)}">＋ Registrar mensualidad</button>`}</div></article>`;
  }).join('');
}

function renderFinancialHealth() {
  const monthExpenses = total(expenses.filter(inCurrentMonth));
  const monthlyMsi = totalInstallmentCommitment();
  const plannedSpend = monthExpenses + monthlyMsi;
  const hasBudget = Number(budgets.total) > 0;
  const ratio = hasBudget ? plannedSpend / Number(budgets.total) : 0;
  const percent = Math.round(ratio * 100);
  const state = !hasBudget ? 'unset' : ratio > 1 ? 'over' : ratio >= .85 ? 'caution' : 'good';
  const health = $('#financialHealth');
  health.className = `financial-health health-${state}`;
  $('#healthExpenseTotal').textContent = money(monthExpenses);
  $('#healthMsiTotal').textContent = money(monthlyMsi);
  $('#healthMeter').style.width = `${Math.min(100, percent)}%`;
  $('#healthBudgetButton').textContent = hasBudget ? 'Ajustar presupuesto' : 'Establecer presupuesto';
  if (!hasBudget) {
    $('#healthIcon').textContent = '✳';
    $('#financialHealthTitle').textContent = 'Define tu presupuesto';
    $('#healthMessage').textContent = 'Agrega un límite mensual para comparar tus gastos y mensualidades MSI con lo que tienes planeado.';
    $('#healthAvailable').textContent = '—';
    $('#healthPercent').textContent = '—';
    $('#healthPercentLabel').textContent = 'presupuesto mensual';
  } else if (ratio > 1) {
    $('#healthIcon').textContent = '↑';
    $('#financialHealthTitle').textContent = 'Presupuesto excedido';
    $('#healthMessage').textContent = `Vas ${money(plannedSpend - budgets.total)} por encima de lo que planeaste este mes.`;
    $('#healthAvailable').textContent = `-${money(plannedSpend - budgets.total)}`;
    $('#healthPercent').textContent = `${percent}%`;
    $('#healthPercentLabel').textContent = 'del presupuesto';
  } else if (ratio >= .85) {
    $('#healthIcon').textContent = '!';
    $('#financialHealthTitle').textContent = 'Cerca de tu límite';
    $('#healthMessage').textContent = `Ya usaste el ${percent}% del presupuesto mensual; queda poco margen para nuevos gastos.`;
    $('#healthAvailable').textContent = money(budgets.total - plannedSpend);
    $('#healthPercent').textContent = `${percent}%`;
    $('#healthPercentLabel').textContent = 'del presupuesto';
  } else {
    $('#healthIcon').textContent = '✓';
    $('#financialHealthTitle').textContent = 'Vas bien este mes';
    $('#healthMessage').textContent = `Llevas ${money(plannedSpend)} de ${money(budgets.total)} programados, incluyendo tus mensualidades MSI.`;
    $('#healthAvailable').textContent = money(budgets.total - plannedSpend);
    $('#healthPercent').textContent = `${percent}%`;
    $('#healthPercentLabel').textContent = 'del presupuesto';
  }
}

function nextCutoff(card) {
  let date = atMonth(today.getFullYear(), today.getMonth(), card.cutoff);
  if (date < today) date = atMonth(today.getFullYear(), today.getMonth() + 1, card.cutoff);
  return { date, days: Math.ceil((date - today) / 86400000) };
}
function renderIncomePanel() {
  const monthIncome = incomes.filter((income) => inCurrentMonth(income));
  const monthReceipts = payments.filter(inCurrentMonth);
  $('#incomeMonthTotal').textContent = money(total(monthIncome) + total(monthReceipts));
  const sorted = [...incomes].sort((a, b) => b.date.localeCompare(a.date) || (b.createdAt || '').localeCompare(a.createdAt || '')).slice(0, 8);
  $('#incomeEmpty').hidden = incomes.length > 0;
  $('#incomeList').innerHTML = sorted.map((income) => `<article class="income-row"><div class="income-row-icon">↙</div><div class="income-row-label"><b>${escapeHTML(income.source)}</b><small>${dateFormatter.format(dateFromISO(income.date))}${income.note ? ` · ${escapeHTML(income.note)}` : ''}</small></div><strong>${money(income.amount)}</strong><button type="button" class="income-delete" data-income-id="${escapeHTML(income.id)}" aria-label="Borrar ingreso de ${escapeHTML(income.source)}">×</button></article>`).join('');
}
function renderDecisionKpis() {
  const monthSelect = $('#decisionMonth');
  const cardSelect = $('#decisionCard');
  const personSelect = $('#decisionPerson');
  const currentMonth = monthKey(today);
  const monthValues = [...new Set([currentMonth, ...expenses.map((expense) => monthKey(dateFromISO(expense.date)))])].sort().reverse();
  const selectedMonth = monthSelect.value || currentMonth;
  monthSelect.innerHTML = `<option value="${currentMonth}">Este mes</option><option value="all">Todo el historial</option>${monthValues.filter((month) => month !== currentMonth).map((month) => `<option value="${month}">${monthFormatter.format(dateFromISO(`${month}-01`))}</option>`).join('')}`;
  monthSelect.value = [...monthValues, 'all'].includes(selectedMonth) ? selectedMonth : currentMonth;
  const trackedCards = cards.filter((card) => card.active || expenses.some((expense) => expense.card === card.id) || installmentPlans.some((plan) => plan.cardId === card.id && plan.installmentsPaid < plan.months));
  const selectedCard = cardSelect.value || 'all';
  cardSelect.innerHTML = `<option value="all">Todas las tarjetas</option>${trackedCards.map((card) => `<option value="${escapeHTML(card.id)}">${escapeHTML(card.label)}</option>`).join('')}`;
  cardSelect.value = trackedCards.some((card) => card.id === selectedCard) || selectedCard === 'all' ? selectedCard : 'all';
  const selectedPerson = personSelect.value || 'all';
  personSelect.innerHTML = `<option value="all">Todas las personas</option>${people.map((person) => `<option value="${escapeHTML(person.id)}">${escapeHTML(person.name)}</option>`).join('')}`;
  personSelect.value = people.some((person) => person.id === selectedPerson) || selectedPerson === 'all' ? selectedPerson : 'all';
  const monthExpenses = expenses.filter((expense) => (monthSelect.value === 'all' || monthKey(dateFromISO(expense.date)) === monthSelect.value) && (cardSelect.value === 'all' || expense.card === cardSelect.value) && (personSelect.value === 'all' || expense.personId === personSelect.value));
  const monthSpend = total(monthExpenses);
  const currentExpenses = expenses.filter(inCurrentMonth);
  const currentSpend = total(currentExpenses);
  const selectedIsCurrent = monthSelect.value === currentMonth;
  const previousDate = selectedIsCurrent ? atMonth(today.getFullYear(), today.getMonth() - 1, 1) : monthSelect.value === 'all' ? null : dateFromISO(`${monthSelect.value}-01`);
  const previousMonth = previousDate ? selectedIsCurrent ? monthKey(previousDate) : monthKey(atMonth(previousDate.getFullYear(), previousDate.getMonth() - 1, 1)) : null;
  const previousDayLimit = selectedIsCurrent && previousDate ? Math.min(today.getDate(), atMonth(previousDate.getFullYear(), previousDate.getMonth() + 1, 0).getDate()) : Infinity;
  const previousSpend = previousMonth ? total(expenses.filter((expense) => monthKey(dateFromISO(expense.date)) === previousMonth && (cardSelect.value === 'all' || expense.card === cardSelect.value) && (personSelect.value === 'all' || expense.personId === personSelect.value) && dateFromISO(expense.date).getDate() <= previousDayLimit)) : 0;
  const monthlyMsi = total(installmentPlans.filter((plan) => plan.installmentsPaid < plan.months && (cardSelect.value === 'all' || plan.cardId === cardSelect.value) && (personSelect.value === 'all' || plan.personId === personSelect.value)).map((plan) => ({ amount: installmentAmount(plan) })));
  const globalMonthlyMsi = totalInstallmentCommitment();
  const plannedSpend = currentSpend + globalMonthlyMsi;
  const budget = Number(budgets.total) || 0;
  const daysInMonth = atMonth(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const spentThroughToday = total(monthExpenses.filter((expense) => dateFromISO(expense.date) <= today));
  const projection = selectedIsCurrent ? (spentThroughToday / Math.max(1, today.getDate())) * daysInMonth + monthlyMsi : monthSelect.value === 'all' ? null : monthSpend;
  const categories = new Map();
  monthExpenses.forEach((expense) => {
    const name = expense.category || 'Otro';
    categories.set(name, (categories.get(name) || 0) + Number(expense.amount || 0));
  });
  const topCategory = [...categories.entries()].sort((a, b) => b[1] - a[1])[0];
  const topCategoryShare = monthSpend > 0 && topCategory ? topCategory[1] / monthSpend : 0;
  const bankBalance = trackedCards.reduce((sum, card) => {
    const end = lastClosedStatementEnd(card.id);
    return sum + Math.max(0, statementTotal(card.id, end) - statementPayments(card.id, end));
  }, 0);
  const cycleCover = trackedCards.reduce((sum, card) => sum + Math.max(0, total(cardExpenses(card.id)) - collectedForCycle(card.id)) + installmentsDueForCycle(card.id), 0);
  const receivable = people.reduce((sum, person) => sum + personDebt(person.id), 0);
  const currentIncome = total(incomes.filter(inCurrentMonth));
  const currentReceipts = total(payments.filter(inCurrentMonth));
  const currentMargin = currentIncome + currentReceipts - total(expenses.filter(inCurrentMonth)) - globalMonthlyMsi;
  const cutoffs = trackedCards.filter((card) => cardSelect.value === 'all' || card.id === cardSelect.value).map((card) => ({ card, ...nextCutoff(card), amount: Math.max(0, total(cardExpenses(card.id)) - collectedForCycle(card.id)) + installmentsDueForCycle(card.id) })).sort((a, b) => a.days - b.days);
  const periodLabel = monthSelect.value === 'all' ? 'todo el historial' : monthFormatter.format(dateFromISO(`${monthSelect.value}-01`));
  const cardLabel = cardSelect.value === 'all' ? 'todas las tarjetas' : cardById(cardSelect.value).label;
  const personLabel = personSelect.value === 'all' ? 'todas las personas' : personById(personSelect.value).name;
  $('#decisionPeriodLabel').textContent = `Gasto: ${periodLabel} · ${cardLabel} · ${personLabel}. Los saldos de tarjeta y presupuesto son generales.`;
  $('#decisionSpend').textContent = money(monthSpend);
  $('#decisionSpendChange').textContent = monthSelect.value === 'all' ? `${monthExpenses.length} registros en todo el historial` : previousSpend > 0 ? `${monthSpend >= previousSpend ? '↑' : '↓'} ${Math.abs(Math.round((monthSpend - previousSpend) / previousSpend * 100))}% vs. periodo anterior (${money(previousSpend)})` : 'Sin base de comparación del periodo anterior';
  $('#decisionSpendChange').className = `decision-kpi-detail ${previousSpend > 0 ? monthSpend > previousSpend ? 'trend-up' : monthSpend < previousSpend ? 'trend-down' : '' : ''}`;
  $('#decisionProjection').textContent = projection === null ? '—' : money(projection);
  $('#decisionProjectionDetail').textContent = selectedIsCurrent ? `Promedio diario ${money(spentThroughToday / Math.max(1, today.getDate()))} + MSI ${money(monthlyMsi)}` : monthSelect.value === 'all' ? 'Elige un mes para ver su proyección o cierre' : 'Gasto registrado en el periodo cerrado';
  $('#decisionBudgetLeft').textContent = budget > 0 ? money(budget - plannedSpend) : 'Sin definir';
  $('#decisionBudgetUsage').textContent = budget > 0 ? `${Math.round(plannedSpend / budget * 100)}% utilizado · ${money(plannedSpend)} de ${money(budget)}` : 'Configura un presupuesto mensual';
  $('#decisionBudgetMeter').style.width = budget > 0 ? `${Math.min(100, plannedSpend / budget * 100)}%` : '0%';
  $('#decisionBudgetMeter').classList.toggle('is-over', budget > 0 && plannedSpend > budget);
  $('#decisionBankBalance').textContent = money(bankBalance);
  $('#decisionCycleCover').textContent = money(cycleCover);
  $('#decisionReceivable').textContent = money(receivable);
  $('#decisionIncome').textContent = money(currentIncome + currentReceipts);
  $('#decisionAvailableCash').textContent = currentIncome + currentReceipts > 0 ? money(currentMargin) : 'Registra ingresos';
  $('#decisionTopCategory').textContent = topCategory ? topCategory[0] : 'Sin datos';
  $('#decisionTopCategoryShare').textContent = topCategory ? `${money(topCategory[1])} · ${Math.round(topCategoryShare * 100)}% del gasto filtrado` : 'Agrega gastos para descubrir patrones';
  let insightTitle = 'Vas dentro de lo planeado';
  let insightText = `El gasto filtrado en ${periodLabel} es ${money(monthSpend)}.`;
  const currentUnfiltered = selectedIsCurrent && cardSelect.value === 'all' && personSelect.value === 'all';
  if (budget <= 0 && currentUnfiltered) { insightTitle = 'Define tu presupuesto mensual'; insightText = `Llevas ${money(plannedSpend)} entre gastos y mensualidades MSI; agrega un límite para medir tu margen.`; }
  else if (currentUnfiltered && plannedSpend > budget) { insightTitle = 'Tu presupuesto ya fue rebasado'; insightText = `Vas ${money(plannedSpend - budget)} por encima del límite. Revisa gastos próximos y mensualidades comprometidas.`; }
  else if (currentUnfiltered && projection > budget) { insightTitle = 'La tendencia podría superar tu límite'; insightText = `Si mantienes el ritmo actual, cerrarías cerca de ${money(projection)}, ${money(projection - budget)} por encima del presupuesto.`; }
  else if (currentUnfiltered && budget > 0 && plannedSpend / budget >= .85) { insightTitle = 'Queda poco margen'; insightText = `Has comprometido ${Math.round(plannedSpend / budget * 100)}% del presupuesto; quedan ${money(budget - plannedSpend)}.`; }
  else if (cutoffs[0]?.days <= 3 && cutoffs[0].amount > 0) { insightTitle = `Se acerca el corte de ${cutoffs[0].card.label}`; insightText = `Faltan ${cutoffs[0].days} ${cutoffs[0].days === 1 ? 'día' : 'días'} y llevas ${money(cutoffs[0].amount)} por cubrir en su ciclo actual.`; }
  else if (topCategoryShare >= .4) { insightTitle = `${topCategory[0]} concentra buena parte de tu gasto`; insightText = `${Math.round(topCategoryShare * 100)}% del gasto filtrado está en esta categoría. Revísala si buscas liberar presupuesto.`; }
  else if (previousSpend > 0 && monthSpend > previousSpend * 1.15) { insightTitle = 'Tu gasto va por encima del periodo anterior'; insightText = `Llevas ${money(monthSpend)}, ${Math.round((monthSpend / previousSpend - 1) * 100)}% más que en el periodo anterior comparable.`; }
  $('#decisionInsightTitle').textContent = insightTitle;
  $('#decisionInsightText').textContent = insightText;
  $('#decisionCutoffStrip').innerHTML = cutoffs.map(({ card, days, amount }) => `<article class="decision-cutoff-card"><span class="decision-cutoff-dot" style="--cutoff-color:${card.color}"></span><div><small>Próximo corte · ${escapeHTML(card.label)}</small><b>${days === 0 ? 'Hoy' : days === 1 ? 'Mañana' : `En ${days} días`}</b></div><span class="decision-cutoff-balance">${money(amount)}<small>por cubrir</small></span></article>`).join('');
}

function renderPersonKpis() {
  const monthSelect = $('#kpiMonth');
  const cardSelect = $('#kpiCard');
  const analyticsExpenses = [...expenses, ...installmentPlans.map((plan) => ({ ...plan, card: plan.cardId, category: 'Meses sin intereses' }))];
  const selectedMonth = monthSelect.value || monthKey(today);
  const months = [...new Set([monthKey(today), ...analyticsExpenses.map((expense) => monthKey(dateFromISO(expense.date)))])].sort().reverse();
  monthSelect.innerHTML = `<option value="${monthKey(today)}">Este mes</option><option value="all">Todo el historial</option>${months.filter((month) => month !== monthKey(today)).map((month) => {
    const [year, number] = month.split('-').map(Number);
    return `<option value="${month}">${monthFormatter.format(new Date(year, number - 1, 1))}</option>`;
  }).join('')}`;
  monthSelect.value = [...months, 'all'].includes(selectedMonth) ? selectedMonth : monthKey(today);
  const selectedCard = cardSelect.value || 'all';
  const trackedCards = cards.filter((card) => card.active || analyticsExpenses.some((expense) => expense.card === card.id));
  cardSelect.innerHTML = `<option value="all">Todas las tarjetas</option>${trackedCards.map((card) => `<option value="${escapeHTML(card.id)}">${escapeHTML(card.label)}${card.active ? '' : ' · baja'}</option>`).join('')}`;
  cardSelect.value = trackedCards.some((card) => card.id === selectedCard) || selectedCard === 'all' ? selectedCard : 'all';

  const filtered = analyticsExpenses.filter((expense) => (monthSelect.value === 'all' || monthKey(dateFromISO(expense.date)) === monthSelect.value) && (cardSelect.value === 'all' || expense.card === cardSelect.value));
  const personIds = [...new Set(filtered.map((expense) => expense.personId))];
  const orderedIds = [...people.filter((person) => personIds.includes(person.id)).map((person) => person.id), ...personIds.filter((id) => !people.some((person) => person.id === id))];
  $('#personKpiEmpty').hidden = filtered.length > 0;
  $('#personKpiList').innerHTML = orderedIds.map((personId) => {
    const person = personById(personId);
    const personExpenses = filtered.filter((expense) => expense.personId === personId);
    const spent = total(personExpenses);
    const count = personExpenses.length;
    const average = count ? spent / count : 0;
    const groupBy = (keyOf) => {
      const groups = new Map();
      personExpenses.forEach((expense) => {
        const label = keyOf(expense);
        const entry = groups.get(label) || { label, amount: 0, count: 0 };
        entry.amount += Number(expense.amount) || 0;
        entry.count += 1;
        groups.set(label, entry);
      });
      return [...groups.values()];
    };
    const categoriesByAmount = groupBy((expense) => expense.category || 'Otro').sort((a, b) => b.amount - a.amount);
    const categoriesByCount = [...categoriesByAmount].sort((a, b) => b.count - a.count || b.amount - a.amount);
    const descriptionsByCount = groupBy((expense) => (expense.description || 'Compra').trim().toLocaleLowerCase('es-MX')).sort((a, b) => b.count - a.count || b.amount - a.amount);
    const topSpent = categoriesByAmount[0];
    const topFrequent = categoriesByCount[0];
    const repeated = descriptionsByCount.find((entry) => entry.count > 1);
    const repeatedLabel = repeated ? personExpenses.find((expense) => (expense.description || 'Compra').trim().toLocaleLowerCase('es-MX') === repeated.label)?.description || repeated.label : '';
    const colorIndex = Math.max(0, people.findIndex((item) => item.id === personId));
    const topCategories = categoriesByAmount.slice(0, 3);
    const mostFrequentCategories = categoriesByCount.slice(0, 3);
    return `<article class="person-kpi-card"><div class="kpi-person-heading"><span class="kpi-avatar" style="--kpi-color:${cardPalette[colorIndex % cardPalette.length]}">${escapeHTML(person.name.trim().charAt(0).toLocaleUpperCase('es-MX') || '?')}</span><div class="kpi-person-name"><b>${escapeHTML(person.name)}</b><small>${count} ${count === 1 ? 'compra' : 'compras'} en ${monthSelect.value === 'all' ? 'todo el historial' : monthFormatter.format(dateFromISO(`${monthSelect.value}-01`))}</small></div><span class="kpi-person-total">${money(spent)}</span></div><div class="kpi-stat-grid"><div><small>Compras</small><b>${count}</b></div><div><small>Ticket promedio</small><b>${money(average)}</b></div><div><small>Tarjeta</small><b>${cardSelect.value === 'all' ? 'Todas' : escapeHTML(cardById(cardSelect.value).label)}</b></div></div><div class="kpi-insights"><div class="kpi-insight"><small>En qué gasta más</small><b>${topSpent ? escapeHTML(topSpent.label) : '—'}</b><span>${topSpent ? `${money(topSpent.amount)} · ${Math.round(topSpent.amount / spent * 100)}% del gasto` : 'Sin datos'}</span></div><div class="kpi-insight"><small>Categoría más frecuente</small><b>${topFrequent ? escapeHTML(topFrequent.label) : '—'}</b><span>${topFrequent ? `${topFrequent.count} ${topFrequent.count === 1 ? 'compra' : 'compras'}` : 'Sin datos'}</span></div></div><div class="kpi-lists"><div><h3>Mayor gasto por categoría</h3>${topCategories.map((entry) => `<div class="kpi-category-row"><div><span>${escapeHTML(entry.label)}</span><b>${money(entry.amount)}</b></div><i><em style="width:${spent ? entry.amount / spent * 100 : 0}%"></em></i></div>`).join('')}</div><div><h3>Gastos más frecuentes</h3>${mostFrequentCategories.map((entry) => `<div class="kpi-frequency-row"><span>${escapeHTML(entry.label)}</span><b>${entry.count} ${entry.count === 1 ? 'vez' : 'veces'}</b></div>`).join('')}${repeated ? `<p class="kpi-repeat-note">Compra repetida: <b>${escapeHTML(repeatedLabel)}</b> · ${repeated.count} veces</p>` : ''}</div></div></article>`;
  }).join('');
}

function renderSummary() {
  const monthly = expenses.filter(inCurrentMonth);
  const monthTotal = total(monthly);
  const plannedMonthTotal = monthTotal + totalInstallmentCommitment();
  $('#monthLabel').textContent = monthFormatter.format(today).toLocaleUpperCase('es-MX');
  $('#welcomeName').textContent = userProfile.name || 'bienvenido';
  $('#todayLabel').textContent = new Intl.DateTimeFormat('es-MX', { weekday: 'short', day: 'numeric', month: 'short' }).format(today);
  $('#monthTotal').textContent = money(monthTotal);
  $('#monthCount').textContent = `${monthly.length} ${monthly.length === 1 ? 'compra registrada' : 'compras registradas'}`;
  $('#monthBudgetLabel').textContent = budgets.total > 0 ? `${money(plannedMonthTotal)} / ${money(budgets.total)}` : 'Sin límite';
  $('#monthMeter').style.width = budgets.total > 0 ? `${Math.min(100, plannedMonthTotal / budgets.total * 100)}%` : '0%';
  $('#monthMeter').style.background = budgets.total > 0 && plannedMonthTotal > budgets.total ? '#df777c' : '';
  const budgetLeft = budgets.total - plannedMonthTotal;
  const budgetRatio = budgets.total > 0 ? Math.min(100, plannedMonthTotal / budgets.total * 100) : 0;
  const budgetRemaining = $('#summaryBudgetRemaining');
  const budgetState = $('#summaryBudgetState');
  const budgetMeter = $('#summaryBudgetMeter');
  budgetRemaining.textContent = budgets.total > 0 ? (budgetLeft >= 0 ? money(budgetLeft) : `Excedido ${money(Math.abs(budgetLeft))}`) : '—';
  budgetState.textContent = budgets.total > 0 ? (budgetLeft >= 0 ? `${money(plannedMonthTotal)} de ${money(budgets.total)} utilizados` : 'Revisa el gasto programado') : 'Define tu límite mensual';
  budgetRemaining.closest('.summary-card').classList.toggle('is-over-budget', budgets.total > 0 && budgetLeft < 0);
  budgetRemaining.closest('.summary-card').classList.toggle('is-budget-tight', budgets.total > 0 && budgetLeft >= 0 && budgetRatio >= 80);
  budgetRemaining.closest('.summary-card').classList.toggle('is-budget-unset', budgets.total <= 0);
  budgetMeter.style.width = `${budgetRatio}%`;
  budgetMeter.style.background = budgetLeft < 0 ? '#df777c' : '';
  const activeCards = cards.filter((card) => card.active || installmentPlans.some((plan) => plan.cardId === card.id && plan.installmentsPaid < plan.months));
  $('#totalReceivable').textContent = money(people.reduce((sum, person) => sum + personDebt(person.id), 0));
  $('#totalCredit').textContent = money(people.reduce((sum, person) => sum + personCredit(person.id), 0));
  const statementCards = cards.filter((card) => card.active || cardExpenses(card.id).length > 0 || installmentPlans.some((plan) => plan.cardId === card.id && plan.installmentsPaid < plan.months));
  $('#totalToCover').textContent = money(statementCards.reduce((sum, card) => sum + Math.max(0, total(cardExpenses(card.id)) - collectedForCycle(card.id)) + installmentsDueForCycle(card.id), 0));
  const nextCard = activeCards.map((card) => ({ card, cutoff: nextCutoff(card) })).sort((a, b) => a.cutoff.date - b.cutoff.date)[0];
  const cutoffCountdown = $('#summaryCutoffCountdown');
  const cutoffName = $('#summaryCutoffName');
  const cutoffDate = $('#summaryCutoffDate');
  const cutoffDue = $('#summaryCutoffDue');
  const cutoffPanel = cutoffCountdown.closest('.summary-card');
  cutoffPanel.classList.toggle('has-no-cutoff', !nextCard);
  if (nextCard) {
    const { card, cutoff } = nextCard;
    const due = Math.max(0, total(cardExpenses(card.id)) - collectedForCycle(card.id)) + installmentsDueForCycle(card.id);
    cutoffCountdown.textContent = cutoff.days === 0 ? 'Hoy' : cutoff.days === 1 ? 'Mañana' : `En ${cutoff.days} días`;
    cutoffName.textContent = card.label;
    cutoffDate.textContent = new Intl.DateTimeFormat('es-MX', { weekday: 'long', day: 'numeric', month: 'long' }).format(cutoff.date);
    cutoffDue.textContent = money(due);
    cutoffPanel.style.setProperty('--cutoff-accent', card.color || 'var(--accent)');
  } else {
    cutoffCountdown.textContent = '—';
    cutoffName.textContent = 'Agrega o activa una tarjeta';
    cutoffDate.textContent = '';
    cutoffDue.textContent = '—';
  }
  $('#cardSummaries').innerHTML = activeCards.map((card) => {
    const list = cardExpenses(card);
    const cycle = cycleFor(card);
    const cutoff = nextCutoff(card);
    const spent = total(list);
    const received = collectedForCycle(card.id);
    const msiDue = installmentsDueForCycle(card.id);
    const toCover = Math.max(0, spent - received) + msiDue;
    const budgetSpend = spent + msiDue;
    return `<article class="summary-card"><div class="summary-top"><span>Ciclo activo · ${escapeHTML(card.label)}</span><span class="pill cutoff-countdown ${cutoff.days <= 3 ? 'cutoff-soon' : ''}" style="--cutoff-color:${card.color}">${cutoff.days === 0 ? 'Corte hoy' : cutoff.days === 1 ? 'Corte mañana' : `Corte en ${cutoff.days} días`}</span></div><strong>${money(spent)}</strong><small>${dateFormatter.format(cycle.start)} – ${dateFormatter.format(cycle.end)} · ${list.length} ${list.length === 1 ? 'compra' : 'compras'}</small><div class="card-money-lines"><span>Cubierto con abonos <b>${money(received)}</b></span>${msiDue > 0 ? `<span>Mensualidades MSI por cubrir <b>${money(msiDue)}</b></span>` : ''}<span>Te falta cubrir <b>${money(toCover)}</b></span></div>${card.cycleBudget > 0 ? `<div class="meter"><i style="width:${Math.min(100, budgetSpend / card.cycleBudget * 100)}%;background:${budgetSpend > card.cycleBudget ? '#df777c' : card.color}"></i></div><div class="meter-caption"><span>Tope del ciclo</span><b>${money(card.cycleBudget)}</b></div>` : ''}</article>`;
  }).join('');
  renderCutoffs();
  renderCardPayments();
}

function renderCutoffs() {
  const cutoffs = cards.filter((card) => card.active || installmentPlans.some((plan) => plan.cardId === card.id && plan.installmentsPaid < plan.months)).map((meta) => {
    let date = atMonth(today.getFullYear(), today.getMonth(), meta.cutoff);
    if (date < today) date = atMonth(today.getFullYear(), today.getMonth() + 1, meta.cutoff);
    return { key: meta.id, meta, date };
  }).sort((a, b) => a.date - b.date);
  $('#cutoffList').innerHTML = cutoffs.map(({ key, meta, date }) => {
    const list = cardExpenses(key);
    const left = Math.ceil((date - today) / 86400000);
    const when = left === 0 ? 'Es hoy' : left === 1 ? 'Mañana' : `En ${left} días`;
    return `<div class="cutoff-item"><div class="date-box"><small>${new Intl.DateTimeFormat('es-MX', { month: 'short' }).format(date).toUpperCase()}</small><b>${String(date.getDate()).padStart(2, '0')}</b></div><div class="cutoff-info"><b>${meta.label}</b><small>${when} · ${list.length} compras</small></div><div class="cutoff-amount">${money(Math.max(0, total(list) - collectedForCycle(key)) + installmentsDueForCycle(key))}<small>te falta cubrir</small></div></div>`;
  }).join('');
}

function monthOptions() {
  const select = $('#filterMonth');
  const existing = select.value;
  const months = new Set([monthKey(today), ...expenses.map((item) => monthKey(dateFromISO(item.date)))]);
  const sorted = [...months].sort().reverse();
  select.innerHTML = '<option value="all">Todos los meses</option>' + sorted.map((key) => {
    const [year, month] = key.split('-').map(Number);
    return `<option value="${key}">${monthFormatter.format(new Date(year, month - 1, 1))}</option>`;
  }).join('');
  select.value = sorted.includes(existing) ? existing : monthKey(today);
  const cardFilter = $('#filterCard');
  const cardValue = cardFilter.value;
  cardFilter.innerHTML = '<option value="all">Todas las tarjetas</option>' + cards.map((card) => `<option value="${card.id}">${escapeHTML(card.label)}${card.active ? '' : ' (baja)'}</option>`).join('');
  cardFilter.value = cards.some((card) => card.id === cardValue) || cardValue === 'all' ? cardValue : 'all';
  const personFilter = $('#filterPerson');
  const personValue = personFilter.value;
  personFilter.innerHTML = '<option value="all">Todas las personas</option>' + people.map((person) => `<option value="${person.id}">${escapeHTML(person.name)}${person.active ? '' : ' (baja)'}</option>`).join('');
  personFilter.value = people.some((person) => person.id === personValue) || personValue === 'all' ? personValue : 'all';
}

function filteredExpenses() {
  const query = $('#searchInput').value.trim().toLocaleLowerCase('es-MX');
  const card = $('#filterCard').value;
  const person = $('#filterPerson').value;
  const month = $('#filterMonth').value;
  return [...expenses].filter((expense) => {
    const matchesQuery = !query || `${expense.description} ${expense.category} ${expense.note || ''}`.toLocaleLowerCase('es-MX').includes(query);
    return matchesQuery && (card === 'all' || expense.card === card) && (person === 'all' || expense.personId === person) && (month === 'all' || monthKey(dateFromISO(expense.date)) === month);
  }).sort((a, b) => b.date.localeCompare(a.date) || (b.createdAt || '').localeCompare(a.createdAt || ''));
}

function renderRecords() {
  monthOptions();
  const expenseRows = filteredExpenses().map((item) => ({ ...item, kind: 'expense' }));
  const query = $('#searchInput').value.trim().toLocaleLowerCase('es-MX');
  const cardFilter = $('#filterCard').value;
  const personFilter = $('#filterPerson').value;
  const monthFilter = $('#filterMonth').value;
  const paymentRows = payments.filter((item) => {
    const person = personById(item.personId);
    const card = cardById(item.cardId);
    const matchesQuery = !query || `${person.name} abono ${item.note || ''} ${card.label}`.toLocaleLowerCase('es-MX').includes(query);
    return matchesQuery && (cardFilter === 'all' || item.cardId === cardFilter) && (personFilter === 'all' || item.personId === personFilter) && (monthFilter === 'all' || monthKey(dateFromISO(item.date)) === monthFilter);
  }).map((item) => ({ ...item, kind: 'payment' }));
  const list = [...expenseRows, ...paymentRows].sort((a, b) => b.date.localeCompare(a.date) || (b.createdAt || '').localeCompare(a.createdAt || ''));
  $('#recordCount').textContent = list.length;
  $('#recordList').innerHTML = list.map((expense) => {
    if (expense.kind === 'payment') {
      const person = personById(expense.personId);
      const card = cardById(expense.cardId);
      return `<article class="record-row payment-row" data-kind="payment" data-id="${expense.id}"><span class="payment-icon" aria-label="Abono recibido">↙</span><div class="record-description"><b>Abono de ${escapeHTML(person.name)}</b><small>${escapeHTML(expense.note || 'Pago recibido')} · ${escapeHTML(card.label)}</small></div><div class="record-date">${dateFormatter.format(dateFromISO(expense.date))}</div><div class="record-person">${escapeHTML(person.name)}</div><span class="tag payment-tag">Abono recibido</span><div class="record-amount payment-amount">+ ${money(expense.amount)}</div><button class="row-menu" data-action="menu" aria-label="Opciones para el abono de ${escapeHTML(person.name)}">···</button></article>`;
    }
    const meta = cardById(expense.card);
    const label = `${expense.paid ? 'Pagado' : 'Pendiente'}: ${expense.description}`;
    const person = personById(expense.personId);
    return `<article class="record-row" data-kind="expense" data-id="${expense.id}"><button class="record-check ${expense.paid ? 'checked' : ''}" data-action="paid" aria-label="${label}" title="Marcar ${expense.paid ? 'pendiente' : 'pagado'}">✓</button><div class="record-description"><b>${escapeHTML(expense.description)}</b><small>${escapeHTML(expense.category)} · ${escapeHTML(person.name)}${expense.receivable ? ' · por cobrar' : ' · gasto propio'}${expense.note ? ` · ${escapeHTML(expense.note)}` : ''}</small></div><div class="record-date">${dateFormatter.format(dateFromISO(expense.date))}</div><div class="record-person">${escapeHTML(person.name)}</div><span class="tag" style="color:${meta.color};background:${meta.color}1c"><i class="tag-dot"></i>${escapeHTML(meta.label)}</span><div class="record-amount">${money(expense.amount)}</div><button class="row-menu" data-action="menu" aria-label="Opciones para ${escapeHTML(expense.description)}">···</button></article>`;
  }).join('');
  $('#emptyState').hidden = list.length !== 0;
  $('#recordList').hidden = list.length === 0;
  $('#chartEmpty').hidden = expenses.some(inCurrentMonth);
  drawChart();
}

function escapeHTML(value = '') { return String(value).replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]); }

function drawChart() {
  const canvas = $('#spendChart');
  if (!canvas || !canvas.clientWidth) return;
  const width = canvas.clientWidth;
  const height = canvas.clientHeight;
  const ratio = window.devicePixelRatio || 1;
  canvas.width = width * ratio;
  canvas.height = height * ratio;
  const ctx = canvas.getContext('2d');
  ctx.scale(ratio, ratio);
  const style = getComputedStyle(document.documentElement);
  const gridColor = style.getPropertyValue('--line').trim();
  const mutedColor = style.getPropertyValue('--muted').trim();
  const mode = $('#chartMode').value;
  const monthly = expenses.filter(inCurrentMonth);
  const pad = { left: 39, right: 8, top: 10, bottom: 24 };
  const plotW = width - pad.left - pad.right;
  const plotH = height - pad.top - pad.bottom;
  ctx.clearRect(0, 0, width, height);
  ctx.font = '10px DM Sans, sans-serif';
  ctx.textBaseline = 'middle';
  if (mode !== 'daily') {
    const personList = people.filter((person) => person.active);
    const labels = mode === 'category' ? ['Comida', 'Supermercado', 'Transporte', 'Hogar', 'Salud', 'Entretenimiento', 'Servicios', 'Compras', 'Otro'] : personList.map((person) => person.name);
    const values = labels.map((label, index) => total(monthly.filter((item) => (mode === 'category' ? item.category === label : item.personId === personList[index]?.id))));
    const max = Math.max(1, ...values);
    const gap = plotW / labels.length;
    const barW = Math.min(mode === 'person' ? 58 : 25, gap * .58);
    ctx.strokeStyle = gridColor; ctx.fillStyle = mutedColor; ctx.lineWidth = 1;
    for (let tick = 0; tick <= 3; tick++) {
      const y = pad.top + plotH - plotH * tick / 3;
      ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(width - pad.right, y); ctx.stroke();
      ctx.textAlign = 'right'; ctx.fillText(tick === 0 ? '0' : shortMoney(max * tick / 3), pad.left - 7, y);
    }
    labels.forEach((label, index) => {
      const x = pad.left + gap * index + (gap - barW) / 2;
      const barH = values[index] / max * (plotH - 4);
      ctx.fillStyle = mode === 'category' ? (index % 2 ? '#a89be9' : '#8270e4') : cardPalette[index % cardPalette.length];
      roundRect(ctx, x, pad.top + plotH - barH, barW, Math.max(2, barH), 5); ctx.fill();
      ctx.fillStyle = mutedColor; ctx.textAlign = 'center';
      const shortLabel = mode === 'category' ? ({ Supermercado: 'Super', Entretenimiento: 'Ocio', Transporte: 'Transp.' }[label] || label) : label.length > 10 ? `${label.slice(0, 9)}…` : label;
      ctx.fillText(shortLabel, pad.left + gap * index + gap / 2, height - 9);
    });
    $('#chartLegend').innerHTML = mode === 'category' ? '<span class="legend-item"><i class="legend-dot" style="background:#8270e4"></i>Gasto por categoría</span>' : personList.map((person, index) => `<span class="legend-item"><i class="legend-dot" style="background:${cardPalette[index % cardPalette.length]}"></i>${escapeHTML(person.name)}</span>`).join('');
    return;
  }
  const dayCount = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const values = Array.from({ length: dayCount }, (_, i) => total(monthly.filter((item) => dateFromISO(item.date).getDate() === i + 1)));
  const max = Math.max(1, ...values);
  ctx.strokeStyle = gridColor; ctx.fillStyle = mutedColor; ctx.lineWidth = 1;
  for (let tick = 0; tick <= 3; tick++) {
    const y = pad.top + plotH - plotH * tick / 3;
    ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(width - pad.right, y); ctx.stroke();
    ctx.textAlign = 'right'; ctx.fillText(tick === 0 ? '0' : shortMoney(max * tick / 3), pad.left - 7, y);
  }
  const points = values.map((value, index) => ({ x: pad.left + index / Math.max(1, dayCount - 1) * plotW, y: pad.top + plotH - (value / max) * (plotH - 5) }));
  const gradient = ctx.createLinearGradient(0, pad.top, 0, height - pad.bottom);
  gradient.addColorStop(0, 'rgba(130,112,228,.26)'); gradient.addColorStop(1, 'rgba(130,112,228,0)');
  ctx.beginPath(); ctx.moveTo(points[0].x, pad.top + plotH); points.forEach((point) => ctx.lineTo(point.x, point.y)); ctx.lineTo(points.at(-1).x, pad.top + plotH); ctx.closePath(); ctx.fillStyle = gradient; ctx.fill();
  ctx.beginPath(); points.forEach((point, index) => index ? ctx.lineTo(point.x, point.y) : ctx.moveTo(point.x, point.y)); ctx.strokeStyle = '#8270e4'; ctx.lineWidth = 2.2; ctx.lineJoin = 'round'; ctx.lineCap = 'round'; ctx.stroke();
  ctx.fillStyle = mutedColor; ctx.textAlign = 'center';
  [1, 8, 15, 22, dayCount].forEach((day) => { const x = pad.left + (day - 1) / Math.max(1, dayCount - 1) * plotW; ctx.fillText(String(day), x, height - 9); });
  $('#chartLegend').innerHTML = '<span class="legend-item"><i class="legend-dot" style="background:#8270e4"></i>Gasto diario</span><span class="legend-item"><i class="legend-dot" style="background:#c8c0f4"></i>Mes en curso</span>';
}
function shortMoney(value) { return value >= 10000 ? `$${Math.round(value / 1000)}k` : value >= 1000 ? `$${(value / 1000).toFixed(1)}k` : `$${Math.round(value)}`; }
function roundRect(ctx, x, y, w, h, radius) { ctx.beginPath(); ctx.roundRect(x, y, w, h, Math.min(radius, w / 2, h / 2)); }
function render() { renderSummary(); renderDebtBreakdown(); renderInstallments(); renderFinancialHealth(); renderIncomePanel(); renderDecisionKpis(); renderPersonKpis(); renderRecords(); }
function showToast(message) { const toast = $('#toast'); toast.textContent = message; toast.classList.add('show'); clearTimeout(toastTimer); toastTimer = setTimeout(() => toast.classList.remove('show'), 2300); }

async function connectExcelFile() {
  if (!window.showSaveFilePicker) {
    setExcelStatus('Este navegador no permite sincronizar directamente; usa importar/descargar respaldo');
    $('#excelFileInput').click();
    return;
  }
  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: 'gfp-base-datos.xlsx',
      types: [{ description: 'Base de datos de Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
    });
    const file = await handle.getFile();
    let imported = null;
    if (file.size > 0) {
      imported = await readExcelDatabase(file);
      if (!imported) { showToast('Ese libro no tiene las hojas de la base de datos GFP'); return; }
      if (window.confirm('Se encontró una base Excel. ¿Cargar sus datos en GFP? Aceptar: reemplazar los datos actuales por los del libro. Cancelar: conservar los datos actuales y escribirlos en el libro.')) {
        applyExcelData(imported);
        await saveAllCollections();
      }
    }
    excelHandle = handle;
    await putDatabaseValue('excelFileHandle', handle);
    await saveExcelHandle();
    render();
    showToast(imported ? 'Libro Excel conectado' : 'Libro Excel creado y conectado');
  } catch (error) {
    if (error?.name !== 'AbortError') { console.error(error); setExcelStatus('No se pudo conectar el Excel; la base local sigue activa'); showToast('No se pudo conectar el archivo Excel'); }
  }
}
async function importExcelFile(file) {
  try {
    const data = await readExcelDatabase(file);
    if (!data) { showToast('El archivo no contiene las hojas de GFP'); return; }
    if (!window.confirm('¿Reemplazar los datos locales con la información de este Excel?')) return;
    applyExcelData(data);
    await saveAllCollections();
    if (excelHandle) await saveExcelHandle();
    render();
    renderManagedCards();
    renderManagedPeople();
    showToast('Base Excel importada');
  } catch (error) { console.error(error); showToast('No se pudo leer ese archivo Excel'); }
}

function openExpense(expense = null) {
  $('#expenseForm').reset(); $('#formError').textContent = '';
  const activeCards = cards.filter((card) => card.active);
  const activePeople = people.filter((person) => person.active);
  $('#card').innerHTML = activeCards.map((card) => `<option value="${card.id}">${escapeHTML(card.label)} · corte ${String(card.cutoff).padStart(2, '0')}</option>`).join('');
  $('#person').innerHTML = activePeople.map((person) => `<option value="${person.id}">${escapeHTML(person.name)}</option>`).join('');
  $('#card').disabled = activeCards.length === 0;
  if (!activeCards.length || !activePeople.length) { showToast('Agrega tarjetas y personas en Presupuestos y tarjetas'); renderManagedCards(); renderManagedPeople(); $('#settingsDialog').showModal(); return; }
  $('#formTitle').textContent = expense ? 'Editar gasto' : 'Agregar gasto';
  $('#expenseId').value = expense?.id || '';
  $('#description').value = expense?.description || '';
  $('#amount').value = expense?.amount || '';
  $('#date').value = expense?.date || isoDate(today);
  $('#card').value = activeCards.some((card) => card.id === expense?.card) ? expense.card : activeCards[0].id;
  $('#person').value = activePeople.some((person) => person.id === expense?.personId) ? expense.personId : activePeople[0].id;
  $('#category').value = expense?.category || 'Comida';
  $('#note').value = expense?.note || '';
  $('#paid').checked = Boolean(expense?.paid);
  $('#expenseDialog').showModal();
  setTimeout(() => $('#description').focus(), 80);
}

$('#addButton').addEventListener('click', () => openExpense());
$('#quickAddFloating').addEventListener('click', () => openExpense());
$('#emptyAdd').addEventListener('click', () => openExpense());
$('#closeDialog').addEventListener('click', () => $('#expenseDialog').close());
$('#expenseForm').addEventListener('submit', (event) => {
  event.preventDefault();
  const amount = Number($('#amount').value);
  if (!Number.isFinite(amount) || amount <= 0) { $('#formError').textContent = 'Ingresa un monto mayor a cero.'; return; }
  const id = $('#expenseId').value;
  const original = expenses.find((item) => item.id === id);
  const personId = $('#person').value;
  const expense = { id: id || crypto.randomUUID(), description: $('#description').value.trim(), amount, date: $('#date').value, card: $('#card').value, personId, receivable: original?.personId === personId ? original.receivable : personById(personId).tracksDebt, category: $('#category').value, note: $('#note').value.trim(), paid: $('#paid').checked, createdAt: original?.createdAt || new Date().toISOString() };
  expenses = id ? expenses.map((item) => item.id === id ? expense : item) : [...expenses, expense];
  persist(); render(); $('#expenseDialog').close(); showToast(id ? 'Gasto actualizado' : 'Gasto guardado');
});

function openRecordActions(row) {
  const isPayment = row.dataset.kind === 'payment';
  const record = isPayment ? payments.find((item) => item.id === row.dataset.id) : expenses.find((item) => item.id === row.dataset.id);
  if (!record) return;
  recordActionTarget = { kind: isPayment ? 'payment' : 'expense', id: record.id };
  $('#recordActionTitle').textContent = isPayment ? `Abono de ${personById(record.personId).name}` : record.description;
  $('#recordActionOptions').innerHTML = `${isPayment ? '' : '<button class="record-action-choice" type="button" data-record-command="edit"><span class="record-action-icon">✎</span><span><b>Editar gasto</b><small>Modificar los datos del registro</small></span></button>'}<button class="record-action-choice delete-choice" type="button" data-record-command="delete"><span class="record-action-icon">⌫</span><span><b>${isPayment ? 'Eliminar abono' : 'Borrar gasto'}</b><small>${isPayment ? 'Quitar este abono de tus cuentas' : 'Eliminar este gasto de tus registros'}</small></span></button>`;
  $('#recordActionsDialog').showModal();
}
$('#recordList').addEventListener('click', (event) => {
  const button = event.target.closest('[data-action]');
  if (!button) return;
  const row = button.closest('.record-row');
  if (button.dataset.action === 'menu') { openRecordActions(row); return; }
  if (row?.dataset.kind === 'payment') return;
  const expense = expenses.find((item) => item.id === row?.dataset.id);
  if (expense && button.dataset.action === 'paid') {
    expense.paid = !expense.paid; persist(); renderRecords(); showToast(expense.paid ? 'Marcado como pagado' : 'Marcado como pendiente');
  }
});
$('#closeRecordActions').addEventListener('click', () => $('#recordActionsDialog').close());
$('#recordActionOptions').addEventListener('click', (event) => {
  const button = event.target.closest('[data-record-command]');
  const target = recordActionTarget;
  if (!button || !target) return;
  const command = button.dataset.recordCommand;
  const item = target.kind === 'payment' ? payments.find((record) => record.id === target.id) : expenses.find((record) => record.id === target.id);
  $('#recordActionsDialog').close();
  recordActionTarget = null;
  if (!item) return;
  if (command === 'edit' && target.kind === 'expense') { openExpense(item); return; }
  const label = target.kind === 'payment' ? `el abono de ${money(item.amount)}` : `“${item.description}”`;
  if (!window.confirm(`¿Eliminar ${label} de tus registros?`)) return;
  if (target.kind === 'payment') { payments = payments.filter((record) => record.id !== target.id); persistPayments(); }
  else { expenses = expenses.filter((record) => record.id !== target.id); persist(); }
  render(); showToast(target.kind === 'payment' ? 'Abono eliminado' : 'Gasto eliminado');
});

['searchInput', 'filterCard', 'filterPerson', 'filterMonth'].forEach((selector) => $(`#${selector}`).addEventListener(selector === 'searchInput' ? 'input' : 'change', renderRecords));
['kpiMonth', 'kpiCard'].forEach((selector) => $(`#${selector}`).addEventListener('change', renderPersonKpis));
['decisionMonth', 'decisionCard', 'decisionPerson'].forEach((selector) => $(`#${selector}`).addEventListener('change', renderDecisionKpis));
function updateSectionNavigation() {
  const sectionIds = ['inicio', 'analisis', 'finanzas', 'registros'];
  let current = sectionIds[0];
  for (const id of sectionIds) {
    const section = document.getElementById(id);
    if (section && section.getBoundingClientRect().top <= window.innerHeight * .36) current = id;
  }
  document.querySelectorAll('.section-nav a, .mobile-nav a').forEach((link) => link.classList.toggle('active', link.hash === `#${current}`));
}
window.addEventListener('scroll', updateSectionNavigation, { passive: true });
window.addEventListener('resize', updateSectionNavigation);
updateSectionNavigation();
$('#chartMode').addEventListener('change', drawChart);
window.addEventListener('resize', drawChart);
$('#themeToggle').addEventListener('click', () => {
  const current = localStorage.getItem('saldo-theme') || 'system';
  setTheme(current === 'dark' ? 'light' : 'dark');
});
const savedTheme = localStorage.getItem('saldo-theme') || 'system';
setTheme(savedTheme);
$('#settingsButton').addEventListener('click', () => {
  $('#budgetTotal').value = budgets.total || '';
  $('#profileName').value = userProfile.name || '';
  renderManagedCards();
  renderManagedPeople();
  hideCardEditor();
  hidePersonEditor();
  $('#settingsDialog').showModal();
});
$('#healthBudgetButton').addEventListener('click', () => $('#settingsButton').click());
$('#closeSettings').addEventListener('click', () => $('#settingsDialog').close());
$('#settingsForm').addEventListener('submit', (event) => {
  event.preventDefault();
  budgets = { ...budgets, total: Math.max(0, Number($('#budgetTotal').value) || 0) };
  userProfile = { name: $('#profileName').value.trim().slice(0, 35) };
  persistBudgets(); persistProfile(); $('#settingsDialog').close(); render(); showToast(userProfile.name ? `¡Hola, ${userProfile.name}! Tu perfil se guardó` : 'Cambios guardados');
});

function openInstallment(plan = null) {
  const activeCards = cards.filter((card) => card.active || card.id === plan?.cardId);
  const activePeople = people.filter((person) => person.active || person.id === plan?.personId);
  if (!activeCards.length || !activePeople.length) {
    showToast('Agrega una tarjeta y una persona antes de registrar una compra MSI');
    renderManagedCards(); renderManagedPeople(); $('#settingsDialog').showModal();
    return;
  }
  $('#installmentForm').reset();
  $('#installmentFormTitle').textContent = plan ? 'Editar compra MSI' : 'Registrar MSI';
  $('#installmentId').value = plan?.id || '';
  $('#installmentDescription').value = plan?.description || '';
  $('#installmentAmount').value = plan?.amount || '';
  $('#installmentDate').value = plan?.date || isoDate(today);
  $('#installmentMonths').value = plan?.months || 12;
  $('#installmentPaid').value = plan?.installmentsPaid || 0;
  $('#installmentNote').value = plan?.note || '';
  $('#installmentError').textContent = '';
  $('#installmentCard').innerHTML = activeCards.map((card) => `<option value="${escapeHTML(card.id)}">${escapeHTML(card.label)} · corte ${String(card.cutoff).padStart(2, '0')}${card.active ? '' : ' · baja'}</option>`).join('');
  $('#installmentPerson').innerHTML = activePeople.map((person) => `<option value="${escapeHTML(person.id)}">${escapeHTML(person.name)}${person.active ? '' : ' · baja'}</option>`).join('');
  $('#installmentCard').value = activeCards.some((card) => card.id === plan?.cardId) ? plan.cardId : activeCards[0].id;
  $('#installmentPerson').value = activePeople.some((person) => person.id === plan?.personId) ? plan.personId : activePeople[0].id;
  $('#installmentDialog').showModal();
  setTimeout(() => $('#installmentDescription').focus(), 80);
}
$('#addInstallmentButton').addEventListener('click', () => openInstallment());
$('#emptyInstallmentAdd').addEventListener('click', () => openInstallment());
$('#closeInstallment').addEventListener('click', () => $('#installmentDialog').close());
$('#installmentForm').addEventListener('submit', (event) => {
  event.preventDefault();
  const amount = Number($('#installmentAmount').value);
  const months = Math.floor(Number($('#installmentMonths').value));
  const installmentsPaid = Math.floor(Number($('#installmentPaid').value));
  if (!Number.isFinite(amount) || amount <= 0) { $('#installmentError').textContent = 'Ingresa un monto mayor a cero.'; return; }
  if (!Number.isInteger(months) || months < 2 || months > 120) { $('#installmentError').textContent = 'El plazo debe estar entre 2 y 120 meses.'; return; }
  if (!Number.isInteger(installmentsPaid) || installmentsPaid < 0 || installmentsPaid > months) { $('#installmentError').textContent = 'Las mensualidades pagadas deben estar entre cero y el plazo.'; return; }
  const id = $('#installmentId').value;
  const original = installmentPlans.find((item) => item.id === id);
  const plan = { id: id || crypto.randomUUID(), date: $('#installmentDate').value, description: $('#installmentDescription').value.trim(), amount, cardId: $('#installmentCard').value, personId: $('#installmentPerson').value, months, installmentsPaid, note: $('#installmentNote').value.trim(), createdAt: original?.createdAt || new Date().toISOString() };
  installmentPlans = id ? installmentPlans.map((item) => item.id === id ? plan : item) : [...installmentPlans, plan];
  persistInstallments(); render(); $('#installmentDialog').close(); showToast(id ? 'Compra MSI actualizada' : 'Compra MSI guardada');
});
$('#installmentList').addEventListener('click', (event) => {
  const button = event.target.closest('[data-msi-action]');
  if (!button) return;
  const plan = installmentPlans.find((item) => item.id === button.dataset.msiId);
  if (!plan) return;
  if (button.dataset.msiAction === 'edit') openInstallment(plan);
  else if (button.dataset.msiAction === 'pay' && plan.installmentsPaid < plan.months) {
    plan.installmentsPaid += 1; persistInstallments(); render(); showToast(`Mensualidad ${plan.installmentsPaid} de ${plan.months} registrada`);
  } else if (button.dataset.msiAction === 'delete' && window.confirm(`¿Eliminar la compra MSI “${plan.description}”?`)) {
    installmentPlans = installmentPlans.filter((item) => item.id !== plan.id); persistInstallments(); render(); showToast('Compra MSI eliminada');
  }
});

function openPayment() {
  if (!cards.length || !people.length) {
    showToast('Agrega al menos una persona y una tarjeta primero');
    renderManagedCards(); renderManagedPeople(); hideCardEditor(); hidePersonEditor(); $('#settingsDialog').showModal();
    return;
  }
  $('#paymentCard').innerHTML = cards.map((card) => `<option value="${card.id}">${escapeHTML(card.label)} · corte ${String(card.cutoff).padStart(2, '0')}${card.active ? '' : ' · baja'}</option>`).join('');
  $('#paymentForm').reset();
  $('#paymentDate').value = isoDate(today);
  $('#paymentCard').value = cards.find((card) => card.active)?.id || cards[0].id;
  updatePaymentPeople();
  $('#paymentDialog').showModal();
}
function updatePaymentPeople() {
  const cardId = $('#paymentCard').value;
  $('#paymentPerson').innerHTML = people.map((person) => {
    const debt = personDebt(person.id, cardId);
    const credit = personCredit(person.id, cardId);
    const label = debt > 0 ? `pendiente ${money(debt)}` : credit > 0 ? `a favor ${money(credit)}` : 'sin saldo';
    return `<option value="${person.id}">${escapeHTML(person.name)} · ${label}${person.active ? '' : ' · baja'}</option>`;
  }).join('');
  $('#paymentPerson').disabled = people.length === 0;
}
$('#paymentButton').addEventListener('click', openPayment);
$('#closePayment').addEventListener('click', () => $('#paymentDialog').close());
$('#paymentCard').addEventListener('change', updatePaymentPeople);
  $('#paymentForm').addEventListener('submit', (event) => {
  event.preventDefault();
  const amount = Number($('#paymentAmount').value);
  if (!Number.isFinite(amount) || amount <= 0) { showToast('Ingresa un abono mayor a cero'); return; }
  payments.push({ id: crypto.randomUUID(), personId: $('#paymentPerson').value, cardId: $('#paymentCard').value, amount, date: $('#paymentDate').value, note: $('#paymentNote').value.trim(), createdAt: new Date().toISOString() });
  persistPayments(); $('#paymentDialog').close(); render(); showToast('Abono guardado; saldo pendiente actualizado');
});

$('#bankPaymentCards').addEventListener('click', (event) => {
  const add = event.target.closest('.add-bank-payment');
  if (add) { openBankPayment(add.dataset.cardId, add.dataset.statementEnd); return; }
  const remove = event.target.closest('[data-bank-payment-id]');
  if (!remove) return;
  cardPayments = cardPayments.filter((payment) => payment.id !== remove.dataset.bankPaymentId);
  persistCardPayments(); render(); showToast('Pago al banco eliminado');
});
$('#addBankPaymentButton').addEventListener('click', () => openBankPayment());
$('#closeBankPayment').addEventListener('click', () => $('#bankPaymentDialog').close());
$('#bankPaymentCard').addEventListener('change', (event) => updateBankStatementOptions(event.target.value));
$('#bankPaymentForm').addEventListener('submit', (event) => {
  event.preventDefault();
  const amount = Number($('#bankPaymentAmount').value);
  if (!Number.isFinite(amount) || amount <= 0) { showToast('Ingresa un pago mayor a cero'); return; }
  cardPayments.push({ id: crypto.randomUUID(), cardId: $('#bankPaymentCard').value, statementEnd: $('#bankStatementEnd').value, amount, date: $('#bankPaymentDate').value, note: $('#bankPaymentNote').value.trim(), createdAt: new Date().toISOString() });
  persistCardPayments(); $('#bankPaymentDialog').close(); render(); showToast('Pago a la tarjeta guardado');
});
function openIncome() {
  $('#incomeForm').reset();
  $('#incomeDate').value = isoDate(today);
  $('#incomeDialog').showModal();
  $('#incomeSource').focus();
}
$('#addIncomeButton').addEventListener('click', openIncome);
$('#incomeEmptyAdd').addEventListener('click', openIncome);
$('#closeIncome').addEventListener('click', () => $('#incomeDialog').close());
$('#incomeForm').addEventListener('submit', (event) => {
  event.preventDefault();
  const amount = Number($('#incomeAmount').value);
  const source = $('#incomeSource').value.trim();
  if (!source || !Number.isFinite(amount) || amount <= 0) { showToast('Completa la fuente y un monto mayor a cero'); return; }
  incomes.push({ id: crypto.randomUUID(), date: $('#incomeDate').value, source, amount, note: $('#incomeNote').value.trim(), createdAt: new Date().toISOString() });
  persistIncomes(); $('#incomeDialog').close(); render(); showToast('Ingreso registrado');
});
$('#incomeList').addEventListener('click', (event) => {
  const button = event.target.closest('[data-income-id]');
  if (!button) return;
  incomes = incomes.filter((income) => income.id !== button.dataset.incomeId);
  persistIncomes(); render(); showToast('Ingreso eliminado');
});

function renderManagedCards() {
  $('#managedCards').innerHTML = cards.map((card) => `<div class="managed-card ${card.active ? '' : 'archived'}"><span class="managed-card-dot" style="background:${card.color}"></span><span class="managed-card-name"><b>${escapeHTML(card.label)}</b><small>Corte ${String(card.cutoff).padStart(2, '0')} · ${card.paymentDueDay ? `pago ${String(card.paymentDueDay).padStart(2, '0')} ${Number(card.paymentDueOffset) ? 'del siguiente mes' : 'del mes'}` : 'fecha límite sin definir'}${card.active ? '' : ' · dada de baja'}</small></span><button class="managed-edit" type="button" data-card-action="edit" data-card-id="${card.id}" aria-label="Editar ${escapeHTML(card.label)}">Editar</button>${card.active ? `<button class="managed-disable" type="button" data-card-action="disable" data-card-id="${card.id}">Dar de baja</button>` : ''}</div>`).join('') || '<p class="no-cards">Aún no tienes tarjetas. Agrega la primera para registrar compras.</p>';
}
function hideCardEditor() { $('#cardEditor').hidden = true; $('#cardId').value = ''; $('#cardName').value = ''; $('#cardCutoff').value = ''; $('#cardPaymentDay').value = ''; $('#cardPaymentOffset').value = '1'; $('#cardBudget').value = ''; }
function renderManagedPeople() {
  $('#managedPeople').innerHTML = people.map((person) => {
    const debt = personDebt(person.id);
    const credit = personCredit(person.id);
    const status = person.tracksDebt ? `Pendiente por cobrar: ${money(debt)}${credit > 0 ? ` · saldo a favor: ${money(credit)}` : ''}` : `Gastos propios${debt > 0 ? ` · deuda anterior: ${money(debt)}` : ''}${credit > 0 ? ` · saldo a favor: ${money(credit)}` : ''}`;
    return `<div class="managed-card ${person.active ? '' : 'archived'}"><span class="managed-card-dot" style="background:${cardPalette[people.indexOf(person) % cardPalette.length]}"></span><span class="managed-card-name"><b>${escapeHTML(person.name)}</b><small>${status}${person.active ? '' : ' · dada de baja'}</small></span><button class="managed-edit" type="button" data-person-action="edit" data-person-id="${person.id}">Editar</button>${person.active ? `<button class="managed-disable" type="button" data-person-action="disable" data-person-id="${person.id}">Dar de baja</button>` : ''}</div>`;
  }).join('') || '<p class="no-cards">Aún no tienes personas. Agrega a quienes registrarán gastos.</p>';
}
function hidePersonEditor() { $('#personEditor').hidden = true; $('#personId').value = ''; $('#personName').value = ''; $('#personOwes').value = 'true'; }
function showPersonEditor(person = null) {
  $('#personEditor').hidden = false;
  $('#personEditorTitle').textContent = person ? 'Editar persona' : 'Nueva persona';
  $('#personId').value = person?.id || '';
  $('#personName').value = person?.name || '';
  $('#personOwes').value = person?.tracksDebt ? 'true' : 'false';
  $('#personName').focus();
}
function showCardEditor(card = null) {
  $('#cardEditor').hidden = false;
  $('#cardEditorTitle').textContent = card ? 'Editar tarjeta' : 'Nueva tarjeta';
  $('#cardId').value = card?.id || '';
  $('#cardName').value = card?.label || '';
  $('#cardCutoff').value = card?.cutoff || '';
  $('#cardPaymentDay').value = card?.paymentDueDay || '';
  $('#cardPaymentOffset').value = String(card?.paymentDueOffset ?? 1);
  $('#cardBudget').value = card?.cycleBudget || '';
  $('#cardName').focus();
}
$('#newCardButton').addEventListener('click', () => showCardEditor());
$('#cancelCardEdit').addEventListener('click', hideCardEditor);
$('#saveCardButton').addEventListener('click', () => {
  const label = $('#cardName').value.trim();
  const cutoff = Number($('#cardCutoff').value);
  const paymentDueDay = Number($('#cardPaymentDay').value) || 0;
  const paymentDueOffset = Number($('#cardPaymentOffset').value) === 0 ? 0 : 1;
  const cycleBudget = Math.max(0, Number($('#cardBudget').value) || 0);
  const id = $('#cardId').value;
  if (!label) { $('#cardName').focus(); showToast('Escribe el nombre de la tarjeta'); return; }
  if (!Number.isInteger(cutoff) || cutoff < 1 || cutoff > 28) { $('#cardCutoff').focus(); showToast('El corte debe ser un día entre 1 y 28'); return; }
  if (paymentDueDay && (!Number.isInteger(paymentDueDay) || paymentDueDay < 1 || paymentDueDay > 31)) { $('#cardPaymentDay').focus(); showToast('El día límite debe estar entre 1 y 31'); return; }
  const existing = cards.find((card) => card.id === id);
  const card = { id: id || crypto.randomUUID(), label, cutoff, paymentDueDay, paymentDueOffset, cycleBudget, color: existing?.color || cardPalette[cards.length % cardPalette.length], active: true };
  cards = id ? cards.map((item) => item.id === id ? card : item) : [...cards, card];
  persistCards(); hideCardEditor(); renderManagedCards(); render(); showToast(id ? 'Tarjeta actualizada' : 'Tarjeta agregada');
});
$('#managedCards').addEventListener('click', (event) => {
  const button = event.target.closest('[data-card-action]');
  if (!button) return;
  const card = cards.find((item) => item.id === button.dataset.cardId);
  if (!card) return;
  if (button.dataset.cardAction === 'edit') showCardEditor(card);
  else if (button.dataset.cardAction === 'disable' && window.confirm(`¿Dar de baja “${card.label}”? Se conservarán sus gastos anteriores.`)) {
    card.active = false; persistCards(); renderManagedCards(); render(); showToast('Tarjeta dada de baja; historial conservado');
  }
});
$('#newPersonButton').addEventListener('click', () => showPersonEditor());
$('#cancelPersonEdit').addEventListener('click', hidePersonEditor);
$('#savePersonButton').addEventListener('click', () => {
  const name = $('#personName').value.trim();
  const id = $('#personId').value;
  if (!name) { $('#personName').focus(); showToast('Escribe el nombre de la persona'); return; }
  const existing = people.find((person) => person.id === id);
  const person = { id: id || crypto.randomUUID(), name, tracksDebt: $('#personOwes').value === 'true', active: true };
  people = id ? people.map((item) => item.id === id ? person : item) : [...people, person];
  persistPeople(); hidePersonEditor(); renderManagedPeople(); render(); showToast(id ? 'Persona actualizada' : 'Persona agregada');
});
$('#managedPeople').addEventListener('click', (event) => {
  const button = event.target.closest('[data-person-action]');
  if (!button) return;
  const person = people.find((item) => item.id === button.dataset.personId);
  if (!person) return;
  if (button.dataset.personAction === 'edit') showPersonEditor(person);
  else if (button.dataset.personAction === 'disable' && window.confirm(`¿Dar de baja a “${person.name}”? Se conservarán sus gastos y abonos.`)) {
    person.active = false; persistPeople(); renderManagedPeople(); render(); showToast('Persona dada de baja; historial conservado');
  }
});

$('#connectExcelButton').addEventListener('click', connectExcelFile);
$('#exportExcelButton').addEventListener('click', () => {
  try { downloadExcel(); showToast('Respaldo Excel descargado'); }
  catch (error) { console.error(error); showToast('No se pudo crear el respaldo Excel'); }
});
$('#importExcelButton').addEventListener('click', () => $('#excelFileInput').click());
$('#excelFileInput').addEventListener('change', async (event) => {
  const file = event.target.files?.[0];
  if (file) await importExcelFile(file);
  event.target.value = '';
});

document.body.classList.add('database-loading');
initDatabase().then(() => {
  render();
  document.body.classList.remove('database-loading');
}).catch((error) => {
  console.error('IndexedDB no disponible, se usará el almacenamiento local del navegador', error);
  database = null;
  normalizeState();
  render();
  document.body.classList.remove('database-loading');
  setExcelStatus('Guardado local del navegador activo · Excel disponible para respaldo');
});
if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('./sw.js').catch((error) => console.info('Service worker no disponible', error));
