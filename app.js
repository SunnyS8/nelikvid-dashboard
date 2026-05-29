// ============ ГЛОБАЛЬНОЕ СОСТОЯНИЕ ============
let stockRaw = [];
let salesRaw = [];
let ropList = [];
let report = [];
let aggregates = { kpi: {}, bySklad: [], byBrand: [], statusCounts: [] };
let state = { rows: [], charts: {} };

const fmt = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 });
const money = new Intl.NumberFormat('ru-RU', { style: 'currency', currency: 'RUB', maximumFractionDigits: 0 });

// ============ ПАРСИНГ EXCEL ============

/**
 * Читает Excel файл и возвращает рабочую книгу
 */
function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        resolve(workbook);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Парсит файл остатков из 1С
 * Ищет строки с названием склада, потом строки данных с артикулом и остатком
 */
function parseStock(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  
  let currentSklad = '';
  const result = [];
  
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length < 3) continue;
    
    // Строка склада: название в col[0], col[1] пусто, col[2] пусто
    if (row[0] && !row[1] && !row[2]) {
      currentSklad = row[0].toString().trim();
      continue;
    }
    
    // Строка с данными: col[0] - артикул/код, col[2] - наименование, col[15] - остаток
    const code = row[0] ? String(row[0]).trim() : '';
    const fullName = row[2] ? String(row[2]).trim() : '';
    const stock = row[15] ? parseFloat(row[15]) : 0;
    
    if (code && fullName && currentSklad) {
      result.push({
        Склад: currentSklad,
        Наименование_raw: fullName,
        Код: code,
        Остаток_ролики: stock
      });
    }
  }
  
  return result;
}

/**
 * Парсит файл продаж из 1С
 * Ищет шапку с полями: Склад, Номенклатура, Характеристика, Количество, Количество м2, Сумма выручки
 */
function parseSales(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  
  let headerIdx = -1;
  const headers = ['Вид номенклатуры', 'Партнер', 'Склад', 'Номенклатура', 'Характеристика'];
  
  // Ищем строку с заголовками
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowStr = row.join('|').toLowerCase();
    if (headers.every(h => rowStr.includes(h.toLowerCase()))) {
      headerIdx = i;
      break;
    }
  }
  
  if (headerIdx === -1) return [];
  
  const result = [];
  const headerRow = rows[headerIdx];
  
  // Найдём индексы столбцов
  let colSklad = -1, colNom = -1, colChar = -1, colQty = -1, colM2 = -1, colRevenue = -1;
  for (let j = 0; j < headerRow.length; j++) {
    const h = String(headerRow[j]).toLowerCase();
    if (h.includes('склад')) colSklad = j;
    else if (h.includes('номенклатур') && !h.includes('вид')) colNom = j;
    else if (h.includes('характеристик')) colChar = j;
    else if (h.includes('количество') && !h.includes('м2')) colQty = j;
    else if (h.includes('м2')) colM2 = j;
    else if (h.includes('сумм') || h.includes('выручк')) colRevenue = j;
  }
  
  // Парсим данные
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const sklad = row[colSklad] ? String(row[colSklad]).trim() : '';
    const nom = row[colNom] ? String(row[colNom]).trim() : '';
    const char = row[colChar] ? String(row[colChar]).trim() : '';
    const qty = row[colQty] ? parseFloat(row[colQty]) : 0;
    const m2 = row[colM2] ? parseFloat(row[colM2]) : 0;
    const revenue = row[colRevenue] ? parseFloat(row[colRevenue]) : 0;
    
    if (sklad && nom && qty > 0) {
      result.push({
        Склад: sklad,
        FullName: nom + (char ? ', ' + char : ''),
        Номенклатура: nom,
        Характеристика: char,
        Количество: qty,
        КоличествоМ2: m2,
        СуммаВыручки: revenue
      });
    }
  }
  
  return result;
}

/**
 * Парсит список неликвидов из листа "Неликвид"
 */
function parseRop(workbook) {
  let sheet = null;
  for (const name of workbook.SheetNames) {
    if (name.toLowerCase().includes('неликвид')) {
      sheet = workbook.Sheets[name];
      break;
    }
  }
  
  if (!sheet) sheet = workbook.Sheets[workbook.SheetNames[0]];
  
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length < 1) continue;
    
    const name = row[0] ? String(row[0]).trim() : '';
    const brand = row[1] ? String(row[1]).trim() : '';
    
    if (name) {
      result.push({
        Наименование: name,
        Бренд: brand
      });
    }
  }
  
  return result;
}

// ============ НОРМАЛИЗАЦИЯ И ПАРСИНГ ИМЁН ============

/**
 * Парсит из текста:
 * - Код (по шаблону: 3 цифры + буквы, например 303EVO, 303PTG, 303S, L305)
 * - Ширину (по шаблону: число + 'м', например 2.7м, 3.2м)
 */
function parseName(text) {
  const result = {
    Код: null,
    Ширина: null
  };
  
  if (!text) return result;
  
  // Ищем Код (3 цифры + буквы)
  const codeMatch = text.match(/(\d{3}[A-Z]*)/i);
  if (codeMatch) {
    result.Код = codeMatch[1];
  }
  
  // Ищем Ширину (число + 'м')
  const widthMatch = text.match(/(\d+\.?\d*)\s*м/i);
  if (widthMatch) {
    result.Ширина = parseFloat(widthMatch[1]);
  }
  
  return result;
}

/**
 * Обогащает RopItem данными из парсинга имени
 */
function enrichRopItem(item) {
  const parsed = parseName(item.Наименование);
  return {
    ...item,
    Код: parsed.Код,
    Ширина: parsed.Ширина
  };
}

/**
 * Обогащает StockItem данными из парсинга имени
 */
function enrichStockItem(item) {
  const parsed = parseName(item.Наименование_raw);
  return {
    ...item,
    Код: parsed.Код,
    Ширина: parsed.Ширина
  };
}

/**
 * Обогащает SaleItem данными из парсинга имени
 */
function enrichSaleItem(item) {
  const parsed = parseName(item.FullName);
  return {
    ...item,
    Код: parsed.Код,
    Ширина: parsed.Ширина
  };
}

// ============ СОПОСТАВЛЕНИЕ И ПОСТРОЕНИЕ ОТЧЁТА ============

/**
 * Строит итоговый отчёт по неликвидам
 */
function buildReport(ropList, stockRaw, salesRaw) {
  // Обогащаем данные
  const rop = ropList.map(enrichRopItem);
  const stock = stockRaw.map(enrichStockItem);
  const sales = salesRaw.map(enrichSaleItem);
  
  const result = [];
  
  // Для каждого неликвида РОПа
  for (const ropItem of rop) {
    // Если нет Кода или Ширины, пропускаем сопоставление
    if (!ropItem.Код || ropItem.Ширина === null) {
      result.push({
        Наименование: ropItem.Наименование,
        Бренд: ropItem.Бренд,
        Склад: 'N/A',
        Код: ropItem.Код,
        Ширина: ropItem.Ширина,
        Остаток_ролики: 0,
        Количество: 0,
        КоличествоМ2: 0,
        СуммаВыручки: 0,
        Статус: 'Нет данных'
      });
      continue;
    }
    
    // Ищем остатки по этому коду и ширине
    const matchingStock = stock.filter(s => 
      s.Код === ropItem.Код && s.Ширина === ropItem.Ширина
    );
    
    if (matchingStock.length === 0) {
      // Нет остатков
      result.push({
        Наименование: ropItem.Наименование,
        Бренд: ropItem.Бренд,
        Склад: 'N/A',
        Код: ropItem.Код,
        Ширина: ropItem.Ширина,
        Остаток_ролики: 0,
        Количество: 0,
        КоличествоМ2: 0,
        СуммаВыручки: 0,
        Статус: 'Нет данных'
      });
      continue;
    }
    
    // Для каждого склада с остатком
    for (const stockItem of matchingStock) {
      // Ищем продажи по складу, коду и ширине
      const matchingSales = sales.filter(s =>
        s.Склад === stockItem.Склад &&
        s.Код === ropItem.Код &&
        s.Ширина === ropItem.Ширина
      );
      
      const saleQty = matchingSales.reduce((sum, s) => sum + s.Количество, 0);
      const saleM2 = matchingSales.reduce((sum, s) => sum + s.КоличествоМ2, 0);
      const saleRevenue = matchingSales.reduce((sum, s) => sum + s.СуммаВыручки, 0);
      
      // Определяем статус
      let status = 'Нет данных';
      if (stockItem.Остаток_ролики > 0 && saleQty === 0) {
        status = 'Остаток есть, продаж нет';
      } else if (stockItem.Остаток_ролики > 0 && saleQty > 0) {
        status = 'Остаток и продажи есть';
      } else if (stockItem.Остаток_ролики === 0 && saleQty > 0) {
        status = 'Продан / остатка нет';
      }
      
      result.push({
        Наименование: ropItem.Наименование,
        Бренд: ropItem.Бренд,
        Склад: stockItem.Склад,
        Код: ropItem.Код,
        Ширина: ropItem.Ширина,
        Остаток_ролики: stockItem.Остаток_ролики,
        Количество: saleQty,
        КоличествоМ2: saleM2,
        СуммаВыручки: saleRevenue,
        Статус: status
      });
    }
  }
  
  return result;
}

// ============ АГРЕГАЦИЯ ДАННЫХ ============

/**
 * Строит KPI из отчёта
 */
function buildKpi(report) {
  return {
    positions: report.length,
    stocks: report.reduce((sum, r) => sum + r.Остаток_ролики, 0),
    sales_qty: report.reduce((sum, r) => sum + r.Количество, 0),
    sales_m2: report.reduce((sum, r) => sum + r.КоличествоМ2, 0),
    revenue: report.reduce((sum, r) => sum + r.СуммаВыручки, 0)
  };
}

/**
 * Группирует по складам
 */
function buildBySklad(report) {
  const map = new Map();
  
  for (const row of report) {
    if (!map.has(row.Склад)) {
      map.set(row.Склад, {
        Склад: row.Склад,
        Остаток_ролики: 0,
        Количество: 0,
        КоличествоМ2: 0,
        СуммаВыручки: 0
      });
    }
    const item = map.get(row.Склад);
    item.Остаток_ролики += row.Остаток_ролики;
    item.Количество += row.Количество;
    item.КоличествоМ2 += row.КоличествоМ2;
    item.СуммаВыручки += row.СуммаВыручки;
  }
  
  return Array.from(map.values()).sort((a, b) => b.Остаток_ролики - a.Остаток_ролики);
}

/**
 * Группирует по брендам
 */
function buildByBrand(report) {
  const map = new Map();
  
  for (const row of report) {
    if (!map.has(row.Бренд)) {
      map.set(row.Бренд, {
        Бренд: row.Бренд,
        Остаток_ролики: 0,
        Количество: 0,
        КоличествоМ2: 0,
        СуммаВыручки: 0
      });
    }
    const item = map.get(row.Бренд);
    item.Остаток_ролики += row.Остаток_ролики;
    item.Количество += row.Количество;
    item.КоличествоМ2 += row.КоличествоМ2;
    item.СуммаВыручки += row.СуммаВыручки;
  }
  
  return Array.from(map.values()).sort((a, b) => b.Остаток_ролики - a.Остаток_ролики);
}

/**
 * Считает по статусам
 */
function buildStatusCounts(report) {
  const map = new Map();
  
  for (const row of report) {
    const status = row.Статус || 'Неизвестно';
    map.set(status, (map.get(status) || 0) + 1);
  }
  
  return Array.from(map.entries()).map(([status, count]) => ({
    Статус: status,
    Количество: count
  }));
}

/**
 * Строит все агрегаты
 */
function buildAggregates(report) {
  return {
    kpi: buildKpi(report),
    bySklad: buildBySklad(report),
    byBrand: buildByBrand(report),
    statusCounts: buildStatusCounts(report)
  };
}

// ============ ВИЗУАЛИЗАЦИЯ ============

/**
 * Очищает и перерисовывает KPI карточки
 */
function renderKpi() {
  const k = aggregates.kpi;
  const items = [
    ['Позиции', fmt.format(k.positions)],
    ['Остаток, ролики', fmt.format(k.stocks)],
    ['Продажи, шт', fmt.format(k.sales_qty)],
    ['Выручка', money.format(k.revenue)]
  ];
  document.getElementById('kpis').innerHTML = items.map(([label, val]) => `
    <div class="card">
      <div class="kpi-label">${label}</div>
      <div class="kpi-value">${val}</div>
    </div>`).join('');
}

/**
 * Заполняет выпадающие списки фильтров
 */
function fillFilters() {
  const fill = (id, values) => {
    const el = document.getElementById(id);
    el.innerHTML = '<option value="">Все</option>';
    [...new Set(values)].sort().forEach(v => {
      if (!v) return;
      const o = document.createElement('option');
      o.value = v;
      o.textContent = v;
      el.appendChild(o);
    });
  };
  fill('skladFilter', report.map(r => r.Склад));
  fill('brandFilter', report.map(r => r.Бренд));
  fill('statusFilter', report.map(r => r.Статус));
}

/**
 * CSS класс для статуса
 */
function statusClass(s) {
  if (!s) return '';
  if (s.includes('продаж нет')) return 's2';
  if (s.includes('продажи есть')) return 's1';
  if (s.includes('Продан')) return 's0';
  return '';
}

/**
 * Перерисовывает таблицу с фильтрацией
 */
function renderTable() {
  const q = document.getElementById('search').value.toLowerCase();
  const skl = document.getElementById('skladFilter').value;
  const br = document.getElementById('brandFilter').value;
  const st = document.getElementById('statusFilter').value;

  const rows = report.filter(r =>
    (!q || r.Наименование.toLowerCase().includes(q)) &&
    (!skl || r.Склад === skl) &&
    (!br || r.Бренд === br) &&
    (!st || r.Статус === st)
  );

  state.rows = rows;

  document.getElementById('tbody').innerHTML = rows.map(r => `
    <tr>
      <td>${r.Наименование}</td>
      <td>${r.Бренд}</td>
      <td>${r.Склад}</td>
      <td>${r.Код ?? ''}</td>
      <td>${r.Ширина ? r.Ширина.toFixed(1) + 'м' : ''}</td>
      <td>${fmt.format(r.Остаток_ролики)}</td>
      <td>${fmt.format(r.Количество)}</td>
      <td><span class="status ${statusClass(r.Статус)}">${r.Статус || ''}</span></td>
    </tr>`).join('');

  document.getElementById('countInfo').textContent = `Строк: ${rows.length}`;
}

/**
 * Перерисовывает графики
 */
function renderCharts() {
  // Уничтожаем старые графики
  Object.values(state.charts).forEach(chart => {
    if (chart) chart.destroy();
  });
  state.charts = {};

  // График по складам
  state.charts.sklad = new Chart(document.getElementById('skladChart'), {
    type: 'bar',
    data: {
      labels: aggregates.bySklad.map(x => x.Склад),
      datasets: [{
        label: 'Остаток, ролики',
        data: aggregates.bySklad.map(x => x.Остаток_ролики),
        backgroundColor: '#01696f'
      }, {
        label: 'Продажи, шт',
        data: aggregates.bySklad.map(x => x.Количество),
        backgroundColor: '#d19900'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: { y: { beginAtZero: true } }
    }
  });

  // График по статусам
  state.charts.status = new Chart(document.getElementById('statusChart'), {
    type: 'doughnut',
    data: {
      labels: aggregates.statusCounts.map(x => x.Статус),
      datasets: [{
        data: aggregates.statusCounts.map(x => x.Количество),
        backgroundColor: ['#a12c7b', '#01696f', '#437a22', '#d19900']
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });

  // График по брендам
  const topBrands = aggregates.byBrand.slice(0, 10);
  state.charts.brand = new Chart(document.getElementById('brandChart'), {
    type: 'bar',
    data: {
      labels: topBrands.map(x => x.Бренд),
      datasets: [{
        label: 'Остаток, ролики',
        data: topBrands.map(x => x.Остаток_ролики),
        backgroundColor: '#4f98a3'
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      scales: { x: { beginAtZero: true } }
    }
  });
}

// ============ СКАЧИВАНИЕ ОТЧЁТА ============

/**
 * Скачивает отфильтрованный отчёт в Excel
 */
function downloadReport() {
  if (!state.rows || state.rows.length === 0) {
    alert('Нет данных для скачивания');
    return;
  }

  // Подготавливаем данные
  const data = state.rows.map(r => ({
    'Наименование': r.Наименование,
    'Бренд': r.Бренд,
    'Склад': r.Склад,
    'Код': r.Код || '',
    'Ширина': r.Ширина ? r.Ширина.toFixed(1) + 'м' : '',
    'Остаток, ролики': r.Остаток_ролики,
    'Продажи, шт': r.Количество,
    'Продажи, м2': r.КоличествоМ2,
    'Выручка': r.СуммаВыручки,
    'Статус': r.Статус
  }));

  // Создаём лист
  const ws = XLSX.utils.json_to_sheet(data);
  
  // Устанавливаем ширину колонок
  ws['!cols'] = [
    { wch: 30 },  // Наименование
    { wch: 12 },  // Бренд
    { wch: 12 },  // Склад
    { wch: 8 },   // Код
    { wch: 10 },  // Ширина
    { wch: 10 },  // Остаток
    { wch: 10 },  // Продажи шт
    { wch: 10 },  // Продажи м2
    { wch: 12 },  // Выручка
    { wch: 25 }   // Статус
  ];

  // Создаём рабочую книгу
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Неликвиды');

  // Скачиваем
  XLSX.writeFile(wb, 'nelikvid_report.xlsx');
}

// ============ ИНИЦИАЛИЗАЦИЯ И ОБРАБОТЧИКИ ============

/**
 * Обработчик кнопки "Обработать"
 */
async function processFiles() {
  try {
    const fileStock = document.getElementById('fileStock').files[0];
    const fileSales = document.getElementById('fileSales').files[0];
    const fileRop = document.getElementById('fileRop').files[0];

    if (!fileStock || !fileSales || !fileRop) {
      document.getElementById('uploadStatus').textContent = 'Ошибка: выберите все три файла';
      return;
    }

    document.getElementById('uploadStatus').textContent = 'Обработка...';

    // Параллельно читаем файлы
    const [wbStock, wbSales, wbRop] = await Promise.all([
      readExcel(fileStock),
      readExcel(fileSales),
      readExcel(fileRop)
    ]);

    // Парсим
    stockRaw = parseStock(wbStock);
    salesRaw = parseSales(wbSales);
    ropList = parseRop(wbRop);

    // Строим отчёт и агрегаты
    report = buildReport(ropList, stockRaw, salesRaw);
    aggregates = buildAggregates(report);

    // Инициализируем state
    state.rows = report.slice();

    // Обновляем UI
    renderKpi();
    fillFilters();
    renderTable();
    renderCharts();

    // Активируем кнопку скачивания
    document.getElementById('btnDownload').disabled = false;

    document.getElementById('uploadStatus').textContent = `Обработано: ${report.length} позиций неликвидов`;

    // Добавляем слушатели на фильтры
    ['search', 'skladFilter', 'brandFilter', 'statusFilter']
      .forEach(id => {
        const el = document.getElementById(id);
        el.removeEventListener('change', renderTable);
        el.removeEventListener('input', renderTable);
        el.addEventListener('input', renderTable);
        el.addEventListener('change', renderTable);
      });

  } catch (error) {
    console.error('Ошибка обработки:', error);
    document.getElementById('uploadStatus').textContent = 'Ошибка: ' + error.message;
  }
}

// Инициализация при загрузке страницы
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('btnProcess').addEventListener('click', processFiles);
  document.getElementById('btnDownload').addEventListener('click', downloadReport);
});
