(() => {
  'use strict';

  // Header Constants
  const RAW_HEADERS = [
    'OUTLET','NAMA OUTLET','TYPE','GRUP','SLSNO','KG','TGL PENGANTARAN',
    'TGLFAKTUR','NO.FAKTUR','PCODE','NAMA BARANG','SATUAN/ISI','......QUANTITY',
    '....HARGA(RP)','....JMLBRUTTO','....RP.DISCPC','.....RP.DISC1','.....RP.DISC2',
    '...PROMO UANG','.....DISCOUNT','..........PPN','........TOTAL'
  ];

  const PROCESSED_HEADERS = [
    'OUTLET ','NAMA OUTLET         ','TYPE ','GRUP ','SLSNO  ','KG ','TGL PENGANTARAN',
    'TGLFAKTUR','NO.FAKTUR ','PCODE  ','NAMA BARANG                         ','SATU',
    'AN/','ISI','......QUANTITY ','....HARGA(RP)  ','....JMLBRUTTO  ','....RP.DISCPC ',
    '.....RP.DISC1 ','.....RP.DISC2 ','...PROMO UANG ','.....DISCOUNT ','..........PPN ',
    '........TOTAL ','HARGA PCS','HARGA CTN'
  ];

  const IMPORT_HEADERS = [
    'ND6TRAN','salesinvoice','attribute','salesmanID','salesorderNumber','salesorderDate',
    'invoiceNumber','invoiceDate','term','soldtoCustomerID','senttoCustomerID',
    'invoicedtoCustomerID','customerPO','sellingType','documentType','cashPayment',
    'giroPayment','giroNumber','giroBank','giroDue','adjustmentAmount','discount1',
    'discount2','discount3','tax1','tax2','tax3','productCode','productvarianCode',
    'qtySold','qtyFreeGood','sellingPrice','lineDiscount1','lineDiscount2','lineDiscount3',
    'lineDiscount4','lineDiscount5','companyID','branchID','divisionID','warehouseID',
    'barcodeList','isSKU'
  ];

  // DOM Elements
  const $ = id => document.getElementById(id);
  const dropzone = $('dropzone');
  const fileInput = $('fileInput');
  const chooseBtn = $('chooseBtn');
  const demoBtn = $('demoBtn');
  const clearBtn = $('clearBtn');
  const processBtn = $('processBtn');
  const downloadProcessedBtn = $('downloadProcessedBtn');
  const downloadImportXlsxBtn = $('downloadImportXlsxBtn');
  const downloadImportBtn = $('downloadImportBtn');
  const downloadImportTxtBtn = $('downloadImportTxtBtn');
  const branchSelect = $('branchSelect');
  const statusBox = $('status');
  const previewTable = $('previewTable');
  const searchInput = $('searchInput');
  const themeToggleBtn = $('themeToggleBtn');

  // Application State
  let selectedFile = null;
  let isDemoFile = false;
  let processedWorkbook = null;
  let processedFilename = '';
  let importWorkbook = null;
  let importXlsxFilename = '';
  let importCsv = '';
  let importFilename = '';
  let importTxt = '';
  let importTxtFilename = '';
  let warnings = [];
  let convertedDiscounts = 0;
  let currentImportRows = [];

  // Helper Functions
  const cleanHeader = v => String(v ?? '').trim().replace(/\s+/g, ' ').toUpperCase();
  const hasValue = v => v !== null && v !== undefined && String(v) !== '';
  
  function setStatus(type, msg) {
    statusBox.className = `status show ${type}`;
    statusBox.innerHTML = msg;
  }
  
  function clearStatus() {
    statusBox.className = 'status';
    statusBox.textContent = '';
  }

  function bytes(n) {
    if (!n) return '0 KB';
    const u = ['B', 'KB', 'MB', 'GB'];
    const i = Math.min(Math.floor(Math.log(n) / Math.log(1024)), u.length - 1);
    return `${(n / (1024 ** i)).toFixed(i ? 1 : 0)} ${u[i]}`;
  }

  function resetResult() {
    processedWorkbook = null;
    processedFilename = '';
    importWorkbook = null;
    importXlsxFilename = '';
    importCsv = '';
    importFilename = '';
    importTxt = '';
    importTxtFilename = '';
    warnings = [];
    convertedDiscounts = 0;
    currentImportRows = [];
    downloadProcessedBtn.disabled = true;
    downloadImportXlsxBtn.disabled = true;
    downloadImportBtn.disabled = true;
    downloadImportTxtBtn.disabled = true;
    $('stats').classList.remove('show');
    $('previewSection').classList.remove('show');
    previewTable.innerHTML = '';
    if (searchInput) searchInput.value = '';
    clearStatus();
  }

  function selectFile(file, isDemo = false) {
    if (!file) return;
    if (!/\.(xlsx|xls)$/i.test(file.name) && !isDemo) {
      setStatus('error', '<strong>File tidak didukung.</strong> Gunakan file .xlsx atau .xls.');
      return;
    }
    selectedFile = file;
    isDemoFile = isDemo;
    resetResult();
    $('fileName').textContent = file.name;
    $('fileMeta').textContent = isDemo ? 'Demo Data (Excel) · Siap Diproses' : `${bytes(file.size)} · Siap Diproses`;
    $('fileInfo').classList.add('show');
    processBtn.disabled = false;
  }

  function clearFile() {
    selectedFile = null;
    isDemoFile = false;
    fileInput.value = '';
    $('fileInfo').classList.remove('show');
    processBtn.disabled = true;
    resetResult();
  }

  function rawRows(sheet) {
    if (!sheet || !sheet['!ref']) return [];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const rows = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c })];
        row.push(cell && cell.v !== undefined ? cell.v : null);
      }
      rows.push(row);
    }
    return rows;
  }

  function normalizedHeaderAt(row, i) {
    const h = cleanHeader(row[i]);
    if (i === 6 && h === 'INVALID DATE') return 'TGL PENGANTARAN';
    return h;
  }

  function startsWithHeaders(row, expected) {
    if (row.length < expected.length) return false;
    for (let i = 0; i < expected.length; i++) {
      if (normalizedHeaderAt(row, i) !== cleanHeader(expected[i])) return false;
    }
    return true;
  }

  function detectType(row) {
    if (startsWithHeaders(row, RAW_HEADERS)) return 'raw';
    if (startsWithHeaders(row, PROCESSED_HEADERS)) return 'processed';
    return null;
  }

  function toNum(v, row, field) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number' && Number.isFinite(v)) return v;
    const n = Number(String(v).trim());
    if (Number.isFinite(n)) return n;
    warnings.push(`Baris ${row}: ${field} tidak dapat diubah menjadi angka (${v}).`);
    return v;
  }

  function numeric(v) {
    if (typeof v === 'number' && Number.isFinite(v)) return v;
    const n = Number(String(v ?? '').trim().replace(',', '.'));
    return Number.isFinite(n) ? n : null;
  }

  function excelSerial(y, m, d) {
    return Math.floor(Date.UTC(y, m - 1, d) / 86400000) + 25569;
  }

  function validDate(y, m, d) {
    const x = new Date(Date.UTC(y, m - 1, d));
    return y >= 1900 && x.getUTCFullYear() === y && x.getUTCMonth() === m - 1 && x.getUTCDate() === d;
  }

  function dateSerial(v, row, field, removeSpaces = false) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number' && Number.isFinite(v)) return v;
    if (v instanceof Date && !isNaN(v)) return excelSerial(v.getFullYear(), v.getMonth() + 1, v.getDate());
    let s = String(v);
    s = removeSpaces ? s.replace(/\s+/g, '') : s.trim();
    if (/^\d+(\.\d+)?$/.test(s)) return Number(s);
    let m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
    if (m && validDate(+m[3], +m[2], +m[1])) return excelSerial(+m[3], +m[2], +m[1]);
    m = s.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})$/);
    if (m && validDate(+m[1], +m[2], +m[3])) return excelSerial(+m[1], +m[2], +m[3]);
    warnings.push(`Baris ${row}: ${field} tidak dikenali sebagai tanggal (${s}).`);
    return v;
  }

  function splitUnit(v, row) {
    const s = String(v ?? '').trim();
    if (!s) return ['', null, ''];
    const m = s.match(/^([^/]+\/)\s*([0-9]+(?:[.,][0-9]+)?)\s*(.*)$/);
    if (!m) {
      warnings.push(`Baris ${row}: SATUAN/ISI tidak sesuai pola format (${s}).`);
      return [s, null, ''];
    }
    const an = Number(m[2].replace(',', '.'));
    return [m[1].trim(), Number.isFinite(an) ? an : m[2], m[3].trim()];
  }

  function rawToProcessed(rows) {
    const out = [PROCESSED_HEADERS.slice()];
    for (let i = 1; i < rows.length; i++) {
      const s = rows[i];
      if (!s.some(hasValue)) continue;
      const rn = i + 1;
      const [satu, an, isi] = splitUnit(s[11], rn);
      const q = numeric(s[12]);
      const br = numeric(s[14]);
      const anN = numeric(an);
      const pcs = (q !== null && q !== 0 && br !== null) ? br / q : null;
      const ctn = (pcs !== null && anN !== null) ? pcs * anN : null;

      if (q === 0) warnings.push(`Baris ${rn}: QUANTITY = 0.`);
      out.push([
        toNum(s[0], rn, 'OUTLET'),
        s[1], s[2], s[3],
        toNum(s[4], rn, 'SLSNO'),
        s[5],
        dateSerial(s[6], rn, 'TGL PENGANTARAN'),
        dateSerial(s[7], rn, 'TGLFAKTUR', true),
        toNum(s[8], rn, 'NO.FAKTUR'),
        toNum(s[9], rn, 'PCODE'),
        s[10], satu, an, isi,
        s[12], s[13], s[14], s[15], s[16], s[17], s[18], s[19], s[20], s[21],
        pcs, ctn
      ]);
    }
    return out;
  }

  function processedNormalize(rows) {
    const out = [PROCESSED_HEADERS.slice()];
    for (let i = 1; i < rows.length; i++) {
      let s = rows[i];
      if (!s.some(hasValue)) continue;
      s = s.slice(0, 26);
      while (s.length < 26) s.push(null);
      const rn = i + 1;
      s[0] = toNum(s[0], rn, 'OUTLET');
      s[4] = toNum(s[4], rn, 'SLSNO');
      s[6] = dateSerial(s[6], rn, 'TGL PENGANTARAN');
      s[7] = dateSerial(s[7], rn, 'TGLFAKTUR', true);
      s[8] = toNum(s[8], rn, 'NO.FAKTUR');
      s[9] = toNum(s[9], rn, 'PCODE');
      
      let q = numeric(s[14]);
      let br = numeric(s[16]);
      let an = numeric(s[12]);
      if (numeric(s[24]) === null && q && br !== null) s[24] = br / q;
      if (numeric(s[25]) === null && numeric(s[24]) !== null && an !== null) s[25] = numeric(s[24]) * an;
      out.push(s);
    }
    return out;
  }

  function serialIso(v) {
    if (typeof v === 'number' && Number.isFinite(v)) {
      const d = new Date((v - 25569) * 86400000);
      return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
    }
    let s = String(v ?? '').trim().replace(/\s+/g, '');
    if (!s) return '';
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
    m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
    if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    if (/^\d+(\.\d+)?$/.test(s)) return serialIso(Number(s));
    return s;
  }

  function plain(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number' && Number.isFinite(v)) return String(v);
    const s = String(v).trim();
    const n = Number(s);
    return (s !== '' && Number.isFinite(n)) ? String(n) : s;
  }

  function two(v, def = '0.00') {
    const n = numeric(v);
    return n === null ? def : n.toFixed(2);
  }

  function discValue(v, selling, row, label) {
    const d = numeric(v);
    if (d === null || d === 0) return '0.00';
    const sp = numeric(selling);
    if (d > 0 && d < 100) {
      if (sp === null || sp === 0) {
        warnings.push(`Baris ${row}: ${label} < 100 tetapi sellingPrice kosong atau 0.`);
        return d.toFixed(2);
      }
      convertedDiscounts++;
      return ((d / sp) * 100).toFixed(2);
    }
    return d.toFixed(2);
  }

  function indexMap(headers) {
    const m = {};
    headers.forEach((h, i) => m[cleanHeader(h)] = i);
    return m;
  }

  function processedToImport(processed, branchId) {
    const h = indexMap(processed[0]);
    const ix = name => {
      const k = cleanHeader(name);
      if (h[k] === undefined) throw new Error(`Kolom ${name} tidak ditemukan pada hasil olah.`);
      return h[k];
    };

    const I = {
      outlet: ix('OUTLET'), sls: ix('SLSNO'), delivery: ix('TGL PENGANTARAN'),
      fakturDate: ix('TGLFAKTUR'), faktur: ix('NO.FAKTUR'), pcode: ix('PCODE'),
      qty: ix('......QUANTITY'), discpc: ix('....RP.DISCPC'), disc1: ix('.....RP.DISC1'),
      disc2: ix('.....RP.DISC2'), promo: ix('...PROMO UANG'), discount: ix('.....DISCOUNT'),
      ctn: ix('HARGA CTN')
    };

    const out = [IMPORT_HEADERS.slice()];
    for (let r = 1; r < processed.length; r++) {
      const s = processed[r];
      if (!s.some(hasValue)) continue;
      const order = plain(s[I.faktur]);
      const outlet = plain(s[I.outlet]);
      const price = two(s[I.ctn], '');
      out.push([
        'ND6TRAN', 'salesinvoice', 'value', plain(s[I.sls]), order,
        serialIso(s[I.fakturDate]), `I${order}`, serialIso(s[I.delivery]),
        'T000', outlet, outlet, outlet, '', 'TO', '', '0', '0', '', '', '', '0',
        '0', '0', '0', '11', '0', '0', plain(s[I.pcode]), '-', plain(s[I.qty]), '0',
        price,
        discValue(s[I.promo], price, r + 1, 'lineDiscount1'),
        discValue(s[I.disc2], price, r + 1, 'lineDiscount2'),
        discValue(s[I.discpc], price, r + 1, 'lineDiscount3'),
        discValue(s[I.disc1], price, r + 1, 'lineDiscount4'),
        discValue(s[I.discount], price, r + 1, 'lineDiscount5'),
        'NS6081030004380', String(branchId), '1717050000000', 'STANDARD', '-', 'N'
      ]);
    }
    return out;
  }

  function buildProcessedWb(rows) {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const last = rows.length;
    for (let r = 2; r <= last; r++) {
      for (const c of ['G', 'H']) {
        const cell = ws[`${c}${r}`];
        if (cell && typeof cell.v === 'number') {
          cell.t = 'n';
          cell.z = 'yyyy-mm-dd';
        }
      }
      const y = ws[`Y${r}`];
      if (y && typeof y.v === 'number' && Number.isFinite(y.v)) {
        y.f = `Q${r}/O${r}`;
        y.t = 'n';
        delete y.z;
      }
      const z = ws[`Z${r}`];
      if (z && typeof z.v === 'number' && Number.isFinite(z.v)) {
        z.f = `Y${r}*M${r}`;
        z.t = 'n';
        delete z.z;
      }
    }
    ws['!cols'] = [11,24,8,8,11,7,17,13,13,11,38,9,9,18,16,16,17,16,16,16,16,16,14,14,14,14].map(wch => ({ wch }));
    ws['!autofilter'] = { ref: `A1:Z${last}` };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'IMPORT');
    wb.CalcPr = { fullCalcOnLoad: true, forceFullCalc: true, calcMode: 'auto' };
    return wb;
  }

  function isoToSerial(v) {
    const s = String(v ?? '').trim();
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? excelSerial(+m[1], +m[2], +m[3]) : v;
  }

  function buildImportWb(rows) {
    const data = rows.map(r => r.slice());
    const ws = XLSX.utils.aoa_to_sheet(data);
    const last = rows.length;
    for (let r = 2; r <= last; r++) {
      for (const c of ['F', 'H']) {
        const cell = ws[`${c}${r}`];
        if (cell && cell.v !== '' && cell.v !== null && cell.v !== undefined) {
          const serial = isoToSerial(cell.v);
          if (typeof serial === 'number' && Number.isFinite(serial)) {
            cell.v = serial;
            cell.t = 'n';
            cell.z = 'yyyy-mm-dd';
          }
        }
      }
      for (const c of ['AM', 'AN']) {
        const cell = ws[`${c}${r}`];
        if (cell && cell.v !== '' && cell.v !== null && cell.v !== undefined) {
          const n = Number(cell.v);
          if (Number.isFinite(n)) {
            cell.v = n;
            cell.t = 'n';
            cell.z = '0';
          }
        }
      }
      for (const c of ['AF', 'AG', 'AH', 'AI', 'AJ', 'AK']) {
        const cell = ws[`${c}${r}`];
        if (!cell || cell.v === '' || cell.v === null || cell.v === undefined) continue;
        const n = Number(cell.v);
        if (Number.isFinite(n)) {
          cell.v = n;
          cell.t = 'n';
          cell.z = '0.00';
        }
      }
    }
    ws['!cols'] = [14,14,12,13,18,14,16,14,10,18,18,22,14,13,14,13,13,14,13,11,18,12,12,12,9,9,9,14,18,11,14,15,15,15,15,15,15,20,18,18,15,13,9].map(wch => ({ wch }));
    ws['!autofilter'] = { ref: `A1:AQ${last}` };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'IMPORT');
    return wb;
  }

  function csvCell(v) {
    const s = String(v ?? '');
    return /[;"\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  }
  function toCsv(rows) {
    return rows.map(row => row.map(csvCell).join(';')).join('\r\n') + '\r\n';
  }

  function txtCell(v) {
    const s = String(v ?? '');
    return /[|"\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  }
  function toTxt(rows) {
    return rows.map(row => row.map(txtCell).join('|')).join('\r\n') + '\r\n';
  }

  function processedName(name) {
    const base = name.replace(/\.(xlsx|xls)$/i, '');
    if (/\(\s*sebelum\s+di\s+olah\s*\)/i.test(base)) {
      return base.replace(/\(\s*sebelum\s+di\s+olah\s*\)/i, '(Sesudah di Olah)') + '.xlsx';
    }
    if (/sebelum\s+di\s+olah/i.test(base)) {
      return base.replace(/sebelum\s+di\s+olah/i, 'Sesudah di Olah') + '.xlsx';
    }
    if (/sesudah\s+di\s+olah/i.test(base)) return base + '.xlsx';
    return base + ' (Sesudah di Olah).xlsx';
  }

  const monthMap = {
    JANUARI:'01',JAN:'01',FEBRUARI:'02',FEB:'02',MARET:'03',MAR:'03',APRIL:'04',APR:'04',MEI:'05',MAY:'05',
    JUNI:'06',JUN:'06',JULI:'07',JUL:'07',AGUSTUS:'08',AGU:'08',AUG:'08',SEPTEMBER:'09',SEP:'09',OKTOBER:'10',
    OKT:'10',OCT:'10',NOVEMBER:'11',NOV:'11',DESEMBER:'12',DES:'12',DEC:'12'
  };

  function dateTag(name) {
    const u = name.toUpperCase();
    let m = u.match(/TGL\s*(\d{1,2})\s+([A-Z]+)(?:\s+\d{4})?/);
    if (m && monthMap[m[2]]) return `${m[1].padStart(2, '0')}${monthMap[m[2]]}`;
    m = u.match(/(?:TGL\s*)?(\d{1,2})[\-_.\/ ](\d{1,2})(?:[\-_.\/ ]\d{2,4})?/);
    if (m) return `${m[1].padStart(2, '0')}${m[2].padStart(2, '0')}`;
    return '';
  }

  function branchInfo() {
    const opt = branchSelect.options[branchSelect.selectedIndex];
    return {
      id: branchSelect.value,
      prefix: opt?.dataset?.prefix || '',
      name: opt ? opt.textContent.split(' — ')[0] : '-'
    };
  }

  function importName(input) {
    const b = branchInfo();
    const tag = dateTag(input);
    return `${b.prefix}_Import NABATI${tag ? ` - ${tag}` : ''}.csv`;
  }
  function importExcelName(input) { return importName(input).replace(/\.csv$/i, '.xlsx'); }
  function importTxtName(input) { return importName(input).replace(/\.csv$/i, '.txt'); }

  function renderPreview(rows, filterKeyword = '') {
    if (!rows || !rows.length) {
      previewTable.innerHTML = '';
      return;
    }

    const keyword = filterKeyword.trim().toLowerCase();
    const filteredBodyRows = [];

    for (let r = 1; r < rows.length; r++) {
      if (!keyword) {
        filteredBodyRows.push(rows[r]);
        if (filteredBodyRows.length >= 10) break; // Display top 10 matching rows
      } else {
        const rowStr = rows[r].map(x => String(x ?? '')).join(' ').toLowerCase();
        if (rowStr.includes(keyword)) {
          filteredBodyRows.push(rows[r]);
          if (filteredBodyRows.length >= 10) break;
        }
      }
    }

    const frag = document.createDocumentFragment();
    const head = document.createElement('thead');
    const trh = document.createElement('tr');
    
    rows[0].forEach(x => {
      const th = document.createElement('th');
      th.textContent = x;
      trh.appendChild(th);
    });
    head.appendChild(trh);
    frag.appendChild(head);

    const body = document.createElement('tbody');
    if (filteredBodyRows.length === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = rows[0].length;
      td.style.textAlign = 'center';
      td.style.padding = '20px';
      td.style.color = 'var(--text-muted)';
      td.textContent = keyword ? `Tidak ada data yang cocok dengan "${filterKeyword}"` : 'Tidak ada data.';
      tr.appendChild(td);
      body.appendChild(tr);
    } else {
      filteredBodyRows.forEach(row => {
        const tr = document.createElement('tr');
        row.forEach(x => {
          const td = document.createElement('td');
          td.textContent = x ?? '';
          tr.appendChild(td);
        });
        body.appendChild(tr);
      });
    }

    frag.appendChild(body);
    previewTable.innerHTML = '';
    previewTable.appendChild(frag);
  }

  // Processing Controller
  async function process() {
    if (!selectedFile) return;
    if (!branchSelect.value) {
      setStatus('error', '<strong>Pilih branch terlebih dahulu.</strong> Silakan pilih Air Molek atau Ujung Batu.');
      return;
    }
    if (typeof XLSX === 'undefined') {
      setStatus('error', 'Library Excel belum termuat. Pastikan koneksi internet aktif lalu refresh halaman.');
      return;
    }

    processBtn.disabled = true;
    downloadProcessedBtn.disabled = true;
    downloadImportXlsxBtn.disabled = true;
    downloadImportBtn.disabled = true;
    downloadImportTxtBtn.disabled = true;
    warnings = [];
    convertedDiscounts = 0;
    setStatus('processing', 'Membaca file, mengolah data, dan menyusun template import...');

    try {
      let buf;
      if (isDemoFile && selectedFile.demoBuffer) {
        buf = selectedFile.demoBuffer;
      } else {
        buf = await selectedFile.arrayBuffer();
      }

      const wb = XLSX.read(buf, { type: 'array', raw: true, cellDates: false, cellFormula: true, cellNF: true, cellText: false });
      const sheetName = wb.SheetNames.includes('IMPORT') ? 'IMPORT' : wb.SheetNames[0];
      if (!sheetName) throw new Error('Workbook tidak memiliki sheet.');

      const rows = rawRows(wb.Sheets[sheetName]);
      if (rows.length < 2) throw new Error('File tidak memiliki data baris.');

      const type = detectType(rows[0]);
      if (!type) throw new Error('Struktur file tidak dikenali sebagai format "Sebelum di Olah" ataupun "Sesudah di Olah".');

      const processed = (type === 'raw') ? rawToProcessed(rows) : processedNormalize(rows);
      if (processed.length < 2) throw new Error('Tidak ada baris data valid yang dapat diproses.');

      processedWorkbook = buildProcessedWb(processed);
      processedFilename = processedName(selectedFile.name);

      const imports = processedToImport(processed, branchSelect.value);
      currentImportRows = imports;

      importWorkbook = buildImportWb(imports);
      importXlsxFilename = importExcelName(selectedFile.name);
      importCsv = toCsv(imports);
      importFilename = importName(selectedFile.name);
      importTxt = toTxt(imports);
      importTxtFilename = importTxtName(selectedFile.name);

      $('inputType').textContent = (type === 'raw') ? 'Sebelum di Olah' : 'Sesudah di Olah';
      $('rowCount').textContent = String(imports.length - 1);
      $('branchName').textContent = branchInfo().name;
      $('discountCount').textContent = String(convertedDiscounts);
      $('stats').classList.add('show');

      renderPreview(imports);
      $('previewSection').classList.add('show');

      downloadProcessedBtn.disabled = false;
      downloadImportXlsxBtn.disabled = false;
      downloadImportBtn.disabled = false;
      downloadImportTxtBtn.disabled = false;

      const warnMsg = warnings.length 
        ? `<br><span style="font-size:11.5px; opacity:0.9">${warnings.slice(0, 3).map(w => '• ' + w).join('<br>')}${warnings.length > 3 ? '<br>…' : ''}</span>`
        : '';
        
      setStatus(
        warnings.length ? 'warning' : 'success',
        `<strong>Berhasil diproses!</strong> Total <strong>${imports.length - 1}</strong> baris data siap diimport untuk branch <strong>${branchInfo().name}</strong>.${convertedDiscounts ? ` Terdapat ${convertedDiscounts} nilai discount < 100 yang dikonversi ke persen.` : ''}${warnMsg}`
      );
    } catch (e) {
      console.error(e);
      setStatus('error', `<strong>Gagal memproses file.</strong><br>${e?.message || 'Terjadi kesalahan sistem.'}`);
      processedWorkbook = null;
      importWorkbook = null;
      importCsv = '';
      importTxt = '';
      currentImportRows = [];
    } finally {
      processBtn.disabled = !selectedFile;
    }
  }

  // Demo Data Generator
  function generateDemoData() {
    const rawDemoRows = [
      RAW_HEADERS.slice(),
      [
        '10002341', 'TOKO REZEKI ABADI', 'RETAIL', 'GROSIR', '50012', 15.5,
        '15/09/2026', '15/09/2026', 'FK-98421', '8992741001', 'RICHEESE WAFER 145G',
        'CTN/ 24 PCS', 10, 5000, 50000, 5, 0, 0, 2500, 10, 5500, 55500
      ],
      [
        '10002342', 'WARUNG BERKAH AMANAH', 'RETAIL', 'SMALL', '50012', 8.2,
        '15/09/2026', '15/09/2026', 'FK-98422', '8992741002', 'NEXTAR CHOCO 112G',
        'CTN/ 12 PCS', 5, 6000, 30000, 0, 0, 0, 0, 5, 3300, 33300
      ]
    ];

    const ws = XLSX.utils.aoa_to_sheet(rawDemoRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'IMPORT');
    const demoArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    const demoFile = {
      name: 'NABATI_Sales_Report_TGL 15 SEPT 2026 (Sebelum di Olah).xlsx',
      size: demoArray.byteLength,
      demoBuffer: demoArray
    };

    if (!branchSelect.value) {
      branchSelect.value = '1764907397498'; // Default Air Molek
    }

    selectFile(demoFile, true);
    setStatus('success', '<strong>Demo Data dimuat!</strong> Klik tombol <strong>Proses Data</strong> di atas untuk menjalankan simulasi pengolahan.');
  }

  // Download Trigger Handlers
  function downloadProcessed() {
    if (!processedWorkbook) return;
    XLSX.writeFile(processedWorkbook, processedFilename, { bookType: 'xlsx', compression: true, cellStyles: true });
  }

  function downloadImportXlsx() {
    if (!importWorkbook) return;
    XLSX.writeFile(importWorkbook, importXlsxFilename, { bookType: 'xlsx', compression: true, cellStyles: true });
  }

  function downloadImport() {
    if (!importCsv) return;
    const blob = new Blob([importCsv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = importFilename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function downloadImportTxt() {
    if (!importTxt) return;
    const blob = new Blob([importTxt], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = importTxtFilename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // Theme Manager
  function initTheme() {
    const savedTheme = localStorage.getItem('nabati_theme');
    if (savedTheme === 'light') {
      document.body.classList.add('light-theme');
    }
  }

  function toggleTheme() {
    document.body.classList.toggle('light-theme');
    const isLight = document.body.classList.contains('light-theme');
    localStorage.setItem('nabati_theme', isLight ? 'light' : 'dark');
  }

  // Event Listeners Registration
  chooseBtn.addEventListener('click', e => { e.stopPropagation(); fileInput.click(); });
  dropzone.addEventListener('click', () => fileInput.click());
  demoBtn.addEventListener('click', e => { e.stopPropagation(); generateDemoData(); });
  
  dropzone.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      fileInput.click();
    }
  });

  fileInput.addEventListener('change', () => selectFile(fileInput.files[0]));
  clearBtn.addEventListener('click', clearFile);
  processBtn.addEventListener('click', process);
  downloadProcessedBtn.addEventListener('click', downloadProcessed);
  downloadImportXlsxBtn.addEventListener('click', downloadImportXlsx);
  downloadImportBtn.addEventListener('click', downloadImport);
  downloadImportTxtBtn.addEventListener('click', downloadImportTxt);

  if (themeToggleBtn) {
    themeToggleBtn.addEventListener('click', toggleTheme);
  }

  if (searchInput) {
    searchInput.addEventListener('input', e => {
      renderPreview(currentImportRows, e.target.value);
    });
  }

  branchSelect.addEventListener('change', () => {
    if (selectedFile && importCsv) {
      resetResult();
      processBtn.disabled = false;
      $('fileInfo').classList.add('show');
    }
  });

  // Drag and Drop Event Handling
  ['dragenter', 'dragover'].forEach(t => dropzone.addEventListener(t, e => {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.add('dragging');
  }));

  ['dragleave', 'drop'].forEach(t => dropzone.addEventListener(t, e => {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.remove('dragging');
  }));

  dropzone.addEventListener('drop', e => selectFile(e.dataTransfer.files?.[0]));

  // Initialize App
  initTheme();
})();
