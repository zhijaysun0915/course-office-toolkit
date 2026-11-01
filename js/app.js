// 全域狀態管理
const AppState = {
  // 模組 1: 雙表比對
  diffDataA: [],
  diffDataB: [],
  diffResult: [],
  diffHeaders: [],
  diffFilter: 'all',
  diffSearchQuery: '',
  workbookA: null,
  workbookB: null,

  // 模組 2: 網頁/Word/公文搬移
  extractedRows: [],
  extractedHeaders: [],

  // 模組 3: 辦公室例行事務 (排程衝突/工時/資產盤點)
  scheduleRows: [],
  inventoryRows: []
};

// 工具通知
function showToast(msg) {
  const container = document.getElementById('toast-container');
  const el = document.createElement('div');
  el.className = 'toast';
  el.innerHTML = `<span>✨</span> <span>${msg}</span>`;
  container.appendChild(el);
  setTimeout(() => {
    el.style.opacity = '0';
    setTimeout(() => el.remove(), 300);
  }, 2800);
}

// 頁面頁籤切換
function switchTab(tabId, navEl) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));

  const target = document.getElementById(`tab-${tabId}`);
  if (target) target.classList.add('active');
  if (navEl) navEl.classList.add('active');

  const titles = {
    'diff': ['📊 雙 Excel 智慧比對器', '自動比對兩份工作表、名冊、排程或報表之「新增、刪除、異動儲存格」並一鍵匯出'],
    'extract': ['📋 網頁 / Word / 公文資料搬移器', '從系統網頁公告、Word 表格、Email 或公文直接貼上，自動清理隱藏空白並轉為乾淨 Excel'],
    'course-ops': ['🏢 辦公室例行事務工作台', '場地會議室與人員排程衝突檢測、工時與加班津貼核算、辦公資產與物品盤點管理'],
    'wand-ai': ['🪄 Excel 萬能公式與清洗魔杖', '參考 Wand.Tools 設計：輸入白話文需求，立即產生辦公室常用 Excel 公式與一鍵批次資料清洗'],
    'guide': ['📖 詳細載用說明 ＆ Gemini / Antigravity 客製化改裝手冊', '一步步教您如何線上/離線跨地使用，以及如何將本專案放入您自己的 Gemini / Antigravity 自由改裝']
  };

  if (titles[tabId]) {
    document.getElementById('page-title').textContent = titles[tabId][0];
    document.getElementById('page-desc').textContent = titles[tabId][1];
  }
}

// ============================================================================
// 模組 1：2 個 Excel 檔案智慧比對
// ============================================================================
function handleFileUpload(file, side) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      if (side === 'A') {
        AppState.workbookA = workbook;
      } else {
        AppState.workbookB = workbook;
      }
      renderSheetSelector(workbook, side, file.name);
      loadSheetData(workbook, workbook.SheetNames[0], side, file.name);
    } catch (err) {
      alert('讀取檔案失敗，請確認是否為有效的 Excel (.xlsx, .xls) 或 CSV 檔案！');
    }
  };
  reader.readAsArrayBuffer(file);
}

function renderSheetSelector(workbook, side, fileName) {
  const containerId = side === 'A' ? 'sheet-select-wrap-a' : 'sheet-select-wrap-b';
  const container = document.getElementById(containerId);
  if (!container) return;

  if (workbook.SheetNames.length <= 1) {
    container.innerHTML = '';
    return;
  }

  container.innerHTML = `
    <div style="margin-top:8px;" onclick="event.stopPropagation();">
      <span style="font-size:0.78rem;color:#475569;margin-right:4px;">📑 選擇工作表：</span>
      <select class="sheet-selector" onchange="loadSheetData(AppState.workbook${side}, this.value, '${side}', '${fileName.replace(/'/g, "\\'")}')">
        ${workbook.SheetNames.map(s => `<option value="${s}">${s}</option>`).join('')}
      </select>
    </div>
  `;
}

function loadSheetData(workbook, sheetName, side, fileName = '') {
  if (!workbook || !workbook.Sheets[sheetName]) return;
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
  if (side === 'A') {
    AppState.diffDataA = rows;
    document.getElementById('file-label-a').innerHTML = `✅ 已載入：<strong>${fileName}</strong> [工作表: ${sheetName}] (${rows.length} 筆)`;
  } else {
    AppState.diffDataB = rows;
    document.getElementById('file-label-b').innerHTML = `✅ 已載入：<strong>${fileName}</strong> [工作表: ${sheetName}] (${rows.length} 筆)`;
  }
  updateKeyColumnSelector();
  runExcelDiff();
  showToast(`已成功讀取 ${fileName || ''} (${sheetName}) 共 ${rows.length} 筆資料`);
}

function loadDemoDiff() {
  AppState.diffDataA = JSON.parse(JSON.stringify(DemoData.diffOld));
  AppState.diffDataB = JSON.parse(JSON.stringify(DemoData.diffNew));
  document.getElementById('file-label-a').innerHTML = `✅ 已載入範例：<strong>業務排程初稿.xlsx</strong> (${AppState.diffDataA.length} 筆)`;
  document.getElementById('file-label-b').innerHTML = `✅ 已載入範例：<strong>業務排程定稿.xlsx</strong> (${AppState.diffDataB.length} 筆)`;
  updateKeyColumnSelector();
  runExcelDiff();
  showToast('已載入「版本初稿 vs 最終定稿」比對範例！');
}

function updateKeyColumnSelector() {
  const headersSet = new Set();
  [...AppState.diffDataA, ...AppState.diffDataB].forEach(row => {
    Object.keys(row).forEach(k => headersSet.add(k));
  });
  AppState.diffHeaders = Array.from(headersSet);
  const select = document.getElementById('diff-key-col');
  select.innerHTML = '';
  AppState.diffHeaders.forEach((h, idx) => {
    const opt = document.createElement('option');
    opt.value = h;
    opt.textContent = h + (idx === 0 ? ' (推薦主鍵)' : '');
    select.appendChild(opt);
  });
}

function runExcelDiff() {
  if (!AppState.diffDataA.length && !AppState.diffDataB.length) {
    alert('請先上傳兩個 Excel 檔案，或點擊右上角「載入對照範例」體驗！');
    return;
  }
  const keyCol = document.getElementById('diff-key-col').value || AppState.diffHeaders[0];
  const ignoreSpaces = document.getElementById('diff-ignore-space').checked;

  const norm = val => {
    let s = String(val ?? '');
    return ignoreSpaces ? s.replace(/\s+/g, ' ').trim() : s;
  };

  const mapA = new Map();
  const mapB = new Map();

  AppState.diffDataA.forEach((row, idx) => {
    const k = norm(row[keyCol]) || `__ROW_A_${idx}`;
    mapA.set(k, row);
  });
  AppState.diffDataB.forEach((row, idx) => {
    const k = norm(row[keyCol]) || `__ROW_B_${idx}`;
    mapB.set(k, row);
  });

  const allKeys = new Set([...mapA.keys(), ...mapB.keys()]);
  const results = [];
  let countAdded = 0, countRemoved = 0, countModified = 0, countSame = 0;

  allKeys.forEach(k => {
    const inA = mapA.has(k);
    const inB = mapB.has(k);
    if (inA && !inB) {
      countRemoved++;
      results.push({ status: 'removed', statusLabel: '已刪除/移除', key: k, rowA: mapA.get(k), rowB: null, changedCols: {} });
    } else if (!inA && inB) {
      countAdded++;
      results.push({ status: 'added', statusLabel: '新增項目', key: k, rowA: null, rowB: mapB.get(k), changedCols: {} });
    } else {
      const rA = mapA.get(k);
      const rB = mapB.get(k);
      const changedCols = {};
      let isDiff = false;
      AppState.diffHeaders.forEach(h => {
        if (norm(rA[h]) !== norm(rB[h])) {
          isDiff = true;
          changedCols[h] = { oldVal: rA[h] ?? '', newVal: rB[h] ?? '' };
        }
      });
      if (isDiff) {
        countModified++;
        results.push({ status: 'modified', statusLabel: '內容異動', key: k, rowA: rA, rowB: rB, changedCols });
      } else {
        countSame++;
        results.push({ status: 'same', statusLabel: '無異動', key: k, rowA: rA, rowB: rB, changedCols: {} });
      }
    }
  });

  AppState.diffResult = results;
  document.getElementById('stat-added').textContent = countAdded;
  document.getElementById('stat-removed').textContent = countRemoved;
  document.getElementById('stat-modified').textContent = countModified;
  document.getElementById('stat-same').textContent = countSame;

  renderDiffTable();
}

function setDiffFilter(filter) {
  AppState.diffFilter = filter;
  renderDiffTable();
}

function onDiffSearch(query) {
  AppState.diffSearchQuery = query;
  renderDiffTable();
}

function renderDiffTable() {
  const thead = document.getElementById('diff-thead');
  const tbody = document.getElementById('diff-tbody');
  const headers = AppState.diffHeaders;

  thead.innerHTML = `<tr>
    <th>比對狀態</th>
    ${headers.map(h => `<th>${h}</th>`).join('')}
  </tr>`;

  const q = (AppState.diffSearchQuery || '').trim().toLowerCase();

  const filtered = AppState.diffResult.filter(item => {
    // 狀態篩選
    if (AppState.diffFilter !== 'all') {
      if (AppState.diffFilter === 'diff-only' && item.status === 'same') return false;
      if (AppState.diffFilter !== 'diff-only' && item.status !== AppState.diffFilter) return false;
    }
    // 關鍵字搜尋
    if (!q) return true;
    const row = item.rowB || item.rowA || {};
    const textMatch = Object.values(row).some(v => String(v).toLowerCase().includes(q));
    const statusMatch = item.statusLabel.includes(q);
    const changeMatch = Object.entries(item.changedCols || {}).some(([k, v]) => 
      k.toLowerCase().includes(q) || String(v.oldVal).toLowerCase().includes(q) || String(v.newVal).toLowerCase().includes(q)
    );
    return textMatch || statusMatch || changeMatch;
  });

  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="${headers.length + 1}" style="text-align:center;padding:28px;color:#64748b;">目前篩選或搜尋條件下沒有找到符合的課程資料</td></tr>`;
    return;
  }

  tbody.innerHTML = filtered.map(item => {
    const pillClass = {
      'added': 'pill-added',
      'removed': 'pill-removed',
      'modified': 'pill-modified',
      'same': 'pill-same'
    }[item.status];

    const cellsHtml = headers.map(h => {
      if (item.status === 'added') {
        return `<td>${item.rowB[h] ?? ''}</td>`;
      }
      if (item.status === 'removed') {
        return `<td><del>${item.rowA[h] ?? ''}</del></td>`;
      }
      if (item.status === 'modified' && item.changedCols[h]) {
        const c = item.changedCols[h];
        return `<td class="cell-changed" title="原值：${c.oldVal}">
          <span style="text-decoration:line-through;color:#b91c1c;font-size:0.78rem;margin-right:4px;">${c.oldVal}</span>
          ➔ <strong>${c.newVal}</strong>
        </td>`;
      }
      return `<td>${item.rowB[h] ?? ''}</td>`;
    }).join('');

    return `<tr class="diff-${item.status}">
      <td><span class="pill ${pillClass}">${item.statusLabel}</span></td>
      ${cellsHtml}
    </tr>`;
  }).join('');
}

function exportDiffExcel() {
  if (!AppState.diffResult.length) {
    alert('請先執行比對後再匯出！');
    return;
  }
  const exportRows = AppState.diffResult.map(item => {
    const obj = { "比對狀態": item.statusLabel };
    AppState.diffHeaders.forEach(h => {
      if (item.status === 'added') obj[h] = item.rowB[h] ?? '';
      else if (item.status === 'removed') obj[h] = item.rowA[h] ?? '';
      else if (item.status === 'modified' && item.changedCols[h]) {
        obj[h] = `${item.changedCols[h].oldVal} -> ${item.changedCols[h].newVal}`;
      } else {
        obj[h] = item.rowB[h] ?? '';
      }
    });
    return obj;
  });

  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "雙表比對結果");
  XLSX.writeFile(wb, `Excel差異比對報告_${new Date().toISOString().slice(0,10)}.xlsx`);
  showToast('已匯出 Excel 比對報告！');
}

// ============================================================================
// 模組 2：資料從網頁 / Word / 公文搬來 (智慧擷取與清洗)
// ============================================================================
function loadExtractDemo(type) {
  const pasteBox = document.getElementById('rich-paste-box');
  if (type === 'html') {
    pasteBox.innerHTML = DemoData.sampleHtmlTable;
    showToast('已模擬從「系統網頁/Word表格」貼上內容，系統已自動擷取！');
  } else {
    pasteBox.innerText = DemoData.sampleMessyText;
    showToast('已模擬從「Word公文/Email調補課申請單」貼上純文字，系統已自動拆解欄位！');
  }
  parsePastedContent();
}

function handleDocxUpload(file) {
  if (!file) return;
  if (typeof mammoth === 'undefined') {
    alert('Word 解析元件載入中，請直接開啟 Word 全選複製貼入下方框內即可！');
    return;
  }
  const reader = new FileReader();
  reader.onload = function(e) {
    mammoth.convertToHtml({ arrayBuffer: e.target.result })
      .then(function(result) {
        document.getElementById('rich-paste-box').innerHTML = result.value;
        parsePastedContent();
        showToast(`已成功從 Word 檔 (${file.name}) 擷取表格與文字！`);
      })
      .catch(function() {
        alert('Word 解析失敗，建議直接打開 Word 複製表格貼到下方區域！');
      });
  };
  reader.readAsArrayBuffer(file);
}

function parsePastedContent() {
  const pasteBox = document.getElementById('rich-paste-box');
  const cleanSpace = document.getElementById('ext-clean-space').checked;
  const halfWidth = document.getElementById('ext-half-width').checked;

  const sanitize = str => {
    let s = String(str ?? '').replace(/\u00A0/g, ' ');
    if (cleanSpace) s = s.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
    if (halfWidth) {
      s = s.replace(/[\uff01-\uff5e]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xfee0)).replace(/\u3000/g, ' ');
    }
    return s;
  };

  // 1. 優先檢查是否有 HTML <table> (從網頁或 Word 複製表格貼上時會帶 HTML table)
  const tables = pasteBox.querySelectorAll('table');
  if (tables.length > 0) {
    const rows = [];
    tables.forEach(tbl => {
      tbl.querySelectorAll('tr').forEach(tr => {
        const cells = Array.from(tr.querySelectorAll('th, td')).map(td => sanitize(td.innerText || td.textContent));
        if (cells.some(c => c !== '')) {
          rows.push(cells);
        }
      });
    });

    if (rows.length > 0) {
      const rawHeaders = rows[0];
      const maxCols = Math.max(...rows.map(r => r.length));
      const headers = Array.from({ length: maxCols }, (_, i) => (rawHeaders && rawHeaders[i]) ? rawHeaders[i] : `欄位_${i+1}`);
      const dataObjs = rows.slice(1).map(r => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = r[i] ?? '');
        return obj;
      });
      AppState.extractedHeaders = headers;
      AppState.extractedRows = dataObjs;
      renderExtractedTable();
      return;
    }
  }

  // 2. 若為純文字：透過複製節點將 <br>、<p>、<div> 轉換為標準換行符號，防止瀏覽器 DOM 吞噬換行
  const clone = pasteBox.cloneNode(true);
  clone.querySelectorAll('br').forEach(br => br.replaceWith('\n'));
  clone.querySelectorAll('p, div, tr').forEach(p => p.after('\n'));
  const rawText = (clone.textContent || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = rawText.split('\n').map(l => l.trim()).filter(Boolean);
  if (!lines.length) {
    AppState.extractedHeaders = [];
    AppState.extractedRows = [];
    renderExtractedTable();
    return;
  }

  // 檢查是否包含「キー：値」格式 (例如：申請人：陳建宏 | 項目代碼：CS1001)
  const kvLines = lines.filter(l => l.includes('：') || l.includes(':'));
  if (kvLines.length >= 2 && (rawText.includes('|') || rawText.includes('｜') || rawText.includes('；'))) {
    const headersSet = new Set();
    const parsedObjs = [];
    lines.forEach(line => {
      if (!line.includes('：') && !line.includes(':')) return;
      const segments = line.split(/[|｜；;]/);
      const obj = {};
      segments.forEach(seg => {
        const m = seg.split(/[:：]/);
        if (m.length >= 2) {
          const k = sanitize(m[0]);
          const v = sanitize(m.slice(1).join(':'));
          if (k) {
            headersSet.add(k);
            obj[k] = v;
          }
        }
      });
      if (Object.keys(obj).length > 0) parsedObjs.push(obj);
    });
    if (parsedObjs.length > 0) {
      AppState.extractedHeaders = Array.from(headersSet);
      AppState.extractedRows = parsedObjs;
      renderExtractedTable();
      return;
    }
  }

  // 3. 一般分隔符號 (Tab, 逗號, 或 2 個以上空白)
  const splitRows = lines.map(line => {
    if (line.includes('\t')) return line.split('\t').map(sanitize);
    if (line.includes(',')) return line.split(',').map(sanitize);
    return line.split(/\s{2,}/).map(sanitize);
  });

  const rawHeaders = splitRows[0];
  const maxCols = Math.max(...splitRows.map(r => r.length));
  const headers = Array.from({ length: maxCols }, (_, i) => (rawHeaders && rawHeaders[i]) ? rawHeaders[i] : `欄位_${i+1}`);

  const dataObjs = splitRows.slice(1).map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i] ?? '');
    return o;
  });

  AppState.extractedHeaders = headers;
  AppState.extractedRows = dataObjs;
  renderExtractedTable();
}

function renderExtractedTable() {
  const thead = document.getElementById('ext-thead');
  const tbody = document.getElementById('ext-tbody');
  const countBadge = document.getElementById('ext-count-badge');

  const headers = AppState.extractedHeaders;
  const rows = AppState.extractedRows;
  countBadge.textContent = `共擷取 ${rows.length} 筆資料、${headers.length} 個欄位`;

  if (!headers.length) {
    thead.innerHTML = '';
    tbody.innerHTML = `<tr><td style="text-align:center;padding:30px;color:#64748b;">尚無資料，請在上方貼上網頁表格、Word 表格或公文清單</td></tr>`;
    return;
  }

  thead.innerHTML = `<tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr>`;
  tbody.innerHTML = rows.map((r, rIdx) => `
    <tr>
      ${headers.map(h => `<td contenteditable="true" onblur="updateExtractedCell(${rIdx}, '${h}', this.innerText)">${r[h] ?? ''}</td>`).join('')}
    </tr>
  `).join('');
}

function updateExtractedCell(rowIdx, colName, val) {
  if (AppState.extractedRows[rowIdx]) {
    AppState.extractedRows[rowIdx][colName] = val.trim();
  }
}

function exportExtractedExcel() {
  if (!AppState.extractedRows.length) {
    alert('目前沒有可匯出的資料！');
    return;
  }
  const ws = XLSX.utils.json_to_sheet(AppState.extractedRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "擷取整理資料");
  XLSX.writeFile(wb, `網頁Word轉Excel_${new Date().toISOString().slice(0,10)}.xlsx`);
  showToast('已成功下載乾淨的 Excel 檔案！');
}

function sendExtractedToDiff(side) {
  if (!AppState.extractedRows.length) {
    alert('請先擷取資料後再傳送！');
    return;
  }
  const copy = JSON.parse(JSON.stringify(AppState.extractedRows));
  if (side === 'A') {
    AppState.diffDataA = copy;
    document.getElementById('file-label-a').innerHTML = `✅ 已從擷取器載入 (${copy.length} 筆)`;
  } else {
    AppState.diffDataB = copy;
    document.getElementById('file-label-b').innerHTML = `✅ 已從擷取器載入 (${copy.length} 筆)`;
  }
  updateKeyColumnSelector();
  showToast(`已將擷取表格送入「雙 Excel 比對器」檔案 ${side}！`);
}

// ============================================================================
// 模組 3：辦公室例行事務（排程/場地衝突檢測、工時統計、資產用品盤點）
// ============================================================================
function initCourseOps() {
  AppState.scheduleRows = JSON.parse(JSON.stringify(DemoData.scheduleCheckSample));
  AppState.inventoryRows = JSON.parse(JSON.stringify(DemoData.inventorySample));
  analyzeScheduleAndHours();
  renderInventoryTable();
}

function handleScheduleUpload(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, { type: 'array' });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
    AppState.scheduleRows = rows;
    analyzeScheduleAndHours();
    showToast(`已匯入排程/工時資料共 ${rows.length} 筆，並完成衝突檢測與工時核算！`);
  };
  reader.readAsArrayBuffer(file);
}

// 拆解「一234」或「二345」為 ['一-2', '一-3', '一-4']
function expandTimeSlots(timeStr) {
  const s = String(timeStr || '').trim();
  if (!s) return [];
  const dayMatch = s.match(/^[一二三四五六日]/);
  if (!dayMatch) return [s];
  const day = dayMatch[0];
  const periods = s.slice(1).split('');
  return periods.map(p => `${day}-${p}`);
}

function analyzeScheduleAndHours() {
  const rows = AppState.scheduleRows;
  const roomSlotMap = new Map();
  const teacherSlotMap = new Map();
  const conflicts = [];

  rows.forEach(r => {
    const slots = expandTimeSlots(r["上課時間"]);
    const room = r["場地會議室"];
    const teacher = r["負責人員"];

    slots.forEach(slot => {
      const parts = slot.split('-');
      const day = parts[0] || '';
      const period = parts[1] || '';

      // 檢查場地衝突
      const roomKey = `${room}@${slot}`;
      if (room && roomSlotMap.has(roomKey)) {
        const prev = roomSlotMap.get(roomKey);
        if (prev["項目代碼"] !== r["項目代碼"]) {
          const coursesKey = [prev["項目代碼"], r["項目代碼"]].sort().join('_');
          conflicts.push({
            type: '場地衝突',
            targetKey: `${room}__${coursesKey}`,
            targetName: room,
            day: day,
            period: period,
            prev: prev,
            curr: r
          });
        }
      } else if (room) {
        roomSlotMap.set(roomKey, r);
      }

      // 檢查人員衝突
      const teacherKey = `${teacher}@${slot}`;
      if (teacher && teacherSlotMap.has(teacherKey)) {
        const prev = teacherSlotMap.get(teacherKey);
        if (prev["項目代碼"] !== r["項目代碼"]) {
          const coursesKey = [prev["項目代碼"], r["項目代碼"]].sort().join('_');
          conflicts.push({
            type: '人員衝突',
            targetKey: `${teacher}__${coursesKey}`,
            targetName: `${teacher} 老師`,
            day: day,
            period: period,
            prev: prev,
            curr: r
          });
        }
      } else if (teacher) {
        teacherSlotMap.set(teacherKey, r);
      }
    });
  });

  // 聚合相同衝突群組 (同一星期幾、同一對象、相同課程之連續或離散多節排程衝突聚合)
  const aggregated = new Map();
  conflicts.forEach(c => {
    const groupKey = `${c.type}__${c.targetKey}__${c.day}`;
    if (!aggregated.has(groupKey)) {
      aggregated.set(groupKey, {
        type: c.type,
        targetName: c.targetName,
        day: c.day,
        periods: [],
        prev: c.prev,
        curr: c.curr
      });
    }
    aggregated.get(groupKey).periods.push(c.period);
  });

  const conflictList = Array.from(aggregated.values());

  // 渲染排程衝突報告
  const conflictBody = document.getElementById('conflict-tbody');
  document.getElementById('conflict-count').textContent = conflictList.length;
  if (!conflictList.length) {
    conflictBody.innerHTML = `<tr><td colspan="3" style="text-align:center;color:#16a34a;padding:18px;font-weight:700;">🎉 太棒了！目前排程無任何場地或人員衝突！</td></tr>`;
  } else {
    conflictBody.innerHTML = conflictList.map(g => `
      <tr class="diff-removed">
        <td><span class="pill pill-removed">⚠️ ${g.type}</span></td>
        <td>
          <strong>${g.targetName}</strong>
          <div style="font-size:0.78rem;color:#b91c1c;margin-top:2px;">星期${g.day} 第 ${g.periods.sort().join('、')} 節</div>
        </td>
        <td>
          【${g.prev["項目代碼"]} ${g.prev["任務名稱"]} (${g.type === '場地衝突' ? g.prev["負責人員"] : g.prev["場地會議室"]})】與
          【${g.curr["項目代碼"]} ${g.curr["任務名稱"]} (${g.type === '場地衝突' ? g.curr["負責人員"] : g.curr["場地會議室"]})】同時排定
        </td>
      </tr>
    `).join('');
  }

  // 核算人員工時/津貼（基本標準時數規則：教授/主管 8、副教授/副理 9、助理教授 9、講師/專員 10）
  const baseHoursRule = { '教授': 8, '副教授': 9, '助理教授': 9, '講師': 10, '經理': 8, '主管': 8, '工程師': 8, '專員': 8 };
  const teacherStats = new Map();

  rows.forEach(r => {
    const name = r["負責人員"];
    if (!name) return;
    if (!teacherStats.has(name)) {
      const title = r["職稱"] || '教授';
      teacherStats.set(name, {
        name,
        title,
        baseHours: baseHoursRule[title] ?? 8,
        reduction: Number(r["減免時數"] || 0),
        totalHours: 0,
        courses: []
      });
    }
    const st = teacherStats.get(name);
    const hrs = Number(r["時數"] || 0);
    st.totalHours += hrs;
    st.courses.push(`${r["任務名稱"]}(${hrs}h)`);
  });

  const hoursBody = document.getElementById('hours-tbody');
  hoursBody.innerHTML = Array.from(teacherStats.values()).map(t => {
    const required = Math.max(0, t.baseHours - t.reduction);
    const overHours = Math.max(0, t.totalHours - required);
    const underHours = Math.max(0, required - t.totalHours);
    let statusHtml = `<span class="pill pill-added">符合基準工時</span>`;
    if (overHours > 0) statusHtml = `<span class="pill pill-modified">超額工時 +${overHours} 時</span>`;
    if (underHours > 0) statusHtml = `<span class="pill pill-removed">不足基本 -${underHours} 時</span>`;

    return `<tr>
      <td><strong>${t.name}</strong></td>
      <td>${t.title}</td>
      <td>${t.baseHours}</td>
      <td>${t.reduction}</td>
      <td><strong>${required}</strong></td>
      <td style="font-weight:700;color:#1e40af;">${t.totalHours}</td>
      <td>${statusHtml}</td>
      <td style="font-size:0.8rem;color:#475569;">${t.courses.join('、')}</td>
    </tr>`;
  }).join('');
}

function renderInventoryTable() {
  const filterCat = document.getElementById('inv-filter-cat').value;
  const tbody = document.getElementById('inv-tbody');
  const rows = AppState.inventoryRows.filter(r => filterCat === 'all' || r["管理類別"] === filterCat);

  tbody.innerHTML = rows.map((r, idx) => {
    const isShort = Number(r["實盤數量"]) < Number(r["應有數量"]) || r["狀態"] !== '正常';
    return `<tr class="${isShort ? 'diff-modified' : ''}">
      <td>${r["編號"]}</td>
      <td>${r["管理類別"]}</td>
      <td><strong>${r["存放地點"]}</strong></td>
      <td>${r["物品名稱"]}</td>
      <td>${r["應有數量"]}</td>
      <td><input type="number" value="${r["實盤數量"]}" style="width:70px;padding:4px 8px;" onchange="updateInvQty(${idx}, this.value)"></td>
      <td><span class="pill ${isShort ? 'pill-modified' : 'pill-added'}">${r["狀態"]}</span></td>
      <td>${r["保管/檢查人"]}</td>
    </tr>`;
  }).join('');
}

function updateInvQty(idx, val) {
  const item = AppState.inventoryRows[idx];
  if (!item) return;
  item["實盤數量"] = Number(val);
  item["狀態"] = item["實盤數量"] < item["應有數量"] ? '缺件/待補' : '正常';
  renderInventoryTable();
  showToast(`已更新「${item["物品名稱"]}」實盤數量！`);
}

function addInventoryItem() {
  const loc = document.getElementById('new-inv-loc').value.trim();
  const name = document.getElementById('new-inv-name').value.trim();
  const cat = document.getElementById('new-inv-cat').value;
  const qty = Number(document.getElementById('new-inv-qty').value || 1);
  const owner = document.getElementById('new-inv-owner').value.trim() || '行政盤點人員';

  if (!loc || !name) {
    alert('請輸入存放地點/會議室與物品名稱！');
    return;
  }

  AppState.inventoryRows.unshift({
    "編號": `EQ-00${AppState.inventoryRows.length + 1}`,
    "管理類別": cat,
    "存放地點": loc,
    "物品名稱": name,
    "應有數量": qty,
    "實盤數量": qty,
    "狀態": "正常",
    "保管/檢查人": owner,
    "最後檢查日": new Date().toISOString().slice(0, 10)
  });

  document.getElementById('new-inv-loc').value = '';
  document.getElementById('new-inv-name').value = '';
  renderInventoryTable();
  showToast(`已新增資產/行政用品：${name}`);
}

function exportInventoryExcel() {
  const ws = XLSX.utils.json_to_sheet(AppState.inventoryRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "辦公室資產與用品清冊");
  XLSX.writeFile(wb, `辦公室_資產與用品盤點表_${new Date().toISOString().slice(0,10)}.xlsx`);
  showToast('已匯出辦公室資產與用品盤點 Excel！');
}

// ============================================================================
// 模組 4：Wand.Tools 風格 Excel 公式與資料魔杖
// ============================================================================
function renderFormulaCards(query = '') {
  const q = query.trim().toLowerCase();
  const list = DemoData.formulaTemplates.filter(item => {
    if (!q) return true;
    return item.title.toLowerCase().includes(q) ||
           item.desc.toLowerCase().includes(q) ||
           item.category.toLowerCase().includes(q) ||
           item.tags.some(t => t.toLowerCase().includes(q));
  });

  const container = document.getElementById('formula-cards-container');
  container.innerHTML = list.map((item, i) => `
    <div class="card" style="margin-bottom:14px;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
        <span class="pill pill-same">${item.category}</span>
        <span style="font-size:0.78rem;color:#64748b;">支援 Excel / Google 試算表</span>
      </div>
      <h4 style="font-size:1rem;margin-bottom:4px;color:#0f172a;">${item.title}</h4>
      <p style="font-size:0.85rem;color:#475569;margin-bottom:8px;">${item.desc}</p>
      <div class="formula-box">
        <code id="formula-code-${i}">${item.formula}</code>
        <button class="btn btn-sm btn-primary" onclick="copyText('${item.formula.replace(/'/g, "\\'")}')">📋 複製公式</button>
      </div>
    </div>
  `).join('');
}

function generateCustomFormula() {
  const colMatch = document.getElementById('wand-col-lookup').value || 'A2';
  const sheetName = document.getElementById('wand-sheet-name').value || '工作排程總表';
  const colReturn = document.getElementById('wand-col-return').value || 'F:F';
  const taskType = document.getElementById('wand-task-type').value;

  let formula = '';
  let explain = '';

  if (taskType === 'xlookup') {
    formula = `=XLOOKUP(${colMatch}, ${sheetName}!A:A, ${sheetName}!${colReturn}, "查無資料")`;
    explain = `在「${sheetName}」的 A 欄尋找 ${colMatch}，找到後傳回對應的 ${colReturn} 欄內容；若找不到則顯示「查無資料」。`;
  } else if (taskType === 'countif') {
    formula = `=IF(COUNTIF(${sheetName}!A:A, ${colMatch})>0, "已存在", "缺少/新增")`;
    explain = `檢查 ${colMatch} 是否存在於「${sheetName}」的 A 欄中，常用於兩份排程或人員清冊核對。`;
  } else if (taskType === 'sumifs') {
    formula = `=SUMIFS(${sheetName}!${colReturn}, ${sheetName}!A:A, ${colMatch})`;
    explain = `將「${sheetName}」中所有符合 ${colMatch}（例如同一位老師或同一系所）的 ${colReturn} 欄數值（如工時時數、鐘點費）自動加總。`;
  } else if (taskType === 'clean') {
    formula = `=TRIM(CLEAN(SUBSTITUTE(${colMatch}, CHAR(160), " ")))`;
    explain = `徹底清除 ${colMatch} 儲存格中從網頁或 Word 複製帶來的不換行空白與隱藏控制字元。`;
  }

  document.getElementById('custom-formula-output').textContent = formula;
  document.getElementById('custom-formula-explain').textContent = explain;
  showToast('🪄 已根據您的欄位參數自動生成專屬 Excel 公式！');
}

function copyText(txt) {
  navigator.clipboard.writeText(txt).then(() => {
    showToast(`已複製到剪貼簿：${txt}`);
  });
}

function setupDropzone(dropzoneEl, onFileDrop) {
  if (!dropzoneEl) return;
  ['dragenter', 'dragover'].forEach(eventName => {
    dropzoneEl.addEventListener(eventName, e => {
      e.preventDefault();
      e.stopPropagation();
      dropzoneEl.classList.add('dragover');
    });
  });
  ['dragleave', 'drop'].forEach(eventName => {
    dropzoneEl.addEventListener(eventName, e => {
      e.preventDefault();
      e.stopPropagation();
      dropzoneEl.classList.remove('dragover');
    });
  });
  dropzoneEl.addEventListener('drop', e => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files.length > 0) {
      onFileDrop(dt.files[0]);
    }
  });
}

// 初始化綁定
window.addEventListener('DOMContentLoaded', () => {
  // 嘗試還原使用者上次未完成的工作階段，若無則載入預設範例
  const hasRestored = restoreWorkspaceFromLocal();
  if (!hasRestored) {
    loadDemoDiff();
    loadExtractDemo('html');
    initCourseOps();
  }
  renderFormulaCards();
  generateCustomFormula();

  // 綁定檔案拖曳上傳 (Drag and Drop)
  setupDropzone(document.getElementById('dropzone-a'), file => handleFileUpload(file, 'A'));
  setupDropzone(document.getElementById('dropzone-b'), file => handleFileUpload(file, 'B'));
  setupDropzone(document.querySelector('label[for="upload-excel-a"]'), file => handleFileUpload(file, 'A'));
  setupDropzone(document.querySelector('label[for="upload-excel-b"]'), file => handleFileUpload(file, 'B'));

  // 綁定貼上區即時解析 (直接擷取剪貼簿純文字與 HTML 表格)
  const pasteBox = document.getElementById('rich-paste-box');
  if (pasteBox) {
    pasteBox.addEventListener('paste', e => {
      setTimeout(parsePastedContent, 60);
    });
    pasteBox.addEventListener('input', () => {
      parsePastedContent();
    });
  }
});

// ============================================================================
// 模組 5：Google Drive 雲端匯入 & 跨地移動辦公 (在家 / 機構辦公室 / 公司) 同步引擎
// ============================================================================
const STORAGE_KEY = 'OFFICE_EXCEL_TOOLKIT_WORKSPACE_V1';

// 自動儲存工作階段至本機 localStorage (重新整理或離線不遺失)
function saveWorkspaceToLocal(silent = true) {
  try {
    const payload = {
      updatedAt: new Date().toISOString(),
      diffDataA: AppState.diffDataA,
      diffDataB: AppState.diffDataB,
      extractedHeaders: AppState.extractedHeaders,
      extractedRows: AppState.extractedRows,
      scheduleRows: AppState.scheduleRows,
      inventoryRows: AppState.inventoryRows,
      gdriveScriptUrl: localStorage.getItem('GDRIVE_APPS_SCRIPT_URL') || ''
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    const timeStr = new Date().toLocaleTimeString('zh-TW', { hour12: false });
    const statusEl = document.getElementById('sync-status-text');
    if (statusEl) {
      statusEl.textContent = `已自動儲存本機進度 (${timeStr})`;
    }
    if (!silent) showToast('💾 工作台進度已儲存！');
  } catch (e) {
    console.warn('儲存工作區失敗:', e);
  }
}

// 載入上次儲存的工作區
function restoreWorkspaceFromLocal() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    const data = JSON.parse(raw);
    if (data.diffDataA?.length || data.diffDataB?.length) {
      AppState.diffDataA = data.diffDataA || [];
      AppState.diffDataB = data.diffDataB || [];
      document.getElementById('file-label-a').innerHTML = `✅ 已還原上次檔案 A (${AppState.diffDataA.length} 筆)`;
      document.getElementById('file-label-b').innerHTML = `✅ 已還原上次檔案 B (${AppState.diffDataB.length} 筆)`;
      updateKeyColumnSelector();
      runExcelDiff();
    }
    if (data.extractedRows?.length) {
      AppState.extractedHeaders = data.extractedHeaders || [];
      AppState.extractedRows = data.extractedRows || [];
      renderExtractedTable();
    }
    if (data.scheduleRows?.length) {
      AppState.scheduleRows = data.scheduleRows;
      analyzeScheduleAndHours();
    }
    if (data.inventoryRows?.length) {
      AppState.inventoryRows = data.inventoryRows;
      renderInventoryTable();
    }
    return true;
  } catch (e) {
    return false;
  }
}

// 匯出完整移動辦公快照檔 (.json) — 可放入 Google Drive 雲端硬碟資料夾跨電腦接續工作
function exportWorkspaceSnapshot() {
  const payload = {
    version: '1.2.0',
    exportedAt: new Date().toISOString(),
    note: '辦公室萬能工具箱 移動辦公雲端快照檔 (可放於 Google Drive 在家/辦公室/差旅跨裝置載入)',
    diffDataA: AppState.diffDataA,
    diffDataB: AppState.diffDataB,
    extractedHeaders: AppState.extractedHeaders,
    extractedRows: AppState.extractedRows,
    scheduleRows: AppState.scheduleRows,
    inventoryRows: AppState.inventoryRows
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `辦公室工作台快照_${new Date().toISOString().slice(0,10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('☁️ 已匯出「移動辦公快照檔」，可放進 Google Drive 隨時跨地點載入！');
}

// 匯入移動辦公快照檔 (.json)
function importWorkspaceSnapshot(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = JSON.parse(e.target.result);
      applyWorkspacePayload(data);
      saveWorkspaceToLocal(true);
      showToast(`🎉 已成功還原工作台快照 (${file.name})！所有比對、排程表與盤點資料皆已就緒。`);
    } catch (err) {
      alert('快照檔案格式不正確，請選擇有效的 .json 工作台快照檔！');
    }
  };
  reader.readAsText(file);
}

function applyWorkspacePayload(data) {
  if (data.diffDataA) AppState.diffDataA = data.diffDataA;
  if (data.diffDataB) AppState.diffDataB = data.diffDataB;
  if (AppState.diffDataA.length || AppState.diffDataB.length) {
    document.getElementById('file-label-a').innerHTML = `✅ 雲端快照檔案 A (${AppState.diffDataA.length} 筆)`;
    document.getElementById('file-label-b').innerHTML = `✅ 雲端快照檔案 B (${AppState.diffDataB.length} 筆)`;
    updateKeyColumnSelector();
    runExcelDiff();
  }
  if (data.extractedHeaders && data.extractedRows) {
    AppState.extractedHeaders = data.extractedHeaders;
    AppState.extractedRows = data.extractedRows;
    renderExtractedTable();
  }
  if (data.scheduleRows) {
    AppState.scheduleRows = data.scheduleRows;
    analyzeScheduleAndHours();
  }
  if (data.inventoryRows) {
    AppState.inventoryRows = data.inventoryRows;
    renderInventoryTable();
  }
}

// 解析 Google Sheets / Google Drive 網址或檔案 ID，轉為可直接讀取的匯出連結
function extractGoogleFileId(urlOrId) {
  const s = String(urlOrId || '').trim();
  if (!s) return null;
  const matchD = s.match(/\/d\/([a-zA-Z0-9_-]{15,})/);
  if (matchD) return matchD[1];
  const matchId = s.match(/[?&]id=([a-zA-Z0-9_-]{15,})/);
  if (matchId) return matchId[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;
  return null;
}

// 從 Google Drive / Google 試算表網址直接匯入資料
async function importFromGoogleDriveUrl(targetModule) {
  let inputId = '';
  if (targetModule === 'diff-A') inputId = 'gdrive-url-diff-a';
  else if (targetModule === 'diff-B') inputId = 'gdrive-url-diff-b';
  else if (targetModule === 'extract') inputId = 'gdrive-url-extract';
  else if (targetModule === 'schedule') inputId = 'gdrive-url-schedule';

  const rawUrl = document.getElementById(inputId)?.value.trim();
  if (!rawUrl) {
    alert('請貼上 Google 試算表 (Google Sheets) 或 Google Drive 分享連結！');
    return;
  }

  const fileId = extractGoogleFileId(rawUrl);
  // 擷取 gid (若連結內指定了特定工作表分頁)
  const gidMatch = rawUrl.match(/[#&?]gid=([0-9]+)/);
  const gidParam = gidMatch ? `&gid=${gidMatch[1]}` : '';

  // 構造讀取候選清單 (優先嘗試 Google Sheets xlsx / csv 匯出端點)
  const candidateUrls = [];
  if (fileId) {
    candidateUrls.push(`https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx${gidParam}`);
    candidateUrls.push(`https://docs.google.com/spreadsheets/d/${fileId}/export?format=csv${gidParam}`);
    candidateUrls.push(`https://docs.google.com/document/d/${fileId}/export?format=txt`);
  } else {
    candidateUrls.push(rawUrl);
  }

  showToast('☁️ 正在從 Google Drive 讀取雲端資料...');

  let loadedBuffer = null;
  for (const url of candidateUrls) {
    try {
      const resp = await fetch(url);
      if (resp.ok) {
        loadedBuffer = await resp.arrayBuffer();
        break;
      }
    } catch (e) {
      // 繼續嘗試下一個端點
    }
  }

  if (!loadedBuffer) {
    alert(
      '⚠️ 無法直接讀取該 Google Drive 連結。\n\n' +
      '請確認：\n' +
      '1. 該 Google 試算表已設定為「知道連結的任何人都能檢視」；或\n' +
      '2. 若為私人檔案，您可以在電腦安裝 Google 雲端硬碟桌面版 (G: 槽)，直接從檔案選擇器選取 Google Drive 資料夾內的檔案！'
    );
    return;
  }

  try {
    const data = new Uint8Array(loadedBuffer);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheet = workbook.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: "" });

    if (!rows.length) {
      alert('已連線至 Google 檔案，但內容似乎是空的！');
      return;
    }

    if (targetModule === 'diff-A') {
      AppState.diffDataA = rows;
      document.getElementById('file-label-a').innerHTML = `☁️ 已從 Google Drive 載入 (${rows.length} 筆)`;
      updateKeyColumnSelector();
      runExcelDiff();
    } else if (targetModule === 'diff-B') {
      AppState.diffDataB = rows;
      document.getElementById('file-label-b').innerHTML = `☁️ 已從 Google Drive 載入 (${rows.length} 筆)`;
      updateKeyColumnSelector();
      runExcelDiff();
    } else if (targetModule === 'extract') {
      AppState.extractedHeaders = Object.keys(rows[0]);
      AppState.extractedRows = rows;
      renderExtractedTable();
    } else if (targetModule === 'schedule') {
      AppState.scheduleRows = rows;
      analyzeScheduleAndHours();
    }

    saveWorkspaceToLocal(true);
    showToast(`✅ 成功從 Google Drive 匯入 ${rows.length} 筆資料！`);
  } catch (err) {
    alert('解析雲端檔案失敗，請確認連結是否為 Google 試算表或 Excel/CSV 檔案。');
  }
}

// 開關雲端移動辦公設定彈窗
function toggleCloudModal(open) {
  const modal = document.getElementById('cloud-modal');
  if (!modal) return;
  if (open) {
    document.getElementById('gas-webapp-url').value = localStorage.getItem('GDRIVE_APPS_SCRIPT_URL') || '';
    modal.classList.add('open');
  } else {
    modal.classList.remove('open');
  }
}

// 開關版本歷程彈窗
function toggleChangelogModal(open) {
  const modal = document.getElementById('changelog-modal');
  if (!modal) return;
  if (open) {
    modal.classList.add('open');
  } else {
    modal.classList.remove('open');
  }
}

// 使用 Google Apps Script Web App 儲存/讀取私人 Google Drive 進度 (免設公開權限)
async function syncWithGoogleAppsScript(mode) {
  const url = document.getElementById('gas-webapp-url').value.trim();
  if (!url) {
    alert('請先填入您的 Google Apps Script 網址，或直接使用左側「匯出/載入雲端快照檔 (.json)」放入 Google Drive！');
    return;
  }
  localStorage.setItem('GDRIVE_APPS_SCRIPT_URL', url);

  try {
    if (mode === 'push') {
      showToast('☁️ 正在備份工作台資料至您的 Google Drive...');
      const payload = {
        action: 'save',
        updatedAt: new Date().toISOString(),
        diffDataA: AppState.diffDataA,
        diffDataB: AppState.diffDataB,
        extractedHeaders: AppState.extractedHeaders,
        extractedRows: AppState.extractedRows,
        scheduleRows: AppState.scheduleRows,
        inventoryRows: AppState.inventoryRows
      };
      await fetch(url, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload)
      });
      showToast('✅ 已發送雲端備份至您的 Google Drive！');
    } else {
      showToast('☁️ 正在從您的 Google Drive 拉取最新進度...');
      const resp = await fetch(`${url}?action=load`);
      const data = await resp.json();
      applyWorkspacePayload(data);
      saveWorkspaceToLocal(true);
      showToast('🎉 已從 Google Drive 同步最新工作台資料！');
    }
  } catch (e) {
    alert('連線至 Google Apps Script 失敗，請檢查網址是否正確且已部署為「所有人可存取」。');
  }
}
