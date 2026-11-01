const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

console.log('=== 開始執行辦公室 Excel 萬能工具箱 自動化測試套件 ===\n');

let passCount = 0;
let failCount = 0;

function assert(condition, message) {
  if (condition) {
    console.log(`✅ [PASS] ${message}`);
    passCount++;
  } else {
    console.error(`❌ [FAIL] ${message}`);
    failCount++;
  }
}

// -------------------------------------------------------------
// 測試 1：真實 Excel 檔案二進位讀取與 Sheet 抽取
// -------------------------------------------------------------
const testDir = path.join(__dirname, 'test-data');
const fileOldPath = path.join(testDir, '115-1_排程初稿_真實測試.xlsx');
const fileNewPath = path.join(testDir, '115-1_加退選後定稿_真實測試.xlsx');
const fileOpsPath = path.join(testDir, '115-1_全機構排程與衝突工時_真實測試.xlsx');

assert(fs.existsSync(fileOldPath), '初稿 Excel 檔案存在');
assert(fs.existsSync(fileNewPath), '定稿 Excel 檔案存在');
assert(fs.existsSync(fileOpsPath), '任務排程排程衝突 Excel 檔案存在');

const wbOld = XLSX.readFile(fileOldPath);
const rowsOld = XLSX.utils.sheet_to_json(wbOld.Sheets[wbOld.SheetNames[0]], { defval: "" });
const wbNew = XLSX.readFile(fileNewPath);
const rowsNew = XLSX.utils.sheet_to_json(wbNew.Sheets[wbNew.SheetNames[0]], { defval: "" });

assert(rowsOld.length === 6, `初稿應有 6 筆資料，實際讀取到 ${rowsOld.length} 筆`);
assert(rowsNew.length === 6, `定稿應有 6 筆資料，實際讀取到 ${rowsNew.length} 筆`);

// -------------------------------------------------------------
// 測試 2：雙 Excel 比對引擎測試 (Diff Engine)
// -------------------------------------------------------------
function computeDiff(dataA, dataB, keyCol = '項目代碼', ignoreSpaces = true) {
  const norm = val => {
    let s = String(val ?? '');
    return ignoreSpaces ? s.replace(/\s+/g, ' ').trim() : s;
  };

  const headersSet = new Set();
  [...dataA, ...dataB].forEach(r => Object.keys(r).forEach(k => headersSet.add(k)));
  const headers = Array.from(headersSet);

  const mapA = new Map();
  const mapB = new Map();
  dataA.forEach((r, i) => mapA.set(norm(r[keyCol]) || `_A_${i}`, r));
  dataB.forEach((r, i) => mapB.set(norm(r[keyCol]) || `_B_${i}`, r));

  const allKeys = new Set([...mapA.keys(), ...mapB.keys()]);
  const results = [];
  let countAdded = 0, countRemoved = 0, countModified = 0, countSame = 0;

  allKeys.forEach(k => {
    const inA = mapA.has(k);
    const inB = mapB.has(k);
    if (inA && !inB) {
      countRemoved++;
      results.push({ status: 'removed', statusLabel: '已刪除/停開', key: k, rowA: mapA.get(k), rowB: null, changedCols: {} });
    } else if (!inA && inB) {
      countAdded++;
      results.push({ status: 'added', statusLabel: '新增項目', key: k, rowA: null, rowB: mapB.get(k), changedCols: {} });
    } else {
      const rA = mapA.get(k);
      const rB = mapB.get(k);
      const changedCols = {};
      let isDiff = false;
      headers.forEach(h => {
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

  return { results, headers, countAdded, countRemoved, countModified, countSame };
}

const diffResult = computeDiff(rowsOld, rowsNew, '項目代碼');
assert(diffResult.countSame === 2, `完全相同課程應為 2 門 (CS1001, BA2008)，實際: ${diffResult.countSame}`);
assert(diffResult.countModified === 3, `內容異動應為 3 門 (CS2004, EE1005, GE1102)，實際: ${diffResult.countModified}`);
assert(diffResult.countRemoved === 1, `停開/刪除應為 1 門 (ME3010)，實際: ${diffResult.countRemoved}`);
assert(diffResult.countAdded === 1, `新開/新增應為 1 門 (AI3001)，實際: ${diffResult.countAdded}`);

const cs2004 = diffResult.results.find(r => r.key === 'CS2004');
assert(cs2004 && cs2004.changedCols['場地會議室'] && cs2004.changedCols['場地會議室'].oldVal === 'TR-309' && cs2004.changedCols['場地會議室'].newVal === 'RB-105', 'CS2004 場地會議室異動 (TR-309 -> RB-105) 偵測正確');
assert(cs2004 && cs2004.changedCols['選課人數'] && Number(cs2004.changedCols['選課人數'].oldVal) === 48 && Number(cs2004.changedCols['選課人數'].newVal) === 85, 'CS2004 選課人數異動 (48 -> 85) 偵測正確');

// -------------------------------------------------------------
// 測試 3：任務排程排程衝突與任務排程節次聚合演算法測試 (Conflict Detection & Aggregation)
// -------------------------------------------------------------
function expandTimeSlots(timeStr) {
  const s = String(timeStr || '').trim();
  if (!s) return [];
  const dayMatch = s.match(/^[一二三四五六日]/);
  if (!dayMatch) return [s];
  const day = dayMatch[0];
  const periods = s.slice(1).split('');
  return periods.map(p => `${day}-${p}`);
}

function analyzeConflicts(rows) {
  const roomSlotMap = new Map();
  const teacherSlotMap = new Map();
  const rawConflicts = [];

  rows.forEach(r => {
    const slots = expandTimeSlots(r["上課時間"]);
    const room = r["場地會議室"];
    const teacher = r["負責人員"];

    slots.forEach(slot => {
      // 教室衝突
      const roomKey = `${room}@${slot}`;
      if (room && roomSlotMap.has(roomKey)) {
        const prev = roomSlotMap.get(roomKey);
        if (prev["項目代碼"] !== r["項目代碼"]) {
          rawConflicts.push({
            type: '場地衝突',
            key: `${room}__${[prev["項目代碼"], r["項目代碼"]].sort().join('_')}`,
            day: slot.split('-')[0],
            period: slot.split('-')[1],
            target: room,
            c1: prev,
            c2: r
          });
        }
      } else if (room) {
        roomSlotMap.set(roomKey, r);
      }

      // 人員衝突
      const teacherKey = `${teacher}@${slot}`;
      if (teacher && teacherSlotMap.has(teacherKey)) {
        const prev = teacherSlotMap.get(teacherKey);
        if (prev["項目代碼"] !== r["項目代碼"]) {
          rawConflicts.push({
            type: '人員衝突',
            key: `${teacher}__${[prev["項目代碼"], r["項目代碼"]].sort().join('_')}`,
            day: slot.split('-')[0],
            period: slot.split('-')[1],
            target: `${teacher} 老師`,
            c1: prev,
            c2: r
          });
        }
      } else if (teacher) {
        teacherSlotMap.set(teacherKey, r);
      }
    });
  });

  // 聚合相同衝突群組 (例如 一-2, 一-3, 一-4 合併為 星期一 2、3、4節)
  const aggregated = new Map();
  rawConflicts.forEach(item => {
    const groupKey = `${item.type}__${item.key}__${item.day}`;
    if (!aggregated.has(groupKey)) {
      aggregated.set(groupKey, {
        type: item.type,
        target: item.target,
        day: item.day,
        periods: [],
        c1: item.c1,
        c2: item.c2
      });
    }
    aggregated.get(groupKey).periods.push(item.period);
  });

  return Array.from(aggregated.values()).map(g => ({
    type: g.type,
    timeText: `星期${g.day} 第 ${g.periods.sort().join('、')} 節`,
    target: g.target,
    detail: `【${g.c1["項目代碼"]} ${g.c1["任務名稱"]} (${g.type === '場地衝突' ? g.c1["負責人員"] : g.c1["場地會議室"]})】與【${g.c2["項目代碼"]} ${g.c2["任務名稱"]} (${g.type === '場地衝突' ? g.c2["負責人員"] : g.c2["場地會議室"]})】排程衝突`
  }));
}

const wbOps = XLSX.readFile(fileOpsPath);
const rowsOps = XLSX.utils.sheet_to_json(wbOps.Sheets[wbOps.SheetNames[0]], { defval: "" });
const conflicts = analyzeConflicts(rowsOps);

assert(conflicts.length === 2, `聚合後排程衝突項目應精簡為 2 項 (1 個場地衝突 + 1 個人員衝突)，實際: ${conflicts.length}`);
const roomConflict = conflicts.find(c => c.type === '場地衝突');
assert(roomConflict && roomConflict.timeText.includes('2、3、4'), `場地衝突正確聚合為 2、3、4 節：${roomConflict.timeText}`);
const teacherConflict = conflicts.find(c => c.type === '人員衝突');
assert(teacherConflict && teacherConflict.timeText.includes('3、4'), `人員衝突正確聚合為 3、4 節：${teacherConflict.timeText}`);

// -------------------------------------------------------------
// 測試 4：教師基本任務工時與超額工時核算 (Teacher Hours Calculation)
// -------------------------------------------------------------
function computeTeacherHours(rows) {
  const baseHoursRule = { '教授': 8, '副教授': 9, '助理教授': 9, '講師': 10 };
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
        totalHours: 0
      });
    }
    const st = teacherStats.get(name);
    st.totalHours += Number(r["時數"] || 0);
  });

  return Array.from(teacherStats.values()).map(t => {
    const required = Math.max(0, t.baseHours - t.reduction);
    const overHours = Math.max(0, t.totalHours - required);
    const underHours = Math.max(0, required - t.totalHours);
    return { ...t, required, overHours, underHours };
  });
}

const hoursStats = computeTeacherHours(rowsOps);
const chen = hoursStats.find(t => t.name === '陳建宏');
assert(chen && chen.totalHours === 8 && chen.overHours === 0 && chen.underHours === 0, '陳建宏 (教授 8h，無減授，執行 8h) 符合基準工時');

const lin = hoursStats.find(t => t.name === '林志明');
assert(lin && lin.baseHours === 9 && lin.reduction === 2 && lin.required === 7 && lin.totalHours === 9 && lin.overHours === 2, '林志明 (副教授 9h，減授 2h，應授 7h，執行 9h) 超額工時 +2h 驗證通過');

const wang = hoursStats.find(t => t.name === '王品翔');
assert(wang && wang.baseHours === 9 && wang.totalHours === 6 && wang.underHours === 3, '王品翔 (助理教授 9h，執行 6h) 不足基準工時 -3h 驗證通過');

// -------------------------------------------------------------
// 測試 5：網頁/Word 資料搬移剖析 (Text / HTML Table Parsing)
// -------------------------------------------------------------
function parseTableData(rawText, cleanSpace = true, halfWidth = true) {
  const sanitize = str => {
    let s = String(str ?? '').replace(/\u00A0/g, ' ');
    if (cleanSpace) s = s.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
    if (halfWidth) {
      s = s.replace(/[\uff01-\uff5e]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xfee0)).replace(/\u3000/g, ' ');
    }
    return s;
  };

  const lines = rawText.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  if (!lines.length) return { headers: [], rows: [] };

  const splitRows = lines.map(line => {
    if (line.includes('\t')) return line.split('\t').map(sanitize);
    if (line.includes(',')) return line.split(',').map(sanitize);
    return line.split(/\s{2,}/).map(sanitize);
  });

  const rawHeaders = splitRows[0];
  const maxCols = Math.max(...splitRows.map(r => r.length));
  const headers = Array.from({ length: maxCols }, (_, i) => {
    return (rawHeaders && rawHeaders[i]) ? rawHeaders[i] : `欄位_${i+1}`;
  });

  const dataObjs = splitRows.slice(1).map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i] ?? '');
    return o;
  });

  return { headers, rows: dataObjs };
}

const tabInput = "項目代碼\t任務名稱\t開課系所\t教師\nCS1001\t計算機概論\t資工系\t陳建宏\nCS2004\t資料結構\t資工系\t林志明";
const parsed = parseTableData(tabInput);
assert(parsed.headers.length === 4 && parsed.headers[0] === '項目代碼', 'Tab 分隔文字正確擷取標題欄位');
assert(parsed.rows.length === 2 && parsed.rows[0]['任務名稱'] === '計算機概論', 'Tab 分隔文字正確擷取 2 筆資料');

console.log(`\n測試總結: ${passCount} 通過, ${failCount} 失敗`);
if (failCount === 0) {
  console.log('🎉 所有邏輯驗證測試全部通過！');
} else {
  process.exit(1);
}
