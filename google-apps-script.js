/**
 * AHP 問卷 Google Apps Script
 * 功能：接收問卷資料、計算 AHP 權重與 CR 值
 * 
 * 部署步驟：
 * 1. 開啟 Google Sheets，點選「擴充功能」→「Apps Script」
 * 2. 貼上此程式碼
 * 3. 點選「部署」→「新增部署」→ 類型選「網頁應用程式」
 * 4. 設定「誰可以存取」為「所有人」
 * 5. 複製部署網址，貼到 index.html 的 SCRIPT_URL
 */

// Random Index (Saaty) for CR calculation
const RI = {
  1: 0, 2: 0, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32, 8: 1.41,
  9: 1.45, 10: 1.49, 11: 1.51, 12: 1.48, 13: 1.56, 14: 1.57, 15: 1.59, 16: 1.60
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 儲存原始資料
    saveRawData(ss, data);
    
    // 計算 AHP 權重與 CR
    const results = calculateAHP(data);
    
    // 儲存計算結果
    saveResults(ss, data, results);
    
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, results: results }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("AHP Survey API is running.");
}

/**
 * 儲存原始比較資料
 */
function saveRawData(ss, data) {
  let sheet = ss.getSheetByName("原始資料");
  if (!sheet) {
    sheet = ss.insertSheet("原始資料");
    sheet.appendRow(["時間戳記", "姓名", "單位", "專長", "年資", "JSON資料"]);
  }
  
  const meta = data.meta || {};
  sheet.appendRow([
    new Date(),
    meta.name || "",
    meta.org || "",
    meta.field || "",
    meta.years || "",
    JSON.stringify(data)
  ]);
}

/**
 * 計算所有構面的 AHP 權重與 CR
 */
function calculateAHP(data) {
  const results = {};
  const sections = ["dimensions", "A", "B", "C", "D"];
  
  for (const section of sections) {
    const comparisons = data.comparisons[section];
    if (!comparisons || comparisons.length === 0) continue;
    
    // 取得所有項目 ID
    const ids = [...new Set(comparisons.flatMap(c => [c.left, c.right]))].sort();
    const n = ids.length;
    
    // 建立成對比較矩陣
    const matrix = buildMatrix(ids, comparisons);
    
    // 計算權重 (geometric mean method)
    const weights = calculateWeights(matrix, n);
    
    // 計算 CR
    const { lambdaMax, CI, CR } = calculateCR(matrix, weights, n);
    
    // 組合結果
    const weightMap = {};
    ids.forEach((id, i) => {
      weightMap[id] = Math.round(weights[i] * 10000) / 10000;
    });
    
    results[section] = {
      n: n,
      weights: weightMap,
      lambdaMax: Math.round(lambdaMax * 10000) / 10000,
      CI: Math.round(CI * 10000) / 10000,
      CR: Math.round(CR * 10000) / 10000,
      consistent: CR < 0.1
    };
  }
  
  return results;
}

/**
 * 建立成對比較矩陣
 */
function buildMatrix(ids, comparisons) {
  const n = ids.length;
  const matrix = [];
  
  // 初始化為單位矩陣
  for (let i = 0; i < n; i++) {
    matrix[i] = [];
    for (let j = 0; j < n; j++) {
      matrix[i][j] = (i === j) ? 1 : null;
    }
  }
  
  // 填入比較值
  for (const comp of comparisons) {
    const i = ids.indexOf(comp.left);
    const j = ids.indexOf(comp.right);
    const ratio = comp.ahp_ratio_aij;
    
    if (i >= 0 && j >= 0 && ratio != null) {
      matrix[i][j] = ratio;
      matrix[j][i] = 1 / ratio;
    }
  }
  
  // 填補未填的值為 1
  for (let i = 0; i < n; i++) {
    for (let j = 0; j < n; j++) {
      if (matrix[i][j] === null) matrix[i][j] = 1;
    }
  }
  
  return matrix;
}

/**
 * 計算權重 (幾何平均法)
 */
function calculateWeights(matrix, n) {
  const geoMeans = [];
  
  for (let i = 0; i < n; i++) {
    let product = 1;
    for (let j = 0; j < n; j++) {
      product *= matrix[i][j];
    }
    geoMeans[i] = Math.pow(product, 1 / n);
  }
  
  // 正規化
  const sum = geoMeans.reduce((a, b) => a + b, 0);
  return geoMeans.map(g => g / sum);
}

/**
 * 計算一致性比率 CR
 */
function calculateCR(matrix, weights, n) {
  // 計算 Aw (矩陣乘以權重向量)
  const Aw = [];
  for (let i = 0; i < n; i++) {
    let sum = 0;
    for (let j = 0; j < n; j++) {
      sum += matrix[i][j] * weights[j];
    }
    Aw[i] = sum;
  }
  
  // 計算 λmax
  let lambdaMax = 0;
  for (let i = 0; i < n; i++) {
    lambdaMax += Aw[i] / weights[i];
  }
  lambdaMax /= n;
  
  // 計算 CI
  const CI = (lambdaMax - n) / (n - 1);
  
  // 計算 CR
  const ri = RI[n] || 1.5;
  const CR = ri > 0 ? CI / ri : 0;
  
  return { lambdaMax, CI, CR };
}

/**
 * 儲存計算結果
 */
function saveResults(ss, data, results) {
  let sheet = ss.getSheetByName("AHP結果");
  if (!sheet) {
    sheet = ss.insertSheet("AHP結果");
    sheet.appendRow([
      "時間戳記", "姓名", "單位",
      "構面CR", "構面一致", "A權重", "B權重", "C權重", "D權重",
      "A_CR", "A一致", "B_CR", "B一致", "C_CR", "C一致", "D_CR", "D一致"
    ]);
  }
  
  const meta = data.meta || {};
  const dim = results.dimensions || {};
  
  sheet.appendRow([
    new Date(),
    meta.name || "",
    meta.org || "",
    dim.CR || "",
    dim.consistent ? "是" : "否",
    dim.weights?.A || "",
    dim.weights?.B || "",
    dim.weights?.C || "",
    dim.weights?.D || "",
    results.A?.CR || "",
    results.A?.consistent ? "是" : "否",
    results.B?.CR || "",
    results.B?.consistent ? "是" : "否",
    results.C?.CR || "",
    results.C?.consistent ? "是" : "否",
    results.D?.CR || "",
    results.D?.consistent ? "是" : "否"
  ]);
  
  // 儲存詳細權重到各自的工作表
  saveDetailedWeights(ss, "權重_A", results.A, data.meta);
  saveDetailedWeights(ss, "權重_B", results.B, data.meta);
  saveDetailedWeights(ss, "權重_C", results.C, data.meta);
  saveDetailedWeights(ss, "權重_D", results.D, data.meta);
}

/**
 * 儲存各構面的詳細權重
 */
function saveDetailedWeights(ss, sheetName, result, meta) {
  if (!result || !result.weights) return;
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = ["時間戳記", "姓名", "CR", "一致性"];
    const ids = Object.keys(result.weights).sort();
    headers.push(...ids);
    sheet.appendRow(headers);
  }
  
  const row = [
    new Date(),
    meta?.name || "",
    result.CR,
    result.consistent ? "是" : "否"
  ];
  
  const ids = Object.keys(result.weights).sort();
  ids.forEach(id => {
    row.push(result.weights[id]);
  });
  
  sheet.appendRow(row);
}
