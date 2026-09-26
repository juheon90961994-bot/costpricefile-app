const COST_TABLE = 'cost';
const ITEM_COLUMN = '품목';
const PRICE_COLUMN = '원화단가';
const MANAGER_COLUMN = '담당자';
const VENDOR_COLUMN = '거래처명';
const UNIT_COLUMN = '단위';
const DOLLAR_PRICE_COLUMN = '달러단가';
const PREVIEW_LIMIT = 50;
const PAGE_SIZE = 1000;
const DEFAULT_SUPABASE_URL = window.COST_APP_CONFIG?.supabaseUrl || '';
const DEFAULT_SUPABASE_PUBLISHABLE_KEY = window.COST_APP_CONFIG?.supabasePublishableKey || '';
const COST_UPLOAD_COLUMNS = [
  { target: '품목명', sources: ['품목명'] },
  { target: '회사', sources: ['회사'] },
  { target: '거래처명', sources: ['거래처명', '거래처', '주거래처'] },
  { target: '담당자', sources: ['담당자명', '담당자'] },
  { target: '단위', sources: ['단위'] },
  { target: '달러단가', sources: ['달러단가'], numeric: true },
  { target: '원화단가', sources: ['원화단가', '원가단가'], numeric: true },
];

const uploadWorkspace = document.querySelector('#uploadWorkspace');
const uploadInput = document.querySelector('#costUploadInput');
const uploadDropzone = document.querySelector('#costUploadDropzone');
const uploadFileName = document.querySelector('#costUploadFileName');
const analyzeButton = document.querySelector('#analyzeUpload');
const executeButton = document.querySelector('#executeUpload');
const uploadMessage = document.querySelector('#uploadMessage');
const preview = document.querySelector('#uploadPreview');
const previewBody = document.querySelector('#uploadPreviewBody');
const progressPanel = document.querySelector('#uploadProgress');
const progressBar = document.querySelector('#progressBar');
const progressText = document.querySelector('#progressText');
const progressPercent = document.querySelector('#progressPercent');
const uploadResult = document.querySelector('#uploadResult');
const settingsPanel = document.querySelector('#supabaseSettings');
const settingsButton = document.querySelector('#toggleSupabaseSettings');
const saveSettingsButton = document.querySelector('#saveSupabaseSettings');
const urlInput = document.querySelector('#supabaseUrl');
const anonKeyInput = document.querySelector('#supabaseAnonKey');
const accessTokenInput = document.querySelector('#supabaseAccessToken');
const emailInput = document.querySelector('#supabaseEmail');
const passwordInput = document.querySelector('#supabasePassword');
const loginButton = document.querySelector('#supabaseLogin');

let uploadFile;
let analysis;
let automaticUploadRunning = false;

loadConnectionSettings();

tabUpload.addEventListener('click', () => {
  document.querySelector('#costWorkspace').hidden = true;
  document.querySelector('#lookupWorkspace').hidden = true;
  uploadWorkspace.hidden = false;
  tabUpload.classList.add('active');
  tabGenerate.classList.remove('active');
  tabLookup.classList.remove('active');
  tabDownload.classList.remove('active');
  uploadWorkspace.scrollIntoView({ behavior: 'smooth', block: 'center' });
});

settingsButton.addEventListener('click', () => {
  settingsPanel.hidden = !settingsPanel.hidden;
});

saveSettingsButton.addEventListener('click', () => {
  const settings = readConnectionInputs();
  if (!settings.url || !settings.anonKey) {
    showMessage('Supabase URL과 publishable key를 입력해 주세요.', 'error');
    return;
  }
  localStorage.setItem('costUploadSupabaseUrl', settings.url);
  localStorage.setItem('costUploadSupabaseAnonKey', settings.anonKey);
  sessionStorage.setItem('costUploadSupabaseAccessToken', settings.accessToken);
  settingsPanel.hidden = true;
  showMessage('Supabase 연결 설정을 저장했습니다.', 'success');
});

loginButton.addEventListener('click', loginToSupabase);

uploadDropzone.addEventListener('click', () => uploadInput.click());
uploadInput.addEventListener('change', () => selectUploadFile(uploadInput.files[0]));
uploadDropzone.addEventListener('dragover', event => {
  event.preventDefault();
  uploadDropzone.classList.add('dragging');
});
uploadDropzone.addEventListener('dragleave', () => uploadDropzone.classList.remove('dragging'));
uploadDropzone.addEventListener('drop', event => {
  event.preventDefault();
  uploadDropzone.classList.remove('dragging');
  selectUploadFile(event.dataTransfer.files[0]);
});

analyzeButton.addEventListener('click', analyzeUpload);
executeButton.addEventListener('click', executeUpload);

function loadConnectionSettings() {
  urlInput.value = localStorage.getItem('costUploadSupabaseUrl') || DEFAULT_SUPABASE_URL;
  anonKeyInput.value = localStorage.getItem('costUploadSupabaseAnonKey') || DEFAULT_SUPABASE_PUBLISHABLE_KEY;
  accessTokenInput.value = sessionStorage.getItem('costUploadSupabaseAccessToken') || '';
  if (!urlInput.value || !anonKeyInput.value) {
    settingsButton.hidden = false;
    settingsPanel.hidden = false;
  }
}

function readConnectionInputs() {
  return {
    url: urlInput.value.trim().replace(/\/$/, ''),
    anonKey: anonKeyInput.value.trim(),
    accessToken: accessTokenInput.value.trim(),
  };
}

function getConnectionSettings() {
  const settings = readConnectionInputs();
  if (!settings.url || !settings.anonKey) {
    settingsPanel.hidden = false;
    throw new Error('Supabase 연결 설정을 먼저 입력해 주세요.');
  }
  if (!/^https:\/\/[a-z0-9-]+\.supabase\.co$/i.test(settings.url)) {
    throw new Error('Supabase URL 형식을 확인해 주세요.');
  }
  return settings;
}

async function selectUploadFile(file) {
  if (!file) return;
  if (automaticUploadRunning) {
    showMessage('현재 파일을 업로드하고 있습니다. 완료 후 다시 선택해 주세요.');
    return;
  }
  if (!/\.(xls|xlsx|xlsm)$/i.test(file.name)) {
    showMessage('Excel 파일(.xls, .xlsx, .xlsm)만 선택할 수 있습니다.', 'error');
    return;
  }
  uploadFile = file;
  analysis = undefined;
  uploadFileName.textContent = file.name;
  analyzeButton.disabled = false;
  preview.hidden = true;
  progressPanel.hidden = true;
  uploadResult.hidden = true;
  showMessage('파일을 확인하고 있습니다.');
  automaticUploadRunning = true;
  uploadInput.disabled = true;
  try {
    await analyzeUpload();
    if (analysis && (analysis.updates.length || analysis.inserts.length)) {
      await executeUpload();
    }
  } finally {
    automaticUploadRunning = false;
    uploadInput.disabled = false;
  }
}

async function analyzeUpload() {
  if (!uploadFile) return;
  setBusy(analyzeButton, true, '분석 중…');
  preview.hidden = true;
  uploadResult.hidden = true;
  try {
    const settings = getConnectionSettings();
    const excel = await readFirstSheet(uploadFile);
    showMessage('업로드 자료를 분석하고 있습니다.');
    const databaseRows = await fetchAllCostRows(settings);
    analysis = classifyRows(excel, databaseRows);
    renderAnalysis(analysis);
    showMessage(`분석 완료 · 전체 ${analysis.items.length}건 · 수정 ${analysis.updates.length}건`, 'success');
  } catch (error) {
    showMessage(toFriendlyError(error), 'error');
  } finally {
    setBusy(analyzeButton, false, '업로드 분석');
  }
}

async function readFirstSheet(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true, raw: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error('첫 번째 시트를 찾을 수 없습니다.');
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
  if (!matrix.length) throw new Error('첫 번째 시트가 비어 있습니다.');

  const headers = matrix[0].map(value => String(value ?? '').trim());
  const itemIndex = headers.indexOf(ITEM_COLUMN);
  if (itemIndex < 0) throw new Error('첫 번째 행에 필수 컬럼 “품목”이 없습니다.');
  const duplicateHeaders = headers.filter((header, index) => header && headers.indexOf(header) !== index);
  if (duplicateHeaders.length) throw new Error(`중복된 컬럼명이 있습니다: ${[...new Set(duplicateHeaders)].join(', ')}`);

  const rows = [];
  matrix.slice(1).forEach((values, index) => {
    if (values.every(value => value === null || value === '')) return;
    const row = {};
    headers.forEach((header, column) => { if (header) row[header] = cleanCellValue(values[column]); });
    rows.push({ excelRow: index + 2, item: normalizeItem(values[itemIndex]), row });
  });
  return { headers: headers.filter(Boolean), rows };
}

function cleanCellValue(value) {
  if (value === '' || value === undefined) return null;
  if (value instanceof Date) return value.toISOString();
  return value;
}

function normalizeItem(value) {
  return normalizeItemKey(value);
}

function classifyRows(excel, databaseRows) {
  const databaseColumns = new Set(databaseRows.flatMap(row => Object.keys(row)));
  const uploadColumns = COST_UPLOAD_COLUMNS.filter(({ target, sources }) =>
    (!databaseColumns.size || databaseColumns.has(target)) && sources.some(source => excel.headers.includes(source))
  );

  const seenExcelItems = new Set();
  const databaseMap = new Map();
  databaseRows.forEach(row => {
    const item = normalizeItem(row[ITEM_COLUMN]);
    if (!databaseMap.has(item)) databaseMap.set(item, []);
    databaseMap.get(item).push(row);
  });

  const items = excel.rows.map(entry => {
    const payload = {};
    uploadColumns.forEach(({ target, sources, numeric }) => {
      const source = sources.find(header => Object.prototype.hasOwnProperty.call(entry.row, header) && hasUploadValue(entry.row[header]));
      if (source !== undefined) {
        const value = normalizeUploadValue(entry.row[source], numeric);
        if (hasUploadValue(value)) payload[target] = value;
      }
    });
    // Blank cells never clear existing database values, including currency prices.
    if (!entry.item) return makeItem('review', entry, payload, '필수값 “품목”이 없습니다.');
    if (seenExcelItems.has(entry.item)) {
      return makeItem('unchanged', entry, payload, 'Excel 중복 품목입니다. 첫 번째 행만 처리하고 건너뜁니다.');
    }
    seenExcelItems.add(entry.item);
    const matches = databaseMap.get(entry.item) || [];
    if (matches.length > 1) return makeItem('review', entry, payload, 'cost 테이블에 동일한 품목이 중복되어 있습니다.');
    if (!matches.length) {
      const insertPayload = { [ITEM_COLUMN]: entry.item, ...payload };
      if (!hasUploadValue(insertPayload['품목명'])) insertPayload['품목명'] = entry.item;
      return makeItem('new', entry, insertPayload, 'Supabase 신규 품목으로 등록', { conflictPayload: payload });
    }
    const matchedRow = matches[0];
    const changedColumns = Object.keys(payload).filter(header => !valuesEqual(payload[header], matchedRow[header]));
    if (!changedColumns.length) return makeItem('unchanged', entry, payload, '변경 없음');
    const changedPayload = Object.fromEntries(changedColumns.map(header => [header, payload[header]]));
    return makeItem(
      'update',
      entry,
      changedPayload,
      `${changedColumns.slice(0, 4).join(', ')}${changedColumns.length > 4 ? ` 외 ${changedColumns.length - 4}개` : ''} 변경`,
      { databaseId: matchedRow.id, databaseItem: matchedRow[ITEM_COLUMN] }
    );
  });

  return {
    items,
    writableHeaders: uploadColumns.map(column => column.target),
    updates: items.filter(item => item.type === 'update'),
    inserts: items.filter(item => item.type === 'new'),
    unchanged: items.filter(item => item.type === 'unchanged'),
    review: items.filter(item => item.type === 'review'),
  };
}

function hasUploadValue(value) {
  if (value === 0 || value === '0') return true;
  return value !== null && value !== undefined && (typeof value !== 'string' || value.trim() !== '');
}

function normalizeUploadValue(value, numeric) {
  if (!numeric) return value;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value !== 'string') return value;
  const normalized = value.replace(/,/g, '').trim();
  if (normalized === '') return value;
  const number = Number(normalized);
  return Number.isFinite(number) ? number : value;
}

function makeItem(type, entry, payload, detail, locator = {}) {
  return { type, excelRow: entry.excelRow, item: entry.item || '-', payload, detail, ...locator };
}

function valuesEqual(left, right) {
  if ((left === null || left === '') && (right === null || right === '')) return true;
  if (typeof left === 'number' && typeof right === 'number') return left === right;
  return String(left ?? '').trim() === String(right ?? '').trim();
}

function renderAnalysis(result) {
  document.querySelector('#newCount').textContent = result.inserts.length;
  document.querySelector('#updateCount').textContent = result.updates.length;
  document.querySelector('#unchangedCount').textContent = result.unchanged.length;
  document.querySelector('#reviewCount').textContent = result.review.length;
  const labels = { new: '신규', update: '수정', unchanged: '변경 없음', review: '확인 필요' };
  previewBody.replaceChildren(...result.items.slice(0, PREVIEW_LIMIT).map(item => {
    const row = document.createElement('tr');
    row.innerHTML = `<td class="type-${item.type}">${labels[item.type]}</td><td>${item.excelRow}</td><td></td><td></td>`;
    row.children[2].textContent = item.item;
    row.children[3].textContent = item.detail;
    return row;
  }));
  executeButton.disabled = result.updates.length + result.inserts.length === 0;
  preview.hidden = false;
}

async function executeUpload() {
  if (!analysis || (!analysis.updates.length && !analysis.inserts.length)) return;
  const settings = getConnectionSettings();
  const actions = [...analysis.updates, ...analysis.inserts];
  const completed = [];
  const failed = [];
  executeButton.disabled = true;
  progressPanel.hidden = false;
  uploadResult.hidden = true;
  updateProgress(0, actions.length);

  for (let index = 0; index < actions.length; index += 1) {
    const action = actions[index];
    try {
      if (action.type === 'update') await updateCostRow(settings, action);
      else await insertCostRow(settings, action);
      completed.push(action);
    } catch (error) {
      failed.push({ action, message: toFriendlyError(error) });
      if (error?.code === 'RLS_UPDATE_DENIED') {
        actions.slice(index + 1).forEach(pending => failed.push({
          action: pending,
          message: '앞선 품목에서 Supabase 수정 권한 오류가 확인되어 실행하지 않았습니다.',
        }));
        updateProgress(actions.length, actions.length);
        break;
      }
    }
    updateProgress(completed.length + failed.length, actions.length);
  }

  let verified = 0;
  let verificationFailed = [];
  try {
    const refreshedRows = await fetchAllCostRows(settings);
    ({ verified, failed: verificationFailed } = verifyChanges(completed, refreshedRows));
  } catch (error) {
    verificationFailed = completed.map(action => ({ action, message: toFriendlyError(error) }));
  }

  const hasError = failed.length > 0 || verificationFailed.length > 0;
  const updatedCount = completed.filter(action => action.type === 'update').length;
  const insertedCount = completed.filter(action => action.type === 'new').length;
  uploadResult.className = `upload-result${hasError ? ' error' : ''}`;
  uploadResult.textContent = hasError
    ? `처리 완료: ${completed.length}건, 처리 실패: ${failed.length}건, 검증 완료: ${verified}건, 검증 필요: ${verificationFailed.length}건. ${[...failed, ...verificationFailed].slice(0, 3).map(item => `${item.action.item}: ${item.message}`).join(' / ')}`
    : `업로드와 재조회 검증을 완료했습니다. 수정 ${updatedCount}건, 신규 ${insertedCount}건, 검증 완료 ${verified}건.`;
  uploadResult.hidden = false;
  showMessage(hasError ? '일부 항목을 확인해 주세요.' : '업로드가 완료되었습니다.', hasError ? 'error' : 'success');
  executeButton.disabled = false;
}

async function loginToSupabase() {
  const settings = getConnectionSettings();
  const email = emailInput.value.trim();
  const password = passwordInput.value;
  if (!email || !password) {
    showMessage('Supabase 로그인 이메일과 비밀번호를 입력해 주세요.', 'error');
    return;
  }
  setBusy(loginButton, true, '로그인 중…');
  try {
    const response = await fetch(`${settings.url}/auth/v1/token?grant_type=password`, {
      method: 'POST',
      headers: { apikey: settings.anonKey, 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password }),
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok || !result.access_token) {
      throw new Error(result.msg || result.message || '이메일 또는 비밀번호를 확인해 주세요.');
    }
    accessTokenInput.value = result.access_token;
    sessionStorage.setItem('costUploadSupabaseAccessToken', result.access_token);
    passwordInput.value = '';
    showMessage('Supabase 로그인에 성공했습니다. 이제 업로드를 실행할 수 있습니다.', 'success');
  } catch (error) {
    showMessage(`Supabase 로그인 실패: ${error.message}`, 'error');
  } finally {
    setBusy(loginButton, false, 'Supabase 로그인');
  }
}

async function fetchAllCostRows(settings) {
  const rows = [];
  for (let start = 0; ; start += PAGE_SIZE) {
    const page = await supabaseRequest(settings, `/rest/v1/${COST_TABLE}?select=*`, {
      headers: { Range: `${start}-${start + PAGE_SIZE - 1}` },
    });
    if (!Array.isArray(page)) throw new Error('cost 테이블 조회 결과를 확인할 수 없습니다.');
    rows.push(...page);
    if (page.length < PAGE_SIZE) break;
  }
  return rows;
}

async function updateCostRow(settings, action) {
  let result = await patchCostRow(settings, action, false);
  if (!Array.isArray(result)) throw new Error('Supabase 수정 결과를 확인할 수 없습니다.');
  if (result.length === 0) {
    const error = new Error('분석 단계에서 찾은 cost 행을 수정하지 못했습니다.');
    error.status = 403;
    error.code = 'RLS_UPDATE_DENIED';
    throw error;
  }
  if (result.length !== 1) throw new Error('cost 테이블에 동일한 품목이 여러 건 존재합니다. 품목 중복을 확인해 주세요.');

  let mismatches = getPayloadMismatches(action.payload, result[0]);
  if (mismatches.length && Object.values(action.payload).some(value => value === 0)) {
    result = await patchCostRow(settings, action, true);
    if (!Array.isArray(result) || result.length !== 1) {
      throw new Error(`${action.item}의 숫자 0 값을 다시 전송했지만 수정 결과를 확인하지 못했습니다.`);
    }
    mismatches = getPayloadMismatches(action.payload, result[0]);
  }
  if (mismatches.length) {
    const detail = mismatches.slice(0, 3).map(({ column, expected, actual }) =>
      `${column}: 전송 ${formatDiagnosticValue(expected)}, 저장 ${formatDiagnosticValue(actual)}`
    ).join(', ');
    throw new Error(`Supabase가 전송값과 다른 값을 반환했습니다. ${detail}`);
  }
}

async function patchCostRow(settings, action, forceItemMatch) {
  const url = new URL(`${settings.url}/rest/v1/${COST_TABLE}`);
  if (!forceItemMatch && action.databaseId !== null && action.databaseId !== undefined && action.databaseId !== '') {
    url.searchParams.set('id', `eq.${action.databaseId}`);
  } else {
    url.searchParams.set(ITEM_COLUMN, `eq.${action.databaseItem ?? action.item}`);
  }
  return supabaseRequest(settings, url.toString(), {
    method: 'PATCH',
    headers: { Prefer: 'return=representation,count=exact' },
    body: JSON.stringify(action.payload),
  });
}

function getPayloadMismatches(payload, row) {
  return Object.entries(payload)
    .filter(([column, value]) => !valuesEqual(value, row?.[column]))
    .map(([column, expected]) => ({ column, expected, actual: row?.[column] }));
}

function formatDiagnosticValue(value) {
  if (value === null || value === undefined || value === '') return 'null';
  return JSON.stringify(value);
}

async function insertCostRow(settings, action) {
  if (Object.keys(action.conflictPayload || {}).length) {
    const url = new URL(`${settings.url}/rest/v1/${COST_TABLE}`);
    url.searchParams.set(ITEM_COLUMN, `eq.${action.item}`);
    const updated = await supabaseRequest(settings, url.toString(), {
      method: 'PATCH',
      headers: { Prefer: 'return=representation,count=exact' },
      body: JSON.stringify(action.conflictPayload),
    });
    if (Array.isArray(updated) && updated.length === 1) {
      action.type = 'update';
      action.payload = action.conflictPayload;
      return;
    }
    if (Array.isArray(updated) && updated.length > 1) {
      throw new Error('cost 테이블에 동일한 품목이 여러 건 존재합니다. 품목 중복을 확인해 주세요.');
    }
  }
  const result = await supabaseRequest(settings, `/rest/v1/${COST_TABLE}`, {
    method: 'POST',
    headers: { Prefer: 'return=representation' },
    body: JSON.stringify(action.payload),
  });
  if (!Array.isArray(result) || result.length !== 1) {
    throw new Error('Supabase 신규 품목 등록 결과를 확인할 수 없습니다. INSERT 권한과 RLS 정책을 확인해 주세요.');
  }
}

async function supabaseRequest(settings, pathOrUrl, options = {}) {
  const url = pathOrUrl.startsWith('http') ? pathOrUrl : `${settings.url}${pathOrUrl}`;
  const authorization = settings.accessToken ? { Authorization: `Bearer ${settings.accessToken}` } : {};
  const response = await fetch(url, {
    ...options,
    headers: {
      apikey: settings.anonKey,
      ...authorization,
      'Content-Type': 'application/json',
      ...options.headers,
    },
  });
  const text = await response.text();
  let payload;
  try { payload = text ? JSON.parse(text) : null; } catch { payload = text; }
  if (!response.ok) {
    const error = new Error(payload?.message || payload?.hint || `Supabase 요청 실패 (${response.status})`);
    error.status = response.status;
    error.code = payload?.code;
    throw error;
  }
  return payload;
}

function verifyChanges(actions, rows) {
  const databaseMap = new Map();
  rows.forEach(row => {
    const item = normalizeItem(row[ITEM_COLUMN]);
    if (!databaseMap.has(item)) databaseMap.set(item, []);
    databaseMap.get(item).push(row);
  });
  let verified = 0;
  const failed = [];
  actions.forEach(action => {
    const matches = databaseMap.get(action.item) || [];
    const valid = matches.length === 1 && Object.entries(action.payload).every(([column, value]) => valuesEqual(value, matches[0][column]));
    if (valid) verified += 1;
    else failed.push({ action, message: '재조회한 값이 업로드 값과 일치하지 않습니다.' });
  });
  return { verified, failed };
}

function updateProgress(done, total) {
  const percent = total ? Math.round((done / total) * 100) : 0;
  progressText.textContent = `${done} / ${total}`;
  progressPercent.textContent = `${percent}%`;
  progressBar.value = percent;
}

function setBusy(target, busy, label) {
  target.disabled = busy;
  target.textContent = label;
}

function showMessage(message, type = '') {
  uploadMessage.className = `upload-message${type ? ` ${type}` : ''}`;
  uploadMessage.textContent = message;
}

function toFriendlyError(error) {
  const message = String(error?.message || error || '알 수 없는 오류');
  if (error?.code === 'RLS_UPDATE_DENIED') {
    return 'Supabase cost 테이블의 수정 권한이 없습니다. anon 역할의 UPDATE 권한과 RLS 정책을 확인해 주세요.';
  }
  if (error?.status === 401 || error?.status === 403 || /permission|policy|rls|row-level security/i.test(message)) {
    return 'Supabase 접근 권한이 없습니다. 로그인 상태와 cost 테이블의 RLS 정책을 확인해 주세요.';
  }
  if (error?.status === 409 || error?.code === '23505' || /duplicate|unique/i.test(message)) {
    return '동일한 품목이 이미 존재하여 등록할 수 없습니다. cost 테이블의 품목 고유 제약을 확인해 주세요.';
  }
  if (error?.code === '22P02' && /integer/i.test(message)) {
    const decimalValue = message.match(/integer:\s*"([^"]+)"/i)?.[1];
    const valueGuide = decimalValue ? ` 소수값 ${decimalValue}을(를)` : ' 소수값을';
    return `Supabase 단가 컬럼이 정수형이라${valueGuide} 저장할 수 없습니다. supabase-price-decimal-migration.sql을 Supabase SQL Editor에서 한 번 실행해 단가 컬럼을 numeric 형식으로 변경해 주세요.`;
  }
  if (/column|schema cache|could not find/i.test(message)) {
    const code = error?.code ? ` (${error.code})` : '';
    return `Supabase 컬럼 오류${code}: ${message}`;
  }
  if (/Failed to fetch|NetworkError/i.test(message)) {
    return 'Supabase에 연결할 수 없습니다. URL과 네트워크 상태를 확인해 주세요.';
  }
  return message;
}

window.costDatabase = {
  async getUnitPriceMap() {
    const settings = getConnectionSettings();
    const rows = await fetchUnitPriceRows(settings);
    if (!rows.length) {
      throw new Error('cost 테이블에서 조회된 데이터가 없습니다. 데이터가 존재한다면 SELECT RLS 정책 또는 사용자 access token을 확인해 주세요.');
    }
    const prices = new Map();
    const records = new Map();
    rows.forEach(row => {
      const item = normalizeItemKey(row[ITEM_COLUMN]);
      if (item && !records.has(item)) records.set(item, row);
      if (item && !prices.has(item) && row[PRICE_COLUMN] !== null && row[PRICE_COLUMN] !== '') prices.set(item, row[PRICE_COLUMN]);
    });
    return { prices, records, priceColumn: PRICE_COLUMN, rowCount: rows.length };
  },
  async downloadCostTable() {
    const settings = getConnectionSettings();
    const rows = await fetchAllCostRows(settings);
    if (!rows.length) throw new Error('cost 테이블에서 다운로드할 데이터가 없습니다.');

    const headers = ['품목', '품목명', '회사', '거래처명', '담당자', '단위', '달러단가', '원화단가'];
    const widths = [25.125, 57.125, 14.125, 20.125, 12.125, 11.75, 12, 12.25];
    const workbook = new ExcelJS.Workbook();
    const today = new Date();
    const dateStamp = [
      today.getFullYear(),
      String(today.getMonth() + 1).padStart(2, '0'),
      String(today.getDate()).padStart(2, '0'),
    ].join('');
    workbook.title = `단가표.${dateStamp}`;
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.columns = headers.map((header, index) => ({ header, key: header, width: widths[index] }));
    rows.forEach(row => sheet.addRow(headers.map(header => row[header] ?? null)));

    const headerRow = sheet.getRow(1);
    headerRow.height = 12.75;
    headerRow.eachCell(cell => {
      cell.font = { name: '굴림체', size: 9, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF969696' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = BORDER;
    });

    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      row.height = 12.75;
      row.eachCell({ includeEmpty: true }, cell => {
        cell.font = { name: '굴림', size: 9, color: { argb: 'FF000000' } };
        cell.alignment = { vertical: 'middle' };
        cell.border = BORDER;
      });
      [1, 2, 4].forEach(column => { row.getCell(column).alignment = { horizontal: 'left', vertical: 'middle' }; });
      [3, 5, 6, 7, 8].forEach(column => { row.getCell(column).alignment = { horizontal: 'center', vertical: 'middle' }; });
      row.getCell(7).numFmt = '#,##0.####';
      row.getCell(8).numFmt = '#,##0.####';
    }

    sheet.autoFilter = { from: 'A1', to: `H${sheet.rowCount}` };
    const data = await workbook.xlsx.writeBuffer();
    const fileName = `단가표.${dateStamp}.xlsx`;
    if (window.desktopFile?.saveExcel) {
      const result = await window.desktopFile.saveExcel(new Uint8Array(data), fileName);
      return { count: rows.length, canceled: Boolean(result?.canceled), filePath: result?.filePath };
    }
    if (window.showSaveFilePicker) {
      try {
        const fileHandle = await window.showSaveFilePicker({
          suggestedName: fileName,
          startIn: 'downloads',
          types: [{
            description: 'Excel 통합 문서',
            accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
          }],
          excludeAcceptAllOption: true,
        });
        const writable = await fileHandle.createWritable();
        await writable.write(data);
        await writable.close();
        return { count: rows.length, canceled: false, filePath: fileHandle.name };
      } catch (error) {
        if (error?.name === 'AbortError') return { count: rows.length, canceled: true };
        throw error;
      }
    }
    throw new Error('이 브라우저에서는 저장 위치 선택 창을 지원하지 않습니다. 설치된 Windows 앱에서 실행해 주세요.');
  },
};

async function fetchUnitPriceRows(settings) {
  const rows = [];
  const select = encodeURIComponent(`${ITEM_COLUMN},${PRICE_COLUMN},${DOLLAR_PRICE_COLUMN},${UNIT_COLUMN},${MANAGER_COLUMN},${VENDOR_COLUMN}`);
  for (let start = 0; ; start += PAGE_SIZE) {
    const page = await supabaseRequest(settings, `/rest/v1/${COST_TABLE}?select=${select}`, {
      headers: { Range: `${start}-${start + PAGE_SIZE - 1}` },
    });
    if (!Array.isArray(page)) throw new Error('cost 테이블의 품목과 원화단가를 조회할 수 없습니다.');
    rows.push(...page);
    if (page.length < PAGE_SIZE) break;
  }
  return rows;
}

function normalizeItemKey(value) {
  return String(value ?? '').normalize('NFKC').replace(/\u00a0/g, ' ').trim().toUpperCase();
}
