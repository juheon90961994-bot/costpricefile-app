const lookupWorkspace = document.querySelector('#lookupWorkspace');
const lookupForm = document.querySelector('#costLookupForm');
const lookupInput = document.querySelector('#costLookupItem');
const lookupButton = document.querySelector('#costLookupButton');
const lookupMessage = document.querySelector('#costLookupMessage');
const lookupResults = document.querySelector('#costLookupResults');
const lookupBody = document.querySelector('#costLookupBody');
const LOOKUP_FIELDS = ['품목', '품목명', '거래처명', '담당자', '단위', '달러단가', '원화단가'];
const LOOKUP_LIMIT = 100;

let lookupRows = [];
let lookupSequence=0;
let lookupMutating=false;
let editingLookupKey = null;

tabLookup.addEventListener('click', () => {
  document.querySelector('#costWorkspace').hidden = true;
  document.querySelector('#uploadWorkspace').hidden = true;
  lookupWorkspace.hidden = false;
  tabGenerate.classList.remove('active');
  tabUpload.classList.remove('active');
  tabDownload.classList.remove('active');
  tabLookup.classList.add('active');
  lookupWorkspace.scrollIntoView({ behavior: 'smooth', block: 'center' });
  lookupInput.focus({ preventScroll: true });
});

lookupForm.addEventListener('submit', event => {
  event.preventDefault();
  searchCosts();
});

async function searchCosts() {
  const sequence=++lookupSequence;
  const keyword = normalizeItemKey(lookupInput.value);
  if (!keyword) {
    showLookupMessage('조회할 품목을 입력해 주세요.', 'error');
    lookupResults.hidden = true;
    return;
  }
  setLookupBusy(true);
  showLookupMessage('Supabase cost 테이블을 조회하고 있습니다.');
  try {
    const rows = await fetchAllCostRows(getConnectionSettings());
    if(sequence!==lookupSequence)return;
    lookupRows = rows
      .filter(row => normalizeItemKey(row[ITEM_COLUMN]).includes(keyword))
      .slice(0, LOOKUP_LIMIT);
    editingLookupKey = null;
    renderLookupRows();
    const limited = lookupRows.length === LOOKUP_LIMIT ? ' 최대 100건까지 표시합니다.' : '';
    showLookupMessage(`${lookupRows.length.toLocaleString()}건을 조회했습니다.${limited}`, 'success');
  } catch (error) {
    if(sequence!==lookupSequence)return;
    lookupRows = [];
    lookupResults.hidden = true;
    showLookupMessage(toFriendlyError(error), 'error');
  } finally {
    if(sequence===lookupSequence)setLookupBusy(false);
  }
}

function renderLookupRows() {
  lookupBody.replaceChildren(...lookupRows.map((row, index) => buildLookupRow(row, index)));
  lookupResults.hidden = lookupRows.length === 0;
}

function buildLookupRow(row, index) {
  const tr = document.createElement('tr');
  const key = lookupRowKey(row, index);
  if (editingLookupKey === key) {
    LOOKUP_FIELDS.forEach(field => {
      const td = document.createElement('td');
      const input = document.createElement('input');
      input.dataset.field = field;
      input.value = row[field] ?? '';
      input.type = field.includes('단가') ? 'text' : 'text';
      input.autocomplete = 'off';
      td.append(input);
      tr.append(td);
    });
    tr.append(buildActionCell([
      ['저장', 'lookup-save', () => saveLookupRow(row, tr)],
      ['취소', 'lookup-cancel', () => { editingLookupKey = null; renderLookupRows(); }],
    ]));
    return tr;
  }

  LOOKUP_FIELDS.forEach(field => {
    const td = document.createElement('td');
    td.textContent = formatLookupValue(row[field], field);
    tr.append(td);
  });
  tr.append(buildActionCell([
    ['수정', 'lookup-edit', () => { editingLookupKey = key; renderLookupRows(); }],
    ['삭제', 'lookup-delete', () => deleteLookupRow(row)],
  ]));
  return tr;
}

function buildActionCell(actions) {
  const td = document.createElement('td');
  const wrap = document.createElement('div');
  wrap.className = 'lookup-actions';
  actions.forEach(([label, className, handler]) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = className;
    button.textContent = label;
    button.addEventListener('click', handler);
    wrap.append(button);
  });
  td.append(wrap);
  return td;
}

async function saveLookupRow(original, tr) {
  const inputs = [...tr.querySelectorAll('input[data-field]')];
  const values = Object.fromEntries(inputs.map(input => [input.dataset.field, input.value.trim()]));
  if (!values['품목']) return showLookupMessage('품목은 비워둘 수 없습니다.', 'error');
  if (!values['품목명']) return showLookupMessage('품목명은 비워둘 수 없습니다.', 'error');

  let dollarPrice;
  let wonPrice;
  try {
    dollarPrice = parseLookupPrice(values['달러단가'], '달러단가');
    wonPrice = parseLookupPrice(values['원화단가'], '원화단가');
  } catch (error) {
    showLookupMessage(error.message, 'error');
    return;
  }
  if (/^(KRW|WON|원화|원)$/i.test(values['단위'])) dollarPrice = null;
  const payload = {
    '품목': values['품목'],
    '품목명': values['품목명'],
    '거래처명': values['거래처명'] || null,
    '담당자': values['담당자'] || null,
    '단위': values['단위'] || null,
    '달러단가': dollarPrice,
    '원화단가': wonPrice,
  };

  lookupMutating=true;
  setLookupBusy(true);
  showLookupMessage(`${original[ITEM_COLUMN]} 품목을 수정하고 있습니다.`);
  try {
    const result = await requestLookupMutation('PATCH', original, payload);
    if (!Array.isArray(result) || result.length !== 1) throw makeLookupPermissionError('수정');
    Object.assign(original, result[0]);
    editingLookupKey = null;
    renderLookupRows();
    showLookupMessage(`${original[ITEM_COLUMN]} 품목을 수정했습니다.`, 'success');
  } catch (error) {
    showLookupMessage(toFriendlyError(error), 'error');
  } finally {
    lookupMutating=false;
    setLookupBusy(false);
  }
}

async function deleteLookupRow(row) {
  const item = row[ITEM_COLUMN];
  if ((row.id === null || row.id === undefined || row.id === '') && lookupRows.filter(candidate => candidate[ITEM_COLUMN] === item).length !== 1) {
    showLookupMessage('동일한 품목이 여러 건 있어 안전하게 삭제할 수 없습니다. cost 테이블의 id 컬럼과 중복 데이터를 확인해 주세요.', 'error');
    return;
  }
  if (!window.confirm(`품목 ${item}을(를) 삭제하시겠습니까?\n삭제한 데이터는 자동으로 복구되지 않습니다.`)) return;
  lookupMutating=true;
  setLookupBusy(true);
  showLookupMessage(`${item} 품목을 삭제하고 있습니다.`);
  try {
    const result = await requestLookupMutation('DELETE', row);
    if (!Array.isArray(result) || result.length !== 1) throw makeLookupPermissionError('삭제');
    lookupRows = lookupRows.filter(candidate => candidate !== row);
    editingLookupKey = null;
    renderLookupRows();
    showLookupMessage(`${item} 품목을 삭제했습니다.`, 'success');
  } catch (error) {
    showLookupMessage(toFriendlyError(error), 'error');
  } finally {
    lookupMutating=false;
    setLookupBusy(false);
  }
}

async function requestLookupMutation(method, row, payload) {
  const settings = getConnectionSettings();
  const url = new URL(`${settings.url}/rest/v1/${COST_TABLE}`);
  if (row.id !== null && row.id !== undefined && row.id !== '') url.searchParams.set('id', `eq.${row.id}`);
  else url.searchParams.set(ITEM_COLUMN, `eq.${row[ITEM_COLUMN]}`);
  return supabaseRequest(settings, url.toString(), {
    method,
    headers: { Prefer: 'return=representation,count=exact' },
    ...(payload ? { body: JSON.stringify(payload) } : {}),
  });
}

function parseLookupPrice(value, label) {
  if (value === '') return null;
  const number = Number(value.replace(/,/g, ''));
  if (!Number.isFinite(number)) throw new Error(`${label}에는 숫자만 입력해 주세요.`);
  return number;
}

function formatLookupValue(value, field) {
  if (value === null || value === undefined || value === '') return '-';
  if (field.includes('단가') && Number.isFinite(Number(value))) return Number(value).toLocaleString('ko-KR', { maximumFractionDigits: 6 });
  return String(value);
}

function lookupRowKey(row, index) {
  return row.id !== null && row.id !== undefined ? `id:${row.id}` : `${row[ITEM_COLUMN]}:${index}`;
}

function makeLookupPermissionError(action) {
  const error = new Error(`Supabase cost 테이블의 ${action} 권한이 없습니다. RLS 정책을 확인해 주세요.`);
  error.status = 403;
  return error;
}

function setLookupBusy(busy) {
  lookupButton.disabled = busy;
  lookupInput.disabled = busy;
}

function showLookupMessage(message, type = '') {
  lookupMessage.className = `lookup-message${type ? ` ${type}` : ''}`;
  lookupMessage.textContent = message;
}
