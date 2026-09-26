// Search and edit order_mobility independently of the cost lookup.
const ORDER_FIELDS = ['품목', '품목명', '프로젝트명', '세부프로젝트명', 'BOM'];
const ORDER_SEARCH_FIELDS = ['품목', '프로젝트명', '세부프로젝트명'];
const ORDER_PAGE_SIZE = 500;
const orderForm = document.querySelector('#orderLookupForm');
const orderKeyword = document.querySelector('#orderLookupKeyword');
const orderField = document.querySelector('#orderLookupField');
const orderMessage = document.querySelector('#orderLookupMessage');
const orderCandidates = document.querySelector('#orderCandidates');
const orderCandidateBody = document.querySelector('#orderCandidateBody');
const orderDetail = document.querySelector('#orderDetail');
const orderDetailFields = document.querySelector('#orderDetailFields');
const orderEdit = document.querySelector('#orderEdit');
const orderSave = document.querySelector('#orderSave');
const orderCancel = document.querySelector('#orderCancel');
const orderDelete = document.querySelector('#orderDelete');
let orderRows = [];
let selectedOrder = null;
let orderEditing = false;
let orderMutating = false;
let orderSearchController;
let orderSearchSequence = 0;
let orderSearchTimer;

orderForm.addEventListener('submit', event => { event.preventDefault(); clearTimeout(orderSearchTimer); searchOrders(); });
orderKeyword.addEventListener('input', scheduleOrderSearch);
orderField.addEventListener('change', scheduleOrderSearch);
orderEdit.addEventListener('click', () => { orderEditing = true; renderOrderDetail(); });
orderCancel.addEventListener('click', () => { orderEditing = false; renderOrderDetail(); });
orderSave.addEventListener('click', saveSelectedOrder);
orderDelete.addEventListener('click', deleteSelectedOrder);

function scheduleOrderSearch() {
  clearTimeout(orderSearchTimer);
  orderSearchController?.abort();
  orderSearchSequence += 1;
  // Clear stale selection immediately when the query changes.
  selectedOrder = null;
  orderRows = [];
  orderCandidates.hidden = true;
  orderDetail.hidden = true;
  orderSearchTimer = setTimeout(searchOrders, 300);
}

function orderLiteral(value) {
  return '"' + String(value).replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
}

function orderSearchUrl(settings, keyword, field, offset = 0) {
  const url = new URL(`${settings.url}/rest/v1/order_mobility`);
  const fields = ORDER_SEARCH_FIELDS.includes(field) ? [field] : ORDER_SEARCH_FIELDS;
  // PostgREST treats * as a wildcard alias. Escape SQL wildcard characters;
  // use * as an optional user wildcard, then filter literal matches below.
  const pattern = '%' + keyword.replace(/\\/g, '\\\\').replace(/[%_]/g, '\\$&') + '%';
  url.searchParams.set('select', ORDER_FIELDS.join(','));
  url.searchParams.set('or', '(' + fields.map(name => `${name}.ilike.${orderLiteral(pattern)}`).join(',') + ')');
  url.searchParams.set('order', ORDER_FIELDS.slice(0, 4).map(name => `${name}.asc.nullsfirst`).join(','));
  url.searchParams.set('offset', String(offset));
  url.searchParams.set('limit', String(ORDER_PAGE_SIZE));
  return url.toString();
}

function matchesOrderKeyword(row, keyword, field) {
  const fields = ORDER_SEARCH_FIELDS.includes(field) ? [field] : ORDER_SEARCH_FIELDS;
  const needle = keyword.toLocaleLowerCase();
  return fields.some(name => String(row[name] ?? '').toLocaleLowerCase().includes(needle));
}

async function fetchOrderMatches(settings, keyword, field, signal) {
  const rows = [];
  let offset = 0;
  while (true) {
    const page = await supabaseRequest(settings, orderSearchUrl(settings, keyword, field, offset), { signal });
    if (!Array.isArray(page)) throw new Error('조회 결과를 확인할 수 없습니다.');
    if (!page.length) break;
    rows.push(...page.filter(row => matchesOrderKeyword(row, keyword, field)));
    // Continue to an empty page so a server page cap never truncates matches.
    offset += page.length;
  }
  return rows;
}

async function searchOrders() {
  if (orderMutating) return;
  orderSearchController?.abort();
  const sequence = ++orderSearchSequence;
  const keyword = orderKeyword.value.trim();
  selectedOrder = null;
  orderEditing = false;
  orderRows = [];
  orderCandidates.hidden = true;
  orderDetail.hidden = true;
  if (!keyword) { showOrderMessage('품목, 프로젝트명 또는 세부프로젝트명의 일부를 입력해 주세요.'); return; }
  orderSearchController = new AbortController();
  showOrderMessage('일치하는 자료를 찾고 있습니다…');
  try {
    const rows = await fetchOrderMatches(getConnectionSettings(), keyword, orderField.value, orderSearchController.signal);
    if (sequence !== orderSearchSequence) return;
    orderRows = rows;
    renderOrderCandidates();
    showOrderMessage(rows.length ? `${rows.length.toLocaleString()}건을 찾았습니다. 자료를 선택하면 전체 내용을 볼 수 있습니다.`
      : '조회 가능한 일치 자료가 없습니다. 자료가 있다면 연결 계정의 조회 권한을 확인해 주세요.', rows.length ? 'success' : '');
  } catch (error) {
    if (sequence !== orderSearchSequence || error.name === 'AbortError') return;
    showOrderMessage(orderErrorMessage(error), 'error');
  }
}

function orderDisplay(value) {
  if (value === null || value === undefined || value === '') return '—';
  return typeof value === 'object' ? JSON.stringify(value, null, 2) : String(value);
}

function renderOrderCandidates() {
  orderCandidateBody.replaceChildren(...orderRows.map((row, index) => {
    const tr = document.createElement('tr');
    tr.classList.toggle('order-selected', row === selectedOrder);
    ORDER_FIELDS.slice(0, 4).forEach(field => {
      const td = document.createElement('td');
      td.textContent = orderDisplay(row[field]);
      tr.append(td);
    });
    const td = document.createElement('td');
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'order-select';
    button.textContent = row === selectedOrder ? '선택됨' : '선택';
    button.setAttribute('aria-label', `${index + 1}번 자료 ${row['품목'] ?? ''} 선택`);
    button.setAttribute('aria-pressed', String(row === selectedOrder));
    button.disabled = orderMutating;
    button.addEventListener('click', () => {
      selectedOrder = row;
      orderEditing = false;
      renderOrderCandidates();
      renderOrderDetail();
      orderDetail.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    });
    td.append(button); tr.append(td);
    return tr;
  }));
  orderCandidates.hidden = !orderRows.length;
}

function renderOrderDetail() {
  orderDetail.hidden = !selectedOrder;
  if (!selectedOrder) return;
  orderDetailFields.replaceChildren(...ORDER_FIELDS.map(field => {
    const label = document.createElement('label');
    const title = document.createElement('span'); title.textContent = field;
    label.append(title);
    if (orderEditing) {
      const input = document.createElement('textarea');
      input.rows = field === 'BOM' ? 5 : 2;
      input.dataset.orderField = field;
      input.value = selectedOrder[field] == null ? '' : orderDisplay(selectedOrder[field]);
      input.disabled = orderMutating;
      label.append(input);
    } else {
      const value = document.createElement('div');
      value.className = 'order-detail-value';
      value.textContent = orderDisplay(selectedOrder[field]);
      label.append(value);
    }
    return label;
  }));
  document.querySelector('#orderGenerateCost').hidden = orderEditing;
  orderEdit.hidden = orderEditing;
  orderSave.hidden = !orderEditing;
  orderCancel.hidden = !orderEditing;
  orderDelete.hidden = orderEditing;
}

function orderMutationUrl(settings, row) {
  const url = new URL(`${settings.url}/rest/v1/order_mobility`);
  // This table has no id. Match every original field, including nulls, to avoid
  // changing a different project's record or overwriting concurrent edits.
  ORDER_FIELDS.forEach(field => {
    if (!Object.prototype.hasOwnProperty.call(row, field)) throw new Error('자료를 다시 조회한 후 처리해 주세요.');
    const value = row[field];
    url.searchParams.set(field, value == null ? 'is.null' : `eq.${typeof value === 'object' ? JSON.stringify(value) : value}`);
  });
  return url;
}

async function mutateOrder(method, original, payload) {
  const settings = getConnectionSettings();
  const url = orderMutationUrl(settings, original);
  const checkUrl = new URL(url);
  checkUrl.searchParams.set('select', ORDER_FIELDS.join(','));
  checkUrl.searchParams.set('limit', '2');
  const matches = await supabaseRequest(settings, checkUrl.toString());
  if (!Array.isArray(matches) || matches.length !== 1) {
    throw new Error(matches?.length > 1 ? '완전히 동일한 자료가 여러 건 있어 한 건만 처리할 수 없습니다. DB에서 고유 번호를 지정해 주세요.'
      : '자료가 변경되었거나 조회 권한이 없습니다. 다시 조회해 주세요.');
  }
  const result = await supabaseRequest(settings, url.toString(), {
    method,
    // The server must enforce this atomically, including records inserted
    // between the preflight read and mutation. Unsupported servers fail closed.
    headers: { Prefer: 'handling=strict,max-affected=1,return=representation' },
    ...(method === 'PATCH' ? { body: JSON.stringify(payload) } : {}),
  });
  if (!Array.isArray(result) || result.length !== 1) throw new Error('처리된 자료가 없습니다. 수정·삭제 권한 또는 다른 사용자의 변경 여부를 확인해 주세요.');
  return result[0];
}

function orderEditPayload(row, inputs) {
  const payload = {};
  for (const input of inputs) {
    const field = input.dataset.orderField;
    const before = row[field] == null ? '' : orderDisplay(row[field]);
    if (input.value === before) continue;
    const text = input.value.trim();
    if (!text) payload[field] = null;
    else if (typeof row[field] === 'object' && row[field] !== null) {
      try { payload[field] = JSON.parse(text); } catch { throw new Error(`${field}의 JSON 형식을 확인해 주세요.`); }
    } else if (typeof row[field] === 'number') {
      const number = Number(text);
      if (!Number.isFinite(number)) throw new Error(`${field}에는 숫자를 입력해 주세요.`);
      payload[field] = number;
    } else if (typeof row[field] === 'boolean') {
      if (!['true', 'false'].includes(text)) throw new Error(`${field}에는 true 또는 false를 입력해 주세요.`);
      payload[field] = text === 'true';
    } else payload[field] = text;
  }
  return payload;
}

async function saveSelectedOrder() {
  if (!selectedOrder || orderMutating) return;
  const original = selectedOrder;
  let payload;
  try { payload = orderEditPayload(original, orderDetailFields.querySelectorAll('[data-order-field]')); }
  catch (error) { showOrderMessage(error.message, 'error'); return; }
  if (!Object.keys(payload).length) { orderEditing = false; renderOrderDetail(); showOrderMessage('변경한 내용이 없습니다.'); return; }
  setOrderMutating(true);
  try {
    const updated = await mutateOrder('PATCH', original, payload);
    Object.assign(original, updated);
    orderEditing = false;
    showOrderMessage('선택한 자료를 수정했습니다.', 'success');
    renderOrderDetail();
  } catch (error) { showOrderMessage(orderErrorMessage(error), 'error'); }
  finally { setOrderMutating(false); renderOrderCandidates(); }
}

async function deleteSelectedOrder() {
  if (!selectedOrder || orderMutating) return;
  const original = selectedOrder;
  if (!window.confirm(`선택한 자료 한 건을 삭제하시겠습니까?\n품목: ${orderDisplay(original['품목'])}\n프로젝트명: ${orderDisplay(original['프로젝트명'])}\n세부프로젝트명: ${orderDisplay(original['세부프로젝트명'])}\n삭제한 자료는 자동으로 복구되지 않습니다.`)) return;
  setOrderMutating(true);
  try {
    await mutateOrder('DELETE', original);
    orderRows = orderRows.filter(row => row !== original);
    selectedOrder = null;
    orderDetail.hidden = true;
    showOrderMessage('선택한 자료 한 건을 삭제했습니다.', 'success');
  } catch (error) { showOrderMessage(orderErrorMessage(error), 'error'); }
  finally { setOrderMutating(false); renderOrderCandidates(); }
}

function setOrderMutating(busy) {
  orderMutating = busy;
  document.querySelectorAll('#orderLookupSection button, #orderLookupSection input, #orderLookupSection select, #orderLookupSection textarea')
    .forEach(element => { element.disabled = busy; });
}

function showOrderMessage(message, type = '') {
  orderMessage.textContent = message;
  orderMessage.className = `lookup-message${type ? ` ${type}` : ''}`;
}

function orderErrorMessage(error) {
  if (error.code === 'PGRST122' || error.code === 'PGRST124') return '한 건만 처리하도록 보호하는 과정에서 요청이 중단되었습니다. DB 관리자에게 고유 번호와 서버 지원 설정을 확인해 주세요.';
  if (error.status === 401 || error.status === 403 || /permission|policy|rls/i.test(error.message)) return '이 자료에 대한 접근 권한이 없습니다. 로그인 상태와 조회·수정·삭제 권한을 확인해 주세요.';
  return String(error.message || '요청을 처리하지 못했습니다.');
}

document.querySelector('#orderGenerateCost').addEventListener('click',async()=>{
  if(!selectedOrder || orderMutating || orderEditing)return;
  const item=selectedOrder['품목'];
  orderMutating=true;
  const controls=[...document.querySelectorAll('#orderLookupSection button, #orderLookupSection input, #orderLookupSection select')];
  controls.forEach(control=>control.disabled=true);
  try {
    orderMessage.textContent='보관함에서 '+item+' 품목의 엑셀 파일을 찾는 중입니다…';
    const file=await ExcelArchive.readForItem(item);
    orderMessage.textContent='원가 자료와 환율을 적용하여 제품 원가를 생성하는 중입니다…';
    const result=await buildProductCost(file);
    const saved=await window.desktopFile.saveExcel(new Uint8Array(result.data),result.name);
    orderMessage.textContent=saved.canceled ? '원가 파일 저장을 취소했습니다.' : '제품 원가 저장 완료: '+result.name;
  } catch(error) {
    orderMessage.textContent='원가 생성 실패: '+error.message;
  } finally {
    orderMutating=false;
    controls.forEach(control=>control.disabled=false);
  }
});
