const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const XLSX = require('../xlsx.full.min.js');

function load() {
  const element = () => ({ value: '', addEventListener() {} });
  const context = vm.createContext({
    window: {}, document: { querySelector: element },
    localStorage: { getItem() {} }, sessionStorage: { getItem() {} },
    tabUpload: element(), XLSX, URL,
  });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../cost-upload.js'), 'utf8'), context);
  return context;
}
const plain = value => JSON.parse(JSON.stringify(value));
function classify(row, stored) {
  return load().classifyRows({ headers: Object.keys(row), rows: [
    { excelRow: 2, item: 'PART', row },
  ] }, [{ id: 7, 품목: 'PART', ...stored }]);
}

test('null, missing, empty and whitespace preserve stored fields; zero is written', () => {
  const result = classify({ 품목: 'PART', 품목명: null, 회사: undefined,
    담당자: '', 거래처명: '  ', 단위: 'KRW', 원화단가: 0 },
  { 품목명: 'original', 회사: 'company', 담당자: 'owner', 거래처명: 'vendor',
    단위: 'USD', 달러단가: 123, 원화단가: 100 });
  assert.deepEqual(plain(result.updates[0].payload), { 단위: 'KRW', 원화단가: 0 });
});

test('explicit dollar value survives KRW and vendor alias is recognized', () => {
  const result = classify({ 품목: 'PART', 주거래처: 'new vendor', 단위: 'KRW', 달러단가: '0', 원화단가: '1,234.5' },
    { 거래처명: 'old vendor', 단위: 'KRW', 달러단가: 10, 원화단가: 20 });
  assert.deepEqual(plain(result.updates[0].payload), { 거래처명: 'new vendor', 달러단가: 0, 원화단가: 1234.5 });
});

test('all blank fields cause no update', () => {
  const result = classify({ 품목: 'PART', 단위: null, 달러단가: null, 원화단가: null },
    { 단위: 'KRW', 달러단가: 10, 원화단가: 20 });
  assert.equal(result.updates.length, 0);
  assert.equal(result.unchanged.length, 1);
});

test('real Excel parsing matches case-insensitively and PATCH preserves omitted fields', async () => {
  const context = load();
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ['품목', '담당자', '주거래처', '단위', '원화단가'],
    [' wbk0-00a1a ', '김주헌', '예일산업', 'KRW', 19740],
  ]), 'Sheet1');
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const excel = await context.readFirstSheet({ arrayBuffer: async () => buffer });
  const stored = { id: 1, 품목: 'WBK0-00A1A', 담당자: 'old', 거래처명: 'old', 단위: 'USD', 원화단가: 1, 달러단가: 99 };
  const result = context.classifyRows(excel, [stored]);
  assert.equal(result.inserts.length, 0);
  assert.equal(result.updates.length, 1);
  let requests = 0;
  context.fetch = async (url, options) => {
    requests++;
    assert.equal(options.method, 'PATCH');
    assert.equal(new URL(url).searchParams.get('id'), 'eq.1');
    const payload = JSON.parse(options.body);
    assert.deepEqual(payload, { 담당자: '김주헌', 거래처명: '예일산업', 단위: 'KRW', 원화단가: 19740 });
    Object.assign(stored, payload);
    return { ok: true, text: async () => JSON.stringify([stored]) };
  };
  await context.updateCostRow({ url: 'https://example.supabase.co', anonKey: 'test' }, result.updates[0]);
  assert.equal(requests, 1);
  assert.equal(stored.달러단가, 99);
  assert.equal(context.verifyChanges(result.updates, [stored]).verified, 1);
});

test('normalized duplicate database keys require review instead of a write', () => {
  const context = load();
  const result = context.classifyRows({ headers: ['품목', '원화단가'], rows: [
    { excelRow: 2, item: 'PART', row: { 품목: 'PART', 원화단가: 5 } },
  ] }, [{ 품목: 'PART', 원화단가: 1 }, { 품목: 'part', 원화단가: 2 }]);
  assert.equal(result.review.length, 1);
  assert.equal(result.updates.length, 0);
});
