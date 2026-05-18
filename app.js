/* global XLSX, ExcelJS, saveAs */

const SHEETS = {
  car: '자가용 이용 업무 수행 경비 사용내역서',
  transit: '대중교통 이용 업무 수행 경비 사용내역서',
  db: 'DB'
};

const ROWS = {
  carStart: 27,
  carEnd: 67,
  transitStart: 21,
  transitEnd: 39
};

const state = {
  parsed: null,
  templateArrayBuffer: null,
  routeResults: [],
  logs: []
};

const $ = (id) => document.getElementById(id);

function log(message) {
  const time = new Date().toLocaleTimeString('ko-KR', { hour12: false });
  state.logs.push(`[${time}] ${message}`);
  $('status').textContent = state.logs.join('\n');
}

function resetLog() {
  state.logs = [];
  $('status').textContent = '';
}

function toArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function readWorkbookRows(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', raw: false, cellDates: false });
  const firstSheetName = wb.SheetNames[0];
  const ws = wb.Sheets[firstSheetName];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
}

function readSheetRows(arrayBuffer, sheetName) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', raw: false, cellDates: false });
  const target = sheetName && wb.Sheets[sheetName] ? sheetName : wb.SheetNames[0];
  return XLSX.utils.sheet_to_json(wb.Sheets[target], { header: 1, defval: '', raw: false });
}

function normalizeText(v) {
  return String(v ?? '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeCompact(v) {
  return normalizeText(v).replace(/\s+/g, '');
}

function parseMoney(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = Number(String(v).replace(/,/g, '').replace(/원/g, '').trim());
  return Number.isFinite(n) ? n : 0;
}

function parseNumber(v) {
  if (v === null || v === undefined || v === '') return null;
  const s = String(v).replace(/,/g, '').replace(/km/gi, '').replace(/원/g, '').trim();
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function parseDateKey(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date && !isNaN(v)) return formatDate(v);

  const s = String(v).trim();
  if (!s) return '';

  // Excel serial date fallback
  if (/^\d{5}$/.test(s)) {
    const serial = Number(s);
    const d = XLSX.SSF.parse_date_code(serial);
    if (d) return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }

  const digits = s.replace(/[^0-9]/g, '');
  if (digits.length >= 8) {
    const y = digits.slice(0, 4);
    const m = digits.slice(4, 6);
    const d = digits.slice(6, 8);
    return `${y}-${m}-${d}`;
  }
  return s;
}

function formatDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function dateKeyToDate(dateKey) {
  const [y, m, d] = dateKey.split('-').map(Number);
  return new Date(y, m - 1, d);
}

function normalizePlaceName(name) {
  const raw = normalizeText(name);
  if (!raw) return '';
  const compact = normalizeCompact(raw);

  for (const place of window.PLACE_DB || []) {
    const names = [place.name, ...(place.aliases || [])];
    if (names.some(n => normalizeCompact(n) === compact)) return place.name;
  }

  if (compact.includes('대구혁신')) return '코스트코 혁신점';
  if (compact.includes('혁신점')) return '코스트코 혁신점';
  if (compact.includes('대구점') && compact.includes('코스트코')) return '코스트코 대구점';
  if (compact === '대구점') return '코스트코 대구점';
  if (compact.includes('대전점')) return '코스트코 대전점';
  if (compact.includes('세종점')) return '코스트코 세종점';
  if (compact.includes('대구사무실')) return '대구 사무실';

  return raw;
}

function getPlace(name) {
  const normalized = normalizePlaceName(name);
  return (window.PLACE_DB || []).find(p => p.name === normalized) || { name: normalized, address: '', shortAddress: '', lon: null, lat: null };
}

function headerIndexMap(rows) {
  let headerRowIndex = rows.findIndex(row => row.some(cell => String(cell).includes('승인일자')));
  if (headerRowIndex < 0) headerRowIndex = 0;
  const map = {};
  rows[headerRowIndex].forEach((h, i) => {
    const key = normalizeCompact(h);
    if (key) map[key] = i;
  });
  return { headerRowIndex, map };
}

function findColumn(map, candidates) {
  const compactCandidates = candidates.map(normalizeCompact);
  for (const c of compactCandidates) if (Object.prototype.hasOwnProperty.call(map, c)) return map[c];
  for (const [header, index] of Object.entries(map)) {
    if (compactCandidates.some(c => header.includes(c))) return index;
  }
  return -1;
}

function parseCardRows(rows) {
  const { headerRowIndex, map } = headerIndexMap(rows);
  const idx = {
    date: findColumn(map, ['승인일자', '이용일자']),
    time: findColumn(map, ['승인시간', '이용시간']),
    merchant: findColumn(map, ['가맹점명', '사용처']),
    amount: findColumn(map, ['승인금액(원화)', '승인금액', '이용금액']),
    cancel: findColumn(map, ['취소여부']),
    vat: findColumn(map, ['부가세', 'VAT']),
    category: findColumn(map, ['가맹점업종', '업종']),
    address: findColumn(map, ['가맹점주소', '주소'])
  };

  const required = ['date', 'merchant', 'amount'];
  const missing = required.filter(k => idx[k] < 0);
  if (missing.length) throw new Error(`카드내역 필수 열을 찾지 못했습니다: ${missing.join(', ')}`);

  const result = [];
  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const row = rows[r];
    const dateKey = parseDateKey(row[idx.date]);
    const merchant = normalizeText(row[idx.merchant]);
    const amount = parseMoney(row[idx.amount]);
    const cancel = idx.cancel >= 0 ? normalizeText(row[idx.cancel]).toUpperCase() : 'N';
    if (!dateKey || !merchant || amount === 0) continue;
    if (cancel === 'Y' || amount < 0) continue;

    const item = {
      dateKey,
      time: idx.time >= 0 ? normalizeText(row[idx.time]) : '',
      merchant,
      amount,
      vat: idx.vat >= 0 ? parseMoney(row[idx.vat]) : 0,
      category: idx.category >= 0 ? normalizeText(row[idx.category]) : '',
      address: idx.address >= 0 ? normalizeText(row[idx.address]) : ''
    };
    item.type = classifyCard(item);
    result.push(item);
  }
  return result;
}

function classifyCard(item) {
  const text = `${item.merchant} ${item.category} ${item.address}`.toLowerCase();
  if (/sr|srt|ktx|코레일|한국철도|레츠코레일|철도/.test(text)) return 'train';
  if (/주유|충전|가스충전|lpg|gs칼텍스|sk에너지|s-oil|에스오일|현대오일|오일뱅크|알뜰주유|ev|전기차|차지비|환경부충전/.test(text)) return 'fuel';
  if (/하이패스|도로공사|통행료|톨게이트|고속도로/.test(text)) return 'toll';
  return 'other';
}

function parseVisits(rows) {
  const header = rows[0] || [];
  const map = {};
  header.forEach((h, i) => { map[normalizeCompact(h)] = i; });
  const dateIdx = findColumn(map, ['방문날짜']);
  const timeIdx = findColumn(map, ['방문시간']);
  const storeIdx = findColumn(map, ['매장명']);
  if (dateIdx < 0 || storeIdx < 0) throw new Error('방문활동관리에서 방문날짜/매장명 열을 찾지 못했습니다.');

  return rows.slice(1)
    .map(row => ({
      dateKey: parseDateKey(row[dateIdx]),
      time: timeIdx >= 0 ? normalizeText(row[timeIdx]) : '',
      store: normalizePlaceName(row[storeIdx])
    }))
    .filter(v => v.dateKey && v.store)
    .sort((a, b) => (a.dateKey + a.time).localeCompare(b.dateKey + b.time));
}

function parseAttendance(rows) {
  const header = rows[0] || [];
  const map = {};
  header.forEach((h, i) => { map[normalizeCompact(h)] = i; });
  const dateIdx = findColumn(map, ['근무일자']);
  const startIdx = findColumn(map, ['출근지점']);
  const endIdx = findColumn(map, ['퇴근지점']);
  if (dateIdx < 0 || startIdx < 0) throw new Error('근태관리에서 근무일자/출근지점 열을 찾지 못했습니다.');

  return rows.slice(1)
    .map(row => ({
      dateKey: parseDateKey(row[dateIdx]),
      start: normalizePlaceName(row[startIdx]),
      end: endIdx >= 0 ? normalizePlaceName(row[endIdx]) : ''
    }))
    .filter(v => v.dateKey && v.start);
}

function parseTemplatePairDb(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', raw: false, cellDates: false });
  const ws = wb.Sheets[SHEETS.db];
  if (!ws) return new Map();
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
  const pairMap = new Map();
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const start = normalizePlaceName(row[6]);
    const end = normalizePlaceName(row[7]);
    const distanceKm = parseNumber(row[8]);
    const toll = parseNumber(row[9]) || 0;
    if (!start || !end || distanceKm === null) continue;
    pairMap.set(`${start}|${end}`, { distanceKm, toll, source: '양식DB' });
    if (!pairMap.has(`${end}|${start}`)) pairMap.set(`${end}|${start}`, { distanceKm, toll, source: '양식DB(역방향)' });
  }
  return pairMap;
}

function buildSegments(attendanceRows, visitRows, pairDb) {
  const visitsByDate = groupBy(visitRows, 'dateKey');
  const segments = [];

  for (const att of attendanceRows) {
    const stops = [];
    if (att.start) stops.push(att.start);
    for (const v of visitsByDate.get(att.dateKey) || []) stops.push(v.store);
    if (att.end) stops.push(att.end);

    const compactStops = [];
    for (const stop of stops) {
      const normalized = normalizePlaceName(stop);
      if (!normalized) continue;
      if (compactStops[compactStops.length - 1] !== normalized) compactStops.push(normalized);
    }

    for (let i = 0; i < compactStops.length - 1; i++) {
      const start = compactStops[i];
      const end = compactStops[i + 1];
      if (!start || !end || start === end) continue;
      const pair = pairDb.get(`${start}|${end}`) || null;
      segments.push({
        id: `${att.dateKey}-${segments.length + 1}`,
        dateKey: att.dateKey,
        start,
        end,
        distanceKm: pair?.distanceKm ?? null,
        toll: pair?.toll ?? 0,
        timeMin: null,
        source: pair?.source ?? '좌표대기'
      });
    }
  }
  return segments;
}

function groupBy(items, key) {
  const map = new Map();
  for (const item of items) {
    const k = item[key];
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(item);
  }
  return map;
}

async function enrichRoutesWithTmap(segments) {
  const payloadRoutes = [];
  for (const seg of segments) {
    const sp = getPlace(seg.start);
    const ep = getPlace(seg.end);
    if (sp.lon == null || sp.lat == null || ep.lon == null || ep.lat == null) continue;
    payloadRoutes.push({
      id: seg.id,
      startName: seg.start,
      endName: seg.end,
      start: { lon: sp.lon, lat: sp.lat },
      end: { lon: ep.lon, lat: ep.lat }
    });
  }

  if (!payloadRoutes.length) return { ok: false, message: '좌표가 있는 이동구간이 없습니다.' };

  try {
    const res = await fetch('/api/tmap-route', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ routes: payloadRoutes })
    });
    const data = await res.json();
    if (!res.ok || !data.ok) throw new Error(data.message || 'TMAP 호출 실패');

    const byId = new Map((data.results || []).map(r => [r.id, r]));
    for (const seg of segments) {
      const found = byId.get(seg.id);
      if (!found || !found.ok) continue;
      seg.distanceKm = found.distanceKm;
      seg.toll = found.toll ?? 0;
      seg.timeMin = found.timeMin ?? null;
      seg.source = 'TMAP';
    }
    return { ok: true, message: `TMAP 조회 완료: ${data.results.filter(r => r.ok).length}건` };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function prepareCardSummary(cardItems, attendanceRows, visitRows) {
  const fuelByDate = new Map();
  const tollByDate = new Map();
  const trains = [];
  const others = [];

  for (const item of cardItems) {
    if (item.type === 'fuel') addAmount(fuelByDate, item.dateKey, item.amount);
    else if (item.type === 'toll') addAmount(tollByDate, item.dateKey, item.amount);
    else if (item.type === 'train') trains.push({ ...item, routeText: inferTransitRoute(item.dateKey, attendanceRows, visitRows) });
    else others.push(item);
  }
  return { fuelByDate, tollByDate, trains, others };
}

function addAmount(map, key, amount) {
  map.set(key, (map.get(key) || 0) + amount);
}

function inferTransitRoute(dateKey, attendanceRows, visitRows) {
  const att = attendanceRows.find(a => a.dateKey === dateKey);
  const visits = visitRows.filter(v => v.dateKey === dateKey);
  if (att?.start && att?.end && att.start !== att.end) return `${att.start} ~ ${att.end}`;
  if (visits.length >= 2) return `${visits[0].store} ~ ${visits[visits.length - 1].store}`;
  if (visits.length === 1) return `${visits[0].store} 방문`;
  return '확인 필요';
}

async function buildPreview() {
  resetLog();
  $('preview').innerHTML = '';
  $('summary').style.display = 'none';
  $('downloadBtn').disabled = true;

  const templateFile = $('templateFile').files[0];
  const visitFile = $('visitFile').files[0];
  const attendanceFile = $('attendanceFile').files[0];
  const cardFile = $('cardFile').files[0];
  if (!templateFile || !visitFile || !attendanceFile || !cardFile) {
    alert('이동 양식, 방문활동관리, 근태관리, 카드승인내역 파일을 모두 넣어주세요.');
    return;
  }

  log('엑셀 파일을 읽는 중입니다.');
  state.templateArrayBuffer = await toArrayBuffer(templateFile);
  const visitAb = await toArrayBuffer(visitFile);
  const attendanceAb = await toArrayBuffer(attendanceFile);
  const cardAb = await toArrayBuffer(cardFile);

  const visits = parseVisits(readWorkbookRows(visitAb));
  const attendance = parseAttendance(readWorkbookRows(attendanceAb));
  const cardItems = parseCardRows(readWorkbookRows(cardAb));
  const pairDb = parseTemplatePairDb(state.templateArrayBuffer);

  log(`방문활동 ${visits.length}건, 근태 ${attendance.length}건, 카드 승인 ${cardItems.length}건을 읽었습니다.`);

  const segments = buildSegments(attendance, visits, pairDb);
  log(`이동구간 후보 ${segments.length}건을 생성했습니다.`);

  const tmapResult = await enrichRoutesWithTmap(segments);
  const badge = $('apiBadge');
  if (tmapResult.ok) {
    badge.className = 'badge green';
    badge.textContent = 'TMAP 연결 완료';
    log(tmapResult.message);
  } else {
    badge.className = 'badge amber';
    badge.textContent = 'TMAP 미사용/실패, 양식DB 보조';
    log(`TMAP 조회는 건너뛰거나 실패했습니다: ${tmapResult.message}`);
  }

  const cardSummary = prepareCardSummary(cardItems, attendance, visits);
  state.parsed = { visits, attendance, cardItems, pairDb, segments, cardSummary };

  renderSummary(segments, cardSummary);
  renderPreviews(segments, cardSummary);
  $('downloadBtn').disabled = false;
  log('미리보기가 준비되었습니다. 이상 없으면 완성본 다운로드를 누르세요.');
}

function renderSummary(segments, cardSummary) {
  const routeCount = segments.length;
  const totalDistance = segments.reduce((sum, s) => sum + (Number(s.distanceKm) || 0), 0);
  const totalToll = segments.reduce((sum, s) => sum + (Number(s.toll) || 0), 0);
  const totalFuel = [...cardSummary.fuelByDate.values()].reduce((a, b) => a + b, 0);
  const trainTotal = cardSummary.trains.reduce((a, b) => a + b.amount, 0);

  $('summary').style.display = 'grid';
  $('summary').innerHTML = `
    <div><span>자가용 이동구간</span><b>${routeCount}건</b></div>
    <div><span>운행거리 합계</span><b>${totalDistance.toFixed(1)} km</b></div>
    <div><span>주유/충전 합계</span><b>${formatWon(totalFuel)}</b></div>
    <div><span>SRT/KTX 합계</span><b>${formatWon(trainTotal)}</b></div>
  `;
  void totalToll;
}

function renderPreviews(segments, cardSummary) {
  const routeRows = segments.map(s => [s.dateKey, s.start, s.end, numberOrDash(s.distanceKm), formatWon(s.toll || 0), s.timeMin ? `${s.timeMin}분` : '-', s.source]);
  const fuelRows = [...cardSummary.fuelByDate.entries()].map(([date, amount]) => [date, formatWon(amount), '자가용 M열 입력']);
  const trainRows = cardSummary.trains.map(t => [t.dateKey, t.merchant, t.routeText, formatWon(t.amount - t.vat), formatWon(t.vat), formatWon(t.amount)]);
  const otherRows = cardSummary.others.slice(0, 20).map(t => [t.dateKey, t.merchant, t.category, formatWon(t.amount), '미입력']);

  $('preview').innerHTML = `
    ${tableBlock('자가용 이동경로 미리보기', ['날짜', '출발지', '도착지', '거리', '통행료', '예상시간', '기준'], routeRows)}
    ${tableBlock('주유/충전 카드내역 → 자가용 M열', ['날짜', '합계', '처리'], fuelRows)}
    ${tableBlock('SRT/KTX/SR → 대중교통 시트', ['날짜', '가맹점', '행선지', '금액', 'VAT', '계'], trainRows)}
    ${tableBlock('미입력 카드내역 참고', ['날짜', '가맹점', '업종', '금액', '처리'], otherRows)}
  `;
}

function tableBlock(title, headers, rows) {
  const body = rows.length
    ? rows.map(r => `<tr>${r.map((c, i) => `<td class="${i > 2 ? 'num' : ''}">${escapeHtml(c)}</td>`).join('')}</tr>`).join('')
    : `<tr><td colspan="${headers.length}">해당 건 없음</td></tr>`;
  return `
    <div class="card">
      <div class="section-title"><h3>${title}</h3><span class="badge">${rows.length}건</span></div>
      <table><thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${body}</tbody></table>
    </div>`;
}

function escapeHtml(v) {
  return String(v ?? '').replace(/[&<>'"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[ch]));
}

function formatWon(v) {
  const n = Number(v) || 0;
  return `${Math.round(n).toLocaleString('ko-KR')}원`;
}

function numberOrDash(v) {
  return v === null || v === undefined || v === '' ? '-' : Number(v).toFixed(1);
}

async function downloadWorkbook() {
  if (!state.parsed || !state.templateArrayBuffer) return;
  log('완성본 엑셀을 작성하는 중입니다.');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(state.templateArrayBuffer);
  const carSheet = workbook.getWorksheet(SHEETS.car);
  const transitSheet = workbook.getWorksheet(SHEETS.transit);
  if (!carSheet || !transitSheet) throw new Error('이동 양식에서 자가용/대중교통 시트를 찾지 못했습니다.');

  const { segments, cardSummary } = state.parsed;
  const allDates = [...new Set([...segments.map(s => s.dateKey), ...cardSummary.trains.map(t => t.dateKey)])].filter(Boolean).sort();
  const firstDate = allDates[0] || formatDate(new Date());
  const year = Number(firstDate.slice(0, 4));
  const month = Number(firstDate.slice(5, 7));

  carSheet.getCell('A1').value = `${year}년도 ( ${month} )월  자가용 이용 업무 수행 경비 사용내역서 `;
  transitSheet.getCell('A1').value = `${year}년도 ( ${month} )월  교통여비 사용내역서 `;

  clearCarRows(carSheet);
  writeCarRows(carSheet, segments, cardSummary);
  clearTransitRows(transitSheet);
  writeTransitRows(transitSheet, cardSummary.trains);

  workbook.calcProperties.fullCalcOnLoad = true;
  const buffer = await workbook.xlsx.writeBuffer();
  const filename = `이동경비_자동작성_${firstDate.slice(0, 7)}.xlsx`;
  saveAs(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename);
  log(`완성본 다운로드 완료: ${filename}`);
}

function clearCarRows(sheet) {
  for (let r = ROWS.carStart; r <= ROWS.carEnd; r++) {
    // A/B/D/F/J/M/N만 값 초기화. C/E/I/K/L은 원본 수식 유지, G/H는 고정값 유지.
    ['A', 'B', 'D', 'F', 'J', 'M', 'N'].forEach(col => { sheet.getCell(`${col}${r}`).value = null; });
  }
}

function writeCarRows(sheet, segments, cardSummary) {
  const fuelWrittenDates = new Set();
  const tollWrittenDates = new Set();
  let row = ROWS.carStart;
  for (const seg of segments) {
    if (row > ROWS.carEnd) break;
    sheet.getCell(`A${row}`).value = dateKeyToDate(seg.dateKey);
    sheet.getCell(`B${row}`).value = seg.start;
    sheet.getCell(`D${row}`).value = seg.end;

    if (seg.distanceKm !== null && seg.distanceKm !== undefined) sheet.getCell(`F${row}`).value = Number(seg.distanceKm);
    if (seg.toll !== null && seg.toll !== undefined) {
      sheet.getCell(`J${row}`).value = Math.round(Number(seg.toll) || 0);
      sheet.getCell(`N${row}`).value = Math.round(Number(seg.toll) || 0);
    }

    if (!fuelWrittenDates.has(seg.dateKey) && cardSummary.fuelByDate.has(seg.dateKey)) {
      sheet.getCell(`M${row}`).value = Math.round(cardSummary.fuelByDate.get(seg.dateKey));
      fuelWrittenDates.add(seg.dateKey);
    }

    // 카드내역에 하이패스/통행료가 있을 때는 해당 날짜 첫 이동행 N열에 우선 반영
    if (!tollWrittenDates.has(seg.dateKey) && cardSummary.tollByDate.has(seg.dateKey)) {
      sheet.getCell(`N${row}`).value = Math.round(cardSummary.tollByDate.get(seg.dateKey));
      tollWrittenDates.add(seg.dateKey);
    }
    row++;
  }
}

function clearTransitRows(sheet) {
  for (let r = ROWS.transitStart; r <= ROWS.transitEnd; r++) {
    ['A', 'B', 'C', 'D', 'F', 'G'].forEach(col => { sheet.getCell(`${col}${r}`).value = null; });
  }
}

function writeTransitRows(sheet, trains) {
  let row = ROWS.transitStart;
  for (const item of trains) {
    if (row > ROWS.transitEnd) break;
    sheet.getCell(`A${row}`).value = dateKeyToDate(item.dateKey);
    sheet.getCell(`B${row}`).value = '기차';
    sheet.getCell(`C${row}`).value = item.routeText || '확인 필요';
    sheet.getCell(`D${row}`).value = '업무 이동';
    sheet.getCell(`F${row}`).value = Math.round(item.amount - item.vat);
    sheet.getCell(`G${row}`).value = Math.round(item.vat || 0);
    row++;
  }
}

$('previewBtn').addEventListener('click', () => buildPreview().catch(err => {
  console.error(err);
  log(`오류: ${err.message}`);
  alert(err.message);
}));

$('downloadBtn').addEventListener('click', () => downloadWorkbook().catch(err => {
  console.error(err);
  log(`오류: ${err.message}`);
  alert(err.message);
}));

$('resetBtn').addEventListener('click', () => {
  state.parsed = null;
  state.templateArrayBuffer = null;
  state.routeResults = [];
  state.logs = [];
  ['templateFile', 'visitFile', 'attendanceFile', 'cardFile'].forEach(id => { $(id).value = ''; });
  $('preview').innerHTML = '';
  $('summary').style.display = 'none';
  $('downloadBtn').disabled = true;
  $('apiBadge').className = 'badge amber';
  $('apiBadge').textContent = 'TMAP 상태 확인 전';
  $('status').textContent = '파일을 업로드하고 “미리보기 생성”을 눌러주세요.';
});
