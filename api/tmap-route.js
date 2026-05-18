const TMAP_URL = 'https://apis.openapi.sk.com/tmap/routes?version=1&format=json';

function send(res, status, payload) {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.end(JSON.stringify(payload));
}

function parseTotalFromTmap(json) {
  const features = Array.isArray(json?.features) ? json.features : [];
  const found = features.find(f => f?.properties && (f.properties.totalDistance !== undefined || f.properties.totalTime !== undefined));
  const props = found?.properties || json?.properties || {};
  const distanceM = Number(props.totalDistance ?? 0);
  const timeSec = Number(props.totalTime ?? 0);
  const toll = Number(props.totalFare ?? props.totalToll ?? 0);

  return {
    distanceKm: distanceM ? Math.round((distanceM / 1000) * 10) / 10 : null,
    timeMin: timeSec ? Math.round(timeSec / 60) : null,
    toll: Number.isFinite(toll) ? toll : 0
  };
}

async function callTmap(route, appKey) {
  const body = {
    startX: String(route.start.lon),
    startY: String(route.start.lat),
    endX: String(route.end.lon),
    endY: String(route.end.lat),
    reqCoordType: 'WGS84GEO',
    resCoordType: 'WGS84GEO',
    searchOption: '0',
    startName: route.startName || '출발지',
    endName: route.endName || '도착지'
  };

  const response = await fetch(TMAP_URL, {
    method: 'POST',
    headers: {
      appKey,
      'Content-Type': 'application/json',
      Accept: 'application/json'
    },
    body: JSON.stringify(body)
  });

  const text = await response.text();
  let json;
  try {
    json = text ? JSON.parse(text) : {};
  } catch (e) {
    json = { raw: text };
  }

  if (!response.ok) {
    const msg = json?.error?.message || json?.message || text || `HTTP ${response.status}`;
    throw new Error(msg);
  }

  return parseTotalFromTmap(json);
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    send(res, 405, { ok: false, message: 'POST만 지원합니다.' });
    return;
  }

  const appKey = process.env.TMAP_APP_KEY;
  if (!appKey) {
    send(res, 200, { ok: false, message: 'TMAP_APP_KEY 환경변수가 없습니다.', results: [] });
    return;
  }

  try {
    const chunks = [];
    for await (const chunk of req) chunks.push(chunk);
    const raw = Buffer.concat(chunks).toString('utf8');
    const body = raw ? JSON.parse(raw) : {};
    const routes = Array.isArray(body.routes) ? body.routes : [];
    if (!routes.length) {
      send(res, 400, { ok: false, message: 'routes 배열이 비어 있습니다.' });
      return;
    }

    const results = [];
    for (const route of routes.slice(0, 50)) {
      try {
        const summary = await callTmap(route, appKey);
        results.push({ id: route.id, ok: true, ...summary });
      } catch (error) {
        results.push({ id: route.id, ok: false, message: error.message });
      }
    }

    send(res, 200, { ok: true, results });
  } catch (error) {
    send(res, 500, { ok: false, message: error.message });
  }
}
