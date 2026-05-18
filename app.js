export default async function handler(req, res) {
  if (req.method === 'OPTIONS') {
    res.setHeader('Allow', 'POST, OPTIONS');
    res.status(204).end();
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ ok: false, error: 'POST만 지원합니다.' });
    return;
  }

  const appKey = process.env.TMAP_APP_KEY;
  if (!appKey) {
    res.status(200).json({ ok: false, error: 'Vercel 환경변수 TMAP_APP_KEY가 없습니다.' });
    return;
  }

  try {
    const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
    const start = body.start || {};
    const end = body.end || {};

    const startX = Number(start.lon);
    const startY = Number(start.lat);
    const endX = Number(end.lon);
    const endY = Number(end.lat);

    if (![startX, startY, endX, endY].every(Number.isFinite)) {
      res.status(200).json({ ok: false, error: '출발/도착 좌표가 올바르지 않습니다.' });
      return;
    }

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 12000);

    const response = await fetch('https://apis.openapi.sk.com/tmap/routes?version=1&format=json', {
      method: 'POST',
      signal: controller.signal,
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'appKey': appKey
      },
      body: JSON.stringify({
        startX: String(startX),
        startY: String(startY),
        endX: String(endX),
        endY: String(endY),
        startName: start.name || '출발지',
        endName: end.name || '도착지',
        reqCoordType: 'WGS84GEO',
        resCoordType: 'WGS84GEO',
        searchOption: '0',
        trafficInfo: 'Y'
      })
    });

    clearTimeout(timeout);
    const data = await response.json().catch(() => ({}));

    if (!response.ok) {
      res.status(200).json({
        ok: false,
        error: data?.error?.message || data?.message || `TMAP HTTP ${response.status}`
      });
      return;
    }

    const summary = extractSummary(data);
    if (!summary || !summary.totalDistance) {
      res.status(200).json({ ok: false, error: 'TMAP 응답에서 거리 정보를 찾지 못했습니다.' });
      return;
    }

    res.status(200).json({
      ok: true,
      distanceKm: Math.round((summary.totalDistance / 1000) * 10) / 10,
      timeMin: Math.round((summary.totalTime || 0) / 60),
      toll: Number(summary.totalFare || 0),
      taxiFare: Number(summary.taxiFare || 0)
    });
  } catch (err) {
    const message = err?.name === 'AbortError' ? 'TMAP 호출 시간 초과' : (err?.message || String(err));
    res.status(200).json({ ok: false, error: message });
  }
}

function extractSummary(data) {
  if (!data) return null;
  if (data.features && Array.isArray(data.features)) {
    const first = data.features.find((f) => f?.properties?.totalDistance);
    if (first) return first.properties;
  }
  if (data.properties?.totalDistance) return data.properties;
  if (data.totalDistance) return data;
  return null;
}
