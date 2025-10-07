import { parseScheduleFromPackage, buildIcs } from './src/schedule.js';

const JSON_HEADERS = {
  'content-type': 'application/json; charset=utf-8',
  'access-control-allow-origin': '*',
};

function errorResponse(message, status = 400) {
  return new Response(JSON.stringify({ error: message }), {
    status,
    headers: JSON_HEADERS,
  });
}

async function handleParse(request, env) {
  if (request.method === 'OPTIONS') {
    return new Response(null, {
      status: 204,
      headers: {
        'access-control-allow-origin': '*',
        'access-control-allow-methods': 'POST, OPTIONS',
        'access-control-allow-headers': 'content-type,x-filename',
      },
    });
  }

  if (request.method !== 'POST') {
    return errorResponse('Only POST is supported', 405);
  }

  const url = new URL(request.url);
  const startDate = url.searchParams.get('start') || '2025-09-08';
  const timeZone = url.searchParams.get('tz') || 'Asia/Shanghai';

  let buffer;
  try {
    buffer = await request.arrayBuffer();
  } catch (error) {
    return errorResponse('Failed to read request body');
  }

  try {
    const { title, events } = await parseScheduleFromPackage(buffer, {
      startDate,
      timeZone,
    });
    const ics = buildIcs(events, timeZone);
    const payload = {
      title,
      startDate,
      timeZone,
      eventCount: events.length,
      occurrenceCount: events.reduce((sum, ev) => sum + ev.occurrences.length, 0),
      events,
      ics,
    };
    return new Response(JSON.stringify(payload), {
      status: 200,
      headers: JSON_HEADERS,
    });
  } catch (error) {
    return errorResponse(error.message || 'Failed to parse schedule');
  }
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    if (url.pathname.startsWith('/api/parse')) {
      return handleParse(request, env);
    }

    if (request.method === 'OPTIONS') {
      return new Response(null, {
        status: 204,
        headers: {
          'access-control-allow-origin': '*',
          'access-control-allow-methods': 'GET, POST, OPTIONS',
          'access-control-allow-headers': 'content-type,x-filename',
        },
      });
    }

    if (env.ASSETS) {
      try {
        const assetResponse = await env.ASSETS.fetch(request);
        if (assetResponse.status !== 404) return assetResponse;
      } catch (error) {
        console.error('Asset fetch failed:', error);
      }
    }

    if (url.pathname === '/' || url.pathname === '') {
      return new Response('Not Found', { status: 404 });
    }

    return new Response('Not Found', { status: 404 });
  },
};
