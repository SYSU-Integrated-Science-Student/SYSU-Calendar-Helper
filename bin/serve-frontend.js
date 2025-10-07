#!/usr/bin/env node
import http from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { extname, join, resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';
import { parseScheduleFromPackage, buildIcs } from '../src/schedule.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const projectRoot = resolve(__dirname, '..');
const frontendDir = join(projectRoot, 'frontend');

function parseOptions(argv) {
  const options = { port: 4173, open: false, host: '127.0.0.1' };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    switch (arg) {
      case '--port':
      case '-p':
        options.port = parseInt(argv[++i], 10);
        break;
      case '--host':
        options.host = argv[++i];
        break;
      case '--open':
      case '-o':
        options.open = true;
        break;
      case '--help':
      case '-h':
        options.help = true;
        break;
      default:
        throw new Error(`Unknown option: ${arg}`);
    }
  }
  return options;
}

function printHelp() {
  console.log(`Usage: serve-frontend [options]\n\n` +
    `Options:\n` +
    `  -p, --port <number>   Port to listen on (default: 4173)\n` +
    `      --host <host>     Host address (default: 127.0.0.1)\n` +
    `  -o, --open            Open default browser after starting\n` +
    `  -h, --help            Show this message`);
}

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.mjs': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.txt': 'text/plain; charset=utf-8'
};

function resolveFilePath(urlPath = '/') {
  const [rawPath] = urlPath.split('?');
  if (rawPath === '/' || rawPath === '' || rawPath === '/index.html') {
    return join(frontendDir, 'index.html');
  }

  let decoded;
  try {
    decoded = decodeURIComponent(rawPath);
  } catch (error) {
    decoded = rawPath;
  }

  const normalizedSegments = decoded
    .replace(/^\/+/, '')
    .split('/')
    .filter(Boolean)
    .filter((segment) => segment !== '..');

  if (!normalizedSegments.length) {
    return join(frontendDir, 'index.html');
  }

  const candidate = resolve(frontendDir, normalizedSegments.join('/'));
  if (!candidate.startsWith(frontendDir)) {
    return null;
  }
  return candidate;
}

async function serveFile(res, filePath) {
  try {
    const stats = await stat(filePath);
    if (stats.isDirectory()) {
      const indexPath = join(filePath, 'index.html');
      await serveFile(res, indexPath);
      return;
    }
    const ext = extname(filePath);
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';
    const data = await readFile(filePath);
    res.writeHead(200, { 'Content-Type': contentType, 'Content-Length': data.length });
    res.end(data);
  } catch (error) {
    if (error.code === 'ENOENT') {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not Found');
      return;
    }
    console.error(error);
    res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Internal Server Error');
  }
}

function collectRequestBody(req, limit = 10 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on('data', (chunk) => {
      size += chunk.length;
      if (size > limit) {
        reject(new Error('Request body too large'));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => {
      resolve(Buffer.concat(chunks));
    });
    req.on('error', reject);
  });
}

async function handleParseRequest(req, res, url) {
  try {
    const startDate = url.searchParams.get('start') || '2025-09-08';
    const timeZone = url.searchParams.get('tz') || 'Asia/Shanghai';
    const fileName = req.headers['x-filename'] ? String(req.headers['x-filename']) : 'schedule.docx';
    const buffer = await collectRequestBody(req);

    const { title, events } = await parseScheduleFromPackage(buffer, {
      startDate,
      timeZone
    });
    const ics = buildIcs(events, timeZone);

    const response = {
      title,
      startDate,
      timeZone,
      fileName,
      eventCount: events.length,
      occurrenceCount: events.reduce((sum, event) => sum + event.occurrences.length, 0),
      events,
      ics
    };

    res.writeHead(200, {
      'Content-Type': 'application/json; charset=utf-8',
      'Access-Control-Allow-Origin': '*'
    });
    res.end(JSON.stringify(response));
  } catch (error) {
    console.error('Parse request failed:', error);
    res.writeHead(400, {
      'Content-Type': 'application/json; charset=utf-8',
      'Access-Control-Allow-Origin': '*'
    });
    res.end(JSON.stringify({ error: error.message || String(error) }));
  }
}

function openBrowser(url) {
  const command = process.platform === 'darwin'
    ? 'open'
    : process.platform === 'win32'
      ? 'cmd'
      : 'xdg-open';
  const args = process.platform === 'win32' ? ['/c', 'start', '""', url] : [url];
  spawn(command, args, { stdio: 'ignore', detached: true }).unref();
}

async function startServer() {
  let options;
  try {
    options = parseOptions(process.argv.slice(2));
  } catch (error) {
    console.error(error.message);
    process.exit(1);
  }

  if (options.help) {
    printHelp();
    return;
  }

  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url || '/', `http://${req.headers.host || `${options.host}:${options.port}`}`);

    if (req.method === 'POST' && url.pathname === '/api/parse') {
      await handleParseRequest(req, res, url);
      return;
    }

    if (req.method === 'OPTIONS' && url.pathname === '/api/parse') {
      res.writeHead(204, {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type,X-Filename'
      });
      res.end();
      return;
    }

    const filePath = resolveFilePath(url.pathname);
    if (!filePath) {
      res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Bad Request');
      return;
    }
    await serveFile(res, filePath);
  });

  server.listen(options.port, options.host, () => {
    const url = `http://${options.host}:${options.port}`;
    console.log(`Serving frontend on ${url}`);
    if (options.open) {
      openBrowser(url);
    }
  });
}

startServer();
