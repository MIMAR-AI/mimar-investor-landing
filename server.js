import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('.', import.meta.url));
const files = new Map([
  ['/', 'index.html'],
  ['/index.html', 'index.html'],
  ['/styles.css', 'styles.css']
]);
const types = { '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8' };

createServer(async (req, res) => {
  const path = new URL(req.url, `http://${req.headers.host}`).pathname;
  if (path === '/healthz') {
    res.writeHead(200, { 'content-type': 'application/json; charset=utf-8' });
    return res.end('{"status":"ok"}');
  }
  const file = files.get(path);
  if (!file || !['GET', 'HEAD'].includes(req.method)) {
    res.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' });
    return res.end('Not found');
  }
  try {
    const body = await readFile(join(root, file));
    res.writeHead(200, {
      'content-type': types[extname(file)],
      'cache-control': file === 'index.html' ? 'no-cache' : 'public, max-age=3600',
      'x-content-type-options': 'nosniff',
      'x-frame-options': 'SAMEORIGIN',
      'referrer-policy': 'strict-origin-when-cross-origin'
    });
    res.end(req.method === 'HEAD' ? undefined : body);
  } catch {
    res.writeHead(500, { 'content-type': 'text/plain; charset=utf-8' });
    res.end('Server error');
  }
}).listen(process.env.PORT || 3000, () => console.log('MIMAR investor landing is running'));
