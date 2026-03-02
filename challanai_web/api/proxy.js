export const config = { runtime: 'edge' };

export default async function handler(req) {
  const url = new URL(req.url);
  const path = url.pathname.replace('/api', '');
  const targetUrl = process.env.BACKEND_URL + path + url.search;

  return fetch(targetUrl, {
    method: req.method,
    headers: req.headers,
    body: req.method !== 'GET' ? req.body : undefined,
    duplex: 'half',
  });
}
