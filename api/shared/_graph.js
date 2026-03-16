// api/shared/_graph.js
// Microsoft Graph API helpers using Managed Identity — no client secret required.
// DefaultAzureCredential automatically uses the Azure Static Web App's system-assigned
// Managed Identity when deployed, and falls back to az CLI / environment for local dev.

const https = require('https');

// Managed Identity token endpoint (available inside Azure at runtime)
const IMDS_ENDPOINT = 'http://169.254.169.254/metadata/instance';
const TOKEN_ENDPOINT = 'http://169.254.169.254/metadata/token?api-version=2018-02-01&resource=https://graph.microsoft.com/';

let _cachedToken = null;
let _tokenExpiry = 0;

/**
 * Get a Managed Identity access token for Microsoft Graph.
 * In local dev: falls back to client credentials if MI env vars aren't present.
 */
async function getAccessToken() {
  const now = Date.now();
  if (_cachedToken && now < _tokenExpiry - 60_000) {
    return _cachedToken;
  }

  // Try Managed Identity first (Azure-hosted environment)
  try {
    const data = await httpGet(TOKEN_ENDPOINT, { Metadata: 'true' });
    _cachedToken = data.access_token;
    _tokenExpiry = now + parseInt(data.expires_in, 10) * 1000;
    return _cachedToken;
  } catch {
    // Not running on Azure — fall back to client credentials for local dev
    return await getTokenViaClientCredentials();
  }
}

async function getTokenViaClientCredentials() {
  const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;
  if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
    throw new Error(
      'Not running on Azure (no Managed Identity) and no local dev credentials found. ' +
      'Set AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET in local.settings.json for dev.'
    );
  }

  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: AZURE_CLIENT_ID,
    client_secret: AZURE_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
  });

  const data = await httpsPost(
    `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`,
    body.toString(),
    { 'Content-Type': 'application/x-www-form-urlencoded' }
  );

  _cachedToken = data.access_token;
  _tokenExpiry = Date.now() + data.expires_in * 1000;
  return _cachedToken;
}

/**
 * Make a Microsoft Graph API request.
 * @param {string} token
 * @param {string} method
 * @param {string} path - Graph path or full URL
 * @param {object} [body]
 */
async function graphRequest(token, method, path, body) {
  const url = path.startsWith('https://') ? path : `https://graph.microsoft.com/v1.0${path}`;
  const headers = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
  };

  return new Promise((resolve, reject) => {
    const bodyStr = body ? JSON.stringify(body) : null;
    if (bodyStr) headers['Content-Length'] = Buffer.byteLength(bodyStr);

    const urlObj = new URL(url);
    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method,
      headers,
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => (data += chunk));
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try { resolve(data ? JSON.parse(data) : {}); }
          catch { resolve({}); }
        } else {
          const err = new Error(`Graph API ${res.statusCode}: ${data}`);
          err.status = res.statusCode;
          reject(err);
        }
      });
    });
    req.on('error', reject);
    if (bodyStr) req.write(bodyStr);
    req.end();
  });
}

/**
 * Upload a file to a OneDrive folder (creates subfolder if needed).
 * Uses upload session for files > 4MB.
 */
async function uploadFileToOneDrive(token, parentFolderId, subfolder, file) {
  const MAX_SIMPLE = 4 * 1024 * 1024;
  const path = subfolder
    ? `/me/drive/items/${parentFolderId}:/${encodeURIComponent(subfolder)}/${encodeURIComponent(file.filename)}:`
    : `/me/drive/items/${parentFolderId}:/${encodeURIComponent(file.filename)}:`;

  const fullUrl = `https://graph.microsoft.com/v1.0${path}/content`;

  if (file.buffer.length <= MAX_SIMPLE) {
    return await uploadBinary(token, fullUrl, file.buffer, file.mimetype);
  } else {
    const session = await graphRequest(token, 'POST',
      `${path}/createUploadSession`,
      { item: { '@microsoft.graph.conflictBehavior': 'rename' } }
    );
    return await uploadInChunks(session.uploadUrl, file.buffer);
  }
}

async function uploadBinary(token, url, buffer, contentType) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': contentType || 'application/octet-stream',
        'Content-Length': buffer.length,
      },
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => (data += chunk));
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve(data ? JSON.parse(data) : {});
        } else {
          const err = new Error(`Upload failed ${res.statusCode}: ${data}`);
          err.status = res.statusCode;
          reject(err);
        }
      });
    });
    req.on('error', reject);
    req.write(buffer);
    req.end();
  });
}

async function uploadInChunks(uploadUrl, buffer) {
  const chunkSize = 5 * 1024 * 1024;
  const total = buffer.length;
  let offset = 0;
  let result;

  while (offset < total) {
    const end = Math.min(offset + chunkSize, total);
    const chunk = buffer.slice(offset, end);
    result = await new Promise((resolve, reject) => {
      const urlObj = new URL(uploadUrl);
      const options = {
        hostname: urlObj.hostname,
        path: urlObj.pathname + urlObj.search,
        method: 'PUT',
        headers: {
          'Content-Length': chunk.length,
          'Content-Range': `bytes ${offset}-${end - 1}/${total}`,
          'Content-Type': 'application/octet-stream',
        },
      };
      const req = https.request(options, (res) => {
        let data = '';
        res.on('data', d => (data += d));
        res.on('end', () => {
          if ([200, 201, 202].includes(res.statusCode)) {
            resolve(data ? JSON.parse(data) : {});
          } else {
            reject(new Error(`Chunk upload failed ${res.statusCode}: ${data}`));
          }
        });
      });
      req.on('error', reject);
      req.write(chunk);
      req.end();
    });
    offset = end;
  }
  return result;
}

// HTTP GET helper (for IMDS token endpoint — http, not https)
function httpGet(url, headers) {
  return new Promise((resolve, reject) => {
    const lib = url.startsWith('https') ? https : require('http');
    const urlObj = new URL(url);
    const options = { hostname: urlObj.hostname, path: urlObj.pathname + urlObj.search, method: 'GET', headers };
    const req = lib.request(options, (res) => {
      let data = '';
      res.on('data', chunk => (data += chunk));
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch { reject(new Error(`Non-JSON response: ${data}`)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

function httpsPost(url, body, headers) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname,
      method: 'POST',
      headers: { ...headers, 'Content-Length': Buffer.byteLength(body) },
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => (data += chunk));
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch { reject(new Error(`Failed to parse token response: ${data}`)); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

module.exports = { getAccessToken, graphRequest, uploadFileToOneDrive };
