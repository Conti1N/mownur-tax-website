// api/shared/_multipart.js
// Parses multipart/form-data from an Azure Functions request body (Buffer).

const MAX_FILE_SIZE = 25 * 1024 * 1024;

/**
 * Azure Functions v4 passes req.body as a Buffer (or string).
 * Pass the raw body buffer and the Content-Type header value.
 */
function parseMultipartBuffer(bodyBuffer, contentType) {
  const boundaryMatch = (contentType || '').match(/boundary=([^\s;]+)/);
  if (!boundaryMatch) throw new Error('No boundary in Content-Type');
  return parseBody(
    Buffer.isBuffer(bodyBuffer) ? bodyBuffer : Buffer.from(bodyBuffer),
    boundaryMatch[1]
  );
}

function parseBody(body, boundary) {
  const fields = {};
  const files = [];

  const delimiter = Buffer.from(`--${boundary}`);
  const CRLFCRLF = Buffer.from('\r\n\r\n');
  let offset = 0;

  while (offset < body.length) {
    const delimPos = indexOf(body, delimiter, offset);
    if (delimPos === -1) break;
    offset = delimPos + delimiter.length;
    if (body.slice(offset, offset + 2).toString() === '--') break;
    if (body.slice(offset, offset + 2).toString() === '\r\n') offset += 2;

    const headerEnd = indexOf(body, CRLFCRLF, offset);
    if (headerEnd === -1) break;

    const headerStr = body.slice(offset, headerEnd).toString();
    offset = headerEnd + 4;

    const nextDelim = indexOf(body, delimiter, offset);
    if (nextDelim === -1) break;

    const partBody = body.slice(offset, nextDelim - 2);
    offset = nextDelim;

    const headers = parsePartHeaders(headerStr);
    const disposition = headers['content-disposition'] || '';
    const nameMatch = disposition.match(/name="([^"]+)"/);
    const filenameMatch = disposition.match(/filename="([^"]+)"/);

    if (!nameMatch) continue;
    const name = nameMatch[1];

    if (filenameMatch) {
      if (partBody.length > MAX_FILE_SIZE) {
        throw new Error(`File "${filenameMatch[1]}" exceeds 25MB limit`);
      }
      files.push({
        fieldName: name,
        filename: filenameMatch[1],
        mimetype: headers['content-type'] || 'application/octet-stream',
        buffer: partBody,
        category: name.startsWith('file_') ? name.replace('file_', '').replace(/_/g, ' ') : 'General',
      });
    } else {
      const value = partBody.toString('utf8');
      if (name in fields) {
        fields[name] = Array.isArray(fields[name]) ? [...fields[name], value] : [fields[name], value];
      } else {
        fields[name] = value;
      }
    }
  }

  return { fields, files };
}

function parsePartHeaders(headerStr) {
  const headers = {};
  for (const line of headerStr.split('\r\n')) {
    const colonIdx = line.indexOf(':');
    if (colonIdx === -1) continue;
    headers[line.slice(0, colonIdx).toLowerCase().trim()] = line.slice(colonIdx + 1).trim();
  }
  return headers;
}

function indexOf(buffer, search, offset = 0) {
  for (let i = offset; i <= buffer.length - search.length; i++) {
    let found = true;
    for (let j = 0; j < search.length; j++) {
      if (buffer[i + j] !== search[j]) { found = false; break; }
    }
    if (found) return i;
  }
  return -1;
}

module.exports = { parseMultipartBuffer };
