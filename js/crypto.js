/*
 * 資料加解密（瀏覽器與 Node 22 共用）：PBKDF2-SHA256 推導金鑰，AES-256-GCM 加密，內容先以 gzip 壓縮。
 * 公開站只放密文；密碼不寫入程式或倉庫。
 */
(function (root) {
  'use strict';
  const ITERATIONS = 1200000;
  const subtle = () => (root.crypto && root.crypto.subtle) || (typeof globalThis !== 'undefined' && globalThis.crypto && globalThis.crypto.subtle);
  const enc = new TextEncoder();
  const dec = new TextDecoder();
  function b64(bytes) { let s = ''; const a = new Uint8Array(bytes); for (let i = 0; i < a.length; i += 0x8000) s += String.fromCharCode.apply(null, a.subarray(i, i + 0x8000)); return btoa(s); }
  function unb64(str) { const s = atob(str); const a = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) a[i] = s.charCodeAt(i); return a; }
  async function pipe(bytes, stream) { return new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(stream)).arrayBuffer()); }
  async function key(password, salt, iterations) {
    const base = await subtle().importKey('raw', enc.encode(password), 'PBKDF2', false, ['deriveKey']);
    return subtle().deriveKey({ name: 'PBKDF2', hash: 'SHA-256', salt, iterations }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  }
  async function encrypt(obj, password, meta = {}) {
    const salt = root.crypto.getRandomValues(new Uint8Array(16));
    const iv = root.crypto.getRandomValues(new Uint8Array(12));
    const packed = await pipe(enc.encode(JSON.stringify(obj)), new CompressionStream('gzip'));
    const ct = await subtle().encrypt({ name: 'AES-GCM', iv }, await key(password, salt, ITERATIONS), packed);
    return { format: 'ama-enc/1', kdf: 'PBKDF2-SHA256', iterations: ITERATIONS, cipher: 'AES-256-GCM', gzip: true, salt: b64(salt), iv: b64(iv), data: b64(ct), meta };
  }
  async function decrypt(envelope, password) {
    if (!envelope || envelope.format !== 'ama-enc/1') throw new Error('資料檔格式不符。');
    let plain;
    try {
      plain = await subtle().decrypt({ name: 'AES-GCM', iv: unb64(envelope.iv) }, await key(password, unb64(envelope.salt), envelope.iterations), unb64(envelope.data));
    } catch (e) { const err = new Error('密碼不正確。'); err.code = 'BAD_PASSWORD'; throw err; }
    const bytes = envelope.gzip ? await pipe(new Uint8Array(plain), new DecompressionStream('gzip')) : new Uint8Array(plain);
    return JSON.parse(dec.decode(bytes));
  }
  const api = { encrypt, decrypt, ITERATIONS };
  root.AMACrypto = api;
})(typeof window === 'undefined' ? globalThis : window);
