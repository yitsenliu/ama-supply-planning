// 將產銷 Excel 解析並加密為 data/data.enc.json。
// 用法：AMA_SITE_PASSWORD='…' node scripts/build-data.mjs <xlsx 路徑>
// 密碼只從環境變數讀取，不寫入任何檔案。
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as XLSX from 'xlsx';

XLSX.set_fs(fs);
const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..');
await import(path.join(root, 'js/parser.js'));
await import(path.join(root, 'js/crypto.js'));

const file = process.argv[2];
const password = process.env.AMA_SITE_PASSWORD;
if (!file || !fs.existsSync(file)) { console.error('請提供 Excel 檔路徑。'); process.exit(2); }
if (!password || password.length < 6) { console.error('請以環境變數 AMA_SITE_PASSWORD 提供至少 6 碼的密碼。'); process.exit(2); }

const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellFormula: true });
const data = globalThis.AMAParser.parseWorkbook(wb, { file: path.basename(file), parsedAt: new Date().toISOString() });
const envelope = await globalThis.AMACrypto.encrypt(data, password, { builtAt: new Date().toISOString(), schema: data.schema });
fs.mkdirSync(path.join(root, 'data'), { recursive: true });
fs.writeFileSync(path.join(root, 'data/data.enc.json'), JSON.stringify(envelope));

// 驗證：重新解密並比對
const back = await globalThis.AMACrypto.decrypt(envelope, password);
if (JSON.stringify(back) !== JSON.stringify(data)) { console.error('解密驗證失敗。'); process.exit(1); }
console.log(JSON.stringify({ ok: true, sheets: data.sheets.length, skus: data.plan?.skus.length, bom: data.bom.length, erpItems: data.erp.items.length, issues: data.issues.length, bytes: fs.statSync(path.join(root, 'data/data.enc.json')).size }));
