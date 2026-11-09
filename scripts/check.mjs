// 發布前檢查：倉庫內不得有 Excel、未加密資料或密碼檔；資料檔必須是密文格式。
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const tracked = execFileSync('git', ['ls-files', '--cached', '--others', '--exclude-standard'], { cwd: root, encoding: 'utf8' }).split('\n').filter(Boolean);
const problems = [];
for (const f of tracked) {
  if (/\.(xlsx|xlsm|xls|csv)$/i.test(f)) problems.push(`不應發布試算表：${f}`);
  if (/(^|\/)\.env|site-password/i.test(f)) problems.push(`不應發布秘密檔：${f}`);
  if (f.startsWith('data/') && f !== 'data/data.enc.json') problems.push(`data/ 只允許 data.enc.json：${f}`);
}
const env = JSON.parse(fs.readFileSync(path.join(root, 'data/data.enc.json'), 'utf8'));
if (env.format !== 'ama-enc/1' || !env.data || !env.salt || !env.iv) problems.push('data.enc.json 不是預期的密文格式。');
const plainKeys = ['skus', 'bom', 'erp', 'issues', 'raw'];
if (plainKeys.some(k => k in env)) problems.push('data.enc.json 含有未加密欄位。');
// 程式檔中不應出現明顯的營運數字樣本（以 ERP 品項編碼格式粗查）
for (const f of tracked.filter(f => /\.(js|html|md|css)$/.test(f) && !f.startsWith('vendor/'))) {
  const t = fs.readFileSync(path.join(root, f), 'utf8');
  const codes = t.match(/\b(?:1FN|FN|SF|NP|WF)[0-9A-Z]{6,9}\b/g) || [];
  if (codes.length) problems.push(`${f} 含疑似 ERP 品項編碼：${[...new Set(codes)].slice(0, 5).join(', ')}`);
}
if (problems.length) { console.error(problems.join('\n')); process.exit(1); }
console.log(`檢查通過：${tracked.length} 個檔案，資料為密文（${env.cipher}，${env.iterations} 次 PBKDF2）。`);
