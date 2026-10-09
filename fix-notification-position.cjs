const fs = require('fs');
const path = require('path');

const target = process.argv[2] || path.join(process.cwd(), 'src', 'app.jsx');
let code = fs.readFileSync(target, 'utf8');
const original = code;

const oldCss = `.rgt-notice-stack{position:fixed;top:24px;right:24px;z-index:10000;width:min(390px,calc(100vw - 32px));pointer-events:none}.rgt-notice{pointer-events:auto;display:grid;grid-template-columns:42px 1fr 30px;gap:12px;align-items:center;padding:14px 14px 14px 16px;border:1px solid rgba(148,163,184,.28);border-radius:16px;background:rgba(15,23,42,.96);box-shadow:0 24px 70px rgba(15,23,42,.32);color:#fff;backdrop-filter:blur(16px);animation:rgtNoticeIn .28s cubic-bezier(.2,.8,.2,1)}`;
const newCss = `.rgt-notice-stack{position:fixed;top:0;left:0;right:0;z-index:10000;width:100%;pointer-events:none}.rgt-notice{pointer-events:auto;min-height:64px;display:grid;grid-template-columns:42px minmax(0,1fr) 30px;gap:12px;align-items:center;padding:11px max(24px,calc((100vw - 1440px)/2 + 24px));border:0;border-bottom:1px solid rgba(148,163,184,.32);border-radius:0;background:linear-gradient(90deg,rgba(8,47,73,.98),rgba(15,23,42,.98));box-shadow:0 8px 28px rgba(15,23,42,.2);color:#fff;backdrop-filter:blur(16px);animation:rgtNoticeDown .3s cubic-bezier(.2,.8,.2,1)}`;

if (!code.includes(oldCss)) {
  throw new Error('Expected notification CSS was not found. Run the fixed update-app-speed.cjs first, then run this file. No source file was changed.');
}
code = code.replace(oldCss, newCss);
code = code.replace(`@keyframes rgtNoticeIn{from{opacity:0;transform:translate3d(24px,-8px,0)}to{opacity:1;transform:none}}`, `@keyframes rgtNoticeDown{from{opacity:0;transform:translate3d(0,-100%,0)}to{opacity:1;transform:none}}`);
code = code.replace(`@media(max-width:640px){.rgt-notice-stack{top:14px;right:16px}}`, `@media(max-width:640px){.rgt-notice{min-height:58px;padding:10px 14px;grid-template-columns:38px minmax(0,1fr) 28px}.rgt-notice-icon{width:36px;height:36px}.rgt-notice p{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}`);

const temp = `${target}.tmp`;
fs.writeFileSync(temp, code);
fs.renameSync(temp, target);
console.log(`Updated successfully: ${target}`);
console.log('Notification is now a fixed full-width top status bar.');
console.log(`Original bytes: ${Buffer.byteLength(original)} | Updated bytes: ${Buffer.byteLength(code)}`);
