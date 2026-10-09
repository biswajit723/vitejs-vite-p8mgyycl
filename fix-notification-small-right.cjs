const fs = require('fs');
const path = require('path');

const target = process.argv[2] || path.join(process.cwd(), 'src', 'app.jsx');
let code = fs.readFileSync(target, 'utf8');
const original = code;

const topBarStack = `.rgt-notice-stack{position:fixed;top:0;left:0;right:0;z-index:10000;width:100%;pointer-events:none}`;
const topBarNotice = `.rgt-notice{pointer-events:auto;min-height:64px;display:grid;grid-template-columns:42px minmax(0,1fr) 30px;gap:12px;align-items:center;padding:11px max(24px,calc((100vw - 1440px)/2 + 24px));border:0;border-bottom:1px solid rgba(148,163,184,.32);border-radius:0;background:linear-gradient(90deg,rgba(8,47,73,.98),rgba(15,23,42,.98));box-shadow:0 8px 28px rgba(15,23,42,.2);color:#fff;backdrop-filter:blur(16px);animation:rgtNoticeDown .3s cubic-bezier(.2,.8,.2,1)}`;

const originalStack = `.rgt-notice-stack{position:fixed;top:24px;right:24px;z-index:10000;width:min(390px,calc(100vw - 32px));pointer-events:none}`;
const currentSmallLeftStack = `.rgt-notice-stack{position:fixed;top:12px;left:12px;z-index:10000;width:min(320px,calc(100vw - 24px));pointer-events:none}`;
const originalNotice = `.rgt-notice{pointer-events:auto;display:grid;grid-template-columns:42px 1fr 30px;gap:12px;align-items:center;padding:14px 14px 14px 16px;border:1px solid rgba(148,163,184,.28);border-radius:16px;background:rgba(15,23,42,.96);box-shadow:0 24px 70px rgba(15,23,42,.32);color:#fff;backdrop-filter:blur(16px);animation:rgtNoticeIn .28s cubic-bezier(.2,.8,.2,1)}`;

const smallStack = `.rgt-notice-stack{position:absolute;top:12px;right:12px;z-index:20;width:min(320px,calc(100vw - 24px));pointer-events:none}`;
const smallNotice = `.rgt-notice{pointer-events:auto;min-height:48px;display:grid;grid-template-columns:34px minmax(0,1fr) 24px;gap:9px;align-items:center;padding:8px 9px 8px 10px;border:1px solid rgba(148,163,184,.3);border-radius:11px;background:rgba(8,47,73,.97);box-shadow:0 10px 28px rgba(15,23,42,.24);color:#fff;backdrop-filter:blur(14px);animation:rgtNoticeRight .26s cubic-bezier(.2,.8,.2,1)}`;

let matched = false;
if (code.includes(topBarStack) && code.includes(topBarNotice)) {
  code = code.replace(topBarStack, smallStack).replace(topBarNotice, smallNotice);
  matched = true;
} else if (code.includes(currentSmallLeftStack) && code.includes(smallNotice.replace('rgtNoticeRight', 'rgtNoticeLeft'))) {
  code = code.replace(currentSmallLeftStack, smallStack).replace(smallNotice.replace('rgtNoticeRight', 'rgtNoticeLeft'), smallNotice);
  matched = true;
} else if (code.includes(originalStack) && code.includes(originalNotice)) {
  code = code.replace(originalStack, smallStack).replace(originalNotice, smallNotice);
  matched = true;
}

if (!matched) {
  throw new Error('Expected notification CSS was not found. No source file was changed.');
}

code = code
  .replace(/\.rgt-notice-icon\{width:40px;height:40px;/, `.rgt-notice-icon{width:34px;height:34px;`)
  .replace(/\.rgt-notice strong\{display:block;font-size:14px;/, `.rgt-notice strong{display:block;font-size:12px;`)
  .replace(/\.rgt-notice p\{margin:4px 0 0;color:#cbd5e1;font-size:12px;/, `.rgt-notice p{margin:2px 0 0;color:#cbd5e1;font-size:10px;`)
  .replace(/\.rgt-notice-close\{border:0;background:transparent;color:#94a3b8;font-size:22px;/, `.rgt-notice-close{border:0;background:transparent;color:#94a3b8;font-size:18px;`)
  .replace(/@keyframes rgtNoticeDown\{[^}]+\}to\{[^}]+\}\}/, `@keyframes rgtNoticeRight{from{opacity:0;transform:translate3d(18px,0,0)}to{opacity:1;transform:none}}`)
  .replace(/@keyframes rgtNoticeIn\{[^}]+\}to\{[^}]+\}\}/, `@keyframes rgtNoticeRight{from{opacity:0;transform:translate3d(18px,0,0)}to{opacity:1;transform:none}}`)
  .replace(/@media\(max-width:640px\)\{\.rgt-notice\{min-height:58px;padding:10px 14px;grid-template-columns:38px minmax\(0,1fr\) 28px\}\.rgt-notice-icon\{width:36px;height:36px\}\.rgt-notice p\{white-space:nowrap;overflow:hidden;text-overflow:ellipsis\}\}/, `@media(max-width:640px){.rgt-notice-stack{top:8px;right:8px;width:min(300px,calc(100vw - 16px))}.rgt-notice{min-height:46px;padding:7px 8px;grid-template-columns:32px minmax(0,1fr) 22px}.rgt-notice-icon{width:32px;height:32px}.rgt-notice p{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}`)
  .replace(/@media\(max-width:640px\)\{\.rgt-notice-stack\{top:14px;right:16px\}\}/, `@media(max-width:640px){.rgt-notice-stack{top:8px;right:8px;width:min(300px,calc(100vw - 16px))}.rgt-notice p{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}`);

if (!code.includes('@keyframes rgtNoticeRight')) {
  code = code.replace('@keyframes rgtSpin{to{transform:rotate(360deg)}}', '@keyframes rgtSpin{to{transform:rotate(360deg)}}@keyframes rgtNoticeRight{from{opacity:0;transform:translate3d(18px,0,0)}to{opacity:1;transform:none}}');
}

const temp = `${target}.tmp`;
fs.writeFileSync(temp, code);
fs.renameSync(temp, target);
console.log(`Updated successfully: ${target}`);
console.log('Notification is now small and anchored at the top-right of the page header.');
console.log(`Original bytes: ${Buffer.byteLength(original)} | Updated bytes: ${Buffer.byteLength(code)}`);
