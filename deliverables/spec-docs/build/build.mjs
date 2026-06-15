// Sobani 仕様書ビルダー: Markdown -> 体裁付きPDF（表紙・目次・章番号・ヘッダ/フッタ）
// 使い方:
//   node build.mjs              … src/ の全 .md をビルド
//   node build.mjs 00_xxx.md    … 指定ファイルのみビルド
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import MarkdownIt from 'markdown-it';
import puppeteer from 'puppeteer-core';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, 'src');
const OUT = path.join(ROOT, 'pdf');
const CSS = fs.readFileSync(path.join(__dirname, 'style.css'), 'utf8');

const CHROME_CANDIDATES = [
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
];
const CHROME = CHROME_CANDIDATES.find((p) => fs.existsSync(p));
if (!CHROME) throw new Error('Chrome/Edge が見つかりません');

const md = new MarkdownIt({ html: true, linkify: false, typographer: false });

function parseMeta(text) {
  const meta = {};
  const m = text.match(/<!--META([\s\S]*?)-->/);
  if (m) {
    for (const line of m[1].trim().split(/\r?\n/)) {
      const i = line.indexOf(':');
      if (i > 0) meta[line.slice(0, i).trim()] = line.slice(i + 1).trim();
    }
  }
  const body = text.replace(/<!--META[\s\S]*?-->/, '').trim();
  return { meta, body };
}

function esc(s = '') {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// 章番号を h2/h3 に自動採番し、目次を生成
function renderBody(body) {
  const tokens = md.parse(body, {});
  const toc = [];
  let maj = 0, min = 0, idx = 0;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type === 'heading_open' && (t.tag === 'h2' || t.tag === 'h3')) {
      const inline = tokens[i + 1];
      let num;
      if (t.tag === 'h2') { maj++; min = 0; num = String(maj); }
      else { min++; num = `${maj}.${min}`; }
      const id = `sec-${idx++}`;
      t.attrSet('id', id);
      const child = inline.children && inline.children[0];
      if (child && child.type === 'text') child.content = `${num}　${child.content}`;
      inline.content = `${num}　${inline.content}`;
      toc.push({ level: t.tag === 'h2' ? 2 : 3, text: inline.content, id });
    }
  }
  const html = md.renderer.render(tokens, md.options, {});
  const tocHtml =
    `<nav class="toc"><div class="toc-title">目次</div><ul>` +
    toc.map((e) => `<li class="lvl${e.level}"><a href="#${e.id}">${esc(e.text)}</a></li>`).join('') +
    `</ul></nav>`;
  return { html, tocHtml };
}

function coverHtml(meta) {
  const rows = [
    ['文書種別', meta.subtitle || '仕様書'],
    ['版数', meta.version || '1.0'],
    ['発行日', meta.date || ''],
    ['提出先', meta.client || 'そばに 御中'],
    ['作成', meta.author || ''],
  ].filter(([, v]) => v);
  return (
    `<section class="cover">` +
    `<div class="doc-kind">SPECIFICATION</div>` +
    `<div class="title">${esc(meta.title || '')}` +
    (meta.subtitle ? `<span class="sub">${esc(meta.subtitle)}</span>` : '') +
    `</div>` +
    `<div class="rule"></div>` +
    `<table class="meta"><tbody>` +
    rows.map(([k, v]) => `<tr><th>${esc(k)}</th><td>${esc(v)}</td></tr>`).join('') +
    `</tbody></table>` +
    `<div class="footer-org">本書は内製運用への引き継ぎを目的とした概要仕様書です。</div>` +
    `</section>`
  );
}

function fullHtml(meta, tocHtml, bodyHtml) {
  return (
    `<!doctype html><html lang="ja"><head><meta charset="utf-8"><style>${CSS}</style></head>` +
    `<body>${coverHtml(meta)}${tocHtml}<main class="doc">${bodyHtml}</main></body></html>`
  );
}

async function buildOne(browser, file) {
  const raw = fs.readFileSync(path.join(SRC, file), 'utf8');
  const { meta, body } = parseMeta(raw);
  const { html, tocHtml } = renderBody(body);
  const page = await browser.newPage();
  await page.setContent(fullHtml(meta, tocHtml, html), { waitUntil: 'networkidle0' });
  const header = `<div style="font-size:7.5pt;color:#9aa0a6;width:100%;padding:0 16mm;text-align:right;">${esc(meta.title || '')}</div>`;
  const footer =
    `<div style="font-size:7.5pt;color:#9aa0a6;width:100%;padding:0 16mm;display:flex;justify-content:space-between;align-items:center;">` +
    `<span>${esc(meta.client || 'そばに 御中')}</span>` +
    `<span><span class="pageNumber"></span> / <span class="totalPages"></span></span>` +
    `<span>${esc(meta.subtitle || '仕様書')} v${esc(meta.version || '1.0')}</span></div>`;
  const out = path.join(OUT, file.replace(/\.md$/, '.pdf'));
  await page.pdf({
    path: out,
    format: 'A4',
    printBackground: true,
    displayHeaderFooter: true,
    headerTemplate: header,
    footerTemplate: footer,
    margin: { top: '16mm', bottom: '16mm', left: '16mm', right: '16mm' },
    tagged: true,
    outline: true,
  });
  await page.close();
  console.log('✓', path.basename(out));
}

const args = process.argv.slice(2);
const files = (args.length ? args : fs.readdirSync(SRC).filter((f) => f.endsWith('.md'))).sort();
fs.mkdirSync(OUT, { recursive: true });
const browser = await puppeteer.launch({ executablePath: CHROME, headless: 'new', args: ['--no-sandbox'] });
for (const f of files) await buildOne(browser, f);
await browser.close();
console.log(`\n完了: ${files.length} 件 -> ${OUT}`);
