// レイアウト確認用: 指定 md を A4幅でレンダリングしPNG化（cover / toc / body 先頭）
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import MarkdownIt from 'markdown-it';
import puppeteer from 'puppeteer-core';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, 'src');
const CSS = fs.readFileSync(path.join(__dirname, 'style.css'), 'utf8');
const CHROME = ['C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe','C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'].find(p=>fs.existsSync(p));
const md = new MarkdownIt({ html: true, linkify: false });

const file = process.argv[2];
let text = fs.readFileSync(path.join(SRC, file), 'utf8');
const mm = text.match(/<!--META([\s\S]*?)-->/); const meta={};
if(mm) for(const l of mm[1].trim().split(/\r?\n/)){const i=l.indexOf(':');if(i>0)meta[l.slice(0,i).trim()]=l.slice(i+1).trim();}
const body = text.replace(/<!--META[\s\S]*?-->/,'').trim();
const esc=s=>(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

const tokens = md.parse(body,{}); let maj=0,min=0,idx=0; const toc=[];
for(let i=0;i<tokens.length;i++){const t=tokens[i];if(t.type==='heading_open'&&(t.tag==='h2'||t.tag==='h3')){const inl=tokens[i+1];let num;if(t.tag==='h2'){maj++;min=0;num=''+maj;}else{min++;num=maj+'.'+min;}const id='sec-'+(idx++);t.attrSet('id',id);const c=inl.children&&inl.children[0];if(c&&c.type==='text')c.content=num+'　'+c.content;inl.content=num+'　'+inl.content;toc.push({level:t.tag==='h2'?2:3,text:inl.content,id});}}
const html=md.renderer.render(tokens,md.options,{});
const tocHtml=`<nav class="toc"><div class="toc-title">目次</div><ul>`+toc.map(e=>`<li class="lvl${e.level}"><a href="#${e.id}">${esc(e.text)}</a></li>`).join('')+`</ul></nav>`;
const cover=`<section class="cover"><div class="doc-kind">SPECIFICATION</div><div class="title">${esc(meta.title)}<span class="sub">${esc(meta.subtitle)}</span></div><div class="rule"></div><table class="meta"><tbody><tr><th>文書種別</th><td>${esc(meta.subtitle)}</td></tr><tr><th>版数</th><td>${esc(meta.version)}</td></tr><tr><th>発行日</th><td>${esc(meta.date)}</td></tr><tr><th>提出先</th><td>${esc(meta.client||'そばに 御中')}</td></tr></tbody></table><div class="footer-org">本書は内製運用への引き継ぎを目的とした概要仕様書です。</div></section>`;
const full=`<!doctype html><html lang="ja"><head><meta charset="utf-8"><style>${CSS}\nhtml,body{background:#fff}body{width:210mm}.cover,.toc{height:auto;min-height:auto;padding:14mm 16mm;border-bottom:2px dashed #f00}.doc{padding:14mm 16mm}</style></head><body>${cover}${tocHtml}<main class="doc">${html}</main></body></html>`;

const browser=await puppeteer.launch({executablePath:CHROME,headless:'new'});
const page=await browser.newPage();
await page.setViewport({width:794,height:1123,deviceScaleFactor:1.5});
await page.setContent(full,{waitUntil:'networkidle0'});
await page.screenshot({path:path.join(__dirname,'_preview.png'),fullPage:true});
const coverEl=await page.$('.cover'); if(coverEl) await coverEl.screenshot({path:path.join(__dirname,'_cover.png')});
await browser.close();
console.log('preview.png written');
