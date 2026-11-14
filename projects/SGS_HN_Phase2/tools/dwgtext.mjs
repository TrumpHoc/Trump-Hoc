import { Dwg_File_Type, LibreDwg } from '/tmp/claude-0/dwgtool/node_modules/@mlightcad/libredwg-web/dist/libredwg-web.js';
import fs from 'fs';
const lib = await LibreDwg.create('/tmp/claude-0/dwgtool/node_modules/@mlightcad/libredwg-web/wasm/');
const f = process.argv[2];
const buf = fs.readFileSync(f);
const dwg = lib.dwg_read_data(buf, Dwg_File_Type.DWG);
const db = lib.convert(dwg);
const out=[];
const clean=s=>String(s??'').replace(/\\P/g,' | ').replace(/\\[A-Za-z][^;]*;/g,'').replace(/[{}]/g,'').trim();
function walk(ents, where){
  for(const e of ents||[]){
    const p=e.insertionPoint||e.startPoint||e.position||e.textMidPoint||{};
    if(e.type==='TEXT'||e.type==='MTEXT'||e.type==='ATTRIB'||e.type==='ATTDEF') out.push([where,e.type,(p.x??0).toFixed(0),(p.y??0).toFixed(0),clean(e.text||e.contents||e.textString||e.defaultValue||'')]);
    if(e.type==='INSERT' && e.attribs) for(const a of e.attribs) out.push([where,'ATTR:'+e.name,(p.x??0).toFixed(0),(p.y??0).toFixed(0),clean(a.text)]);
  }
}
walk(db.entities,'MS');
const blocks = db.tables?.BLOCK_RECORD?.entries||[];
for(const b of blocks){ if(b.entities && !/^\*Model/i.test(b.name)) walk(b.entities,'BLK:'+b.name); }
console.log(out.map(r=>r.join('\t')).join('\n'));
const cnt={}; for(const e of db.entities||[]) cnt[e.type]=(cnt[e.type]||0)+1; console.error(JSON.stringify(cnt));
