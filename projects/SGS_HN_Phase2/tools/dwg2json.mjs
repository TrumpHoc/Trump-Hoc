import { Dwg_File_Type, LibreDwg } from '/tmp/claude-0/dwgtool/node_modules/@mlightcad/libredwg-web/dist/libredwg-web.js';
import fs from 'fs';
const lib = await LibreDwg.create('/tmp/claude-0/dwgtool/node_modules/@mlightcad/libredwg-web/wasm/');
const db = lib.convert(lib.dwg_read_data(fs.readFileSync(process.argv[2]), Dwg_File_Type.DWG));
const blocks={}; for(const b of db.tables?.BLOCK_RECORD?.entries||[]) blocks[b.name]=b.entities||[];
fs.writeFileSync(process.argv[3], JSON.stringify({ents:db.entities, blocks}, (k,v)=> typeof v==='bigint'? Number(v): v));
