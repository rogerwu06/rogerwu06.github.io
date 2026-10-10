const PYODIDE_INDEX = "https://cdn.jsdelivr.net/pyodide/v314.0.7/full/";
let pyodidePromise = null;
let scientificReady = false;
let seismicWithPad = false;
const $ = id => document.getElementById(id);
const $$ = s => [...document.querySelectorAll(s)];

function status(id, message, isError=false){ const el=$(id); if(!el) return; el.textContent=message; el.classList.toggle('error',isError); }
function rowsToTsv(rows){ return rows.map(row=>row.map(v=>String(v)).join('\t')).join('\n'); }
async function fetchToFS(py,url,fsName){ const r=await fetch(url); if(!r.ok) throw new Error(`Could not load ${url}`); py.FS.writeFile(`/home/pyodide/site/${fsName}`,await r.text()); }
async function getPyodideRuntime(){
  if(pyodidePromise) return pyodidePromise;
  pyodidePromise=(async()=>{
    const g=$('global-runtime-status'); if(g) g.textContent='Loading Python runtime…';
    const py=await loadPyodide({indexURL:PYODIDE_INDEX});
    py.FS.mkdirTree('/home/pyodide/site');
    await Promise.all([
      fetchToFS(py,'python/seismic_core.py','seismic_core.py'),
      fetchToFS(py,'python/wind_core.py','wind_core.py'),
      fetchToFS(py,'python/pressure_graph.py','pressure_graph.py')
    ]);
    py.runPython(`
import sys
if "/home/pyodide/site" not in sys.path: sys.path.insert(0,"/home/pyodide/site")
import seismic_core, wind_core
`);
    if(g) g.textContent='Python runtime ready — calculations run locally in your browser.';
    return py;
  })().catch(err=>{ pyodidePromise=null; const g=$('global-runtime-status'); if(g) g.textContent=`Python runtime error: ${err.message}`; throw err; });
  return pyodidePromise;
}
async function calculateSeismic(){ try{ status('seismic-status','Running seismic calculation in Python…'); const py=await getPyodideRuntime(); py.globals.set('web_node_text',$('seis-nodes').value); py.globals.set('web_reaction_text',$('seis-reactions').value); py.globals.set('web_sa02',Number($('seis-sa02').value)); py.globals.set('web_ie',Number($('seis-ie').value)); py.globals.set('web_rd',Number($('seis-rd').value)); py.globals.set('web_ro',Number($('seis-ro').value)); py.globals.set('web_height',Number($('seis-height').value)); py.globals.set('web_with_pad',seismicWithPad); const json=py.runPython(`
import json,seismic_core
nodes=seismic_core._parse_clipboard_table(web_node_text,4)
reactions=seismic_core._parse_clipboard_table(web_reaction_text,8)
r=seismic_core.calculate_seismic(nodes,reactions,sa02=web_sa02,ie=web_ie,rd=web_rd,ro=web_ro,height_adjustment=web_height,with_pad=web_with_pad)
json.dumps(r)
`); const r=JSON.parse(json); $('m-coeff').textContent=r.coefficient.toFixed(9); $('m-total').textContent=`${r.total_vertical.toFixed(3)} kN`; $('m-excluded').textContent=`${r.excluded_weight.toFixed(3)} kN`; $('m-effective').textContent=`${r.effective_weight.toFixed(3)} kN`; $('m-vmax').textContent=`${r.vmax.toFixed(3)} kN`; $('m-count').textContent=String(r.count); $('seis-fz').textContent=r.fz_lines.join('\n'); $('seis-fx').textContent=r.fx_lines.join('\n'); status('seismic-status',`Calculation complete · Σ(Wi·hi) = ${r.wihi_sum.toFixed(3)} kN·m · ${r.count} matched reaction nodes.`); }catch(e){status('seismic-status',e.message,true)} }
async function calculateWind(){ try{ status('wind-status','Running wind calculation in Python…'); const py=await getPyodideRuntime(); py.globals.set('web_w_nodes',$('wind-nodes').value); py.globals.set('web_w_members',$('wind-members').value); py.globals.set('web_w_sections',$('wind-sections').value); py.globals.set('web_w_height',Number($('wind-height').value)); py.globals.set('web_w_q',Number($('wind-q').value)); py.globals.set('web_w_cg',Number($('wind-cg').value)); py.globals.set('web_w_compat',$('wind-compat').checked); const json=py.runPython(`
import json,wind_core
nodes=wind_core._parse_clipboard_table(web_w_nodes,4)
members=wind_core._parse_clipboard_table(web_w_members,7)
sections=wind_core._parse_clipboard_table(web_w_sections,10)
max_height,gz,gx=wind_core.calculate_wind_outputs(nodes,members,sections,web_w_height,web_w_q,web_w_cg,exact_workbook_compatibility=web_w_compat)
json.dumps({"max_height":max_height,"gz":gz,"gx":gx})
`); const r=JSON.parse(json); $('wind-maxheight').textContent=`${r.max_height.toFixed(3)} m`; $('wind-gz').textContent=r.gz.join('\n'); $('wind-gx').textContent=r.gx.join('\n'); status('wind-status',`Calculation complete · ${r.gz.length} GZ groups · ${r.gx.length} GX groups.`); }catch(e){status('wind-status',e.message,true)} }
function resetPressureSample(){ $('pressure-design').value='9.5'; $('pressure-reinforced').value=`6\t0\n7\t0.002209\n8\t0.05521\n9\t0.1156\n10\t0.186\n11\t0.2731\n12\t0.3772\n13\t0.5281\n14\t0.7744\n15\t1.174\n16\t1.979\n17\t4.41\n17.529192\t20.06`; $('pressure-unreinforced').value=`6\t0.3915\n7\t0.6452\n8\t1.087\n9\t1.667\n10\t2.39\n11\t3.355\n12\t4.682\n13\t6.607\n14\t9.842\n15\t18.29\n16\t20.22`; status('pressure-status','Sample data restored.'); }
async function calculatePressureGraph(){ try{ status('pressure-status',scientificReady?'Generating graph in Python…':'Loading NumPy, SciPy, and Matplotlib for the first run…'); const py=await getPyodideRuntime(); if(!scientificReady){ await py.loadPackage(['numpy','scipy','matplotlib']); py.runPython('import pressure_graph'); scientificReady=true; } py.globals.set('web_p_reinforced',$('pressure-reinforced').value); py.globals.set('web_p_unreinforced',$('pressure-unreinforced').value); py.globals.set('web_p_design',Number($('pressure-design').value)); const b64=await py.runPythonAsync(`
import pressure_graph
pressure_graph.render_graph(web_p_reinforced,web_p_unreinforced,web_p_design)
`); $('pressure-image').src=`data:image/png;base64,${b64}`; $('pressure-image').hidden=false; $('graph-placeholder').hidden=true; status('pressure-status','Graph generated with Python in the browser.'); }catch(e){status('pressure-status',e.message,true)} }
function showProject(name){ $$('.project-panel').forEach(p=>p.classList.toggle('active',p.id===`project-${name}`)); $$('.project-tab').forEach(b=>b.classList.toggle('active',b.dataset.project===name)); history.replaceState(null,'',`#${name}`); }
async function copyTarget(id,button){ await navigator.clipboard.writeText($(id).textContent); const old=button.textContent; button.textContent='Copied'; setTimeout(()=>button.textContent=old,1000); }
document.addEventListener('DOMContentLoaded',()=>{
  $$('.project-tab').forEach(b=>b.addEventListener('click',()=>showProject(b.dataset.project)));
  $$('[data-seismic-mode]').forEach(b=>b.addEventListener('click',()=>{ seismicWithPad=b.dataset.seismicMode==='true'; $$('[data-seismic-mode]').forEach(x=>x.classList.toggle('active',x===b)); status('seismic-status',`Mode changed to ${seismicWithPad?'WITH PAD':'NO PAD'}. Paste the corresponding STAAD data before calculating.`); }));
$('seismic-calc')?.addEventListener('click',calculateSeismic); $('wind-calc')?.addEventListener('click',calculateWind); $('pressure-sample')?.addEventListener('click',resetPressureSample); $('pressure-calc')?.addEventListener('click',calculatePressureGraph);
  $$('[data-copy-target]').forEach(b=>b.addEventListener('click',()=>copyTarget(b.dataset.copyTarget,b)));
  const initial=location.hash.slice(1); if(initial && document.getElementById(`project-${initial}`)) showProject(initial);
});
