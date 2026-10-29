<!DOCTYPE html>
<html lang="en"><body>
Word Document Merger Upload, select, reorder, and download one merged Word document.


<script>
let docs=[],uid=0,dragId=null;
const $=id=>document.getElementById(id), esc=s=>s.replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
$('drop').onclick=()=>$('file').click();
$('file').onchange=e=>{addFiles(e.target.files);e.target.value=''};
['dragover','dragenter'].forEach(t=>$('drop').addEventListener(t,e=>{e.preventDefault();$('drop').classList.add('over')}));
['dragleave','drop'].forEach(t=>$('drop').addEventListener(t,e=>{e.preventDefault();$('drop').classList.remove('over')}));
$('drop').addEventListener('drop',e=>addFiles(e.dataTransfer.files));

async function addFiles(files){
 for(const f of files){
  const d={id:++uid,name:f.name,file:f,selected:true,warnings:[],error:null};
  docs.push(d); render();
  await inspect(d); render();
 }
}
async function inspect(d){
 const ext=d.name.split('.').pop().toLowerCase();
 if(ext==='doc'){d.error='Old .doc format cannot be merged. Open it in Word and Save As .docx.';d.selected=false;return}
 try{
  d.buf=await d.file.arrayBuffer();
  const z=await JSZip.loadAsync(d.buf);
  const doc=z.file('word/document.xml');
  if(!doc){d.error='Not a valid Word document.';d.selected=false;return}
  const x=await doc.async('string');
  if(ext==='docm'||z.file('word/vbaProject.bin'))d.warnings.push('Contains macros; macros will not be carried into the merged file.');
  if(ext==='dotx')d.warnings.push('This is a template; its content will be merged as a normal document.');
  if(/<w:(ins|del)\b/.test(x))d.warnings.push('Has tracked changes; they may be accepted or flattened when Word combines the files.');
  if(z.file('word/comments.xml'))d.warnings.push('Has comments; they may not carry over.');
  if(/<w:sectPr[\s\S]*?<w:(headerReference|footerReference)/.test(x))d.warnings.push('Has headers/footers; Word may use the first document\'s headers/footers for merged pages.');
  if(!/<w:t[ >]/.test(x)&&!/<w:drawing|<w:pict/.test(x))d.warnings.push('Appears to be empty.');
  if(docs.some(o=>o!==d&&o.name===d.name))d.warnings.push('Duplicate file name — check you meant to add it twice.');
 }catch(e){d.error='Could not read file (it may be password-protected or corrupted).';d.selected=false}
}
function render(){
 const L=$('list');L.innerHTML='';
 docs.forEach((d,i)=>{
  const li=document.createElement('li');li.draggable=true;li.dataset.id=d.id;
  li.innerHTML=`☰${i+1}.
  <input type="checkbox" ${d.selected?'checked':''} ${d.error?'disabled':''}>
  ${esc(d.name)}${d.error?` ✖ ${d.error} `:''}${d.warnings.map(w=>` ⚠ ${w} `).join('')}
  
  
  <input type="number" min="1" max="${docs.length}" value="${i+1}" style="width:52px" title="Move to position">
  `;
  li.querySelector('input[type=checkbox]').onchange=e=>{d.selected...