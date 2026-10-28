'use strict';
// All indices follow the screen: back row first, then left to right.
const fixedDesks=new Set();
let arrangeHistory=[];
function clearArrangeHistory(){arrangeHistory=[];$('undo-arrange').disabled=true;}
function refreshDeskLabels(){
  desks.forEach((desk,d)=>{
    let badge=desk.querySelector('.desk-label');
    if(!badge){badge=document.createElement('span');badge.className='desk-label';desk.append(badge);}
    const number=selectedDesks.indexOf(d);
    badge.textContent=`${fixedDesks.has(d)?'🔒 ':''}Bàn ${number>=0?number+1:d+1}`;
    desk.classList.toggle('fixed-desk',fixedDesks.has(d));
    desk.title=`Dãy ${d%4+1}, hàng ${Math.floor(d/4)+1} từ phía sau${fixedDesks.has(d)?' · Cố định':''}`;
  });
}
function syncCurrentResults(){
  clearArrangeHistory();
  rebuildCurrentResults();
}
function rebuildCurrentResults(){
  // Once manually rearranged, desk numbers refer to the 24 physical desks.
  selectedDesks=Array.from({length:24},(_,i)=>i);
  randomAssignments=seats.flatMap((seat,i)=>seat.value.trim()?[{name:seat.value.trim(),seat:i,deskNumber:Math.floor(i/2)+1,physicalDesk:Math.floor(i/2)}]:[]);
  $('export-random').disabled=!randomAssignments.length;
  refreshDeskLabels();updateCount();
}
function resetArrangementState(){
  fixedDesks.clear();clearArrangeHistory();refreshDeskLabels();
}
$('lock-mode').onchange=()=>{
  if($('lock-mode').checked){$('swap').checked=false;resetSelection();}
  seats.forEach(s=>s.readOnly=$('lock-mode').checked||$('swap').checked);
  $('paper').classList.toggle('locking',$('lock-mode').checked);
};
$('swap').addEventListener('change',()=>{
  if($('swap').checked){$('lock-mode').checked=false;$('paper').classList.remove('locking');}
  seats.forEach(s=>s.readOnly=$('lock-mode').checked||$('swap').checked);
});
desks.forEach((desk,d)=>{
  desk.addEventListener('click',e=>{
    if(!$('lock-mode').checked)return;
    e.preventDefault();e.stopPropagation();
    if(fixedDesks.has(d))fixedDesks.delete(d);else fixedDesks.add(d);
    refreshDeskLabels();
  },true);
  desk.addEventListener('keydown',e=>{
    if($('lock-mode').checked&&(e.key==='Enter'||e.key===' ')){e.preventDefault();desk.click();}
  });
});
seats.forEach(s=>s.addEventListener('input',syncCurrentResults));
$('unlock-all').onclick=()=>{fixedDesks.clear();refreshDeskLabels();};
function rotateArrangement(axis,step){
  const before=seats.map(s=>s.value),after=before.slice();let changed=false;
  for(let line=0;line<(axis==='rows'?4:6);line++){
    const ring=Array.from({length:axis==='rows'?6:4},(_,n)=>axis==='rows'?n*4+line:line*4+n).filter(d=>!fixedDesks.has(d));
    if(ring.length<2)continue;
    ring.forEach((from,i)=>{
      const to=ring[(i+step+ring.length)%ring.length];
      for(let side=0;side<2;side++)after[to*2+side]=before[from*2+side];
    });
  }
  changed=after.some((v,i)=>v!==before[i]);
  if(!changed){$('status').textContent='Không có thay đổi: các bàn còn lại trống, giống nhau hoặc đã cố định.';return;}
  arrangeHistory.push({names:before,assignments:randomAssignments.map(x=>({...x})),selected:selectedDesks.slice()});
  if(arrangeHistory.length>30)arrangeHistory.shift();
  seats.forEach((s,i)=>s.value=after[i]);resetSelection();rebuildCurrentResults();
  $('undo-arrange').disabled=false;
  $('status').textContent=axis==='rows'?`Đã đổi bàn ${step===1?'về phía bảng':'ra phía sau'}, giữ nguyên bàn cố định.`:`Đã đổi dãy sang ${step===1?'phải':'trái'}, giữ nguyên bàn cố định.`;
}
$('rotate-desks').onclick=()=>rotateArrangement('rows',Number($('row-direction').value));
$('rotate-columns').onclick=()=>rotateArrangement('columns',Number($('column-direction').value));
$('undo-arrange').onclick=()=>{
  const prev=arrangeHistory.pop();if(!prev)return;
  seats.forEach((s,i)=>s.value=prev.names[i]);randomAssignments=prev.assignments;selectedDesks=prev.selected;
  resetSelection();updateCount();refreshDeskLabels();$('export-random').disabled=!randomAssignments.length;
  $('undo-arrange').disabled=!arrangeHistory.length;$('status').textContent='Đã hoàn tác lần đổi gần nhất.';
};

const WORD_NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP_NS='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
function descendants(node,ns,name){return [...node.getElementsByTagNameNS(ns,name)];}
function wordText(node){
  return descendants(node,WORD_NS,'p').map(p=>descendants(p,WORD_NS,'t').map(t=>t.textContent).join('')).join(' ').normalize('NFC').trim();
}
function parseClassDocxXml(xml){
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  if(doc.getElementsByTagName('parsererror').length)throw Error('Nội dung Word bị lỗi.');
  const body=doc.getElementsByTagNameNS(WORD_NS,'body')[0];
  if(!body)throw Error('Không tìm thấy nội dung sơ đồ.');
  const paras=[...body.children].filter(p=>p.localName==='p');
  const rows=[];
  for(const p of paras){
    const anchors=descendants(p,WP_NS,'anchor').filter(a=>descendants(a,WORD_NS,'txbxContent').length);
    if(anchors.length!==8)continue;
    const cells=anchors.map(a=>{
      const position=descendants(a,WP_NS,'positionH')[0];
      const off=position&&descendants(position,WP_NS,'posOffset')[0];
      if(!off||!Number.isFinite(Number(off.textContent)))throw Error('Không xác định được vị trí ô trong Word.');
      return {x:Number(off.textContent),name:wordText(descendants(a,WORD_NS,'txbxContent')[0])};
    }).sort((a,b)=>a.x-b.x);
    if(new Set(cells.map(c=>c.x)).size!==8)throw Error('Các ô trong Word bị trùng vị trí.');
    rows.push(cells.map(c=>c.name));
  }
  if(rows.length!==6)throw Error('Chưa hỗ trợ bố cục Word này. Hãy dùng mẫu 4 dãy × 6 bàn đã gửi hoặc file Word xuất từ web.');
  const names=rows.flat();
  if(names.some(n=>n.length>70||/\{\{SEAT_/.test(n)))throw Error('Có ô tên không hợp lệ hoặc dài quá 70 ký tự.');
  const plain=paras.filter(p=>!descendants(p,WP_NS,'anchor').length).map(p=>descendants(p,WORD_NS,'t').map(t=>t.textContent).join('').normalize('NFC').trim());
  const header=plain.find(t=>/SƠ ĐỒ LỚP/i.test(t))||'';
  const match=header.match(/SƠ ĐỒ LỚP\s*(.*?)\s*NĂM HỌC\s*:\s*(.*)/i);
  const teacher=plain.find(t=>/^GVCN\s*:/i.test(t));
  return {names,className:match?.[1]||'',year:match?.[2]||'',teacher:teacher?teacher.replace(/^GVCN\s*:\s*/i,''):''};
}
$('docx-file').onchange=async()=>{
  const input=$('docx-file'),file=input.files[0];if(!file)return;
  input.disabled=true;
  try{
    if(!window.JSZip)throw Error('Thiếu jszip.min.js. Hãy tải đủ các file của web.');
    if(file.size>10*1024*1024)throw Error('Vui lòng chọn file DOCX dưới 10 MB.');
    const zip=await JSZip.loadAsync(await file.arrayBuffer());
    const entry=zip.file('word/document.xml');if(!entry)throw Error('File không phải tài liệu DOCX hợp lệ.');
    const data=parseClassDocxXml(await entry.async('string'));
    seats.forEach((s,i)=>s.value=data.names[i]);
    ['teacher','className','year'].forEach(k=>$(k).value=data[k]);
    importedNames=data.names.filter(Boolean);importedStudents=importedNames.map(n=>({name:n,fixedDesk:null}));resetSelection();resetArrangementState();rebuildCurrentResults();updateHeader();
    $('random').disabled=!importedNames.length;
    $('status').textContent=`Đã nạp sơ đồ Word: ${importedNames.length} học sinh, giữ đúng vị trí và ô trống.`;
  }catch(error){$('status').textContent=`Không nạp được Word: ${error.message}`;}
  finally{input.disabled=false;input.value='';}
};
refreshDeskLabels();
