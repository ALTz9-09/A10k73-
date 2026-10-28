/* Giữ nguyên hình khối, vị trí bàn, khổ giấy của template.docx. */
(function(root){
  const escapeXml=value=>String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
  root.makeClassDocx=async function({teacher,className,year,names},outputType='blob'){
    const zip=await root.JSZip.loadAsync(root.CLASS_TEMPLATE_BASE64,{base64:true});
    let xml=await zip.file('word/document.xml').async('string');
    const values={TITLE:`SƠ ĐỒ LỚP ${className||'…'} NĂM HỌC: ${year||'…'}`,TEACHER:teacher?`GVCN: ${teacher}`:''};
    for(let i=0;i<48;i++)values['SEAT_'+i]=(names[i]||'').trim();
    xml=xml.replace(/\{\{(TITLE|TEACHER|SEAT_\d+)\}\}/g,(_,key)=>escapeXml(values[key]));
    zip.file('word/document.xml',xml);
    return zip.generateAsync({type:outputType,compression:'DEFLATE',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  };
})(typeof window!=='undefined'?window:globalThis);
