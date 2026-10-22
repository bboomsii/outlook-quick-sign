/* Quick Sign - entirely client-side PDF stamping + Outlook reply draft.
 * Requires Office.js, pdf-lib 1.17.1, PDF.js 3.11.174. */
'use strict';
(() => {
  const $ = id => document.getElementById(id);
  const state = {
    outlook: false, signature: null, signatureRatio: 3.0,
    pdfBytes: null, pdfName: '', pdfDoc: null, pdfJsDoc: null,
    pageNumber: 1, pageCount: 0, renderCounter: 0,
    signedBytes: null, signedName: '', dragging: null
  };
  let drawing = false, lastPt = null, moved = false;
  const SIGNATURE_KEY = 'outlook-quick-sign:signature-v1';
  const PLACEMENT_KEY = 'outlook-quick-sign:last-placement-v1';
  function setStatus(message, kind='') { $('status').textContent=message; $('status').className='status '+kind; }
  function safeError(e) { return e && e.message ? e.message : String(e); }
  function readyForSign() {
    $('signReply').disabled = !(state.signature && state.pdfBytes && state.pdfJsDoc);
    $('downloadSigned').disabled = !state.signedBytes;
  }
  function loadSavedSignature() {
    try { const data=localStorage.getItem(SIGNATURE_KEY); if(data && data.startsWith('data:image/')) updateSignature(data,false); }
    catch(e) { console.warn('Signature storage blocked',e); }
  }
  function updateSignature(dataUrl, save=true) {
    const img = new Image();
    img.onload=()=>{
      state.signature=dataUrl;
      state.signatureRatio=img.naturalWidth / img.naturalHeight || 3;
      $('savedSig').src=dataUrl;
      $('sigOverlayImage').src=dataUrl;
      $('savedSigWrap').hidden=false;
      $('signatureSaved').hidden=false;
      if(save) {
        try { localStorage.setItem(SIGNATURE_KEY,dataUrl); }
        catch(e) { setStatus('Signature works for this session, but the browser prevented saving it. Try a smaller image.','error'); }
      }
      if (state.pdfJsDoc) positionOverlay(false);
      readyForSign();
    };
    img.onerror=()=>setStatus('Could not read that signature image.','error');
    img.src=dataUrl;
  }
  function trimSignature(canvas) {
    const ctx=canvas.getContext('2d',{willReadFrequently:true});
    const {width:w,height:h}=canvas, pixels=ctx.getImageData(0,0,w,h).data;
    let x0=w,y0=h,x1=-1,y1=-1;
    for(let y=0;y<h;y++) for(let x=0;x<w;x++){
      const i=(y*w+x)*4; if(pixels[i+3]>35){x0=Math.min(x0,x);y0=Math.min(y0,y);x1=Math.max(x1,x);y1=Math.max(y1,y);}
    }
    if(x1<0) throw new Error('Draw your signature before saving.');
    const pad=10; x0=Math.max(0,x0-pad); y0=Math.max(0,y0-pad); x1=Math.min(w-1,x1+pad);y1=Math.min(h-1,y1+pad);
    const dest=document.createElement('canvas'); dest.width=x1-x0+1;dest.height=y1-y0+1;
    dest.getContext('2d').drawImage(canvas,x0,y0,dest.width,dest.height,0,0,dest.width,dest.height);
    return dest.toDataURL('image/png');
  }
  function setupDrawing() {
    const canvas=$('drawCanvas'), ctx=canvas.getContext('2d');
    ctx.strokeStyle='#142b50';ctx.lineWidth=3.8;ctx.lineCap='round';ctx.lineJoin='round';
    const coords=e=>{ const r=canvas.getBoundingClientRect(); return {x:(e.clientX-r.left)*canvas.width/r.width,y:(e.clientY-r.top)*canvas.height/r.height}; };
    canvas.addEventListener('pointerdown',e=>{e.preventDefault();drawing=true;lastPt=coords(e);moved=true;canvas.setPointerCapture(e.pointerId);});
    canvas.addEventListener('pointermove',e=>{if(!drawing)return;e.preventDefault();const p=coords(e);ctx.beginPath();ctx.moveTo(lastPt.x,lastPt.y);ctx.lineTo(p.x,p.y);ctx.stroke();lastPt=p;});
    const end=()=>{drawing=false;lastPt=null;};canvas.addEventListener('pointerup',end);canvas.addEventListener('pointercancel',end);
    $('clearCanvas').onclick=()=>{ctx.clearRect(0,0,canvas.width,canvas.height);moved=false;};
    $('saveSignature').onclick=()=>{try{if(!moved)throw Error('Please sign in the box first.');updateSignature(trimSignature(canvas));setStatus('Signature saved. You only need to do this once.','ok');}catch(e){setStatus(safeError(e),'error');}};
    $('signatureFile').onchange=async e=>{
      try{
        const f=e.target.files?.[0];if(!f)return;
        if(!['image/png','image/jpeg'].includes(f.type))throw Error('Use a PNG, JPG or WebP signature image.');
        if(f.size>3*1024*1024)throw Error('Signature image should be smaller than 3 MB.');
        const data=await fileToDataURL(f);updateSignature(data);setStatus('Signature image saved.','ok');
      }catch(err){setStatus(safeError(err),'error');}finally{e.target.value='';}
    };
    $('forgetSignature').onclick=()=>{
      try {localStorage.removeItem(SIGNATURE_KEY);}catch(e){}
      state.signature=null;$('savedSigWrap').hidden=true;$('signatureSaved').hidden=true;$('signatureOverlay').hidden=true;
      readyForSign();setStatus('Saved signature removed from this browser.','ok');
    };
  }
  function fileToDataURL(file) { return new Promise((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(r.result);r.onerror=()=>reject(Error('Cannot open file.'));r.readAsDataURL(file);}); }
  function b64ToBytes(s) {
    const b64=s.includes(',')?s.slice(s.indexOf(',')+1):s;
    const bin=atob(b64.replace(/\s/g,'')), out=new Uint8Array(bin.length);
    for(let i=0;i<bin.length;i++)out[i]=bin.charCodeAt(i);
    return out;
  }
  function bytesToB64(bytes) {
    let s='';const stride=0x4000;
    for(let i=0;i<bytes.length;i+=stride)s+=String.fromCharCode(...bytes.subarray(i,i+stride));
    return btoa(s);
  }
  function detectPDF(name,bytes) {
    if(!name.toLowerCase().endsWith('.pdf'))throw Error('Choose a PDF file.');
    if(bytes.length<5 || String.fromCharCode(...bytes.subarray(0,5))!=='%PDF-')throw Error('This attachment does not appear to be a valid PDF.');
  }
  async function loadPDF(name, bytes) {
    if(!window.pdfjsLib || !window.PDFLib)throw Error('PDF libraries did not load. Check your Internet connection or content-blocking policy.');
    detectPDF(name,bytes);
    state.pdfBytes=bytes;state.pdfName=name;state.signedBytes=null;state.signedName='';
    const doc=await window.pdfjsLib.getDocument({data:bytes.slice()}).promise;
    state.pdfJsDoc=doc;state.pageCount=doc.numPages;state.pageNumber=doc.numPages;
    $('currentFile').textContent=`${name} • ${doc.numPages} page(s)`;
    $('pdfControls').hidden=false;
    await renderPage();readyForSign();setStatus('PDF loaded. Move the blue signature box onto the signing area.','ok');
  }
  function getPlacement() {
    const c=$('pdfCanvas'),o=$('signatureOverlay');if(!c.clientWidth || !c.clientHeight)return null;
    return {
      x:o.offsetLeft/c.clientWidth,
      y:o.offsetTop/c.clientHeight,
      width:o.offsetWidth/c.clientWidth
    };
  }
  function positionOverlay(keepLast=true) {
    if(!state.signature)return;
    const c=$('pdfCanvas'),o=$('signatureOverlay');
    if(!c.clientWidth||!c.clientHeight)return;
    const w=c.clientWidth, h=c.clientHeight;
    let remembered=null;
    if(keepLast)try{remembered=JSON.parse(localStorage.getItem(PLACEMENT_KEY));}catch(e){}
    const wp=remembered?.width ?? (Number($('sizeSlider').value)/100);
    $('sizeSlider').value=Math.round(Math.max(12,Math.min(65,wp*100)));
    const ow=Math.min(w*wp,w-8),oh=ow/state.signatureRatio;
    let left=remembered?.x != null ? remembered.x*w : w-ow-19;
    let top=remembered?.y != null ? remembered.y*h : h-oh-31;
    left=Math.max(0,Math.min(w-ow,left));top=Math.max(0,Math.min(h-oh,top));
    Object.assign(o.style,{width:`${ow}px`,height:`${oh}px`,left:`${left}px`,top:`${top}px`});
    o.hidden=false;
  }
  async function renderPage() {
    const cycle=++state.renderCounter;
    const p=await state.pdfJsDoc.getPage(state.pageNumber);
    const base=p.getViewport({scale:1});
    const available=Math.max(180,Math.min($('paperWrap').clientWidth-18,900));
    const scale=available/base.width;
    const viewport=p.getViewport({scale});
    const canvas=$('pdfCanvas');canvas.width=Math.round(viewport.width);canvas.height=Math.round(viewport.height);
    const ctx=canvas.getContext('2d');
    await p.render({canvasContext:ctx,viewport}).promise;
    if(cycle!==state.renderCounter)return;
    $('paper').hidden=false;$('paperEmpty').hidden=true;
    $('pageNumber').textContent=String(state.pageNumber);$('pageCount').textContent=String(state.pageCount);
    $('prevPage').disabled=state.pageNumber<=1;$('nextPage').disabled=state.pageNumber>=state.pageCount;
    positionOverlay();
  }
  function moveOverlay(x,y) {
    const c=$('pdfCanvas'),o=$('signatureOverlay');
    o.style.left=`${Math.max(0,Math.min(c.clientWidth-o.offsetWidth,x))}px`;
    o.style.top=`${Math.max(0,Math.min(c.clientHeight-o.offsetHeight,y))}px`;
  }
  function setupOverlay() {
    const o=$('signatureOverlay');
    o.addEventListener('pointerdown',e=>{e.preventDefault();state.dragging={startX:e.clientX,startY:e.clientY,x:o.offsetLeft,y:o.offsetTop};o.setPointerCapture(e.pointerId);});
    o.addEventListener('pointermove',e=>{if(!state.dragging)return;const d=state.dragging;moveOverlay(d.x+e.clientX-d.startX,d.y+e.clientY-d.startY);});
    o.addEventListener('pointerup',()=>state.dragging=null);
    o.addEventListener('pointercancel',()=>state.dragging=null);
    o.addEventListener('keydown',e=>{const delta=e.shiftKey?10:2;const direction={ArrowLeft:[-delta,0],ArrowRight:[delta,0],ArrowUp:[0,-delta],ArrowDown:[0,delta]}[e.key];if(direction){e.preventDefault();moveOverlay(o.offsetLeft+direction[0],o.offsetTop+direction[1]);}});
    $('sizeSlider').addEventListener('input',()=>{
      const c=$('pdfCanvas');if(!state.signature || !c.clientWidth)return;
      const cx=o.offsetLeft+o.offsetWidth/2,cy=o.offsetTop+o.offsetHeight/2;
      const w=Number($('sizeSlider').value)/100*c.clientWidth, h=w/state.signatureRatio;
      o.style.width=`${w}px`;o.style.height=`${h}px`;
      moveOverlay(cx-w/2,cy-h/2);
    });
    $('prevPage').onclick=async()=>{if(state.pageNumber>1){state.pageNumber--;await renderPage();}};
    $('nextPage').onclick=async()=>{if(state.pageNumber<state.pageCount){state.pageNumber++;await renderPage();}};
    let delay=null;window.addEventListener('resize',()=>{if(!state.pdfJsDoc)return;clearTimeout(delay);delay=setTimeout(()=>renderPage().catch(e=>setStatus(safeError(e),'error')),250);});
  }
  function setupOutlook() {
    const item=window.Office?.context?.mailbox?.item;
    const compatible=window.Office?.context?.requirements?.isSetSupported('Mailbox','1.8');
    if(item && compatible && typeof item.getAttachmentContentAsync==='function'){
      state.outlook=true;
      $('connection').textContent='Connected to Outlook — select a PDF attachment below.';
      $('connection').className='notice success';
      const attachments=(item.attachments||[]).filter(a=>a.name?.toLowerCase().endsWith('.pdf'));
      $('attachmentSelect').innerHTML='';
      if(!attachments.length){$('attachmentSelect').add(new Option('No attached PDFs (cloud links are not supported)','')); $('loadFromOutlook').disabled=true;}
      else attachments.forEach((a,i)=>$('attachmentSelect').add(new Option(a.name,String(i))));
      $('loadFromOutlook').onclick=()=>{
        const idx=Number($('attachmentSelect').value);const a=attachments[idx];if(!a)return;
        $('loadFromOutlook').disabled=true;setStatus('Opening PDF attachment…');
        item.getAttachmentContentAsync(a.id,async result=>{
          $('loadFromOutlook').disabled=false;
          try{
            if(result.status!==Office.AsyncResultStatus.Succeeded)throw Error(result.error?.message||'Could not read Outlook attachment.');
            if(result.value.format!==Office.MailboxEnums.AttachmentContentFormat.Base64)throw Error('This is a cloud-linked attachment; attach the PDF directly to the email instead.');
            await loadPDF(a.name,b64ToBytes(result.value.content));
          }catch(e){setStatus(safeError(e),'error');}
        });
      };
      if(attachments.length===1)$('loadFromOutlook').click();
    }else{
      state.outlook=false;
      $('connection').textContent='Standalone test mode. Choose a local PDF; Outlook reply requires the installed Outlook add-in.';
      $('connection').className='notice warning';$('attachmentSelect').disabled=true;$('loadFromOutlook').disabled=true;
    }
  }
  async function buildSignedPDF() {
    if(!state.signature||!state.pdfBytes)throw Error('Load a PDF and save your signature first.');
    if(!window.PDFLib)throw Error('PDF library unavailable.');
    const pdf=await window.PDFLib.PDFDocument.load(state.pdfBytes,{updateMetadata:false});
    const page=pdf.getPage(state.pageNumber-1);
    // Standard orientation only: rotation changes displayed axes.
    if(((page.getRotation().angle%360)+360)%360!==0) throw Error('This PDF page is rotated. Rotate it to normal orientation before signing.');
    const crop=page.getCropBox(), c=$('pdfCanvas'),pos=getPlacement(),o=$('signatureOverlay');
    if(!pos || o.hidden)throw Error('Place the signature first.');
    const w=pos.width*crop.width;
    const h=(o.offsetHeight/c.clientHeight)*crop.height;
    const x=crop.x+pos.x*crop.width;
    const y=crop.y+(1-pos.y-o.offsetHeight/c.clientHeight)*crop.height;
    const sigData=b64ToBytes(state.signature);
    let image;
    if(state.signature.startsWith('data:image/png'))image=await pdf.embedPng(sigData);
    else if(state.signature.startsWith('data:image/jpeg'))image=await pdf.embedJpg(sigData);
    else throw Error('Use PNG or JPEG as your signature.');
    page.drawImage(image,{x,y,width:w,height:h});
    const signed=await pdf.save({useObjectStreams:true});
    const signedName=state.pdfName.replace(/\.pdf$/i,'')+'_signed.pdf';
    state.signedBytes=signed;state.signedName=signedName;
    try{localStorage.setItem(PLACEMENT_KEY,JSON.stringify(pos));}catch(e){}
    readyForSign();
    return {signed,signedName};
  }
  function download(bytes,name){
    const blob=new Blob([bytes],{type:'application/pdf'}),url=URL.createObjectURL(blob);
    const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),2500);
  }
  async function signAndReply() {
    const button=$('signReply');button.disabled=true;setStatus('Adding the signature to the PDF…');
    try{
      const {signed,signedName}=await buildSignedPDF();
      const item=window.Office?.context?.mailbox?.item;
      if(!state.outlook || !item?.displayReplyFormAsync){
        setStatus('Signed PDF prepared. Standalone mode cannot open an Outlook reply; use Download signed PDF or run from the Outlook add-in.','ok');return;
      }
      if(signed.length>25*1024*1024)throw Error('The signed PDF exceeds the 25 MB attachment limit for New Outlook. Download it instead.');
      setStatus('Preparing an Outlook reply with the signed PDF attached…');
      const attachment={type:Office.MailboxEnums.AttachmentType.Base64,name:signedName,base64file:bytesToB64(signed),inLine:false};
      item.displayReplyFormAsync({htmlBody:'<p>Please find the signed document attached.</p>',attachments:[attachment]},result=>{
        if(result.status===Office.AsyncResultStatus.Succeeded)setStatus('Reply opened with the signed PDF attached. Check the attachment and recipient, then press Send.','ok');
        else setStatus('Outlook could not create the reply: '+(result.error?.message||'unknown error')+'. You can download the signed PDF instead.','error');
      });
    }catch(e){setStatus(safeError(e),'error');}
    finally{readyForSign();}
  }
  function start() {
    if(window.pdfjsLib)window.pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    setupDrawing();setupOverlay();loadSavedSignature();
    $('filePicker').onchange=async e=>{try{const f=e.target.files?.[0];if(f)await loadPDF(f.name,new Uint8Array(await f.arrayBuffer()));}catch(err){setStatus(safeError(err),'error');}finally{e.target.value='';}};
    $('signReply').onclick=signAndReply;
    $('downloadSigned').onclick=()=>{if(state.signedBytes)download(state.signedBytes,state.signedName);};
    // Office.onReady works only when embedded in Outlook. Permit browser-only testing.
    if(window.Office && typeof Office.onReady==='function'){
      let done=false;const fallback=setTimeout(()=>{if(!done){done=true;setupOutlook();}},2200);
      Office.onReady(()=>{if(!done){done=true;clearTimeout(fallback);setupOutlook();}});
    }else setupOutlook();
    readyForSign();
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start);else start();
})();
