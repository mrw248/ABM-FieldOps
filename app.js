
let account=null, msalApp=null, accessToken=null, driveFolderPath="/MOP_SOP";
let recognition=null; let transcriptEl, jsonEl, previewEl, qualityEl, auditEl;
let approverList=[]; let auditLog=[]; let currentDoc=null;

const SENTENCE_SPLIT=/(?<=[.!?])\s+(?=[A-Z(])/;
const ENUM_CUES=/\b(step\s*\d+|^\d+[.)-]|first|second|third|fourth|fifth|next|then|after that|finally)\b[ :\-.,]*/i;
const STOP_WORD_PREFIX=/^(okay|ok|today|so|now|um|uh|alright|basically|like|you know|let's|i will|we will|we are going to|i'm going to|i'm|this is)\b[\s,]*/i;
const MAX_STEP_LEN=160;

function qs(id){ return document.getElementById(id); }
function escapeHtml(s){ return (s||"").replace(/[&<>\"']/g, m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])); }
function setStatus(s, who){ const el=document.getElementById('status'); if(el) el.textContent=s; const w=document.getElementById('who'); if(w) w.textContent=who?(" — "+who):""; }
function showCloudRow(show){ const el=document.getElementById('cloudRow'); if(el) el.style.display = show ? "flex" : "none"; }

function titleFromTranscript(t){
  const first=(t.split(SENTENCE_SPLIT)[0]||"").replace(STOP_WORD_PREFIX,"").trim();
  return first.replace(/^today\s+we('|’)ll\s+/i,"").replace(/^today\s+i('|’)ll\s+explain\s+the\s+procedure\s+to\s+/i,"").replace(/^let's\s+/i,"").replace(/\.$/, "") || "Procedure";
}
function normalizeStep(s){ let t=s.trim(); t=t.replace(STOP_WORD_PREFIX,"").trim().replace(ENUM_CUES,"").trim();
  t=t.replace(/^(to\s+)/i,""); t=t.replace(/^(make sure to|ensure|confirm|verify that)\b/i,"Verify"); t=t.replace(/^(please|you (need|should) to|we (need|should) to)\b/i,"");
  return t.charAt(0).toUpperCase()+t.slice(1);
}
function splitCandidates(text){
  const lines=text.split(/[\r\n]+/).map(l=>l.trim()).filter(Boolean);
  let parts=[];
  if(lines.length>1){ for(const line of lines){ if(ENUM_CUES.test(line)||/^[\-\•]/.test(line)) parts.push(line); else parts=parts.concat(line.split(SENTENCE_SPLIT)); } }
  else { const injected=text.replace(/\b(Step\s*\d+\.?|First|Second|Third|Fourth|Fifth|Next|Then|After that|Finally)\b/gi,"\n$1 ");
         for(const seg of injected.split(/\n+/)){ const sents=seg.split(SENTENCE_SPLIT); for(const s of sents){ if(s.trim()) parts.push(s.trim()); } } }
  return parts.map(p=>p.replace(/^[\-\•\d.)\s]+/,"").trim()).filter(p=>p.length>2);
}
function enforceBullets(parts){
  const out=[];
  for(let st of parts){ if(st.length<=MAX_STEP_LEN){ out.push(st); continue; }
    const chunks=st.split(/;|\band\b/gi).map(x=>x.trim()).filter(Boolean);
    if(chunks.length>1){ for(const c of chunks) out.push(c); } else out.push(st);
  } return out;
}
function dedupe(arr){ const seen=new Set(); const out=[]; for(const s of arr){ const k=s.toLowerCase(); if(!seen.has(k)){seen.add(k); out.push(s);} } return out; }

function applyDomainHints(doc){
  const t=(doc.title+" "+doc.scope).toLowerCase();
  if(/\bpatch|windows|server|reboot|vm snapshot|services?\b/.test(t)){
    doc.prerequisites = doc.prerequisites.concat(["Maintenance window approved","Patch bundle downloaded and checksum verified","VM snapshot available"]).filter(Boolean);
    doc.risks = doc.risks.concat(["Service downtime during reboot","Patch incompatibility or rollback failure"]).filter(Boolean);
    doc.validation = doc.validation.concat(["Services start without errors","Application smoke test passed with owner"]).filter(Boolean);
    doc.rollback = doc.rollback.concat(["Revert to previous VM snapshot","Restore configuration backups"]).filter(Boolean);
  }
  if(/\bswitch|cisco|stp|vlan|port-?channel\b/.test(t)){
    doc.risks = doc.risks.concat(["Loss of connectivity during cutover","STP loop or VLAN mismatch"]).filter(Boolean);
    doc.validation = doc.validation.concat(["Uplinks up with correct STP roles","Endpoints reachable; VLANs correct"]).filter(Boolean);
    doc.rollback = doc.rollback.concat(["Reconnect old switch; restore config"]).filter(Boolean);
  }
  return doc;
}

function assessQuality(doc){
  const issues=[]; if(!doc.steps?.length || doc.steps.length<3) issues.push("Too few steps.");
  (doc.steps||[]).forEach(s=>{ if(!/^[A-Z]/.test(s.instruction)) issues.push(`Step ${s.number} not capitalized.`); if(s.instruction.length>MAX_STEP_LEN) issues.push(`Step ${s.number} long.`); });
  if(!doc.rollback?.length) issues.push("No rollback steps."); if(!doc.validation?.length) issues.push("No validation steps.");
  return issues;
}

function toPreviewHTML(doc){
  const ul = arr => `<ul>${(arr||[]).map(x=>`<li>${escapeHtml(x)}</li>`).join('')}</ul>`;
  const steps = (doc.steps||[]).map(s=>`<li>${escapeHtml(s.instruction||s)}</li>`).join('');
  return `
    <h4>${(doc.type||"Procedure")}: ${escapeHtml(doc.title||"")}</h4>
    <div><b>Purpose:</b> ${escapeHtml(doc.purpose||"")}</div>
    <div><b>Scope:</b> ${escapeHtml(doc.scope||"")}</div>
    <h4>Prerequisites</h4>${ul(doc.prerequisites)}
    <h4>Risks</h4>${ul(doc.risks)}
    <h4>Materials</h4>${ul(doc.materials)}
    <h4>Procedure</h4><ol>${steps}</ol>
    <h4>Rollback</h4>${ul(doc.rollback)}
    <h4>Validation</h4>${ul(doc.validation)}
    <h4>Approvals</h4>${(doc.approvals||[]).map(a=>`${escapeHtml(a.role||"Approver")} ${a.name?("("+escapeHtml(a.name)+")"):""} ${a.status?(" - "+escapeHtml(a.status)):""}`).join("<br/>")}
  `;
}

function structureLocal(){
  const type=qs('docType').value;
  const text=(qs('transcript').value||"").replace(/\[[^\]]+\]/g," ").replace(/\s+/g," ").trim();
  const titleGuess = titleFromTranscript(text);
  let parts = splitCandidates(text).map(normalizeStep).filter(Boolean);
  if(qs('enforceBullets').checked) parts = enforceBullets(parts);
  parts = dedupe(parts);

  const steps = parts.map((s,i)=>({number:i+1,actor:"",instruction:s,expected:"Completed",notes:""}));
  let doc = { type, title:titleGuess, purpose:`${type==="MOP"?"Execute":"Perform"} ${titleGuess.toLowerCase()}.`, scope:"As described in the recording.", prerequisites:[], risks:[], materials:[], steps, rollback:["Revert to previous known-good state"], validation:["Verify expected outcomes with no errors"], approvals:[{role:"Requestor"},{role:"Reviewer"},{role:"Approver"}], versionHistory:[{version:1, notes:"initial", date:new Date().toISOString().slice(0,10)}] };
  doc = applyDomainHints(doc);

  currentDoc = doc;
  var __o=qs('jsonOut'); if(__o){ __o.value = JSON.stringify(doc, null, 2); }
  previewEl.innerHTML = toPreviewHTML(doc);
  const issues=assessQuality(doc);
  qualityEl.textContent = issues.length? ("Quality: ⚠ " + issues.join(" ")) : "Quality: ✅ Solid.";
  qs('fileName').value = `${doc.type}-${doc.title}`.replace(/[^\w\-]+/g,"_");
}

function buildCopilotPrompt(){
  const t = qs('transcript').value || "";
  const kind = qs('docType').value || "MOP";
  return `You are a strict MOP/SOP formatter. Convert the following transcript into a precise ${kind}.
Return ONLY compact JSON with fields: type, title, purpose, scope, prerequisites[], risks[], materials[], steps[] (strings), rollback[], validation[].
Enforce: imperative, single-action bullets; remove filler; normalize titles; realistic prerequisites/risks.
Transcript:\\n\"\"\"${t}\"\"\"`;
}
function sendToCopilot(){
  const url = "https://copilot.microsoft.com/?q=" + encodeURIComponent(buildCopilotPrompt());
  window.open(url, "_blank");
}
function applyCopilotJSON(){
  let txt = qs('copilotIn').value.trim();
  if(!txt) return alert("Paste Copilot's JSON first.");
  const m = txt.match(/\{[\s\S]*\}/);
  if(m) txt = m[0];
  let obj=null; try{ obj = JSON.parse(txt); }catch(e){ return alert("Could not parse JSON."); }
  if(Array.isArray(obj.steps) && typeof obj.steps[0]==="string"){
    obj.steps = obj.steps.map((s,i)=>({number:i+1, actor:"", instruction:s, expected:"Completed", notes:""}));
  }
  currentDoc = obj;
  var __o=qs('jsonOut'); if(__o){ __o.value = JSON.stringify(currentDoc, null, 2); }
  previewEl.innerHTML = toPreviewHTML(currentDoc);
  const issues=assessQuality(currentDoc);
  qualityEl.textContent = issues.length? ("Quality: ⚠ " + issues.join(" ")) : "Quality: ✅ Solid.";
  qs('fileName').value = `${currentDoc.type||"Procedure"}-${currentDoc.title||"Untitled"}`.replace(/[^\w\-]+/g,"_");
}

function initSpeech(){ const SR=window.SpeechRecognition||window.webkitSpeechRecognition; if(!SR){ qs('srStatus').textContent="Web Speech not available."; qs('btnStart').disabled=true; return; }
  recognition=new SR(); recognition.continuous=qs('chkContinuous').checked; recognition.interimResults=true; recognition.lang=navigator.language||"en-US"; let buffer="";
  recognition.onstart=()=> qs('srStatus').textContent="Listening…";
  recognition.onresult=(ev)=>{ let interim=""; for(let i=ev.resultIndex;i<ev.results.length;i++){ const tr=ev.results[i][0].transcript; if(ev.results[i].isFinal){ buffer+= (qs('chkPunct').checked?smartPunct(tr):tr)+" "; } else { interim+=tr; } } qs('transcript').value=buffer+(interim?` [${interim}]`:""); };
  recognition.onerror=e=> qs('srStatus').textContent="Speech error: "+e.error; recognition.onend=()=> qs('srStatus').textContent="Stopped.";
}
function smartPunct(s){ let t=s.trim(); if(!t) return t; if(!/[.!?]$/.test(t)) t+="."; return t.charAt(0).toUpperCase()+t.slice(1); }
function initSignature(){ const c=qs('sig'); const ctx=c.getContext('2d'); let drawing=false; c.addEventListener('pointerdown',e=>{drawing=true;draw(e)}); c.addEventListener('pointermove',e=>{ if(drawing) draw(e); }); window.addEventListener('pointerup',()=> drawing=false);
  function draw(e){ const r=c.getBoundingClientRect(); const x=e.clientX-r.left, y=e.clientY-r.top; ctx.fillStyle="#fff"; ctx.beginPath(); ctx.arc(x,y,1.4,0,Math.PI*2); ctx.fill(); }
  qs('btnClearSig').onclick=()=> ctx.clearRect(0,0,c.width,c.height);
}

function download(name,mime,data){ const blob=new Blob([data],{type:mime}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1000); }
function docToHTML(doc){ const sig=localStorage.getItem("sigDataUrl33"); return `<!doctype html><meta charset="utf-8"><title>${escapeHtml((doc.type||"Procedure")+": "+(doc.title||"Untitled"))}</title>
  <style>body{font:14px/1.5 Segoe UI,Arial;margin:24px}h1{font-size:20px}h2{font-size:16px;margin:14px 0 6px}ul,ol{margin:6px 0 12px 22px}img.sig{{border:1px solid #aaa;padding:4px;border-radius:6px}}</style>
  <h1>${escapeHtml((doc.type||"Procedure")+": "+(doc.title||"Untitled"))}</h1>
  <div><b>Purpose:</b> ${escapeHtml(doc.purpose||"")}</div>
  <div><b>Scope:</b> ${escapeHtml(doc.scope||"")}</div>
  <h2>Prerequisites</h2><ul>${(doc.prerequisites||[]).map(x=>`<li>${escapeHtml(x)}</li>`).join("")}</ul>
  <h2>Risks</h2><ul>${(doc.risks||[]).map(x=>`<li>${escapeHtml(x)}</li>`).join("")}</ul>
  <h2>Materials</h2><ul>${(doc.materials||[]).map(x=>`<li>${escapeHtml(x)}</li>`).join("")}</ul>
  <h2>Procedure</h2><ol>${(doc.steps||[]).map(s=>`<li>${escapeHtml(s.instruction||s)}</li>`).join("")}</ol>
  <h2>Rollback</h2><ul>${(doc.rollback||[]).map(x=>`<li>${escapeHtml(x)}</li>`).join("")}</ul>
  <h2>Validation</h2><ul>${(doc.validation||[]).map(x=>`<li>${escapeHtml(x)}</li>`).join("")}</ul>
  <h2>Approvals</h2><ul>${(doc.approvals||[]).map(a=>`<li>${escapeHtml(a.role||"Approver")} ${a.name?("("+escapeHtml(a.name)+")"):""} ${a.status?(" - "+escapeHtml(a.status)):""}</li>`).join("")}</ul>
  ${sig?`<h2>Signature</h2><img class="sig" src="${sig}" width="320" height="96"/>`:""}
  <hr><small>Generated ${new Date().toLocaleString()}</small>`; }

function wireUI(){
  transcriptEl=qs('transcript'); jsonEl=qs('jsonOut'); previewEl=qs('preview'); qualityEl=qs('quality'); auditEl=qs('audit');
  initSpeech(); initSignature();

  qs('btnStart').onclick=()=>{ recognition&&recognition.start(); qs('btnStart').disabled=true; qs('btnStop').disabled=false; };
  qs('btnStop').onclick =()=>{ recognition&&recognition.stop(); qs('btnStart').disabled=false; qs('btnStop').disabled=true; };

  if (qs('btnStructure')) { qs('btnStructure').onclick = structureLocal; }
  qs('btnCopilot').onclick=sendToCopilot;
  qs('btnApplyCopilot').onclick=applyCopilotJSON;

  
  if (qs('btnPasteJSON')) { qs('btnPasteJSON').onclick = async () => {
    try { const t = await navigator.clipboard.readText(); const ta = qs('copilotIn'); if(ta){ ta.value = t || ta.value; ta.focus(); } } catch(e){ alert('Clipboard read failed. Please paste manually (Ctrl/Cmd+V).'); }
  }; }
qs('btnAddApprover').onclick=()=>{ const name=qs('approverName').value.trim(), email=qs('approverEmail').value.trim(); if(!name) return alert("Name required"); (currentDoc.approvals=currentDoc.approvals||[]).push({role:"Approver",name,status:"PENDING"}); if(jsonEl){ jsonEl.value=JSON.stringify(currentDoc,null,2); } previewEl.innerHTML=toPreviewHTML(currentDoc); };
  qs('btnSign').onclick=()=>{ const c=qs('sig'); const data=c.toDataURL("image/png"); localStorage.setItem("sigDataUrl33", data); if(currentDoc){ const first=(currentDoc.approvals||[]).find(a=>a.status==="PENDING"); if(first) first.status="APPROVED"; if(jsonEl){ jsonEl.value=JSON.stringify(currentDoc,null,2); } previewEl.innerHTML=toPreviewHTML(currentDoc); } };

  qs('btnDownloadHTML').onclick=()=>{ if(!currentDoc){ try{ var __tmp=qs('jsonOut'); if(__tmp){ currentDoc = JSON.parse(__tmp.value); } } catch{} } if(!currentDoc) return alert("Structure or paste Copilot JSON first."); download((qs('fileName').value||"procedure")+".html","text/html",docToHTML(currentDoc)); };
  qs('btnDownloadJSON').onclick=()=>{ if(!currentDoc){ try{ var __tmp=qs('jsonOut'); if(__tmp){ currentDoc = JSON.parse(__tmp.value); } } catch{} } if(!currentDoc) return alert("Structure or paste Copilot JSON first."); download((qs('fileName').value||"procedure")+".json","application/json",JSON.stringify(currentDoc,null,2)); };
  qs('btnPrint').onclick=()=>{ if(!currentDoc){ try{ var __tmp=qs('jsonOut'); if(__tmp){ currentDoc = JSON.parse(__tmp.value); } } catch{} } if(!currentDoc) return alert("Structure or paste Copilot JSON first."); const w=window.open("", "_blank"); w.document.write(docToHTML(currentDoc)+`<script>window.onload=()=>window.print()</`+`script>`); w.document.close(); };
}

window.addEventListener('DOMContentLoaded', wireUI);
