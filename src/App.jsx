import { useState, useEffect, useMemo, useCallback } from "react";
import { supabase } from "./supabase";

const C = {
  forest:"#3e554f", forestDk:"#2a3b36", forestLt:"#e8efe8", forestMd:"#5a7a72",
  warm:"#f7f5f0", white:"#ffffff", alert:"#c8472a", alertLt:"#fdf0ed",
  gold:"#b8860b", goldLt:"#fdf8e8", text:"#1a2a25", muted:"#6b7c77", border:"#d4ddd9",
};

const GLOBAL_CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Inter', sans-serif; background: ${C.warm}; color: ${C.text}; }
  input, textarea, select { font-family: inherit; }
  button { cursor: pointer; font-family: inherit; }
  ::-webkit-scrollbar { width: 4px; }
  ::-webkit-scrollbar-track { background: ${C.forestLt}; }
  ::-webkit-scrollbar-thumb { background: ${C.forestMd}; border-radius: 2px; }
`;

// ── PDF ──────────────────────────────────────────────────────────────────────
const loadJsPDF = () => new Promise(res => {
  if (window.jspdf) { res(window.jspdf.jsPDF); return; }
  const s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
  s.onload = () => res(window.jspdf.jsPDF);
  document.head.appendChild(s);
});
const loadXLSX = () => new Promise(res => {
  if (window.XLSX) { res(window.XLSX); return; }
  const s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
  s.onload = () => res(window.XLSX);
  document.head.appendChild(s);
});
const hexRGB = h => [parseInt(h.slice(1,3),16),parseInt(h.slice(3,5),16),parseInt(h.slice(5,7),16)];

async function generatePDF(title, subtitle, rows, columns, orgName) {
  const JsPDF = await loadJsPDF();
  const doc = new JsPDF({ orientation:"portrait", unit:"mm", format:"a4" });
  const W=210, pH=297, m=16, cW=W-m*2;
  const addHeader = (n) => {
    doc.setFillColor(...hexRGB(C.forest)); doc.rect(0,0,W,28,"F");
    doc.setFillColor(...hexRGB(C.forestMd)); doc.roundedRect(m,6,36,16,2,2,"F");
    doc.setFont("helvetica","bold"); doc.setFontSize(9); doc.setTextColor(255,255,255);
    doc.text("BidElevate",m+18,14,{align:"center"});
    doc.setFontSize(6); doc.setFont("helvetica","normal");
    doc.text("Evidence Vault",m+18,19,{align:"center"});
    doc.setFont("helvetica","bold"); doc.setFontSize(13); doc.setTextColor(255,255,255);
    doc.text(title,m+42,13);
    doc.setFont("helvetica","normal"); doc.setFontSize(7.5); doc.setTextColor(200,220,215);
    doc.text(subtitle,m+42,20);
    doc.setFontSize(7);
    doc.text(`Page ${n}`,W-m,20,{align:"right"});
    doc.text(`Generated: ${new Date().toLocaleDateString("en-GB",{day:"2-digit",month:"long",year:"numeric"})}`,W-m,13,{align:"right"});
  };
  const addFooter = () => {
    doc.setFillColor(...hexRGB(C.forestLt)); doc.rect(0,pH-12,W,12,"F");
    doc.setFont("helvetica","normal"); doc.setFontSize(6.5); doc.setTextColor(...hexRGB(C.muted));
    doc.text("bidelevate.co.uk  |  Exported from BidElevate Evidence Vault  |  For tender preparation use only",W/2,pH-5,{align:"center"});
  };
  let pn=1; addHeader(pn); addFooter();
  doc.setFillColor(...hexRGB(C.forestDk)); doc.rect(m,32,cW,11,"F");
  doc.setFont("helvetica","bold"); doc.setFontSize(8.5); doc.setTextColor(255,255,255);
  doc.text(`Prepared by: ${orgName||"Your Organisation"}`,m+4,39);
  doc.setFont("helvetica","normal"); doc.setFontSize(7); doc.setTextColor(200,220,215);
  doc.text("Submitted to BidElevate for tender preparation",W-m-4,39,{align:"right"});
  doc.setFillColor(...hexRGB(C.forestLt)); doc.rect(m,46,cW,9,"F");
  doc.setFont("helvetica","normal"); doc.setFontSize(8); doc.setTextColor(...hexRGB(C.forestDk));
  doc.text(`${rows.length} record${rows.length!==1?"s":""} exported`,m+4,51.5);
  let y=60; const cp=5, lH=5.5;
  rows.forEach((row,idx) => {
    let eH=cp*2+6;
    columns.forEach(col => { const v=row[col.key]||"—"; const ls=doc.splitTextToSize(String(v),cW-cp*2-28); eH+=Math.max(ls.length,1)*lH+2; });
    eH+=4;
    if (y+eH>pH-20) { addFooter(); doc.addPage(); pn++; addHeader(pn); addFooter(); y=36; }
    doc.setFillColor(...(idx%2===0?[255,255,255]:hexRGB(C.warm)));
    doc.roundedRect(m,y,cW,eH,2,2,"F");
    doc.setDrawColor(...hexRGB(C.forest)); doc.setLineWidth(0.3); doc.roundedRect(m,y,cW,eH,2,2,"S");
    doc.setFillColor(...hexRGB(C.forest)); doc.roundedRect(m+cp,y+cp,8,6,1,1,"F");
    doc.setFont("helvetica","bold"); doc.setFontSize(6.5); doc.setTextColor(255,255,255);
    doc.text(`${idx+1}`,m+cp+4,y+cp+4.2,{align:"center"});
    let ry=y+cp+4;
    columns.forEach(col => {
      const v=row[col.key]||"—"; const lx=m+cp+12; const vx=lx+32; const vw=cW-cp-12-32;
      doc.setFont("helvetica","bold"); doc.setFontSize(7); doc.setTextColor(...hexRGB(C.forestDk)); doc.text(col.label+":",lx,ry);
      doc.setFont("helvetica","normal"); doc.setFontSize(7.5); doc.setTextColor(...hexRGB(C.text));
      const ls=doc.splitTextToSize(String(v),vw); doc.text(ls,vx,ry); ry+=Math.max(ls.length,1)*lH+2;
    });
    y+=eH+4;
  });
  doc.save(`BidElevate_${title.replace(/\s+/g,"_")}_${Date.now()}.pdf`);
}

async function generateExcel(title, rows, columns, orgName) {
  const XLSX = await loadXLSX();
  const ws = XLSX.utils.aoa_to_sheet([
    [`Prepared by: ${orgName||"Your Organisation"} — BidElevate Evidence Vault`],
    [`Exported: ${new Date().toLocaleDateString("en-GB")}`],[],
    columns.map(c=>c.label),...rows.map(r=>columns.map(c=>r[c.key]||""))
  ]);
  ws["!cols"]=columns.map(()=>({wch:28}));
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,title.slice(0,31));
  XLSX.writeFile(wb,`BidElevate_${title.replace(/\s+/g,"_")}_${Date.now()}.xlsx`);
}

// ── Helpers ───────────────────────────────────────────────────────────────────
const uid = () => Math.random().toString(36).slice(2)+Date.now().toString(36);
const fmt = d => d?new Date(d).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}):"—";
const weeksBetween = (s,e) => { if(!s) return null; const diff=(( e?new Date(e):new Date())-new Date(s))/(1000*60*60*24*7); return Math.max(0,Math.round(diff)); };
const formatCurrency = n => n?`£${n.toLocaleString("en-GB",{minimumFractionDigits:2,maximumFractionDigits:2})}`:"";

const SERVICE_TYPES = ["Domiciliary Care","Supported Living","Floating Support","Supported Accommodation","Residential Care","Children & Young People / SEND","Other"];
const SV_CATEGORIES = ["Local Employment","Apprenticeships","Work Experience","Volunteering","Community Events","Charitable Donations","Environmental Initiatives","Staff Wellbeing","Other"];

// ── Push notifications ────────────────────────────────────────────────────────
async function requestNotificationPermission() {
  if (!("Notification" in window)) return false;
  if (Notification.permission === "granted") return true;
  const perm = await Notification.requestPermission();
  return perm === "granted";
}

function scheduleWeeklyCheck() {
  // Check every hour if it's Monday 8:30am
  setInterval(() => {
    const now = new Date();
    if (now.getDay()===1 && now.getHours()===8 && now.getMinutes()<30) {
      if (Notification.permission==="granted") {
        navigator.serviceWorker.ready.then(reg => {
          reg.showNotification("BidElevate Evidence Vault", {
            body: "📋 Have you added any evidence this week? It only takes a minute.",
            icon: "/icon-192.png", badge: "/icon-192.png", tag: "weekly-reminder",
          });
        });
      }
    }
  }, 60*60*1000);
}

// ── UI primitives ─────────────────────────────────────────────────────────────
const Btn = ({children,onClick,variant="primary",size="md",disabled,style:s})=>{
  const base={display:"inline-flex",alignItems:"center",justifyContent:"center",gap:6,border:"none",borderRadius:10,fontWeight:600,cursor:disabled?"not-allowed":"pointer",transition:"all .15s",opacity:disabled?.5:1,fontSize:size==="sm"?12:size==="lg"?16:14,padding:size==="sm"?"7px 13px":size==="lg"?"15px 28px":"10px 18px",...s};
  const v={primary:{background:C.forest,color:C.white},secondary:{background:C.forestLt,color:C.forestDk},ghost:{background:"transparent",color:C.forest,border:`1.5px solid ${C.forest}`},danger:{background:C.alertLt,color:C.alert,border:`1px solid ${C.alert}`}};
  return <button style={{...base,...v[variant]}} onClick={onClick} disabled={disabled}>{children}</button>;
};
const Field=({label,children,required,hint})=>(
  <div style={{marginBottom:16}}>
    <label style={{display:"block",fontSize:13,fontWeight:600,color:C.forestDk,marginBottom:6}}>{label}{required&&<span style={{color:C.alert}}> *</span>}</label>
    {children}
    {hint&&<p style={{fontSize:11,color:C.muted,marginTop:4}}>{hint}</p>}
  </div>
);
const iS={width:"100%",padding:"12px 14px",borderRadius:10,border:`1.5px solid ${C.border}`,fontSize:14,color:C.text,background:C.white,outline:"none"};
const Input=p=><input style={iS} {...p}/>;
const Textarea=p=><textarea style={{...iS,resize:"vertical",minHeight:90}} {...p}/>;
const Select=({children,...p})=><select style={{...iS,appearance:"none",backgroundImage:`url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath fill='%233e554f' d='M1 1l5 5 5-5'/%3E%3C/svg%3E")`,backgroundRepeat:"no-repeat",backgroundPosition:"right 14px center"}} {...p}>{children}</select>;
const PIIWarning=()=>(<div style={{background:C.alertLt,border:`1.5px solid ${C.alert}`,borderRadius:10,padding:"12px 14px",marginBottom:20,display:"flex",gap:10,alignItems:"flex-start"}}><span style={{fontSize:18,flexShrink:0}}>⚠️</span><div><p style={{fontSize:13,fontWeight:700,color:C.alert,marginBottom:2}}>Anonymise all information</p><p style={{fontSize:12,color:C.alert,lineHeight:1.5}}>Do not enter names, addresses, NHS numbers, dates of birth, or any other personal information. Use role descriptions only (e.g. "65-year-old male with dementia").</p></div></div>);
const Badge=({children,color=C.forest})=><span style={{fontSize:11,fontWeight:600,padding:"3px 9px",borderRadius:20,background:color+"22",color,whiteSpace:"nowrap"}}>{children}</span>;
const EmptyState=({icon,title,sub,action})=>(<div style={{textAlign:"center",padding:"60px 20px"}}><div style={{fontSize:48,marginBottom:12}}>{icon}</div><p style={{fontSize:17,fontWeight:700,color:C.forestDk,marginBottom:6}}>{title}</p><p style={{fontSize:14,color:C.muted,marginBottom:24,lineHeight:1.6}}>{sub}</p>{action}</div>);
const Modal=({open,onClose,title,children})=>{
  if(!open) return null;
  return <div style={{position:"fixed",inset:0,background:"rgba(42,59,54,.55)",zIndex:1000,display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={onClose}><div style={{background:C.white,borderRadius:"20px 20px 0 0",width:"100%",maxWidth:480,maxHeight:"92vh",overflowY:"auto",padding:"24px 20px 32px"}} onClick={e=>e.stopPropagation()}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}><h2 style={{fontSize:18,fontWeight:700,color:C.forestDk}}>{title}</h2><button onClick={onClose} style={{background:C.forestLt,border:"none",borderRadius:8,width:32,height:32,fontSize:16,cursor:"pointer",color:C.muted}}>✕</button></div>{children}</div></div>;
};
const SearchBar=({value,onChange,placeholder})=>(<div style={{position:"relative",marginBottom:12}}><span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",fontSize:16,color:C.muted}}>🔍</span><input value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder||"Search…"} style={{...iS,paddingLeft:38}}/></div>);
const Ring=({value,max,size=88})=>{
  const pct=max?Math.min(value/max,1):0, r=(size-14)/2, circ=2*Math.PI*r, dash=circ*pct;
  return <svg width={size} height={size} style={{transform:"rotate(-90deg)"}}><circle cx={size/2} cy={size/2} r={r} fill="none" stroke={C.forestLt} strokeWidth={7}/><circle cx={size/2} cy={size/2} r={r} fill="none" stroke={C.forest} strokeWidth={7} strokeDasharray={`${dash} ${circ}`} strokeLinecap="round" style={{transition:"stroke-dasharray .6s ease"}}/></svg>;
};

// ── Loading spinner ───────────────────────────────────────────────────────────
const Spinner=()=><div style={{display:"flex",justifyContent:"center",padding:40}}><div style={{width:32,height:32,border:`3px solid ${C.forestLt}`,borderTop:`3px solid ${C.forest}`,borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/><style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style></div>;

// ── useSupabase hook ──────────────────────────────────────────────────────────
function useSupabaseTable(table, orgId) {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);

  const load = useCallback(async () => {
    if (!orgId) return;
    setLoading(true);
    const { data: rows } = await supabase.from(table).select("*").eq("org_id", orgId).order("created_at", { ascending: false });
    setData(rows || []);
    setLoading(false);
  }, [table, orgId]);

  useEffect(() => { load(); }, [load]);

  const insert = async (row) => {
    const { data: inserted } = await supabase.from(table).insert([{ ...row, org_id: orgId }]).select();
    if (inserted) setData(p => [inserted[0], ...p]);
  };

  const update = async (row) => {
    await supabase.from(table).update(row).eq("id", row.id);
    setData(p => p.map(r => r.id === row.id ? row : r));
  };

  const remove = async (id) => {
    await supabase.from(table).delete().eq("id", id);
    setData(p => p.filter(r => r.id !== id));
  };

  return { data, loading, insert, update, remove };
}

// ── SETUP SCREEN ──────────────────────────────────────────────────────────────
function SetupScreen({ onComplete }) {
  const [orgName, setOrgName] = useState("");
  const [services, setServices] = useState([]);
  const [loading, setLoading] = useState(false);
  const toggleService = s => setServices(p => p.includes(s) ? p.filter(x=>x!==s) : [...p,s]);
  const ok = orgName.trim().length > 2;

  const handleSubmit = async () => {
    setLoading(true);
    const orgId = uid();
    const org = { id: orgId, orgName: orgName.trim(), services };
    // Save org to supabase
    await supabase.from("organisations").insert([org]);
    localStorage.setItem("be_org_id", orgId);
    localStorage.setItem("be_org", JSON.stringify(org));
    // Request push notification permission
    await requestNotificationPermission();
    scheduleWeeklyCheck();
    onComplete(org);
    setLoading(false);
  };

  return (
    <div style={{minHeight:"100vh",background:C.warm,display:"flex",flexDirection:"column"}}>
      <div style={{background:C.forestDk,padding:"32px 24px 28px"}}>
        <img src="https://bidelevate.co.uk/wp-content/uploads/2026/05/BidElevate-Logo.png" alt="BidElevate" style={{height:36,marginBottom:8}}/>
        <p style={{fontSize:26,fontWeight:800,color:C.white,lineHeight:1.2,marginBottom:8}}>Evidence Vault</p>
        <p style={{fontSize:14,color:"#a8c5bc",lineHeight:1.5}}>Your year-round tender evidence bank. Let's get you set up — it takes less than a minute.</p>
      </div>
      <div style={{flex:1,padding:"28px 20px 40px",overflowY:"auto"}}>
        <div style={{background:C.white,borderRadius:16,border:`1.5px solid ${C.border}`,padding:"20px 18px",marginBottom:24}}>
          <p style={{fontSize:15,fontWeight:700,color:C.forestDk,marginBottom:8}}>👋 Welcome to Evidence Vault</p>
          <p style={{fontSize:13,color:C.muted,lineHeight:1.6}}>This app helps you capture case studies, social value activities, and contract records throughout the year — so when a tender arrives, your evidence is ready to go.</p>
        </div>
        <Field label="Your organisation name" required hint="This will appear on all exported documents.">
          <Input value={orgName} onChange={e=>setOrgName(e.target.value)} placeholder="e.g. Sunrise Care Services Ltd" autoFocus />
        </Field>
        <Field label="Which services do you deliver?" hint="Select all that apply. You can change this in Settings.">
          <div style={{display:"flex",flexDirection:"column",gap:8,marginTop:4}}>
            {SERVICE_TYPES.map(s=>(
              <div key={s} onClick={()=>toggleService(s)} style={{display:"flex",alignItems:"center",gap:12,background:services.includes(s)?C.forestLt:C.white,border:`1.5px solid ${services.includes(s)?C.forest:C.border}`,borderRadius:10,padding:"12px 14px",cursor:"pointer"}}>
                <div style={{width:20,height:20,borderRadius:6,border:`2px solid ${services.includes(s)?C.forest:C.border}`,background:services.includes(s)?C.forest:"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                  {services.includes(s)&&<span style={{color:C.white,fontSize:13,fontWeight:700}}>✓</span>}
                </div>
                <span style={{fontSize:14,fontWeight:services.includes(s)?600:400,color:services.includes(s)?C.forestDk:C.text}}>{s}</span>
              </div>
            ))}
          </div>
        </Field>
        <div style={{background:C.alertLt,border:`1.5px solid ${C.alert}`,borderRadius:12,padding:"14px 16px",marginBottom:28}}>
          <p style={{fontSize:13,fontWeight:700,color:C.alert,marginBottom:4}}>⚠️ Important — data privacy</p>
          <p style={{fontSize:12,color:C.alert,lineHeight:1.6}}>This app is for anonymised evidence only. Never enter service user names, addresses, NHS numbers, or any personal information.</p>
        </div>
        <Btn onClick={handleSubmit} disabled={!ok||loading} size="lg" style={{width:"100%"}}>{loading?"Setting up…":"Get started →"}</Btn>
        <p style={{fontSize:11,color:C.muted,textAlign:"center",marginTop:12}}>Provided free to BidElevate clients · bidelevate.co.uk</p>
      </div>
    </div>
  );
}

// ── CASE STUDY VAULT ──────────────────────────────────────────────────────────
const CS_COLUMNS=[{key:"date",label:"Date"},{key:"service_type",label:"Service Type"},{key:"challenge",label:"Need / Challenge"},{key:"intervention",label:"Intervention"},{key:"outcome",label:"Outcome"},{key:"lessons",label:"Lessons Learned"},{key:"tags",label:"Keywords / Tags"}];

function CaseStudyForm({initial,onSave,onCancel}){
  const blank={date:"",service_type:"",challenge:"",intervention:"",outcome:"",lessons:"",tags:""};
  const [f,setF]=useState(initial||blank);
  const set=k=>e=>setF(p=>({...p,[k]:e.target.value}));
  const ok=f.date&&f.service_type&&f.challenge&&f.outcome;
  return(
    <div>
      <PIIWarning/>
      <Field label="Date" required><Input type="date" value={f.date} onChange={set("date")}/></Field>
      <Field label="Service Type" required>
        <Select value={f.service_type} onChange={set("service_type")}>
          <option value="">Select service type…</option>
          {SERVICE_TYPES.map(s=><option key={s}>{s}</option>)}
        </Select>
      </Field>
      <Field label="Need / Challenge" required hint="Describe the presenting need in anonymised terms.">
        <Textarea value={f.challenge} onChange={set("challenge")} rows={3} placeholder="e.g. 65-year-old male with vascular dementia presenting with increasing falls risk…"/>
      </Field>
      <Field label="Intervention" hint="What did your service do?">
        <Textarea value={f.intervention} onChange={set("intervention")} rows={3} placeholder="e.g. Implemented a structured daily routine with two-carer visits…"/>
      </Field>
      <Field label="Outcome" required hint="What changed as a result?">
        <Textarea value={f.outcome} onChange={set("outcome")} rows={3} placeholder="e.g. Falls reduced by 80% over three months."/>
      </Field>
      <Field label="Lessons Learned"><Textarea value={f.lessons} onChange={set("lessons")} rows={2}/></Field>
      <Field label="Keywords / Tags" hint="Comma-separated. e.g. falls prevention, dementia, reablement">
        <Input value={f.tags} onChange={set("tags")} placeholder="falls prevention, dementia, reablement"/>
      </Field>
      <div style={{display:"flex",gap:10,marginTop:8}}>
        <Btn onClick={()=>onSave(f)} disabled={!ok} size="lg" style={{flex:1}}>Save Case Study</Btn>
        <Btn onClick={onCancel} variant="ghost">Cancel</Btn>
      </div>
    </div>
  );
}

function CaseStudyVault({db,org}){
  const {data,loading,insert,update,remove}=db;
  const [showForm,setShowForm]=useState(false);
  const [editing,setEditing]=useState(null);
  const [search,setSearch]=useState("");
  const [filterSvc,setFilterSvc]=useState("");
  const [filterTag,setFilterTag]=useState("");
  const [selected,setSelected]=useState([]);
  const [viewItem,setViewItem]=useState(null);

  const filtered=useMemo(()=>data.filter(r=>{
    const q=search.toLowerCase();
    const mQ=!q||[r.challenge,r.outcome,r.intervention,r.tags,r.lessons].some(f=>f?.toLowerCase().includes(q));
    const mS=!filterSvc||r.service_type===filterSvc;
    const mT=!filterTag||r.tags?.toLowerCase().includes(filterTag.toLowerCase());
    return mQ&&mS&&mT;
  }),[data,search,filterSvc,filterTag]);

  const allTags=useMemo(()=>[...new Set(data.flatMap(r=>r.tags?.split(",").map(t=>t.trim()).filter(Boolean)||[]))].sort(),[data]);
  const save=async f=>{ if(editing){await update(f);}else{await insert(f);} setShowForm(false);setEditing(null); };
  const del=async id=>{ if(window.confirm("Delete this case study?")){await remove(id);setViewItem(null);} };
  const toggleSel=id=>setSelected(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);
  const exportRows=selected.length?data.filter(r=>selected.includes(r.id)):filtered;

  if(loading) return <Spinner/>;
  return(
    <div style={{paddingBottom:20}}>
      <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
        <Btn onClick={()=>{setEditing(null);setShowForm(true);}} size="lg" style={{flex:1,minWidth:160}}>＋ Add Case Study</Btn>
        <Btn onClick={()=>generatePDF("Case Study Vault",`${exportRows.length} records`,exportRows.map(r=>({...r,date:fmt(r.date)})),CS_COLUMNS,org.orgName)} variant="secondary" size="sm">PDF</Btn>
        <Btn onClick={()=>generateExcel("Case Studies",exportRows.map(r=>({...r,date:fmt(r.date)})),CS_COLUMNS,org.orgName)} variant="secondary" size="sm">Excel</Btn>
      </div>
      <SearchBar value={search} onChange={setSearch} placeholder="Search case studies…"/>
      <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap"}}>
        <Select value={filterSvc} onChange={e=>setFilterSvc(e.target.value)} style={{...iS,flex:1,minWidth:140,fontSize:13,padding:"9px 12px"}}>
          <option value="">All service types</option>
          {SERVICE_TYPES.map(s=><option key={s}>{s}</option>)}
        </Select>
        <Select value={filterTag} onChange={e=>setFilterTag(e.target.value)} style={{...iS,flex:1,minWidth:120,fontSize:13,padding:"9px 12px"}}>
          <option value="">All tags</option>
          {allTags.map(t=><option key={t}>{t}</option>)}
        </Select>
      </div>
      {selected.length>0&&(
        <div style={{background:C.goldLt,border:`1px solid ${C.gold}`,borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          <span style={{fontSize:13,color:C.gold,fontWeight:600}}>{selected.length} selected</span>
          <Btn onClick={()=>generatePDF("Case Study Vault",`${selected.length} selected`,data.filter(r=>selected.includes(r.id)).map(r=>({...r,date:fmt(r.date)})),CS_COLUMNS,org.orgName)} size="sm" variant="secondary">Export PDF</Btn>
          <Btn onClick={()=>generateExcel("Case Studies",data.filter(r=>selected.includes(r.id)).map(r=>({...r,date:fmt(r.date)})),CS_COLUMNS,org.orgName)} size="sm" variant="secondary">Export Excel</Btn>
          <button onClick={()=>setSelected([])} style={{marginLeft:"auto",background:"none",border:"none",fontSize:12,color:C.muted,cursor:"pointer"}}>Clear</button>
        </div>
      )}
      {filtered.length===0?(
        <EmptyState icon="📋" title={data.length===0?"No case studies yet":"No results"} sub={data.length===0?"Start building your evidence bank. Add your first anonymised case study.":"Try adjusting your search or filters."} action={data.length===0&&<Btn onClick={()=>setShowForm(true)} size="lg">Add your first case study</Btn>}/>
      ):(
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
            <span style={{fontSize:12,color:C.muted}}>{filtered.length} record{filtered.length!==1?"s":""}</span>
            <button onClick={()=>selected.length===filtered.length?setSelected([]):setSelected(filtered.map(r=>r.id))} style={{fontSize:12,color:C.forest,background:"none",border:"none",fontWeight:600,cursor:"pointer"}}>{selected.length===filtered.length?"Deselect all":"Select all"}</button>
          </div>
          {filtered.map(r=>(
            <div key={r.id} onClick={()=>setViewItem(r)} style={{background:C.white,borderRadius:14,border:`1.5px solid ${selected.includes(r.id)?C.forest:C.border}`,padding:"14px 16px",cursor:"pointer"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:6}}>
                    <Badge>{r.service_type||"Unspecified"}</Badge>
                    <span style={{fontSize:11,color:C.muted,fontFamily:"monospace"}}>{fmt(r.date)}</span>
                  </div>
                  <p style={{fontSize:14,fontWeight:600,color:C.text,marginBottom:4,lineHeight:1.4}}>{r.challenge?.slice(0,100)}{r.challenge?.length>100?"…":""}</p>
                  <p style={{fontSize:12,color:C.muted,lineHeight:1.4}}>{r.outcome?.slice(0,80)}{r.outcome?.length>80?"…":""}</p>
                  {r.tags&&<div style={{marginTop:8,display:"flex",gap:4,flexWrap:"wrap"}}>{r.tags.split(",").filter(Boolean).map(t=><Badge key={t} color={C.forestMd}>{t.trim()}</Badge>)}</div>}
                </div>
                <input type="checkbox" checked={selected.includes(r.id)} onChange={e=>{e.stopPropagation();toggleSel(r.id);}} style={{width:18,height:18,accentColor:C.forest,flexShrink:0,marginTop:2}} onClick={e=>e.stopPropagation()}/>
              </div>
            </div>
          ))}
        </div>
      )}
      <Modal open={showForm||!!editing} onClose={()=>{setShowForm(false);setEditing(null);}} title={editing?"Edit Case Study":"New Case Study"}>
        <CaseStudyForm initial={editing} onSave={save} onCancel={()=>{setShowForm(false);setEditing(null);}}/>
      </Modal>
      <Modal open={!!viewItem} onClose={()=>setViewItem(null)} title="Case Study">
        {viewItem&&(<div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}><Badge>{viewItem.service_type}</Badge><span style={{fontSize:12,color:C.muted,fontFamily:"monospace"}}>{fmt(viewItem.date)}</span></div>
          {CS_COLUMNS.filter(c=>c.key!=="date"&&c.key!=="service_type").map(col=>viewItem[col.key]&&(<div key={col.key} style={{marginBottom:14}}><p style={{fontSize:11,fontWeight:700,color:C.forestMd,textTransform:"uppercase",letterSpacing:.5,marginBottom:4}}>{col.label}</p><p style={{fontSize:14,color:C.text,lineHeight:1.6}}>{viewItem[col.key]}</p></div>))}
          <div style={{display:"flex",gap:8,marginTop:20}}>
            <Btn onClick={()=>{setViewItem(null);setEditing(viewItem);}} variant="secondary" style={{flex:1}}>✏️ Edit</Btn>
            <Btn onClick={()=>del(viewItem.id)} variant="danger">Delete</Btn>
          </div>
        </div>)}
      </Modal>
    </div>
  );
}

// ── SOCIAL VALUE VAULT ────────────────────────────────────────────────────────
const SV_COLUMNS=[{key:"date",label:"Date"},{key:"category",label:"Category"},{key:"activity",label:"Activity"},{key:"description",label:"Description"},{key:"outcome",label:"Outcome"},{key:"estimated_value",label:"Estimated Value / Impact"}];
const CAT_COLORS={"Local Employment":"#2e7d5e","Apprenticeships":"#1565c0","Volunteering":"#6a1b9a","Community Events":"#e65100","Charitable Donations":"#c62828","Environmental Initiatives":"#2e7d32","Staff Wellbeing":"#4527a0","Work Experience":"#00695c","Other":"#546e7a"};

function SVForm({initial,onSave,onCancel}){
  const blank={date:"",category:"",activity:"",description:"",outcome:"",estimated_value:""};
  const [f,setF]=useState(initial||blank);
  const set=k=>e=>setF(p=>({...p,[k]:e.target.value}));
  const ok=f.date&&f.category&&f.activity;
  return(<div>
    <Field label="Date" required><Input type="date" value={f.date} onChange={set("date")}/></Field>
    <Field label="Category" required><Select value={f.category} onChange={set("category")}><option value="">Select category…</option>{SV_CATEGORIES.map(c=><option key={c}>{c}</option>)}</Select></Field>
    <Field label="Activity" required hint="Brief title for this social value activity."><Input value={f.activity} onChange={set("activity")} placeholder="e.g. Hired two local care workers from job fair"/></Field>
    <Field label="Description"><Textarea value={f.description} onChange={set("description")} rows={3}/></Field>
    <Field label="Outcome"><Textarea value={f.outcome} onChange={set("outcome")} rows={2} placeholder="e.g. Two permanent contracts created, both residents within 5 miles of the service."/></Field>
    <Field label="Estimated Value / Impact" hint="Monetary value, hours contributed, number of people reached, etc."><Input value={f.estimated_value} onChange={set("estimated_value")} placeholder="e.g. £4,200 in local wages / 48 volunteer hours"/></Field>
    <div style={{display:"flex",gap:10,marginTop:8}}><Btn onClick={()=>onSave(f)} disabled={!ok} size="lg" style={{flex:1}}>Save Activity</Btn><Btn onClick={onCancel} variant="ghost">Cancel</Btn></div>
  </div>);
}

function SocialValueVault({db,org}){
  const {data,loading,insert,update,remove}=db;
  const [showForm,setShowForm]=useState(false);
  const [editing,setEditing]=useState(null);
  const [search,setSearch]=useState("");
  const [filterCat,setFilterCat]=useState("");
  const [selected,setSelected]=useState([]);
  const [viewItem,setViewItem]=useState(null);
  const filtered=useMemo(()=>data.filter(r=>{const q=search.toLowerCase();return(!q||[r.activity,r.description,r.outcome,r.estimated_value].some(f=>f?.toLowerCase().includes(q)))&&(!filterCat||r.category===filterCat);}),[data,search,filterCat]);
  const save=async f=>{if(editing){await update(f);}else{await insert(f);}setShowForm(false);setEditing(null);};
  const del=async id=>{if(window.confirm("Delete this entry?")){await remove(id);setViewItem(null);}};
  const toggleSel=id=>setSelected(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);
  const exportRows=selected.length?data.filter(r=>selected.includes(r.id)):filtered;
  if(loading) return <Spinner/>;
  return(<div style={{paddingBottom:20}}>
    <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
      <Btn onClick={()=>{setEditing(null);setShowForm(true);}} size="lg" style={{flex:1,minWidth:160}}>＋ Add Activity</Btn>
      <Btn onClick={()=>generatePDF("Social Value Vault",`${exportRows.length} records`,exportRows.map(r=>({...r,date:fmt(r.date)})),SV_COLUMNS,org.orgName)} variant="secondary" size="sm">PDF</Btn>
      <Btn onClick={()=>generateExcel("Social Value",exportRows.map(r=>({...r,date:fmt(r.date)})),SV_COLUMNS,org.orgName)} variant="secondary" size="sm">Excel</Btn>
    </div>
    <SearchBar value={search} onChange={setSearch} placeholder="Search activities…"/>
    <div style={{marginBottom:16}}><Select value={filterCat} onChange={e=>setFilterCat(e.target.value)} style={{...iS,fontSize:13,padding:"9px 12px"}}><option value="">All categories</option>{SV_CATEGORIES.map(c=><option key={c}>{c}</option>)}</Select></div>
    {selected.length>0&&(<div style={{background:C.goldLt,border:`1px solid ${C.gold}`,borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}><span style={{fontSize:13,color:C.gold,fontWeight:600}}>{selected.length} selected</span><Btn onClick={()=>generatePDF("Social Value Vault",`${selected.length} selected`,data.filter(r=>selected.includes(r.id)).map(r=>({...r,date:fmt(r.date)})),SV_COLUMNS,org.orgName)} size="sm" variant="secondary">Export PDF</Btn><Btn onClick={()=>generateExcel("Social Value",data.filter(r=>selected.includes(r.id)).map(r=>({...r,date:fmt(r.date)})),SV_COLUMNS,org.orgName)} size="sm" variant="secondary">Export Excel</Btn><button onClick={()=>setSelected([])} style={{marginLeft:"auto",background:"none",border:"none",fontSize:12,color:C.muted,cursor:"pointer"}}>Clear</button></div>)}
    {filtered.length===0?(<EmptyState icon="🌱" title={data.length===0?"No social value activities yet":"No results"} sub={data.length===0?"Track community impact, employment, volunteering and more throughout the year.":"Try adjusting your search or filter."} action={data.length===0&&<Btn onClick={()=>setShowForm(true)} size="lg">Add your first activity</Btn>}/>):(<div style={{display:"flex",flexDirection:"column",gap:10}}><span style={{fontSize:12,color:C.muted}}>{filtered.length} record{filtered.length!==1?"s":""}</span>{filtered.map(r=>(<div key={r.id} onClick={()=>setViewItem(r)} style={{background:C.white,borderRadius:14,border:`1.5px solid ${selected.includes(r.id)?C.forest:C.border}`,padding:"14px 16px",cursor:"pointer"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}><div style={{flex:1,minWidth:0}}><div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:6}}><Badge color={CAT_COLORS[r.category]||C.forest}>{r.category}</Badge><span style={{fontSize:11,color:C.muted,fontFamily:"monospace"}}>{fmt(r.date)}</span></div><p style={{fontSize:14,fontWeight:600,color:C.text,marginBottom:4}}>{r.activity}</p>{r.estimated_value&&<p style={{fontSize:12,color:C.forestMd,fontWeight:500}}>💰 {r.estimated_value}</p>}</div><input type="checkbox" checked={selected.includes(r.id)} onChange={e=>{e.stopPropagation();toggleSel(r.id);}} style={{width:18,height:18,accentColor:C.forest,flexShrink:0}} onClick={e=>e.stopPropagation()}/></div></div>))}</div>)}
    <Modal open={showForm||!!editing} onClose={()=>{setShowForm(false);setEditing(null);}} title={editing?"Edit Activity":"New Social Value Activity"}><SVForm initial={editing} onSave={save} onCancel={()=>{setShowForm(false);setEditing(null);}}/></Modal>
    <Modal open={!!viewItem} onClose={()=>setViewItem(null)} title="Social Value Activity">{viewItem&&(<div><div style={{display:"flex",gap:8,marginBottom:16}}><Badge color={CAT_COLORS[viewItem.category]||C.forest}>{viewItem.category}</Badge><span style={{fontSize:12,color:C.muted,fontFamily:"monospace"}}>{fmt(viewItem.date)}</span></div>{SV_COLUMNS.filter(c=>c.key!=="date"&&c.key!=="category").map(col=>viewItem[col.key]&&(<div key={col.key} style={{marginBottom:14}}><p style={{fontSize:11,fontWeight:700,color:C.forestMd,textTransform:"uppercase",letterSpacing:.5,marginBottom:4}}>{col.label}</p><p style={{fontSize:14,color:C.text,lineHeight:1.6}}>{viewItem[col.key]}</p></div>))}<div style={{display:"flex",gap:8,marginTop:20}}><Btn onClick={()=>{setViewItem(null);setEditing(viewItem);}} variant="secondary" style={{flex:1}}>✏️ Edit</Btn><Btn onClick={()=>del(viewItem.id)} variant="danger">Delete</Btn></div></div>)}</Modal>
  </div>);
}

// ── CONTRACT TRACKER ──────────────────────────────────────────────────────────
const CT_COLUMNS=[{key:"contract_name",label:"Contract Name"},{key:"commissioner",label:"Commissioner / Client"},{key:"client_ref",label:"Client Reference"},{key:"service_type",label:"Service Type"},{key:"client_needs",label:"Client Needs"},{key:"start_date",label:"Start Date"},{key:"end_date",label:"End Date"},{key:"weekly_hours",label:"Weekly Hours"},{key:"hourly_rate",label:"Hourly Rate (£)"},{key:"calculated_value",label:"Calculated Lifetime Value"},{key:"calculation",label:"Calculation"},{key:"framework_value",label:"Total Framework Value"},{key:"service_users",label:"Service Users Supported"},{key:"commissioner_contact",label:"Commissioner Contact"},{key:"description",label:"Description of Services"},{key:"outcome",label:"Outcome Achieved"}];

function CTForm({initial,onSave,onCancel}){
  const blank={contract_name:"",commissioner:"",client_ref:"",service_type:"",client_needs:"",start_date:"",end_date:"",weekly_hours:"",hourly_rate:"",calculated_value:"",calculation:"",framework_value:"",service_users:"",commissioner_contact:"",description:"",outcome:""};
  const [f,setF]=useState(initial||blank);
  const set=k=>e=>setF(p=>({...p,[k]:e.target.value}));
  const hours=parseFloat(f.weekly_hours)||0, rate=parseFloat(f.hourly_rate)||0, weeks=weeksBetween(f.start_date,f.end_date||null);
  const hasCalc=hours>0&&rate>0&&weeks!==null&&weeks>0;
  const weeklyVal=hours*rate, lifetimeVal=weeklyVal*weeks;
  const calcStr=hasCalc?`${hours} hrs × £${rate} × ${weeks} weeks = ${formatCurrency(lifetimeVal)}`:"";
  useEffect(()=>{if(hasCalc)setF(p=>({...p,calculated_value:formatCurrency(lifetimeVal),calculation:calcStr}));},[f.weekly_hours,f.hourly_rate,f.start_date,f.end_date]);
  const ok=f.contract_name&&f.commissioner&&(f.calculated_value||f.weekly_hours);
  return(<div>
    <div style={{background:C.goldLt,border:`1px solid ${C.gold}`,borderRadius:10,padding:"12px 14px",marginBottom:20}}><p style={{fontSize:13,fontWeight:700,color:C.gold,marginBottom:2}}>⚠️ Framework value vs contract value</p><p style={{fontSize:12,color:"#7a5c00",lineHeight:1.5}}>Record what <strong>your organisation</strong> delivered — not the total framework value. Use the calculator below to get your auditable contract reference value.</p></div>
    <Field label="Contract Name" required><Input value={f.contract_name} onChange={set("contract_name")} placeholder="e.g. Domiciliary Care Framework — Lot 2"/></Field>
    <Field label="Commissioner / Client" required><Input value={f.commissioner} onChange={set("commissioner")} placeholder="e.g. Sandwell Metropolitan Borough Council"/></Field>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
      <Field label="Service Type"><Select value={f.service_type} onChange={set("service_type")}><option value="">Select…</option>{SERVICE_TYPES.map(s=><option key={s}>{s}</option>)}</Select></Field>
      <Field label="Client Reference" hint="Internal ID — no names."><Input value={f.client_ref} onChange={set("client_ref")} placeholder="e.g. SU-2023-014"/></Field>
    </div>
    <Field label="Client Needs" hint="Anonymised — no names or identifying details."><Input value={f.client_needs} onChange={set("client_needs")} placeholder="e.g. Adult with learning disability, tenancy instability"/></Field>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
      <Field label="Start Date" required><Input type="date" value={f.start_date} onChange={set("start_date")}/></Field>
      <Field label="End Date" hint="Leave blank if ongoing."><Input type="date" value={f.end_date} onChange={set("end_date")}/></Field>
    </div>
    <div style={{background:C.forestLt,border:`1.5px solid ${C.forest}`,borderRadius:14,padding:"16px",marginBottom:20}}>
      <p style={{fontSize:13,fontWeight:700,color:C.forestDk,marginBottom:8}}>📊 Contract Value Calculator</p>
      <p style={{fontSize:11,color:C.forestMd,marginBottom:12,lineHeight:1.5}}>Hours per week × Hourly rate × Total weeks = your auditable contract reference value</p>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
        <Field label="Weekly Hours"><Input type="number" value={f.weekly_hours} onChange={set("weekly_hours")} placeholder="e.g. 14" min="0" step="0.5"/></Field>
        <Field label="Hourly Rate (£)"><Input type="number" value={f.hourly_rate} onChange={set("hourly_rate")} placeholder="e.g. 21.50" min="0" step="0.01"/></Field>
      </div>
      {hasCalc?(
        <div style={{background:C.white,borderRadius:10,padding:"14px",border:`1.5px solid ${C.forest}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:8}}>
            <div><p style={{fontSize:11,fontWeight:700,color:C.forestMd,textTransform:"uppercase",letterSpacing:.5}}>Lifetime Contract Value</p><p style={{fontSize:28,fontWeight:800,color:C.forestDk,lineHeight:1.1}}>{formatCurrency(lifetimeVal)}</p></div>
            <div style={{textAlign:"right"}}><p style={{fontSize:11,color:C.muted}}>Weekly value</p><p style={{fontSize:16,fontWeight:700,color:C.forest}}>{formatCurrency(weeklyVal)}</p></div>
          </div>
          <div style={{background:C.forestLt,borderRadius:8,padding:"8px 10px"}}>
            <p style={{fontSize:11,color:C.forestDk,fontFamily:"monospace"}}>{hours} hrs × £{rate} × {weeks} weeks = {formatCurrency(lifetimeVal)}</p>
            <p style={{fontSize:10,color:C.muted,marginTop:2}}>{f.end_date?`${fmt(f.start_date)} → ${fmt(f.end_date)}`:`${fmt(f.start_date)} → today (${weeks} weeks so far)`}</p>
          </div>
        </div>
      ):(
        <div style={{background:C.white,borderRadius:10,padding:"14px",border:`1.5px dashed ${C.border}`,textAlign:"center"}}><p style={{fontSize:13,color:C.muted}}>Enter hours, rate and dates above to calculate</p></div>
      )}
    </div>
    <Field label="Total Framework Value (optional)" hint="The headline value across all providers — not your value alone."><Input value={f.framework_value} onChange={set("framework_value")} placeholder="e.g. £2,400,000"/></Field>
    <Field label="Number of Service Users Supported"><Input type="number" value={f.service_users} onChange={set("service_users")} placeholder="e.g. 24"/></Field>
    <Field label="Commissioner Contact" hint="Name and role of the council officer who placed the referral."><Input value={f.commissioner_contact} onChange={set("commissioner_contact")} placeholder="e.g. J. Smith, Contracts Manager, Sandwell MBC"/></Field>
    <Field label="Description of Services Delivered"><Textarea value={f.description} onChange={set("description")} rows={3} placeholder="e.g. Floating support delivered to an adult with a learning disability experiencing tenancy instability…"/></Field>
    <Field label="Outcome Achieved"><Textarea value={f.outcome} onChange={set("outcome")} rows={2} placeholder="e.g. Client maintained independent tenancy. Zero safeguarding incidents."/></Field>
    <div style={{display:"flex",gap:10,marginTop:8}}><Btn onClick={()=>onSave(f)} disabled={!ok} size="lg" style={{flex:1}}>Save Contract</Btn><Btn onClick={onCancel} variant="ghost">Cancel</Btn></div>
  </div>);
}

function ContractTracker({db,org}){
  const {data,loading,insert,update,remove}=db;
  const [showForm,setShowForm]=useState(false);
  const [editing,setEditing]=useState(null);
  const [search,setSearch]=useState("");
  const [selected,setSelected]=useState([]);
  const [viewItem,setViewItem]=useState(null);
  const filtered=useMemo(()=>data.filter(r=>{const q=search.toLowerCase();return!q||[r.contract_name,r.commissioner,r.description,r.outcome].some(f=>f?.toLowerCase().includes(q));}),[data,search]);
  const save=async f=>{if(editing){await update(f);}else{await insert(f);}setShowForm(false);setEditing(null);};
  const del=async id=>{if(window.confirm("Delete this contract record?")){await remove(id);setViewItem(null);}};
  const toggleSel=id=>setSelected(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);
  const exportRows=selected.length?data.filter(r=>selected.includes(r.id)):filtered;
  if(loading) return <Spinner/>;
  return(<div style={{paddingBottom:20}}>
    <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
      <Btn onClick={()=>{setEditing(null);setShowForm(true);}} size="lg" style={{flex:1,minWidth:160}}>＋ Add Contract</Btn>
      <Btn onClick={()=>generatePDF("Contract Value Tracker",`${exportRows.length} records`,exportRows.map(r=>({...r,start_date:fmt(r.start_date),end_date:fmt(r.end_date)})),CT_COLUMNS,org.orgName)} variant="secondary" size="sm">PDF</Btn>
      <Btn onClick={()=>generateExcel("Contracts",exportRows.map(r=>({...r,start_date:fmt(r.start_date),end_date:fmt(r.end_date)})),CT_COLUMNS,org.orgName)} variant="secondary" size="sm">Excel</Btn>
    </div>
    <SearchBar value={search} onChange={setSearch} placeholder="Search contracts…"/>
    {selected.length>0&&(<div style={{background:C.goldLt,border:`1px solid ${C.gold}`,borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}><span style={{fontSize:13,color:C.gold,fontWeight:600}}>{selected.length} selected</span><Btn onClick={()=>generatePDF("Contract Tracker",`${selected.length} selected`,data.filter(r=>selected.includes(r.id)).map(r=>({...r,start_date:fmt(r.start_date),end_date:fmt(r.end_date)})),CT_COLUMNS,org.orgName)} size="sm" variant="secondary">Export PDF</Btn><button onClick={()=>setSelected([])} style={{marginLeft:"auto",background:"none",border:"none",fontSize:12,color:C.muted,cursor:"pointer"}}>Clear</button></div>)}
    {filtered.length===0?(<EmptyState icon="📄" title={data.length===0?"No contracts recorded yet":"No results"} sub={data.length===0?"Record contracts as you deliver them. Accurate contract values are essential for tender submissions.":"Try adjusting your search."} action={data.length===0&&<Btn onClick={()=>setShowForm(true)} size="lg">Add your first contract</Btn>}/>):(<div style={{display:"flex",flexDirection:"column",gap:10}}><span style={{fontSize:12,color:C.muted}}>{filtered.length} record{filtered.length!==1?"s":""}</span>{filtered.map(r=>(<div key={r.id} onClick={()=>setViewItem(r)} style={{background:C.white,borderRadius:14,border:`1.5px solid ${selected.includes(r.id)?C.forest:C.border}`,padding:"14px 16px",cursor:"pointer"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}><div style={{flex:1,minWidth:0}}><p style={{fontSize:15,fontWeight:700,color:C.forestDk,marginBottom:2}}>{r.contract_name}</p><p style={{fontSize:13,color:C.muted,marginBottom:6}}>{r.commissioner}</p>{r.calculated_value&&(<div style={{background:C.forestLt,borderRadius:8,padding:"8px 10px",marginBottom:6,display:"inline-block"}}><p style={{fontSize:11,color:C.forestMd,fontWeight:600}}>Lifetime Contract Value</p><p style={{fontSize:20,fontWeight:800,color:C.forestDk}}>{r.calculated_value}</p>{r.calculation&&<p style={{fontSize:10,color:C.muted,fontFamily:"monospace",marginTop:2}}>{r.calculation}</p>}</div>)}<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{r.service_type&&<Badge color={C.forest}>{r.service_type}</Badge>}{r.service_users&&<Badge color={C.forestMd}>{r.service_users} service users</Badge>}{r.start_date&&<span style={{fontSize:11,color:C.muted,fontFamily:"monospace"}}>{fmt(r.start_date)} → {r.end_date?fmt(r.end_date):"ongoing"}</span>}</div></div><input type="checkbox" checked={selected.includes(r.id)} onChange={e=>{e.stopPropagation();toggleSel(r.id);}} style={{width:18,height:18,accentColor:C.forest,flexShrink:0}} onClick={e=>e.stopPropagation()}/></div></div>))}</div>)}
    <Modal open={showForm||!!editing} onClose={()=>{setShowForm(false);setEditing(null);}} title={editing?"Edit Contract":"New Contract Record"}><CTForm initial={editing} onSave={save} onCancel={()=>{setShowForm(false);setEditing(null);}}/></Modal>
    <Modal open={!!viewItem} onClose={()=>setViewItem(null)} title="Contract Record">{viewItem&&(<div><p style={{fontSize:18,fontWeight:700,color:C.forestDk,marginBottom:2}}>{viewItem.contract_name}</p><p style={{fontSize:14,color:C.muted,marginBottom:16}}>{viewItem.commissioner}</p>{viewItem.calculated_value&&(<div style={{background:`linear-gradient(135deg, ${C.forestDk}, ${C.forest})`,borderRadius:14,padding:"16px",marginBottom:20,color:C.white}}><p style={{fontSize:11,fontWeight:600,color:"#a8c5bc",textTransform:"uppercase",letterSpacing:.5,marginBottom:4}}>Auditable Contract Reference Value</p><p style={{fontSize:32,fontWeight:800,lineHeight:1,marginBottom:8}}>{viewItem.calculated_value}</p>{viewItem.calculation&&(<div style={{background:"rgba(255,255,255,.1)",borderRadius:8,padding:"8px 10px"}}><p style={{fontSize:11,fontFamily:"monospace",color:"#c8ddd8"}}>{viewItem.calculation}</p></div>)}</div>)}{CT_COLUMNS.filter(c=>!["contract_name","commissioner","calculated_value","calculation"].includes(c.key)).map(col=>viewItem[col.key]&&(<div key={col.key} style={{marginBottom:14}}><p style={{fontSize:11,fontWeight:700,color:C.forestMd,textTransform:"uppercase",letterSpacing:.5,marginBottom:4}}>{col.label}</p><p style={{fontSize:14,color:C.text,lineHeight:1.6,fontFamily:["start_date","end_date"].includes(col.key)?"monospace":"inherit"}}>{["start_date","end_date"].includes(col.key)?fmt(viewItem[col.key]):viewItem[col.key]}</p></div>))}<div style={{display:"flex",gap:8,marginTop:20}}><Btn onClick={()=>{setViewItem(null);setEditing(viewItem);}} variant="secondary" style={{flex:1}}>✏️ Edit</Btn><Btn onClick={()=>del(viewItem.id)} variant="danger">Delete</Btn></div></div>)}</Modal>
  </div>);
}

// ── DASHBOARD ─────────────────────────────────────────────────────────────────
function Dashboard({cs,sv,ct,setTab,org}){
  const total=cs.length+sv.length+ct.length;
  const typeColor={"Case Study":C.forest,"Social Value":"#2e7d5e","Contract":"#1565c0"};
  const recent=[...cs.map(r=>({...r,_type:"Case Study",_label:r.challenge?.slice(0,60)||"Case study"})),...sv.map(r=>({...r,_type:"Social Value",_label:r.activity})),...ct.map(r=>({...r,_type:"Contract",_label:r.contract_name}))].sort((a,b)=>new Date(b.created_at)-new Date(a.created_at)).slice(0,5);
  const motivate=total===0?"Your evidence bank is empty. Populate it with your evidence for future tenders.":total<=5?"Good start. Keep going!":total<15?"Building nicely. This is good — keep going.":"Strong evidence bank. Keep adding to it!";
  return(<div style={{paddingBottom:20}}>
    <div style={{background:C.white,border:`1.5px solid ${C.border}`,borderRadius:14,padding:"12px 16px",marginBottom:16,display:"flex",alignItems:"center",gap:12}}>
      <div style={{width:40,height:40,borderRadius:10,background:C.forestLt,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>🏥</div>
      <div><p style={{fontSize:15,fontWeight:700,color:C.forestDk}}>{org.orgName}</p>{org.services?.length>0&&<p style={{fontSize:11,color:C.muted}}>{org.services.slice(0,2).join(" · ")}{org.services.length>2?` +${org.services.length-2} more`:""}</p>}</div>
    </div>
    <div style={{background:`linear-gradient(135deg, ${C.forestDk} 0%, ${C.forest} 100%)`,borderRadius:20,padding:"24px 20px",marginBottom:20,color:C.white}}>
      <p style={{fontSize:12,fontWeight:600,letterSpacing:1,textTransform:"uppercase",color:"#a8c5bc",marginBottom:8}}>Evidence Strength</p>
      <div style={{display:"flex",alignItems:"center",gap:20,marginBottom:16}}>
        <div><p style={{fontSize:42,fontWeight:800,lineHeight:1}}>{total}</p><p style={{fontSize:13,color:"#a8c5bc"}}>total records</p></div>
        <Ring value={total} max={30} size={80}/>
      </div>
      <p style={{fontSize:13,color:"#c8ddd8",lineHeight:1.5}}>{motivate}</p>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:20}}>
      {[{icon:"📋",label:"Case Studies",count:cs.length,tab:1},{icon:"🌱",label:"Social Value",count:sv.length,tab:2},{icon:"📄",label:"Contracts",count:ct.length,tab:3}].map(s=>(
        <div key={s.tab} onClick={()=>setTab(s.tab)} style={{background:C.white,borderRadius:14,border:`1.5px solid ${C.border}`,padding:"14px 12px",textAlign:"center",cursor:"pointer"}}>
          <p style={{fontSize:22,marginBottom:4}}>{s.icon}</p>
          <p style={{fontSize:24,fontWeight:800,color:C.forestDk}}>{s.count}</p>
          <p style={{fontSize:11,color:C.muted,lineHeight:1.3}}>{s.label}</p>
        </div>
      ))}
    </div>
    <div style={{marginBottom:20}}>
      <p style={{fontSize:13,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:.5,marginBottom:10}}>Quick Add</p>
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {[{icon:"📋",label:"New Case Study",tab:1},{icon:"🌱",label:"New Social Value Activity",tab:2},{icon:"📄",label:"New Contract Record",tab:3}].map(q=>(
          <Btn key={q.tab} onClick={()=>setTab(q.tab)} size="lg" style={{width:"100%",justifyContent:"flex-start",gap:12,background:C.white,color:C.forestDk,border:`1.5px solid ${C.border}`}}>
            <span style={{fontSize:20}}>{q.icon}</span><span>{q.label}</span><span style={{marginLeft:"auto",color:C.muted,fontSize:18}}>›</span>
          </Btn>
        ))}
      </div>
    </div>
    {recent.length>0&&(<div>
      <p style={{fontSize:13,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:.5,marginBottom:10}}>Recent Activity</p>
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {recent.map((r,i)=>(<div key={i} style={{background:C.white,borderRadius:12,border:`1.5px solid ${C.border}`,padding:"12px 14px",display:"flex",alignItems:"center",gap:12}}>
          <div style={{width:8,height:8,borderRadius:"50%",background:typeColor[r._type],flexShrink:0}}/>
          <div style={{flex:1,minWidth:0}}><p style={{fontSize:13,fontWeight:600,color:C.text,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{r._label}</p><p style={{fontSize:11,color:C.muted}}>{r._type}</p></div>
          <span style={{fontSize:11,color:C.muted,fontFamily:"monospace",flexShrink:0}}>{fmt(r.date||r.start_date||r.created_at?.slice(0,10))}</span>
        </div>))}
      </div>
    </div>)}
  </div>);
}

// ── ROOT ──────────────────────────────────────────────────────────────────────
export default function App(){
  const [org,setOrg]=useState(null);
  const [orgLoading,setOrgLoading]=useState(true);
  const [tab,setTab]=useState(0);
  const [showSettings,setShowSettings]=useState(false);

  useEffect(()=>{
    const style=document.createElement("style");
    style.textContent=GLOBAL_CSS;
    document.head.appendChild(style);
    return()=>document.head.removeChild(style);
  },[]);

  // Load org from localStorage on boot
  useEffect(()=>{
    const stored=localStorage.getItem("be_org");
    if(stored){
      setOrg(JSON.parse(stored));
      scheduleWeeklyCheck();
    }
    setOrgLoading(false);
  },[]);

  const orgId=org?.id;
  const csDb=useSupabaseTable("case_studies",orgId);
  const svDb=useSupabaseTable("social_value",orgId);
  const ctDb=useSupabaseTable("contracts",orgId);

  if(orgLoading) return <div style={{display:"flex",justifyContent:"center",padding:80}}><Spinner/></div>;
  if(!org) return <SetupScreen onComplete={o=>{setOrg(o);localStorage.setItem("be_org",JSON.stringify(o));}}/>;

  if(showSettings) return(
    <div style={{maxWidth:480,margin:"0 auto",minHeight:"100vh",background:C.warm}}>
      <div style={{background:C.forestDk,padding:"16px 20px 14px"}}>
        <img src="https://bidelevate.co.uk/wp-content/uploads/2026/05/BidElevate-Logo.png" alt="BidElevate" style={{height:28,marginBottom:2}}/>
        <p style={{fontSize:18,fontWeight:800,color:C.white}}>⚙️ Settings</p>
      </div>
      <div style={{padding:"20px 16px"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:20}}>
          <button onClick={()=>setShowSettings(false)} style={{background:C.forestLt,border:"none",borderRadius:8,width:34,height:34,fontSize:16,cursor:"pointer"}}>‹</button>
          <p style={{fontSize:16,fontWeight:700,color:C.forestDk}}>Organisation Settings</p>
        </div>
        <Field label="Organisation name" required hint="Appears on all exported PDFs and spreadsheets.">
          <Input value={org.orgName} onChange={e=>setOrg(p=>({...p,orgName:e.target.value}))}/>
        </Field>
        <Field label="Services delivered">
          <div style={{display:"flex",flexDirection:"column",gap:8,marginTop:4}}>
            {SERVICE_TYPES.map(s=>{
              const sel=org.services?.includes(s);
              return(<div key={s} onClick={()=>setOrg(p=>({...p,services:sel?p.services.filter(x=>x!==s):[...(p.services||[]),s]}))} style={{display:"flex",alignItems:"center",gap:12,background:sel?C.forestLt:C.white,border:`1.5px solid ${sel?C.forest:C.border}`,borderRadius:10,padding:"12px 14px",cursor:"pointer"}}>
                <div style={{width:20,height:20,borderRadius:6,border:`2px solid ${sel?C.forest:C.border}`,background:sel?C.forest:"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{sel&&<span style={{color:C.white,fontSize:13,fontWeight:700}}>✓</span>}</div>
                <span style={{fontSize:14,color:C.text}}>{s}</span>
              </div>);
            })}
          </div>
        </Field>
        <Btn onClick={()=>{localStorage.setItem("be_org",JSON.stringify(org));setShowSettings(false);}} size="lg" style={{width:"100%"}}>Save changes</Btn>
      </div>
    </div>
  );

  const TABS=[
    {label:"Home",icon:"🏠",content:<Dashboard cs={csDb.data} sv={svDb.data} ct={ctDb.data} setTab={setTab} org={org}/>},
    {label:"Cases",icon:"📋",content:<CaseStudyVault db={csDb} org={org}/>},
    {label:"Social",icon:"🌱",content:<SocialValueVault db={svDb} org={org}/>},
    {label:"Contracts",icon:"📄",content:<ContractTracker db={ctDb} org={org}/>},
  ];

  return(
    <div style={{maxWidth:480,margin:"0 auto",minHeight:"100vh",display:"flex",flexDirection:"column",background:C.warm}}>
      <div style={{background:C.forestDk,padding:"16px 20px 14px",position:"sticky",top:0,zIndex:100}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div>
            <img src="https://bidelevate.co.uk/wp-content/uploads/2026/05/BidElevate-Logo.png" alt="BidElevate" style={{height:28,marginBottom:2}}/>
            <p style={{fontSize:18,fontWeight:800,color:C.white,lineHeight:1.2}}>{TABS[tab].icon} {tab===0?"Evidence Vault":TABS[tab].label}</p>
          </div>
          <button onClick={()=>setShowSettings(true)} style={{background:"rgba(255,255,255,.1)",border:"none",borderRadius:10,padding:"8px 12px",color:"#a8c5bc",fontSize:12,fontWeight:600,cursor:"pointer"}}>⚙️ {org.orgName?.split(" ")[0]}</button>
        </div>
      </div>
      <div style={{flex:1,padding:"20px 16px 0",overflowY:"auto"}}>{TABS[tab].content}</div>
      <div style={{background:C.white,borderTop:`1px solid ${C.border}`,display:"grid",gridTemplateColumns:"repeat(4,1fr)",position:"sticky",bottom:0,zIndex:100}}>
        {TABS.map((t,i)=>(
          <button key={i} onClick={()=>setTab(i)} style={{background:"none",border:"none",padding:"10px 4px 12px",display:"flex",flexDirection:"column",alignItems:"center",gap:3,cursor:"pointer",borderTop:`2.5px solid ${tab===i?C.forest:"transparent"}`}}>
            <span style={{fontSize:20}}>{t.icon}</span>
            <span style={{fontSize:10,fontWeight:tab===i?700:500,color:tab===i?C.forest:C.muted}}>{t.label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}
