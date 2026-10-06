"use client";
import React, { useState, useEffect, useRef, useMemo } from "react";

const MONTHS=["Ocak","Şubat","Mart","Nisan","Mayıs","Haziran","Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"];
const YEARS=[2025,2026,2027];
function getBiz(m:number,y:number):string[]{const d:string[]=[],n=new Date(y,m+1,0).getDate();for(let i=1;i<=n;i++){const dt=new Date(y,m,i);if(dt.getDay()!==0&&dt.getDay()!==6)d.push(i+" "+MONTHS[m].slice(0,3));}return d;}
const sK=(p:string,m:number,y:number)=>p+"_"+y+"_"+m;

// ═══════════════════════════════════════════════════════════════════════
// GLOBAL UNDO SYSTEM — tracks all localStorage writes across the app.
// Every write records {key, prevValue, newValue} into a stack.
// Ctrl+Z or the Undo button pops the last entry and restores prevValue.
// Stack is coalesced: rapid writes to the same key within 400ms merge
// so a user typing doesn't create 50 undo steps for one field.
// ═══════════════════════════════════════════════════════════════════════
type UndoEntry={key:string,prev:string|null,next:string|null,ts:number};
const __undoStack:UndoEntry[]=[];
const __undoListeners=new Set<()=>void>();
let __undoEnabled=true;
let __isUndoing=false;
const MAX_UNDO=200;

function notifyUndo(){__undoListeners.forEach(fn=>{try{fn();}catch{}});}

function trackedSetItem(key:string,val:string){
  if(typeof window==="undefined")return;
  if(__isUndoing){try{localStorage.setItem(key,val);}catch{}return;}
  if(!__undoEnabled){try{localStorage.setItem(key,val);}catch{}return;}
  let prev:string|null=null;
  try{prev=localStorage.getItem(key);}catch{}
  // Coalesce: merge with previous entry if same key and within 400ms
  const now=Date.now();
  const last=__undoStack[__undoStack.length-1];
  if(last&&last.key===key&&(now-last.ts)<400){
    last.next=val;last.ts=now;
  }else{
    __undoStack.push({key,prev,next:val,ts:now});
    if(__undoStack.length>MAX_UNDO)__undoStack.shift();
  }
  try{localStorage.setItem(key,val);}catch{}
  notifyUndo();
}

function trackedRemoveItem(key:string){
  if(typeof window==="undefined")return;
  if(__isUndoing){try{localStorage.removeItem(key);}catch{}return;}
  if(!__undoEnabled){try{localStorage.removeItem(key);}catch{}return;}
  let prev:string|null=null;
  try{prev=localStorage.getItem(key);}catch{}
  __undoStack.push({key,prev,next:null,ts:Date.now()});
  if(__undoStack.length>MAX_UNDO)__undoStack.shift();
  try{localStorage.removeItem(key);}catch{}
  notifyUndo();
}

function performUndo():boolean{
  if(__undoStack.length===0)return false;
  const entry=__undoStack.pop()!;
  __isUndoing=true;
  try{
    if(entry.prev===null){localStorage.removeItem(entry.key);}
    else{localStorage.setItem(entry.key,entry.prev);}
    // Fire storage event so components listening to storage events reload
    window.dispatchEvent(new StorageEvent("storage",{key:entry.key,newValue:entry.prev,oldValue:entry.next}));
  }catch{}
  __isUndoing=false;
  notifyUndo();
  return true;
}

// Hook used by React components to re-render when undo stack size changes
function useUndoStackSize():number{
  const[n,setN]=useState(__undoStack.length);
  useEffect(()=>{const fn=()=>setN(__undoStack.length);__undoListeners.add(fn);return()=>{__undoListeners.delete(fn);};},[]);
  return n;
}

// Global unread-chat count hook (for header badge)
function useChatUnread(username:string):number{
  const[n,setN]=useState(0);
  useEffect(()=>{
    const calc=()=>{
      try{
        const all=JSON.parse(localStorage.getItem("team_chat_messages")||"[]") as any[];
        const lr=parseInt(localStorage.getItem("team_chat_last_read_"+username)||"0");
        setN(all.filter((m:any)=>m.ts>lr).length);
      }catch{setN(0);}
    };
    calc();
    const onS=(e:StorageEvent)=>{if(!e.key||e.key==="team_chat_messages"||e.key?.startsWith("team_chat_last_read_"))calc();};
    window.addEventListener("storage",onS);
    const poll=setInterval(calc,3000);
    return()=>{window.removeEventListener("storage",onS);clearInterval(poll);};
  },[username]);
  return n;
}

// Drop-in replacements for existing svD function. trackedSetItem wraps writes so undo works.
function svD(p:string,m:number,y:number,d:any){try{trackedSetItem(sK(p,m,y),JSON.stringify(d));}catch{}}
function ldD(p:string,m:number,y:number){try{const d=localStorage.getItem(sK(p,m,y));return d?JSON.parse(d):{}}catch{return{}}}

function exportXLS(tableId:string,filename:string){
  const tbl=document.getElementById(tableId);if(!tbl)return;
  const clone=tbl.cloneNode(true) as HTMLElement;
  clone.querySelectorAll("input").forEach(inp=>{const td=document.createElement("td");td.textContent=inp.value||"";td.style.cssText=inp.parentElement?.style.cssText||"";if(inp.parentElement)inp.parentElement.replaceChild(td,inp);});
  clone.querySelectorAll("select").forEach(sel=>{const td=document.createElement("td");td.textContent=(sel as HTMLSelectElement).value||"";if(sel.parentElement)sel.parentElement.replaceChild(td,sel);});
  clone.querySelectorAll("button").forEach(b=>b.remove());
  const html='<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:spreadsheet"><head><meta charset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Sheet</x:Name></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body>'+clone.outerHTML+'</body></html>';
  const blob=new Blob([html],{type:'application/vnd.ms-excel'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename+'.xls';a.click();URL.revokeObjectURL(url);
}
function ExportBtn({tableId,name}:{tableId:string,name:string}){return<button onClick={()=>exportXLS(tableId,name)} className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[11px] font-semibold flex items-center gap-1.5 transition">📥 Excel\'e Aktar</button>;}

const COLL=[
  {id:"us_export",label:"US Export Receivables"},
  {id:"uk_export",label:"UK Export Receivables"},
  {id:"swe_recv",label:"SW Europe Receivables"},
  {id:"tr_export",label:"TR Export Receivables"},
  {id:"dom_cash",label:"Domestic Collection Cash"},
  {id:"dom_cheque",label:"Domestic Cheque Discount"},
  {id:"electricity_c",label:"Electricity"},
  {id:"interest_tr",label:"Interest Income (TR)"},
  {id:"interest_uk",label:"Interest Income (UK)"},
  {id:"vat_refund",label:"VAT Refund"},
  {id:"loan_util",label:"Loan Utilisation"},
  {id:"rent_kazan",label:"Rent Income (Kazan)"},
  {id:"ciner_in",label:"Incoming from Ciner Glass"},
  {id:"group_recv",label:"Receivables from Group"},
];
const PAY=[
  {id:"loan_rep_tr",label:"Loan Repayment (TR)"},
  {id:"loan_rep_uk",label:"Loan Repayment (UK)"},
  {id:"tcc_iet",label:"Kazan Soda TCC & IET (inv.)"},
  {id:"eti_royalty",label:"Eti Royalty - Licence"},
  {id:"nat_gas",label:"Kazan Soda Natural Gas"},
  {id:"payroll",label:"Payroll"},
  {id:"port",label:"Safi Derince - Port"},
  {id:"ssk",label:"Social Security"},
  {id:"tax",label:"Tax (VAT+Corp)"},
  {id:"siemens",label:"Kazan Eti Siemens (inv.)"},
  {id:"epias",label:"EPİAŞ Settlement"},
  {id:"coal",label:"Eti Soda - Coal"},
  {id:"transport",label:"Transportation (TR)"},
  {id:"fuel",label:"Fuel"},
  {id:"import_p",label:"Import"},
  {id:"export_p",label:"Export"},
  {id:"freight",label:"Export Freight"},
  {id:"trade_pay",label:"Trade Payables"},
  {id:"elec_kazan",label:"Electricity Kazan"},
  {id:"elec_eti",label:"Electricity Eti"},
  {id:"other_inv",label:"Other Investment"},
  {id:"maintenance",label:"Maintenance"},
  {id:"swe_freight",label:"SW Europe Freight"},
  {id:"wesoda_pay",label:"We Soda Payroll & Other"},
  {id:"sw_reseller",label:"SW Reseller Fee & Others"},
  {id:"we_ent",label:"We Enterprises Inc"},
  {id:"ciner_fund",label:"Ciner Glass Funding"},
  {id:"ansac",label:"Ansac Transfer"},
  {id:"group_pay",label:"Payables to Group"},
];

const USERS=[{username:"admin",password:"soda2026",role:"admin",name:"Hazine Yöneticisi"},{username:"analyst",password:"analyst2026",role:"analyst",name:"Hazine Analisti"},{username:"viewer",password:"viewer2026",role:"viewer",name:"Üst Yönetim"}];
const FX_D={usd_try:44.001,eur_usd:1.1553,gbp_usd:1.3335,eur_try:50.834,source:"Manuel",date:"10.03.2026"};
const fmt=(v:number|null|undefined)=>{if(v==null)return"";const n=Number(v);if(isNaN(n)||n===0)return"-";const a=Math.abs(n),s=a>=1?a.toFixed(2):a.toFixed(4),p=s.split(".");p[0]=p[0].replace(/\B(?=(\d{3})+(?!\d))/g,",");return n<0?"("+p.join(".")+")":p.join(".");};
const fmtC=(v:number)=>{if(!v)return"-";const a=Math.abs(v),p=Math.floor(a).toString().replace(/\B(?=(\d{3})+(?!\d))/g,",");return v<0?"("+p+")":p;};
const fmtM=(v:number|null|undefined)=>{const n=Number(v);return isNaN(n)||n===0?"-":n.toFixed(3);};
const addC=(v:string):string=>{const r=v.replace(/,/g,"");if(!r||isNaN(Number(r)))return v;const p=r.split(".");p[0]=p[0].replace(/\B(?=(\d{3})+(?!\d))/g,",");return p.join(".");};
const stripC=(v:string):string=>v.replace(/,/g,"");
const fetchTCMB=async():Promise<typeof FX_D>=>{
  // TCMB XML has no CORS headers — go through Next.js API proxy at /api/tcmb
  // Try previous business days (skip weekends/holidays), up to 7 days back
  for(let back=1;back<=7;back++){
    const d=new Date();d.setDate(d.getDate()-back);
    // Skip weekends (Saturday=6, Sunday=0)
    const dow=d.getDay();if(dow===0||dow===6)continue;
    const dd=String(d.getDate()).padStart(2,"0");
    const mm=String(d.getMonth()+1).padStart(2,"0");
    const yyyy=String(d.getFullYear());
    try{
      const r=await fetch(`/api/tcmb?date=${yyyy}-${mm}-${dd}`);
      if(!r.ok)continue;
      const t=await r.text();
      const p=new DOMParser();const x=p.parseFromString(t,"text/xml");
      const rates:Record<string,number>={};
      x.querySelectorAll("Currency").forEach(c=>{
        const code=c.getAttribute("CurrencyCode")||"";
        const b=parseFloat(c.querySelector("ForexBuying")?.textContent||"0");
        const s=parseFloat(c.querySelector("ForexSelling")?.textContent||"0");
        if(b>0&&s>0)rates[code]=(b+s)/2;
      });
      if(rates.USD&&rates.EUR&&rates.GBP){
        return{usd_try:rates.USD,eur_try:rates.EUR,gbp_usd:rates.GBP/rates.USD,eur_usd:rates.EUR/rates.USD,source:"TCMB",date:dd+"."+mm+"."+yyyy};
      }
    }catch{/* try older date */}
  }
  return FX_D;
};

// Global numeric style — Calibri preferred, with tabular nums for clean column alignment.
// Applied as `className={mn}` everywhere or via the `numFont` inline style for larger numbers.
const mn="text-xs num-cell";
const numFont={fontFamily:"'Calibri', 'Segoe UI', 'Arial', system-ui, sans-serif",fontVariantNumeric:"tabular-nums" as const,letterSpacing:"0.01em"};
const stL="sticky left-0 bg-white dark:bg-gray-900 z-10 px-2 py-1 text-xs whitespace-nowrap";
const ddS="px-3 py-1.5 border border-gray-300 rounded-lg text-sm bg-white text-gray-800 font-semibold cursor-pointer outline-none focus:ring-2 focus:ring-blue-400";
const inp="w-full px-1 py-0.5 border border-gray-200 rounded text-xs bg-white outline-none";

function Login({onLogin}:{onLogin:(u:typeof USERS[0])=>void}){const[u,setU]=useState("");const[p,setP]=useState("");const[err,setErr]=useState("");const go=()=>{const f=USERS.find(x=>x.username===u&&x.password===p);if(f)onLogin(f);else setErr("Hatalı kullanıcı adı veya şifre");};
  return(<div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-slate-50 to-blue-50"><div className="w-80 bg-white rounded-2xl shadow-xl p-8 border border-gray-100"><h2 className="text-xl font-bold mb-1">Soda Grubu Hazine</h2><p className="text-xs text-gray-400 mb-6">Nakit yönetim sistemi</p>{err&&<div className="bg-red-50 text-red-600 text-xs p-2 rounded-lg mb-3">{err}</div>}<label className="text-xs font-medium text-gray-600">Kullanıcı adı</label><input value={u} onChange={e=>setU(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()} className="w-full mt-1 mb-3 px-3 py-2 border border-gray-200 rounded-lg text-sm outline-none focus:ring-2 focus:ring-blue-400"/><label className="text-xs font-medium text-gray-600">Şifre</label><input type="password" value={p} onChange={e=>setP(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()} className="w-full mt-1 mb-4 px-3 py-2 border border-gray-200 rounded-lg text-sm outline-none focus:ring-2 focus:ring-blue-400"/><button onClick={go} className="w-full py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-sm font-semibold">Giriş yap</button><p className="mt-4 text-[10px] text-gray-400">admin / soda2026</p></div></div>);}

function Tb({active,color,children,onClick}:{active:boolean,color:string,children:React.ReactNode,onClick:()=>void}){return<button onClick={onClick} className={"px-3 py-2 text-xs whitespace-nowrap border-b-2 transition "+(active?"font-semibold border-current":"border-transparent text-gray-400 hover:text-gray-600")} style={active?{color}:{}}>{children}</button>;}
function Sec({children,color}:{children:React.ReactNode,color:string}){return<div className="rounded-md px-3 py-1.5 mt-3 mb-1.5 text-xs font-bold" style={{background:color}}>{children}</div>;}

function NoteCell({id,di,value,ed,onChange,notes,setNotes,sm,sy,storageKey}:{id:string,di:number,value:number,ed:boolean,onChange:(v:string)=>void,notes:Record<string,string>,setNotes:(n:Record<string,string>)=>void,sm:number,sy:number,storageKey?:string}){
  const nk=id+"_"+di;const note=notes[nk]||"";const[showE,setShowE]=useState(false);const[draft,setDraft]=useState("");const[hover,setHover]=useState(false);const[dv,setDv]=useState("");
  useEffect(()=>{setDv(value?addC(String(value)):"");},[value]);
  const saveN=(t:string)=>{const nn={...notes};if(t.trim())nn[nk]=t.trim();else delete nn[nk];setNotes(nn);svD(storageKey||"notes",sm,sy,nn);setShowE(false);};
  return(<div className="relative" onMouseEnter={()=>setHover(true)} onMouseLeave={()=>setHover(false)} onContextMenu={e=>{e.preventDefault();setDraft(note);setShowE(true);}}>
    {note&&<div className="absolute top-0 right-0 w-0 h-0 z-20" style={{borderLeft:"6px solid transparent",borderTop:"6px solid #ef4444"}}/>}
    {note&&hover&&!showE&&<div className="absolute z-30 bottom-full right-0 mb-1 px-2 py-1 bg-gray-800 text-white text-[10px] rounded shadow-lg max-w-[200px] whitespace-pre-wrap">{note}</div>}
    {ed?<input className={"w-full px-1 py-0.5 border border-gray-200 rounded text-right "+mn+" bg-amber-50 text-blue-700 outline-none"} value={dv} onChange={e=>{setDv(addC(e.target.value));onChange(stripC(e.target.value));}}/>
      :<span className={"block text-right px-1 "+mn+(value<0?" text-red-500":"")}>{fmt(value)}</span>}
    {showE&&<div className="absolute z-40 top-full left-0 mt-1 p-2 bg-white border border-gray-300 rounded-lg shadow-xl" style={{minWidth:200}}><textarea value={draft} onChange={e=>setDraft(e.target.value)} className="w-full h-14 px-2 py-1 border border-gray-200 rounded text-xs outline-none resize-none" autoFocus/><div className="flex gap-1 mt-1"><button onClick={()=>saveN(draft)} className="px-2 py-0.5 bg-blue-600 text-white rounded text-[10px]">Kaydet</button>{note&&<button onClick={()=>saveN("")} className="px-2 py-0.5 bg-red-100 text-red-600 rounded text-[10px]">Sil</button>}<button onClick={()=>setShowE(false)} className="px-2 py-0.5 bg-gray-100 rounded text-[10px]">İptal</button></div></div>}
  </div>);
}

function CashflowChart({sm,sy,fx}:{sm:number,sy:number,fx:typeof FX_D}){
  const biz=getBiz(sm,sy);
  const[cfData,setCfData]=useState<Record<string,any>>({});
  useEffect(()=>{
    const reload=()=>{setCfData(ldD("cf",sm,sy));};
    reload();
    const onStorage=(e:StorageEvent)=>{if(e.key&&e.key.startsWith("cf_"))reload();};
    const onFocus=()=>reload();
    window.addEventListener("storage",onStorage);
    window.addEventListener("focus",onFocus);
    const poll=setInterval(reload,2000);
    return()=>{window.removeEventListener("storage",onStorage);window.removeEventListener("focus",onFocus);clearInterval(poll);};
  },[sm,sy]);
  if(biz.length===0)return null;
  // Read nested format: {id: {di: value}}
  const cfV=(id:string,di:number)=>{const entry=cfData[id];if(entry&&typeof entry==="object"){return parseFloat(stripC((entry as any)[di]||""))||0;}return 0;};
  // Closing Balance USD per day = (Beginning TL day 0 + cumulative (tC - tP)) / fxRate / 1e6
  const beginTL0=cfV("begin_tl",0);
  const closingUSD:number[]=[];let running=beginTL0;
  biz.forEach((_,di)=>{const c=COLL.reduce((s,r)=>s+cfV(r.id,di),0);const p=PAY.reduce((s,r)=>s+cfV(r.id,di),0);running=running+c-p;closingUSD.push(running/fx.usd_try/1e6);});
  const W=700,H=240,PL=55,PR=15,PT=20,PB=35;const cW=W-PL-PR,cH=H-PT-PB;
  const mn2=Math.min(0,...closingUSD),mx=Math.max(0,...closingUSD)||1;
  const pad=(mx-mn2)*0.15||1;const yMin=mn2-pad,yMax=mx+pad,rng=yMax-yMin||1;
  const x=(i:number)=>PL+(closingUSD.length>1?(i/(closingUSD.length-1))*cW:cW/2);
  const y=(v:number)=>PT+cH-((v-yMin)/rng)*cH;
  const zero=y(0);
  const pts=closingUSD.map((v,i)=>`${x(i)},${y(v)}`).join(" ");
  const area=pts+` ${x(closingUSD.length-1)},${zero} ${x(0)},${zero}`;
  const gridVals=[yMin,(yMin+yMax)/2,yMax];
  const labelIdx=biz.length<=10?biz.map((_,i)=>i):[0,Math.floor(biz.length*0.25),Math.floor(biz.length*0.5),Math.floor(biz.length*0.75),biz.length-1];
  const last=closingUSD[closingUSD.length-1]||0;
  const hasData=closingUSD.some(v=>Math.abs(v)>0.001);
  if(!hasData)return<div className="py-8 text-center text-gray-400 text-xs">Cashflow\'a veri girildikçe grafik burada şekillenecek</div>;
  return<svg viewBox={`0 0 ${W} ${H}`} className="w-full" style={{maxWidth:W}}>
    <defs>
      <linearGradient id="cfGradPos" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#10b981" stopOpacity="0.4"/><stop offset="100%" stopColor="#10b981" stopOpacity="0.02"/></linearGradient>
      <linearGradient id="cfGradNeg" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#ef4444" stopOpacity="0.02"/><stop offset="100%" stopColor="#ef4444" stopOpacity="0.4"/></linearGradient>
    </defs>
    {gridVals.map((v,i)=><g key={i}><line x1={PL} y1={y(v)} x2={W-PR} y2={y(v)} stroke="#e5e7eb" strokeWidth="0.5" strokeDasharray="3,3"/><text x={PL-6} y={y(v)+3} textAnchor="end" fill="#9ca3af" fontSize="9" fontFamily="monospace">{v.toFixed(1)}</text></g>)}
    <line x1={PL} y1={zero} x2={W-PR} y2={zero} stroke="#6b7280" strokeWidth="1"/>
    <polygon points={area} fill={last>=0?"url(#cfGradPos)":"url(#cfGradNeg)"}/>
    <polyline points={pts} fill="none" stroke={last>=0?"#10b981":"#ef4444"} strokeWidth="2" strokeLinejoin="round" strokeLinecap="round"/>
    {closingUSD.map((v,i)=><circle key={i} cx={x(i)} cy={y(v)} r={i===closingUSD.length-1?4:2} fill={v>=0?"#059669":"#dc2626"} stroke="white" strokeWidth={i===closingUSD.length-1?1.5:0.5}/>)}
    {labelIdx.map(i=><text key={i} x={x(i)} y={H-10} textAnchor="middle" fill="#6b7280" fontSize="9" fontFamily="monospace">{biz[i]}</text>)}
    <text x={x(closingUSD.length-1)} y={y(last)-10} textAnchor="middle" fill={last>=0?"#059669":"#dc2626"} fontSize="11" fontWeight="bold" fontFamily="monospace">{last.toFixed(2)}M USD</text>
  </svg>;
}

// ═══════════════════════════════════════════════════════════════════════
// TEAM CHAT — internal chatbox on Dashboard.
// Messages persist in localStorage under 'team_chat_messages'.
// Each user has a 'team_chat_last_read' timestamp; new messages show badge.
// Uses native browser notifications + an audio ping when a new message arrives.
// Cross-tab updates via storage event.
// Single-browser limitation: for truly multi-user chat across different
// devices, a backend (WebSocket/server) is needed — this works across
// tabs and browser sessions on the same device.
// ═══════════════════════════════════════════════════════════════════════

type ChatMsg={id:string,user:string,text:string,ts:number};

function playPing(){
  try{
    const ctx=new (window.AudioContext||(window as any).webkitAudioContext)();
    const osc=ctx.createOscillator();const gain=ctx.createGain();
    osc.connect(gain);gain.connect(ctx.destination);
    osc.frequency.value=880;osc.type="sine";
    gain.gain.setValueAtTime(0.0001,ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.2,ctx.currentTime+0.02);
    gain.gain.exponentialRampToValueAtTime(0.0001,ctx.currentTime+0.4);
    osc.start();osc.stop(ctx.currentTime+0.4);
  }catch{}
}

function TeamChat({user}:{user:typeof USERS[0]}){
  const[msgs,setMsgs]=useState<ChatMsg[]>([]);
  const[draft,setDraft]=useState("");
  const[lastRead,setLastRead]=useState(0);
  const[open,setOpen]=useState(true);
  const[permReq,setPermReq]=useState(false);
  const listRef=useRef<HTMLDivElement>(null);
  const prevCountRef=useRef(0);

  // Load messages and last-read timestamp
  useEffect(()=>{
    const reload=()=>{
      try{
        const s=localStorage.getItem("team_chat_messages");
        const raw=s?JSON.parse(s):[];
        // Defensive: only keep messages with all required fields
        const all:ChatMsg[]=(Array.isArray(raw)?raw:[]).filter((m:any)=>m&&typeof m==="object"&&typeof m.text==="string"&&typeof m.user==="string"&&typeof m.ts==="number"&&m.id);
        setMsgs(all);
        const lr=parseInt(localStorage.getItem("team_chat_last_read_"+user.username)||"0");
        setLastRead(lr);
        // Detect new messages from OTHER users since last poll
        if(all.length>prevCountRef.current&&prevCountRef.current>0){
          const incoming=all.slice(prevCountRef.current);
          const fromOthers=incoming.filter(m=>m.user!==user.name);
          if(fromOthers.length>0){
            playPing();
            if("Notification" in window&&Notification.permission==="granted"){
              const last=fromOthers[fromOthers.length-1];
              try{new Notification(`${last.user} — Soda Grubu Chat`,{body:last.text.slice(0,120),icon:"/wesoda_logo.jpg",tag:"team-chat"});}catch{}
            }
          }
        }
        prevCountRef.current=all.length;
      }catch(err){console.error("TeamChat reload error:",err);}
    };
    reload();
    const onStorage=(e:StorageEvent)=>{if(!e.key||e.key==="team_chat_messages"||e.key?.startsWith("team_chat_last_read_"))reload();};
    window.addEventListener("storage",onStorage);
    const poll=setInterval(reload,3000);
    return()=>{window.removeEventListener("storage",onStorage);clearInterval(poll);};
  },[user]);

  // Auto-scroll to bottom when messages change or chat opens
  useEffect(()=>{
    if(open&&listRef.current){listRef.current.scrollTop=listRef.current.scrollHeight;}
  },[msgs,open]);

  // Mark-as-read when chat is open
  useEffect(()=>{
    if(open&&msgs.length>0){
      const newestTs=msgs[msgs.length-1].ts;
      if(newestTs>lastRead){
        // Write silently so it doesn't hit the undo stack
        __undoEnabled=false;try{localStorage.setItem("team_chat_last_read_"+user.username,String(newestTs));}catch{}__undoEnabled=true;
        setLastRead(newestTs);
      }
    }
  },[open,msgs,lastRead,user]);

  // Request notification permission on first open
  useEffect(()=>{
    if(open&&!permReq&&"Notification" in window&&Notification.permission==="default"){
      setPermReq(true);
      Notification.requestPermission().catch(()=>{});
    }
  },[open,permReq]);

  const send=()=>{
    const t=draft.trim();if(!t)return;
    const msg:ChatMsg={id:"msg_"+Date.now()+"_"+Math.random().toString(36).slice(2,6),user:user.name,text:t,ts:Date.now()};
    try{
      const cur=JSON.parse(localStorage.getItem("team_chat_messages")||"[]");
      const next=[...cur,msg].slice(-500);// cap at 500 messages
      // Write silently so chat messages don't clutter undo stack
      __undoEnabled=false;localStorage.setItem("team_chat_messages",JSON.stringify(next));__undoEnabled=true;
      setMsgs(next);
      // Also dispatch storage event so other tabs update immediately
      window.dispatchEvent(new StorageEvent("storage",{key:"team_chat_messages",newValue:JSON.stringify(next)}));
      // Mark my own read-state forward
      __undoEnabled=false;try{localStorage.setItem("team_chat_last_read_"+user.username,String(msg.ts));}catch{}__undoEnabled=true;
      setLastRead(msg.ts);
    }catch{}
    setDraft("");
  };

  const clearHistory=()=>{
    if(!confirm("Tüm mesajlar silinecek. Emin misiniz?"))return;
    __undoEnabled=false;try{localStorage.setItem("team_chat_messages","[]");}catch{}__undoEnabled=true;
    setMsgs([]);
    window.dispatchEvent(new StorageEvent("storage",{key:"team_chat_messages",newValue:"[]"}));
  };

  const unread=msgs.filter(m=>m.ts>lastRead&&m.user!==user.name).length;
  const fmtTime=(ts:number)=>{const d=new Date(ts);const now=new Date();const sameDay=d.toDateString()===now.toDateString();const yd=new Date();yd.setDate(yd.getDate()-1);const isYest=d.toDateString()===yd.toDateString();const hm=`${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`;if(sameDay)return hm;if(isYest)return `Dün ${hm}`;return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")} ${hm}`;};

  // Color per user (hash-based)
  const userColor=(name:string)=>{let h=0;for(let i=0;i<name.length;i++)h=(h*31+name.charCodeAt(i))>>>0;const colors=["#2563eb","#059669","#d97706","#7c3aed","#db2777","#0891b2","#ea580c","#16a34a"];return colors[h%colors.length];};

  return<div className="rounded-xl border border-gray-200 bg-white overflow-hidden">
    <div className="flex items-center justify-between px-3 py-2 bg-gradient-to-r from-indigo-50 to-blue-50 border-b border-gray-200 cursor-pointer" onClick={()=>setOpen(!open)}>
      <div className="flex items-center gap-2">
        <span className="text-xs font-bold text-indigo-700">💬 Takım Mesajlaşması</span>
        {unread>0&&<span className="px-1.5 py-0 bg-red-500 text-white rounded-full text-[9px] font-bold animate-pulse">{unread} yeni</span>}
        <span className="text-[10px] text-gray-400">{msgs.length} mesaj</span>
      </div>
      <div className="flex items-center gap-2">
        {msgs.length>0&&open&&<button onClick={(e)=>{e.stopPropagation();clearHistory();}} className="text-[9px] px-1.5 py-0.5 text-gray-400 hover:text-red-500" title="Tüm mesajları sil">🗑</button>}
        <span className="text-gray-400 text-[10px]">{open?"▼":"▶"}</span>
      </div>
    </div>
    {open&&<>
      <div ref={listRef} className="max-h-96 overflow-y-auto px-3 py-2 bg-gray-50/30 space-y-2" style={{minHeight:200}}>
        {msgs.length===0?<div className="text-center text-gray-400 text-xs py-12">Henüz mesaj yok.<br/>İlk mesajı siz yazın!</div>:msgs.map((m,idx)=>{
          if(!m||!m.id||!m.text)return null;
          const mine=m.user===user.name;
          const isNew=m.ts>lastRead&&!mine;
          return<div key={m.id||("m_"+idx)} className={"flex "+(mine?"justify-end":"justify-start")}>
            <div className={"max-w-[80%] rounded-lg px-3 py-1.5 "+(mine?"bg-indigo-500 text-white":"bg-white border border-gray-200")+(isNew?" ring-2 ring-red-300":"")}>
              {!mine&&<div className="text-[9px] font-bold mb-0.5" style={{color:userColor(m.user||"?")}}>{m.user||"?"}</div>}
              <div className={"text-[11px] whitespace-pre-wrap break-words "+(mine?"text-white":"text-gray-800")}>{m.text}</div>
              <div className={"text-[8px] mt-0.5 "+(mine?"text-indigo-100":"text-gray-400")}>{fmtTime(m.ts)}</div>
            </div>
          </div>;
        })}
      </div>
      <div className="flex gap-2 p-2 border-t border-gray-200 bg-white">
        <input value={draft} onChange={e=>setDraft(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();send();}}} placeholder={`${user.name} olarak mesaj yazın…`} className="flex-1 px-3 py-1.5 border border-gray-300 rounded-lg text-xs outline-none focus:border-indigo-400"/>
        <button onClick={send} disabled={!draft.trim()} className={"px-4 py-1.5 rounded-lg text-xs font-semibold transition "+(draft.trim()?"bg-indigo-600 hover:bg-indigo-700 text-white":"bg-gray-100 text-gray-300 cursor-not-allowed")}>Gönder</button>
      </div>
    </>}
  </div>;
}

function Dashboard({fx,setFx,sm,sy,user}:{fx:typeof FX_D,setFx:(f:typeof FX_D)=>void,sm:number,sy:number,user:typeof USERS[0]}){
  const biz=getBiz(sm,sy);
  const[upcoming,setUpcoming]=useState<{loan:string,date:string,total:string,ccy:string,overdue:boolean}[]>([]);
  const[overdueInvoices,setOverdueInvoices]=useState<{description:string,person:string,dueDate:string,totalUSD:number,days:number}[]>([]);
  // Read Cashflow & Kasa from localStorage
  const cf=ldD("cf",sm,sy);const kasaD=ldD("kasa2",sy,sm);
  const cfV=(id:string,d:number)=>parseFloat(stripC(cf[id+"_"+d]||""))||0;
  const tC=(d:number)=>COLL.reduce((s,r)=>s+cfV(r.id,d),0);
  const tP=(d:number)=>PAY.reduce((s,r)=>s+cfV(r.id,d),0);

  // Load critically overdue invoices (> 1 month past due)
  useEffect(()=>{
    const reload=()=>{
      try{
        const raw=localStorage.getItem("invoices_history");
        if(!raw)return;
        const all=JSON.parse(raw);
        if(!Array.isArray(all))return;
        const today=new Date();today.setHours(0,0,0,0);
        const crit=all.map((inv:any)=>{
          const p=(inv.dueDate||"").split(/[\/\-\.]/);
          if(p.length!==3||p[2]?.length!==4)return null;
          const d=new Date(+p[2],+p[1]-1,+p[0]).getTime();
          const days=Math.round((d-today.getTime())/86400000);
          if(days>=-30)return null; // only > 1 month overdue
          return{description:inv.description||"",person:inv.person||"",dueDate:inv.dueDate||"",totalUSD:inv.totalUSD||0,days};
        }).filter((x:any):x is {description:string,person:string,dueDate:string,totalUSD:number,days:number}=>x!==null).sort((a:any,b:any)=>b.totalUSD-a.totalUSD);
        setOverdueInvoices(crit);
      }catch{}
    };
    reload();
    const onS=(e:StorageEvent)=>{if(!e.key||e.key==="invoices_history")reload();};
    window.addEventListener("storage",onS);
    const poll=setInterval(reload,5000);
    return()=>{window.removeEventListener("storage",onS);clearInterval(poll);};
  },[]);

  useEffect(()=>{
    const reload=()=>{try{const reps=JSON.parse(localStorage.getItem("loan_reps")||"{}");const edits=JSON.parse(localStorage.getItem("loan_edits")||"{}");const customs=JSON.parse(localStorage.getItem("custom_loans")||"[]");const allL=[...LOANS,...customs];const now=new Date();now.setHours(0,0,0,0);const in30=new Date(now);in30.setDate(now.getDate()+30);const items:{loan:string,date:string,total:string,ccy:string,overdue:boolean}[]=[];
    Object.entries(reps).forEach(([lid,rows])=>{const loan=allL.find(l=>l.id===Number(lid));if(!loan)return;const lname=edits[lid]?.type||loan.type;const co=edits[lid]?.company||loan.company;(rows as RepRow[]).forEach(r=>{if(r.paid)return;if(!r.date)return;const parts=r.date.replace(/\//g,"-").split("-");let d:Date;if(parts[0]?.length===4)d=new Date(r.date);else if(parts[2]?.length===4)d=new Date(+parts[2],+parts[1]-1,+parts[0]);else d=new Date(r.date);if(isNaN(d.getTime()))return;d.setHours(0,0,0,0);const isOverdue=d<now;const isUpcoming=d>=now&&d<=in30;if(isOverdue||isUpcoming)items.push({loan:co+" — "+lname,date:r.date,total:r.total||r.principal||"0",ccy:loan.ccy,overdue:isOverdue});});});
    items.sort((a,b)=>a.date.localeCompare(b.date));setUpcoming(items);}catch(e){console.error(e);}};
    reload();
    const onS=(e:StorageEvent)=>{if(!e.key||e.key==="loan_reps"||e.key==="loan_edits"||e.key==="custom_loans")reload();};
    window.addEventListener("storage",onS);
    const poll=setInterval(reload,3000);
    return()=>{window.removeEventListener("storage",onS);clearInterval(poll);};
  },[]);

  const activeLoans=LOANS.filter(l=>l.status==="Active").length;
  const crd2="rounded-xl p-4 border";

  // Kasa closing values — prefer Cashflow-derived values (closing balance in USD millions)
  // falls back to explicit Kasa tab values if set.
  const kgn=(k:string)=>parseFloat(kasaD[k]||"")||0;

  // Compute closing balance from Cashflow (USD, millions) for last business day of month
  const cfClosingUSD=(()=>{
    if(biz.length===0)return 0;
    const usd=fx.usd_try||30;
    let openTL=parseFloat(stripC(cf["beg_tl_0"]||""))||0;
    let openUSD=parseFloat(stripC(cf["beg_usd_0"]||""))||0;
    for(let i=0;i<biz.length;i++){
      const tlIn=tC(i); // collections from Cashflow main sheet
      const tlOut=tP(i); // payments
      openTL+=tlIn-tlOut;
    }
    // Convert TL to USD then add USD directly; return in millions
    return (openTL/usd+openUSD)/1_000_000;
  })();

  // TR Kasa = today's Cashflow closing balance (TL→USD + USD combined, in USD).
  // Falls back: if today has no data, use the most recent non-zero business day.
  // Final fallback: the explicit Kasa tab value.
  const trKapLast=(() => {
    if(biz.length===0)return 0;
    // Walk Cashflow days from the most recent business day backwards; first non-zero closing wins
    const today=new Date();today.setHours(0,0,0,0);
    // Find today's index in biz (or closest past business day)
    let startIdx=biz.length-1;
    for(let i=0;i<biz.length;i++){const d=new Date(sy,sm,parseInt(biz[i]));d.setHours(0,0,0,0);if(d.getTime()>today.getTime())break;startIdx=i;}
    const usdRate=fx.usd_try||1;
    // Helper: compute closing USD-equivalent for a specific business-day index
    const closingAt=(endIdx:number)=>{
      // Seed with opening
      let openTL=parseFloat(stripC(cf["opening_tl"]||""))||0;
      let openUSD=parseFloat(stripC(cf["opening_usd"]||""))||0;
      for(let i=0;i<=endIdx;i++){
        openTL+=tC(i)-tP(i);  // net TL flow
      }
      return (openTL/usdRate+openUSD)/1_000_000;
    };
    // Try today (startIdx), then walk back to find a meaningful value
    for(let i=startIdx;i>=0;i--){
      const v=closingAt(i);
      if(Math.abs(v)>0.0001)return v;
    }
    // Fall back to explicit Kasa tab
    const explicit=(() => {let v=kgn("tr_acilis_init");for(let i=0;i<biz.length;i++){v=v+kgn("tr_ukgelen_"+i)+kgn("tr_swegelen_"+i)+kgn("tr_kdviade_"+i)+kgn("tr_tahsilat_"+i)-kgn("tr_odemeler_"+i);}return v;})();
    return Math.abs(explicit)>0.001?explicit:cfClosingUSD;
  })();

  // UK Kasa = today's closing from Kasa tab data (accumulated through today)
  const ukKapLast=(() => {
    if(biz.length===0)return 0;
    const today=new Date();today.setHours(0,0,0,0);
    let startIdx=biz.length-1;
    for(let i=0;i<biz.length;i++){const d=new Date(sy,sm,parseInt(biz[i]));d.setHours(0,0,0,0);if(d.getTime()>today.getTime())break;startIdx=i;}
    // Accumulate UK Kasa values up to today's index
    let v=kgn("uk_acilis_init");
    for(let i=0;i<=startIdx;i++){v=v+kgn("uk_export_"+i)+kgn("uk_usexport_"+i)-kgn("uk_odemeler_"+i)-kgn("uk_ansac_"+i)-kgn("uk_trye_"+i)-kgn("uk_sweye_"+i);}
    return v;
  })();

  // SWE GmbH (hidden from UI but still used internally)
  const sweKapLast=biz.length>0?(() => {let v=kgn("swe_acilis_init");for(let i=0;i<biz.length;i++){v=v+kgn("swe_ukgelen_"+i)+kgn("swe_export_"+i)-kgn("swe_odemeler_"+i)-kgn("swe_trye_"+i);}return v;})():0;
  const totalPos=trKapLast+ukKapLast+sweKapLast;

  const today=new Date();
  const DAYS_TR=["Pazar","Pazartesi","Salı","Çarşamba","Perşembe","Cuma","Cumartesi"];
  const dateStr=today.getDate()+" "+MONTHS[today.getMonth()]+" "+today.getFullYear()+" "+DAYS_TR[today.getDay()];
  return<div className="space-y-4">
    <div className="relative rounded-2xl p-5 overflow-hidden" style={{background:"linear-gradient(135deg,#1e3a8a 0%,#3b82f6 50%,#10b981 100%)"}}>
      <div className="absolute inset-0 opacity-20" style={{backgroundImage:"radial-gradient(circle at 20% 50%,white 0%,transparent 50%),radial-gradient(circle at 80% 50%,white 0%,transparent 50%)"}}></div>
      <div className="relative">
        <div className="text-[10px] text-blue-100 uppercase tracking-widest font-medium">Bugün</div>
        <div className="text-2xl font-bold text-white mt-1" style={{textShadow:"0 2px 4px rgba(0,0,0,0.2)"}}>{dateStr}</div>
      </div>
    </div>
    <div className="grid grid-cols-1 lg:grid-cols-3 gap-3">
      <div className={crd2+" bg-gradient-to-br from-blue-50 to-white border-blue-100"}>
        <div className="text-[10px] text-blue-400 font-medium uppercase tracking-wider">TR Kasa</div>
        <div className="text-2xl font-semibold text-blue-700 mt-1" style={numFont}>{trKapLast.toFixed(1)}<span className="text-sm ml-0.5 text-blue-400 font-normal">M USD</span></div>
        <div className="text-[9px] text-blue-400/80 mt-1">bugünkü kapanış</div>
      </div>
      <div className={crd2+" bg-gradient-to-br from-emerald-50 to-white border-emerald-100"}>
        <div className="text-[10px] text-emerald-400 font-medium uppercase tracking-wider">UK Kasa</div>
        <div className="text-2xl font-semibold text-emerald-700 mt-1" style={numFont}>{ukKapLast.toFixed(1)}<span className="text-sm ml-0.5 text-emerald-400 font-normal">M USD</span></div>
        <div className="text-[9px] text-emerald-400/80 mt-1">bugünkü kapanış</div>
      </div>
      <div className={crd2+" bg-gradient-to-br from-amber-50 to-white border-amber-100"}>
        <div className="text-[10px] text-amber-500 font-medium uppercase tracking-wider">Aktif Kredi</div>
        <div className="text-2xl font-semibold text-amber-700 mt-1" style={numFont}>{activeLoans}<span className="text-sm ml-1 text-amber-400 font-normal">/ {LOANS.length}</span></div>
        <div className="text-[9px] text-amber-500/80 mt-1">canlı kredi adedi</div>
      </div>
    </div>
    <div className={crd2+" bg-white border-gray-200"}>
      <div className="flex justify-between items-center mb-2">
        <span className="text-sm font-bold text-gray-700">Günlük Closing Balance USD <span className="font-normal text-gray-400 text-xs ml-1">{MONTHS[sm]} {sy} — milyon USD</span></span>
        <span className="text-[10px] px-2 py-0.5 bg-emerald-50 text-emerald-700 rounded-full font-semibold">Cashflow\'dan beslenir</span>
      </div>
      <CashflowChart sm={sm} sy={sy} fx={fx}/>
    </div>
    <div className={crd2+" bg-white border-gray-200"}>
      <div className="flex items-center justify-between mb-1">
        <div className="text-xs font-bold text-gray-700">Döviz Kurları <span className={"font-normal ml-1 "+(fx.source==="TCMB"?"text-emerald-600":"text-amber-600")}>• {fx.source==="TCMB"?"✓":"⚠"} {fx.source} {fx.date}</span></div>
        <button onClick={()=>{fetchTCMB().then(setFx);}} className="text-[10px] px-2 py-0.5 border border-gray-200 rounded hover:bg-gray-50" title="Kurları TCMB'den yenile">↻ Yenile</button>
      </div>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mt-2">{[{l:"USD/TRY",v:fx.usd_try,c:"#3b82f6"},{l:"EUR/USD",v:fx.eur_usd,c:"#10b981"},{l:"GBP/USD",v:fx.gbp_usd,c:"#7c3aed"},{l:"EUR/TRY",v:fx.eur_try,c:"#f59e0b"}].map((r,i)=><div key={i} className="text-center p-3 rounded-lg bg-gray-50"><span className="text-[11px] text-gray-500 block">{r.l}</span><span className={"text-base font-bold "+mn} style={{color:r.c}}>{r.v.toFixed(4)}</span></div>)}</div>
    </div>
    {upcoming.length>0&&<div className="rounded-xl p-5 bg-white border border-slate-200 shadow-sm">
      <div className="flex items-start justify-between mb-4 gap-4">
        <div className="flex items-start gap-3">
          <div className="w-9 h-9 rounded-lg bg-rose-50 border border-rose-200 flex items-center justify-center flex-shrink-0">
            <span className="text-rose-600 text-lg">📅</span>
          </div>
          <div>
            <div className="text-[13px] font-semibold text-slate-800">Yaklaşan &amp; Gecikmiş Kredi Ödemeleri</div>
            <div className="text-[10px] text-slate-500 mt-0.5">Önümüzdeki 30 gün içinde vadesi gelen ve vadesi geçmiş taksitler</div>
          </div>
        </div>
        <div className="text-right">
          <div className="text-[9px] text-slate-400 uppercase tracking-wider">{upcoming.length} taksit</div>
          <div className="text-[10px] text-rose-600 font-medium mt-0.5">{upcoming.filter(u=>u.overdue).length} gecikmiş</div>
        </div>
      </div>
      <div className="space-y-1 max-h-64 overflow-y-auto">{upcoming.map((u,i)=><div key={i} className={"flex justify-between items-center py-2 px-3 rounded-md text-[11px] transition-colors "+(u.overdue?"bg-rose-50/60 hover:bg-rose-50":"bg-slate-50/60 hover:bg-slate-50")}>
        <div className="flex-1 min-w-0">
          <div className="font-medium text-slate-700 truncate" title={u.loan}>{u.loan}</div>
          <div className="text-[9px] text-slate-500 mt-0.5">Vade: {u.date}{u.overdue&&<span className="ml-2 text-rose-600 font-medium">• gecikmiş</span>}</div>
        </div>
        <div className="text-right ml-3 flex-shrink-0">
          <div className={"font-semibold "+(u.overdue?"text-rose-700":"text-slate-800")} style={{...numFont,fontSize:"12px"}}>{addC(u.total)} <span className="text-[9px] text-slate-400 font-normal">{u.ccy}</span></div>
        </div>
      </div>)}</div>
    </div>}
    {overdueInvoices.length>0&&<div className="rounded-xl p-5 bg-white border border-slate-200 shadow-sm">
      <div className="flex items-start justify-between mb-4 gap-4">
        <div className="flex items-start gap-3">
          <div className="w-9 h-9 rounded-lg bg-amber-50 border border-amber-200 flex items-center justify-center flex-shrink-0">
            <span className="text-amber-600 text-lg">⚠</span>
          </div>
          <div>
            <div className="text-[13px] font-semibold text-slate-800">Uzun Süredir Gecikmiş Faturalar</div>
            <div className="text-[10px] text-slate-500 mt-0.5">Vadesi 30+ gün önce geçmiş ve hâlâ ödenmemiş</div>
          </div>
        </div>
        <div className="text-right">
          <div className="text-[9px] text-slate-400 uppercase tracking-wider">Toplam</div>
          <div className="text-[20px] font-semibold text-slate-800 mt-0.5" style={{fontFamily:"'Calibri', 'Segoe UI', system-ui, sans-serif",fontVariantNumeric:"tabular-nums"}}>{fmt(overdueInvoices.reduce((s,i)=>s+i.totalUSD,0))}<span className="text-[11px] ml-1 text-slate-400 font-normal">USD</span></div>
          <div className="text-[10px] text-amber-600 font-medium">{overdueInvoices.length} fatura</div>
        </div>
      </div>
      <div className="space-y-1 max-h-64 overflow-y-auto">
        {overdueInvoices.slice(0,15).map((inv,i)=><div key={i} className="flex justify-between items-center py-2 px-3 bg-slate-50/60 hover:bg-amber-50/40 rounded-md text-[11px] transition-colors">
          <div className="flex-1 min-w-0">
            <div className="font-medium text-slate-700 truncate" title={inv.description}>{inv.description}</div>
            <div className="text-[9px] text-slate-500 mt-0.5">{inv.person||"—"} <span className="text-slate-300 mx-1">•</span> Vade: {inv.dueDate}</div>
          </div>
          <div className="text-right ml-3 flex-shrink-0">
            <div className="font-semibold text-slate-800" style={{fontFamily:"'Calibri', 'Segoe UI', system-ui, sans-serif",fontVariantNumeric:"tabular-nums",fontSize:"12px"}}>{fmt(inv.totalUSD)} <span className="text-[9px] text-slate-400">USD</span></div>
            <div className="text-[9px] text-amber-600 font-medium mt-0.5">{Math.abs(inv.days)} gün gecikmiş</div>
          </div>
        </div>)}
        {overdueInvoices.length>15&&<div className="text-center text-[10px] text-slate-400 py-2">... ve {overdueInvoices.length-15} fatura daha</div>}
      </div>
    </div>}
    <TeamChat user={user}/>
  </div>;
}

// Sub-sheet firm arrays for feeding
const UK_F=["ALKIMIA PACKAGING","AARNA INTERNATIONAL (INDIA)","AB ETI PRODUCTS OY","ARDAGH LTD","AGC GLASS EUROPE SA","AGC INC","AL SAMA CO FOR IMPORT & EXPORT","ALINDA - VELCO SA","ALINDA BULGARIA","ARDAGH GLASS","BERNER OY","CERMIN MINERALS / CERMIN DMCC","CEBRACE","CHINA TIANCHEN ENGINEERING (TCC)","CONTINENTAL","DORMEX TRADE & INVESTMENTS","DORMEX IMPORT AND MARKETING LTD","ENCIRC","ELLERBROOK ROHSTOFFHANDEL","FOSFITALIA SPA","GOVCREST INTERNATIONAL SRL","GREENLINE CHEMICALS LTD (Cash)","GUARDIAN DO BRASIL","GUARDIAN INDUSTRIES CORP LTD (TH)","GUARDIAN INDUSTRIES RAYONG CO LTD (TH)","GUARDIAN Zoujaj","GUARDIAN Europe","GUARDIAN UK","Intersac","ICST","Horus Shipping","HUNAN ZHONGWEI HOLDING","ITAU BBA","GLOBAL SILICATES","KEMİKA","LEESON","KUWAIT INSULATING MATERIAL","MANUCHAR","MIRASCO","MITSUI","MMCT","MULTIMIN EGYPT","NEWPORT INDUSTRIES LTD","O-I ITALY SPA","O-I GLASS","O-I NETHERLANDS BV","O-I ESTONİA","O-I DO BRASIL INDUSTRIA","O-I FRANCE","On Route Logistics","PROCTER AND GAMBLE EGYPT","PQ Silicas UK","PQ SILICAS BRAZIL LTDA","GLOBAL SILICATES EGYPT","SISECAM ITALY (BNP)","SAINT-GOBAIN UK","SAINT-GOBAIN LATAM","SAINT-GOBAIN DEUTSCHLAND","SAISA CHEMICALS SA","SAUDI GUARDIAN","SOC. CHIMICA EMILIO FEDELI SPA","SCS","SALCHEM","TAM TRADING","TRADEASIA","TRAXYS EUROPE S.A.","UNILEVER ASIA PRIVATE LIMITED","UNILEVER MASHREQ EGYPT","UNIGLOBAL","UG Impex","VERALLIA UK","VIDIERA","VETRERIE MERIDIONALI SPA"];
const SWE_F=["AB Eti","AGC GLASS EUROPE SA","AGC Flat Glass Czech","O-I NETHERLANDS BV","ENCIRC","O-I ESTONİA","PQ Silicas BV","Guardian Europe","ST Gobain Deutschland","TRAXYS EUROPE S.A."];
const US_F=["CO VIDRIERA GMBH","CRISA LIBBEY MEXICO S DE RL DE CV","Cristaleria del Ecuador S.A. CRIDESA","Cristaleria Peldar S.A.","Desmoltandes","INDUSTRIA DEL ALCALI, S.A. DE C.V.","Minera del Altiplano S.A.","NATRIO","OWENS AMERICA S DE RL DE CV","Raupater S.A.","SAINT GOBAIN MEXICO SA DE CV","Sales de Jujuy S.A.","SAVERGLASS S DE RL DE CV","SQM SALAR SPA","VIDRIO PLANO DE MEXICO, SA DE CV","VITRO ENVASES, S.A. DE C.V.","VITRO VIDRIO AUTOMOTRIZ SA DE CV"];
const DOM_F=["VESKİM KİMYEVİ MADDELER A.Ş.","GÜROK TURİZM VE MADENCİLİK A.Ş (ÇEK)","DÜZCE CAM SANAYİ VE TİCARET A. (ÇEK)","PARK CAM","TÜRKİYE ŞİŞE VE CAM FABRİKALAR","RECKITT BENCKISER EV VE HİJYEN","TURK HENKEL","UNILEVER SAN. VE TİC. TÜR","ETİ ALUMİNUM","ETİ BAKIR","ETI MADEN","FATER TEMİZLİK ÜRÜNLERİ LTD.ŞTİ.","PG TÜKETİM MALLARI SAN. A.Ş.","SBS KIMYA","BAŞTÜRK CAM SAN. VE TİC.A.Ş. (USD)","VOTORANTİM ÇİMENTO","LİMAK ÇİMENTO","AG CİNER","Other"];

function Cashflow({canEdit,sm,sy,fx}:{canEdit:boolean,sm:number,sy:number,fx:typeof FX_D}){
  const fxRate=fx.usd_try;
  const[ov,setOv]=useState<Record<string,Record<number,string>>>({});
  const[notes,setNotes]=useState<Record<string,string>>({});
  const[customColl,setCustomColl]=useState<{id:string,label:string}[]>([]);
  const[customPay,setCustomPay]=useState<{id:string,label:string}[]>([]);
  const[addingTo,setAddingTo]=useState<"coll"|"pay"|null>(null);
  const[newRowLabel,setNewRowLabel]=useState("");
  const biz=getBiz(sm,sy);
  // Sub-sheet feeding
  const[subD,setSubD]=useState<Record<string,Record<string,string>>>({});
  const FEED_MAP:{cfId:string,pfx:string,firms:string[],cols:string[]}[]=[
    {cfId:"uk_export",pfx:"euk",firms:UK_F,cols:["EUR","USD"]},
    {cfId:"swe_recv",pfx:"eswe",firms:SWE_F,cols:["EUR","USD"]},
    {cfId:"us_export",pfx:"eus",firms:US_F,cols:["USD"]},
    {cfId:"dom_cash",pfx:"edom",firms:DOM_F,cols:["EUR","USD","TRY"]},
  ];
  useEffect(()=>{
    const reload=()=>{
      setOv(ldD("cf",sm,sy));setNotes(ldD("notes",sm,sy));
      const sd:Record<string,Record<string,string>>={};FEED_MAP.forEach(m=>sd[m.pfx]=ldD(m.pfx,sm,sy));setSubD(sd);
      try{const cc=localStorage.getItem("cf_custom_coll");if(cc)setCustomColl(JSON.parse(cc));else setCustomColl([]);
      const cp=localStorage.getItem("cf_custom_pay");if(cp)setCustomPay(JSON.parse(cp));else setCustomPay([]);}catch{}
    };
    reload();
    const onStorage=(e:StorageEvent)=>{if(!e.key||e.key.startsWith("cf_")||e.key.startsWith("notes_")||e.key.startsWith("euk_")||e.key.startsWith("eswe_")||e.key.startsWith("eus_")||e.key.startsWith("edom_"))reload();};
    window.addEventListener("storage",onStorage);
    return()=>window.removeEventListener("storage",onStorage);
  },[sm,sy]);

  const addRow=(section:"coll"|"pay")=>{
    if(!newRowLabel.trim())return;
    const id=(section==="coll"?"custColl_":"custPay_")+Date.now();
    const newRow={id,label:newRowLabel.trim()};
    if(section==="coll"){const nc=[...customColl,newRow];setCustomColl(nc);try{trackedSetItem("cf_custom_coll",JSON.stringify(nc));}catch{}}
    else{const nc=[...customPay,newRow];setCustomPay(nc);try{trackedSetItem("cf_custom_pay",JSON.stringify(nc));}catch{}}
    setNewRowLabel("");setAddingTo(null);
  };
  const removeRow=(section:"coll"|"pay",id:string)=>{
    if(!confirm("Bu satırı silmek istediğinize emin misiniz?"))return;
    if(section==="coll"){const nc=customColl.filter(r=>r.id!==id);setCustomColl(nc);try{trackedSetItem("cf_custom_coll",JSON.stringify(nc));}catch{}}
    else{const nc=customPay.filter(r=>r.id!==id);setCustomPay(nc);try{trackedSetItem("cf_custom_pay",JSON.stringify(nc));}catch{}}
  };

  const allColl=[...COLL,...customColl];
  const allPay=[...PAY,...customPay];

  const subHasData=(pfx:string)=>{const d=subD[pfx];if(!d)return false;return Object.values(d).some(v=>v!==""&&parseFloat(v)!==0);};
  const subDayTL=(m:{pfx:string,firms:string[],cols:string[]},di:number)=>{const d=subD[m.pfx];if(!d)return 0;let t=0;m.firms.forEach(f=>{m.cols.forEach(c=>{const v=parseFloat(d[f+"__"+di+"__"+c.toLowerCase()]||"")||0;if(c==="EUR")t+=v*fx.eur_try;else if(c==="USD")t+=v*fx.usd_try;else t+=v;});});return t;};
  const gV=(id:string,d:number)=>{
    const fm=FEED_MAP.find(m=>m.cfId===id);
    if(fm&&subHasData(fm.pfx)&&!(ov[id]?.[d]!==undefined&&ov[id][d]!==""))return subDayTL(fm,d);
    if(ov[id]?.[d]!==undefined&&ov[id][d]!=="")return parseFloat(stripC(ov[id][d]))||0;
    return 0;};
  const sV=(id:string,d:number,v:string)=>{if(!canEdit)return;setOv(p=>{const n={...p,[id]:{...(p[id]||{}),[d]:v}};svD("cf",sm,sy,n);return n;});};
  const isFed=(id:string)=>{const fm=FEED_MAP.find(m=>m.cfId===id);return fm&&subHasData(fm.pfx);};

  // Compute previous month's last-day closing balance TL
  // This chains months: April closing → May beginning, May closing → June beginning, etc.
  const prevMonthClosingTL=():number=>{
    const prevM=sm===0?11:sm-1;const prevY=sm===0?sy-1:sy;
    const prevOv=ldD("cf",prevM,prevY);const prevSubs:Record<string,Record<string,string>>={};
    FEED_MAP.forEach(m=>{prevSubs[m.pfx]=ldD(m.pfx,prevM,prevY);});
    const prevBiz=getBiz(prevM,prevY);if(prevBiz.length===0)return 0;
    // Read prev month cf values — same format as current
    const pGV=(id:string,d:number):number=>{
      const fm=FEED_MAP.find(m=>m.cfId===id);
      if(fm){const pS=prevSubs[fm.pfx];if(pS){const hasAny=Object.values(pS).some(v=>v!==""&&parseFloat(v)!==0);if(hasAny){const manualVal=prevOv[id]?.[d];if(!(manualVal!==undefined&&manualVal!=="")){let t=0;fm.firms.forEach(f=>{fm.cols.forEach(c=>{const v=parseFloat(pS[f+"__"+d+"__"+c.toLowerCase()]||"")||0;if(c==="EUR")t+=v*fx.eur_try;else if(c==="USD")t+=v*fx.usd_try;else t+=v;});});return t;}}}}
      const v=prevOv[id]?.[d];if(v!==undefined&&v!=="")return parseFloat(stripC(v))||0;
      return 0;
    };
    // Compute closing day by day, starting from prev month's own beginning balance
    let bal=parseFloat(stripC(prevOv["begin_tl"]?.[0]||""))||0;
    for(let i=0;i<prevBiz.length;i++){
      const c=allColl.reduce((s,r)=>s+pGV(r.id,i),0);
      const p=allPay.reduce((s,r)=>s+pGV(r.id,i),0);
      bal=bal+c-p;
    }
    return bal;
  };

  // Begin balance logic:
  // - Day 0: manual entry, OR auto = previous month's last-day closing balance
  // - Day N: = closing balance of Day N-1
  const beginTLRaw=(di:number)=>{const v=ov["begin_tl"]?.[di];if(v!==undefined&&v!=="")return parseFloat(stripC(v))||0;return 0;};
  const beginTLDay0=():number=>{const manual=ov["begin_tl"]?.[0];if(manual!==undefined&&manual!=="")return parseFloat(stripC(manual))||0;return prevMonthClosingTL();};
  const beginTL=(di:number):number=>di===0?beginTLDay0():closingTL(di-1);
  const beginUSD=(di:number):number=>beginTL(di)/fxRate;

  // Totals
  const tC=(di:number)=>allColl.reduce((s,r)=>s+gV(r.id,di),0);
  const tP=(di:number)=>allPay.reduce((s,r)=>s+gV(r.id,di),0);
  const inflowTL=(di:number):number=>beginTL(di)+tC(di);
  const inflowUSD=(di:number):number=>inflowTL(di)/fxRate;
  const grandTL=(di:number):number=>inflowTL(di)-tP(di);
  const grandUSD=(di:number):number=>grandTL(di)/fxRate;
  const closingTL=(di:number):number=>grandTL(di);
  const closingUSD=(di:number):number=>closingTL(di)/fxRate;
  const rowTot=(id:string)=>biz.reduce((s,_,i)=>s+gV(id,i),0);

  return<div>
    <div className="flex justify-between items-center mb-2">
      <div className="text-[10px] text-gray-400">Sağ tıkla → not ekle</div>
      <ExportBtn tableId="cfTable" name={"Cashflow_"+MONTHS[sm]+"_"+sy}/>
    </div>
    <div className="overflow-x-auto"><table id="cfTable" className="border-collapse text-[11px] w-full" style={{minWidth:220+(biz.length+2)*100+"px"}}>
      <thead><tr className="border-b border-gray-200"><th className={stL+" text-left font-semibold"} style={{minWidth:200}}>Kalem</th>{biz.map((d,i)=><th key={i} className="text-right px-1 py-1 text-[10px] whitespace-nowrap">{d}</th>)}<th className="text-right px-2 py-1 text-[10px] bg-gray-100 font-bold">Ay TL</th><th className="text-right px-2 py-1 text-[10px] bg-gray-100 font-bold">Ay USD</th></tr></thead>
      <tbody>
        {/* BEGINNING BALANCE TL — Day 0 editable (auto from prev month closing), rest = previous day closing */}
        <tr className="bg-indigo-50 border-b border-indigo-200"><td className={stL+" font-bold bg-indigo-50 text-indigo-800"}>Beginning Balance TL{(()=>{const manualD0=ov["begin_tl"]?.[0];const isAuto=!(manualD0!==undefined&&manualD0!=="");return isAuto&&beginTL(0)!==0?<span className="ml-1 text-[8px] text-cyan-600 font-semibold" title={"Önceki ayın ("+MONTHS[sm===0?11:sm-1]+") kapanış TL bakiyesinden geliyor"}>← prev ay</span>:null;})()}</td>{biz.map((_,i)=><td key={i} className="px-0.5 py-0.5">{i===0?<NoteCell id="begin_tl" di={i} value={beginTL(i)} ed={canEdit} onChange={v=>sV("begin_tl",i,v)} notes={notes} setNotes={setNotes} sm={sm} sy={sy}/>:<span className={"block text-right px-1 "+mn+" text-indigo-700 font-semibold"} title="Önceki günün Closing Balance TL'si">{fmt(beginTL(i))}</span>}</td>)}<td className={"text-right px-2 bg-indigo-100 "+mn+" font-bold"}>{fmt(beginTL(0))}</td><td className={"text-right px-2 bg-indigo-100 "+mn+" font-bold text-gray-500"}>{fmt(beginTL(0)/fxRate)}</td></tr>
        {/* BEGINNING BALANCE USD — Day 0 editable in USD, rest = derived */}
        <tr className="bg-indigo-50/50 border-b-2 border-indigo-300"><td className={stL+" font-bold bg-indigo-50/50 text-indigo-700"}>Beginning Balance USD</td>{biz.map((_,i)=><td key={i} className={"text-right px-1 "+mn+" text-indigo-600 font-semibold"}>{fmt(beginUSD(i))}</td>)}<td className="bg-indigo-100"></td><td className={"text-right px-2 bg-indigo-100 "+mn+" font-bold text-gray-500"}>{fmt(beginUSD(0))}</td></tr>

        <tr><td colSpan={biz.length+3}><Sec color="#dcfce7">Tahsilatlar ({allColl.length}){canEdit&&<button onClick={()=>setAddingTo("coll")} className="ml-3 px-2 py-0.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[10px] font-semibold">+ Satır Ekle</button>}</Sec></td></tr>
        {addingTo==="coll"&&<tr><td colSpan={biz.length+3} className="p-2 bg-emerald-50/30"><div className="inline-flex items-center gap-2"><input autoFocus value={newRowLabel} onChange={e=>setNewRowLabel(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addRow("coll")} placeholder="Yeni kalem adı" className="px-2 py-1 border border-gray-300 rounded text-xs outline-none" style={{minWidth:200}}/><button onClick={()=>addRow("coll")} className="px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-xs font-semibold">Ekle</button><button onClick={()=>{setAddingTo(null);setNewRowLabel("");}} className="px-2 py-1 text-gray-400 hover:text-gray-600 text-xs">İptal</button></div></td></tr>}
        {allColl.map(r=>{const rt=rowTot(r.id);const fed=isFed(r.id);const isCust=customColl.some(c=>c.id===r.id);return<tr key={r.id} className={"border-b border-gray-100 "+(fed?"bg-cyan-50/40":"")}><td className={stL}>{isCust&&<span className="text-[8px] text-blue-500 font-bold mr-1">★</span>}{r.label}{fed&&<span className="ml-1 text-[8px] text-cyan-600 font-semibold" title="Sub-sheet'ten besleniyor">⟵</span>}{isCust&&canEdit&&<button onClick={()=>removeRow("coll",r.id)} className="ml-2 text-red-400 hover:text-red-600 text-[10px]" title="Sil">×</button>}</td>{biz.map((_,i)=><td key={i} className="px-0.5 py-0.5"><NoteCell id={r.id} di={i} value={gV(r.id,i)} ed={canEdit&&!fed} onChange={v=>sV(r.id,i,v)} notes={notes} setNotes={setNotes} sm={sm} sy={sy}/></td>)}<td className={"text-right px-2 bg-gray-50 "+mn+" font-semibold"}>{fmt(rt)}</td><td className={"text-right px-2 bg-gray-50 "+mn+" text-gray-400"}>{fmt(rt/fxRate)}</td></tr>;})}

        {/* Total TL (of inflows only, excl begin) */}
        <tr className="bg-emerald-50 font-semibold"><td className={stL+" font-semibold bg-emerald-50"}>Total TL</td>{biz.map((_,i)=><td key={i} className={"text-right px-1 "+mn}>{fmt(tC(i))}</td>)}<td className={"text-right px-2 bg-emerald-100 "+mn+" font-bold"}>{fmt(biz.reduce((s,_,i)=>s+tC(i),0))}</td><td className={"text-right px-2 bg-emerald-100 "+mn+" font-bold text-gray-500"}>{fmt(biz.reduce((s,_,i)=>s+tC(i),0)/fxRate)}</td></tr>
        {/* Total USD */}
        <tr className="bg-emerald-50/70 border-b-2 border-emerald-300"><td className={stL+" font-semibold bg-emerald-50/70 text-gray-500"}>Total USD</td>{biz.map((_,i)=><td key={i} className={"text-right px-1 "+mn+" text-gray-500"}>{fmt(tC(i)/fxRate)}</td>)}<td className="bg-emerald-100"></td><td className={"text-right px-2 bg-emerald-100 "+mn+" font-bold text-gray-500"}>{fmt(biz.reduce((s,_,i)=>s+tC(i)/fxRate,0))}</td></tr>

        {/* TOTAL INFLOW TL */}
        <tr className="bg-blue-100 font-bold border-t-2 border-blue-400"><td className={stL+" font-bold bg-blue-100 text-blue-900"}>TOTAL INFLOW TL</td>{biz.map((_,i)=><td key={i} className={"text-right px-1 "+mn+" font-bold text-blue-900"}>{fmt(inflowTL(i))}</td>)}<td className={"text-right px-2 bg-blue-200 "+mn+" font-bold text-blue-900"}>{fmt(biz.reduce((s,_,i)=>s+inflowTL(i),0))}</td><td className={"text-right px-2 bg-blue-200 "+mn+" font-bold text-blue-800"}>{fmt(biz.reduce((s,_,i)=>s+inflowTL(i),0)/fxRate)}</td></tr>
        {/* TOTAL INFLOW USD */}
        <tr className="bg-blue-50 font-bold border-b-2 border-blue-400"><td className={stL+" font-bold bg-blue-50 text-blue-800"}>TOTAL INFLOW USD</td>{biz.map((_,i)=><td key={i} className={"text-right px-1 "+mn+" font-bold text-blue-800"}>{fmt(inflowUSD(i))}</td>)}<td className="bg-blue-200"></td><td className={"text-right px-2 bg-blue-200 "+mn+" font-bold text-blue-800"}>{fmt(biz.reduce((s,_,i)=>s+inflowUSD(i),0))}</td></tr>

        <tr><td colSpan={biz.length+3}><Sec color="#fee2e2">Ödemeler ({allPay.length}){canEdit&&<button onClick={()=>setAddingTo("pay")} className="ml-3 px-2 py-0.5 bg-red-600 hover:bg-red-700 text-white rounded text-[10px] font-semibold">+ Satır Ekle</button>}</Sec></td></tr>
        {addingTo==="pay"&&<tr><td colSpan={biz.length+3} className="p-2 bg-red-50/30"><div className="inline-flex items-center gap-2"><input autoFocus value={newRowLabel} onChange={e=>setNewRowLabel(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addRow("pay")} placeholder="Yeni kalem adı" className="px-2 py-1 border border-gray-300 rounded text-xs outline-none" style={{minWidth:200}}/><button onClick={()=>addRow("pay")} className="px-3 py-1 bg-red-600 hover:bg-red-700 text-white rounded text-xs font-semibold">Ekle</button><button onClick={()=>{setAddingTo(null);setNewRowLabel("");}} className="px-2 py-1 text-gray-400 hover:text-gray-600 text-xs">İptal</button></div></td></tr>}
        {allPay.map(r=>{const rt=rowTot(r.id);const isCust=customPay.some(c=>c.id===r.id);return<tr key={r.id} className="border-b border-gray-100"><td className={stL}>{isCust&&<span className="text-[8px] text-blue-500 font-bold mr-1">★</span>}{r.label}{isCust&&canEdit&&<button onClick={()=>removeRow("pay",r.id)} className="ml-2 text-red-400 hover:text-red-600 text-[10px]" title="Sil">×</button>}</td>{biz.map((_,i)=><td key={i} className="px-0.5 py-0.5"><NoteCell id={r.id} di={i} value={gV(r.id,i)} ed={canEdit} onChange={v=>sV(r.id,i,v)} notes={notes} setNotes={setNotes} sm={sm} sy={sy}/></td>)}<td className={"text-right px-2 bg-gray-50 "+mn+" font-semibold"}>{fmt(rt)}</td><td className={"text-right px-2 bg-gray-50 "+mn+" text-gray-400"}>{fmt(rt/fxRate)}</td></tr>;})}

        {/* Grand Total TL */}
        <tr className="bg-purple-100 font-bold border-t-2 border-purple-400"><td className={stL+" font-bold bg-purple-100 text-purple-900"}>Grand Total TL</td>{biz.map((_,i)=><td key={i} className={"text-right px-1 "+mn+" font-bold "+(grandTL(i)<0?"text-red-600":"text-purple-900")}>{fmt(grandTL(i))}</td>)}<td className={"text-right px-2 bg-purple-200 "+mn+" font-bold text-purple-900"}>{fmt(biz.reduce((s,_,i)=>s+grandTL(i),0))}</td><td className={"text-right px-2 bg-purple-200 "+mn+" font-bold text-purple-800"}>{fmt(biz.reduce((s,_,i)=>s+grandTL(i),0)/fxRate)}</td></tr>
        {/* Grand Total USD */}
        <tr className="bg-purple-50 font-bold border-b border-purple-400"><td className={stL+" font-bold bg-purple-50 text-purple-800"}>Grand Total USD</td>{biz.map((_,i)=><td key={i} className={"text-right px-1 "+mn+" font-bold "+(grandUSD(i)<0?"text-red-500":"text-purple-800")}>{fmt(grandUSD(i))}</td>)}<td className="bg-purple-200"></td><td className={"text-right px-2 bg-purple-200 "+mn+" font-bold text-purple-800"}>{fmt(biz.reduce((s,_,i)=>s+grandUSD(i),0))}</td></tr>

        {/* Closing Balance TL */}
        <tr className="bg-indigo-200 font-bold border-t-2 border-indigo-500"><td className={stL+" font-bold bg-indigo-200 text-indigo-950"}>Closing Balance TL</td>{biz.map((_,i)=><td key={i} className={"text-right px-1 "+mn+" font-bold "+(closingTL(i)<0?"text-red-700":"text-indigo-950")}>{fmt(closingTL(i))}</td>)}<td className={"text-right px-2 bg-indigo-300 "+mn+" font-bold text-indigo-950"}>{fmt(biz.reduce((s,_,i)=>s+closingTL(i),0))}</td><td className={"text-right px-2 bg-indigo-300 "+mn+" font-bold text-indigo-900"}>{fmt(biz.reduce((s,_,i)=>s+closingTL(i),0)/fxRate)}</td></tr>
        {/* Closing Balance USD */}
        <tr className="bg-indigo-100 font-bold"><td className={stL+" font-bold bg-indigo-100 text-indigo-900"}>Closing Balance USD</td>{biz.map((_,i)=><td key={i} className={"text-right px-1 "+mn+" font-bold "+(closingUSD(i)<0?"text-red-600":"text-indigo-900")}>{fmt(closingUSD(i))}</td>)}<td className="bg-indigo-300"></td><td className={"text-right px-2 bg-indigo-300 "+mn+" font-bold text-indigo-900"}>{fmt(biz.reduce((s,_,i)=>s+closingUSD(i),0))}</td></tr>
      </tbody></table></div></div>;
}

// Sub-sheet firm arrays for feeding

type ST="uk"|"us"|"dom";
const SCFG:Record<ST,{title:string,tables:{pfx:string,firms:string[],label:string,cf:string,clr:string,cols:string[]}[]}> = {uk:{title:"Export",tables:[{pfx:"euk",firms:UK_F,label:"UK Export Receivables",cf:"UK Export Receivables",clr:"#dcfce7",cols:["EUR","USD"]},{pfx:"eswe",firms:SWE_F,label:"SW Europe Receivables",cf:"SW Europe Receivables",clr:"#e0f2fe",cols:["EUR","USD"]}]},us:{title:"US - Export",tables:[{pfx:"eus",firms:US_F,label:"US Export Receivables",cf:"US Export Receivables",clr:"#dcfce7",cols:["USD"]}]},dom:{title:"Domestic",tables:[{pfx:"edom",firms:DOM_F,label:"Domestic Collection",cf:"Domestic Collection Cash",clr:"#fef9c3",cols:["EUR","USD","TRY"]}]}};
const CC:Record<string,string>={"EUR":"text-blue-600","USD":"text-green-600","TRY":"text-amber-700"};
const CB:Record<string,string>={"EUR":"bg-blue-50/50 text-blue-700","USD":"bg-green-50/50 text-green-700","TRY":"bg-amber-50/50 text-amber-700"};
function SubSheet({type,sm,sy,fx,canEdit}:{type:ST,sm:number,sy:number,fx:typeof FX_D,canEdit:boolean}){const cfg=SCFG[type];const biz=getBiz(sm,sy);
  const[datas,setDatas]=useState<Record<string,Record<string,string>>>({});
  const[col,setCol]=useState<Record<string,boolean>>(()=>{const o:Record<string,boolean>={};cfg.tables.forEach((t,i)=>o[t.pfx]=i>0);return o;});
  // Custom firms per prefix
  const[customFirms,setCustomFirms]=useState<Record<string,string[]>>({});
  // Deleted default firms per prefix
  const[deletedFirms,setDeletedFirms]=useState<Record<string,string[]>>({});
  const[addingFor,setAddingFor]=useState<string|null>(null);
  const[newFirmName,setNewFirmName]=useState("");
  useEffect(()=>{
    const reload=()=>{
      const d:Record<string,Record<string,string>>={};cfg.tables.forEach(t=>d[t.pfx]=ldD(t.pfx,sm,sy));setDatas(d);
      try{const cf=localStorage.getItem("sub_custom_firms");if(cf)setCustomFirms(JSON.parse(cf));
      const df=localStorage.getItem("sub_deleted_firms");if(df)setDeletedFirms(JSON.parse(df));}catch{}
    };
    reload();
    const onStorage=(e:StorageEvent)=>{if(!e.key||cfg.tables.some(t=>e.key?.startsWith(t.pfx+"_"))||e.key==="sub_custom_firms"||e.key==="sub_deleted_firms")reload();};
    window.addEventListener("storage",onStorage);
    return()=>window.removeEventListener("storage",onStorage);
  },[sm,sy,type]);
  const getFirms=(pfx:string,defaults:string[])=>{const extra=customFirms[pfx]||[];const del=deletedFirms[pfx]||[];return[...defaults.filter(f=>!del.includes(f)),...extra];};
  const addFirm=(pfx:string)=>{if(!newFirmName.trim())return;setCustomFirms(p=>{const n={...p,[pfx]:[...(p[pfx]||[]),newFirmName.trim()]};try{trackedSetItem("sub_custom_firms",JSON.stringify(n));}catch{}return n;});setNewFirmName("");setAddingFor(null);};
  const removeFirm=(pfx:string,firm:string,defaults:string[])=>{if(!confirm(`"${firm}" firmasını silmek istediğinize emin misiniz?`))return;
    if(defaults.includes(firm)){setDeletedFirms(p=>{const n={...p,[pfx]:[...(p[pfx]||[]),firm]};try{trackedSetItem("sub_deleted_firms",JSON.stringify(n));}catch{}return n;});}
    else{setCustomFirms(p=>{const n={...p,[pfx]:(p[pfx]||[]).filter(f=>f!==firm)};try{trackedSetItem("sub_custom_firms",JSON.stringify(n));}catch{}return n;});}
  };
  const gV=(pfx:string,f:string,d:number,c:string)=>(datas[pfx]||{})[f+"__"+d+"__"+c.toLowerCase()]||"";
  const gN=(pfx:string,f:string,d:number,c:string)=>parseFloat(gV(pfx,f,d,c))||0;
  const sV=(pfx:string,f:string,d:number,c:string,v:string)=>{const k=f+"__"+d+"__"+c.toLowerCase();setDatas(prev=>{const nd={...(prev[pfx]||{}),[k]:v};svD(pfx,sm,sy,nd);return{...prev,[pfx]:nd};});};
  const dTot=(pfx:string,firms:string[],d:number,c:string)=>firms.reduce((s,f)=>s+gN(pfx,f,d,c),0);
  const dTL=(pfx:string,firms:string[],cols:string[],d:number)=>{let t=0;if(cols.includes("EUR"))t+=dTot(pfx,firms,d,"eur")*fx.eur_try;if(cols.includes("USD"))t+=dTot(pfx,firms,d,"usd")*fx.usd_try;if(cols.includes("TRY"))t+=dTot(pfx,firms,d,"try");return t;};
  const firmUSD=(pfx:string,f:string,cols:string[])=>{let t=0;for(let di=0;di<biz.length;di++){if(cols.includes("USD"))t+=gN(pfx,f,di,"usd");if(cols.includes("EUR"))t+=gN(pfx,f,di,"eur")*fx.eur_usd;if(cols.includes("TRY"))t+=gN(pfx,f,di,"try")/fx.usd_try;}return t;};
  const totalMonthUSD=(pfx:string,firms:string[],cols:string[])=>firms.reduce((s,f)=>s+firmUSD(pfx,f,cols),0);
  const tblId="sub_"+type;
  return<div><div className="flex justify-between items-center mb-4"><span className="text-sm font-bold">{cfg.title}</span><ExportBtn tableId={tblId} name={cfg.title+"_"+MONTHS[sm]+"_"+sy}/></div>
  {cfg.tables.map(tbl=>{const ic=col[tbl.pfx];const nc=tbl.cols.length;const firms=getFirms(tbl.pfx,tbl.firms);return<div key={tbl.pfx} className="mb-6">
    <div className="flex items-center gap-3 mb-2">
      <span className="text-sm font-semibold cursor-pointer select-none" onClick={()=>setCol(p=>({...p,[tbl.pfx]:!p[tbl.pfx]}))}>{ic?"▶":"▼"} {tbl.label}</span>
      <span className="text-[10px] px-2 py-0.5 rounded-full" style={{background:tbl.clr}}>{firms.length}</span>
      {canEdit&&!ic&&<button onClick={()=>setAddingFor(tbl.pfx)} className="px-2 py-0.5 bg-blue-100 hover:bg-blue-200 text-blue-700 rounded text-[10px] font-semibold">+ Customer</button>}
    </div>
    {addingFor===tbl.pfx&&<div className="mb-2 p-2 bg-white border border-blue-200 rounded-lg inline-flex items-center gap-2"><input autoFocus value={newFirmName} onChange={e=>setNewFirmName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addFirm(tbl.pfx)} placeholder="Yeni müşteri adı" className="px-2 py-1 border border-gray-300 rounded text-xs outline-none" style={{minWidth:200}}/><button onClick={()=>addFirm(tbl.pfx)} className="px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-xs font-semibold">Ekle</button><button onClick={()=>{setAddingFor(null);setNewFirmName("");}} className="px-2 py-1 text-gray-400 hover:text-gray-600 text-xs">İptal</button></div>}
    {!ic&&<div className="overflow-x-auto"><table id={tblId} className="border-collapse text-[10px] w-full" style={{minWidth:360+biz.length*nc*65+"px"}}>
      <thead><tr className="border-b-2 border-gray-300">
        <th className="sticky left-0 bg-white z-10 px-1 py-1" style={{minWidth:30}}>#</th>
        <th className="sticky bg-white z-10 px-1 py-1" style={{minWidth:200,left:30}}>Customer</th>
        {biz.map((d,i)=><th key={i} colSpan={nc} className="text-center px-0.5 py-1 text-[9px]">{d}</th>)}
        <th className="text-right px-2 py-1 text-[9px] bg-indigo-50 text-indigo-700 font-bold whitespace-nowrap" style={{minWidth:80}}>Ay USD</th>
        {canEdit&&<th className="w-6"></th>}
      </tr>{nc>1&&<tr className="border-b border-gray-200"><th className="sticky left-0 bg-white"></th><th className="sticky bg-white" style={{left:30}}></th>{biz.map((_,i)=>tbl.cols.map(c=><th key={i+c} className={"text-center px-0.5 text-[9px] "+CC[c]}>{c}</th>))}<th className="bg-indigo-50"></th>{canEdit&&<th></th>}</tr>}</thead>
      <tbody>{firms.map((f,fi)=>{const fUsd=firmUSD(tbl.pfx,f,tbl.cols);const isCustom=(customFirms[tbl.pfx]||[]).includes(f);return<tr key={fi} className="border-b border-gray-100 hover:bg-gray-50">
        <td className="sticky left-0 bg-white z-10 px-1 text-[9px] text-gray-400">{fi+1}</td>
        <td className="sticky bg-white z-10 px-1 text-[10px] font-medium truncate" style={{left:30,maxWidth:200}} title={f}>{isCustom&&<span className="text-[8px] text-blue-500 font-bold mr-1">★</span>}{f}</td>
        {biz.map((_,di)=>tbl.cols.map(c=><td key={di+c} className="px-0.5 py-0.5"><input className={"w-full px-0.5 py-0.5 border border-gray-200 rounded text-right "+mn+" "+CB[c]+" outline-none text-[10px]"} style={{minWidth:55}} value={gV(tbl.pfx,f,di,c)} onChange={e=>sV(tbl.pfx,f,di,c,e.target.value)} placeholder="-"/></td>))}
        <td className={"text-right px-2 bg-indigo-50/50 "+mn+" font-semibold text-indigo-700"}>{fUsd?fmt(fUsd):"-"}</td>
        {canEdit&&<td className="px-1"><button onClick={()=>removeFirm(tbl.pfx,f,tbl.firms)} className="px-1 py-0.5 text-red-400 hover:bg-red-50 hover:text-red-600 rounded text-[11px]" title="Firmayı sil">×</button></td>}
      </tr>;})}
      <tr className="bg-gray-100 font-bold border-t-2 border-gray-300"><td className="sticky left-0 bg-gray-100"></td><td className="sticky bg-gray-100 font-bold text-[10px]" style={{left:30}}>TOTAL</td>{biz.map((_,di)=>tbl.cols.map(c=><td key={di+c} className={"text-right px-1 "+mn+" "+CC[c]}>{fmt(dTot(tbl.pfx,firms,di,c))}</td>))}<td className={"text-right px-2 bg-indigo-100 "+mn+" font-bold text-indigo-800"}>{fmt(totalMonthUSD(tbl.pfx,firms,tbl.cols))}</td>{canEdit&&<td></td>}</tr>
      <tr className="bg-amber-50 font-bold"><td className="sticky left-0 bg-amber-50"></td><td className="sticky bg-amber-50 font-bold text-[10px]" style={{left:30}}>TOTAL TL</td>{biz.map((_,di)=><td key={di} colSpan={nc} className={"text-right px-1 "+mn+" font-bold text-amber-700"}>{fmt(dTL(tbl.pfx,firms,tbl.cols,di))}</td>)}<td className="bg-amber-50"></td>{canEdit&&<td></td>}</tr>
      </tbody></table></div>}{ic&&<div className="text-[10px] text-gray-400 px-2 py-1">Tıklayarak açın — {firms.length} firma</div>}</div>;})}</div>;
}

function Kasa({canEdit,sm,sy,fx}:{canEdit:boolean,sm:number,sy:number,fx:typeof FX_D}){
  const biz=getBiz(sm,sy);const conv=fx.usd_try*1e6; // TL->USD mio divisor
  const[d,setD]=useState<Record<string,string>>({});
  const[exp,setExp]=useState<Record<string,boolean>>({});
  const[notes,setNotes]=useState<Record<string,string>>({});
  const[cfData,setCfData]=useState<Record<string,string>>({});
  const[refreshTick,setRefreshTick]=useState(0);
  const kk="kasa2_"+sy+"_"+sm;
  // Reload kasa own data + cashflow data when sm/sy change or when window gets focus
  useEffect(()=>{
    const reload=()=>{
      try{const s=localStorage.getItem(kk);setD(s?JSON.parse(s):{});
        setNotes(ldD("kasa_notes",sm,sy));
        setCfData(ldD("cf",sm,sy));}catch{setD({});setCfData({});}
    };
    reload();
    // Listen for storage events (cross-tab) AND window focus (same-tab updates)
    const onStorage=(e:StorageEvent)=>{if(e.key&&(e.key.startsWith("cf_")||e.key===kk))reload();};
    const onFocus=()=>reload();
    window.addEventListener("storage",onStorage);
    window.addEventListener("focus",onFocus);
    // Also poll every 2s so switching tabs within the app refreshes
    const poll=setInterval(reload,2000);
    return()=>{window.removeEventListener("storage",onStorage);window.removeEventListener("focus",onFocus);clearInterval(poll);};
  },[sm,sy,refreshTick]);
  const sv=(k:string,v:string)=>{setD(p=>{const n={...p,[k]:v};try{trackedSetItem(kk,JSON.stringify(n));}catch{}return n;});};
  const gv=(k:string)=>d[k]||"";const gn=(k:string)=>parseFloat(d[k]||"")||0;

  // Cashflow data (reactive via cfData state)
  const cfV=(id:string,di:number)=>{
    // cf stored as nested: {id: {di: val}} - check format
    const entry=cfData[id];
    if(entry&&typeof entry==="object"){const v=(entry as any)[di];return parseFloat(stripC(v||""))||0;}
    // fallback flat format
    return parseFloat(stripC(cfData[id+"_"+di]||""))||0;
  };
  const cfToMio=(v:number)=>v/conv; // TL to USD mio

  // TR Tahsilat from Cashflow
  const TR_TAH_IDS=["dom_cash","dom_cheque","electricity_c","interest_tr","interest_uk","vat_refund","rent_kazan"];
  const trTah=(di:number)=>cfToMio(TR_TAH_IDS.reduce((s,id)=>s+cfV(id,di),0));
  // Loan Utilisation
  const loanUtil=(di:number)=>cfToMio(cfV("loan_util",di));
  // TR Ödemeler from Cashflow
  const TR_OD_IDS=["loan_rep_tr","tcc_iet","eti_royalty","nat_gas","payroll","port","ssk","tax","siemens","epias","coal","transport","import_p","export_p","freight","fuel","trade_pay","elec_kazan","elec_eti","other_inv","maintenance"];
  const trOd=(di:number)=>cfToMio(TR_OD_IDS.reduce((s,id)=>s+cfV(id,di),0));
  // UK Export & US Export from Cashflow
  const ukExp=(di:number)=>cfToMio(cfV("uk_export",di));
  const usExp=(di:number)=>cfToMio(cfV("us_export",di));
  // UK Ödemeler: wesoda_pay + sw_reseller + we_ent + ciner_fund
  const UK_OD_IDS=["wesoda_pay","sw_reseller","we_ent","ciner_fund"];
  const ukOd=(di:number)=>cfToMio(UK_OD_IDS.reduce((s,id)=>s+cfV(id,di),0));
  // Ansac Transfer
  const ansacT=(di:number)=>cfToMio(cfV("ansac",di));
  // SWE Export = SW Europe Receivables
  const sweExp=(di:number)=>cfToMio(cfV("swe_recv",di));
  // SWE Ödemeler = SW Europe Freight
  const sweOd=(di:number)=>cfToMio(cfV("swe_freight",di));

  // TR calc: Açılış + UK'den Gelen + SWE'den Gelen + KDV İade + TR Tahsilat + Kredi Kullanımı - TR Ödemeler
  // Manual override helper: if manual value entered, use it; otherwise fall back to auto
  const mOv=(k:string,autoVal:number)=>{const v=gv(k);return v!==""?parseFloat(v)||0:autoVal;};
  const trKap=(di:number):number=>{const ac=di===0?gn("tr_acilis_init"):trKap(di-1);return ac+gn("tr_ukgelen_"+di)+gn("tr_swegelen_"+di)+gn("tr_kdviade_"+di)+mOv("tr_tahsilat_"+di,trTah(di))+mOv("tr_kredikullanim_"+di,loanUtil(di))-mOv("tr_odemeler_"+di,trOd(di));};
  const trAc=(di:number):number=>di===0?gn("tr_acilis_init"):trKap(di-1);
  // UK calc
  const ukKap=(di:number):number=>{const ac=di===0?gn("uk_acilis_init"):ukKap(di-1);return ac+mOv("uk_export_"+di,ukExp(di))+mOv("uk_usexport_"+di,usExp(di))-mOv("uk_odemeler_"+di,ukOd(di))-mOv("uk_ansac_"+di,ansacT(di))-gn("uk_trye_"+di)-gn("uk_sweye_"+di);};
  const ukAc=(di:number):number=>di===0?gn("uk_acilis_init"):ukKap(di-1);
  // SWE calc
  const sweKap=(di:number):number=>{const ac=di===0?gn("swe_acilis_init"):sweKap(di-1);return ac+gn("swe_ukgelen_"+di)+mOv("swe_export_"+di,sweExp(di))-mOv("swe_odemeler_"+di,sweOd(di))-gn("swe_trye_"+di);};
  const sweAc=(di:number):number=>di===0?gn("swe_acilis_init"):sweKap(di-1);
  const totalKap=(di:number)=>trKap(di)+ukKap(di)+sweKap(di);

  const ci="w-full px-1 py-0.5 border border-gray-200 rounded text-right text-[10px] outline-none bg-amber-50/60 text-blue-700 "+mn;
  const cv="text-right px-1 text-[10px] "+mn;
  const hdr="text-right px-1 py-1.5 text-[9px] whitespace-nowrap font-semibold";
  const lbl="sticky left-0 z-10 px-2 py-1 text-[10px] whitespace-nowrap bg-white";

  // Ödemeler detail item IDs for check validation
  const TR_OD_LABELS=["İthalat","İhracat","EPIAŞ","Akaryakıt","Kömür","Nakliye","Navlun","Satıcılar","Doğalgaz","Exim Kredi G.Ö."];
  const UK_OD_LABELS=["We Soda Ödemeler","Soda World Ödemeler"];
  const SWE_OD_LABELS=["SWE Navlun"];

  type RowDef={key:string,label:string,neg?:boolean,clr?:string,auto?:(di:number)=>number};

  const trRows:RowDef[]=[
    {key:"tr_ukgelen",label:"UK'den Gelen"},
    {key:"tr_swegelen",label:"SWE'den Gelen"},
    {key:"tr_kdviade",label:"KDV İade"},
    {key:"tr_tahsilat",label:"TR Tahsilat",clr:"green",auto:trTah},
    ...(biz.some((_,i)=>loanUtil(i)>0)?[{key:"tr_kredikullanim",label:"Kredi Kullanımı",clr:"green",auto:loanUtil}]:[]),
    {key:"tr_odemeler",label:"TR Ödemeler",neg:true,auto:trOd},
  ];
  const ukRows:RowDef[]=[
    {key:"uk_export",label:"Export",clr:"green",auto:ukExp},
    {key:"uk_usexport",label:"US Export",clr:"green",auto:usExp},
    {key:"uk_odemeler",label:"UK Ödemeler",neg:true,auto:ukOd},
    {key:"uk_ansac",label:"Ansac Transfer",neg:true,clr:"orange",auto:ansacT},
    {key:"uk_trye",label:"TR'ye Gönderilen",neg:true,clr:"orange"},
    {key:"uk_sweye",label:"SWE'ye gönderilen",neg:true,clr:"orange"},
  ];
  const sweRows:RowDef[]=[
    {key:"swe_ukgelen",label:"UK'den gelen"},
    {key:"swe_export",label:"Export",clr:"green",auto:sweExp},
    {key:"swe_odemeler",label:"SW GmbH Ödemeler",neg:true,auto:sweOd},
    {key:"swe_trye",label:"TR'ye Gönderilen",neg:true,clr:"orange"},
  ];

  type SecDef={id:string,title:string,color:string,acFn:(di:number)=>number,kapFn:(di:number)=>number,initKey:string,rows:RowDef[],odLabels:string[],odPrefix:string};
  const sections:SecDef[]=[
    {id:"tr",title:"TR Bölümü",color:"#dbeafe",acFn:trAc,kapFn:trKap,initKey:"tr_acilis_init",rows:trRows,odLabels:TR_OD_LABELS,odPrefix:"trod"},
    {id:"uk",title:"UK Bölümü",color:"#dcfce7",acFn:ukAc,kapFn:ukKap,initKey:"uk_acilis_init",rows:ukRows,odLabels:UK_OD_LABELS,odPrefix:"ukod"},
    {id:"swe",title:"Soda World GmbH",color:"#fef9c3",acFn:sweAc,kapFn:sweKap,initKey:"swe_acilis_init",rows:sweRows,odLabels:SWE_OD_LABELS,odPrefix:"sweod"},
  ];

  const renderSec=(sec:SecDef)=>{const odKey="exp_od_"+sec.id;
    const rowClr=(r:RowDef)=>r.clr==="green"?" text-emerald-700":r.clr==="orange"?" text-orange-600":r.neg?" text-red-600":"";
    const secLabel=sec.id==="swe"?"SW GmbH":sec.id==="tr"?"TR":"UK";
    const odSum=(di:number)=>sec.odLabels.reduce((s,_,ii)=>s+(parseFloat(gv(sec.odPrefix+"_"+ii+"_"+di))||0),0);
    const mainOdRow=sec.rows.find(r=>r.key.includes("odemeler"));
    const mainOdVal=(di:number)=>mainOdRow?.auto?mainOdRow.auto(di):gn(mainOdRow?.key+"_"+di);
    const chk=(di:number)=>{const os=odSum(di);const mo=mainOdVal(di);if(os===0&&mo===0)return 0;return mo-os;};
    return<div key={sec.id} className="mb-3"><table className="border-collapse text-[10px] w-full" style={{minWidth:180+biz.length*90+"px"}}>
      <thead><tr style={{background:sec.color}}><th className={lbl+" font-bold text-[11px]"} style={{minWidth:170,background:sec.color}}>{sec.title}</th>{biz.map((dd,i)=><th key={i} className={hdr}>{dd}</th>)}</tr></thead>
      <tbody>
        <tr className="bg-blue-50/60 font-semibold border-b border-gray-200"><td className={lbl+" font-bold bg-blue-50/60"}>{secLabel} Açılış</td>{biz.map((_,i)=><td key={i} className={cv+" font-bold"}>{i===0?<input className={ci+" bg-blue-100/50 font-bold"} value={gv(sec.initKey)} onChange={e=>sv(sec.initKey,e.target.value)} placeholder="0"/>:fmtM(sec.acFn(i))}</td>)}</tr>
        {sec.rows.map(r=><tr key={r.key} className="border-b border-gray-100"><td className={lbl+rowClr(r)}>{r.label}{r.auto&&<span className="ml-1 text-[7px] text-cyan-500 font-semibold" title="Cashflow'dan beslenir (manuel override mümkün)">CF</span>}</td>{biz.map((_,i)=>{
          const cfVal=r.auto?r.auto(i):0;
          const manual=gv(r.key+"_"+i);
          const hasManual=manual!==""&&manual!==undefined;
          const displayVal=hasManual?parseFloat(manual)||0:cfVal;
          const nk=r.key+"_"+i;const note=notes[nk]||"";
          return<td key={i} className="px-0.5 py-0.5 relative group" onContextMenu={e=>{e.preventDefault();const t=prompt("Not ekle/düzenle:",note);if(t===null)return;const nn={...notes};if(t.trim())nn[nk]=t.trim();else delete nn[nk];setNotes(nn);svD("kasa_notes",sm,sy,nn);}}>
          {note&&<div className="absolute top-0 right-0 w-0 h-0 z-20 pointer-events-none" style={{borderLeft:"5px solid transparent",borderTop:"5px solid #ef4444"}}/>}
          {note&&<div className="absolute z-30 bottom-full right-0 mb-1 px-2 py-1 bg-gray-800 text-white text-[10px] rounded shadow-lg max-w-[200px] whitespace-pre-wrap opacity-0 group-hover:opacity-100 pointer-events-none transition">{note}</div>}
          {canEdit?<input className={ci+rowClr(r)+(r.auto&&!hasManual?" bg-cyan-50/40 text-cyan-700":"")+(hasManual&&r.auto?" ring-1 ring-amber-400":"")} value={hasManual?manual:(r.auto?fmtM(cfVal)||"":"")} onChange={e=>sv(r.key+"_"+i,e.target.value)} placeholder={r.auto?fmtM(cfVal)||"-":"-"}/>
          :<span className={cv+(r.neg?" text-red-500":"")}>{fmtM(displayVal)}</span>}
        </td>;})}</tr>)}
        <tr className="bg-blue-50/60 font-bold border-t-2 border-b-2 border-gray-300"><td className={lbl+" font-bold bg-blue-50/60"}>{secLabel} Kapanış</td>{biz.map((_,i)=><td key={i} className={cv+" font-bold text-blue-800"}>{fmtM(sec.kapFn(i))}</td>)}</tr>
        <tr className="cursor-pointer hover:bg-gray-50" onClick={()=>setExp(p=>({...p,[odKey]:!p[odKey]}))}><td className={lbl+" text-gray-500 underline italic"} colSpan={biz.length+1}>{exp[odKey]?"▼":"▶"} Ödemeler detay</td></tr>
        {exp[odKey]&&<>{sec.odLabels.map((item,ii)=><tr key={ii} className="border-b border-gray-50"><td className={lbl+" pl-4 text-gray-500 text-[9px]"}>{item}</td>{biz.map((_,i)=><td key={i} className="px-0.5 py-0.5"><input className={ci+" bg-gray-50 text-gray-600 text-[9px]"} value={gv(sec.odPrefix+"_"+ii+"_"+i)} onChange={e=>sv(sec.odPrefix+"_"+ii+"_"+i,e.target.value)} placeholder="-"/></td>)}</tr>)}</>}
        <tr className="border-b border-gray-200 bg-gray-50/50"><td className={lbl+" text-gray-400 italic text-[9px]"}>Check</td>{biz.map((_,i)=>{const c=chk(i);return<td key={i} className={cv+" text-[9px] "+(c===0?"text-gray-300":Math.abs(c)<0.01?"text-emerald-500":"text-red-500 font-semibold")}>{c===0?"-":fmtM(c)}</td>;})}</tr>
      </tbody></table></div>;};

  return<div>
    <div className="flex justify-between items-center mb-3 flex-wrap gap-2">
      <span className="text-sm font-bold">TR / UK / SWE GmbH Kasa — USD milyon</span>
      <ExportBtn tableId="kasaFull" name={"Kasa_"+MONTHS[sm]+"_"+sy}/>
    </div>
    <div className="text-[10px] text-gray-400 mb-2"><span className="text-cyan-600 font-semibold">CF</span> = Cashflow\'dan otomatik beslenir</div>
    <div id="kasaFull" className="overflow-x-auto space-y-1">
      {sections.map(renderSec)}
      <table className="border-collapse text-[10px] w-full" style={{minWidth:180+biz.length*90+"px"}}><tbody>
        <tr className="bg-indigo-100 font-bold border-t-2 border-indigo-300"><td className={lbl+" font-bold bg-indigo-100"} style={{minWidth:170}}>TR-UK-DE Kapanış</td>{biz.map((_,i)=><td key={i} className={cv+" font-bold text-indigo-800"}>{fmtM(totalKap(i))}</td>)}</tr>
      </tbody></table>
    </div>
  </div>;
}

const LOANS=[
  {id:1,grp:"Soda",company:"We Soda",type:"Bond-5Y",bank:"",start:"2023-10-06",end:"2028-10-06",principal:980000000,outstanding:980000000,ccy:"USD",rate:"9.50%",security:"N/A",status:"Active"},
  {id:2,grp:"Soda",company:"We Soda",type:"Bond-7Y",bank:"",start:"2024-02-14",end:"2031-02-14",principal:750000000,outstanding:750000000,ccy:"USD",rate:"9.375%",security:"N/A",status:"Active"},
  {id:3,grp:"Soda",company:"We Soda US",type:"US Acquisition TLA",bank:"",start:"",end:"2029-02-26",principal:420000000,outstanding:412125000,ccy:"USD",rate:"See link",security:"N/A",status:"Active"},
  {id:4,grp:"Soda",company:"We Soda",type:"US ORRI Bond",bank:"",start:"",end:"2042-03-31",principal:410044346,outstanding:400284995,ccy:"USD",rate:"",security:"N/A",status:"Active"},
  {id:5,grp:"Soda",company:"We Soda",type:"RCF (USD)",bank:"BNP, DB, GS, JPM, PNC",start:"2022-12-12",end:"2026-08-14",principal:435000000,outstanding:100000000,ccy:"USD",rate:"SOFR+3.25%",security:"Pledges",status:"Active"},
  {id:6,grp:"Soda",company:"We Soda",type:"RCF (EUR)",bank:"BNP, DB, GS, JPM, PNC",start:"2022-12-12",end:"2026-08-14",principal:0,outstanding:30000000,ccy:"EUR",rate:"EURIBOR+3.25%",security:"Pledges",status:"Active"},
  {id:7,grp:"Soda",company:"WE Industries Holdings",type:"Isbank Term Loan",bank:"Isbank",start:"2025-04-25",end:"2030-04-25",principal:289000000,outstanding:289000000,ccy:"EUR",rate:"7.604%",security:"Share Pledge",status:"Active"},
  {id:8,grp:"Soda",company:"WE Industries Holdings",type:"Isbank (2)",bank:"Isbank",start:"2025-04-28",end:"2026-04-27",principal:88144557,outstanding:88144557,ccy:"EUR",rate:"6.634%",security:"Share Pledge | Mortgage",status:"Active"},
  {id:9,grp:"Soda",company:"We Soda Enterprises",type:"RCF",bank:"PNC",start:"2025-02-27",end:"2026-08-14",principal:100000000,outstanding:90000000,ccy:"USD",rate:"7.58%",security:"N/A",status:"Active"},
  {id:10,grp:"Soda",company:"Kazan",type:"RCF (1)",bank:"Denizbank",start:"2024-08-05",end:"",principal:125000000,outstanding:44000000,ccy:"EUR",rate:"6.75%",security:"N/A",status:"Active"},
  {id:11,grp:"Soda",company:"Kazan",type:"RCF (2)",bank:"Denizbank",start:"2025-12-08",end:"",principal:125000000,outstanding:8500000,ccy:"EUR",rate:"6.75%",security:"N/A",status:"Active"},
  {id:12,grp:"Soda",company:"WIDT",type:"Eximbank (TL)",bank:"Eximbank",start:"",end:"",principal:0,outstanding:560000000,ccy:"TL",rate:"",security:"Teminat mekt.",status:"Active"},
  {id:13,grp:"Soda",company:"WIDT",type:"Eximbank (EUR)",bank:"Eximbank",start:"",end:"",principal:0,outstanding:5000000,ccy:"EUR",rate:"",security:"Teminat mekt.",status:"Active"},
  {id:14,grp:"Glass",company:"Ciner Glass UK",type:"Refinancing",bank:"Denizbank",start:"2024-07-18",end:"2029-01-18",principal:175000000,outstanding:0,ccy:"EUR",rate:"Euribor+4.5",security:"Guarantee",status:"Closed"},
  {id:15,grp:"Glass",company:"Ciner Glass UK",type:"Acquisition",bank:"Denizbank",start:"2024-11-18",end:"2028-05-25",principal:149000000,outstanding:97500000,ccy:"EUR",rate:"EURIBOR+3.79%",security:"TC guarantee",status:"Active"},
  {id:16,grp:"Glass",company:"Park Cam",type:"Refinancing",bank:"Denizbank",start:"2025-03-28",end:"2029-01-18",principal:170625000,outstanding:161875000,ccy:"EUR",rate:"Euribor+4.25",security:"Pledge | Mortgage",status:"Active"},
  {id:17,grp:"Glass",company:"Park Cam",type:"Exim (EUR)",bank:"Eximbank",start:"2025-04-22",end:"2027-04-26",principal:2000000,outstanding:1500000,ccy:"EUR",rate:"",security:"Teminat mekt.",status:"Active"},
  {id:18,grp:"Glass",company:"Park Cam",type:"Exim (USD)",bank:"Eximbank",start:"2024-03-26",end:"2027-04-05",principal:3600000,outstanding:1780000,ccy:"USD",rate:"",security:"Teminat mekt.",status:"Active"},
  {id:19,grp:"Glass",company:"Park Cam",type:"Factoring",bank:"Deniz Faktoring",start:"2026-01-29",end:"2026-03-10",principal:543743000,outstanding:543743000,ccy:"TRY",rate:"",security:"",status:"Active"},
  {id:20,grp:"Glass",company:"Ciner Glass Belgium",type:"Bridge Loan",bank:"BNPP, Denizbank, DB",start:"2024-02-05",end:"2026-02-05",principal:250000000,outstanding:0,ccy:"EUR",rate:"EURIBOR+3%",security:"Guarantee",status:"Closed"},
  {id:21,grp:"Glass",company:"Ciner Glass Belgium",type:"Incentive",bank:"Belgium Gov",start:"2024-05-29",end:"2033-04-27",principal:5000000,outstanding:3600000,ccy:"EUR",rate:"5.11%",security:"N/A",status:"Active"},
  {id:22,grp:"Glass",company:"Ciner Glass Belgium",type:"Synd. Commercial",bank:"BNPP, KBC, Belfius",start:"2025-08-08",end:"2032-08-08",principal:252400000,outstanding:214942665,ccy:"EUR",rate:"See schedule",security:"Belgium Sec.",status:"Active"},
  {id:23,grp:"Glass",company:"Ciner Glass Belgium",type:"Synd. ECA",bank:"SACE, UKEF, SERV",start:"2025-08-08",end:"2037-08-08",principal:252400000,outstanding:109546567,ccy:"EUR",rate:"",security:"Belgium Sec.",status:"Active"},
  {id:24,grp:"Other",company:"College Hill",type:"Mortgage",bank:"Deutsche Bank",start:"2024-06-28",end:"2025-08-29",principal:14500000,outstanding:0,ccy:"GBP",rate:"6.16%",security:"Mortgage",status:"Closed"},
  {id:25,grp:"Other",company:"Cloak Lane",type:"Mortgage",bank:"Deutsche Bank",start:"2024-06-28",end:"2025-09-04",principal:11075454,outstanding:0,ccy:"GBP",rate:"6.10%",security:"Mortgage",status:"Closed"},
  {id:26,grp:"Other",company:"Memo Aviation",type:"Aircraft Loan",bank:"Deutsche Bank",start:"2023-11-03",end:"2030-08-03",principal:49000000,outstanding:0,ccy:"USD",rate:"SOFR+2.65%",security:"Aircraft Mortgage",status:"Closed"},
];
const LOAN_COS=[...new Set(LOANS.map(l=>l.company))].sort();
type RepRow={date:string,principal:string,interest:string,total:string,remaining:string,paid?:boolean};

function Krediler(){
  const[selCo,setSelCo]=useState("all");const[edits,setEdits]=useState<Record<string,Record<string,string>>>({});
  const[expanded,setExpanded]=useState<number|null>(null);const[repData,setRepData]=useState<Record<number,RepRow[]>>({});
  const[pasteMode,setPasteMode]=useState(false);const[pasteText,setPasteText]=useState("");
  const[customLoans,setCustomLoans]=useState<typeof LOANS>([]);
  useEffect(()=>{try{const d=localStorage.getItem("loan_edits");if(d)setEdits(JSON.parse(d));const r=localStorage.getItem("loan_reps");if(r)setRepData(JSON.parse(r));const c=localStorage.getItem("custom_loans");if(c)setCustomLoans(JSON.parse(c));}catch{}},[]);
  const allLoans=[...LOANS,...customLoans];
  const getF=(lid:number,f:string)=>edits[lid+""]?.[f];
  const setF=(lid:number,f:string,v:string)=>{setEdits(p=>{const n={...p,[lid+""]:{...(p[lid+""]||{}),[f]:v}};try{trackedSetItem("loan_edits",JSON.stringify(n));}catch{}return n;});};
  const val=(l:any,f:string)=>{const e=getF(l.id,f);return e!==undefined?e:String(l[f]||"");};
  const getReps=(lid:number):RepRow[]=>repData[lid]||[];
  const saveReps=(lid:number,r:RepRow[])=>{const nd={...repData,[lid]:r};setRepData(nd);try{trackedSetItem("loan_reps",JSON.stringify(nd));}catch{}};
  const addRep=(lid:number)=>{saveReps(lid,[...getReps(lid),{date:"",principal:"",interest:"",total:"",remaining:""}]);};
  const updRep=(lid:number,idx:number,f:string,v:string)=>{const r=[...getReps(lid)];if(f==="paid")r[idx]={...r[idx],paid:v==="1"};else r[idx]={...r[idx],[f]:v};saveReps(lid,r);};
  const delRep=(lid:number,idx:number)=>{const r=[...getReps(lid)];r.splice(idx,1);saveReps(lid,r);};
  const parsePaste=(lid:number)=>{const lines=pasteText.trim().split("\n").filter(l=>l.trim());const rows:RepRow[]=lines.map(line=>{const cols=line.split("\t");return{date:cols[0]?.trim()||"",principal:cols[1]?.trim().replace(/,/g,"")||"",interest:cols[2]?.trim().replace(/,/g,"")||"",total:cols[3]?.trim().replace(/,/g,"")||"",remaining:cols[4]?.trim().replace(/,/g,"")||""};});saveReps(lid,[...getReps(lid),...rows]);setPasteText("");setPasteMode(false);};
  const ccyClr=(c:string)=>c==="USD"?"text-green-600":c==="EUR"?"text-blue-600":c==="TL"||c==="TRY"?"text-amber-700":"text-purple-600";
  const nextId=Math.max(...[...LOANS.map(l=>l.id),...customLoans.map(l=>l.id),100])+1;
  const[showAddModal,setShowAddModal]=useState(false);
  const addLoan=(grp:string)=>{const nl={id:nextId,grp,company:"",type:"",bank:"",start:"",end:"",principal:0,outstanding:0,ccy:"USD",rate:"",security:"",status:"Active"};const nc=[...customLoans,nl];setCustomLoans(nc);try{trackedSetItem("custom_loans",JSON.stringify(nc));}catch{}setShowAddModal(false);};
  const delLoan=(lid:number)=>{if(!confirm("Delete this loan?"))return;const nc=customLoans.filter(l=>l.id!==lid);setCustomLoans(nc);try{trackedSetItem("custom_loans",JSON.stringify(nc));const re={...repData};delete re[lid];setRepData(re);trackedSetItem("loan_reps",JSON.stringify(re));const ed={...edits};delete ed[lid+""];setEdits(ed);trackedSetItem("loan_edits",JSON.stringify(ed));}catch{}};
  const isCustom=(lid:number)=>customLoans.some(l=>l.id===lid);
  const filteredAll=selCo==="all"?allLoans:allLoans.filter(l=>(getF(l.id,"company")||l.company)===selCo);
  const allCos=[...new Set(allLoans.map(l=>getF(l.id,"company")||l.company))].sort();
  return<div>
    <div className="flex flex-wrap items-center gap-3 mb-4">
      <span className="text-sm font-bold">Loan Portfolio</span>
      <select value={selCo} onChange={e=>setSelCo(e.target.value)} className={ddS}><option value="all">All Companies ({allLoans.length})</option>{allCos.map(c=><option key={c} value={c}>{c} ({allLoans.filter(l=>(getF(l.id,"company")||l.company)===c).length})</option>)}</select>
      <button onClick={()=>setShowAddModal(true)} className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-[11px] font-semibold">+ Add Loan</button>
      <ExportBtn tableId="loanTable" name="Loans"/>
    </div>
    {showAddModal&&<div className="mb-4 p-4 bg-white border border-blue-300 rounded-lg shadow-lg inline-flex items-center gap-3"><span className="text-xs font-semibold">Select Group:</span>{["Soda","Glass","Other"].map(g=><button key={g} onClick={()=>addLoan(g)} className={"px-3 py-1.5 rounded-lg text-xs font-semibold transition "+(g==="Soda"?"bg-blue-100 hover:bg-blue-200 text-blue-700":g==="Glass"?"bg-emerald-100 hover:bg-emerald-200 text-emerald-700":"bg-gray-200 hover:bg-gray-300 text-gray-700")}>{g}</button>)}<button onClick={()=>setShowAddModal(false)} className="px-2 py-1 text-gray-400 hover:text-gray-600 text-xs">Cancel</button></div>}
    <div className="overflow-x-auto"><table id="loanTable" className="border-collapse text-[11px] w-full" style={{minWidth:1500}}>
      <thead><tr className="border-b-2 border-gray-300 bg-gray-50"><th className="px-2 py-2 text-[10px] text-left">ID</th><th className="px-2 py-2 text-[10px] text-left">Group</th><th className="px-2 py-2 text-[10px] text-center" style={{minWidth:80}}>Repayment</th><th className="px-2 py-2 text-[10px] text-left" style={{minWidth:130}}>Company</th><th className="px-2 py-2 text-[10px] text-left" style={{minWidth:140}}>Loan Type</th><th className="px-2 py-2 text-[10px] text-left" style={{minWidth:110}}>Bank</th><th className="px-2 py-2 text-[10px] text-center">Start</th><th className="px-2 py-2 text-[10px] text-center">End</th><th className="px-2 py-2 text-[10px] text-right">Principal</th><th className="px-2 py-2 text-[10px] text-right">Outstanding</th><th className="px-2 py-2 text-[10px] text-center">CCY</th><th className="px-2 py-2 text-[10px] text-left">Rate</th><th className="px-2 py-2 text-[10px] text-left" style={{minWidth:110}}>Security</th><th className="px-2 py-2 text-[10px] text-center">Status</th><th className="px-1 py-2 text-[10px]"></th></tr></thead>
      <tbody>{filteredAll.map(loan=>{
        const isClosed=val(loan,"status")==="Closed";const isExp=expanded===loan.id;const reps=getReps(loan.id);
        return<><tr key={loan.id} className={(isClosed?"bg-gray-100 text-gray-400":"border-b border-gray-100 hover:bg-blue-50/30")+(isExp?" ring-1 ring-blue-300":"")}>
          <td className="px-2 py-1.5 text-gray-400">{loan.id}</td>
          <td className="px-2 py-1.5"><span className={"px-1.5 py-0.5 rounded text-[9px] font-semibold "+(loan.grp==="Soda"?"bg-blue-100 text-blue-700":loan.grp==="Glass"?"bg-emerald-100 text-emerald-700":"bg-gray-200 text-gray-600")}>{loan.grp}</span></td>
          <td className="px-2 py-1.5 text-center"><button onClick={(e)=>{e.stopPropagation();setExpanded(isExp?null:loan.id);}} className="px-2 py-0.5 bg-blue-100 hover:bg-blue-200 text-blue-700 rounded text-[10px] font-semibold transition">Schedule</button></td>
          <td className="px-1 py-0.5"><input className={inp} value={val(loan,"company")} onChange={e=>setF(loan.id,"company",e.target.value)}/></td>
          <td className="px-1 py-0.5"><input className={inp} value={val(loan,"type")} onChange={e=>setF(loan.id,"type",e.target.value)}/></td>
          <td className="px-1 py-0.5"><input className={inp} value={val(loan,"bank")} onChange={e=>setF(loan.id,"bank",e.target.value)}/></td>
          <td className="px-1 py-0.5"><input className={inp+" text-center"} value={val(loan,"start")} onChange={e=>setF(loan.id,"start",e.target.value)}/></td>
          <td className="px-1 py-0.5"><input className={inp+" text-center"} value={val(loan,"end")} onChange={e=>setF(loan.id,"end",e.target.value)}/></td>
          <td className="px-1 py-0.5"><input className={inp+" text-right "+mn} value={addC(val(loan,"principal"))} onChange={e=>setF(loan.id,"principal",stripC(e.target.value))}/></td>
          <td className="px-1 py-0.5"><input className={inp+" text-right "+mn+" bg-amber-50 text-blue-700 font-semibold"} value={addC(val(loan,"outstanding"))} onChange={e=>setF(loan.id,"outstanding",stripC(e.target.value))}/></td>
          <td className={"text-center px-2 py-1.5 font-semibold "+ccyClr(loan.ccy)}>{loan.ccy}</td>
          <td className="px-1 py-0.5"><input className={inp} value={val(loan,"rate")} onChange={e=>setF(loan.id,"rate",e.target.value)}/></td>
          <td className="px-1 py-0.5"><input className={inp} value={val(loan,"security")} onChange={e=>setF(loan.id,"security",e.target.value)}/></td>
          <td className="text-center px-1 py-0.5"><select className={"px-1 py-0.5 border rounded text-[10px] "+(val(loan,"status")==="Active"?"border-emerald-300 bg-emerald-50 text-emerald-700":"border-gray-300 bg-gray-100 text-gray-500")} value={val(loan,"status")} onChange={e=>setF(loan.id,"status",e.target.value)}><option value="Active">Active</option><option value="Closed">Closed</option></select></td>
          <td className="px-1 py-0.5">{isCustom(loan.id)&&<button onClick={()=>delLoan(loan.id)} className="px-1.5 py-0.5 bg-red-50 hover:bg-red-100 text-red-500 rounded text-[10px]" title="Delete loan">×</button>}</td>
        </tr>
        {isExp&&<tr key={loan.id+"_rep"}><td colSpan={15} className="p-4 bg-blue-50/40 border-b-2 border-blue-200">
          <div className="flex justify-between items-center mb-3 flex-wrap gap-2"><span className="text-sm font-bold text-blue-700">Repayment Schedule — {val(loan,"type")} ({loan.ccy})</span><div className="flex gap-2"><button onClick={()=>setPasteMode(!pasteMode)} className="px-3 py-1 bg-amber-100 hover:bg-amber-200 text-amber-700 rounded text-[11px] font-semibold">Paste from Excel</button><button onClick={()=>addRep(loan.id)} className="px-3 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded text-[11px] font-semibold">+ Add Row</button></div></div>
          {pasteMode&&expanded===loan.id&&<div className="mb-3 p-3 bg-white border border-amber-300 rounded-lg"><div className="text-[11px] text-gray-600 mb-2">Paste tab-separated data (5 columns: Date, Principal, Interest, Total, Remaining Balance)</div><textarea value={pasteText} onChange={e=>setPasteText(e.target.value)} className="w-full h-24 px-3 py-2 border border-gray-300 rounded text-xs font-mono outline-none resize-none" placeholder={"2026-01-15\t5000000\t250000\t5250000\t95000000\n2026-04-15\t5000000\t237500\t5237500\t90000000"}/><div className="flex gap-2 mt-2"><button onClick={()=>parsePaste(loan.id)} className="px-3 py-1 bg-emerald-600 text-white rounded text-[11px] font-semibold">Import ({pasteText.trim().split("\n").filter(l=>l.trim()).length} rows)</button><button onClick={()=>{setPasteMode(false);setPasteText("");}} className="px-3 py-1 bg-gray-200 rounded text-[11px]">Cancel</button></div></div>}
          {reps.length===0&&!pasteMode&&<div className="flex flex-col items-center py-6 gap-3 bg-white rounded-lg border border-dashed border-gray-300"><div className="text-gray-400 text-sm">No repayment schedule entered yet</div><div className="flex gap-3"><button onClick={()=>addRep(loan.id)} className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-sm font-semibold shadow">+ Add Row</button><button onClick={()=>setPasteMode(true)} className="px-4 py-2 bg-amber-100 hover:bg-amber-200 text-amber-700 rounded-lg text-sm font-semibold">Paste from Excel</button></div></div>}
          {reps.length>0&&<><table className="border-collapse text-[11px] w-full max-w-[1000px]"><thead><tr className="border-b-2 border-gray-300 bg-white"><th className="text-center px-2 py-1.5 w-8">Paid</th><th className="text-left px-2 py-1.5">Date</th><th className="text-right px-2 py-1.5">Principal</th><th className="text-right px-2 py-1.5">Interest</th><th className="text-right px-2 py-1.5">Total</th><th className="text-right px-2 py-1.5">Remaining</th><th className="px-1"></th></tr></thead>
            <tbody>{reps.map((r,ri)=><tr key={ri} className={"border-b border-gray-100 "+(r.paid?"bg-emerald-50":"hover:bg-white")}><td className="text-center px-1 py-0.5"><input type="checkbox" checked={!!r.paid} onChange={e=>{updRep(loan.id,ri,"paid",e.target.checked?"1":"");}} className="w-3.5 h-3.5 accent-emerald-600 cursor-pointer"/></td><td className="px-1 py-0.5"><input className={inp+" text-center"+(r.paid?" line-through text-gray-400":"")} value={r.date} onChange={e=>updRep(loan.id,ri,"date",e.target.value)}/></td><td className="px-1 py-0.5"><input className={inp+" text-right "+mn+(r.paid?" line-through text-gray-400":"")} value={addC(r.principal)} onChange={e=>updRep(loan.id,ri,"principal",stripC(e.target.value))}/></td><td className="px-1 py-0.5"><input className={inp+" text-right "+mn+(r.paid?" line-through text-gray-400":"")} value={addC(r.interest)} onChange={e=>updRep(loan.id,ri,"interest",stripC(e.target.value))}/></td><td className="px-1 py-0.5"><input className={inp+" text-right "+mn+(r.paid?" line-through text-gray-400":"")} value={addC(r.total)} onChange={e=>updRep(loan.id,ri,"total",stripC(e.target.value))}/></td><td className="px-1 py-0.5"><input className={inp+" text-right "+mn+" font-semibold"+(r.paid?" line-through text-gray-400 bg-emerald-50":" bg-amber-50")} value={addC(r.remaining)} onChange={e=>updRep(loan.id,ri,"remaining",stripC(e.target.value))}/></td><td className="px-1 py-0.5"><button onClick={()=>delRep(loan.id,ri)} className="px-1.5 py-0.5 bg-red-50 text-red-500 rounded text-[10px] hover:bg-red-100">×</button></td></tr>)}</tbody></table>
          <div className="mt-2"><button onClick={()=>addRep(loan.id)} className="px-3 py-1 bg-blue-100 hover:bg-blue-200 text-blue-700 rounded text-[11px] font-semibold">+ Add Row</button></div></>}
        </td></tr>}</>;})}</tbody></table></div>
    <div className="grid grid-cols-3 gap-3 mt-4">{[{l:"Active Loans",v:filteredAll.filter(l=>val(l,"status")!=="Closed").length,c:"#2563eb"},{l:"Total Outstanding",v:filteredAll.filter(l=>val(l,"status")!=="Closed").reduce((s,l)=>s+(parseFloat(stripC(val(l,"outstanding")))||0),0),c:"#d97706",f:true},{l:"Closed",v:filteredAll.filter(l=>val(l,"status")==="Closed").length,c:"#6b7280"}].map((c,i)=><div key={i} className="rounded-xl border border-gray-200 p-4"><div className="text-[10px] text-gray-400">{c.l}</div><div className={"text-lg font-bold "+mn} style={{color:c.c}}>{c.f?fmtC(c.v):c.v}</div></div>)}</div>
  </div>;
}


// ═══════════════════════════════════════
// CASHFLOW WITH SUB-TABS
// Container that holds the main Cashflow plus 12 sub-tabs:
// Export, US Export, Domestic, BNP TR UK, ABC, SAISA RF, BNP US,
// Other Disc., Electric, Cheques, Garanti Factoring, YK Factoring
// ═══════════════════════════════════════
const CF_SUB_TABS:{id:string,label:string,color:string}[]=[
  {id:"main",label:"Cashflow",color:"#2563eb"},
  {id:"export",label:"Export",color:"#0d9488"},
  {id:"usexport",label:"US Export",color:"#0891b2"},
  {id:"domestic",label:"Domestic",color:"#ca8a04"},
  {id:"bnp_truk",label:"BNP TR UK",color:"#7c3aed"},
  {id:"abc",label:"ABC",color:"#db2777"},
  {id:"saisa_rf",label:"SAISA RF",color:"#ea580c"},
  {id:"bnp_us",label:"BNP US",color:"#0284c7"},
  {id:"other_disc",label:"Other Disc.",color:"#65a30d"},
  {id:"electric",label:"Electric",color:"#f59e0b"},
  {id:"cheques",label:"Cheques",color:"#8b5cf6"},
  {id:"garanti_factoring",label:"Garanti Factoring",color:"#16a34a"},
  {id:"yk_factoring",label:"YK Factoring",color:"#e11d48"},
];

function PlaceholderSheet({title,subId}:{title:string,subId:string}){
  return<div className="py-10 px-6 text-center rounded-xl border-2 border-dashed border-gray-200 bg-gray-50/50">
    <div className="text-2xl font-bold text-gray-300 mb-2">{title}</div>
    <div className="text-xs text-gray-400 mb-1">Alt sekme iskeleti hazır</div>
    <div className="text-[10px] text-gray-300 font-mono">id: {subId}</div>
    <div className="text-[11px] text-gray-500 mt-4">İçerik yapısı konuşuldukça bu sekme doldurulacak.</div>
  </div>;
}

// ═══════════════════════════════════════
// INVOICE DISCOUNT SHEET — for BNP TR UK, ABC, SAISA RF, BNP US, Other Disc., Garanti/YK Factoring etc.
// Tracks invoices sold/discounted before due date (kırdırılan faturalar)
// Columns: Invoice Number | Discount Date | Obligor | CCY | Invoice Amount | Due Date
// ═══════════════════════════════════════
type InvRow={id:string,invoice:string,discDate:string,obligor:string,ccy:string,amount:string,dueDate:string};

function InvoiceDiscountSheet({title,storageKey,sm,sy,canEdit,accent}:{title:string,storageKey:string,sm:number,sy:number,canEdit:boolean,accent:string}){
  const[rows,setRows]=useState<InvRow[]>([]);
  const[adding,setAdding]=useState(false);
  const[sortBy,setSortBy]=useState<keyof InvRow|"">("dueDate");
  const[filterObligor,setFilterObligor]=useState("");
  const[filterCcy,setFilterCcy]=useState("");
  const key=storageKey+"_"+sy+"_"+sm;
  useEffect(()=>{try{const s=localStorage.getItem(key);setRows(s?JSON.parse(s):[]);}catch{setRows([]);}},[key]);
  const save=(r:InvRow[])=>{setRows(r);try{trackedSetItem(key,JSON.stringify(r));}catch{}};
  const addRow=()=>{const nr:InvRow={id:"inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:"",discDate:"",obligor:"",ccy:"EUR",amount:"",dueDate:""};save([...rows,nr]);setAdding(true);};
  const updRow=(id:string,f:keyof InvRow,v:string)=>{save(rows.map(r=>r.id===id?{...r,[f]:v}:r));};
  const delRow=(id:string)=>{if(!confirm("Bu faturayı silmek istediğinize emin misiniz?"))return;save(rows.filter(r=>r.id!==id));};

  // Paste from Excel (tab-separated)
  const[pasteMode,setPasteMode]=useState(false);
  const[pasteText,setPasteText]=useState("");
  const parsePaste=()=>{
    const lines=pasteText.trim().split("\n").filter(l=>l.trim());
    const parsed:InvRow[]=lines.map(line=>{const cols=line.split("\t");return{id:"inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:(cols[0]||"").trim(),discDate:(cols[1]||"").trim(),obligor:(cols[2]||"").trim(),ccy:(cols[3]||"").trim().toUpperCase(),amount:(cols[4]||"").trim().replace(/,/g,"").replace(/\s/g,""),dueDate:(cols[5]||"").trim()};});
    save([...rows,...parsed]);setPasteText("");setPasteMode(false);
  };

  // Filter & sort
  const obligors=[...new Set(rows.map(r=>r.obligor).filter(o=>o))].sort();
  const currencies=[...new Set(rows.map(r=>r.ccy).filter(c=>c))].sort();
  let view=[...rows];
  if(filterObligor)view=view.filter(r=>r.obligor===filterObligor);
  if(filterCcy)view=view.filter(r=>r.ccy===filterCcy);
  if(sortBy){
    view.sort((a,b)=>{const av=(a[sortBy as keyof InvRow]||"")+"";const bv=(b[sortBy as keyof InvRow]||"")+"";
      if(sortBy==="dueDate"||sortBy==="discDate"){
        // Parse DD/MM/YYYY
        const parse=(s:string)=>{const p=s.split(/[\/\-\.]/);if(p.length===3&&p[2]?.length===4)return new Date(+p[2],+p[1]-1,+p[0]).getTime();return new Date(s).getTime()||0;};
        return parse(av)-parse(bv);
      }
      if(sortBy==="amount")return (parseFloat(stripC(av))||0)-(parseFloat(stripC(bv))||0);
      return av.localeCompare(bv);
    });
  }

  // Totals by currency
  const totalsByCcy:Record<string,number>={};
  view.forEach(r=>{const a=parseFloat(stripC(r.amount))||0;totalsByCcy[r.ccy]=(totalsByCcy[r.ccy]||0)+a;});

  const ccyColor=(c:string)=>c==="USD"?"text-green-600":c==="EUR"?"text-blue-600":c==="GBP"?"text-purple-600":c==="TRY"||c==="TL"?"text-amber-700":"text-gray-600";

  return<div>
    <div className="flex justify-between items-center mb-3 flex-wrap gap-2">
      <div><span className="text-sm font-bold" style={{color:accent}}>{title}</span><span className="text-[10px] text-gray-400 ml-2">Kırdırılan faturalar — {MONTHS[sm]} {sy}</span></div>
      <div className="flex gap-2 flex-wrap">
        {canEdit&&<button onClick={()=>setPasteMode(!pasteMode)} className="px-3 py-1 bg-amber-100 hover:bg-amber-200 text-amber-700 rounded text-[11px] font-semibold">Excel'den Yapıştır</button>}
        {canEdit&&<button onClick={addRow} className="px-3 py-1 text-white rounded text-[11px] font-semibold" style={{background:accent}}>+ Fatura Ekle</button>}
        <ExportBtn tableId={"inv_"+storageKey} name={title+"_"+MONTHS[sm]+"_"+sy}/>
      </div>
    </div>

    {pasteMode&&<div className="mb-3 p-3 bg-white border border-amber-300 rounded-lg">
      <div className="text-[11px] text-gray-600 mb-2">Tab-separated data yapıştırın (6 sütun: Invoice | Discount Date | Obligor | CCY | Amount | Due Date)</div>
      <textarea value={pasteText} onChange={e=>setPasteText(e.target.value)} className="w-full h-32 px-3 py-2 border border-gray-300 rounded text-xs font-mono outline-none resize-none" placeholder={"IM20250000000498\t08/12/2025\tGUARDIAN EUROPE\tEUR\t92191.27\t23/01/2026"}/>
      <div className="flex gap-2 mt-2"><button onClick={parsePaste} className="px-3 py-1 bg-emerald-600 text-white rounded text-[11px] font-semibold">İçeri Aktar ({pasteText.trim().split("\n").filter(l=>l.trim()).length} satır)</button><button onClick={()=>{setPasteMode(false);setPasteText("");}} className="px-3 py-1 bg-gray-200 rounded text-[11px]">İptal</button></div>
    </div>}

    {/* Filters */}
    {rows.length>0&&<div className="flex gap-2 mb-3 flex-wrap text-[10px]">
      <div className="flex items-center gap-1"><span className="text-gray-500">Obligor:</span><select value={filterObligor} onChange={e=>setFilterObligor(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="">Tümü ({obligors.length})</option>{obligors.map(o=><option key={o} value={o}>{o}</option>)}</select></div>
      <div className="flex items-center gap-1"><span className="text-gray-500">CCY:</span><select value={filterCcy} onChange={e=>setFilterCcy(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="">Tümü</option>{currencies.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
      <div className="flex items-center gap-1"><span className="text-gray-500">Sırala:</span><select value={sortBy} onChange={e=>setSortBy(e.target.value as any)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="dueDate">Due Date</option><option value="discDate">Discount Date</option><option value="obligor">Obligor</option><option value="amount">Amount</option><option value="invoice">Invoice</option><option value="">—</option></select></div>
      {(filterObligor||filterCcy)&&<button onClick={()=>{setFilterObligor("");setFilterCcy("");}} className="px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded text-[10px]">Filtreyi Temizle</button>}
      <span className="ml-auto text-gray-400">{view.length} / {rows.length} fatura</span>
    </div>}

    {rows.length===0&&<div className="py-10 px-6 text-center rounded-xl border-2 border-dashed border-gray-200 bg-gray-50/50">
      <div className="text-gray-400 text-sm mb-3">Henüz bu ay için fatura girilmedi</div>
      {canEdit&&<div className="flex gap-3 justify-center"><button onClick={addRow} className="px-4 py-2 text-white rounded-lg text-sm font-semibold shadow" style={{background:accent}}>+ Fatura Ekle</button><button onClick={()=>setPasteMode(true)} className="px-4 py-2 bg-amber-100 hover:bg-amber-200 text-amber-700 rounded-lg text-sm font-semibold">Excel'den Yapıştır</button></div>}
    </div>}

    {rows.length>0&&<div className="overflow-x-auto"><table id={"inv_"+storageKey} className="border-collapse text-[11px] w-full" style={{minWidth:1100}}>
      <thead><tr className="border-b-2 border-gray-300" style={{background:accent+"12"}}>
        <th className="px-2 py-2 text-[10px] text-left font-semibold" style={{minWidth:40}}>#</th>
        <th className="px-2 py-2 text-[10px] text-left font-semibold" style={{minWidth:180}}>Invoice Number</th>
        <th className="px-2 py-2 text-[10px] text-center font-semibold" style={{minWidth:100}}>Discount Date</th>
        <th className="px-2 py-2 text-[10px] text-left font-semibold" style={{minWidth:220}}>Obligor</th>
        <th className="px-2 py-2 text-[10px] text-center font-semibold" style={{minWidth:55}}>CCY</th>
        <th className="px-2 py-2 text-[10px] text-right font-semibold" style={{minWidth:130}}>Invoice Amount</th>
        <th className="px-2 py-2 text-[10px] text-center font-semibold" style={{minWidth:100}}>Due Date</th>
        {canEdit&&<th className="w-6"></th>}
      </tr></thead>
      <tbody>{view.map((r,i)=><tr key={r.id} className="border-b border-gray-100 hover:bg-gray-50">
        <td className="px-2 py-1 text-[10px] text-gray-400">{i+1}</td>
        <td className="px-1 py-0.5"><input className={inp+" font-mono text-[10px]"} value={r.invoice} onChange={e=>updRow(r.id,"invoice",e.target.value)} placeholder="SW20260000000000"/></td>
        <td className="px-1 py-0.5"><input className={inp+" text-center text-[10px]"} value={r.discDate} onChange={e=>updRow(r.id,"discDate",e.target.value)} placeholder="DD/MM/YYYY"/></td>
        <td className="px-1 py-0.5"><input className={inp+" text-[10px]"} value={r.obligor} onChange={e=>updRow(r.id,"obligor",e.target.value)} placeholder="Obligor name"/></td>
        <td className="px-1 py-0.5"><select value={r.ccy} onChange={e=>updRow(r.id,"ccy",e.target.value)} className={"w-full px-1 py-0.5 border border-gray-200 rounded text-[10px] bg-white text-center font-semibold outline-none "+ccyColor(r.ccy)}><option value="EUR">EUR</option><option value="USD">USD</option><option value="GBP">GBP</option><option value="TRY">TRY</option><option value="CHF">CHF</option><option value="JPY">JPY</option></select></td>
        <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={addC(r.amount)} onChange={e=>updRow(r.id,"amount",stripC(e.target.value))} placeholder="0.00"/></td>
        <td className="px-1 py-0.5"><input className={inp+" text-center text-[10px]"} value={r.dueDate} onChange={e=>updRow(r.id,"dueDate",e.target.value)} placeholder="DD/MM/YYYY"/></td>
        {canEdit&&<td className="px-1 py-0.5"><button onClick={()=>delRow(r.id)} className="px-1 py-0.5 text-red-400 hover:bg-red-50 hover:text-red-600 rounded text-[11px]" title="Faturayı sil">×</button></td>}
      </tr>)}</tbody>
    </table></div>}

    {/* Totals by currency */}
    {view.length>0&&<div className="mt-4 grid grid-cols-2 md:grid-cols-4 gap-2">
      {Object.entries(totalsByCcy).map(([c,t])=><div key={c} className="rounded-lg border border-gray-200 p-3 bg-white">
        <div className="text-[10px] text-gray-400 uppercase tracking-wide">Toplam {c}</div>
        <div className={"text-base font-bold font-mono mt-0.5 "+ccyColor(c)}>{fmt(t)}</div>
        <div className="text-[9px] text-gray-400 mt-0.5">{view.filter(r=>r.ccy===c).length} fatura</div>
      </div>)}
    </div>}
  </div>;
}



// ═══════════════════════════════════════════════════════════════════════════
// BNP TR UK — exact replica of BNP_TR_UK.xlsx
// Left: Invoice table (Invoice Number, Discount Date, Obligor, CCY, Amount, Due Date)
// Right: Parities (I3=TotalLimit, I5=GBP/USD, I7=EUR/USD)
//        + Obligor limits table (I8:P41) with live SUMIF formulas
// Formulas propagate to new rows automatically.
// ═══════════════════════════════════════════════════════════════════════════

const BNP_DEFAULT_INVOICES=[{invoice:"IM20250000000498",discDate:"08/12/2025",obligor:"GUARDIAN EUROPE (S.A.R.L)",ccy:"EUR",amount:92191.27,dueDate:"23/01/2026"},{invoice:"SW20250100002986",discDate:"24/12/2025",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:105545.98,dueDate:"30/01/2026"},{invoice:"SW20250100003075",discDate:"08/12/2025",obligor:"GUARDIAN INDUSTRIES CORP LTD",ccy:"USD",amount:112300.98,dueDate:"03/02/2026"},{invoice:"SW20250100003049",discDate:"24/12/2025",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:105573.2,dueDate:"05/02/2026"},{invoice:"SW20250100003141",discDate:"24/12/2025",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:79088.29,dueDate:"11/02/2026"},{invoice:"SW20250100003237",discDate:"24/12/2025",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:75709.63,dueDate:"22/02/2026"},{invoice:"SW20250100003257",discDate:"16/01/2025",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:50559.43,dueDate:"24/02/2026"},{invoice:"SD20250000000531",discDate:"16/01/2025",obligor:"GUARDIAN EUROPE SARL (GOOLE DIV.)",ccy:"GBP",amount:49902.08,dueDate:"27/02/2026"},{invoice:"SW20250100003382",discDate:"16/01/2025",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:86315.89,dueDate:"27/02/2026"},{invoice:"SD20250000000539",discDate:"16/01/2025",obligor:"GUARDIAN EUROPE SARL (GOOLE DIV.)",ccy:"GBP",amount:49701.14,dueDate:"01/03/2026"},{invoice:"SW20260100000066",discDate:"16/01/2025",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:61539.45,dueDate:"14/03/2026"},{invoice:"SW20260100000096",discDate:"02/02/2026",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:51170.21,dueDate:"17/03/2026"},{invoice:"SD20260000000020",discDate:"22/01/2026",obligor:"GUARDIAN EUROPE SARL (GOOLE DIV.)",ccy:"GBP",amount:82947.5,dueDate:"20/03/2026"},{invoice:"SD20250000000519",discDate:"16/01/2025",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:116789.3,dueDate:"22/03/2026"},{invoice:"SD20250000000522",discDate:"16/01/2025",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:90621.78,dueDate:"22/03/2026"},{invoice:"SD20250000000524",discDate:"16/01/2025",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:59286.79,dueDate:"22/03/2026"},{invoice:"SD20250000000525",discDate:"16/01/2025",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:82540.37,dueDate:"22/03/2026"},{invoice:"SD20250000000526",discDate:"16/01/2025",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:89654.33,dueDate:"22/03/2026"},{invoice:"SW20260100000143",discDate:"02/02/2026",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:112679.21,dueDate:"24/03/2026"},{invoice:"IM20260000000029",discDate:"28/01/2026",obligor:"GUARDIAN EUROPE (S.A.R.L)",ccy:"EUR",amount:98310.0,dueDate:"27/03/2026"},{invoice:"SW20260100000164",discDate:"02/02/2026",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:51259.53,dueDate:"27/03/2026"},{invoice:"SW20260100000187",discDate:"02/02/2026",obligor:"GUARDIAN CZESTOCHOWA SP. ZO.O.",ccy:"EUR",amount:184560.21,dueDate:"29/03/2026"},{invoice:"SW20250100003369",discDate:"30/12/2025",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:310000.0,dueDate:"29/03/2026"},{invoice:"SW20260100000170",discDate:"04/02/2026",obligor:"SAUDI GUARDIAN INTERNATIONAL",ccy:"USD",amount:1674000.0,dueDate:"30/03/2026"},{invoice:"SD20250000000546",discDate:"16/01/2025",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:86317.49,dueDate:"31/03/2026"},{invoice:"SD20250000000550",discDate:"16/01/2025",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:57583.61,dueDate:"31/03/2026"},{invoice:"IM20260000000040",discDate:"13/02/2026",obligor:"GUARDIAN EUROPE (S.A.R.L)",ccy:"EUR",amount:98418.48,dueDate:"01/04/2026"},{invoice:"SW20250100003368",discDate:"22/01/2026",obligor:"Cermin DMCC",ccy:"USD",amount:110400.0,dueDate:"05/04/2026"},{invoice:"SW20250100003365",discDate:"09/03/2026",obligor:"MANUCHAR NV",ccy:"USD",amount:400000.0,dueDate:"05/04/2026"},{invoice:"SW20260100000030",discDate:"16/01/2025",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:310000.0,dueDate:"06/04/2026"},{invoice:"SD20260000000012",discDate:"22/01/2026",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:113863.68,dueDate:"12/04/2026"},{invoice:"SW20260100000258",discDate:"19/02/2026",obligor:"GUARDIAN CZESTOCHOWA SP. ZO.O.",ccy:"EUR",amount:210774.63,dueDate:"13/04/2026"},{invoice:"SW20260100000264",discDate:"19/02/2026",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:112620.34,dueDate:"13/04/2026"},{invoice:"SW20260100000063",discDate:"16/01/2025",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:68750.0,dueDate:"13/04/2026"},{invoice:"SW20250100003223",discDate:"16/01/2025",obligor:"UNIGLOBAL GENERAL TRADING DWC-LLC",ccy:"USD",amount:3460000.0,dueDate:"14/04/2026"},{invoice:"SW20250100002941",discDate:"19/11/2025",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:595000.0,dueDate:"15/04/2026"},{invoice:"IM20260000000060",discDate:"06/03/2026",obligor:"GUARDIAN EUROPE (S.A.R.L)",ccy:"EUR",amount:104967.96,dueDate:"15/04/2026"},{invoice:"SW20260100000086",discDate:"16/01/2025",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:71200.0,dueDate:"15/04/2026"},{invoice:"SW20260100000344",discDate:"23/02/2026",obligor:"CHINA TIANCHEN ENGINEERING",ccy:"USD",amount:3100000.0,dueDate:"17/04/2026"},{invoice:"SW20260100000102",discDate:"22/01/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:200645.9,dueDate:"19/04/2026"},{invoice:"SW20260100000104",discDate:"22/01/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:72520.0,dueDate:"19/04/2026"},{invoice:"SD20260000000018",discDate:"22/01/2026",obligor:"ARDAGH GLASS LTD",ccy:"GBP",amount:103430.88,dueDate:"19/04/2026"},{invoice:"SD20260000000025",discDate:"22/01/2026",obligor:"ENCIRC LTD (IN LINE)",ccy:"GBP",amount:223347.6,dueDate:"19/04/2026"},{invoice:"SW20260100000037",discDate:"16/01/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:77313.6,dueDate:"20/04/2026"},{invoice:"SW20260100000024",discDate:"16/01/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:48289.14,dueDate:"20/04/2026"},{invoice:"SW20260100000046",discDate:"16/01/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:48349.32,dueDate:"20/04/2026"},{invoice:"SW20260100000047",discDate:"16/01/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:77441.04,dueDate:"20/04/2026"},{invoice:"SW20250100003383",discDate:"16/01/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:43083.44,dueDate:"20/04/2026"},{invoice:"SW20260100000289",discDate:"03/03/2026",obligor:"GUARDIAN CZESTOCHOWA SP. ZO.O.",ccy:"EUR",amount:184284.84,dueDate:"20/04/2026"},{invoice:"SW20260100000318",discDate:"03/03/2026",obligor:"GUARDIAN CZESTOCHOWA SP. ZO.O.",ccy:"EUR",amount:131391.69,dueDate:"20/04/2026"},{invoice:"SW20260100000312",discDate:"03/03/2026",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:112575.68,dueDate:"20/04/2026"},{invoice:"SW20260100000116",discDate:"22/01/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:71200.0,dueDate:"20/04/2026"},{invoice:"SW20250100003340",discDate:"29/12/2025",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:55966.76,dueDate:"20/04/2026"},{invoice:"SW20250100003324",discDate:"29/12/2025",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:64863.74,dueDate:"20/04/2026"},{invoice:"SW20260100000119",discDate:"22/01/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:77375.55,dueDate:"20/04/2026"},{invoice:"SW20260100000341",discDate:"09/03/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:206250.0,dueDate:"21/04/2026"},{invoice:"SW20260100000129",discDate:"22/01/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:190195.2,dueDate:"21/04/2026"},{invoice:"SW20260100000122",discDate:"22/01/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:83500.0,dueDate:"21/04/2026"},{invoice:"SW20260100000420",discDate:"16/03/2026",obligor:"CHINA TIANCHEN ENGINEERING",ccy:"USD",amount:540000.0,dueDate:"21/04/2026"},{invoice:"SW20260100000342",discDate:"24/03/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:206250.0,dueDate:"21/04/2026"},{invoice:"IM20260000000079",discDate:"06/03/2026",obligor:"GUARDIAN EUROPE (S.A.R.L)",ccy:"EUR",amount:104610.88,dueDate:"22/04/2026"},{invoice:"SW20260100000131",discDate:"27/01/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:310000.0,dueDate:"22/04/2026"},{invoice:"SW20260100000137",discDate:"10/02/2026",obligor:"GOVCREST INTERNATIONAL SRL",ccy:"EUR",amount:1176777.0,dueDate:"23/04/2026"},{invoice:"SW20260100000093",discDate:"09/03/2026",obligor:"AGC INC",ccy:"USD",amount:111164.02,dueDate:"24/04/2026"},{invoice:"SD20260000000027",discDate:"19/03/2026",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:92378.16,dueDate:"24/04/2026"},{invoice:"IM20260000000016",discDate:"19/03/2026",obligor:"SAINT-GOBAIN SEKURIT DEUTSCHLA",ccy:"EUR",amount:124656.0,dueDate:"24/04/2026"},{invoice:"SD20260000000019",discDate:"24/03/2026",obligor:"SAINT GOBAIN UK",ccy:"GBP",amount:93922.8,dueDate:"24/04/2026"},{invoice:"SW20260100000262",discDate:"10/02/2026",obligor:"SISECAM FLAT GLASS SOUTH ITALY",ccy:"EUR",amount:665000.0,dueDate:"25/04/2026"},{invoice:"SW20260100000369",discDate:"24/03/2026",obligor:"GUARDIAN INDUSTRIES CORP LTD",ccy:"USD",amount:50265.12,dueDate:"25/04/2026"},{invoice:"SD20260000000037",discDate:"23/02/2026",obligor:"ENCIRC LTD (IN LINE)",ccy:"GBP",amount:200566.8,dueDate:"26/04/2026"},{invoice:"SA20260000000013",discDate:"24/03/2026",obligor:"GUARDIAN EUROPE SARL (GOOLE DIV.)",ccy:"GBP",amount:65021.02,dueDate:"27/04/2026"},{invoice:"SA20260000000025",discDate:"24/03/2026",obligor:"GUARDIAN EUROPE SARL (GOOLE DIV.)",ccy:"GBP",amount:73585.2,dueDate:"27/04/2026"},{invoice:"SW20260100000171",discDate:"30/01/2026",obligor:"GUARDIAN ZOUJAJ INTERNATIONAL",ccy:"USD",amount:1337000.0,dueDate:"29/04/2026"},{invoice:"IM20260000000089",discDate:"06/03/2026",obligor:"GUARDIAN EUROPE (S.A.R.L)",ccy:"EUR",amount:111115.16,dueDate:"29/04/2026"},{invoice:"SD20260000000067",discDate:"24/03/2026",obligor:"GUARDIAN EUROPE SARL (GOOLE DIV.)",ccy:"GBP",amount:73324.82,dueDate:"29/04/2026"},{invoice:"IM20250000000590",discDate:"20/01/2026",obligor:"AGC FLAT GLASS CZECH A.S.",ccy:"EUR",amount:725883.84,dueDate:"30/04/2026"},{invoice:"SW20260100000415",discDate:"03/03/2026",obligor:"GUARDIAN CZESTOCHOWA SP. ZO.O.",ccy:"EUR",amount:131646.96,dueDate:"30/04/2026"},{invoice:"SW20260100000416",discDate:"24/03/2026",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:61385.17,dueDate:"30/04/2026"},{invoice:"SD20260000000042",discDate:"09/02/2026",obligor:"ENCIRC LTD (IN LINE)",ccy:"GBP",amount:209734.56,dueDate:"01/05/2026"},{invoice:"SW20260100000489",discDate:"16/03/2026",obligor:"CHINA TIANCHEN ENGINEERING",ccy:"USD",amount:540000.0,dueDate:"01/05/2026"},{invoice:"SD20260000000036",discDate:"19/03/2026",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:106389.36,dueDate:"01/05/2026"},{invoice:"IM20260000000031",discDate:"19/03/2026",obligor:"SAINT-GOBAIN SEKURIT DEUTSCHLA",ccy:"EUR",amount:108600.0,dueDate:"01/05/2026"},{invoice:"SD20260000000030",discDate:"24/03/2026",obligor:"SAINT GOBAIN UK",ccy:"GBP",amount:63841.92,dueDate:"01/05/2026"},{invoice:"IM20260000000105",discDate:"30/03/2026",obligor:"GUARDIAN EUROPE (S.A.R.L)",ccy:"EUR",amount:85418.96,dueDate:"03/05/2026"},{invoice:"SW20260100000188",discDate:"19/03/2026",obligor:"BERNER OY",ccy:"EUR",amount:101000.0,dueDate:"04/05/2026"},{invoice:"SD20260000000077",discDate:"24/03/2026",obligor:"GUARDIAN EUROPE SARL (GOOLE DIV.)",ccy:"GBP",amount:73460.68,dueDate:"04/05/2026"},{invoice:"SW20260100000447",discDate:"24/03/2026",obligor:"GUARDIAN INDUSTRIES CORP LTD",ccy:"USD",amount:50254.08,dueDate:"04/05/2026"},{invoice:"SW20260100000520",discDate:"16/03/2026",obligor:"CHINA TIANCHEN ENGINEERING",ccy:"USD",amount:540000.0,dueDate:"05/05/2026"},{invoice:"SW20260100000225",discDate:"09/02/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:130118.3,dueDate:"06/05/2026"},{invoice:"SW20260100000218",discDate:"09/02/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:204662.7,dueDate:"06/05/2026"},{invoice:"SD20260000000039",discDate:"24/03/2026",obligor:"SAINT GOBAIN UK",ccy:"GBP",amount:56530.8,dueDate:"06/05/2026"},{invoice:"SD20260000000041",discDate:"24/03/2026",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:84021.84,dueDate:"06/05/2026"},{invoice:"SW20260100000207",discDate:"24/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:63356.8,dueDate:"07/05/2026"},{invoice:"SW20260100000186",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:53400.0,dueDate:"07/05/2026"},{invoice:"SW20260100000202",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:71200.0,dueDate:"07/05/2026"},{invoice:"SW20260100000203",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:53400.0,dueDate:"07/05/2026"},{invoice:"SW20260100000455",discDate:"24/03/2026",obligor:"GUARDIAN INDUSTRIES CORP LTD",ccy:"USD",amount:50360.8,dueDate:"09/05/2026"},{invoice:"SW20260100000414",discDate:"26/03/2026",obligor:"SISECAM FLAT GLASS SOUTH ITALY",ccy:"EUR",amount:570000.0,dueDate:"09/05/2026"},{invoice:"SD20260000000047",discDate:"23/02/2026",obligor:"ENCIRC LTD (IN LINE)",ccy:"GBP",amount:209790.0,dueDate:"10/05/2026"},{invoice:"SW20260100000568",discDate:"16/03/2026",obligor:"CHINA TIANCHEN ENGINEERING",ccy:"USD",amount:180000.0,dueDate:"10/05/2026"},{invoice:"SW20260100000206",discDate:"24/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:91408.2,dueDate:"10/05/2026"},{invoice:"SW20260100000584",discDate:"16/03/2026",obligor:"CHINA TIANCHEN ENGINEERING",ccy:"USD",amount:180000.0,dueDate:"11/05/2026"},{invoice:"SW20260100000583",discDate:"16/03/2026",obligor:"CHINA TIANCHEN ENGINEERING",ccy:"USD",amount:180000.0,dueDate:"11/05/2026"},{invoice:"SW20260100000265",discDate:"24/03/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:155000.0,dueDate:"11/05/2026"},{invoice:"SW20260100000734",discDate:"31/03/2026",obligor:"NEWPORT INDUSTRIES LTD",ccy:"EUR",amount:745280.0,dueDate:"11/05/2026"},{invoice:"SW20260100000735",discDate:"31/03/2026",obligor:"NEWPORT INDUSTRIES LTD",ccy:"EUR",amount:1042040.0,dueDate:"11/05/2026"},{invoice:"SW20260100000276",discDate:"04/03/2026",obligor:"Cermin DMCC",ccy:"USD",amount:241296.3,dueDate:"12/05/2026"},{invoice:"SW20260100000594",discDate:"16/03/2026",obligor:"CHINA TIANCHEN ENGINEERING",ccy:"USD",amount:144000.0,dueDate:"12/05/2026"},{invoice:"SW20260100000595",discDate:"16/03/2026",obligor:"CHINA TIANCHEN ENGINEERING",ccy:"USD",amount:180000.0,dueDate:"12/05/2026"},{invoice:"SW20260100000596",discDate:"16/03/2026",obligor:"CHINA TIANCHEN ENGINEERING",ccy:"USD",amount:180000.0,dueDate:"12/05/2026"},{invoice:"SW20260100000597",discDate:"16/03/2026",obligor:"CHINA TIANCHEN ENGINEERING",ccy:"USD",amount:180000.0,dueDate:"12/05/2026"},{invoice:"SD20260000000091",discDate:"24/03/2026",obligor:"GUARDIAN EUROPE SARL (GOOLE DIV.)",ccy:"GBP",amount:65015.35,dueDate:"13/05/2026"},{invoice:"IM20260000000125",discDate:"30/03/2026",obligor:"GUARDIAN EUROPE (S.A.R.L)",ccy:"EUR",amount:84957.92,dueDate:"13/05/2026"},{invoice:"SW20260100000263",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:50100.0,dueDate:"14/05/2026"},{invoice:"SW20260100000278",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:50100.0,dueDate:"14/05/2026"},{invoice:"IM20260000000062",discDate:"19/03/2026",obligor:"SAINT-GOBAIN SEKURIT DEUTSCHLA",ccy:"EUR",amount:121468.8,dueDate:"15/05/2026"},{invoice:"SD20260000000043",discDate:"24/03/2026",obligor:"SAINT GOBAIN UK",ccy:"GBP",amount:57263.88,dueDate:"15/05/2026"},{invoice:"SD20260000000046",discDate:"24/03/2026",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:69476.4,dueDate:"15/05/2026"},{invoice:"SW20260100000510",discDate:"24/03/2026",obligor:"GUARDIAN CZESTOCHOWA SP. ZO.O.",ccy:"EUR",amount:131628.87,dueDate:"15/05/2026"},{invoice:"SW20260100000511",discDate:"24/03/2026",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:61263.37,dueDate:"15/05/2026"},{invoice:"SW20260100000437",discDate:"03/03/2026",obligor:"SISECAM FLAT GLASS SOUTH ITALY",ccy:"EUR",amount:617500.0,dueDate:"16/05/2026"},{invoice:"SW20260100000275",discDate:"04/03/2026",obligor:"Cermin DMCC",ccy:"USD",amount:237544.65,dueDate:"17/05/2026"},{invoice:"SW20260100000266",discDate:"04/03/2026",obligor:"Cermin DMCC",ccy:"USD",amount:182668.5,dueDate:"17/05/2026"},{invoice:"SW20260100000335",discDate:"04/03/2026",obligor:"Cermin DMCC",ccy:"USD",amount:175489.2,dueDate:"17/05/2026"},{invoice:"SW20260100000302",discDate:"04/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:150000.0,dueDate:"17/05/2026"},{invoice:"SW20260100000339",discDate:"04/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:204827.4,dueDate:"17/05/2026"},{invoice:"SW20260100000279",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:53400.0,dueDate:"17/05/2026"},{invoice:"SW20260100000297",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:53400.0,dueDate:"17/05/2026"},{invoice:"SW20260100000359",discDate:"31/03/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:155000.0,dueDate:"17/05/2026"},{invoice:"SW20250100003404",discDate:"16/01/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:69050.74,dueDate:"20/05/2026"},{invoice:"SW20260100000267",discDate:"19/02/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:125521.32,dueDate:"20/05/2026"},{invoice:"SD20260000000103",discDate:"30/03/2026",obligor:"GUARDIAN EUROPE SARL (GOOLE DIV.)",ccy:"GBP",amount:74904.07,dueDate:"20/05/2026"},{invoice:"IM20260000000144",discDate:"30/03/2026",obligor:"GUARDIAN EUROPE (S.A.R.L)",ccy:"EUR",amount:85242.68,dueDate:"20/05/2026"},{invoice:"SW20260100000482",discDate:"10/03/2026",obligor:"ALINDA - VELCO SA",ccy:"EUR",amount:516000.0,dueDate:"21/05/2026"},{invoice:"SD20260000000057",discDate:"09/03/2026",obligor:"ARDAGH GLASS LTD",ccy:"GBP",amount:139590.36,dueDate:"22/05/2026"},{invoice:"SD20260000000059",discDate:"09/03/2026",obligor:"ARDAGH GLASS LTD",ccy:"GBP",amount:101490.48,dueDate:"22/05/2026"},{invoice:"IM20260000000078",discDate:"19/03/2026",obligor:"SAINT-GOBAIN SEKURIT DEUTSCHLA",ccy:"EUR",amount:138868.8,dueDate:"22/05/2026"},{invoice:"SD20260000000055",discDate:"24/03/2026",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:69793.92,dueDate:"22/05/2026"},{invoice:"SD20260000000056",discDate:"24/03/2026",obligor:"ENCIRC LTD (IN LINE)",ccy:"GBP",amount:127713.6,dueDate:"22/05/2026"},{invoice:"SW20260100000657",discDate:"31/03/2026",obligor:"GUARDIAN CZESTOCHOWA SP. ZO.O.",ccy:"EUR",amount:158090.52,dueDate:"22/05/2026"},{invoice:"SW20260100000659",discDate:"31/03/2026",obligor:"GUARDIAN CZESTOCHOWA SP. ZO.O.",ccy:"EUR",amount:93686.1,dueDate:"22/05/2026"},{invoice:"SW20260100000358",discDate:"24/03/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:72520.0,dueDate:"23/05/2026"},{invoice:"SW20260100000322",discDate:"24/02/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:1076250.0,dueDate:"24/05/2026"},{invoice:"SW20260100000323",discDate:"24/02/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:410000.0,dueDate:"24/05/2026"},{invoice:"SW20260100000324",discDate:"24/02/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:265240.0,dueDate:"24/05/2026"},{invoice:"SW20260100000325",discDate:"24/02/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:530000.0,dueDate:"24/05/2026"},{invoice:"SW20260100000543",discDate:"31/03/2026",obligor:"GUARDIAN CZESTOCHOWA SP. ZO.O.",ccy:"EUR",amount:131447.97,dueDate:"24/05/2026"},{invoice:"SW20260100000607",discDate:"31/03/2026",obligor:"GUARDIAN CZESTOCHOWA SP. ZO.O.",ccy:"EUR",amount:158050.32,dueDate:"24/05/2026"},{invoice:"SW20260100000354",discDate:"04/03/2026",obligor:"Cermin DMCC",ccy:"USD",amount:179321.85,dueDate:"25/05/2026"},{invoice:"SW20260100000355",discDate:"04/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:237595.95,dueDate:"25/05/2026"},{invoice:"SW20260100000356",discDate:"04/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:140530.6,dueDate:"25/05/2026"},{invoice:"SW20260100000361",discDate:"24/03/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:155000.0,dueDate:"25/05/2026"},{invoice:"SW20260100000406",discDate:"04/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:179192.25,dueDate:"26/05/2026"},{invoice:"SW20260100000409",discDate:"04/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:179211.15,dueDate:"26/05/2026"},{invoice:"SW20260100000362",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:50100.0,dueDate:"26/05/2026"},{invoice:"SD20260000000112",discDate:"30/03/2026",obligor:"GUARDIAN EUROPE SARL (GOOLE DIV.)",ccy:"GBP",amount:58075.7,dueDate:"27/05/2026"},{invoice:"IM20260000000161",discDate:"30/03/2026",obligor:"GUARDIAN EUROPE (S.A.R.L)",ccy:"EUR",amount:91805.72,dueDate:"27/05/2026"},{invoice:"SD20260000000065",discDate:"09/03/2026",obligor:"ARDAGH GLASS LTD",ccy:"GBP",amount:212224.32,dueDate:"29/05/2026"},{invoice:"SD20260000000070",discDate:"19/03/2026",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:93673.44,dueDate:"29/05/2026"},{invoice:"IM20260000000088",discDate:"19/03/2026",obligor:"SAINT-GOBAIN SEKURIT DEUTSCHLA",ccy:"EUR",amount:108720.0,dueDate:"29/05/2026"},{invoice:"SD20260000000066",discDate:"24/03/2026",obligor:"SAINT GOBAIN UK",ccy:"GBP",amount:57150.72,dueDate:"29/05/2026"},{invoice:"SD20260000000069",discDate:"24/03/2026",obligor:"ENCIRC LTD (IN LINE)",ccy:"GBP",amount:140006.16,dueDate:"29/05/2026"},{invoice:"SW20260100000748",discDate:"31/03/2026",obligor:"GUARDIAN CZESTOCHOWA SP. ZO.O.",ccy:"EUR",amount:124690.35,dueDate:"29/05/2026"},{invoice:"SW20260100000357",discDate:"24/03/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:310000.0,dueDate:"30/05/2026"},{invoice:"SW20260100000422",discDate:"30/03/2026",obligor:"MANUCHAR NV",ccy:"USD",amount:82500.0,dueDate:"30/05/2026"},{invoice:"IM20260000000049",discDate:"10/02/2026",obligor:"AGC FLAT GLASS CZECH A.S.",ccy:"EUR",amount:584494.6,dueDate:"31/05/2026"},{invoice:"IM20260000000050",discDate:"13/02/2026",obligor:"AGC FLAT GLASS CZECH A.S.",ccy:"EUR",amount:69899.76,dueDate:"31/05/2026"},{invoice:"SW20260100000441",discDate:"04/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:237687.75,dueDate:"01/06/2026"},{invoice:"SW20260100000444",discDate:"04/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:182830.5,dueDate:"01/06/2026"},{invoice:"SW20260100000474",discDate:"04/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:204700.5,dueDate:"01/06/2026"},{invoice:"SW20260100000466",discDate:"24/03/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:111264.0,dueDate:"01/06/2026"},{invoice:"SW20260100000467",discDate:"24/03/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:94000.0,dueDate:"01/06/2026"},{invoice:"SW20260100000439",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:53400.0,dueDate:"01/06/2026"},{invoice:"SW20260100000440",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:53400.0,dueDate:"01/06/2026"},{invoice:"SW20260100000445",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:53400.0,dueDate:"01/06/2026"},{invoice:"IM20260000000106",discDate:"19/03/2026",obligor:"SAINT-GOBAIN SEKURIT DEUTSCHLA",ccy:"EUR",amount:169910.4,dueDate:"02/06/2026"},{invoice:"SD20260000000074",discDate:"09/03/2026",obligor:"ARDAGH GLASS LTD",ccy:"GBP",amount:234707.76,dueDate:"03/06/2026"},{invoice:"SD20260000000085",discDate:"24/03/2026",obligor:"ENCIRC LTD (IN LINE)",ccy:"GBP",amount:169545.6,dueDate:"03/06/2026"},{invoice:"SD20260000000087",discDate:"24/03/2026",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:73246.32,dueDate:"03/06/2026"},{invoice:"IM20260000000175",discDate:"13/04/2026",obligor:"GUARDIAN EUROPE (S.A.R.L)",ccy:"EUR",amount:118306.48,dueDate:"03/06/2026"},{invoice:"SW20260100000483",discDate:"31/03/2026",obligor:"GOVCREST INTERNATIONAL SRL",ccy:"EUR",amount:961968.5,dueDate:"04/06/2026"},{invoice:"SW20260100000502",discDate:"24/03/2026",obligor:"Distribuidora Portland S.A.",ccy:"USD",amount:61250.0,dueDate:"07/06/2026"},{invoice:"SD20260000000088",discDate:"13/04/2026",obligor:"ARDAGH GLASS LTD",ccy:"GBP",amount:183435.84,dueDate:"07/06/2026"},{invoice:"SW20260100000546",discDate:"16/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:140091.0,dueDate:"08/06/2026"},{invoice:"SW20260100000555",discDate:"19/03/2026",obligor:"ELLERBROOK ROHSTOFFHANDEL GMBH",ccy:"USD",amount:568000.0,dueDate:"08/06/2026"},{invoice:"IM20260000000123",discDate:"31/03/2026",obligor:"ARDAGH GLASS GMBH",ccy:"EUR",amount:1066000.0,dueDate:"09/06/2026"},{invoice:"SW20260100000578",discDate:"16/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:197518.5,dueDate:"10/06/2026"},{invoice:"SW20260100000559",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:50100.0,dueDate:"10/06/2026"},{invoice:"SW20260100000560",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:50100.0,dueDate:"10/06/2026"},{invoice:"SW20260100000561",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:50100.0,dueDate:"10/06/2026"},{invoice:"SW20260100000545",discDate:"30/03/2026",obligor:"Distribuidora Portland S.A.",ccy:"USD",amount:104000.0,dueDate:"10/06/2026"},{invoice:"IM20260000000128",discDate:"19/03/2026",obligor:"SAINT-GOBAIN SEKURIT DEUTSCHLA",ccy:"EUR",amount:151065.6,dueDate:"12/06/2026"},{invoice:"SD20260000000090",discDate:"24/03/2026",obligor:"SAINT GOBAIN UK",ccy:"GBP",amount:70862.76,dueDate:"12/06/2026"},{invoice:"SD20260000000099",discDate:"24/03/2026",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:103390.56,dueDate:"12/06/2026"},{invoice:"SD20260000000097",discDate:"07/04/2026",obligor:"ENCIRC LTD (IN LINE)",ccy:"GBP",amount:109967.76,dueDate:"12/06/2026"},{invoice:"SW20260100000633",discDate:"24/03/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:53400.0,dueDate:"16/06/2026"},{invoice:"SW20260100000644",discDate:"31/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:110049.8,dueDate:"16/06/2026"},{invoice:"SW20260100000632",discDate:"31/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:150000.0,dueDate:"16/06/2026"},{invoice:"SD20260000000102",discDate:"30/03/2026",obligor:"SAINT GOBAIN UK",ccy:"GBP",amount:57401.64,dueDate:"19/06/2026"},{invoice:"SD20260000000108",discDate:"07/04/2026",obligor:"ENCIRC LTD (IN LINE)",ccy:"GBP",amount:146961.36,dueDate:"19/06/2026"},{invoice:"SW20260100000163",discDate:"06/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:77278.2,dueDate:"20/06/2026"},{invoice:"SW20260100000212",discDate:"06/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:77410.95,dueDate:"20/06/2026"},{invoice:"SW20260100000284",discDate:"06/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:77379.09,dueDate:"20/06/2026"},{invoice:"SW20260100000288",discDate:"06/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:125682.39,dueDate:"20/06/2026"},{invoice:"SW20260100000351",discDate:"06/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:77464.05,dueDate:"20/06/2026"},{invoice:"SW20260100000383",discDate:"06/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:125799.21,dueDate:"20/06/2026"},{invoice:"SW20260100000446",discDate:"16/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:77501.22,dueDate:"20/06/2026"},{invoice:"SW20260100000448",discDate:"16/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:125689.47,dueDate:"20/06/2026"},{invoice:"SW20260100000538",discDate:"16/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:91850.61,dueDate:"20/06/2026"},{invoice:"SW20260100000539",discDate:"16/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:169288.11,dueDate:"20/06/2026"},{invoice:"SW20260100000799",discDate:"10/04/2026",obligor:"ALINDA - VELCO SA",ccy:"EUR",amount:516000.0,dueDate:"20/06/2026"},{invoice:"IM20260000000084",discDate:"06/03/2026",obligor:"AGC FLAT GLASS CZECH A.S.",ccy:"EUR",amount:127423.28,dueDate:"21/06/2026"},{invoice:"SW20260100000636",discDate:"30/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:140281.4,dueDate:"22/06/2026"},{invoice:"SW20260100000700",discDate:"31/03/2026",obligor:"BERNER OY",ccy:"EUR",amount:101000.0,dueDate:"23/06/2026"},{invoice:"SW20260100000688",discDate:"31/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:201016.2,dueDate:"23/06/2026"},{invoice:"SW20260100000706",discDate:"31/03/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:310000.0,dueDate:"24/06/2026"},{invoice:"SW20260100000724",discDate:"31/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:121198.0,dueDate:"25/06/2026"},{invoice:"SD20260000000111",discDate:"30/03/2026",obligor:"SAINT GOBAIN UK",ccy:"GBP",amount:57568.92,dueDate:"26/06/2026"},{invoice:"SD20260000000116",discDate:"30/03/2026",obligor:"VERALLIA UK LTD",ccy:"GBP",amount:51357.6,dueDate:"26/06/2026"},{invoice:"SD20260000000118",discDate:"07/04/2026",obligor:"ENCIRC LTD (IN LINE)",ccy:"GBP",amount:177075.36,dueDate:"26/06/2026"},{invoice:"SW20260100000166",discDate:"30/01/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:1700000.0,dueDate:"28/06/2026"},{invoice:"SW20260100000167",discDate:"02/03/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:2210000.0,dueDate:"28/06/2026"},{invoice:"SW20260100000168",discDate:"02/03/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:255000.0,dueDate:"28/06/2026"},{invoice:"SW20260100000169",discDate:"02/03/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:510000.0,dueDate:"28/06/2026"},{invoice:"SW20260100000172",discDate:"02/03/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:425000.0,dueDate:"28/06/2026"},{invoice:"SW20260100000173",discDate:"24/03/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:170000.0,dueDate:"28/06/2026"},{invoice:"SW20260100000746",discDate:"31/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:185361.4,dueDate:"28/06/2026"},{invoice:"SW20260100000747",discDate:"31/03/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:246365.0,dueDate:"28/06/2026"},{invoice:"IM20260000000103",discDate:"06/03/2026",obligor:"AGC FLAT GLASS CZECH A.S.",ccy:"EUR",amount:322752.4,dueDate:"29/06/2026"},{invoice:"SW20260100000772",discDate:"10/04/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:768750.0,dueDate:"30/06/2026"},{invoice:"SW20260100000774",discDate:"10/04/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:188552.0,dueDate:"30/06/2026"},{invoice:"SW20260100000775",discDate:"10/04/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:159000.0,dueDate:"30/06/2026"},{invoice:"SW20260100000776",discDate:"10/04/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:159144.0,dueDate:"30/06/2026"},{invoice:"SW20260100000784",discDate:"10/04/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:212259.6,dueDate:"01/07/2026"},{invoice:"SD20260000000128",discDate:"07/04/2026",obligor:"ENCIRC LTD (IN LINE)",ccy:"GBP",amount:178133.76,dueDate:"03/07/2026"},{invoice:"SD20260000000129",discDate:"07/04/2026",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:115564.68,dueDate:"04/07/2026"},{invoice:"SD20260000000130",discDate:"07/04/2026",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:118781.21,dueDate:"04/07/2026"},{invoice:"SD20260000000131",discDate:"07/04/2026",obligor:"O-I GLASS LIMITED",ccy:"GBP",amount:114644.38,dueDate:"04/07/2026"},{invoice:"SW20260100000816",discDate:"10/04/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:155000.0,dueDate:"05/07/2026"},{invoice:"SW20260100000756",discDate:"10/04/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:155000.0,dueDate:"06/07/2026"},{invoice:"SW20260100000757",discDate:"10/04/2026",obligor:"TAM TRADING FZ-LLC",ccy:"USD",amount:155000.0,dueDate:"06/07/2026"},{invoice:"IM20260000000134",discDate:"13/03/2026",obligor:"AGC FLAT GLASS CZECH A.S.",ccy:"EUR",amount:363193.6,dueDate:"07/07/2026"},{invoice:"SW20260100000835",discDate:"10/04/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:108500.0,dueDate:"08/07/2026"},{invoice:"SW20260100000601",discDate:"19/03/2026",obligor:"SCS - COMERCIAL E SERVICOS",ccy:"USD",amount:376000.0,dueDate:"15/07/2026"},{invoice:"SW20260100000600",discDate:"19/03/2026",obligor:"SCS - COMERCIAL E SERVICOS",ccy:"USD",amount:376000.0,dueDate:"15/07/2026"},{invoice:"SW20260100000604",discDate:"19/03/2026",obligor:"SCS - COMERCIAL E SERVICOS",ccy:"USD",amount:564000.0,dueDate:"15/07/2026"},{invoice:"SW20260100000605",discDate:"19/03/2026",obligor:"SCS - COMERCIAL E SERVICOS",ccy:"USD",amount:564000.0,dueDate:"15/07/2026"},{invoice:"SW20260100000602",discDate:"30/03/2026",obligor:"SCS - COMERCIAL E SERVICOS",ccy:"USD",amount:3605840.0,dueDate:"15/07/2026"},{invoice:"SW20260100000603",discDate:"09/04/2026",obligor:"SCS - COMERCIAL E SERVICOS",ccy:"USD",amount:3605840.0,dueDate:"15/07/2026"},{invoice:"SW20260100000589",discDate:"31/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:169247.4,dueDate:"20/07/2026"},{invoice:"SW20260100000590",discDate:"31/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:91795.74,dueDate:"20/07/2026"},{invoice:"SW20260100000656",discDate:"31/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:144807.24,dueDate:"20/07/2026"},{invoice:"SW20260100000725",discDate:"31/03/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:169325.28,dueDate:"20/07/2026"},{invoice:"SW20260100000754",discDate:"10/04/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:169311.12,dueDate:"20/07/2026"},{invoice:"SW20260100000782",discDate:"10/04/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:169254.48,dueDate:"20/07/2026"},{invoice:"SW20260100000794",discDate:"10/04/2026",obligor:"GUARDIAN DO BRASIL",ccy:"USD",amount:144959.46,dueDate:"20/07/2026"},{invoice:"SW20260100000637",discDate:"31/03/2026",obligor:"NATRIO PTY LTD.",ccy:"USD",amount:108555.28,dueDate:"29/07/2026"},{invoice:"SW20260100000638",discDate:"31/03/2026",obligor:"NATRIO PTY LTD.",ccy:"USD",amount:108533.48,dueDate:"29/07/2026"},{invoice:"SW20260100000673",discDate:"31/03/2026",obligor:"NATRIO PTY LTD.",ccy:"USD",amount:226476.96,dueDate:"29/07/2026"},{invoice:"IM20260000000195",discDate:"13/04/2026",obligor:"AGC FLAT GLASS CZECH A.S.",ccy:"EUR",amount:170946.48,dueDate:"03/08/2026"},{invoice:"SW20260100000540",discDate:"16/04/2026",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:81831.33,dueDate:"16/05/2026"},{invoice:"SW20260100000542",discDate:"16/04/2026",obligor:"GUARDIAN INDUSTRIES CORP LTD",ccy:"USD",amount:50272.48,dueDate:"22/05/2026"},{invoice:"SW20260100000591",discDate:"16/04/2026",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:91857.5,dueDate:"24/05/2026"},{invoice:"SW20260100000592",discDate:"16/04/2026",obligor:"GUARDIAN INDUSTRIES CORP LTD",ccy:"USD",amount:50322.16,dueDate:"01/06/2026"},{invoice:"SW20260100000550",discDate:"16/04/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:79500.0,dueDate:"23/05/2026"},{invoice:"SW20260100000551",discDate:"16/04/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:79500.0,dueDate:"23/05/2026"},{invoice:"SW20260100000552",discDate:"16/04/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:79500.0,dueDate:"23/05/2026"},{invoice:"SW20260100000553",discDate:"16/04/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:79500.0,dueDate:"23/05/2026"},{invoice:"SW20260100000598",discDate:"16/04/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:79500.0,dueDate:"23/05/2026"},{invoice:"SW20260100000599",discDate:"16/04/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:79500.0,dueDate:"23/05/2026"},{invoice:"SW20260100000750",discDate:"16/04/2026",obligor:"INTERSAC INTERNATIONAL",ccy:"USD",amount:94000.0,dueDate:"05/07/2026"},{invoice:"SW20260100000687",discDate:"16/04/2026",obligor:"GUARDIAN INDUSTRIES CORP LTD",ccy:"USD",amount:50095.84,dueDate:"04/06/2026"},{invoice:"SW20260100000613",discDate:"16/04/2026",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:81926.74,dueDate:"27/05/2026"},{invoice:"SW20260100000615",discDate:"16/04/2026",obligor:"NEWPORT INDUSTRIES LTD",ccy:"EUR",amount:60702.0,dueDate:"08/05/2026"},{invoice:"SW20260100000738",discDate:"16/04/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:50100.0,dueDate:"01/07/2026"},{invoice:"SW20260100000654",discDate:"16/04/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:68202.4,dueDate:"24/06/2026"},{invoice:"SW20260100000661",discDate:"16/04/2026",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:61500.88,dueDate:"27/05/2026"},{invoice:"SW20260100000740",discDate:"16/04/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:79500.0,dueDate:"02/06/2026"},{invoice:"SW20260100000741",discDate:"16/04/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:79500.0,dueDate:"02/06/2026"},{invoice:"SW20260100000742",discDate:"16/04/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:79500.0,dueDate:"02/06/2026"},{invoice:"SW20260100000743",discDate:"16/04/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:79500.0,dueDate:"02/06/2026"},{invoice:"SW20260100000744",discDate:"16/04/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:79500.0,dueDate:"02/06/2026"},{invoice:"SW20260100000745",discDate:"16/04/2026",obligor:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",amount:79500.0,dueDate:"02/06/2026"},{invoice:"SW20260100000699",discDate:"16/04/2026",obligor:"CONTINENTAL INDUSTRIES GROUP I",ccy:"USD",amount:71200.0,dueDate:"12/07/2026"},{invoice:"SW20260100000749",discDate:"16/04/2026",obligor:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"EUR",amount:82068.84,dueDate:"03/06/2026"},{invoice:"SW20260100000685",discDate:"16/04/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:79625.0,dueDate:"24/06/2026"},{invoice:"SW20260100000751",discDate:"16/04/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:77500.0,dueDate:"07/07/2026"},{invoice:"SW20260100000766",discDate:"16/04/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:56250.0,dueDate:"07/07/2026"},{invoice:"SW20260100000781",discDate:"16/04/2026",obligor:"CERMIN MINERALS & LOGISTICS PT",ccy:"USD",amount:56250.0,dueDate:"12/07/2026"}];

const BNP_DEFAULT_OBLIGORS=[{name:"AGC FLAT GLASS CZECH A.S.",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"AGC INC",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"ALINDA - VELCO SA",ccy:"EUR",allianz:0,atradius:0,mercury:0},{name:"ARDAGH GLASS LTD",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"Cermin DMCC",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"CERMIN MINERALS & LOGISTICS PTE LTD",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"CHINA TIANCHEN ENGINEERING",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"CONTINENTAL INDUSTRIES GROUP INC",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"ELLERBROOK ROHSTOFFHANDEL GMBH",ccy:"GBP",allianz:0,atradius:0,mercury:0},{name:"ENCIRC LTD (IN LINE)",ccy:"EUR",allianz:0,atradius:0,mercury:0},{name:"GOVCREST INTERNATIONAL SRL",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"GUARDIAN CZESTOCHOWA SP. ZO.O.",ccy:"EUR",allianz:0,atradius:0,mercury:0},{name:"GUARDIAN DO BRASIL",ccy:"EUR",allianz:0,atradius:0,mercury:0},{name:"GUARDIAN EUROPE (S.A.R.L)",ccy:"EUR",allianz:0,atradius:183659.97,mercury:0},{name:"GUARDIAN EUROPE DIV. THALHEIM",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"GUARDIAN EUROPE SARL (GOOLE DIV.)",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"GUARDIAN INDUSTRIES CORP LTD",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"GUARDIAN INDUSTRIES RAYONG CO LTD",ccy:"EUR",allianz:0,atradius:0,mercury:0},{name:"GUARDIAN ZOUJAJ INTERNATIONAL",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"INTERSAC INTERNATIONAL",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"ITAU BBA TRADING S.A.",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"LEESON ENTERPRISE PTE.LTD.",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"MANUCHAR NV",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"NEWPORT INDUSTRIES LTD",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"O-I GLASS LIMITED",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"SAINT GOBAIN UK",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"SAINT-GOBAIN ACHATS",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"SAINT-GOBAIN GLASS DEUTSCHLAND GMBH",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"SAUDI GUARDIAN INTERNATIONAL",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"SCS - COMERCIAL E SERVICOS",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"SISECAM FLAT GLASS SOUTH ITALY",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"St Gobain Achtas",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"TAM TRADING FZ-LLC",ccy:"USD",allianz:0,atradius:0,mercury:0},{name:"UNIGLOBAL GENERAL TRADING DWC-LLC",ccy:"USD",allianz:0,atradius:0,mercury:0}];


type BnpInv={id:string,invoice:string,discDate:string,obligor:string,ccy:string,amount:number,dueDate:string};
type BnpObl={id:string,name:string,ccy:string,allianz:number,atradius:number,mercury:number};

function parseDDMMYYYY(s:string):number{if(!s)return 0;const p=s.split(/[\/\-\.]/);if(p.length===3&&p[2]?.length===4)return new Date(+p[2],+p[1]-1,+p[0]).getTime();return 0;}

// DD/MM/YYYY ↔ YYYY-MM-DD conversion helpers for HTML native date input
function ddmmyyyyToIso(s:string):string{if(!s)return "";const p=s.split(/[\/\-\.]/);if(p.length===3&&p[2]?.length===4)return `${p[2]}-${p[1].padStart(2,"0")}-${p[0].padStart(2,"0")}`;return "";}
function isoToDDMMYYYY(s:string):string{if(!s)return "";const p=s.split("-");if(p.length===3&&p[0].length===4)return `${p[2]}/${p[1]}/${p[0]}`;return s;}

// Date cell: shows DD/MM/YYYY as a friendly text display but triggers native date picker on click.
function DateCell({value,onChange,disabled}:{value:string,onChange:(v:string)=>void,disabled?:boolean}){
  const ref=useRef<HTMLInputElement>(null);
  const iso=ddmmyyyyToIso(value);
  return<div className="relative w-full">
    <input type="date" ref={ref} value={iso} onChange={e=>onChange(isoToDDMMYYYY(e.target.value))} disabled={disabled} className="w-full px-1 py-0.5 border border-gray-200 rounded text-[10px] bg-white text-center outline-none cursor-pointer font-mono focus:border-violet-400 focus:bg-violet-50/30" style={{colorScheme:"light"}}/>
  </div>;
}

// Combo box: searchable dropdown that accepts custom values.
// Click chevron → see full list. Type to filter. Press "Enter" or click "✏️ Özel: …" to accept custom text.
function ObligorCombo({value,options,onChange,disabled}:{value:string,options:string[],onChange:(v:string)=>void,disabled?:boolean}){
  const[open,setOpen]=useState(false);
  const[query,setQuery]=useState("");
  const[inputVal,setInputVal]=useState(value);
  const ref=useRef<HTMLDivElement>(null);
  useEffect(()=>{setInputVal(value);},[value]);
  useEffect(()=>{
    const onDoc=(e:MouseEvent)=>{if(ref.current&&!ref.current.contains(e.target as Node))setOpen(false);};
    document.addEventListener("mousedown",onDoc);
    return()=>document.removeEventListener("mousedown",onDoc);
  },[]);
  const q=query.toLowerCase();
  const filtered=q?options.filter(o=>o.toLowerCase().includes(q)):options;
  const isCustom=inputVal&&!options.includes(inputVal);
  const commit=(v:string)=>{onChange(v);setInputVal(v);setOpen(false);setQuery("");};
  return<div ref={ref} className="relative w-full">
    <div className="flex items-center w-full border border-gray-200 rounded bg-white hover:border-violet-300 focus-within:border-violet-400">
      <input className="flex-1 px-1 py-0.5 text-[10px] bg-transparent outline-none min-w-0" value={open?query:inputVal} placeholder={open?"Yazarak ara veya seç…":(isCustom?"★ "+inputVal:"Obligor seç")} onFocus={()=>{if(!disabled){setOpen(true);setQuery(inputVal);}}} onChange={e=>{setQuery(e.target.value);setInputVal(e.target.value);setOpen(true);}} onKeyDown={e=>{if(e.key==="Enter"){e.preventDefault();commit(query||inputVal);}if(e.key==="Escape"){setOpen(false);setInputVal(value);setQuery("");}}} disabled={disabled}/>
      <button type="button" onClick={()=>{if(!disabled){setOpen(!open);setQuery("");}}} className="px-1 text-gray-400 hover:text-violet-600 text-[10px] border-l border-gray-100" tabIndex={-1} disabled={disabled}>{open?"▲":"▼"}</button>
    </div>
    {open&&<div className="absolute left-0 right-0 top-full mt-0.5 z-40 max-h-56 overflow-y-auto bg-white border border-violet-200 rounded-lg shadow-lg text-[10px]">
      {query&&!options.some(o=>o.toLowerCase()===q)&&<div onClick={()=>commit(query)} className="px-2 py-1.5 cursor-pointer hover:bg-emerald-50 bg-emerald-50/50 text-emerald-700 font-semibold border-b border-emerald-100">✏️ Özel obligor ekle: "{query}"</div>}
      {filtered.length===0&&!query&&<div className="px-2 py-1.5 text-gray-400 italic">Listede obligor yok</div>}
      {filtered.map(o=><div key={o} onClick={()=>commit(o)} className={"px-2 py-1 cursor-pointer hover:bg-violet-50 "+(o===inputVal?"bg-violet-100 font-semibold text-violet-800":"")}>{o}</div>)}
      {query&&filtered.length===0&&<div className="px-2 py-1.5 text-gray-400 italic">"{query}" eşleşen yok — Enter ile özel olarak ekleyin</div>}
    </div>}
  </div>;
}

function BnpTrUkSheet({sm,sy,canEdit}:{sm:number,sy:number,canEdit:boolean}){
  const[invoices,setInvoices]=useState<BnpInv[]>([]);
  const[obligors,setObligors]=useState<BnpObl[]>([]);
  const[totalLimit,setTotalLimit]=useState<number>(3037000);
  const[gbpUsd,setGbpUsd]=useState<number>(1.3389);
  const[eurUsd,setEurUsd]=useState<number>(1.1554);
  const[adding,setAdding]=useState(false);
  const[sortBy,setSortBy]=useState<string>("dueDate");
  const[filterObligor,setFilterObligor]=useState("");
  const[filterCcy,setFilterCcy]=useState("");
  // Per-column filters (header-level search)
  const[fInvoice,setFInvoice]=useState("");
  const[fDiscDate,setFDiscDate]=useState("");
  const[fObligor,setFObligor]=useState("");
  const[fCcy,setFCcy]=useState("");
  const[fAmount,setFAmount]=useState("");
  const[fDueDate,setFDueDate]=useState("");
  const[pasteMode,setPasteMode]=useState(false);
  const[pasteText,setPasteText]=useState("");
  const[showLimits,setShowLimits]=useState(true);

  const invKey="bnp_truk_inv_"+sy+"_"+sm;
  const oblKey="bnp_truk_obl_"+sy+"_"+sm;
  const cfgKey="bnp_truk_cfg_"+sy+"_"+sm;

  useEffect(()=>{
    const reload=()=>{
      try{
        const si=localStorage.getItem(invKey);
        if(si){setInvoices(JSON.parse(si));}
        else{
          // Seed with default invoices (silent — don't pollute undo stack)
          const seeded=BNP_DEFAULT_INVOICES.map((d,i)=>({id:"inv_seed_"+i,...d}));
          setInvoices(seeded);
          __undoEnabled=false;try{localStorage.setItem(invKey,JSON.stringify(seeded));}catch{}__undoEnabled=true;
        }
        const so=localStorage.getItem(oblKey);
        if(so){setObligors(JSON.parse(so));}
        else{
          const seeded=BNP_DEFAULT_OBLIGORS.map((d,i)=>({id:"obl_seed_"+i,...d}));
          setObligors(seeded);
          __undoEnabled=false;try{localStorage.setItem(oblKey,JSON.stringify(seeded));}catch{}__undoEnabled=true;
        }
        const sc=localStorage.getItem(cfgKey);
        if(sc){const cfg=JSON.parse(sc);if(cfg.totalLimit!==undefined)setTotalLimit(cfg.totalLimit);if(cfg.gbpUsd!==undefined)setGbpUsd(cfg.gbpUsd);if(cfg.eurUsd!==undefined)setEurUsd(cfg.eurUsd);}
      }catch{}
    };
    reload();
    const onStorage=(e:StorageEvent)=>{if(!e.key||e.key===invKey||e.key===oblKey||e.key===cfgKey)reload();};
    window.addEventListener("storage",onStorage);
    return()=>window.removeEventListener("storage",onStorage);
  },[sm,sy]);

  const saveInv=(v:BnpInv[])=>{setInvoices(v);try{trackedSetItem(invKey,JSON.stringify(v));}catch{}};
  const saveObl=(v:BnpObl[])=>{setObligors(v);try{trackedSetItem(oblKey,JSON.stringify(v));}catch{}};
  const saveCfg=(patch:Partial<{totalLimit:number,gbpUsd:number,eurUsd:number}>)=>{try{const cur=JSON.parse(localStorage.getItem(cfgKey)||"{}");const next={totalLimit,gbpUsd,eurUsd,...cur,...patch};trackedSetItem(cfgKey,JSON.stringify(next));}catch{}};

  // === INVOICE CRUD ===
  const addInv=()=>{const nr:BnpInv={id:"inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:"",discDate:"",obligor:"",ccy:"EUR",amount:0,dueDate:""};saveInv([...invoices,nr]);setAdding(true);};
  const updInv=(id:string,f:keyof BnpInv,v:any)=>{saveInv(invoices.map(r=>r.id===id?{...r,[f]:v}:r));};
  const delInv=(id:string)=>{if(!confirm("Bu faturayı silmek istediğinize emin misiniz?"))return;saveInv(invoices.filter(r=>r.id!==id));};

  // === OBLIGOR CRUD ===
  const addObl=()=>{const nr:BnpObl={id:"obl_"+Date.now(),name:"",ccy:"USD",allianz:0,atradius:0,mercury:0};saveObl([...obligors,nr]);};
  const updObl=(id:string,f:keyof BnpObl,v:any)=>{saveObl(obligors.map(r=>r.id===id?{...r,[f]:v}:r));};
  const delObl=(id:string)=>{if(!confirm("Bu obligor limitini silmek istediğinize emin misiniz?"))return;saveObl(obligors.filter(r=>r.id!==id));};

  // === PASTE FROM EXCEL ===
  const parsePaste=()=>{
    const lines=pasteText.trim().split("\n").filter(l=>l.trim());
    const parsed:BnpInv[]=lines.map(line=>{const cols=line.split("\t");const amt=parseFloat((cols[4]||"0").replace(/[,\s]/g,""))||0;return{id:"inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:(cols[0]||"").trim(),discDate:(cols[1]||"").trim(),obligor:(cols[2]||"").trim(),ccy:(cols[3]||"EUR").trim().toUpperCase(),amount:amt,dueDate:(cols[5]||"").trim()};});
    saveInv([...invoices,...parsed]);setPasteText("");setPasteMode(false);
  };

  // === DRAG-DROP FILE IMPORT ===
  const[isDraggingFile,setIsDraggingFile]=useState(false);
  const[dropStatus,setDropStatus]=useState<string>("");
  // Load SheetJS dynamically on first use
  const ensureSheetJS=async()=>{
    if((window as any).XLSX)return (window as any).XLSX;
    return new Promise<any>((resolve,reject)=>{
      const s=document.createElement("script");
      s.src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";
      s.onload=()=>resolve((window as any).XLSX);
      s.onerror=()=>reject(new Error("XLSX yüklenemedi"));
      document.head.appendChild(s);
    });
  };
  const rowToInv=(row:any[]):BnpInv|null=>{
    if(!row||row.length<1)return null;
    const str=(v:any)=>v===null||v===undefined?"":String(v).trim();
    const dateStr=(v:any)=>{
      if(!v)return "";
      // Excel serial number to date
      if(typeof v==="number"&&v>25569&&v<200000){const d=new Date((v-25569)*86400*1000);return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;}
      if(v instanceof Date)return `${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;
      return String(v).trim();
    };
    const inv=str(row[0]);if(!inv)return null;
    const amt=typeof row[4]==="number"?row[4]:parseFloat(str(row[4]).replace(/[,\s]/g,""))||0;
    return{id:"inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:inv,discDate:dateStr(row[1]),obligor:str(row[2]),ccy:str(row[3]).toUpperCase()||"EUR",amount:amt,dueDate:dateStr(row[5])};
  };
  const handleFile=async(file:File)=>{
    try{
      setDropStatus("İşleniyor: "+file.name);
      const name=file.name.toLowerCase();
      let rows:any[][]=[];
      if(name.endsWith(".csv")||name.endsWith(".tsv")||name.endsWith(".txt")){
        const txt=await file.text();
        const sep=name.endsWith(".tsv")||txt.includes("\t")?"\t":(txt.split("\n")[0]?.includes(";")?";":",");
        rows=txt.split(/\r?\n/).filter(l=>l.trim()).map(l=>l.split(sep));
        // Skip header row if starts with common header text
        if(rows[0]&&rows[0][0]&&/invoice/i.test(String(rows[0][0])))rows=rows.slice(1);
      }else if(name.endsWith(".xlsx")||name.endsWith(".xls")){
        const XLSX=await ensureSheetJS();
        const buf=await file.arrayBuffer();
        const wb=XLSX.read(buf,{type:"array",cellDates:true});
        const ws=wb.Sheets[wb.SheetNames[0]];
        rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}) as any[][];
        // Skip header row
        if(rows[0]&&rows[0][0]&&/invoice/i.test(String(rows[0][0])))rows=rows.slice(1);
      }else{
        setDropStatus("Desteklenmeyen format. .xlsx, .csv, .tsv veya .txt kabul edilir.");
        setTimeout(()=>setDropStatus(""),3500);return;
      }
      const parsed=rows.map(rowToInv).filter((x):x is BnpInv=>!!x);
      if(parsed.length===0){setDropStatus("Dosyada okunabilir fatura bulunamadı.");setTimeout(()=>setDropStatus(""),3500);return;}
      saveInv([...invoices,...parsed]);
      setDropStatus(`✓ ${parsed.length} fatura içeri aktarıldı`);
      setTimeout(()=>setDropStatus(""),3000);
    }catch(err){
      setDropStatus("Hata: "+(err instanceof Error?err.message:"dosya okunamadı"));
      setTimeout(()=>setDropStatus(""),4000);
    }
  };
  const onDrop=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(false);const files=Array.from(e.dataTransfer.files);if(files.length>0)handleFile(files[0]);};
  const onDragOver=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(true);};
  const onDragLeave=(e:React.DragEvent)=>{if(e.currentTarget===e.target)setIsDraggingFile(false);};

  // === FORMULA REPLICAS (SUMIF equivalents) ===
  // J3 = sum of amounts where D=EUR (minus IM* EUR invoices which go to J6)
  // J4 = sum where D=USD
  // J5 = sum where D=GBP
  // J6 = sum where A starts with "IM" AND D=EUR
  const sumByCcy=(ccy:string):number=>invoices.reduce((s,r)=>s+(r.ccy===ccy?r.amount:0),0);
  const J6=invoices.reduce((s,r)=>s+(r.invoice.toUpperCase().startsWith("IM")&&r.ccy==="EUR"?r.amount:0),0);
  const J3=sumByCcy("EUR")-J6;  // EUR (excluding IM*)
  const J4=sumByCcy("USD");       // USD
  const J5=sumByCcy("GBP");       // GBP
  // J7 = J3*eurUsd + J4 + J5*gbpUsd + J6*eurUsd (total in USD)
  const J7=J3*eurUsd+J4+J5*gbpUsd+J6*eurUsd;

  // Per-obligor: sum of invoice amounts where obligor matches
  const obligorTotal=(name:string):number=>invoices.reduce((s,r)=>s+(r.obligor===name?r.amount:0),0);
  // Total limit per obligor: L+M+N
  const obligorTotalLimit=(o:BnpObl):number=>o.allianz+o.atradius+o.mercury;
  // Remaining: if ccy=USD → O-J; if EUR → O-(J*eurUsd); else blank
  const obligorRemaining=(o:BnpObl):number|null=>{const J=obligorTotal(o.name);const O=obligorTotalLimit(o);if(o.ccy==="USD")return O-J;if(o.ccy==="EUR")return O-(J*eurUsd);return null;};

  // Filtered + sorted invoice view
  const obligorNamesSet=[...new Set(invoices.map(r=>r.obligor).filter(o=>o))].sort();
  const ccySet=[...new Set(invoices.map(r=>r.ccy).filter(c=>c))].sort();
  let view=[...invoices];
  if(filterObligor)view=view.filter(r=>r.obligor===filterObligor);
  if(filterCcy)view=view.filter(r=>r.ccy===filterCcy);
  // Per-column contains filters (case-insensitive)
  const ci=(s:string)=>s.toLowerCase();
  if(fInvoice)view=view.filter(r=>ci(r.invoice).includes(ci(fInvoice)));
  if(fDiscDate)view=view.filter(r=>ci(r.discDate).includes(ci(fDiscDate)));
  if(fObligor)view=view.filter(r=>ci(r.obligor).includes(ci(fObligor)));
  if(fCcy)view=view.filter(r=>ci(r.ccy).includes(ci(fCcy)));
  if(fAmount)view=view.filter(r=>r.amount.toString().includes(fAmount.replace(/[,\s]/g,"")));
  if(fDueDate)view=view.filter(r=>ci(r.dueDate).includes(ci(fDueDate)));
  view.sort((a,b)=>{
    if(sortBy==="dueDate"||sortBy==="discDate"){const ka=sortBy as "dueDate"|"discDate";return parseDDMMYYYY(a[ka])-parseDDMMYYYY(b[ka]);}
    if(sortBy==="amount")return a.amount-b.amount;
    if(sortBy==="obligor")return a.obligor.localeCompare(b.obligor);
    if(sortBy==="invoice")return a.invoice.localeCompare(b.invoice);
    return 0;
  });

  const ccyColor=(c:string)=>c==="USD"?"text-green-600":c==="EUR"?"text-blue-600":c==="GBP"?"text-purple-600":c==="TRY"?"text-amber-700":"text-gray-600";

  const resetDefaults=()=>{
    if(!confirm("Tüm faturalar ve obligor limitleri silinip Excel default'larıyla doldurulacak. Emin misiniz?"))return;
    const si=BNP_DEFAULT_INVOICES.map((d,i)=>({id:"inv_seed_"+i,...d}));
    const so=BNP_DEFAULT_OBLIGORS.map((d,i)=>({id:"obl_seed_"+i,...d}));
    saveInv(si);saveObl(so);setTotalLimit(3037000);setGbpUsd(1.3389);setEurUsd(1.1554);saveCfg({totalLimit:3037000,gbpUsd:1.3389,eurUsd:1.1554});
  };

  return<div className="relative" onDragOver={canEdit?onDragOver:undefined} onDragLeave={canEdit?onDragLeave:undefined} onDrop={canEdit?onDrop:undefined}>
    {/* Drop overlay — visible when file is being dragged over the area */}
    {isDraggingFile&&canEdit&&<div className="absolute inset-0 z-50 bg-violet-500/20 border-4 border-dashed border-violet-500 rounded-lg flex items-center justify-center pointer-events-none">
      <div className="bg-white px-6 py-4 rounded-xl shadow-xl border border-violet-200">
        <div className="text-2xl font-bold text-violet-700 mb-1">📥 Dosyayı buraya bırakın</div>
        <div className="text-[11px] text-gray-500 text-center">.xlsx • .csv • .tsv • .txt kabul edilir</div>
      </div>
    </div>}
    {dropStatus&&<div className={"mb-3 px-3 py-2 rounded-lg text-xs font-semibold "+(dropStatus.startsWith("✓")?"bg-emerald-50 text-emerald-700 border border-emerald-200":dropStatus.startsWith("Hata")||dropStatus.startsWith("Desteklenmeyen")||dropStatus.startsWith("Dosyada")?"bg-red-50 text-red-700 border border-red-200":"bg-blue-50 text-blue-700 border border-blue-200")}>{dropStatus}</div>}

    {/* Header + controls */}
    <div className="flex justify-between items-center mb-3 flex-wrap gap-2">
      <div><span className="text-sm font-bold text-violet-700">BNP TR UK</span><span className="text-[10px] text-gray-400 ml-2">Invoice discount ledger — {MONTHS[sm]} {sy} — {invoices.length} fatura</span><span className="text-[10px] text-violet-500 ml-2 hidden sm:inline">• Excel/CSV dosyasını sürükleyip bırakabilirsiniz</span></div>
      <div className="flex gap-2 flex-wrap">
        {canEdit&&<button onClick={()=>setPasteMode(!pasteMode)} className="px-3 py-1 bg-amber-100 hover:bg-amber-200 text-amber-700 rounded text-[11px] font-semibold">Excel\'den Yapıştır</button>}
        {canEdit&&<label className="px-3 py-1 bg-blue-100 hover:bg-blue-200 text-blue-700 rounded text-[11px] font-semibold cursor-pointer">📁 Dosya Seç<input type="file" accept=".xlsx,.xls,.csv,.tsv,.txt" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f);e.target.value="";}}/></label>}
        {canEdit&&<button onClick={resetDefaults} className="px-3 py-1 bg-gray-100 hover:bg-gray-200 text-gray-600 rounded text-[11px]" title="Excel default değerlerine sıfırla">↻ Sıfırla</button>}
        <ExportBtn tableId="bnpInvTbl" name={"BNP_TR_UK_"+MONTHS[sm]+"_"+sy}/>
      </div>
    </div>

    {pasteMode&&<div className="mb-3 p-3 bg-white border border-amber-300 rounded-lg">
      <div className="text-[11px] text-gray-600 mb-2">Tab-separated — 6 sütun: Invoice | Discount Date | Obligor | CCY | Amount | Due Date</div>
      <textarea value={pasteText} onChange={e=>setPasteText(e.target.value)} className="w-full h-32 px-3 py-2 border border-gray-300 rounded text-xs font-mono outline-none resize-none" placeholder={"IM20250000000498  08/12/2025  GUARDIAN EUROPE  EUR  92191.27  23/01/2026\n(her sütun arasında TAB, her satır ayrı fatura)"}/>
      <div className="flex gap-2 mt-2"><button onClick={parsePaste} className="px-3 py-1 bg-emerald-600 text-white rounded text-[11px] font-semibold">İçeri Aktar ({pasteText.trim().split("\n").filter(l=>l.trim()).length} satır)</button><button onClick={()=>{setPasteMode(false);setPasteText("");}} className="px-3 py-1 bg-gray-200 rounded text-[11px]">İptal</button></div>
    </div>}

    {/* Summary strip — currency totals and total in USD */}
    <div className="mb-4 p-3 bg-gradient-to-r from-violet-50 via-white to-violet-50 border border-violet-200 rounded-lg">
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <div>
          <div className="text-[9px] text-gray-500 uppercase">EUR Toplam (J3)</div>
          <div className="text-[11px] font-mono font-bold text-blue-600">{fmt(J3)}</div>
          <div className="text-[8px] text-gray-400">hariç IM*</div>
        </div>
        <div>
          <div className="text-[9px] text-gray-500 uppercase">USD Toplam (J4)</div>
          <div className="text-[11px] font-mono font-bold text-green-600">{fmt(J4)}</div>
        </div>
        <div>
          <div className="text-[9px] text-gray-500 uppercase">GBP Toplam (J5)</div>
          <div className="text-[11px] font-mono font-bold text-purple-600">{fmt(J5)}</div>
        </div>
        <div>
          <div className="text-[9px] text-gray-500 uppercase">SW EUROPE (J6)</div>
          <div className="text-[11px] font-mono font-bold text-indigo-600">{fmt(J6)}</div>
        </div>
      </div>
      <div className="mt-3 pt-3 border-t border-violet-200 flex items-center justify-between flex-wrap gap-2">
        <div><span className="text-[10px] text-gray-500 uppercase mr-2">Total in USD</span><span className="text-lg font-mono font-bold text-violet-800">{fmt(J7)}</span></div>
        <div><span className="text-[10px] text-gray-500 uppercase mr-2">Total Limit Usage</span><span className={"text-base font-mono font-bold "+(J7>totalLimit?"text-red-600":"text-emerald-600")}>{totalLimit>0?((J7/totalLimit)*100).toFixed(1):"0.0"}%</span></div>
      </div>
    </div>

    {/* Filters + Add Invoice (right side) */}
    {invoices.length>0&&<div className="flex gap-2 mb-3 flex-wrap text-[10px] items-center">
      <div className="flex items-center gap-1"><span className="text-gray-500">Obligor:</span><select value={filterObligor} onChange={e=>setFilterObligor(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="">Tümü ({obligorNamesSet.length})</option>{obligorNamesSet.map(o=><option key={o} value={o}>{o}</option>)}</select></div>
      <div className="flex items-center gap-1"><span className="text-gray-500">CCY:</span><select value={filterCcy} onChange={e=>setFilterCcy(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="">Tümü</option>{ccySet.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
      <div className="flex items-center gap-1"><span className="text-gray-500">Sırala:</span><select value={sortBy} onChange={e=>setSortBy(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="dueDate">Due Date</option><option value="discDate">Discount Date</option><option value="obligor">Obligor</option><option value="amount">Amount</option><option value="invoice">Invoice No</option></select></div>
      {(filterObligor||filterCcy||fInvoice||fDiscDate||fObligor||fCcy||fAmount||fDueDate)&&<button onClick={()=>{setFilterObligor("");setFilterCcy("");setFInvoice("");setFDiscDate("");setFObligor("");setFCcy("");setFAmount("");setFDueDate("");}} className="px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded text-[10px]">Tüm Filtreleri Temizle</button>}
      <span className="text-gray-400">{view.length} / {invoices.length} fatura</span>
      {canEdit&&<button onClick={addInv} className="ml-auto px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[11px] font-semibold shadow-sm">+ Fatura Ekle</button>}
    </div>}

    {/* Main layout: invoices on left, obligor limits on right */}
    <div className="grid lg:grid-cols-[1fr_520px] gap-4">
      {/* LEFT: Invoice table */}
      <div>
        <table id="bnpInvTbl" className="border-collapse text-[10px] w-full" style={{minWidth:760}}>
          <thead><tr className="border-b-2 border-violet-300 bg-violet-50 sticky top-0">
            <th className="px-2 py-2 text-left text-[9px] font-semibold text-violet-800" style={{minWidth:30}}>#</th>
            <th className="px-2 py-2 text-left text-[9px] font-semibold text-violet-800" style={{minWidth:160}}>Invoice Number</th>
            <th className="px-2 py-2 text-center text-[9px] font-semibold text-violet-800" style={{minWidth:85}}>Discount<br/>Date</th>
            <th className="px-2 py-2 text-left text-[9px] font-semibold text-violet-800" style={{minWidth:200}}>Obligor</th>
            <th className="px-2 py-2 text-center text-[9px] font-semibold text-violet-800" style={{minWidth:45}}>CCY</th>
            <th className="px-2 py-2 text-right text-[9px] font-semibold text-violet-800" style={{minWidth:100}}>Invoice<br/>Amount</th>
            <th className="px-2 py-2 text-center text-[9px] font-semibold text-violet-800" style={{minWidth:85}}>Due Date</th>
            {canEdit&&<th className="w-6"></th>}
          </tr>
          <tr className="border-b border-violet-200 bg-violet-50/60">
            <th className="px-1 py-1"></th>
            <th className="px-1 py-1"><input value={fInvoice} onChange={e=>setFInvoice(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] font-mono bg-white outline-none placeholder:text-gray-300"/></th>
            <th className="px-1 py-1"><input value={fDiscDate} onChange={e=>setFDiscDate(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
            <th className="px-1 py-1"><input value={fObligor} onChange={e=>setFObligor(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
            <th className="px-1 py-1"><input value={fCcy} onChange={e=>setFCcy(e.target.value)} placeholder="🔍" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] text-center bg-white outline-none placeholder:text-gray-300"/></th>
            <th className="px-1 py-1"><input value={fAmount} onChange={e=>setFAmount(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] text-right font-mono bg-white outline-none placeholder:text-gray-300"/></th>
            <th className="px-1 py-1"><input value={fDueDate} onChange={e=>setFDueDate(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
            {canEdit&&<th className="w-6"></th>}
          </tr></thead>
          <tbody>{view.length===0?<tr><td colSpan={canEdit?8:7} className="py-10 text-center text-gray-400">Henüz fatura yok. "+ Fatura Ekle" ile başlayın.</td></tr>:view.map((r,i)=><tr key={r.id} className="border-b border-gray-100 hover:bg-violet-50/30">
            <td className="px-2 py-0.5 text-[9px] text-gray-400">{i+1}</td>
            <td className="px-1 py-0.5"><input className={inp+" font-mono text-[10px]"} value={r.invoice} onChange={e=>updInv(r.id,"invoice",e.target.value)} placeholder="IM/SW/SD…"/></td>
            <td className="px-1 py-0.5"><DateCell value={r.discDate} onChange={v=>updInv(r.id,"discDate",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><ObligorCombo value={r.obligor} options={obligors.map(o=>o.name).filter(n=>n).sort()} onChange={v=>updInv(r.id,"obligor",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><select value={r.ccy} onChange={e=>updInv(r.id,"ccy",e.target.value)} className={"w-full px-1 py-0.5 border border-gray-200 rounded text-[10px] bg-white text-center font-semibold outline-none "+ccyColor(r.ccy)}><option value="EUR">EUR</option><option value="USD">USD</option><option value="GBP">GBP</option><option value="TRY">TRY</option><option value="CHF">CHF</option></select></td>
            <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={addC(r.amount.toString())} onChange={e=>updInv(r.id,"amount",parseFloat(stripC(e.target.value))||0)} placeholder="0.00"/></td>
            <td className="px-1 py-0.5"><DateCell value={r.dueDate} onChange={v=>updInv(r.id,"dueDate",v)} disabled={!canEdit}/></td>
            {canEdit&&<td className="px-1 py-0.5"><button onClick={()=>delInv(r.id)} className="px-1 py-0.5 text-red-400 hover:bg-red-50 hover:text-red-600 rounded text-[11px]" title="Sil">×</button></td>}
          </tr>)}</tbody>
        </table>
      </div>

      {/* RIGHT: Obligor limits table (I8:P41 replica) */}
      <div className="lg:sticky lg:top-2 lg:self-start">
        <div className="rounded-lg border border-violet-200 bg-white">
          <div className="flex justify-between items-center px-3 py-2 border-b border-violet-200 bg-violet-50 cursor-pointer" onClick={()=>setShowLimits(!showLimits)}>
            <div className="text-[11px] font-bold text-violet-800">{showLimits?"▼":"▶"} Obligor Limits ({obligors.length})</div>
            {canEdit&&showLimits&&<button onClick={(e)=>{e.stopPropagation();addObl();}} className="px-2 py-0.5 bg-violet-600 hover:bg-violet-700 text-white rounded text-[9px] font-semibold">+ Obligor</button>}
          </div>
          {showLimits&&<div className="overflow-x-auto max-h-[700px] overflow-y-auto">
            <table className="border-collapse text-[9px] w-full" style={{minWidth:500}}>
              <thead className="sticky top-0"><tr className="bg-violet-100/70 border-b border-violet-200">
                <th className="px-1 py-1 text-left text-[8px] font-bold text-violet-900" style={{minWidth:130}}>Obligor</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-violet-900" style={{minWidth:65}}>J=Σinv</th>
                <th className="px-1 py-1 text-center text-[8px] font-bold text-violet-900">K</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-violet-900" title="Allianz" style={{minWidth:55}}>Alli</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-violet-900" title="Atradius" style={{minWidth:55}}>Atr</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-violet-900" title="Mercury" style={{minWidth:55}}>Mrc</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-violet-900" title="L+M+N" style={{minWidth:55}}>O=Σ</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-violet-900" title="Remaining" style={{minWidth:60}}>P=Rem</th>
                {canEdit&&<th className="w-4"></th>}
              </tr></thead>
              <tbody>{obligors.map(o=>{const J=obligorTotal(o.name);const O=obligorTotalLimit(o);const P=obligorRemaining(o);return<tr key={o.id} className="border-b border-gray-100 hover:bg-violet-50/30">
                <td className="px-1 py-0.5 text-[9px]"><input className={"w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white outline-none"} value={o.name} onChange={e=>updObl(o.id,"name",e.target.value)}/></td>
                <td className={"px-1 py-0.5 text-right font-mono text-[9px] "+(J>0?"text-violet-700 font-semibold":"text-gray-300")}>{J?fmt(J):"-"}</td>
                <td className="px-1 py-0.5"><select value={o.ccy} onChange={e=>updObl(o.id,"ccy",e.target.value)} className={"w-full px-0.5 py-0.5 border border-gray-200 rounded text-[9px] bg-white text-center font-semibold "+ccyColor(o.ccy)}><option value="USD">USD</option><option value="EUR">EUR</option><option value="GBP">GBP</option></select></td>
                <td className="px-0.5 py-0.5"><input type="number" className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white text-right font-mono outline-none" value={o.allianz||""} onChange={e=>updObl(o.id,"allianz",parseFloat(e.target.value)||0)}/></td>
                <td className="px-0.5 py-0.5"><input type="number" className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white text-right font-mono outline-none" value={o.atradius||""} onChange={e=>updObl(o.id,"atradius",parseFloat(e.target.value)||0)}/></td>
                <td className="px-0.5 py-0.5"><input type="number" className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white text-right font-mono outline-none" value={o.mercury||""} onChange={e=>updObl(o.id,"mercury",parseFloat(e.target.value)||0)}/></td>
                <td className={"px-1 py-0.5 text-right font-mono text-[9px] "+(O>0?"font-bold text-emerald-700":"text-gray-300")}>{O?fmt(O):"-"}</td>
                <td className={"px-1 py-0.5 text-right font-mono text-[9px] font-semibold "+(P===null?"text-gray-300":P<0?"text-red-600":"text-emerald-600")}>{P===null?"—":(P<0?"("+fmt(Math.abs(P))+")":fmt(P))}</td>
                {canEdit&&<td className="px-0.5 py-0.5"><button onClick={()=>delObl(o.id)} className="px-0.5 py-0 text-red-400 hover:text-red-600 text-[10px]" title="Sil">×</button></td>}
              </tr>;})}</tbody>
            </table>
          </div>}
          <div className="px-3 py-2 bg-violet-50/40 border-t border-violet-200 text-[9px] text-violet-700 space-y-0.5">
            <div><b>J</b> = ΣInvoice by obligor (SUMIF)</div>
            <div><b>O</b> = Allianz + Atradius + Mercury</div>
            <div><b>P (Remaining)</b> = USD ? O-J : EUR ? O-(J×{eurUsd.toFixed(4)}) : —</div>
          </div>
        </div>
      </div>
    </div>
  </div>;
}




// ═══════════════════════════════════════════════════════════════════════════
// ABC — invoice discount sheet (same structure as BNP TR UK but with:
//   - 90% amount column (Discounted Amount = 100% × 0.9)
//   - Company column (IMP or SW)
//   - Obligor limits: Allianz + Atradius only (no Mercury)
//   - Parite (EUR/USD) used for "in USD" totals
// ═══════════════════════════════════════════════════════════════════════════

const ABC_DEFAULT_INVOICES=[{invoice:"SW20260100000292",discDate:"16/02/2026",obligor:"SOC. CHIMICA EMILIO FEDELI SPA",ccy:"EUR",amount:548000.0,dueDate:"11/04/2026",company:"SW"},{invoice:"IM20250000000563",discDate:"23/12/2025",obligor:"AGC GLASS EUROPE SA",ccy:"EUR",amount:307500.0,dueDate:"17/04/2026",company:"IMP"},{invoice:"IM20250000000576",discDate:"31/12/2025",obligor:"AGC GLASS EUROPE SA",ccy:"EUR",amount:820000.0,dueDate:"22/04/2026",company:"IMP"},{invoice:"SW20260100000418",discDate:"02/03/2026",obligor:"SOC. CHIMICA EMILIO FEDELI SPA",ccy:"EUR",amount:190000.0,dueDate:"24/04/2026",company:"SW"},{invoice:"SW20250100003339",discDate:"30/12/2025",obligor:"AGC GLASS EUROPE SA",ccy:"EUR",amount:1476000.0,dueDate:"25/04/2026",company:"SW"},{invoice:"SW20260100000226",discDate:"10/03/2026",obligor:"FOSFITALIA",ccy:"EUR",amount:67800.0,dueDate:"25/04/2026",company:"SW"},{invoice:"SW20260100000153",discDate:"02/03/2026",obligor:"AGC GLASS EUROPE SA",ccy:"EUR",amount:615000.0,dueDate:"26/04/2026",company:"SW"},{invoice:"SW20250100003353",discDate:"30/12/2025",obligor:"SAISA",ccy:"EUR",amount:1118295.0,dueDate:"28/04/2026",company:"SW"},{invoice:"SW20250100003355",discDate:"30/12/2025",obligor:"SAISA",ccy:"EUR",amount:1593600.0,dueDate:"28/04/2026",company:"SW"},{invoice:"SW20250100003354",discDate:"31/12/2025",obligor:"SAISA",ccy:"EUR",amount:1056000.0,dueDate:"28/04/2026",company:"SW"},{invoice:"SW20250100003384",discDate:"31/12/2025",obligor:"SAISA",ccy:"EUR",amount:497020.0,dueDate:"29/04/2026",company:"SW"},{invoice:"SW20250100003386",discDate:"31/12/2025",obligor:"SAISA",ccy:"EUR",amount:1710800.0,dueDate:"29/04/2026",company:"SW"},{invoice:"SW20250100003387",discDate:"31/12/2025",obligor:"SAISA",ccy:"EUR",amount:1169703.75,dueDate:"29/04/2026",company:"SW"},{invoice:"SW20250100003388",discDate:"31/12/2025",obligor:"SAISA",ccy:"EUR",amount:1536000.0,dueDate:"29/04/2026",company:"SW"},{invoice:"SW20260100000293",discDate:"10/03/2026",obligor:"FOSFITALIA",ccy:"EUR",amount:1139285.0,dueDate:"01/05/2026",company:"SW"},{invoice:"SW20260100000513",discDate:"10/03/2026",obligor:"SOC. CHIMICA EMILIO FEDELI SPA",ccy:"EUR",amount:760000.0,dueDate:"04/05/2026",company:"SW"},{invoice:"SW20260100000512",discDate:"10/03/2026",obligor:"SOC. CHIMICA EMILIO FEDELI SPA",ccy:"EUR",amount:798000.0,dueDate:"04/05/2026",company:"SW"},{invoice:"IM20260000000104",discDate:"06/03/2026",obligor:"ENCIRC",ccy:"EUR",amount:1127000.0,dueDate:"28/05/2026",company:"IMP"},{invoice:"IM20260000000058",discDate:"16/02/2026",obligor:"AGC GLASS EUROPE SA",ccy:"EUR",amount:820000.0,dueDate:"09/06/2026",company:"IMP"},{invoice:"SW20260100000317",discDate:"16/02/2026",obligor:"SAISA",ccy:"EUR",amount:979200.0,dueDate:"11/06/2026",company:"SW"},{invoice:"SW20260100000315",discDate:"16/02/2026",obligor:"SAISA",ccy:"EUR",amount:806400.0,dueDate:"11/06/2026",company:"SW"},{invoice:"SW20260100000281",discDate:"16/02/2026",obligor:"SAISA",ccy:"EUR",amount:86840.0,dueDate:"17/06/2026",company:"SW"},{invoice:"SW20260100000280",discDate:"16/02/2026",obligor:"SAISA",ccy:"EUR",amount:72273.0,dueDate:"17/06/2026",company:"SW"},{invoice:"IM20260000000118",discDate:"06/03/2026",obligor:"AGC GLASS EUROPE SA",ccy:"EUR",amount:410000.0,dueDate:"02/07/2026",company:"IMP"},{invoice:"SW20260100000641",discDate:"25/03/2026",obligor:"AGC GLASS EUROPE SA",ccy:"EUR",amount:615000.0,dueDate:"16/06/2026",company:"SW"},{invoice:"SW20260100000642",discDate:"25/03/2026",obligor:"SOC. CHIMICA EMILIO FEDELI SPA",ccy:"EUR",amount:285000.0,dueDate:"17/05/2026",company:"SW"},{invoice:"SW20260100000643",discDate:"25/03/2026",obligor:"SOC. CHIMICA EMILIO FEDELI SPA",ccy:"EUR",amount:570000.0,dueDate:"17/05/2026",company:"SW"},{invoice:"SW20260100000645",discDate:"27/03/2026",obligor:"SOC. CHIMICA EMILIO FEDELI SPA",ccy:"EUR",amount:646000.0,dueDate:"25/05/2026",company:"SW"},{invoice:"SW20260100000646",discDate:"27/03/2026",obligor:"SOC. CHIMICA EMILIO FEDELI SPA",ccy:"EUR",amount:493200.0,dueDate:"25/05/2026",company:"SW"},{invoice:"SW20260100000647",discDate:"27/03/2026",obligor:"SOC. CHIMICA EMILIO FEDELI SPA",ccy:"EUR",amount:380000.0,dueDate:"25/05/2026",company:"SW"},{invoice:"SW20260100000648",discDate:"27/03/2026",obligor:"SOC. CHIMICA EMILIO FEDELI SPA",ccy:"EUR",amount:570000.0,dueDate:"25/05/2026",company:"SW"},{invoice:"IM20260000000211",discDate:"14/04/2026",obligor:"AGC GLASS EUROPE SA",ccy:"EUR",amount:656000.0,dueDate:"09/08/2026",company:"SW"}];

const ABC_DEFAULT_OBLIGORS=[{name:"SAISA",ccy:"EUR",allianz:6000000.0,atradius:4000000.0},{name:"ARDAGH",ccy:"EUR",allianz:1400000.0,atradius:0},{name:"AGC GLASS EUROPE SA",ccy:"EUR",allianz:29000000.0,atradius:6000000.0},{name:"SOC. CHIMICA EMILIO FEDELI",ccy:"EUR",allianz:5000000.0,atradius:7000000.0},{name:"OI DO BRASIL",ccy:"USD",allianz:3600000.0,atradius:6000000.0},{name:"AB ETI",ccy:"EUR",allianz:1200000.0,atradius:800000.0},{name:"ENCIRC LTD",ccy:"EUR",allianz:4000000.0,atradius:3000000.0},{name:"TRAXYS EUROPE S.A.",ccy:"EUR",allianz:3000000.0,atradius:2000000.0},{name:"OI NETHERLANDS",ccy:"EUR",allianz:1800000.0,atradius:1200000.0},{name:"FOSFITALIA",ccy:"EUR",allianz:1500000.0,atradius:1000000.0},{name:"VETRERIE MERIDIONALI SPA",ccy:"EUR",allianz:1200000.0,atradius:800000.0},{name:"OI ITALY",ccy:"EUR",allianz:900000.0,atradius:600000.0},{name:"OI ESTONIA",ccy:"EUR",allianz:1200000.0,atradius:800000.0},{name:"O-I FRANCE SAS",ccy:"EUR",allianz:1200000.0,atradius:800000.0},{name:"SAINT GOBAIN GLASS EGYPT",ccy:"USD",allianz:5000000.0,atradius:1000000.0}];


type AbcInv={id:string,invoice:string,discDate:string,obligor:string,ccy:string,amount:number,dueDate:string,company:string};
type AbcObl={id:string,name:string,ccy:string,allianz:number,atradius:number};

function AbcSheet({sm,sy,canEdit}:{sm:number,sy:number,canEdit:boolean}){
  const[invoices,setInvoices]=useState<AbcInv[]>([]);
  const[obligors,setObligors]=useState<AbcObl[]>([]);
  const[parite,setParite]=useState<number>(1.161);
  const[sortBy,setSortBy]=useState<string>("dueDate");
  const[filterObligor,setFilterObligor]=useState("");
  const[filterCcy,setFilterCcy]=useState("");
  const[filterCompany,setFilterCompany]=useState("");
  // Per-column filters
  const[fInvoice,setFInvoice]=useState("");
  const[fDiscDate,setFDiscDate]=useState("");
  const[fObligor,setFObligor]=useState("");
  const[fCcy,setFCcy]=useState("");
  const[fAmount,setFAmount]=useState("");
  const[f90,setF90]=useState("");
  const[fDueDate,setFDueDate]=useState("");
  const[fCompany,setFCompany]=useState("");
  const[pasteMode,setPasteMode]=useState(false);
  const[pasteText,setPasteText]=useState("");
  const[showLimits,setShowLimits]=useState(true);
  const[isDraggingFile,setIsDraggingFile]=useState(false);
  const[dropStatus,setDropStatus]=useState<string>("");

  const invKey="abc_inv_"+sy+"_"+sm;
  const oblKey="abc_obl_"+sy+"_"+sm;
  const cfgKey="abc_cfg_"+sy+"_"+sm;

  useEffect(()=>{
    const reload=()=>{
      try{
        const si=localStorage.getItem(invKey);
        if(si){setInvoices(JSON.parse(si));}
        else{
          const seeded=ABC_DEFAULT_INVOICES.map((d,i)=>({id:"abc_inv_seed_"+i,...d}));
          setInvoices(seeded);
          __undoEnabled=false;try{localStorage.setItem(invKey,JSON.stringify(seeded));}catch{}__undoEnabled=true;
        }
        const so=localStorage.getItem(oblKey);
        if(so){setObligors(JSON.parse(so));}
        else{
          const seeded=ABC_DEFAULT_OBLIGORS.map((d,i)=>({id:"abc_obl_seed_"+i,...d}));
          setObligors(seeded);
          __undoEnabled=false;try{localStorage.setItem(oblKey,JSON.stringify(seeded));}catch{}__undoEnabled=true;
        }
        const sc=localStorage.getItem(cfgKey);
        if(sc){const cfg=JSON.parse(sc);if(cfg.parite!==undefined)setParite(cfg.parite);}
      }catch{}
    };
    reload();
    const onStorage=(e:StorageEvent)=>{if(!e.key||e.key===invKey||e.key===oblKey||e.key===cfgKey)reload();};
    window.addEventListener("storage",onStorage);
    return()=>window.removeEventListener("storage",onStorage);
  },[sm,sy]);

  const saveInv=(v:AbcInv[])=>{setInvoices(v);try{trackedSetItem(invKey,JSON.stringify(v));}catch{}};
  const saveObl=(v:AbcObl[])=>{setObligors(v);try{trackedSetItem(oblKey,JSON.stringify(v));}catch{}};
  const saveCfg=(patch:Partial<{parite:number}>)=>{try{const cur=JSON.parse(localStorage.getItem(cfgKey)||"{}");const next={parite,...cur,...patch};trackedSetItem(cfgKey,JSON.stringify(next));}catch{}};

  const addInv=()=>{const nr:AbcInv={id:"abc_inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:"",discDate:"",obligor:"",ccy:"EUR",amount:0,dueDate:"",company:"SW"};saveInv([...invoices,nr]);};
  const updInv=(id:string,f:keyof AbcInv,v:any)=>{saveInv(invoices.map(r=>r.id===id?{...r,[f]:v}:r));};
  const delInv=(id:string)=>{if(!confirm("Bu faturayı silmek istediğinize emin misiniz?"))return;saveInv(invoices.filter(r=>r.id!==id));};

  const addObl=()=>{const nr:AbcObl={id:"abc_obl_"+Date.now(),name:"",ccy:"EUR",allianz:0,atradius:0};saveObl([...obligors,nr]);};
  const updObl=(id:string,f:keyof AbcObl,v:any)=>{saveObl(obligors.map(r=>r.id===id?{...r,[f]:v}:r));};
  const delObl=(id:string)=>{if(!confirm("Bu obligor limitini silmek istediğinize emin misiniz?"))return;saveObl(obligors.filter(r=>r.id!==id));};

  // Paste from Excel (tab-separated, 7 cols: invoice, discDate, obligor, ccy, amount, dueDate, company)
  const parsePaste=()=>{
    const lines=pasteText.trim().split("\n").filter(l=>l.trim());
    const parsed:AbcInv[]=lines.map(line=>{const cols=line.split("\t");const amt=parseFloat((cols[4]||"0").replace(/[,\s]/g,""))||0;return{id:"abc_inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:(cols[0]||"").trim(),discDate:(cols[1]||"").trim(),obligor:(cols[2]||"").trim(),ccy:(cols[3]||"EUR").trim().toUpperCase(),amount:amt,dueDate:(cols[5]||"").trim(),company:(cols[6]||"SW").trim().toUpperCase()};});
    saveInv([...invoices,...parsed]);setPasteText("");setPasteMode(false);
  };

  // Drag-drop file import
  const ensureSheetJS=async()=>{
    if((window as any).XLSX)return (window as any).XLSX;
    return new Promise<any>((resolve,reject)=>{const s=document.createElement("script");s.src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";s.onload=()=>resolve((window as any).XLSX);s.onerror=()=>reject(new Error("XLSX yüklenemedi"));document.head.appendChild(s);});
  };
  const rowToInv=(row:any[]):AbcInv|null=>{
    if(!row||row.length<1)return null;
    const str=(v:any)=>v===null||v===undefined?"":String(v).trim();
    const dateStr=(v:any)=>{if(!v)return "";if(typeof v==="number"&&v>25569&&v<200000){const d=new Date((v-25569)*86400*1000);return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;}if(v instanceof Date)return `${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;return String(v).trim();};
    const inv=str(row[0]);if(!inv)return null;
    const amt=typeof row[4]==="number"?row[4]:parseFloat(str(row[4]).replace(/[,\s]/g,""))||0;
    // Column F (index 5) may be 90% amount (skip it, auto-computed) — due date is col G (index 6), company col H (index 7)
    // But CSV might be only 7 cols without 90% col. Detect by trying col 6 as date, col 7 as company
    const maybeDate=str(row[6]);
    const isDate=/\d/.test(maybeDate);
    const dueIdx=isDate?6:5;
    const compIdx=isDate?7:6;
    return{id:"abc_inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:inv,discDate:dateStr(row[1]),obligor:str(row[2]),ccy:str(row[3]).toUpperCase()||"EUR",amount:amt,dueDate:dateStr(row[dueIdx]),company:str(row[compIdx]).toUpperCase()||"SW"};
  };
  const handleFile=async(file:File)=>{
    try{
      setDropStatus("İşleniyor: "+file.name);
      const name=file.name.toLowerCase();
      let rows:any[][]=[];
      if(name.endsWith(".csv")||name.endsWith(".tsv")||name.endsWith(".txt")){
        const txt=await file.text();
        const sep=name.endsWith(".tsv")||txt.includes("\t")?"\t":(txt.split("\n")[0]?.includes(";")?";":",");
        rows=txt.split(/\r?\n/).filter(l=>l.trim()).map(l=>l.split(sep));
        if(rows[0]&&rows[0][0]&&/inv/i.test(String(rows[0][0])))rows=rows.slice(1);
      }else if(name.endsWith(".xlsx")||name.endsWith(".xls")){
        const XLSX=await ensureSheetJS();
        const buf=await file.arrayBuffer();
        const wb=XLSX.read(buf,{type:"array",cellDates:true});
        const ws=wb.Sheets[wb.SheetNames[0]];
        rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}) as any[][];
        // ABC Excel has empty row 1, headers in row 2; skip until data
        while(rows.length>0&&(!rows[0]||!rows[0][0]||/inv/i.test(String(rows[0][0]))||String(rows[0][0]).trim()===""))rows=rows.slice(1);
      }else{setDropStatus("Desteklenmeyen format. .xlsx, .csv, .tsv veya .txt kabul edilir.");setTimeout(()=>setDropStatus(""),3500);return;}
      const parsed=rows.map(rowToInv).filter((x):x is AbcInv=>!!x);
      if(parsed.length===0){setDropStatus("Dosyada okunabilir fatura bulunamadı.");setTimeout(()=>setDropStatus(""),3500);return;}
      saveInv([...invoices,...parsed]);
      setDropStatus(`✓ ${parsed.length} fatura içeri aktarıldı`);
      setTimeout(()=>setDropStatus(""),3000);
    }catch(err){setDropStatus("Hata: "+(err instanceof Error?err.message:"dosya okunamadı"));setTimeout(()=>setDropStatus(""),4000);}
  };
  const onDrop=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(false);const files=Array.from(e.dataTransfer.files);if(files.length>0)handleFile(files[0]);};
  const onDragOver=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(true);};
  const onDragLeave=(e:React.DragEvent)=>{if(e.currentTarget===e.target)setIsDraggingFile(false);};

  // === FORMULAS ===
  // L (100%) by currency: SUMIF(D=ccy, E)
  // M (90%) by currency: SUM(F) where D=ccy = 0.9 × L
  // N IMP: SUMIF(H="IMP", F) — total of 90% amount where company=IMP
  // O SW: SUMIF(H="SW", F) — total of 90% amount where company=SW
  const amt90=(r:AbcInv)=>r.amount*0.9;
  const sumAmtByCcy=(ccy:string)=>invoices.filter(r=>r.ccy===ccy).reduce((s,r)=>s+r.amount,0);
  const sum90ByCcy=(ccy:string)=>invoices.filter(r=>r.ccy===ccy).reduce((s,r)=>s+amt90(r),0);
  const sum90ByCompany=(comp:string)=>invoices.filter(r=>r.company.toUpperCase()===comp).reduce((s,r)=>s+amt90(r),0);
  const L_EUR=sumAmtByCcy("EUR");
  const L_USD=sumAmtByCcy("USD");
  const M_EUR=sum90ByCcy("EUR");
  const M_USD=sum90ByCcy("USD");
  const N_IMP=sum90ByCompany("IMP");
  const O_SW=sum90ByCompany("SW");
  // in USD: L_USD direct, plus L_EUR × parite; same for 90% 
  const L_inUSD=L_EUR*parite+L_USD;
  const M_inUSD=M_EUR*parite+M_USD;

  // Per-obligor: SUMIF on column C (obligor name) for both 100% and 90%
  const obligorTotal=(name:string)=>invoices.filter(r=>r.obligor===name).reduce((s,r)=>s+r.amount,0);
  const obligorTotal90=(name:string)=>obligorTotal(name)*0.9;
  const obligorTotalLimit=(o:AbcObl)=>o.allianz+o.atradius;
  const obligorBalance=(o:AbcObl)=>obligorTotalLimit(o)-obligorTotal(o.name);

  // Filter & sort
  const obligorNamesSet=[...new Set(invoices.map(r=>r.obligor).filter(o=>o))].sort();
  const ccySet=[...new Set(invoices.map(r=>r.ccy).filter(c=>c))].sort();
  const compSet=[...new Set(invoices.map(r=>r.company).filter(c=>c))].sort();
  let view=[...invoices];
  if(filterObligor)view=view.filter(r=>r.obligor===filterObligor);
  if(filterCcy)view=view.filter(r=>r.ccy===filterCcy);
  if(filterCompany)view=view.filter(r=>r.company===filterCompany);
  const ci=(s:string)=>s.toLowerCase();
  if(fInvoice)view=view.filter(r=>ci(r.invoice).includes(ci(fInvoice)));
  if(fDiscDate)view=view.filter(r=>ci(r.discDate).includes(ci(fDiscDate)));
  if(fObligor)view=view.filter(r=>ci(r.obligor).includes(ci(fObligor)));
  if(fCcy)view=view.filter(r=>ci(r.ccy).includes(ci(fCcy)));
  if(fAmount)view=view.filter(r=>r.amount.toString().includes(fAmount.replace(/[,\s]/g,"")));
  if(f90)view=view.filter(r=>amt90(r).toString().includes(f90.replace(/[,\s]/g,"")));
  if(fDueDate)view=view.filter(r=>ci(r.dueDate).includes(ci(fDueDate)));
  if(fCompany)view=view.filter(r=>ci(r.company).includes(ci(fCompany)));
  view.sort((a,b)=>{
    if(sortBy==="dueDate"||sortBy==="discDate"){const ka=sortBy as "dueDate"|"discDate";return parseDDMMYYYY(a[ka])-parseDDMMYYYY(b[ka]);}
    if(sortBy==="amount")return a.amount-b.amount;
    if(sortBy==="obligor")return a.obligor.localeCompare(b.obligor);
    if(sortBy==="invoice")return a.invoice.localeCompare(b.invoice);
    if(sortBy==="company")return a.company.localeCompare(b.company);
    return 0;
  });

  const ccyColor=(c:string)=>c==="USD"?"text-green-600":c==="EUR"?"text-blue-600":c==="GBP"?"text-purple-600":c==="TRY"?"text-amber-700":"text-gray-600";
  const compColor=(c:string)=>c==="SW"?"bg-blue-50 text-blue-700 border-blue-200":c==="IMP"?"bg-amber-50 text-amber-700 border-amber-200":"bg-gray-50 text-gray-600 border-gray-200";

  const resetDefaults=()=>{
    if(!confirm("Tüm faturalar ve obligor limitleri silinip Excel default'larıyla doldurulacak. Emin misiniz?"))return;
    const si=ABC_DEFAULT_INVOICES.map((d,i)=>({id:"abc_inv_seed_"+i,...d}));
    const so=ABC_DEFAULT_OBLIGORS.map((d,i)=>({id:"abc_obl_seed_"+i,...d}));
    saveInv(si);saveObl(so);setParite(1.161);saveCfg({parite:1.161});
  };

  return<div className="relative" onDragOver={canEdit?onDragOver:undefined} onDragLeave={canEdit?onDragLeave:undefined} onDrop={canEdit?onDrop:undefined}>
    {isDraggingFile&&canEdit&&<div className="absolute inset-0 z-50 bg-pink-500/20 border-4 border-dashed border-pink-500 rounded-lg flex items-center justify-center pointer-events-none">
      <div className="bg-white px-6 py-4 rounded-xl shadow-xl border border-pink-200"><div className="text-2xl font-bold text-pink-700 mb-1">📥 Dosyayı buraya bırakın</div><div className="text-[11px] text-gray-500 text-center">.xlsx • .csv • .tsv • .txt kabul edilir</div></div>
    </div>}
    {dropStatus&&<div className={"mb-3 px-3 py-2 rounded-lg text-xs font-semibold "+(dropStatus.startsWith("✓")?"bg-emerald-50 text-emerald-700 border border-emerald-200":dropStatus.startsWith("Hata")||dropStatus.startsWith("Desteklenmeyen")||dropStatus.startsWith("Dosyada")?"bg-red-50 text-red-700 border border-red-200":"bg-blue-50 text-blue-700 border border-blue-200")}>{dropStatus}</div>}

    <div className="flex justify-between items-center mb-3 flex-wrap gap-2">
      <div><span className="text-sm font-bold text-pink-700">ABC</span><span className="text-[10px] text-gray-400 ml-2">Invoice discount ledger — {MONTHS[sm]} {sy} — {invoices.length} fatura</span><span className="text-[10px] text-pink-500 ml-2 hidden sm:inline">• Excel/CSV dosyasını sürükleyip bırakabilirsiniz</span></div>
      <div className="flex gap-2 flex-wrap">
        {canEdit&&<button onClick={()=>setPasteMode(!pasteMode)} className="px-3 py-1 bg-amber-100 hover:bg-amber-200 text-amber-700 rounded text-[11px] font-semibold">Excel\'den Yapıştır</button>}
        {canEdit&&<label className="px-3 py-1 bg-blue-100 hover:bg-blue-200 text-blue-700 rounded text-[11px] font-semibold cursor-pointer">📁 Dosya Seç<input type="file" accept=".xlsx,.xls,.csv,.tsv,.txt" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f);e.target.value="";}}/></label>}
        {canEdit&&<button onClick={resetDefaults} className="px-3 py-1 bg-gray-100 hover:bg-gray-200 text-gray-600 rounded text-[11px]" title="Excel default değerlerine sıfırla">↻ Sıfırla</button>}
        <ExportBtn tableId="abcInvTbl" name={"ABC_"+MONTHS[sm]+"_"+sy}/>
      </div>
    </div>

    {pasteMode&&<div className="mb-3 p-3 bg-white border border-amber-300 rounded-lg">
      <div className="text-[11px] text-gray-600 mb-2">Tab-separated — 7 sütun: Invoice | Discount Date | Obligor | CCY | Amount | Due Date | Company (SW/IMP)</div>
      <textarea value={pasteText} onChange={e=>setPasteText(e.target.value)} className="w-full h-32 px-3 py-2 border border-gray-300 rounded text-xs font-mono outline-none resize-none" placeholder={"SW20260100000292  16/02/2026  SOC. CHIMICA EMILIO FEDELI SPA  EUR  548000  11/04/2026  SW"}/>
      <div className="flex gap-2 mt-2"><button onClick={parsePaste} className="px-3 py-1 bg-emerald-600 text-white rounded text-[11px] font-semibold">İçeri Aktar ({pasteText.trim().split("\n").filter(l=>l.trim()).length} satır)</button><button onClick={()=>{setPasteMode(false);setPasteText("");}} className="px-3 py-1 bg-gray-200 rounded text-[11px]">İptal</button></div>
    </div>}

    {/* Summary strip */}
    <div className="mb-4 p-3 bg-gradient-to-r from-pink-50 via-white to-pink-50 border border-pink-200 rounded-lg">
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
        <div><div className="text-[9px] text-gray-500 uppercase">Parite (EUR/USD)</div>{canEdit?<input type="number" step="0.001" value={parite} onChange={e=>{const v=parseFloat(e.target.value)||0;setParite(v);saveCfg({parite:v});}} className="w-full px-1 py-0.5 border border-gray-300 rounded text-[11px] font-mono font-bold bg-white"/>:<div className="text-[11px] font-mono font-bold">{parite.toFixed(4)}</div>}</div>
        <div><div className="text-[9px] text-gray-500 uppercase">EUR Toplam 100%</div><div className="text-[11px] font-mono font-bold text-blue-600">{fmt(L_EUR)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">EUR Toplam 90%</div><div className="text-[11px] font-mono font-bold text-blue-500">{fmt(M_EUR)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">USD Toplam 100%</div><div className="text-[11px] font-mono font-bold text-green-600">{fmt(L_USD)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">IMP 90%</div><div className="text-[11px] font-mono font-bold text-amber-700">{fmt(N_IMP)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">SW 90%</div><div className="text-[11px] font-mono font-bold text-indigo-600">{fmt(O_SW)}</div></div>
      </div>
      <div className="mt-3 pt-3 border-t border-pink-200 flex items-center justify-between flex-wrap gap-2">
        <div className="flex gap-6 flex-wrap">
          <div><span className="text-[10px] text-gray-500 uppercase mr-2">Toplam 100% (USD)</span><span className="text-lg font-mono font-bold text-pink-800">{fmt(L_inUSD)}</span></div>
          <div><span className="text-[10px] text-gray-500 uppercase mr-2">Toplam 90% (USD)</span><span className="text-base font-mono font-bold text-pink-700">{fmt(M_inUSD)}</span></div>
        </div>
      </div>
    </div>

    {/* Filters */}
    {invoices.length>0&&<div className="flex gap-2 mb-3 flex-wrap text-[10px] items-center">
      <div className="flex items-center gap-1"><span className="text-gray-500">Obligor:</span><select value={filterObligor} onChange={e=>setFilterObligor(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="">Tümü ({obligorNamesSet.length})</option>{obligorNamesSet.map(o=><option key={o} value={o}>{o}</option>)}</select></div>
      <div className="flex items-center gap-1"><span className="text-gray-500">CCY:</span><select value={filterCcy} onChange={e=>setFilterCcy(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="">Tümü</option>{ccySet.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
      <div className="flex items-center gap-1"><span className="text-gray-500">Company:</span><select value={filterCompany} onChange={e=>setFilterCompany(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="">Tümü</option>{compSet.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
      <div className="flex items-center gap-1"><span className="text-gray-500">Sırala:</span><select value={sortBy} onChange={e=>setSortBy(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="dueDate">Due Date</option><option value="discDate">Discount Date</option><option value="obligor">Obligor</option><option value="amount">Amount</option><option value="invoice">Invoice No</option><option value="company">Company</option></select></div>
      {(filterObligor||filterCcy||filterCompany||fInvoice||fDiscDate||fObligor||fCcy||fAmount||f90||fDueDate||fCompany)&&<button onClick={()=>{setFilterObligor("");setFilterCcy("");setFilterCompany("");setFInvoice("");setFDiscDate("");setFObligor("");setFCcy("");setFAmount("");setF90("");setFDueDate("");setFCompany("");}} className="px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded text-[10px]">Tüm Filtreleri Temizle</button>}
      <span className="text-gray-400">{view.length} / {invoices.length} fatura</span>
      {canEdit&&<button onClick={addInv} className="ml-auto px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[11px] font-semibold shadow-sm">+ Fatura Ekle</button>}
    </div>}

    {/* Main layout: invoices on left, obligor limits on right */}
    <div className="grid lg:grid-cols-[1fr_480px] gap-4">
      <div>
        <table id="abcInvTbl" className="border-collapse text-[10px] w-full" style={{minWidth:900}}>
          <thead>
            <tr className="border-b-2 border-pink-300 bg-pink-50 sticky top-0">
              <th className="px-2 py-2 text-left text-[9px] font-semibold text-pink-800" style={{minWidth:30}}>#</th>
              <th className="px-2 py-2 text-left text-[9px] font-semibold text-pink-800" style={{minWidth:150}}>Invoice Number</th>
              <th className="px-2 py-2 text-center text-[9px] font-semibold text-pink-800" style={{minWidth:85}}>Discount<br/>Date</th>
              <th className="px-2 py-2 text-left text-[9px] font-semibold text-pink-800" style={{minWidth:180}}>Obligor</th>
              <th className="px-2 py-2 text-center text-[9px] font-semibold text-pink-800" style={{minWidth:45}}>CCY</th>
              <th className="px-2 py-2 text-right text-[9px] font-semibold text-pink-800" style={{minWidth:95}}>100%<br/>Amount</th>
              <th className="px-2 py-2 text-right text-[9px] font-semibold text-pink-800 bg-pink-100" style={{minWidth:90}}>90%<br/>Amount</th>
              <th className="px-2 py-2 text-center text-[9px] font-semibold text-pink-800" style={{minWidth:85}}>Due Date</th>
              <th className="px-2 py-2 text-center text-[9px] font-semibold text-pink-800" style={{minWidth:55}}>Company</th>
              {canEdit&&<th className="w-6"></th>}
            </tr>
            <tr className="border-b border-pink-200 bg-pink-50/60">
              <th></th>
              <th className="px-1 py-1"><input value={fInvoice} onChange={e=>setFInvoice(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] font-mono bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fDiscDate} onChange={e=>setFDiscDate(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fObligor} onChange={e=>setFObligor(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fCcy} onChange={e=>setFCcy(e.target.value)} placeholder="🔍" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] text-center bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fAmount} onChange={e=>setFAmount(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] text-right font-mono bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1 bg-pink-100/50"><input value={f90} onChange={e=>setF90(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] text-right font-mono bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fDueDate} onChange={e=>setFDueDate(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fCompany} onChange={e=>setFCompany(e.target.value)} placeholder="🔍" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] text-center bg-white outline-none placeholder:text-gray-300"/></th>
              {canEdit&&<th></th>}
            </tr>
          </thead>
          <tbody>{view.length===0?<tr><td colSpan={canEdit?10:9} className="py-10 text-center text-gray-400">Henüz fatura yok. "+ Fatura Ekle" ile başlayın veya Excel dosyasını sürükleyin.</td></tr>:view.map((r,i)=><tr key={r.id} className="border-b border-gray-100 hover:bg-pink-50/30">
            <td className="px-2 py-0.5 text-[9px] text-gray-400">{i+1}</td>
            <td className="px-1 py-0.5"><input className={inp+" font-mono text-[10px]"} value={r.invoice} onChange={e=>updInv(r.id,"invoice",e.target.value)} placeholder="IM/SW…"/></td>
            <td className="px-1 py-0.5"><DateCell value={r.discDate} onChange={v=>updInv(r.id,"discDate",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><ObligorCombo value={r.obligor} options={obligors.map(o=>o.name).filter(n=>n).sort()} onChange={v=>updInv(r.id,"obligor",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><select value={r.ccy} onChange={e=>updInv(r.id,"ccy",e.target.value)} className={"w-full px-1 py-0.5 border border-gray-200 rounded text-[10px] bg-white text-center font-semibold outline-none "+ccyColor(r.ccy)}><option value="EUR">EUR</option><option value="USD">USD</option><option value="GBP">GBP</option><option value="TRY">TRY</option></select></td>
            <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={addC(r.amount.toString())} onChange={e=>updInv(r.id,"amount",parseFloat(stripC(e.target.value))||0)} placeholder="0.00"/></td>
            <td className="px-1 py-0.5 bg-pink-50/30"><span className={"block text-right px-1 "+mn+" text-[10px] font-semibold text-pink-700"} title="100% × 0.9 (otomatik)">{fmt(amt90(r))}</span></td>
            <td className="px-1 py-0.5"><DateCell value={r.dueDate} onChange={v=>updInv(r.id,"dueDate",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><select value={r.company} onChange={e=>updInv(r.id,"company",e.target.value)} className={"w-full px-1 py-0.5 border rounded text-[10px] text-center font-semibold outline-none "+compColor(r.company)}><option value="SW">SW</option><option value="IMP">IMP</option></select></td>
            {canEdit&&<td className="px-1 py-0.5"><button onClick={()=>delInv(r.id)} className="px-1 py-0.5 text-red-400 hover:bg-red-50 hover:text-red-600 rounded text-[11px]" title="Sil">×</button></td>}
          </tr>)}</tbody>
        </table>
      </div>

      {/* RIGHT: Obligor limits (Allianz + Atradius only) */}
      <div className="lg:sticky lg:top-2 lg:self-start">
        <div className="rounded-lg border border-pink-200 bg-white">
          <div className="flex justify-between items-center px-3 py-2 border-b border-pink-200 bg-pink-50 cursor-pointer" onClick={()=>setShowLimits(!showLimits)}>
            <div className="text-[11px] font-bold text-pink-800">{showLimits?"▼":"▶"} Credit Limits ({obligors.length})</div>
            {canEdit&&showLimits&&<button onClick={(e)=>{e.stopPropagation();addObl();}} className="px-2 py-0.5 bg-pink-600 hover:bg-pink-700 text-white rounded text-[9px] font-semibold">+ Obligor</button>}
          </div>
          {showLimits&&<div className="overflow-x-auto max-h-[700px] overflow-y-auto">
            <table className="border-collapse text-[9px] w-full" style={{minWidth:460}}>
              <thead className="sticky top-0"><tr className="bg-pink-100/70 border-b border-pink-200">
                <th className="px-1 py-1 text-left text-[8px] font-bold text-pink-900" style={{minWidth:130}}>Obligor</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-pink-900" style={{minWidth:70}}>100%</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-pink-900" style={{minWidth:70}}>90%</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-pink-900" title="Allianz" style={{minWidth:65}}>Allianz</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-pink-900" title="Atradius" style={{minWidth:65}}>Atradius</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-pink-900" title="Total Limit" style={{minWidth:65}}>TOTAL</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-pink-900" title="Balance" style={{minWidth:70}}>Balance</th>
                {canEdit&&<th className="w-4"></th>}
              </tr></thead>
              <tbody>{obligors.map(o=>{const J=obligorTotal(o.name);const J90=obligorTotal90(o.name);const O=obligorTotalLimit(o);const P=obligorBalance(o);return<tr key={o.id} className="border-b border-gray-100 hover:bg-pink-50/30">
                <td className="px-1 py-0.5 text-[9px]"><input className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white outline-none" value={o.name} onChange={e=>updObl(o.id,"name",e.target.value)}/></td>
                <td className={"px-1 py-0.5 text-right font-mono text-[9px] "+(J>0?"text-pink-700 font-semibold":"text-gray-300")}>{J?fmt(J):"-"}</td>
                <td className={"px-1 py-0.5 text-right font-mono text-[9px] "+(J90>0?"text-pink-600":"text-gray-300")}>{J90?fmt(J90):"-"}</td>
                <td className="px-0.5 py-0.5"><input type="number" className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white text-right font-mono outline-none" value={o.allianz||""} onChange={e=>updObl(o.id,"allianz",parseFloat(e.target.value)||0)}/></td>
                <td className="px-0.5 py-0.5"><input type="number" className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white text-right font-mono outline-none" value={o.atradius||""} onChange={e=>updObl(o.id,"atradius",parseFloat(e.target.value)||0)}/></td>
                <td className={"px-1 py-0.5 text-right font-mono text-[9px] "+(O>0?"font-bold text-emerald-700":"text-gray-300")}>{O?fmt(O):"-"}</td>
                <td className={"px-1 py-0.5 text-right font-mono text-[9px] font-semibold "+(P<0?"text-red-600":"text-emerald-600")}>{P<0?"("+fmt(Math.abs(P))+")":fmt(P)}</td>
                {canEdit&&<td className="px-0.5 py-0.5"><button onClick={()=>delObl(o.id)} className="px-0.5 py-0 text-red-400 hover:text-red-600 text-[10px]" title="Sil">×</button></td>}
              </tr>;})}</tbody>
            </table>
          </div>}
          <div className="px-3 py-2 bg-pink-50/40 border-t border-pink-200 text-[9px] text-pink-700 space-y-0.5">
            <div><b>100%</b> = ΣInvoice by obligor (SUMIF)</div>
            <div><b>90%</b> = 100% × 0.9</div>
            <div><b>TOTAL</b> = Allianz + Atradius</div>
            <div><b>Balance</b> = TOTAL − 100% Total</div>
          </div>
        </div>
      </div>
    </div>
  </div>;
}




// ═══════════════════════════════════════════════════════════════════════════
// BNP US — invoice discount sheet (simpler than BNP TR UK)
// Only Obligor totals (SUMIF) — no Allianz/Atradius/Mercury limits
// ═══════════════════════════════════════════════════════════════════════════

const BNPUS_DEFAULT_INVOICES=[{invoice:"MX46R0168",discDate:"10/12/2025",obligor:"SAINT GOBAIN MEXICO SA DE CV",ccy:"USD",amount:25838.8,dueDate:"08/02/2026"},{invoice:"MX46R0169",discDate:"26/12/2025",obligor:"SAINT GOBAIN MEXICO SA DE CV",ccy:"USD",amount:126934.6,dueDate:"24/02/2026"},{invoice:"MV196738",discDate:"08/12/2025",obligor:"VITRO VIDRIO AUTOMOTRIZ SA DE CV",ccy:"USD",amount:57724.02,dueDate:"08/03/2026"},{invoice:"MV19A0841",discDate:"08/12/2025",obligor:"VIDRIO PLANO DE MEXICO, SA DE CV",ccy:"USD",amount:72256.86,dueDate:"08/03/2026"},{invoice:"MV196739",discDate:"09/12/2025",obligor:"VITRO VIDRIO AUTOMOTRIZ SA DE CV",ccy:"USD",amount:28823.32,dueDate:"09/03/2026"},{invoice:"MV19A0842",discDate:"09/12/2025",obligor:"VIDRIO PLANO DE MEXICO, SA DE CV",ccy:"USD",amount:28791.2,dueDate:"09/03/2026"},{invoice:"MV550391",discDate:"09/12/2025",obligor:"INDUSTRIA DEL ALCALI, S.A. DE C.V.",ccy:"USD",amount:29125.54,dueDate:"09/03/2026"},{invoice:"MV19A0843",discDate:"10/12/2025",obligor:"VIDRIO PLANO DE MEXICO, SA DE CV",ccy:"USD",amount:72055.38,dueDate:"10/03/2026"},{invoice:"MV550392",discDate:"10/12/2025",obligor:"INDUSTRIA DEL ALCALI, S.A. DE C.V.",ccy:"USD",amount:72207.22,dueDate:"10/03/2026"},{invoice:"MV550394",discDate:"12/12/2025",obligor:"INDUSTRIA DEL ALCALI, S.A. DE C.V.",ccy:"USD",amount:57599.92,dueDate:"12/03/2026"},{invoice:"MV550395",discDate:"12/12/2025",obligor:"INDUSTRIA DEL ALCALI, S.A. DE C.V.",ccy:"USD",amount:43255.42,dueDate:"12/03/2026"},{invoice:"MV603337",discDate:"02/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:66464.56,dueDate:"12/03/2026"},{invoice:"MV642164",discDate:"03/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:66430.81,dueDate:"13/03/2026"},{invoice:"MV196740",discDate:"16/12/2025",obligor:"VITRO VIDRIO AUTOMOTRIZ SA DE CV",ccy:"USD",amount:57449.54,dueDate:"16/03/2026"},{invoice:"MV196741",discDate:"16/12/2025",obligor:"VITRO VIDRIO AUTOMOTRIZ SA DE CV",ccy:"USD",amount:28994.14,dueDate:"16/03/2026"},{invoice:"MV19A0844",discDate:"16/12/2025",obligor:"VIDRIO PLANO DE MEXICO, SA DE CV",ccy:"USD",amount:57423.26,dueDate:"16/03/2026"},{invoice:"MV19A0845",discDate:"16/12/2025",obligor:"VIDRIO PLANO DE MEXICO, SA DE CV",ccy:"USD",amount:28675.86,dueDate:"16/03/2026"},{invoice:"MV550396",discDate:"16/12/2025",obligor:"INDUSTRIA DEL ALCALI, S.A. DE C.V.",ccy:"USD",amount:43632.1,dueDate:"16/03/2026"},{invoice:"MV19A0846",discDate:"17/12/2025",obligor:"VIDRIO PLANO DE MEXICO, SA DE CV",ccy:"USD",amount:86651.0,dueDate:"17/03/2026"},{invoice:"MV622286",discDate:"08/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:66367.36,dueDate:"18/03/2026"},{invoice:"MV622289",discDate:"12/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:66044.71,dueDate:"22/03/2026"},{invoice:"MV196742",discDate:"22/12/2025",obligor:"VITRO VIDRIO AUTOMOTRIZ SA DE CV",ccy:"USD",amount:57519.62,dueDate:"22/03/2026"},{invoice:"MV19A0847",discDate:"22/12/2025",obligor:"VIDRIO PLANO DE MEXICO, SA DE CV",ccy:"USD",amount:86058.24,dueDate:"22/03/2026"},{invoice:"MV550397",discDate:"22/12/2025",obligor:"INDUSTRIA DEL ALCALI, S.A. DE C.V.",ccy:"USD",amount:143007.0,dueDate:"22/03/2026"},{invoice:"MV19A0848",discDate:"23/12/2025",obligor:"VIDRIO PLANO DE MEXICO, SA DE CV",ccy:"USD",amount:86541.5,dueDate:"23/03/2026"},{invoice:"MV603343",discDate:"16/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:65882.71,dueDate:"26/03/2026"},{invoice:"MV612173",discDate:"16/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:66116.26,dueDate:"26/03/2026"},{invoice:"MV642168",discDate:"16/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:66201.31,dueDate:"26/03/2026"},{invoice:"MV603344",discDate:"17/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:26511.31,dueDate:"27/03/2026"},{invoice:"MV642169",discDate:"17/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:26593.65,dueDate:"27/03/2026"},{invoice:"MV642170",discDate:"17/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:40212.45,dueDate:"27/03/2026"},{invoice:"MV196744",discDate:"30/12/2025",obligor:"VITRO VIDRIO AUTOMOTRIZ SA DE CV",ccy:"USD",amount:57588.24,dueDate:"30/03/2026"},{invoice:"MV19A0849",discDate:"30/12/2025",obligor:"VIDRIO PLANO DE MEXICO, SA DE CV",ccy:"USD",amount:86481.64,dueDate:"30/03/2026"},{invoice:"MV19A0850",discDate:"30/12/2025",obligor:"VIDRIO PLANO DE MEXICO, SA DE CV",ccy:"USD",amount:86538.58,dueDate:"30/03/2026"},{invoice:"MV612174",discDate:"22/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:66660.32,dueDate:"01/04/2026"},{invoice:"MV603346",discDate:"23/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:79231.52,dueDate:"02/04/2026"},{invoice:"MV622290",discDate:"23/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:53573.41,dueDate:"02/04/2026"},{invoice:"MV622291",discDate:"23/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:53307.46,dueDate:"02/04/2026"},{invoice:"MV622292",discDate:"26/12/2025",obligor:"OWENS AMERICA S DE RL DE CV",ccy:"USD",amount:26566.65,dueDate:"05/04/2026"},{invoice:"00-27506",discDate:"11/02/2026",obligor:"Brenntag Latin America, Inc.",ccy:"USD",amount:1209500.0,dueDate:"12/04/2026"},{invoice:"00-27509",discDate:"11/02/2026",obligor:"Brenntag Latin America, Inc.",ccy:"USD",amount:102500.0,dueDate:"12/04/2026"},{invoice:"00-27507",discDate:"11/02/2026",obligor:"Brenntag Latin America, Inc.",ccy:"USD",amount:41000.0,dueDate:"12/04/2026"},{invoice:"00-27508",discDate:"11/02/2026",obligor:"Brenntag Latin America, Inc.",ccy:"USD",amount:41000.0,dueDate:"12/04/2026"},{invoice:"00-27460",discDate:"23/11/2025",obligor:"SCS - COMERCIAL E SERVICOS QUIMICOS LTDA",ccy:"USD",amount:932856.0,dueDate:"22/04/2026"},{invoice:"00-27452",discDate:"23/11/2025",obligor:"SCS - COMERCIAL E SERVICOS QUIMICOS LTDA",ccy:"USD",amount:4238460.0,dueDate:"22/04/2026"},{invoice:"00-27459",discDate:"23/11/2025",obligor:"SCS - COMERCIAL E SERVICOS QUIMICOS LTDA",ccy:"USD",amount:4230000.0,dueDate:"22/04/2026"},{invoice:"MV660248",discDate:"11/02/2026",obligor:"CO VIDRIERA GMBH",ccy:"USD",amount:83696.23,dueDate:"22/04/2026"},{invoice:"SW20260000000007",discDate:"21/02/2026",obligor:"Livent USA Corp.",ccy:"USD",amount:6274560.0,dueDate:"22/04/2026"},{invoice:"MV660249",discDate:"11/02/2026",obligor:"CO VIDRIERA GMBH",ccy:"USD",amount:83564.17,dueDate:"23/04/2026"},{invoice:"MV660250",discDate:"11/02/2026",obligor:"CO VIDRIERA GMBH",ccy:"USD",amount:83527.25,dueDate:"23/04/2026"},{invoice:"MV660251",discDate:"11/02/2026",obligor:"CO VIDRIERA GMBH",ccy:"USD",amount:83589.74,dueDate:"25/04/2026"},{invoice:"00-27485",discDate:"28/12/2025",obligor:"KOSAC",ccy:"USD",amount:490000.0,dueDate:"27/04/2026"},{invoice:"MV660252",discDate:"11/02/2026",obligor:"CO VIDRIERA GMBH",ccy:"USD",amount:83726.05,dueDate:"29/04/2026"},{invoice:"SW20260100000268",discDate:"11/02/2026",obligor:"PT. LAUTAN LUAS TBK",ccy:"USD",amount:1701000.0,dueDate:"30/04/2026"},{invoice:"SW20260100000269",discDate:"11/02/2026",obligor:"PT. LAUTAN LUAS TBK",ccy:"USD",amount:83106.0,dueDate:"30/04/2026"},{invoice:"MV660254",discDate:"11/02/2026",obligor:"CO VIDRIERA GMBH",ccy:"USD",amount:83646.54,dueDate:"01/05/2026"},{invoice:"MV660253",discDate:"11/02/2026",obligor:"CO VIDRIERA GMBH",ccy:"USD",amount:83599.67,dueDate:"01/05/2026"},{invoice:"MV660255",discDate:"11/02/2026",obligor:"CO VIDRIERA GMBH",ccy:"USD",amount:83317.09,dueDate:"01/05/2026"},{invoice:"SW20260100000401",discDate:"02/02/2026",obligor:"PT. LAUTAN LUAS TBK",ccy:"USD",amount:84315.0,dueDate:"03/05/2026"},{invoice:"SW20260100000403",discDate:"02/02/2026",obligor:"PT. LAUTAN LUAS TBK",ccy:"USD",amount:80685.0,dueDate:"03/05/2026"},{invoice:"SW20260100000400",discDate:"02/02/2026",obligor:"PT. LAUTAN LUAS TBK",ccy:"USD",amount:1440000.0,dueDate:"03/05/2026"},{invoice:"SW20260100000397",discDate:"02/02/2026",obligor:"KOSAC",ccy:"USD",amount:668500.0,dueDate:"03/05/2026"},{invoice:"MV660256",discDate:"11/02/2026",obligor:"CO VIDRIERA GMBH",ccy:"USD",amount:83111.19,dueDate:"06/05/2026"},{invoice:"MV660257",discDate:"11/02/2026",obligor:"CO VIDRIERA GMBH",ccy:"USD",amount:84056.91,dueDate:"07/05/2026"},{invoice:"MV660258",discDate:"11/02/2026",obligor:"CO VIDRIERA GMBH",ccy:"USD",amount:83856.69,dueDate:"07/05/2026"},{invoice:"SA20260000000048",discDate:"27/03/2026",obligor:"KOSAC",ccy:"USD",amount:2646875.0,dueDate:"08/05/2026"},{invoice:"MV660259",discDate:"11/02/2026",obligor:"CO VIDRIERA GMBH",ccy:"USD",amount:83335.55,dueDate:"09/05/2026"},{invoice:"SA202600000000001",discDate:"24/03/2026",obligor:"Vidriera Argentina S.A.",ccy:"USD",amount:2000000.0,dueDate:"22/05/2026"},{invoice:"SW20260100000402",discDate:"02/02/2026",obligor:"PT. LAUTAN LUAS TBK",ccy:"USD",amount:82500.0,dueDate:"02/06/2026"},{invoice:"SA20260000000040",discDate:"26/03/2026",obligor:"Distribuidora Portland S.A.",ccy:"USD",amount:129390.0,dueDate:"02/06/2026"},{invoice:"SA20260000000041",discDate:"26/03/2026",obligor:"Distribuidora Portland S.A.",ccy:"USD",amount:285000.0,dueDate:"02/06/2026"},{invoice:"SA20260000000047",discDate:"26/03/2026",obligor:"Livent USA Corp.",ccy:"USD",amount:2422500.0,dueDate:"02/06/2026"},{invoice:"SA20260000000046",discDate:"26/03/2026",obligor:"Livent USA Corp.",ccy:"USD",amount:712500.0,dueDate:"02/06/2026"},{invoice:"SA20260000000043",discDate:"26/03/2026",obligor:"Vidrios Lirquen S.A.",ccy:"USD",amount:820000.0,dueDate:"02/06/2026"},{invoice:"SA20260000000051",discDate:"27/03/2026",obligor:"TAIWAN GLASS IND. CORP.",ccy:"USD",amount:808000.0,dueDate:"08/06/2026"},{invoice:"00-27479",discDate:"15/12/2025",obligor:"PT. LAUTAN LUAS TBK",ccy:"USD",amount:1800000.0,dueDate:"13/06/2026"},{invoice:"00-27482",discDate:"15/12/2025",obligor:"PT. LAUTAN LUAS TBK",ccy:"USD",amount:214500.0,dueDate:"13/06/2026"},{invoice:"00-27480",discDate:"15/12/2025",obligor:"PT. LAUTAN LUAS TBK",ccy:"USD",amount:82500.0,dueDate:"13/06/2026"},{invoice:"00-27481",discDate:"15/12/2025",obligor:"PT. LAUTAN LUAS TBK",ccy:"USD",amount:1571955.0,dueDate:"13/06/2026"},{invoice:"SA20260000000035",discDate:"26/03/2026",obligor:"SCS - COMERCIAL E SERVICOS QUIMICOS LTDA",ccy:"USD",amount:3841014.32,dueDate:"26/06/2026"},{invoice:"SA20260000000034",discDate:"10/04/2026",obligor:"SCS - COMERCIAL E SERVICOS QUIMICOS LTDA",ccy:"USD",amount:3841014.32,dueDate:"26/06/2026"},{invoice:"00-27496",discDate:"08/01/2026",obligor:"KOSAC",ccy:"USD",amount:3067750.0,dueDate:"08/05/2026"},{invoice:"00-27496A",discDate:"08/01/2026",obligor:"KOSAC",ccy:"USD",amount:39375.0,dueDate:"08/05/2026"},{invoice:"SU20260100000008",discDate:"16/04/2026",obligor:"MANUCHAR PERU S.A.C",ccy:"USD",amount:1665000.0,dueDate:"26/06/2026"},{invoice:"SU20260100000002",discDate:"16/04/2026",obligor:"MANUCHAR PERU S.A.C",ccy:"USD",amount:555000.0,dueDate:"26/06/2026"},{invoice:"SU20260100000001",discDate:"16/04/2026",obligor:"MANUCHAR PERU S.A.C",ccy:"USD",amount:721440.0,dueDate:"27/05/2026"},{invoice:"SU20260100000005",discDate:"16/04/2026",obligor:"Distribuidora Portland S.A.",ccy:"USD",amount:475000.0,dueDate:"26/06/2026"},{invoice:"SU20260100000006",discDate:"16/04/2026",obligor:"Distribuidora Portland S.A.",ccy:"USD",amount:190380.0,dueDate:"26/06/2026"},{invoice:"SU20260100000004",discDate:"16/04/2026",obligor:"LIVENT USA CORP",ccy:"USD",amount:430425.0,dueDate:"27/05/2026"},{invoice:"SU20260100000003",discDate:"16/04/2026",obligor:"LIVENT USA CORP",ccy:"USD",amount:1004325.0,dueDate:"27/05/2026"}];

const BNPUS_DEFAULT_OBLIGORS=[{name:"AG QUIMICA SA DE CV",ccy:"USD"},{name:"Alchemco Philippines, Inc.",ccy:"USD"},{name:"Brenntag Latin America, Inc.",ccy:"USD"},{name:"CO VIDRIERA GMBH",ccy:"USD"},{name:"CRISA LIBBEY MEXICO S DE RL DE CV",ccy:"USD"},{name:"Cristaleria del Ecuador S.A. CRIDESA",ccy:"USD"},{name:"Cristaleria Peldar S.A.",ccy:"USD"},{name:"CRISTALERIAS DE CHILE",ccy:"USD"},{name:"Distribuidora Portland S.A.",ccy:"USD"},{name:"Goodwill Marketing",ccy:"USD"},{name:"INDUSTRIA DEL ALCALI, S.A. DE C.V.",ccy:"USD"},{name:"KOSAC",ccy:"USD"},{name:"Livent USA Corp.",ccy:"USD"},{name:"MANUCHAR PERU S.A.C",ccy:"USD"},{name:"NATRIO SAS",ccy:"USD"},{name:"OWENS AMERICA S DE RL DE CV",ccy:"USD"},{name:"Owens Illinois Peru S.A.",ccy:"USD"},{name:"PT. LAUTAN LUAS TBK",ccy:"USD"},{name:"PT. Muliaglass",ccy:"USD"},{name:"SAINT GOBAIN MEXICO SA DE CV",ccy:"USD"},{name:"SAVERGLASS S DE RL DE CV",ccy:"USD"},{name:"SCS - COMERCIAL E SERVICOS QUIMICOS LTDA",ccy:"USD"},{name:"SESODA CORPORATION",ccy:"USD"},{name:"SQM SALAR SPA",ccy:"USD"},{name:"TAIWAN GLASS IND. CORP.",ccy:"USD"},{name:"TMK CHEMICAL BHD.",ccy:"USD"},{name:"Vidriera Guatemalteca, S.A.",ccy:"USD"},{name:"VIDRIO PLANO DE MEXICO, SA DE CV",ccy:"USD"},{name:"VITRO ENVASES, S.A. DE C.V.",ccy:"USD"},{name:"VITRO VIDRIO AUTOMOTRIZ SA DE CV",ccy:"USD"},{name:"PT. Muliaglass",ccy:"USD"},{name:"Vidriera Argentina S.A.",ccy:"USD"},{name:"Pt. Asahimas Flat Glass TBK",ccy:"USD"},{name:"PT Wings Surya",ccy:"USD"},{name:"ORIENTAL SILICAS CORPORATION",ccy:"USD"},{name:"Vidrios Lirquen S.A.",ccy:"USD"}];


type BnpUsInv={id:string,invoice:string,discDate:string,obligor:string,ccy:string,amount:number,dueDate:string};
type BnpUsObl={id:string,name:string,ccy:string};

function BnpUsSheet({sm,sy,canEdit}:{sm:number,sy:number,canEdit:boolean}){
  const[invoices,setInvoices]=useState<BnpUsInv[]>([]);
  const[obligors,setObligors]=useState<BnpUsObl[]>([]);
  const[sortBy,setSortBy]=useState<string>("dueDate");
  const[filterObligor,setFilterObligor]=useState("");
  const[filterCcy,setFilterCcy]=useState("");
  const[fInvoice,setFInvoice]=useState("");
  const[fDiscDate,setFDiscDate]=useState("");
  const[fObligor,setFObligor]=useState("");
  const[fCcy,setFCcy]=useState("");
  const[fAmount,setFAmount]=useState("");
  const[fDueDate,setFDueDate]=useState("");
  const[pasteMode,setPasteMode]=useState(false);
  const[pasteText,setPasteText]=useState("");
  const[showLimits,setShowLimits]=useState(true);
  const[isDraggingFile,setIsDraggingFile]=useState(false);
  const[dropStatus,setDropStatus]=useState<string>("");

  const invKey="bnpus_inv_"+sy+"_"+sm;
  const oblKey="bnpus_obl_"+sy+"_"+sm;

  useEffect(()=>{
    const reload=()=>{
      try{
        const si=localStorage.getItem(invKey);
        if(si){setInvoices(JSON.parse(si));}
        else{const seeded=BNPUS_DEFAULT_INVOICES.map((d,i)=>({id:"bnpus_inv_seed_"+i,...d}));setInvoices(seeded);__undoEnabled=false;try{localStorage.setItem(invKey,JSON.stringify(seeded));}catch{}__undoEnabled=true;}
        const so=localStorage.getItem(oblKey);
        if(so){setObligors(JSON.parse(so));}
        else{const seeded=BNPUS_DEFAULT_OBLIGORS.map((d,i)=>({id:"bnpus_obl_seed_"+i,...d}));setObligors(seeded);__undoEnabled=false;try{localStorage.setItem(oblKey,JSON.stringify(seeded));}catch{}__undoEnabled=true;}
      }catch{}
    };
    reload();
    const onStorage=(e:StorageEvent)=>{if(!e.key||e.key===invKey||e.key===oblKey)reload();};
    window.addEventListener("storage",onStorage);
    return()=>window.removeEventListener("storage",onStorage);
  },[sm,sy]);

  const saveInv=(v:BnpUsInv[])=>{setInvoices(v);try{trackedSetItem(invKey,JSON.stringify(v));}catch{}};
  const saveObl=(v:BnpUsObl[])=>{setObligors(v);try{trackedSetItem(oblKey,JSON.stringify(v));}catch{}};

  const addInv=()=>{saveInv([...invoices,{id:"bnpus_inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:"",discDate:"",obligor:"",ccy:"USD",amount:0,dueDate:""}]);};
  const updInv=(id:string,f:keyof BnpUsInv,v:any)=>{saveInv(invoices.map(r=>r.id===id?{...r,[f]:v}:r));};
  const delInv=(id:string)=>{if(!confirm("Bu faturayı silmek istediğinize emin misiniz?"))return;saveInv(invoices.filter(r=>r.id!==id));};
  const addObl=()=>{saveObl([...obligors,{id:"bnpus_obl_"+Date.now(),name:"",ccy:"USD"}]);};
  const updObl=(id:string,f:keyof BnpUsObl,v:any)=>{saveObl(obligors.map(r=>r.id===id?{...r,[f]:v}:r));};
  const delObl=(id:string)=>{if(!confirm("Sil?"))return;saveObl(obligors.filter(r=>r.id!==id));};

  const parsePaste=()=>{
    const lines=pasteText.trim().split("\n").filter(l=>l.trim());
    const parsed:BnpUsInv[]=lines.map(line=>{const cols=line.split("\t");const amt=parseFloat((cols[4]||"0").replace(/[,\s]/g,""))||0;return{id:"bnpus_inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:(cols[0]||"").trim(),discDate:(cols[1]||"").trim(),obligor:(cols[2]||"").trim(),ccy:(cols[3]||"USD").trim().toUpperCase(),amount:amt,dueDate:(cols[5]||"").trim()};});
    saveInv([...invoices,...parsed]);setPasteText("");setPasteMode(false);
  };

  const ensureSheetJS=async()=>{if((window as any).XLSX)return (window as any).XLSX;return new Promise<any>((resolve,reject)=>{const s=document.createElement("script");s.src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";s.onload=()=>resolve((window as any).XLSX);s.onerror=()=>reject(new Error("XLSX yüklenemedi"));document.head.appendChild(s);});};
  const rowToInv=(row:any[]):BnpUsInv|null=>{
    if(!row||row.length<1)return null;
    const str=(v:any)=>v===null||v===undefined?"":String(v).trim();
    const dateStr=(v:any)=>{if(!v)return "";if(typeof v==="number"&&v>25569&&v<200000){const d=new Date((v-25569)*86400*1000);return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;}if(v instanceof Date)return `${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;return String(v).trim();};
    const inv=str(row[0]);if(!inv)return null;
    const amt=typeof row[4]==="number"?row[4]:parseFloat(str(row[4]).replace(/[,\s]/g,""))||0;
    return{id:"bnpus_inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:inv,discDate:dateStr(row[1]),obligor:str(row[2]),ccy:str(row[3]).toUpperCase()||"USD",amount:amt,dueDate:dateStr(row[5])};
  };
  const handleFile=async(file:File)=>{
    try{
      setDropStatus("İşleniyor: "+file.name);
      const name=file.name.toLowerCase();let rows:any[][]=[];
      if(name.endsWith(".csv")||name.endsWith(".tsv")||name.endsWith(".txt")){
        const txt=await file.text();const sep=name.endsWith(".tsv")||txt.includes("\t")?"\t":(txt.split("\n")[0]?.includes(";")?";":",");
        rows=txt.split(/\r?\n/).filter(l=>l.trim()).map(l=>l.split(sep));
        if(rows[0]&&rows[0][0]&&/invoice/i.test(String(rows[0][0])))rows=rows.slice(1);
      }else if(name.endsWith(".xlsx")||name.endsWith(".xls")){
        const XLSX=await ensureSheetJS();const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array",cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];
        rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}) as any[][];
        // BNP US excel: header in row 1 with invoice starting col B (index 1); shift each row by 1
        if(rows[0]&&rows[0][1]&&/invoice/i.test(String(rows[0][1]))){rows=rows.slice(1).map(r=>r.slice(1));}
        else if(rows[0]&&rows[0][0]&&/invoice/i.test(String(rows[0][0])))rows=rows.slice(1);
      }else{setDropStatus("Desteklenmeyen format.");setTimeout(()=>setDropStatus(""),3500);return;}
      const parsed=rows.map(rowToInv).filter((x):x is BnpUsInv=>!!x);
      if(parsed.length===0){setDropStatus("Fatura bulunamadı.");setTimeout(()=>setDropStatus(""),3500);return;}
      saveInv([...invoices,...parsed]);setDropStatus(`✓ ${parsed.length} fatura içeri aktarıldı`);setTimeout(()=>setDropStatus(""),3000);
    }catch(err){setDropStatus("Hata: "+(err instanceof Error?err.message:"dosya okunamadı"));setTimeout(()=>setDropStatus(""),4000);}
  };
  const onDrop=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(false);const files=Array.from(e.dataTransfer.files);if(files.length>0)handleFile(files[0]);};
  const onDragOver=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(true);};
  const onDragLeave=(e:React.DragEvent)=>{if(e.currentTarget===e.target)setIsDraggingFile(false);};

  // Totals per currency
  const sumByCcy=(ccy:string)=>invoices.filter(r=>r.ccy===ccy).reduce((s,r)=>s+r.amount,0);
  const total_USD=sumByCcy("USD");
  const total_EUR=sumByCcy("EUR");
  const total_All=invoices.reduce((s,r)=>s+r.amount,0);

  // Per-obligor
  const obligorTotal=(name:string)=>invoices.filter(r=>r.obligor===name).reduce((s,r)=>s+r.amount,0);

  const obligorNamesSet=[...new Set(invoices.map(r=>r.obligor).filter(o=>o))].sort();
  const ccySet=[...new Set(invoices.map(r=>r.ccy).filter(c=>c))].sort();
  let view=[...invoices];
  if(filterObligor)view=view.filter(r=>r.obligor===filterObligor);
  if(filterCcy)view=view.filter(r=>r.ccy===filterCcy);
  const ci=(s:string)=>s.toLowerCase();
  if(fInvoice)view=view.filter(r=>ci(r.invoice).includes(ci(fInvoice)));
  if(fDiscDate)view=view.filter(r=>ci(r.discDate).includes(ci(fDiscDate)));
  if(fObligor)view=view.filter(r=>ci(r.obligor).includes(ci(fObligor)));
  if(fCcy)view=view.filter(r=>ci(r.ccy).includes(ci(fCcy)));
  if(fAmount)view=view.filter(r=>r.amount.toString().includes(fAmount.replace(/[,\s]/g,"")));
  if(fDueDate)view=view.filter(r=>ci(r.dueDate).includes(ci(fDueDate)));
  view.sort((a,b)=>{
    if(sortBy==="dueDate"||sortBy==="discDate"){const ka=sortBy as "dueDate"|"discDate";return parseDDMMYYYY(a[ka])-parseDDMMYYYY(b[ka]);}
    if(sortBy==="amount")return a.amount-b.amount;
    if(sortBy==="obligor")return a.obligor.localeCompare(b.obligor);
    if(sortBy==="invoice")return a.invoice.localeCompare(b.invoice);
    return 0;
  });

  const ccyColor=(c:string)=>c==="USD"?"text-green-600":c==="EUR"?"text-blue-600":c==="GBP"?"text-purple-600":"text-gray-600";

  const resetDefaults=()=>{
    if(!confirm("Tüm faturalar silinip Excel default'larıyla doldurulacak. Emin misiniz?"))return;
    const si=BNPUS_DEFAULT_INVOICES.map((d,i)=>({id:"bnpus_inv_seed_"+i,...d}));
    const so=BNPUS_DEFAULT_OBLIGORS.map((d,i)=>({id:"bnpus_obl_seed_"+i,...d}));
    saveInv(si);saveObl(so);
  };

  return<div className="relative" onDragOver={canEdit?onDragOver:undefined} onDragLeave={canEdit?onDragLeave:undefined} onDrop={canEdit?onDrop:undefined}>
    {isDraggingFile&&canEdit&&<div className="absolute inset-0 z-50 bg-cyan-500/20 border-4 border-dashed border-cyan-500 rounded-lg flex items-center justify-center pointer-events-none"><div className="bg-white px-6 py-4 rounded-xl shadow-xl border border-cyan-200"><div className="text-2xl font-bold text-cyan-700 mb-1">📥 Dosyayı buraya bırakın</div><div className="text-[11px] text-gray-500 text-center">.xlsx • .csv • .tsv • .txt</div></div></div>}
    {dropStatus&&<div className={"mb-3 px-3 py-2 rounded-lg text-xs font-semibold "+(dropStatus.startsWith("✓")?"bg-emerald-50 text-emerald-700 border border-emerald-200":dropStatus.startsWith("Hata")||dropStatus.startsWith("Desteklenmeyen")||dropStatus.startsWith("Fatura")?"bg-red-50 text-red-700 border border-red-200":"bg-blue-50 text-blue-700 border border-blue-200")}>{dropStatus}</div>}

    <div className="flex justify-between items-center mb-3 flex-wrap gap-2">
      <div><span className="text-sm font-bold text-cyan-700">BNP US</span><span className="text-[10px] text-gray-400 ml-2">Invoice discount ledger — {MONTHS[sm]} {sy} — {invoices.length} fatura</span><span className="text-[10px] text-cyan-500 ml-2 hidden sm:inline">• Excel/CSV dosyasını sürükleyip bırakabilirsiniz</span></div>
      <div className="flex gap-2 flex-wrap">
        {canEdit&&<button onClick={()=>setPasteMode(!pasteMode)} className="px-3 py-1 bg-amber-100 hover:bg-amber-200 text-amber-700 rounded text-[11px] font-semibold">Excel\'den Yapıştır</button>}
        {canEdit&&<label className="px-3 py-1 bg-blue-100 hover:bg-blue-200 text-blue-700 rounded text-[11px] font-semibold cursor-pointer">📁 Dosya Seç<input type="file" accept=".xlsx,.xls,.csv,.tsv,.txt" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f);e.target.value="";}}/></label>}
        {canEdit&&<button onClick={resetDefaults} className="px-3 py-1 bg-gray-100 hover:bg-gray-200 text-gray-600 rounded text-[11px]">↻ Sıfırla</button>}
        <ExportBtn tableId="bnpusInvTbl" name={"BNP_US_"+MONTHS[sm]+"_"+sy}/>
      </div>
    </div>

    {pasteMode&&<div className="mb-3 p-3 bg-white border border-amber-300 rounded-lg">
      <div className="text-[11px] text-gray-600 mb-2">Tab-separated — 6 sütun: Invoice | Discount Date | Obligor | CCY | Amount | Due Date</div>
      <textarea value={pasteText} onChange={e=>setPasteText(e.target.value)} className="w-full h-32 px-3 py-2 border border-gray-300 rounded text-xs font-mono outline-none resize-none"/>
      <div className="flex gap-2 mt-2"><button onClick={parsePaste} className="px-3 py-1 bg-emerald-600 text-white rounded text-[11px] font-semibold">İçeri Aktar ({pasteText.trim().split("\n").filter(l=>l.trim()).length} satır)</button><button onClick={()=>{setPasteMode(false);setPasteText("");}} className="px-3 py-1 bg-gray-200 rounded text-[11px]">İptal</button></div>
    </div>}

    {/* Summary strip */}
    <div className="mb-4 p-3 bg-gradient-to-r from-cyan-50 via-white to-cyan-50 border border-cyan-200 rounded-lg">
      <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
        <div><div className="text-[9px] text-gray-500 uppercase">USD Toplam</div><div className="text-[13px] font-mono font-bold text-green-600">{fmt(total_USD)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">EUR Toplam</div><div className="text-[13px] font-mono font-bold text-blue-600">{fmt(total_EUR)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Genel Toplam</div><div className="text-lg font-mono font-bold text-cyan-800">{fmt(total_All)}</div></div>
      </div>
    </div>

    {invoices.length>0&&<div className="flex gap-2 mb-3 flex-wrap text-[10px] items-center">
      <div className="flex items-center gap-1"><span className="text-gray-500">Obligor:</span><select value={filterObligor} onChange={e=>setFilterObligor(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="">Tümü ({obligorNamesSet.length})</option>{obligorNamesSet.map(o=><option key={o} value={o}>{o}</option>)}</select></div>
      <div className="flex items-center gap-1"><span className="text-gray-500">CCY:</span><select value={filterCcy} onChange={e=>setFilterCcy(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="">Tümü</option>{ccySet.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
      <div className="flex items-center gap-1"><span className="text-gray-500">Sırala:</span><select value={sortBy} onChange={e=>setSortBy(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="dueDate">Due Date</option><option value="discDate">Discount Date</option><option value="obligor">Obligor</option><option value="amount">Amount</option><option value="invoice">Invoice No</option></select></div>
      {(filterObligor||filterCcy||fInvoice||fDiscDate||fObligor||fCcy||fAmount||fDueDate)&&<button onClick={()=>{setFilterObligor("");setFilterCcy("");setFInvoice("");setFDiscDate("");setFObligor("");setFCcy("");setFAmount("");setFDueDate("");}} className="px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded text-[10px]">Tüm Filtreleri Temizle</button>}
      <span className="text-gray-400">{view.length} / {invoices.length} fatura</span>
      {canEdit&&<button onClick={addInv} className="ml-auto px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[11px] font-semibold shadow-sm">+ Fatura Ekle</button>}
    </div>}

    <div className="grid lg:grid-cols-[1fr_380px] gap-4">
      <div>
        <table id="bnpusInvTbl" className="border-collapse text-[10px] w-full" style={{minWidth:780}}>
          <thead>
            <tr className="border-b-2 border-cyan-300 bg-cyan-50 sticky top-0">
              <th className="px-2 py-2 text-left text-[9px] font-semibold text-cyan-800" style={{minWidth:30}}>#</th>
              <th className="px-2 py-2 text-left text-[9px] font-semibold text-cyan-800" style={{minWidth:140}}>Invoice Number</th>
              <th className="px-2 py-2 text-center text-[9px] font-semibold text-cyan-800" style={{minWidth:85}}>Discount<br/>Date</th>
              <th className="px-2 py-2 text-left text-[9px] font-semibold text-cyan-800" style={{minWidth:200}}>Obligor</th>
              <th className="px-2 py-2 text-center text-[9px] font-semibold text-cyan-800" style={{minWidth:45}}>CCY</th>
              <th className="px-2 py-2 text-right text-[9px] font-semibold text-cyan-800" style={{minWidth:110}}>Invoice<br/>Amount</th>
              <th className="px-2 py-2 text-center text-[9px] font-semibold text-cyan-800" style={{minWidth:85}}>Due Date</th>
              {canEdit&&<th className="w-6"></th>}
            </tr>
            <tr className="border-b border-cyan-200 bg-cyan-50/60">
              <th></th>
              <th className="px-1 py-1"><input value={fInvoice} onChange={e=>setFInvoice(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] font-mono bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fDiscDate} onChange={e=>setFDiscDate(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fObligor} onChange={e=>setFObligor(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fCcy} onChange={e=>setFCcy(e.target.value)} placeholder="🔍" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] text-center bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fAmount} onChange={e=>setFAmount(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] text-right font-mono bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fDueDate} onChange={e=>setFDueDate(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
              {canEdit&&<th></th>}
            </tr>
          </thead>
          <tbody>{view.length===0?<tr><td colSpan={canEdit?8:7} className="py-10 text-center text-gray-400">Henüz fatura yok.</td></tr>:view.map((r,i)=><tr key={r.id} className="border-b border-gray-100 hover:bg-cyan-50/30">
            <td className="px-2 py-0.5 text-[9px] text-gray-400">{i+1}</td>
            <td className="px-1 py-0.5"><input className={inp+" font-mono text-[10px]"} value={r.invoice} onChange={e=>updInv(r.id,"invoice",e.target.value)}/></td>
            <td className="px-1 py-0.5"><DateCell value={r.discDate} onChange={v=>updInv(r.id,"discDate",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><ObligorCombo value={r.obligor} options={obligors.map(o=>o.name).filter(n=>n).sort()} onChange={v=>updInv(r.id,"obligor",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><select value={r.ccy} onChange={e=>updInv(r.id,"ccy",e.target.value)} className={"w-full px-1 py-0.5 border border-gray-200 rounded text-[10px] bg-white text-center font-semibold outline-none "+ccyColor(r.ccy)}><option value="USD">USD</option><option value="EUR">EUR</option><option value="GBP">GBP</option></select></td>
            <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={addC(r.amount.toString())} onChange={e=>updInv(r.id,"amount",parseFloat(stripC(e.target.value))||0)}/></td>
            <td className="px-1 py-0.5"><DateCell value={r.dueDate} onChange={v=>updInv(r.id,"dueDate",v)} disabled={!canEdit}/></td>
            {canEdit&&<td className="px-1 py-0.5"><button onClick={()=>delInv(r.id)} className="px-1 py-0.5 text-red-400 hover:bg-red-50 hover:text-red-600 rounded text-[11px]">×</button></td>}
          </tr>)}</tbody>
        </table>
      </div>

      <div className="lg:sticky lg:top-2 lg:self-start">
        <div className="rounded-lg border border-cyan-200 bg-white">
          <div className="flex justify-between items-center px-3 py-2 border-b border-cyan-200 bg-cyan-50 cursor-pointer" onClick={()=>setShowLimits(!showLimits)}>
            <div className="text-[11px] font-bold text-cyan-800">{showLimits?"▼":"▶"} Customer Totals ({obligors.length})</div>
            {canEdit&&showLimits&&<button onClick={(e)=>{e.stopPropagation();addObl();}} className="px-2 py-0.5 bg-cyan-600 hover:bg-cyan-700 text-white rounded text-[9px] font-semibold">+ Customer</button>}
          </div>
          {showLimits&&<div className="overflow-x-auto max-h-[700px] overflow-y-auto">
            <table className="border-collapse text-[9px] w-full" style={{minWidth:360}}>
              <thead className="sticky top-0"><tr className="bg-cyan-100/70 border-b border-cyan-200">
                <th className="px-1 py-1 text-left text-[8px] font-bold text-cyan-900" style={{minWidth:180}}>Customer Name</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-cyan-900" style={{minWidth:90}}>Invoice Amount</th>
                <th className="px-1 py-1 text-center text-[8px] font-bold text-cyan-900" style={{minWidth:40}}>CCY</th>
                {canEdit&&<th className="w-4"></th>}
              </tr></thead>
              <tbody>{obligors.map(o=>{const J=obligorTotal(o.name);return<tr key={o.id} className="border-b border-gray-100 hover:bg-cyan-50/30">
                <td className="px-1 py-0.5"><input className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white outline-none" value={o.name} onChange={e=>updObl(o.id,"name",e.target.value)}/></td>
                <td className={"px-1 py-0.5 text-right font-mono text-[9px] "+(J>0?"text-cyan-700 font-semibold":"text-gray-300")}>{J?fmt(J):"-"}</td>
                <td className="px-0.5 py-0.5"><select value={o.ccy} onChange={e=>updObl(o.id,"ccy",e.target.value)} className={"w-full px-0.5 py-0.5 border border-gray-200 rounded text-[9px] bg-white text-center font-semibold "+ccyColor(o.ccy)}><option value="USD">USD</option><option value="EUR">EUR</option><option value="GBP">GBP</option></select></td>
                {canEdit&&<td className="px-0.5 py-0.5"><button onClick={()=>delObl(o.id)} className="px-0.5 py-0 text-red-400 hover:text-red-600 text-[10px]">×</button></td>}
              </tr>;})}
              <tr className="bg-cyan-100 border-t-2 border-cyan-300 font-bold"><td className="px-1 py-1 text-[9px] text-cyan-900">Total Invoice Amount</td><td className="px-1 py-1 text-right font-mono text-[10px] text-cyan-900">{fmt(total_All)}</td><td></td>{canEdit&&<td></td>}</tr>
              </tbody>
            </table>
          </div>}
          <div className="px-3 py-2 bg-cyan-50/40 border-t border-cyan-200 text-[9px] text-cyan-700">
            <b>Invoice Amount</b> = SUMIF(Obligor, Amount) — anlık
          </div>
        </div>
      </div>
    </div>
  </div>;
}




// ═══════════════════════════════════════════════════════════════════════════
// SAISA RF — bank-based invoice discount (BBVA, Santander, Bankinter)
// Each invoice is assigned to a bank (not obligor). Tracks limit/risk by bank.
// ═══════════════════════════════════════════════════════════════════════════

const SAISA_DEFAULT_INVOICES=[{invoice:"SW20250100000808",discDate:"",amount:1200600.0,dueDate:"",bank:"Santander"},{invoice:"SW20250100002317",discDate:"16/09/2025",amount:380265.0,dueDate:"14/01/2026",bank:"Bankinter"},{invoice:"SW20250100002319",discDate:"16/09/2025",amount:392265.0,dueDate:"14/01/2026",bank:"Bankinter"},{invoice:"SW20250100002321",discDate:"16/09/2025",amount:1253580.0,dueDate:"14/01/2026",bank:"Santander"},{invoice:"SW20250100002316",discDate:"16/09/2025",amount:3670740.0,dueDate:"14/01/2026",bank:"BBVA"},{invoice:"SW20250100002461",discDate:"03/10/2025",amount:1462510.0,dueDate:"28/01/2026",bank:"Bankinter"},{invoice:"SW20250100002462",discDate:"03/10/2025",amount:626790.0,dueDate:"28/01/2026",bank:"Bankinter"},{invoice:"SW20250100002452",discDate:"07/10/2025",amount:149346.0,dueDate:"28/01/2026",bank:"Bankinter"},{invoice:"SW20250100002478",discDate:"07/10/2025",amount:735179.0,dueDate:"31/01/2026",bank:"Bankinter"},{invoice:"SW20250100002479",discDate:"07/10/2025",amount:152106.0,dueDate:"31/01/2026",bank:"Bankinter"},{invoice:"SW20250100002482",discDate:"07/10/2025",amount:952720.8,dueDate:"31/01/2026",bank:"Bankinter"},{invoice:"SW20250100002480",discDate:"08/10/2025",amount:1044650.0,dueDate:"02/02/2026",bank:"BBVA"},{invoice:"SW20250100002494",discDate:"10/10/2025",amount:181237.0,dueDate:"02/02/2026",bank:"Bankinter"},{invoice:"SW20250100002510",discDate:"10/10/2025",amount:102065.0,dueDate:"02/02/2026",bank:"Bankinter"},{invoice:"SW20250100002908",discDate:"22/01/2026",amount:129455.0,dueDate:"16/03/2026",bank:"BBVA"},{invoice:"SW20250100002968",discDate:"20/11/2025",amount:1926960.0,dueDate:"20/03/2026",bank:"Bankinter"},{invoice:"SW20250100002969",discDate:"20/11/2025",amount:1926960.0,dueDate:"20/03/2026",bank:"BBVA"},{invoice:"SW20250100003029",discDate:"27/11/2025",amount:256510.0,dueDate:"27/03/2026",bank:"Bankinter"},{invoice:"SW20250100003031",discDate:"27/11/2025",amount:1666560.0,dueDate:"27/03/2026",bank:"Bankinter"},{invoice:"SW20250100003021",discDate:"27/11/2025",amount:372765.0,dueDate:"27/03/2026",bank:"Bankinter"},{invoice:"SW20250100003022",discDate:"27/11/2025",amount:372765.0,dueDate:"27/03/2026",bank:"Bankinter"},{invoice:"SW20250100003032",discDate:"27/11/2025",amount:3115879.0,dueDate:"27/03/2026",bank:"BBVA"},{invoice:"SW20250100002959",discDate:"03/02/2026",amount:146346.0,dueDate:"27/03/2026",bank:"BBVA"},{invoice:"SW20250100003224",discDate:"18/12/2025",amount:2940000.0,dueDate:"17/04/2026",bank:"Bankinter"},{invoice:"SW20250100003225",discDate:"18/12/2025",amount:1528000.0,dueDate:"17/04/2026",bank:"BBVA"},{invoice:"SW20250100003226",discDate:"18/12/2025",amount:980000.0,dueDate:"17/04/2026",bank:"Santander"},{invoice:"SW20250100003258",discDate:"16/02/2026",amount:61056.0,dueDate:"19/04/2026",bank:"BBVA"},{invoice:"SW20250100003270",discDate:"16/02/2026",amount:50182.0,dueDate:"02/05/2026",bank:"BBVA"},{invoice:"SW20250100003374",discDate:"16/02/2026",amount:72273.0,dueDate:"02/05/2026",bank:"BBVA"},{invoice:"SW20250100003398",discDate:"16/02/2026",amount:73473.0,dueDate:"04/05/2026",bank:"BBVA"},{invoice:"SW20250100003385",discDate:"22/01/2026",amount:384765.0,dueDate:"06/05/2026",bank:"BBVA"},{invoice:"SW20260100000084",discDate:"03/02/2026",amount:146346.0,dueDate:"18/05/2026",bank:"BBVA"},{invoice:"SW20260100000071",discDate:"16/02/2026",amount:56160.0,dueDate:"18/05/2026",bank:"BBVA"},{invoice:"SW20260100000107",discDate:"16/02/2026",amount:51182.0,dueDate:"25/05/2026",bank:"Bankinter"},{invoice:"SW20260100000105",discDate:"16/02/2026",amount:50182.0,dueDate:"25/05/2026",bank:"BBVA"},{invoice:"SW20260100000151",discDate:"03/02/2026",amount:960000.0,dueDate:"26/05/2026",bank:"Santander"},{invoice:"SW20260100000152",discDate:"03/02/2026",amount:124255.0,dueDate:"26/05/2026",bank:"Santander"},{invoice:"SW20260100000148",discDate:"06/02/2026",amount:940800.0,dueDate:"26/05/2026",bank:"Bankinter"},{invoice:"SW20260100000150",discDate:"06/02/2026",amount:745530.0,dueDate:"26/05/2026",bank:"Bankinter"},{invoice:"SW20260100000138",discDate:"16/02/2026",amount:177737.0,dueDate:"29/05/2026",bank:"Bankinter"},{invoice:"SW20260100000210",discDate:"12/02/2026",amount:119455.0,dueDate:"30/05/2026",bank:"BBVA"},{invoice:"SW20260100000103",discDate:"16/02/2026",amount:72273.0,dueDate:"31/05/2026",bank:"Bankinter"},{invoice:"SW20260100000316",discDate:"19/02/2026",amount:1766400.0,dueDate:"13/06/2026",bank:"Santander"},{invoice:"SW20260100000370",discDate:"27/02/2206",amount:410416.0,dueDate:"23/06/2026",bank:"Bankinter"},{invoice:"SW20260100000374",discDate:"27/02/2206",amount:1344000.0,dueDate:"23/06/2026",bank:"Santander"},{invoice:"SW20260100000375",discDate:"27/02/2206",amount:3102000.0,dueDate:"23/06/2026",bank:"BBVA"},{invoice:"SW20260100000387",discDate:"27/02/2206",amount:497020.0,dueDate:"23/06/2026",bank:"BBVA"},{invoice:"SW20260100000477",discDate:"06/03/2026",amount:1118295.0,dueDate:"01/07/2026",bank:"Bankinter"},{invoice:"SW20260100000478",discDate:"06/03/2026",amount:652800.0,dueDate:"02/07/2026",bank:"Santander"},{invoice:"SW20260100000479",discDate:"06/03/2026",amount:672000.0,dueDate:"01/07/2026",bank:"Santander"},{invoice:"SW20260100000481",discDate:"06/03/2026",amount:1344000.0,dueDate:"01/07/2026",bank:"BBVA"},{invoice:"SW20260100000639",discDate:"25/03/2026",amount:1018891.0,dueDate:"22/07/2026",bank:"Bankinter"},{invoice:"SW20260100000640",discDate:"25/03/2026",amount:1344000.0,dueDate:"22/07/2026",bank:"BBVA"},{invoice:"SW20260100000209",discDate:"31/03/2026",amount:119455.0,dueDate:"02/06/2026",bank:"Santander"},{invoice:"SW20260100000244",discDate:"31/03/2026",amount:70773.0,dueDate:"10/06/2026",bank:"Santander"},{invoice:"SW20260100000393",discDate:"31/03/2026",amount:177737.0,dueDate:"26/06/2026",bank:"Santander"},{invoice:"SW20260100000372",discDate:"31/03/2026",amount:59964.75,dueDate:"23/06/2026",bank:"Santander"},{invoice:"SW20260100000493",discDate:"31/03/2026",amount:51782.0,dueDate:"05/07/2026",bank:"Santander"},{invoice:"SW20260100000498",discDate:"31/03/2026",amount:119455.0,dueDate:"05/07/2026",bank:"Santander"},{invoice:"SW20260100000514",discDate:"31/03/2026",amount:76773.0,dueDate:"10/07/2026",bank:"Santander"},{invoice:"SW20260100000569",discDate:"31/03/2026",amount:72273.0,dueDate:"23/07/2026",bank:"Santander"},{invoice:"SW20260100000570",discDate:"31/03/2026",amount:94364.0,dueDate:"17/07/2026",bank:"Santander"},{invoice:"SW20260100000631",discDate:"31/03/2026",amount:93600.0,dueDate:"19/07/2026",bank:"Santander"},{invoice:"SW20260100000731",discDate:"31/03/2026",amount:177737.0,dueDate:"31/07/2026",bank:"Santander"},{invoice:"SW20260100000662",discDate:"31/03/2026",amount:497020.0,dueDate:"30/07/2026",bank:"Bankinter"},{invoice:"SW20260100000663",discDate:"31/03/2026",amount:513020.0,dueDate:"30/07/2026",bank:"Bankinter"},{invoice:"SW20260100000666",discDate:"31/03/2026",amount:53302.0,dueDate:"30/07/2026",bank:"Santander"},{invoice:"SW20260100000668",discDate:"31/03/2026",amount:2256000.0,dueDate:"30/07/2026",bank:"Santander"},{invoice:"SW20260100000669",discDate:"31/03/2026",amount:1444800.0,dueDate:"30/07/2026",bank:"Bankinter"},{invoice:"SW20260100000670",discDate:"31/03/2026",amount:672000.0,dueDate:"30/07/2026",bank:"Bankinter"},{invoice:"SW20260100000733",discDate:"31/03/2026",amount:1516800.0,dueDate:"29/07/2026",bank:"BBVA"}];

const SAISA_DEFAULT_BANKS=[
  {name:"BBVA",limit:49300000,rate:"E+ %1,2 + 0,15%"},
  {name:"Santander",limit:22000000,rate:"E+ %1,2 + 0,15%"},
  {name:"Bankinter",limit:19000000,rate:"E+ %1,2 + 0,15%"}
];


type SaisaInv={id:string,invoice:string,discDate:string,amount:number,dueDate:string,bank:string};
type SaisaBank={id:string,name:string,limit:number,rate:string};

function SaisaSheet({sm,sy,canEdit}:{sm:number,sy:number,canEdit:boolean}){
  const[invoices,setInvoices]=useState<SaisaInv[]>([]);
  const[banks,setBanks]=useState<SaisaBank[]>([]);
  const[sortBy,setSortBy]=useState<string>("dueDate");
  const[filterBank,setFilterBank]=useState("");
  const[fInvoice,setFInvoice]=useState("");
  const[fDiscDate,setFDiscDate]=useState("");
  const[fAmount,setFAmount]=useState("");
  const[fDueDate,setFDueDate]=useState("");
  const[fBank,setFBank]=useState("");
  const[pasteMode,setPasteMode]=useState(false);
  const[pasteText,setPasteText]=useState("");
  const[showLimits,setShowLimits]=useState(true);
  const[isDraggingFile,setIsDraggingFile]=useState(false);
  const[dropStatus,setDropStatus]=useState<string>("");

  const invKey="saisa_inv_"+sy+"_"+sm;
  const bankKey="saisa_bank_"+sy+"_"+sm;

  useEffect(()=>{
    const reload=()=>{
      try{
        const si=localStorage.getItem(invKey);
        if(si){setInvoices(JSON.parse(si));}
        else{const seeded=SAISA_DEFAULT_INVOICES.map((d,i)=>({id:"saisa_inv_seed_"+i,...d}));setInvoices(seeded);__undoEnabled=false;try{localStorage.setItem(invKey,JSON.stringify(seeded));}catch{}__undoEnabled=true;}
        const sb=localStorage.getItem(bankKey);
        if(sb){setBanks(JSON.parse(sb));}
        else{const seeded=SAISA_DEFAULT_BANKS.map((d,i)=>({id:"saisa_bank_seed_"+i,...d}));setBanks(seeded);__undoEnabled=false;try{localStorage.setItem(bankKey,JSON.stringify(seeded));}catch{}__undoEnabled=true;}
      }catch{}
    };
    reload();
    const onStorage=(e:StorageEvent)=>{if(!e.key||e.key===invKey||e.key===bankKey)reload();};
    window.addEventListener("storage",onStorage);
    return()=>window.removeEventListener("storage",onStorage);
  },[sm,sy]);

  const saveInv=(v:SaisaInv[])=>{setInvoices(v);try{trackedSetItem(invKey,JSON.stringify(v));}catch{}};
  const saveBank=(v:SaisaBank[])=>{setBanks(v);try{trackedSetItem(bankKey,JSON.stringify(v));}catch{}};

  const addInv=()=>{saveInv([...invoices,{id:"saisa_inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:"",discDate:"",amount:0,dueDate:"",bank:banks[0]?.name||"BBVA"}]);};
  const updInv=(id:string,f:keyof SaisaInv,v:any)=>{saveInv(invoices.map(r=>r.id===id?{...r,[f]:v}:r));};
  const delInv=(id:string)=>{if(!confirm("Bu faturayı silmek istediğinize emin misiniz?"))return;saveInv(invoices.filter(r=>r.id!==id));};
  const addBank=()=>{saveBank([...banks,{id:"saisa_bank_"+Date.now(),name:"",limit:0,rate:""}]);};
  const updBank=(id:string,f:keyof SaisaBank,v:any)=>{saveBank(banks.map(r=>r.id===id?{...r,[f]:v}:r));};
  const delBank=(id:string)=>{if(!confirm("Banka sil?"))return;saveBank(banks.filter(r=>r.id!==id));};

  const parsePaste=()=>{
    const lines=pasteText.trim().split("\n").filter(l=>l.trim());
    const parsed:SaisaInv[]=lines.map(line=>{const cols=line.split("\t");const amt=parseFloat((cols[2]||"0").replace(/[,\s]/g,""))||0;return{id:"saisa_inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:(cols[1]||"").trim(),discDate:(cols[0]||"").trim(),amount:amt,dueDate:(cols[3]||"").trim(),bank:(cols[4]||"BBVA").trim()};});
    saveInv([...invoices,...parsed]);setPasteText("");setPasteMode(false);
  };

  const ensureSheetJS=async()=>{if((window as any).XLSX)return (window as any).XLSX;return new Promise<any>((resolve,reject)=>{const s=document.createElement("script");s.src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";s.onload=()=>resolve((window as any).XLSX);s.onerror=()=>reject(new Error("XLSX yüklenemedi"));document.head.appendChild(s);});};
  // SAISA Excel columns: B=Discount Date, C=Invoice, D=Amount, E=Due Date, F=Bank
  const rowToInv=(row:any[]):SaisaInv|null=>{
    if(!row||row.length<1)return null;
    const str=(v:any)=>v===null||v===undefined?"":String(v).trim();
    const dateStr=(v:any)=>{if(!v)return "";if(typeof v==="number"&&v>25569&&v<200000){const d=new Date((v-25569)*86400*1000);return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;}if(v instanceof Date)return `${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;return String(v).trim();};
    const inv=str(row[1]);if(!inv)return null;
    const amt=typeof row[2]==="number"?row[2]:parseFloat(str(row[2]).replace(/[,\s]/g,""))||0;
    return{id:"saisa_inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),invoice:inv,discDate:dateStr(row[0]),amount:amt,dueDate:dateStr(row[3]),bank:str(row[4])||"BBVA"};
  };
  const handleFile=async(file:File)=>{
    try{
      setDropStatus("İşleniyor: "+file.name);
      const name=file.name.toLowerCase();let rows:any[][]=[];
      if(name.endsWith(".csv")||name.endsWith(".tsv")||name.endsWith(".txt")){
        const txt=await file.text();const sep=name.endsWith(".tsv")||txt.includes("\t")?"\t":(txt.split("\n")[0]?.includes(";")?";":",");
        rows=txt.split(/\r?\n/).filter(l=>l.trim()).map(l=>l.split(sep));
        if(rows[0]&&/disc|date|invoice/i.test(String(rows[0][0]||"")+String(rows[0][1]||"")))rows=rows.slice(1);
      }else if(name.endsWith(".xlsx")||name.endsWith(".xls")){
        const XLSX=await ensureSheetJS();const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array",cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];
        rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}) as any[][];
        // SAISA original excel: data starts around row 13 after many header rows; cols B-F (indices 1-5)
        // Skip rows until we find first real invoice row
        rows=rows.map(r=>r.slice(1));// shift by 1 to start from col B
        rows=rows.filter(r=>r[1]&&String(r[1]).trim().length>0&&!/^invoice|^discount/i.test(String(r[1])));
      }else{setDropStatus("Desteklenmeyen format.");setTimeout(()=>setDropStatus(""),3500);return;}
      const parsed=rows.map(rowToInv).filter((x):x is SaisaInv=>!!x);
      if(parsed.length===0){setDropStatus("Fatura bulunamadı.");setTimeout(()=>setDropStatus(""),3500);return;}
      saveInv([...invoices,...parsed]);setDropStatus(`✓ ${parsed.length} fatura içeri aktarıldı`);setTimeout(()=>setDropStatus(""),3000);
    }catch(err){setDropStatus("Hata: "+(err instanceof Error?err.message:"dosya okunamadı"));setTimeout(()=>setDropStatus(""),4000);}
  };
  const onDrop=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(false);const files=Array.from(e.dataTransfer.files);if(files.length>0)handleFile(files[0]);};
  const onDragOver=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(true);};
  const onDragLeave=(e:React.DragEvent)=>{if(e.currentTarget===e.target)setIsDraggingFile(false);};

  // Per-bank total risk
  const bankRisk=(name:string)=>invoices.filter(r=>r.bank===name).reduce((s,r)=>s+r.amount,0);
  const totalRisk=invoices.reduce((s,r)=>s+r.amount,0);
  const totalLimit=banks.reduce((s,b)=>s+b.limit,0);
  const totalRemaining=totalLimit-totalRisk;

  const banksSet=[...new Set(invoices.map(r=>r.bank).filter(b=>b))].sort();
  let view=[...invoices];
  if(filterBank)view=view.filter(r=>r.bank===filterBank);
  const ci=(s:string)=>s.toLowerCase();
  if(fInvoice)view=view.filter(r=>ci(r.invoice).includes(ci(fInvoice)));
  if(fDiscDate)view=view.filter(r=>ci(r.discDate).includes(ci(fDiscDate)));
  if(fAmount)view=view.filter(r=>r.amount.toString().includes(fAmount.replace(/[,\s]/g,"")));
  if(fDueDate)view=view.filter(r=>ci(r.dueDate).includes(ci(fDueDate)));
  if(fBank)view=view.filter(r=>ci(r.bank).includes(ci(fBank)));
  view.sort((a,b)=>{
    if(sortBy==="dueDate"||sortBy==="discDate"){const ka=sortBy as "dueDate"|"discDate";return parseDDMMYYYY(a[ka])-parseDDMMYYYY(b[ka]);}
    if(sortBy==="amount")return a.amount-b.amount;
    if(sortBy==="invoice")return a.invoice.localeCompare(b.invoice);
    if(sortBy==="bank")return a.bank.localeCompare(b.bank);
    return 0;
  });

  const bankColor=(n:string)=>n==="BBVA"?"bg-blue-50 text-blue-700 border-blue-200":n==="Santander"?"bg-red-50 text-red-700 border-red-200":n==="Bankinter"?"bg-orange-50 text-orange-700 border-orange-200":"bg-gray-50 text-gray-600 border-gray-200";

  const resetDefaults=()=>{
    if(!confirm("Tüm faturalar silinip Excel default'larıyla doldurulacak. Emin misiniz?"))return;
    const si=SAISA_DEFAULT_INVOICES.map((d,i)=>({id:"saisa_inv_seed_"+i,...d}));
    const sb=SAISA_DEFAULT_BANKS.map((d,i)=>({id:"saisa_bank_seed_"+i,...d}));
    saveInv(si);saveBank(sb);
  };

  return<div className="relative" onDragOver={canEdit?onDragOver:undefined} onDragLeave={canEdit?onDragLeave:undefined} onDrop={canEdit?onDrop:undefined}>
    {isDraggingFile&&canEdit&&<div className="absolute inset-0 z-50 bg-orange-500/20 border-4 border-dashed border-orange-500 rounded-lg flex items-center justify-center pointer-events-none"><div className="bg-white px-6 py-4 rounded-xl shadow-xl border border-orange-200"><div className="text-2xl font-bold text-orange-700 mb-1">📥 Dosyayı buraya bırakın</div><div className="text-[11px] text-gray-500 text-center">.xlsx • .csv • .tsv • .txt</div></div></div>}
    {dropStatus&&<div className={"mb-3 px-3 py-2 rounded-lg text-xs font-semibold "+(dropStatus.startsWith("✓")?"bg-emerald-50 text-emerald-700 border border-emerald-200":dropStatus.startsWith("Hata")||dropStatus.startsWith("Desteklenmeyen")||dropStatus.startsWith("Fatura")?"bg-red-50 text-red-700 border border-red-200":"bg-blue-50 text-blue-700 border border-blue-200")}>{dropStatus}</div>}

    <div className="flex justify-between items-center mb-3 flex-wrap gap-2">
      <div><span className="text-sm font-bold text-orange-700">SAISA RF</span><span className="text-[10px] text-gray-400 ml-2">Bank-based invoice discount — {MONTHS[sm]} {sy} — {invoices.length} fatura</span><span className="text-[10px] text-orange-500 ml-2 hidden sm:inline">• Excel/CSV dosyasını sürükleyip bırakabilirsiniz</span></div>
      <div className="flex gap-2 flex-wrap">
        {canEdit&&<button onClick={()=>setPasteMode(!pasteMode)} className="px-3 py-1 bg-amber-100 hover:bg-amber-200 text-amber-700 rounded text-[11px] font-semibold">Excel\'den Yapıştır</button>}
        {canEdit&&<label className="px-3 py-1 bg-blue-100 hover:bg-blue-200 text-blue-700 rounded text-[11px] font-semibold cursor-pointer">📁 Dosya Seç<input type="file" accept=".xlsx,.xls,.csv,.tsv,.txt" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f);e.target.value="";}}/></label>}
        {canEdit&&<button onClick={resetDefaults} className="px-3 py-1 bg-gray-100 hover:bg-gray-200 text-gray-600 rounded text-[11px]">↻ Sıfırla</button>}
        <ExportBtn tableId="saisaInvTbl" name={"SAISA_RF_"+MONTHS[sm]+"_"+sy}/>
      </div>
    </div>

    {pasteMode&&<div className="mb-3 p-3 bg-white border border-amber-300 rounded-lg">
      <div className="text-[11px] text-gray-600 mb-2">Tab-separated — 5 sütun: Discount Date | Invoice Number | Amount | Due Date | Bank</div>
      <textarea value={pasteText} onChange={e=>setPasteText(e.target.value)} className="w-full h-32 px-3 py-2 border border-gray-300 rounded text-xs font-mono outline-none resize-none"/>
      <div className="flex gap-2 mt-2"><button onClick={parsePaste} className="px-3 py-1 bg-emerald-600 text-white rounded text-[11px] font-semibold">İçeri Aktar ({pasteText.trim().split("\n").filter(l=>l.trim()).length} satır)</button><button onClick={()=>{setPasteMode(false);setPasteText("");}} className="px-3 py-1 bg-gray-200 rounded text-[11px]">İptal</button></div>
    </div>}

    {/* Summary */}
    <div className="mb-4 p-3 bg-gradient-to-r from-orange-50 via-white to-orange-50 border border-orange-200 rounded-lg">
      <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
        <div><div className="text-[9px] text-gray-500 uppercase">Total Risk</div><div className="text-[13px] font-mono font-bold text-orange-700">{fmt(totalRisk)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Total Limit</div><div className="text-[13px] font-mono font-bold text-emerald-700">{fmt(totalLimit)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Remaining Limit</div><div className={"text-lg font-mono font-bold "+(totalRemaining<0?"text-red-600":"text-emerald-800")}>{totalRemaining<0?"("+fmt(Math.abs(totalRemaining))+")":fmt(totalRemaining)}</div></div>
      </div>
    </div>

    {invoices.length>0&&<div className="flex gap-2 mb-3 flex-wrap text-[10px] items-center">
      <div className="flex items-center gap-1"><span className="text-gray-500">Bank:</span><select value={filterBank} onChange={e=>setFilterBank(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="">Tümü ({banksSet.length})</option>{banksSet.map(b=><option key={b} value={b}>{b}</option>)}</select></div>
      <div className="flex items-center gap-1"><span className="text-gray-500">Sırala:</span><select value={sortBy} onChange={e=>setSortBy(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="dueDate">Due Date</option><option value="discDate">Discount Date</option><option value="bank">Bank</option><option value="amount">Amount</option><option value="invoice">Invoice No</option></select></div>
      {(filterBank||fInvoice||fDiscDate||fAmount||fDueDate||fBank)&&<button onClick={()=>{setFilterBank("");setFInvoice("");setFDiscDate("");setFAmount("");setFDueDate("");setFBank("");}} className="px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded text-[10px]">Tüm Filtreleri Temizle</button>}
      <span className="text-gray-400">{view.length} / {invoices.length} fatura</span>
      {canEdit&&<button onClick={addInv} className="ml-auto px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[11px] font-semibold shadow-sm">+ Fatura Ekle</button>}
    </div>}

    <div className="grid lg:grid-cols-[1fr_440px] gap-4">
      <div>
        <table id="saisaInvTbl" className="border-collapse text-[10px] w-full" style={{minWidth:720}}>
          <thead>
            <tr className="border-b-2 border-orange-300 bg-orange-50 sticky top-0">
              <th className="px-2 py-2 text-left text-[9px] font-semibold text-orange-800" style={{minWidth:30}}>#</th>
              <th className="px-2 py-2 text-center text-[9px] font-semibold text-orange-800" style={{minWidth:85}}>Discount<br/>Date</th>
              <th className="px-2 py-2 text-left text-[9px] font-semibold text-orange-800" style={{minWidth:160}}>Invoice Number</th>
              <th className="px-2 py-2 text-right text-[9px] font-semibold text-orange-800" style={{minWidth:110}}>Invoice<br/>Amount</th>
              <th className="px-2 py-2 text-center text-[9px] font-semibold text-orange-800" style={{minWidth:85}}>Due Date</th>
              <th className="px-2 py-2 text-center text-[9px] font-semibold text-orange-800" style={{minWidth:100}}>Bank</th>
              {canEdit&&<th className="w-6"></th>}
            </tr>
            <tr className="border-b border-orange-200 bg-orange-50/60">
              <th></th>
              <th className="px-1 py-1"><input value={fDiscDate} onChange={e=>setFDiscDate(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fInvoice} onChange={e=>setFInvoice(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] font-mono bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fAmount} onChange={e=>setFAmount(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] text-right font-mono bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fDueDate} onChange={e=>setFDueDate(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fBank} onChange={e=>setFBank(e.target.value)} placeholder="🔍 ara" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] text-center bg-white outline-none placeholder:text-gray-300"/></th>
              {canEdit&&<th></th>}
            </tr>
          </thead>
          <tbody>{view.length===0?<tr><td colSpan={canEdit?7:6} className="py-10 text-center text-gray-400">Henüz fatura yok.</td></tr>:view.map((r,i)=><tr key={r.id} className="border-b border-gray-100 hover:bg-orange-50/30">
            <td className="px-2 py-0.5 text-[9px] text-gray-400">{i+1}</td>
            <td className="px-1 py-0.5"><DateCell value={r.discDate} onChange={v=>updInv(r.id,"discDate",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><input className={inp+" font-mono text-[10px]"} value={r.invoice} onChange={e=>updInv(r.id,"invoice",e.target.value)}/></td>
            <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={addC(r.amount.toString())} onChange={e=>updInv(r.id,"amount",parseFloat(stripC(e.target.value))||0)}/></td>
            <td className="px-1 py-0.5"><DateCell value={r.dueDate} onChange={v=>updInv(r.id,"dueDate",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><select value={r.bank} onChange={e=>updInv(r.id,"bank",e.target.value)} className={"w-full px-1 py-0.5 border rounded text-[10px] text-center font-semibold outline-none "+bankColor(r.bank)}>{banks.map(b=><option key={b.id} value={b.name}>{b.name}</option>)}</select></td>
            {canEdit&&<td className="px-1 py-0.5"><button onClick={()=>delInv(r.id)} className="px-1 py-0.5 text-red-400 hover:bg-red-50 hover:text-red-600 rounded text-[11px]">×</button></td>}
          </tr>)}</tbody>
        </table>
      </div>

      {/* RIGHT: Bank Limits */}
      <div className="lg:sticky lg:top-2 lg:self-start">
        <div className="rounded-lg border border-orange-200 bg-white">
          <div className="flex justify-between items-center px-3 py-2 border-b border-orange-200 bg-orange-50 cursor-pointer" onClick={()=>setShowLimits(!showLimits)}>
            <div className="text-[11px] font-bold text-orange-800">{showLimits?"▼":"▶"} Bank Limits ({banks.length})</div>
            {canEdit&&showLimits&&<button onClick={(e)=>{e.stopPropagation();addBank();}} className="px-2 py-0.5 bg-orange-600 hover:bg-orange-700 text-white rounded text-[9px] font-semibold">+ Bank</button>}
          </div>
          {showLimits&&<div className="overflow-x-auto">
            <table className="border-collapse text-[9px] w-full" style={{minWidth:420}}>
              <thead><tr className="bg-orange-100/70 border-b border-orange-200">
                <th className="px-1 py-1 text-left text-[8px] font-bold text-orange-900" style={{minWidth:80}}>Bank</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-orange-900" title="Current Risk" style={{minWidth:75}}>Current Risk</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-orange-900" title="Remaining" style={{minWidth:75}}>Remaining</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-orange-900" style={{minWidth:80}}>Total Limit</th>
                <th className="px-1 py-1 text-left text-[8px] font-bold text-orange-900" style={{minWidth:90}}>Interest Rate</th>
                {canEdit&&<th className="w-4"></th>}
              </tr></thead>
              <tbody>{banks.map(b=>{const risk=bankRisk(b.name);const rem=b.limit-risk;return<tr key={b.id} className="border-b border-gray-100 hover:bg-orange-50/30">
                <td className="px-1 py-0.5"><input className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white outline-none font-semibold" value={b.name} onChange={e=>updBank(b.id,"name",e.target.value)}/></td>
                <td className={"px-1 py-0.5 text-right font-mono text-[9px] font-semibold "+(risk>0?"text-orange-700":"text-gray-300")}>{risk?fmt(risk):"-"}</td>
                <td className={"px-1 py-0.5 text-right font-mono text-[9px] font-semibold "+(rem<0?"text-red-600":"text-emerald-600")}>{rem<0?"("+fmt(Math.abs(rem))+")":fmt(rem)}</td>
                <td className="px-0.5 py-0.5"><input type="number" className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white text-right font-mono outline-none" value={b.limit||""} onChange={e=>updBank(b.id,"limit",parseFloat(e.target.value)||0)}/></td>
                <td className="px-0.5 py-0.5"><input className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white outline-none" value={b.rate} onChange={e=>updBank(b.id,"rate",e.target.value)} placeholder="e.g. E+1.2%"/></td>
                {canEdit&&<td className="px-0.5 py-0.5"><button onClick={()=>delBank(b.id)} className="px-0.5 py-0 text-red-400 hover:text-red-600 text-[10px]">×</button></td>}
              </tr>;})}
              <tr className="bg-orange-100 border-t-2 border-orange-300 font-bold"><td className="px-1 py-1 text-[9px] text-orange-900">Total Risk</td><td className="px-1 py-1 text-right font-mono text-[9px] text-orange-900">{fmt(totalRisk)}</td><td className={"px-1 py-1 text-right font-mono text-[9px] "+(totalRemaining<0?"text-red-700":"text-emerald-700")}>{totalRemaining<0?"("+fmt(Math.abs(totalRemaining))+")":fmt(totalRemaining)}</td><td className="px-1 py-1 text-right font-mono text-[9px] text-orange-900">{fmt(totalLimit)}</td><td></td>{canEdit&&<td></td>}</tr>
              </tbody>
            </table>
          </div>}
          <div className="px-3 py-2 bg-orange-50/40 border-t border-orange-200 text-[9px] text-orange-700 space-y-0.5">
            <div><b>Current Risk</b> = SUMIF(Bank, Amount) — anlık</div>
            <div><b>Remaining</b> = Total Limit − Current Risk</div>
          </div>
        </div>
      </div>
    </div>
  </div>;
}





// ═══════════════════════════════════════════════════════════════════════════
// SHARED DATA (Cheques / Factoring / Electric / Other Disc.)
// ═══════════════════════════════════════════════════════════════════════════

const CHEQUES_DEFAULT=[{bank:"DENİZBANK",branch:"ZİNCİRLİKUYU TİC.MRKZ.ŞB",number:"1616838",maturity:"21/04/2026",valor:"",amount:250000.0,endoser:"DÜZCE CAM",company:"Kazan Soda"},{bank:"DENİZBANK",branch:"ZİNCİRLİKUYU TİC.MRKZ.ŞB",number:"1616839",maturity:"23/04/2026",valor:"24/04/2026",amount:250000.0,endoser:"DÜZCE CAM",company:"Kazan Soda"},{bank:"DENİZBANK",branch:"ZİNCİRLİKUYU TİC.MRKZ.ŞB",number:"1616840",maturity:"29/04/2026",valor:"",amount:370104.0,endoser:"DÜZCE CAM",company:"Kazan Soda"},{bank:"DENİZBANK",branch:"ZİNCİRLİKUYU TİC.MRKZ.ŞB",number:"1616846",maturity:"07/05/2026",valor:"",amount:250000.0,endoser:"DÜZCE CAM",company:"Kazan Soda"},{bank:"DENİZBANK",branch:"ZİNCİRLİKUYU TİC.MRKZ.ŞB",number:"1616847",maturity:"12/05/2026",valor:"",amount:250000.0,endoser:"DÜZCE CAM",company:"Kazan Soda"},{bank:"DENİZBANK",branch:"ZİNCİRLİKUYU TİC.MRKZ.ŞB",number:"1616848",maturity:"14/05/2026",valor:"",amount:250000.0,endoser:"DÜZCE CAM",company:"Kazan Soda"},{bank:"DENİZBANK",branch:"ZİNCİRLİKUYU TİC.MRKZ.ŞB",number:"1616849",maturity:"19/05/2026",valor:"",amount:250000.0,endoser:"DÜZCE CAM",company:"Kazan Soda"},{bank:"DENİZBANK",branch:"ZİNCİRLİKUYU TİC.MRKZ.ŞB",number:"1616850",maturity:"21/05/2026",valor:"",amount:250000.0,endoser:"DÜZCE CAM",company:"Kazan Soda"},{bank:"DENİZBANK",branch:"ZİNCİRLİKUYU TİC.MRKZ.ŞB",number:"1616851",maturity:"26/05/2026",valor:"01/06/2026",amount:250000.0,endoser:"DÜZCE CAM",company:"Kazan Soda"},{bank:"DENİZBANK",branch:"ZİNCİRLİKUYU TİC.MRKZ.ŞB",number:"1616852",maturity:"28/05/2026",valor:"01/06/2026",amount:175159.2,endoser:"DÜZCE CAM",company:"Kazan Soda"}];

const OTHERDISC_DEFAULT=[{discDate:"24/10/2024",customer:"Tam Trade",invoice:"SW20240100002209",amount:215000.0,dueDate:"01/12/2024",netAmount:213147.7602739726,beneBank:"",discCost:0},{discDate:"24/10/2024",customer:"Tam Trade",invoice:"SW20240100002523",amount:120000.0,dueDate:"05/01/2025",netAmount:118014.0,beneBank:"",discCost:0},{discDate:"24/10/2024",customer:"Tam Trade",invoice:"SW20240100002416",amount:0,dueDate:"29/12/2024",netAmount:151695.69589041095,beneBank:"",discCost:0},{discDate:"24/10/2024",customer:"Tam Trade",invoice:"SW20240100002425",amount:0,dueDate:"29/12/2024",netAmount:433416.27397260274,beneBank:"",discCost:0},{discDate:"24/10/2024",customer:"Tam Trade",invoice:"SW20240100002441",amount:0,dueDate:"29/12/2024",netAmount:125592.21575342465,beneBank:"",discCost:0},{discDate:"24/10/2024",customer:"Tam Trade",invoice:"SW20240100002498",amount:0,dueDate:"29/12/2024",netAmount:125592.21575342465,beneBank:"",discCost:0},{discDate:"24/10/2024",customer:"Tam Trade",invoice:"SW20240100002500",amount:0,dueDate:"29/12/2024",netAmount:433416.27397260274,beneBank:"",discCost:0},{discDate:"24/10/2024",customer:"Tam Trade",invoice:"SW20240100002499",amount:132000.0,dueDate:"05/01/2025",netAmount:129815.4,beneBank:"",discCost:0}];

const OTHERDISC_CUSTOMERS=[{name:"Tam Trade",limit:0,rate:0.08275},{name:"UG Impex",limit:0,rate:0}];

const ELECTRIC_DEFAULT=[{date:"01/04/2026",total:1640054.68},{date:"02/04/2026",total:1463489.78},{date:"03/04/2026",total:862775.1},{date:"06/04/2026",total:772546.1799999999},{date:"07/04/2026",total:2479808.75},{date:"08/04/2026",total:32567022.669999998},{date:"09/04/2026",total:458188.24},{date:"10/04/2026",total:289074.04},{date:"13/04/2026",total:668609.37},{date:"14/04/2026",total:1452335.0399999998},{date:"15/04/2026",total:265447.13},{date:"16/04/2026",total:335214.29},{date:"17/04/2026",total:461024.49},{date:"20/04/2026",total:850000.0},{date:"21/04/2026",total:3000000.0},{date:"22/04/2026",total:850000.0},{date:"23/04/2026",total:0},{date:"24/04/2026",total:1667092.3},{date:"27/04/2026",total:850000.0},{date:"28/04/2026",total:3000000.0},{date:"29/04/2026",total:860000.0},{date:"30/04/2026",total:850000.0}];


// ═══════════════════════════════════════════════════════════════════════════
// CHEQUES (@safebox) — Kazan Soda + Eti Soda sections with TL/USD totals
// ═══════════════════════════════════════════════════════════════════════════

type Cheque={id:string,bank:string,branch:string,number:string,maturity:string,valor:string,amount:number,ccy:string,endoser:string,company:string};

function ChequesSheet({sm,sy,canEdit,fx}:{sm:number,sy:number,canEdit:boolean,fx:typeof FX_D}){
  const[rows,setRows]=useState<Cheque[]>([]);
  const[isDraggingFile,setIsDraggingFile]=useState(false);
  const[dropStatus,setDropStatus]=useState<string>("");

  const dataKey="cheques_"+sy+"_"+sm;
  useEffect(()=>{
    const reload=()=>{
      try{
        const s=localStorage.getItem(dataKey);
        if(s){setRows(JSON.parse(s));}
        else{const seeded=CHEQUES_DEFAULT.map((d,i)=>({id:"chq_seed_"+i,ccy:"USD",...d}));setRows(seeded);__undoEnabled=false;try{localStorage.setItem(dataKey,JSON.stringify(seeded));}catch{}__undoEnabled=true;}
      }catch{}
    };
    reload();
    const onS=(e:StorageEvent)=>{if(!e.key||e.key===dataKey)reload();};
    window.addEventListener("storage",onS);return()=>window.removeEventListener("storage",onS);
  },[sm,sy]);

  const save=(v:Cheque[])=>{setRows(v);try{trackedSetItem(dataKey,JSON.stringify(v));}catch{}};
  const addRow=(company:string)=>{save([...rows,{id:"chq_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),bank:"",branch:"",number:"",maturity:"",valor:"",amount:0,ccy:"USD",endoser:"",company}]);};
  const upd=(id:string,f:keyof Cheque,v:any)=>{save(rows.map(r=>r.id===id?{...r,[f]:v}:r));};
  const del=(id:string)=>{if(!confirm("Sil?"))return;save(rows.filter(r=>r.id!==id));};

  // Drag-drop
  const ensureSheetJS=async()=>{if((window as any).XLSX)return (window as any).XLSX;return new Promise<any>((resolve,reject)=>{const s=document.createElement("script");s.src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";s.onload=()=>resolve((window as any).XLSX);s.onerror=()=>reject(new Error("XLSX"));document.head.appendChild(s);});};
  const rowToChq=(row:any[],company:string):Cheque|null=>{
    if(!row||row.length<1)return null;
    const str=(v:any)=>v===null||v===undefined?"":String(v).trim();
    const dateStr=(v:any)=>{if(!v)return "";if(typeof v==="number"&&v>25569&&v<200000){const d=new Date((v-25569)*86400*1000);return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;}if(v instanceof Date)return `${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;return String(v).trim();};
    const bank=str(row[0]);if(!bank)return null;
    const amt=typeof row[5]==="number"?row[5]:parseFloat(str(row[5]).replace(/[,\s]/g,""))||0;
    return{id:"chq_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),bank,branch:str(row[1]),number:str(row[2]),maturity:dateStr(row[3]),valor:dateStr(row[4]),amount:amt,ccy:str(row[6]).toUpperCase()||"USD",endoser:str(row[7]),company};
  };
  const handleFile=async(file:File,company:string)=>{
    try{
      setDropStatus("İşleniyor: "+file.name);
      const name=file.name.toLowerCase();let raw:any[][]=[];
      if(name.endsWith(".csv")||name.endsWith(".tsv")||name.endsWith(".txt")){
        const txt=await file.text();const sep=name.endsWith(".tsv")||txt.includes("\t")?"\t":(txt.split("\n")[0]?.includes(";")?";":",");
        raw=txt.split(/\r?\n/).filter(l=>l.trim()).map(l=>l.split(sep));
        if(raw[0]&&/bank|branch/i.test(String(raw[0][0])))raw=raw.slice(1);
      }else if(name.endsWith(".xlsx")||name.endsWith(".xls")){
        const XLSX=await ensureSheetJS();const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array",cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];
        raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}) as any[][];
        raw=raw.map(r=>r.slice(1));// skip col A
        raw=raw.filter(r=>r[0]&&!/bank|kazan|eti|total|toplam/i.test(String(r[0])));
      }else{setDropStatus("Desteklenmeyen format.");setTimeout(()=>setDropStatus(""),3500);return;}
      const parsed=raw.map(r=>rowToChq(r,company)).filter((x):x is Cheque=>!!x);
      if(parsed.length===0){setDropStatus("Çek bulunamadı.");setTimeout(()=>setDropStatus(""),3500);return;}
      save([...rows,...parsed]);setDropStatus(`✓ ${parsed.length} çek içeri aktarıldı (${company})`);setTimeout(()=>setDropStatus(""),3000);
    }catch(err){setDropStatus("Hata: "+(err instanceof Error?err.message:""));setTimeout(()=>setDropStatus(""),4000);}
  };

  const resetDefaults=()=>{if(!confirm("Tüm çekler silinip Excel default'larıyla doldurulacak. Emin misiniz?"))return;save(CHEQUES_DEFAULT.map((d,i)=>({id:"chq_seed_"+i,ccy:"USD",...d})));};

  const renderSection=(company:string,title:string,colorHex:string)=>{
    const sectionRows=rows.filter(r=>r.company===company);
    const totalTL=sectionRows.filter(r=>r.ccy==="TRY").reduce((s,r)=>s+r.amount,0);
    const totalUSD=sectionRows.filter(r=>r.ccy==="USD").reduce((s,r)=>s+r.amount,0);
    const totalEUR=sectionRows.filter(r=>r.ccy==="EUR").reduce((s,r)=>s+r.amount,0);
    const grandTotalTL=totalTL+totalUSD*fx.usd_try+totalEUR*fx.eur_try;
    return<div className="mb-6 rounded-lg border-2 relative" style={{borderColor:colorHex+"55"}}
      onDragOver={canEdit?(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(true);}:undefined}
      onDragLeave={canEdit?(e:React.DragEvent)=>{if(e.currentTarget===e.target)setIsDraggingFile(false);}:undefined}
      onDrop={canEdit?(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(false);const files=Array.from(e.dataTransfer.files);if(files.length>0)handleFile(files[0],company);}:undefined}>
      <div className="flex justify-between items-center px-3 py-2 border-b" style={{backgroundColor:colorHex+"15",borderColor:colorHex+"55"}}>
        <div className="text-xs font-bold" style={{color:colorHex}}>{title} ({sectionRows.length} çek)</div>
        <div className="flex gap-2">
          {canEdit&&<label className="px-2 py-0.5 bg-blue-100 hover:bg-blue-200 text-blue-700 rounded text-[10px] font-semibold cursor-pointer">📁 Dosya Seç<input type="file" accept=".xlsx,.xls,.csv,.tsv,.txt" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f,company);e.target.value="";}}/></label>}
          {canEdit&&<button onClick={()=>addRow(company)} className="px-2 py-0.5 rounded text-[10px] font-semibold text-white" style={{backgroundColor:colorHex}}>+ Çek Ekle</button>}
        </div>
      </div>
      <div className="overflow-x-auto">
        <table className="border-collapse text-[10px] w-full" style={{minWidth:880}}>
          <thead><tr className="border-b text-[9px] font-semibold" style={{backgroundColor:colorHex+"10",color:colorHex,borderColor:colorHex+"33"}}>
            <th className="px-2 py-1 text-left" style={{minWidth:25}}>#</th>
            <th className="px-2 py-1 text-left" style={{minWidth:110}}>Bank</th>
            <th className="px-2 py-1 text-left" style={{minWidth:140}}>Branch</th>
            <th className="px-2 py-1 text-left" style={{minWidth:90}}>Number</th>
            <th className="px-2 py-1 text-center" style={{minWidth:85}}>Maturity</th>
            <th className="px-2 py-1 text-center" style={{minWidth:85}}>Valör-Vade</th>
            <th className="px-2 py-1 text-right" style={{minWidth:100}}>Amount</th>
            <th className="px-2 py-1 text-center" style={{minWidth:45}}>CCY</th>
            <th className="px-2 py-1 text-left" style={{minWidth:110}}>Endoser</th>
            {canEdit&&<th className="w-6"></th>}
          </tr></thead>
          <tbody>{sectionRows.length===0?<tr><td colSpan={canEdit?10:9} className="py-6 text-center text-gray-400 text-[10px]">Çek yok. "+ Çek Ekle" veya Excel sürükle-bırak</td></tr>:sectionRows.map((r,i)=><tr key={r.id} className="border-b border-gray-100 hover:bg-gray-50">
            <td className="px-2 py-0.5 text-[9px] text-gray-400">{i+1}</td>
            <td className="px-1 py-0.5"><input className={inp+" text-[10px]"} value={r.bank} onChange={e=>upd(r.id,"bank",e.target.value)}/></td>
            <td className="px-1 py-0.5"><input className={inp+" text-[10px]"} value={r.branch} onChange={e=>upd(r.id,"branch",e.target.value)}/></td>
            <td className="px-1 py-0.5"><input className={inp+" font-mono text-[10px]"} value={r.number} onChange={e=>upd(r.id,"number",e.target.value)}/></td>
            <td className="px-1 py-0.5"><DateCell value={r.maturity} onChange={v=>upd(r.id,"maturity",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><DateCell value={r.valor} onChange={v=>upd(r.id,"valor",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={addC(r.amount.toString())} onChange={e=>upd(r.id,"amount",parseFloat(stripC(e.target.value))||0)}/></td>
            <td className="px-1 py-0.5"><select value={r.ccy} onChange={e=>upd(r.id,"ccy",e.target.value)} className="w-full px-1 py-0.5 border border-gray-200 rounded text-[10px] bg-white text-center font-semibold outline-none"><option value="USD">USD</option><option value="TRY">TRY</option><option value="EUR">EUR</option><option value="GBP">GBP</option></select></td>
            <td className="px-1 py-0.5"><input className={inp+" text-[10px]"} value={r.endoser} onChange={e=>upd(r.id,"endoser",e.target.value)}/></td>
            {canEdit&&<td className="px-1 py-0.5"><button onClick={()=>del(r.id)} className="px-1 py-0.5 text-red-400 hover:bg-red-50 hover:text-red-600 rounded text-[11px]">×</button></td>}
          </tr>)}</tbody>
        </table>
      </div>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 px-3 py-2 border-t" style={{backgroundColor:colorHex+"08",borderColor:colorHex+"33"}}>
        <div><div className="text-[9px] text-gray-500 uppercase">Toplam TL</div><div className="text-[11px] font-mono font-bold">{fmt(totalTL)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Toplam USD</div><div className="text-[11px] font-mono font-bold text-green-600">{fmt(totalUSD)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Toplam EUR</div><div className="text-[11px] font-mono font-bold text-blue-600">{fmt(totalEUR)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Genel Toplam (TL)</div><div className="text-sm font-mono font-bold" style={{color:colorHex}}>{fmt(grandTotalTL)}</div></div>
      </div>
    </div>;
  };

  const grandTL=rows.filter(r=>r.ccy==="TRY").reduce((s,r)=>s+r.amount,0);
  const grandUSD=rows.filter(r=>r.ccy==="USD").reduce((s,r)=>s+r.amount,0);
  const grandEUR=rows.filter(r=>r.ccy==="EUR").reduce((s,r)=>s+r.amount,0);
  const grandAllTL=grandTL+grandUSD*fx.usd_try+grandEUR*fx.eur_try;

  return<div className="relative">
    {isDraggingFile&&canEdit&&<div className="fixed inset-0 z-50 bg-purple-500/10 pointer-events-none"/>}
    {dropStatus&&<div className={"mb-3 px-3 py-2 rounded-lg text-xs font-semibold "+(dropStatus.startsWith("✓")?"bg-emerald-50 text-emerald-700 border border-emerald-200":dropStatus.startsWith("Hata")||dropStatus.startsWith("Desteklenmeyen")||dropStatus.startsWith("Çek")?"bg-red-50 text-red-700 border border-red-200":"bg-blue-50 text-blue-700 border border-blue-200")}>{dropStatus}</div>}

    <div className="flex justify-between items-center mb-3 flex-wrap gap-2">
      <div><span className="text-sm font-bold text-purple-700">Cheques @ Safebox</span><span className="text-[10px] text-gray-400 ml-2">Kasadaki çekler — {MONTHS[sm]} {sy} — {rows.length} çek</span><span className="text-[10px] text-purple-500 ml-2 hidden sm:inline">• Her bölüme ayrı Excel sürükleyebilirsiniz</span></div>
      <div className="flex gap-2 flex-wrap">
        {canEdit&&<button onClick={resetDefaults} className="px-3 py-1 bg-gray-100 hover:bg-gray-200 text-gray-600 rounded text-[11px]">↻ Sıfırla</button>}
        <ExportBtn tableId="chequesTbl" name={"Cheques_"+MONTHS[sm]+"_"+sy}/>
      </div>
    </div>

    {/* Overall summary */}
    <div className="mb-4 p-3 bg-gradient-to-r from-purple-50 via-white to-purple-50 border border-purple-200 rounded-lg">
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <div><div className="text-[9px] text-gray-500 uppercase">TOPLAM TL (Kasa)</div><div className="text-[11px] font-mono font-bold">{fmt(grandTL)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">TOPLAM USD</div><div className="text-[11px] font-mono font-bold text-green-600">{fmt(grandUSD)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">TOPLAM EUR</div><div className="text-[11px] font-mono font-bold text-blue-600">{fmt(grandEUR)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">GENEL TOPLAM (TL)</div><div className="text-lg font-mono font-bold text-purple-800">{fmt(grandAllTL)}</div></div>
      </div>
    </div>

    <div id="chequesTbl">
      {renderSection("Kazan Soda","KAZAN SODA CHEQUES","#7c3aed")}
      {renderSection("Eti Soda","ETİ SODA CHEQUES","#059669")}
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════
// FACTORING (shared for Garanti + YK) — Kazan + Eti sections, TL-only totals
// ═══════════════════════════════════════════════════════════════════════════

type Factoring={id:string,bank:string,branch:string,number:string,maturity:string,amount:number,endoser:string,company:string};

function FactoringSheet({sm,sy,canEdit,kind,color,label}:{sm:number,sy:number,canEdit:boolean,kind:string,color:string,label:string}){
  const[rows,setRows]=useState<Factoring[]>([]);
  const[isDraggingFile,setIsDraggingFile]=useState(false);
  const[dropStatus,setDropStatus]=useState<string>("");

  const dataKey=kind+"_fact_"+sy+"_"+sm;
  useEffect(()=>{
    const reload=()=>{try{const s=localStorage.getItem(dataKey);if(s)setRows(JSON.parse(s));else setRows([]);}catch{}};
    reload();
    const onS=(e:StorageEvent)=>{if(!e.key||e.key===dataKey)reload();};
    window.addEventListener("storage",onS);return()=>window.removeEventListener("storage",onS);
  },[sm,sy,kind]);

  const save=(v:Factoring[])=>{setRows(v);try{trackedSetItem(dataKey,JSON.stringify(v));}catch{}};
  const addRow=(company:string)=>{save([...rows,{id:kind+"_fact_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),bank:"",branch:"",number:"",maturity:"",amount:0,endoser:"",company}]);};
  const upd=(id:string,f:keyof Factoring,v:any)=>{save(rows.map(r=>r.id===id?{...r,[f]:v}:r));};
  const del=(id:string)=>{if(!confirm("Sil?"))return;save(rows.filter(r=>r.id!==id));};

  const ensureSheetJS=async()=>{if((window as any).XLSX)return (window as any).XLSX;return new Promise<any>((resolve,reject)=>{const s=document.createElement("script");s.src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";s.onload=()=>resolve((window as any).XLSX);s.onerror=()=>reject(new Error("XLSX"));document.head.appendChild(s);});};
  const rowToFact=(row:any[],company:string):Factoring|null=>{
    if(!row||row.length<1)return null;
    const str=(v:any)=>v===null||v===undefined?"":String(v).trim();
    const dateStr=(v:any)=>{if(!v)return "";if(typeof v==="number"&&v>25569&&v<200000){const d=new Date((v-25569)*86400*1000);return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;}if(v instanceof Date)return `${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;return String(v).trim();};
    const bank=str(row[0]);if(!bank)return null;
    const amt=typeof row[4]==="number"?row[4]:parseFloat(str(row[4]).replace(/[,\s]/g,""))||0;
    return{id:kind+"_fact_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),bank,branch:str(row[1]),number:str(row[2]),maturity:dateStr(row[3]),amount:amt,endoser:str(row[5]),company};
  };
  const handleFile=async(file:File,company:string)=>{
    try{
      setDropStatus("İşleniyor: "+file.name);
      const name=file.name.toLowerCase();let raw:any[][]=[];
      if(name.endsWith(".csv")||name.endsWith(".tsv")||name.endsWith(".txt")){
        const txt=await file.text();const sep=name.endsWith(".tsv")||txt.includes("\t")?"\t":(txt.split("\n")[0]?.includes(";")?";":",");
        raw=txt.split(/\r?\n/).filter(l=>l.trim()).map(l=>l.split(sep));
        if(raw[0]&&/bank|branch/i.test(String(raw[0][0])))raw=raw.slice(1);
      }else if(name.endsWith(".xlsx")||name.endsWith(".xls")){
        const XLSX=await ensureSheetJS();const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array",cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];
        raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}) as any[][];
        raw=raw.map(r=>r.slice(1));
        raw=raw.filter(r=>r[0]&&!/bank|kazan|eti|total|toplam|factoring/i.test(String(r[0])));
      }else{setDropStatus("Desteklenmeyen format.");setTimeout(()=>setDropStatus(""),3500);return;}
      const parsed=raw.map(r=>rowToFact(r,company)).filter((x):x is Factoring=>!!x);
      if(parsed.length===0){setDropStatus("Kayıt bulunamadı.");setTimeout(()=>setDropStatus(""),3500);return;}
      save([...rows,...parsed]);setDropStatus(`✓ ${parsed.length} satır içeri aktarıldı (${company})`);setTimeout(()=>setDropStatus(""),3000);
    }catch(err){setDropStatus("Hata: "+(err instanceof Error?err.message:""));setTimeout(()=>setDropStatus(""),4000);}
  };

  const renderSection=(company:string,title:string,colorHex:string)=>{
    const sectionRows=rows.filter(r=>r.company===company);
    const total=sectionRows.reduce((s,r)=>s+r.amount,0);
    return<div className="mb-6 rounded-lg border-2" style={{borderColor:colorHex+"55"}}
      onDragOver={canEdit?(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(true);}:undefined}
      onDragLeave={canEdit?(e:React.DragEvent)=>{if(e.currentTarget===e.target)setIsDraggingFile(false);}:undefined}
      onDrop={canEdit?(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(false);const files=Array.from(e.dataTransfer.files);if(files.length>0)handleFile(files[0],company);}:undefined}>
      <div className="flex justify-between items-center px-3 py-2 border-b" style={{backgroundColor:colorHex+"15",borderColor:colorHex+"55"}}>
        <div className="text-xs font-bold" style={{color:colorHex}}>{title} ({sectionRows.length} kayıt)</div>
        <div className="flex gap-2">
          {canEdit&&<label className="px-2 py-0.5 bg-blue-100 hover:bg-blue-200 text-blue-700 rounded text-[10px] font-semibold cursor-pointer">📁 Dosya Seç<input type="file" accept=".xlsx,.xls,.csv,.tsv,.txt" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f,company);e.target.value="";}}/></label>}
          {canEdit&&<button onClick={()=>addRow(company)} className="px-2 py-0.5 rounded text-[10px] font-semibold text-white" style={{backgroundColor:colorHex}}>+ Satır Ekle</button>}
        </div>
      </div>
      <div className="overflow-x-auto">
        <table className="border-collapse text-[10px] w-full" style={{minWidth:780}}>
          <thead><tr className="border-b text-[9px] font-semibold" style={{backgroundColor:colorHex+"10",color:colorHex,borderColor:colorHex+"33"}}>
            <th className="px-2 py-1 text-left" style={{minWidth:25}}>#</th>
            <th className="px-2 py-1 text-left" style={{minWidth:110}}>Bank</th>
            <th className="px-2 py-1 text-left" style={{minWidth:140}}>Branch</th>
            <th className="px-2 py-1 text-left" style={{minWidth:90}}>Number</th>
            <th className="px-2 py-1 text-center" style={{minWidth:85}}>Maturity</th>
            <th className="px-2 py-1 text-right" style={{minWidth:110}}>Amount (TL)</th>
            <th className="px-2 py-1 text-left" style={{minWidth:110}}>Endoser</th>
            {canEdit&&<th className="w-6"></th>}
          </tr></thead>
          <tbody>{sectionRows.length===0?<tr><td colSpan={canEdit?8:7} className="py-6 text-center text-gray-400 text-[10px]">Kayıt yok. "+ Satır Ekle" veya Excel sürükle-bırak</td></tr>:sectionRows.map((r,i)=><tr key={r.id} className="border-b border-gray-100 hover:bg-gray-50">
            <td className="px-2 py-0.5 text-[9px] text-gray-400">{i+1}</td>
            <td className="px-1 py-0.5"><input className={inp+" text-[10px]"} value={r.bank} onChange={e=>upd(r.id,"bank",e.target.value)}/></td>
            <td className="px-1 py-0.5"><input className={inp+" text-[10px]"} value={r.branch} onChange={e=>upd(r.id,"branch",e.target.value)}/></td>
            <td className="px-1 py-0.5"><input className={inp+" font-mono text-[10px]"} value={r.number} onChange={e=>upd(r.id,"number",e.target.value)}/></td>
            <td className="px-1 py-0.5"><DateCell value={r.maturity} onChange={v=>upd(r.id,"maturity",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={addC(r.amount.toString())} onChange={e=>upd(r.id,"amount",parseFloat(stripC(e.target.value))||0)}/></td>
            <td className="px-1 py-0.5"><input className={inp+" text-[10px]"} value={r.endoser} onChange={e=>upd(r.id,"endoser",e.target.value)}/></td>
            {canEdit&&<td className="px-1 py-0.5"><button onClick={()=>del(r.id)} className="px-1 py-0.5 text-red-400 hover:bg-red-50 hover:text-red-600 rounded text-[11px]">×</button></td>}
          </tr>)}</tbody>
          <tfoot><tr className="font-bold border-t-2" style={{backgroundColor:colorHex+"15",borderColor:colorHex+"55"}}>
            <td colSpan={5} className="px-2 py-1 text-right text-[10px]" style={{color:colorHex}}>TOTAL</td>
            <td className="px-2 py-1 text-right font-mono text-[11px]" style={{color:colorHex}}>{fmt(total)}</td>
            <td colSpan={canEdit?2:1}></td>
          </tr></tfoot>
        </table>
      </div>
    </div>;
  };

  const grandTotal=rows.reduce((s,r)=>s+r.amount,0);

  return<div className="relative">
    {dropStatus&&<div className={"mb-3 px-3 py-2 rounded-lg text-xs font-semibold "+(dropStatus.startsWith("✓")?"bg-emerald-50 text-emerald-700 border border-emerald-200":dropStatus.startsWith("Hata")||dropStatus.startsWith("Desteklenmeyen")||dropStatus.startsWith("Kayıt")?"bg-red-50 text-red-700 border border-red-200":"bg-blue-50 text-blue-700 border border-blue-200")}>{dropStatus}</div>}

    <div className="flex justify-between items-center mb-3 flex-wrap gap-2">
      <div><span className="text-sm font-bold" style={{color}}>{label}</span><span className="text-[10px] text-gray-400 ml-2">Factoring utilisations — {MONTHS[sm]} {sy} — {rows.length} kayıt</span><span className="text-[10px] ml-2 hidden sm:inline" style={{color}}>• Her bölüme ayrı Excel sürükleyebilirsiniz</span></div>
      <div className="flex gap-2 flex-wrap">
        <ExportBtn tableId={kind+"FactTbl"} name={label.replace(/ /g,"_")+"_"+MONTHS[sm]+"_"+sy}/>
      </div>
    </div>

    <div className="mb-4 p-3 rounded-lg border" style={{backgroundColor:color+"08",borderColor:color+"44"}}>
      <div className="flex items-center justify-between">
        <div className="text-[10px] text-gray-500 uppercase">TOPLAM (TL) — Kazan + Eti</div>
        <div className="text-lg font-mono font-bold" style={{color}}>{fmt(grandTotal)}</div>
      </div>
    </div>

    <div id={kind+"FactTbl"}>
      {renderSection("Kazan Soda","KAZAN SODA FACTORING UTILISATIONS",color)}
      {renderSection("Eti Soda","ETİ SODA FACTORING UTILISATIONS","#059669")}
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════
// ELECTRIC — Daily electricity receipts (Grup Dışı/İçi with KDV breakdown)
// ═══════════════════════════════════════════════════════════════════════════

type ElectricRow={id:string,date:string,total:number,epiasQty:number,epiasAmt:number,parkcam:number,sisecamQty:number,sisecamAmt:number,etiQty:number,etiAmt:number,kdvEpias:number,kdvSisecam:number,kdvEti:number,pfk:number,cengizDsg:number,note:string};

function ElectricSheet({sm,sy,canEdit}:{sm:number,sy:number,canEdit:boolean}){
  const[rows,setRows]=useState<ElectricRow[]>([]);
  const[isDraggingFile,setIsDraggingFile]=useState(false);
  const[dropStatus,setDropStatus]=useState<string>("");

  const dataKey="electric_"+sy+"_"+sm;
  useEffect(()=>{
    const reload=()=>{
      try{
        const s=localStorage.getItem(dataKey);
        if(s){setRows(JSON.parse(s));}
        else{
          const seeded=ELECTRIC_DEFAULT.map((d,i)=>({id:"elec_seed_"+i,date:d.date,total:d.total,epiasQty:0,epiasAmt:0,parkcam:0,sisecamQty:0,sisecamAmt:0,etiQty:0,etiAmt:0,kdvEpias:0,kdvSisecam:0,kdvEti:0,pfk:0,cengizDsg:0,note:""}));
          setRows(seeded);__undoEnabled=false;try{localStorage.setItem(dataKey,JSON.stringify(seeded));}catch{}__undoEnabled=true;
        }
      }catch{}
    };
    reload();
    const onS=(e:StorageEvent)=>{if(!e.key||e.key===dataKey)reload();};
    window.addEventListener("storage",onS);return()=>window.removeEventListener("storage",onS);
  },[sm,sy]);

  const save=(v:ElectricRow[])=>{setRows(v);try{trackedSetItem(dataKey,JSON.stringify(v));}catch{}};
  const addRow=()=>{save([...rows,{id:"elec_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),date:"",total:0,epiasQty:0,epiasAmt:0,parkcam:0,sisecamQty:0,sisecamAmt:0,etiQty:0,etiAmt:0,kdvEpias:0,kdvSisecam:0,kdvEti:0,pfk:0,cengizDsg:0,note:""}]);};
  const upd=(id:string,f:keyof ElectricRow,v:any)=>{save(rows.map(r=>r.id===id?{...r,[f]:v}:r));};
  const del=(id:string)=>{if(!confirm("Sil?"))return;save(rows.filter(r=>r.id!==id));};

  // Drag-drop (Excel import - row B=date, C=total)
  const ensureSheetJS=async()=>{if((window as any).XLSX)return (window as any).XLSX;return new Promise<any>((resolve,reject)=>{const s=document.createElement("script");s.src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";s.onload=()=>resolve((window as any).XLSX);s.onerror=()=>reject(new Error("XLSX"));document.head.appendChild(s);});};
  const rowToElec=(row:any[]):ElectricRow|null=>{
    if(!row||row.length<1)return null;
    const str=(v:any)=>v===null||v===undefined?"":String(v).trim();
    const dateStr=(v:any)=>{if(!v)return "";if(typeof v==="number"&&v>25569&&v<200000){const d=new Date((v-25569)*86400*1000);return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;}if(v instanceof Date)return `${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;return String(v).trim();};
    const num=(v:any)=>typeof v==="number"?v:parseFloat(str(v).replace(/[,\s]/g,""))||0;
    const d=dateStr(row[0]);if(!d&&!num(row[1]))return null;
    return{id:"elec_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),date:d,total:num(row[1]),epiasQty:num(row[2]),epiasAmt:num(row[3]),parkcam:num(row[4]),sisecamQty:num(row[5]),sisecamAmt:num(row[6]),etiQty:num(row[7]),etiAmt:num(row[8]),kdvEpias:num(row[9]),kdvSisecam:num(row[10]),kdvEti:num(row[11]),pfk:num(row[12]),cengizDsg:num(row[13]),note:str(row[14])};
  };
  const handleFile=async(file:File)=>{
    try{
      setDropStatus("İşleniyor: "+file.name);
      const name=file.name.toLowerCase();let raw:any[][]=[];
      if(name.endsWith(".csv")||name.endsWith(".tsv")||name.endsWith(".txt")){
        const txt=await file.text();const sep=name.endsWith(".tsv")||txt.includes("\t")?"\t":(txt.split("\n")[0]?.includes(";")?";":",");
        raw=txt.split(/\r?\n/).filter(l=>l.trim()).map(l=>l.split(sep));
      }else if(name.endsWith(".xlsx")||name.endsWith(".xls")){
        const XLSX=await ensureSheetJS();const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array",cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];
        raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}) as any[][];
      }else{setDropStatus("Desteklenmeyen format.");setTimeout(()=>setDropStatus(""),3500);return;}
      // Electric xlsx has date in col A and total in col B (first 2 data cols). Filter for rows with a date
      raw=raw.filter(r=>r[0]&&(typeof r[0]==="number"||r[0] instanceof Date||/\d/.test(String(r[0]))));
      const parsed=raw.map(rowToElec).filter((x):x is ElectricRow=>!!x);
      if(parsed.length===0){setDropStatus("Kayıt bulunamadı.");setTimeout(()=>setDropStatus(""),3500);return;}
      save([...rows,...parsed]);setDropStatus(`✓ ${parsed.length} gün içeri aktarıldı`);setTimeout(()=>setDropStatus(""),3000);
    }catch(err){setDropStatus("Hata: "+(err instanceof Error?err.message:""));setTimeout(()=>setDropStatus(""),4000);}
  };
  const onDrop=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(false);const files=Array.from(e.dataTransfer.files);if(files.length>0)handleFile(files[0]);};
  const onDragOver=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(true);};
  const onDragLeave=(e:React.DragEvent)=>{if(e.currentTarget===e.target)setIsDraggingFile(false);};

  const resetDefaults=()=>{if(!confirm("Tüm günler silinip Excel default'larıyla doldurulacak. Emin misiniz?"))return;const seeded=ELECTRIC_DEFAULT.map((d,i)=>({id:"elec_seed_"+i,date:d.date,total:d.total,epiasQty:0,epiasAmt:0,parkcam:0,sisecamQty:0,sisecamAmt:0,etiQty:0,etiAmt:0,kdvEpias:0,kdvSisecam:0,kdvEti:0,pfk:0,cengizDsg:0,note:""}));save(seeded);};

  // Sums
  const sum=(f:keyof ElectricRow)=>rows.reduce((s,r)=>s+(typeof r[f]==="number"?(r[f] as number):0),0);
  const sumTotal=sum("total");
  const sumEpiasAmt=sum("epiasAmt");
  const sumParkcam=sum("parkcam");
  const sumSisecamAmt=sum("sisecamAmt");
  const sumEtiAmt=sum("etiAmt");
  const sumKdv=sum("kdvEpias")+sum("kdvSisecam")+sum("kdvEti");
  const sumPfk=sum("pfk");
  const sumCengiz=sum("cengizDsg");

  return<div className="relative" onDragOver={canEdit?onDragOver:undefined} onDragLeave={canEdit?onDragLeave:undefined} onDrop={canEdit?onDrop:undefined}>
    {isDraggingFile&&canEdit&&<div className="absolute inset-0 z-50 bg-amber-500/20 border-4 border-dashed border-amber-500 rounded-lg flex items-center justify-center pointer-events-none"><div className="bg-white px-6 py-4 rounded-xl shadow-xl border border-amber-200"><div className="text-2xl font-bold text-amber-700 mb-1">📥 Dosyayı buraya bırakın</div><div className="text-[11px] text-gray-500 text-center">.xlsx • .csv • .tsv</div></div></div>}
    {dropStatus&&<div className={"mb-3 px-3 py-2 rounded-lg text-xs font-semibold "+(dropStatus.startsWith("✓")?"bg-emerald-50 text-emerald-700 border border-emerald-200":dropStatus.startsWith("Hata")||dropStatus.startsWith("Desteklenmeyen")||dropStatus.startsWith("Kayıt")?"bg-red-50 text-red-700 border border-red-200":"bg-blue-50 text-blue-700 border border-blue-200")}>{dropStatus}</div>}

    <div className="flex justify-between items-center mb-3 flex-wrap gap-2">
      <div><span className="text-sm font-bold text-amber-700">Electric</span><span className="text-[10px] text-gray-400 ml-2">Günlük elektrik tahsilatları — {MONTHS[sm]} {sy} — {rows.length} gün</span><span className="text-[10px] text-amber-500 ml-2 hidden sm:inline">• Excel sürükleyip bırakabilirsiniz</span></div>
      <div className="flex gap-2 flex-wrap">
        {canEdit&&<label className="px-3 py-1 bg-blue-100 hover:bg-blue-200 text-blue-700 rounded text-[11px] font-semibold cursor-pointer">📁 Dosya Seç<input type="file" accept=".xlsx,.xls,.csv,.tsv,.txt" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f);e.target.value="";}}/></label>}
        {canEdit&&<button onClick={resetDefaults} className="px-3 py-1 bg-gray-100 hover:bg-gray-200 text-gray-600 rounded text-[11px]">↻ Sıfırla</button>}
        <ExportBtn tableId="electricTbl" name={"Electric_"+MONTHS[sm]+"_"+sy}/>
      </div>
    </div>

    {/* Summary */}
    <div className="mb-4 p-3 bg-gradient-to-r from-amber-50 via-white to-amber-50 border border-amber-200 rounded-lg">
      <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-3">
        <div><div className="text-[9px] text-gray-500 uppercase">Toplam Tahsilat</div><div className="text-sm font-mono font-bold text-amber-800">{fmt(sumTotal)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Epiaş</div><div className="text-[11px] font-mono font-bold">{fmt(sumEpiasAmt)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Parkcam</div><div className="text-[11px] font-mono font-bold">{fmt(sumParkcam)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Şişecam</div><div className="text-[11px] font-mono font-bold">{fmt(sumSisecamAmt)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Eti Soda</div><div className="text-[11px] font-mono font-bold">{fmt(sumEtiAmt)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">KDV Top.</div><div className="text-[11px] font-mono font-bold text-green-700">{fmt(sumKdv)}</div></div>
      </div>
    </div>

    {canEdit&&<div className="mb-3 flex justify-end"><button onClick={addRow} className="px-3 py-1 bg-amber-600 hover:bg-amber-700 text-white rounded text-[11px] font-semibold shadow-sm">+ Gün Ekle</button></div>}

    <div className="overflow-x-auto border border-amber-200 rounded-lg">
      <table id="electricTbl" className="border-collapse text-[10px] w-full" style={{minWidth:1400}}>
        <thead>
          <tr className="bg-amber-50 border-b-2 border-amber-300">
            <th className="px-1 py-1 text-[9px] font-bold text-amber-900 text-left" rowSpan={2}>#</th>
            <th className="px-1 py-1 text-[9px] font-bold text-amber-900 text-center" rowSpan={2}>Ödeme Tarihi</th>
            <th className="px-1 py-1 text-[9px] font-bold text-amber-900 text-right" rowSpan={2}>Toplam Tahsilat</th>
            <th className="px-1 py-1 text-[9px] font-bold text-red-700 text-center" colSpan={4} style={{borderLeft:"2px solid #fcd34d"}}>Grup Dışı</th>
            <th className="px-1 py-1 text-[9px] font-bold text-blue-700 text-center" colSpan={2} style={{borderLeft:"2px solid #fcd34d"}}>Grup İçi</th>
            <th className="px-1 py-1 text-[9px] font-bold text-green-700 text-center" colSpan={3} style={{borderLeft:"2px solid #fcd34d"}}>KDV Gelirleri</th>
            <th className="px-1 py-1 text-[9px] font-bold text-purple-700 text-center" colSpan={2} style={{borderLeft:"2px solid #fcd34d"}}>Ek</th>
            <th className="px-1 py-1 text-[9px] font-bold text-amber-900" rowSpan={2} style={{borderLeft:"2px solid #fcd34d"}}>Not</th>
            {canEdit&&<th className="w-6" rowSpan={2}></th>}
          </tr>
          <tr className="bg-amber-50 border-b border-amber-300">
            <th className="px-1 py-1 text-[8px] font-semibold text-red-700 text-right" style={{borderLeft:"2px solid #fcd34d"}}>Epiaş Mkt</th>
            <th className="px-1 py-1 text-[8px] font-semibold text-red-700 text-right">Epiaş Tutar</th>
            <th className="px-1 py-1 text-[8px] font-semibold text-red-700 text-right">Parkcam</th>
            <th className="px-1 py-1 text-[8px] font-semibold text-red-700 text-right">Şişecam Mkt</th>
            <th className="px-1 py-1 text-[8px] font-semibold text-blue-700 text-right" style={{borderLeft:"2px solid #fcd34d"}}>Eti Mkt</th>
            <th className="px-1 py-1 text-[8px] font-semibold text-blue-700 text-right">Eti Tutar</th>
            <th className="px-1 py-1 text-[8px] font-semibold text-green-700 text-right" style={{borderLeft:"2px solid #fcd34d"}}>Epiaş KDV</th>
            <th className="px-1 py-1 text-[8px] font-semibold text-green-700 text-right">Şişecam KDV</th>
            <th className="px-1 py-1 text-[8px] font-semibold text-green-700 text-right">Eti KDV</th>
            <th className="px-1 py-1 text-[8px] font-semibold text-purple-700 text-right" style={{borderLeft:"2px solid #fcd34d"}}>PFK</th>
            <th className="px-1 py-1 text-[8px] font-semibold text-purple-700 text-right">Cengiz DSG</th>
          </tr>
        </thead>
        <tbody>{rows.length===0?<tr><td colSpan={canEdit?16:15} className="py-10 text-center text-gray-400">Kayıt yok.</td></tr>:rows.map((r,i)=><tr key={r.id} className="border-b border-gray-100 hover:bg-amber-50/30">
          <td className="px-1 py-0.5 text-[9px] text-gray-400">{i+1}</td>
          <td className="px-1 py-0.5" style={{minWidth:85}}><DateCell value={r.date} onChange={v=>upd(r.id,"date",v)} disabled={!canEdit}/></td>
          <td className="px-1 py-0.5 bg-amber-50/30" style={{minWidth:90}}><input className={inp+" text-right font-mono text-[10px] font-semibold"} value={addC(r.total.toString())} onChange={e=>upd(r.id,"total",parseFloat(stripC(e.target.value))||0)}/></td>
          <td className="px-1 py-0.5" style={{borderLeft:"2px solid #fcd34d"}}><input className={inp+" text-right font-mono text-[10px]"} value={r.epiasQty||""} onChange={e=>upd(r.id,"epiasQty",parseFloat(e.target.value)||0)} placeholder="0"/></td>
          <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={r.epiasAmt?addC(r.epiasAmt.toString()):""} onChange={e=>upd(r.id,"epiasAmt",parseFloat(stripC(e.target.value))||0)} placeholder="0"/></td>
          <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={r.parkcam?addC(r.parkcam.toString()):""} onChange={e=>upd(r.id,"parkcam",parseFloat(stripC(e.target.value))||0)} placeholder="0"/></td>
          <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={r.sisecamQty||""} onChange={e=>upd(r.id,"sisecamQty",parseFloat(e.target.value)||0)} placeholder="0"/></td>
          <td className="px-1 py-0.5" style={{borderLeft:"2px solid #fcd34d"}}><input className={inp+" text-right font-mono text-[10px]"} value={r.etiQty||""} onChange={e=>upd(r.id,"etiQty",parseFloat(e.target.value)||0)} placeholder="0"/></td>
          <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={r.etiAmt?addC(r.etiAmt.toString()):""} onChange={e=>upd(r.id,"etiAmt",parseFloat(stripC(e.target.value))||0)} placeholder="0"/></td>
          <td className="px-1 py-0.5" style={{borderLeft:"2px solid #fcd34d"}}><input className={inp+" text-right font-mono text-[10px]"} value={r.kdvEpias?addC(r.kdvEpias.toString()):""} onChange={e=>upd(r.id,"kdvEpias",parseFloat(stripC(e.target.value))||0)} placeholder="0"/></td>
          <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={r.kdvSisecam?addC(r.kdvSisecam.toString()):""} onChange={e=>upd(r.id,"kdvSisecam",parseFloat(stripC(e.target.value))||0)} placeholder="0"/></td>
          <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={r.kdvEti?addC(r.kdvEti.toString()):""} onChange={e=>upd(r.id,"kdvEti",parseFloat(stripC(e.target.value))||0)} placeholder="0"/></td>
          <td className="px-1 py-0.5" style={{borderLeft:"2px solid #fcd34d"}}><input className={inp+" text-right font-mono text-[10px]"} value={r.pfk?addC(r.pfk.toString()):""} onChange={e=>upd(r.id,"pfk",parseFloat(stripC(e.target.value))||0)} placeholder="0"/></td>
          <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={r.cengizDsg?addC(r.cengizDsg.toString()):""} onChange={e=>upd(r.id,"cengizDsg",parseFloat(stripC(e.target.value))||0)} placeholder="0"/></td>
          <td className="px-1 py-0.5" style={{borderLeft:"2px solid #fcd34d",minWidth:80}}><input className={inp+" text-[10px]"} value={r.note} onChange={e=>upd(r.id,"note",e.target.value)} placeholder="not"/></td>
          {canEdit&&<td className="px-1 py-0.5"><button onClick={()=>del(r.id)} className="px-1 py-0.5 text-red-400 hover:bg-red-50 hover:text-red-600 rounded text-[11px]">×</button></td>}
        </tr>)}</tbody>
        <tfoot><tr className="bg-amber-100 font-bold border-t-2 border-amber-400">
          <td colSpan={2} className="px-2 py-1 text-[10px] text-amber-900 text-right">TOPLAM</td>
          <td className="px-1 py-1 text-right font-mono text-[11px] text-amber-900">{fmt(sumTotal)}</td>
          <td style={{borderLeft:"2px solid #fcd34d"}}></td>
          <td className="px-1 py-1 text-right font-mono text-[10px] text-red-700">{fmt(sumEpiasAmt)}</td>
          <td className="px-1 py-1 text-right font-mono text-[10px] text-red-700">{fmt(sumParkcam)}</td>
          <td></td>
          <td style={{borderLeft:"2px solid #fcd34d"}}></td>
          <td className="px-1 py-1 text-right font-mono text-[10px] text-blue-700">{fmt(sumEtiAmt)}</td>
          <td style={{borderLeft:"2px solid #fcd34d"}} className="px-1 py-1 text-right font-mono text-[10px] text-green-700">{fmt(sum("kdvEpias"))}</td>
          <td className="px-1 py-1 text-right font-mono text-[10px] text-green-700">{fmt(sum("kdvSisecam"))}</td>
          <td className="px-1 py-1 text-right font-mono text-[10px] text-green-700">{fmt(sum("kdvEti"))}</td>
          <td style={{borderLeft:"2px solid #fcd34d"}} className="px-1 py-1 text-right font-mono text-[10px] text-purple-700">{fmt(sumPfk)}</td>
          <td className="px-1 py-1 text-right font-mono text-[10px] text-purple-700">{fmt(sumCengiz)}</td>
          <td style={{borderLeft:"2px solid #fcd34d"}}></td>
          {canEdit&&<td></td>}
        </tr></tfoot>
      </table>
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════
// OTHER DISC. — invoice discount with Net Amount + Disc.Cost per customer
// ═══════════════════════════════════════════════════════════════════════════

type OtherDiscInv={id:string,discDate:string,customer:string,invoice:string,amount:number,dueDate:string,netAmount:number,beneBank:string,discCost:number};
type OtherDiscCust={id:string,name:string,limit:number,rate:number};

function OtherDiscSheet({sm,sy,canEdit}:{sm:number,sy:number,canEdit:boolean}){
  const[invoices,setInvoices]=useState<OtherDiscInv[]>([]);
  const[customers,setCustomers]=useState<OtherDiscCust[]>([]);
  const[sortBy,setSortBy]=useState<string>("dueDate");
  const[filterCust,setFilterCust]=useState("");
  const[fInvoice,setFInvoice]=useState("");
  const[fDiscDate,setFDiscDate]=useState("");
  const[fCust,setFCust]=useState("");
  const[fAmount,setFAmount]=useState("");
  const[fDueDate,setFDueDate]=useState("");
  const[showLimits,setShowLimits]=useState(true);
  const[isDraggingFile,setIsDraggingFile]=useState(false);
  const[dropStatus,setDropStatus]=useState<string>("");

  const invKey="otherdisc_inv_"+sy+"_"+sm;
  const custKey="otherdisc_cust_"+sy+"_"+sm;

  useEffect(()=>{
    const reload=()=>{
      try{
        const si=localStorage.getItem(invKey);
        if(si){setInvoices(JSON.parse(si));}
        else{const seeded=OTHERDISC_DEFAULT.map((d,i)=>({id:"od_inv_seed_"+i,...d}));setInvoices(seeded);__undoEnabled=false;try{localStorage.setItem(invKey,JSON.stringify(seeded));}catch{}__undoEnabled=true;}
        const sc=localStorage.getItem(custKey);
        if(sc){setCustomers(JSON.parse(sc));}
        else{const seeded=OTHERDISC_CUSTOMERS.map((d,i)=>({id:"od_cust_seed_"+i,...d}));setCustomers(seeded);__undoEnabled=false;try{localStorage.setItem(custKey,JSON.stringify(seeded));}catch{}__undoEnabled=true;}
      }catch{}
    };
    reload();
    const onS=(e:StorageEvent)=>{if(!e.key||e.key===invKey||e.key===custKey)reload();};
    window.addEventListener("storage",onS);return()=>window.removeEventListener("storage",onS);
  },[sm,sy]);

  const saveInv=(v:OtherDiscInv[])=>{setInvoices(v);try{trackedSetItem(invKey,JSON.stringify(v));}catch{}};
  const saveCust=(v:OtherDiscCust[])=>{setCustomers(v);try{trackedSetItem(custKey,JSON.stringify(v));}catch{}};

  const addInv=()=>{saveInv([...invoices,{id:"od_inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),discDate:"",customer:"",invoice:"",amount:0,dueDate:"",netAmount:0,beneBank:"",discCost:0}]);};
  const updInv=(id:string,f:keyof OtherDiscInv,v:any)=>{saveInv(invoices.map(r=>r.id===id?{...r,[f]:v}:r));};
  const delInv=(id:string)=>{if(!confirm("Sil?"))return;saveInv(invoices.filter(r=>r.id!==id));};
  const addCust=()=>{saveCust([...customers,{id:"od_cust_"+Date.now(),name:"",limit:0,rate:0}]);};
  const updCust=(id:string,f:keyof OtherDiscCust,v:any)=>{saveCust(customers.map(r=>r.id===id?{...r,[f]:v}:r));};
  const delCust=(id:string)=>{if(!confirm("Sil?"))return;saveCust(customers.filter(r=>r.id!==id));};

  const ensureSheetJS=async()=>{if((window as any).XLSX)return (window as any).XLSX;return new Promise<any>((resolve,reject)=>{const s=document.createElement("script");s.src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";s.onload=()=>resolve((window as any).XLSX);s.onerror=()=>reject(new Error("XLSX"));document.head.appendChild(s);});};
  const rowToInv=(row:any[]):OtherDiscInv|null=>{
    if(!row||row.length<1)return null;
    const str=(v:any)=>v===null||v===undefined?"":String(v).trim();
    const dateStr=(v:any)=>{if(!v)return "";if(typeof v==="number"&&v>25569&&v<200000){const d=new Date((v-25569)*86400*1000);return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;}if(v instanceof Date)return `${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;return String(v).trim();};
    const num=(v:any)=>typeof v==="number"?v:parseFloat(str(v).replace(/[,\s]/g,""))||0;
    const inv=str(row[2]);const cust=str(row[1]);if(!inv&&!cust)return null;
    return{id:"od_inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),discDate:dateStr(row[0]),customer:cust,invoice:inv,amount:num(row[3]),dueDate:dateStr(row[4]),netAmount:num(row[5]),beneBank:str(row[6]),discCost:num(row[7])};
  };
  const handleFile=async(file:File)=>{
    try{
      setDropStatus("İşleniyor: "+file.name);
      const name=file.name.toLowerCase();let raw:any[][]=[];
      if(name.endsWith(".csv")||name.endsWith(".tsv")||name.endsWith(".txt")){
        const txt=await file.text();const sep=name.endsWith(".tsv")||txt.includes("\t")?"\t":(txt.split("\n")[0]?.includes(";")?";":",");
        raw=txt.split(/\r?\n/).filter(l=>l.trim()).map(l=>l.split(sep));
        if(raw[0]&&/disc|date|customer|inv/i.test(String(raw[0][0]||"")+String(raw[0][1]||"")))raw=raw.slice(1);
      }else if(name.endsWith(".xlsx")||name.endsWith(".xls")){
        const XLSX=await ensureSheetJS();const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array",cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];
        raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}) as any[][];
        raw=raw.map(r=>r.slice(1));// skip col A (empty in Excel)
        raw=raw.filter(r=>r[1]&&!/customer|banks|total|remained|discount/i.test(String(r[1])));
      }else{setDropStatus("Desteklenmeyen format.");setTimeout(()=>setDropStatus(""),3500);return;}
      const parsed=raw.map(rowToInv).filter((x):x is OtherDiscInv=>!!x);
      if(parsed.length===0){setDropStatus("Fatura bulunamadı.");setTimeout(()=>setDropStatus(""),3500);return;}
      saveInv([...invoices,...parsed]);setDropStatus(`✓ ${parsed.length} fatura içeri aktarıldı`);setTimeout(()=>setDropStatus(""),3000);
    }catch(err){setDropStatus("Hata: "+(err instanceof Error?err.message:""));setTimeout(()=>setDropStatus(""),4000);}
  };
  const onDrop=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(false);const files=Array.from(e.dataTransfer.files);if(files.length>0)handleFile(files[0]);};
  const onDragOver=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(true);};
  const onDragLeave=(e:React.DragEvent)=>{if(e.currentTarget===e.target)setIsDraggingFile(false);};

  // Sums
  const totalAmount=invoices.reduce((s,r)=>s+r.amount,0);
  const totalNet=invoices.reduce((s,r)=>s+r.netAmount,0);
  const totalDiscCost=invoices.reduce((s,r)=>s+r.discCost,0);
  const custRisk=(name:string)=>invoices.filter(r=>r.customer===name).reduce((s,r)=>s+r.amount,0);
  const totalRisk=customers.reduce((s,c)=>s+custRisk(c.name),0);
  const totalLimit=customers.reduce((s,c)=>s+c.limit,0);

  const custSet=[...new Set(invoices.map(r=>r.customer).filter(c=>c))].sort();
  let view=[...invoices];
  if(filterCust)view=view.filter(r=>r.customer===filterCust);
  const ci=(s:string)=>s.toLowerCase();
  if(fInvoice)view=view.filter(r=>ci(r.invoice).includes(ci(fInvoice)));
  if(fDiscDate)view=view.filter(r=>ci(r.discDate).includes(ci(fDiscDate)));
  if(fCust)view=view.filter(r=>ci(r.customer).includes(ci(fCust)));
  if(fAmount)view=view.filter(r=>r.amount.toString().includes(fAmount.replace(/[,\s]/g,"")));
  if(fDueDate)view=view.filter(r=>ci(r.dueDate).includes(ci(fDueDate)));
  view.sort((a,b)=>{
    if(sortBy==="dueDate"||sortBy==="discDate"){const ka=sortBy as "dueDate"|"discDate";return parseDDMMYYYY(a[ka])-parseDDMMYYYY(b[ka]);}
    if(sortBy==="amount")return a.amount-b.amount;
    if(sortBy==="customer")return a.customer.localeCompare(b.customer);
    if(sortBy==="invoice")return a.invoice.localeCompare(b.invoice);
    return 0;
  });

  const resetDefaults=()=>{if(!confirm("Tüm kayıtlar silinip Excel default'larıyla doldurulacak. Emin misiniz?"))return;saveInv(OTHERDISC_DEFAULT.map((d,i)=>({id:"od_inv_seed_"+i,...d})));saveCust(OTHERDISC_CUSTOMERS.map((d,i)=>({id:"od_cust_seed_"+i,...d})));};

  return<div className="relative" onDragOver={canEdit?onDragOver:undefined} onDragLeave={canEdit?onDragLeave:undefined} onDrop={canEdit?onDrop:undefined}>
    {isDraggingFile&&canEdit&&<div className="absolute inset-0 z-50 bg-lime-500/20 border-4 border-dashed border-lime-500 rounded-lg flex items-center justify-center pointer-events-none"><div className="bg-white px-6 py-4 rounded-xl shadow-xl border border-lime-200"><div className="text-2xl font-bold text-lime-700 mb-1">📥 Dosyayı buraya bırakın</div><div className="text-[11px] text-gray-500 text-center">.xlsx • .csv • .tsv</div></div></div>}
    {dropStatus&&<div className={"mb-3 px-3 py-2 rounded-lg text-xs font-semibold "+(dropStatus.startsWith("✓")?"bg-emerald-50 text-emerald-700 border border-emerald-200":dropStatus.startsWith("Hata")||dropStatus.startsWith("Desteklenmeyen")||dropStatus.startsWith("Fatura")?"bg-red-50 text-red-700 border border-red-200":"bg-blue-50 text-blue-700 border border-blue-200")}>{dropStatus}</div>}

    <div className="flex justify-between items-center mb-3 flex-wrap gap-2">
      <div><span className="text-sm font-bold text-lime-700">Other Disc.</span><span className="text-[10px] text-gray-400 ml-2">Other invoice discount — {MONTHS[sm]} {sy} — {invoices.length} fatura</span><span className="text-[10px] text-lime-500 ml-2 hidden sm:inline">• Excel sürükleyip bırakabilirsiniz</span></div>
      <div className="flex gap-2 flex-wrap">
        {canEdit&&<label className="px-3 py-1 bg-blue-100 hover:bg-blue-200 text-blue-700 rounded text-[11px] font-semibold cursor-pointer">📁 Dosya Seç<input type="file" accept=".xlsx,.xls,.csv,.tsv,.txt" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f);e.target.value="";}}/></label>}
        {canEdit&&<button onClick={resetDefaults} className="px-3 py-1 bg-gray-100 hover:bg-gray-200 text-gray-600 rounded text-[11px]">↻ Sıfırla</button>}
        <ExportBtn tableId="odInvTbl" name={"Other_Disc_"+MONTHS[sm]+"_"+sy}/>
      </div>
    </div>

    <div className="mb-4 p-3 bg-gradient-to-r from-lime-50 via-white to-lime-50 border border-lime-200 rounded-lg">
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <div><div className="text-[9px] text-gray-500 uppercase">Toplam Inv. Amt</div><div className="text-[12px] font-mono font-bold text-lime-700">{fmt(totalAmount)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Toplam Net Amt</div><div className="text-[12px] font-mono font-bold text-emerald-600">{fmt(totalNet)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Toplam Disc.Cost</div><div className="text-[12px] font-mono font-bold text-orange-600">{fmt(totalDiscCost)}</div></div>
        <div><div className="text-[9px] text-gray-500 uppercase">Total Risk / Limit</div><div className="text-[11px] font-mono font-bold text-lime-800">{fmt(totalRisk)} / {fmt(totalLimit)}</div></div>
      </div>
    </div>

    {invoices.length>0&&<div className="flex gap-2 mb-3 flex-wrap text-[10px] items-center">
      <div className="flex items-center gap-1"><span className="text-gray-500">Customer:</span><select value={filterCust} onChange={e=>setFilterCust(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="">Tümü ({custSet.length})</option>{custSet.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
      <div className="flex items-center gap-1"><span className="text-gray-500">Sırala:</span><select value={sortBy} onChange={e=>setSortBy(e.target.value)} className="px-2 py-1 border border-gray-300 rounded text-[10px] bg-white"><option value="dueDate">Due Date</option><option value="discDate">Discount Date</option><option value="customer">Customer</option><option value="amount">Amount</option><option value="invoice">Invoice No</option></select></div>
      {(filterCust||fInvoice||fDiscDate||fCust||fAmount||fDueDate)&&<button onClick={()=>{setFilterCust("");setFInvoice("");setFDiscDate("");setFCust("");setFAmount("");setFDueDate("");}} className="px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded text-[10px]">Tüm Filtreleri Temizle</button>}
      <span className="text-gray-400">{view.length} / {invoices.length} fatura</span>
      {canEdit&&<button onClick={addInv} className="ml-auto px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[11px] font-semibold shadow-sm">+ Fatura Ekle</button>}
    </div>}

    <div className="grid lg:grid-cols-[1fr_380px] gap-4">
      <div>
        <table id="odInvTbl" className="border-collapse text-[10px] w-full" style={{minWidth:980}}>
          <thead>
            <tr className="border-b-2 border-lime-300 bg-lime-50 sticky top-0">
              <th className="px-2 py-2 text-left text-[9px] font-semibold text-lime-800" style={{minWidth:25}}>#</th>
              <th className="px-2 py-2 text-center text-[9px] font-semibold text-lime-800" style={{minWidth:85}}>Discount Date</th>
              <th className="px-2 py-2 text-left text-[9px] font-semibold text-lime-800" style={{minWidth:110}}>Customer</th>
              <th className="px-2 py-2 text-left text-[9px] font-semibold text-lime-800" style={{minWidth:140}}>Inv Number</th>
              <th className="px-2 py-2 text-right text-[9px] font-semibold text-lime-800" style={{minWidth:100}}>Inv Amount</th>
              <th className="px-2 py-2 text-center text-[9px] font-semibold text-lime-800" style={{minWidth:85}}>Due Date</th>
              <th className="px-2 py-2 text-right text-[9px] font-semibold text-lime-800" style={{minWidth:100}}>Net Amount</th>
              <th className="px-2 py-2 text-left text-[9px] font-semibold text-lime-800" style={{minWidth:90}}>Bene. Bank</th>
              <th className="px-2 py-2 text-right text-[9px] font-semibold text-lime-800" style={{minWidth:90}}>Disc. Cost</th>
              {canEdit&&<th className="w-6"></th>}
            </tr>
            <tr className="border-b border-lime-200 bg-lime-50/60">
              <th></th>
              <th className="px-1 py-1"><input value={fDiscDate} onChange={e=>setFDiscDate(e.target.value)} placeholder="🔍" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fCust} onChange={e=>setFCust(e.target.value)} placeholder="🔍" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fInvoice} onChange={e=>setFInvoice(e.target.value)} placeholder="🔍" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] font-mono bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fAmount} onChange={e=>setFAmount(e.target.value)} placeholder="🔍" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] text-right font-mono bg-white outline-none placeholder:text-gray-300"/></th>
              <th className="px-1 py-1"><input value={fDueDate} onChange={e=>setFDueDate(e.target.value)} placeholder="🔍" className="w-full px-1 py-0.5 border border-gray-300 rounded text-[9px] bg-white outline-none placeholder:text-gray-300"/></th>
              <th></th><th></th><th></th>
              {canEdit&&<th></th>}
            </tr>
          </thead>
          <tbody>{view.length===0?<tr><td colSpan={canEdit?10:9} className="py-10 text-center text-gray-400">Fatura yok.</td></tr>:view.map((r,i)=><tr key={r.id} className="border-b border-gray-100 hover:bg-lime-50/30">
            <td className="px-2 py-0.5 text-[9px] text-gray-400">{i+1}</td>
            <td className="px-1 py-0.5"><DateCell value={r.discDate} onChange={v=>updInv(r.id,"discDate",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><ObligorCombo value={r.customer} options={customers.map(c=>c.name).filter(n=>n).sort()} onChange={v=>updInv(r.id,"customer",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><input className={inp+" font-mono text-[10px]"} value={r.invoice} onChange={e=>updInv(r.id,"invoice",e.target.value)}/></td>
            <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={addC(r.amount.toString())} onChange={e=>updInv(r.id,"amount",parseFloat(stripC(e.target.value))||0)}/></td>
            <td className="px-1 py-0.5"><DateCell value={r.dueDate} onChange={v=>updInv(r.id,"dueDate",v)} disabled={!canEdit}/></td>
            <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={addC(r.netAmount.toString())} onChange={e=>updInv(r.id,"netAmount",parseFloat(stripC(e.target.value))||0)}/></td>
            <td className="px-1 py-0.5"><input className={inp+" text-[10px]"} value={r.beneBank} onChange={e=>updInv(r.id,"beneBank",e.target.value)}/></td>
            <td className="px-1 py-0.5"><input className={inp+" text-right font-mono text-[10px]"} value={addC(r.discCost.toString())} onChange={e=>updInv(r.id,"discCost",parseFloat(stripC(e.target.value))||0)}/></td>
            {canEdit&&<td className="px-1 py-0.5"><button onClick={()=>delInv(r.id)} className="px-1 py-0.5 text-red-400 hover:bg-red-50 hover:text-red-600 rounded text-[11px]">×</button></td>}
          </tr>)}</tbody>
        </table>
      </div>

      <div className="lg:sticky lg:top-2 lg:self-start">
        <div className="rounded-lg border border-lime-200 bg-white">
          <div className="flex justify-between items-center px-3 py-2 border-b border-lime-200 bg-lime-50 cursor-pointer" onClick={()=>setShowLimits(!showLimits)}>
            <div className="text-[11px] font-bold text-lime-800">{showLimits?"▼":"▶"} Customer Limits ({customers.length})</div>
            {canEdit&&showLimits&&<button onClick={(e)=>{e.stopPropagation();addCust();}} className="px-2 py-0.5 bg-lime-600 hover:bg-lime-700 text-white rounded text-[9px] font-semibold">+ Customer</button>}
          </div>
          {showLimits&&<div className="overflow-x-auto">
            <table className="border-collapse text-[9px] w-full" style={{minWidth:360}}>
              <thead><tr className="bg-lime-100/70 border-b border-lime-200">
                <th className="px-1 py-1 text-left text-[8px] font-bold text-lime-900" style={{minWidth:100}}>Customer</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-lime-900" style={{minWidth:80}}>Current Risk</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-lime-900" style={{minWidth:80}}>Limit</th>
                <th className="px-1 py-1 text-right text-[8px] font-bold text-lime-900" style={{minWidth:60}}>Rate</th>
                {canEdit&&<th className="w-4"></th>}
              </tr></thead>
              <tbody>{customers.map(c=>{const risk=custRisk(c.name);return<tr key={c.id} className="border-b border-gray-100 hover:bg-lime-50/30">
                <td className="px-1 py-0.5"><input className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white outline-none" value={c.name} onChange={e=>updCust(c.id,"name",e.target.value)}/></td>
                <td className={"px-1 py-0.5 text-right font-mono text-[9px] "+(risk>0?"text-lime-700 font-semibold":"text-gray-300")}>{risk?fmt(risk):"-"}</td>
                <td className="px-0.5 py-0.5"><input type="number" className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white text-right font-mono outline-none" value={c.limit||""} onChange={e=>updCust(c.id,"limit",parseFloat(e.target.value)||0)}/></td>
                <td className="px-0.5 py-0.5"><input type="number" step="0.0001" className="w-full px-1 py-0.5 border border-gray-200 rounded text-[9px] bg-white text-right font-mono outline-none" value={c.rate||""} onChange={e=>updCust(c.id,"rate",parseFloat(e.target.value)||0)}/></td>
                {canEdit&&<td className="px-0.5 py-0.5"><button onClick={()=>delCust(c.id)} className="px-0.5 py-0 text-red-400 hover:text-red-600 text-[10px]">×</button></td>}
              </tr>;})}
              <tr className="bg-lime-100 border-t-2 border-lime-300 font-bold"><td className="px-1 py-1 text-[9px] text-lime-900">Total Risk</td><td className="px-1 py-1 text-right font-mono text-[9px] text-lime-900">{fmt(totalRisk)}</td><td className="px-1 py-1 text-right font-mono text-[9px] text-lime-900">{fmt(totalLimit)}</td><td></td>{canEdit&&<td></td>}</tr>
              <tr className="bg-emerald-50 font-bold"><td className="px-1 py-1 text-[9px] text-emerald-900">Remained Limit</td><td colSpan={2} className="px-1 py-1 text-right font-mono text-[9px] text-emerald-800">{fmt(totalLimit-totalRisk)}</td><td></td>{canEdit&&<td></td>}</tr>
              </tbody>
            </table>
          </div>}
          <div className="px-3 py-2 bg-lime-50/40 border-t border-lime-200 text-[9px] text-lime-700">
            <div><b>Current Risk</b> = SUMIF(Customer, Inv Amount)</div>
            <div><b>Rate</b> = faiz oranı (örn. 0.08275 = %8.275)</div>
          </div>
        </div>
      </div>
    </div>
  </div>;
}


function CashflowWithSubs({canEdit,sm,sy,fx}:{canEdit:boolean,sm:number,sy:number,fx:typeof FX_D}){
  const[subTab,setSubTab]=useState("main");
  return<div>
    <div className="flex border-b border-gray-200 overflow-x-auto mb-4 bg-gray-50/60 rounded-t-lg">
      {CF_SUB_TABS.map(t=><button key={t.id} onClick={()=>setSubTab(t.id)} className={"px-3 py-2 text-[11px] whitespace-nowrap border-b-2 transition "+(subTab===t.id?"font-semibold border-current bg-white":"border-transparent text-gray-400 hover:text-gray-600 hover:bg-white/60")} style={subTab===t.id?{color:t.color}:{}}>{t.label}</button>)}
    </div>
    {subTab==="main"&&<Cashflow canEdit={canEdit} sm={sm} sy={sy} fx={fx}/>}
    {subTab==="export"&&<SubSheet type="uk" sm={sm} sy={sy} fx={fx} canEdit={canEdit}/>}
    {subTab==="usexport"&&<SubSheet type="us" sm={sm} sy={sy} fx={fx} canEdit={canEdit}/>}
    {subTab==="domestic"&&<SubSheet type="dom" sm={sm} sy={sy} fx={fx} canEdit={canEdit}/>}
    {subTab==="bnp_truk"&&<BnpTrUkSheet sm={sm} sy={sy} canEdit={canEdit}/>}
    {subTab==="abc"&&<AbcSheet sm={sm} sy={sy} canEdit={canEdit}/>}
    {subTab==="saisa_rf"&&<SaisaSheet sm={sm} sy={sy} canEdit={canEdit}/>}
    {subTab==="bnp_us"&&<BnpUsSheet sm={sm} sy={sy} canEdit={canEdit}/>}
    {subTab==="other_disc"&&<OtherDiscSheet sm={sm} sy={sy} canEdit={canEdit}/>}
    {subTab==="electric"&&<ElectricSheet sm={sm} sy={sy} canEdit={canEdit}/>}
    {subTab==="cheques"&&<ChequesSheet sm={sm} sy={sy} canEdit={canEdit} fx={fx}/>}
    {subTab==="garanti_factoring"&&<FactoringSheet sm={sm} sy={sy} canEdit={canEdit} kind="garanti" color="#16a34a" label="Garanti Factoring"/>}
    {subTab==="yk_factoring"&&<FactoringSheet sm={sm} sy={sy} canEdit={canEdit} kind="yk" color="#e11d48" label="YK Factoring"/>}
  </div>;
}



// ═══════════════════════════════════════════════════════════════════════════
// INVOICES — We Soda invoice portal tracker
// Paste the "Daily Reporting" Excel export and get a categorized overview.
// Historical invoices are remembered across pastes so responsible persons
// and notes can be recovered automatically for recurring invoices.
// ═══════════════════════════════════════════════════════════════════════════

const INVOICE_HISTORY_DEFAULT=[{description:"Pwc (Sermaye Piyasalari Işlem Desteği ve Yürütülmesi ve Bono Danismanlik)",usd:0,eur:0,gbp:181200.0,totalUSD:245127.36,person:"Ali Bekce",dueDate:"22/04/2026",days:107,note:""},{description:"Colliers (College Hill Q4 2025 Tum Katlar Aidat ve Faturalar)",usd:0,eur:0,gbp:159047.63,totalUSD:215159.633864,person:"Henry Penn",dueDate:"31/12/2025",days:0,note:""},{description:"Colliers (College Hill Q1 2026 Tum Katlar Aidat ve Faturalar)",usd:0,eur:0,gbp:159047.63,totalUSD:215159.633864,person:"Henry Penn",dueDate:"31/03/2026",days:17,note:"Dikkat"},{description:"Tag (Ucus ve Konaklama)",usd:0,eur:0,gbp:156388.82,totalUSD:211562.79569600002,person:"",dueDate:"17/04/2026",days:37,note:"TAG Bakiyesi"},{description:"Pwc (Sermaye Piyasalari Işlem Desteği ve Yürütülmesi)",usd:0,eur:0,gbp:156000.0,totalUSD:211036.8,person:"Ali Bekce",dueDate:"11/03/2026",days:21,note:"Colliers"},{description:"Paul Hastings (Hukuki Danismanlik)",usd:192910.95,eur:0,gbp:0,totalUSD:192910.95,person:"M. Ali Erdogan",dueDate:"24/04/2026",days:0,note:""},{description:"DSS (WPS ve PSM Faz 2)",usd:192411.27,eur:0,gbp:0,totalUSD:192411.27,person:"A. Warren",dueDate:"27/03/2026",days:0,note:""},{description:"Paul Hastings (Hukuki Danismanlik)",usd:103041.67000000001,eur:0,gbp:0,totalUSD:103041.67000000001,person:"M. Ali Erdogan",dueDate:"09/03/2026",days:1,note:""},{description:"Cefic (Batarya Grubu Yillik Uyelik Ucreti)",usd:0,eur:65570.33,gbp:0,totalUSD:77274.63390500001,person:"Alan Knight",dueDate:"19/03/2026",days:29,note:""},{description:"Heads Resourcing (Ise Alim Danismanlik-Senior Director, Total Rewards)",usd:0,eur:0,gbp:57120.0,totalUSD:77271.936,person:"Sevgul Erdogan",dueDate:"23/02/2026",days:8,note:""},{description:"Empresaria (Ise Alim Danismanlik - Digital Transformation Lead)",usd:0,eur:0,gbp:43200.0,totalUSD:58440.96,person:"Sevgul Erdogan",dueDate:"09/04/2026",days:63,note:""},{description:"SSI (Ise Alim Danismanlik - Senior Vice President, Sales Latin America)",usd:45832.0,eur:0,gbp:0,totalUSD:45832.0,person:"Bob Katsiouleris",dueDate:"13/02/2026",days:0,note:""},{description:"Kaban (Muhasebe ve Vergi Danismanligi)",usd:0,eur:0,gbp:33322.8,totalUSD:45079.08384000001,person:"M. Ali Erdogan",dueDate:"22/04/2026",days:21,note:""},{description:"Better Sustainability (Surdurulebilirlik Danismanligi)",usd:0,eur:0,gbp:30591.0,totalUSD:41383.5048,person:"Alan Knight",dueDate:"27/03/2026",days:18,note:""},{description:"Computershare (Yonetim Kurulu Raporlama ve Kayit Hizmetleri)",usd:0,eur:0,gbp:27043.2,totalUSD:36584.04096,person:"Jeremy Small",dueDate:"30/03/2026",days:21,note:""},{description:"The Design Portfolio (Web Sitesi Gelistirme)",usd:0,eur:0,gbp:21000.0,totalUSD:28408.8,person:"Chris Perry",dueDate:"01/04/2026",days:16,note:""},{description:"Robertsbridge (Sustainability High Ground Projesi Danismanligi)",usd:0,eur:0,gbp:19200.0,totalUSD:25973.76,person:"Alan Knight",dueDate:"10/03/2026",days:38,note:""},{description:"Cloudrock Partners (Insan Kaynaklari Bilgi Sistemleri Danismanligi)",usd:0,eur:0,gbp:16368.0,totalUSD:22142.630400000002,person:"Enver Citkin",dueDate:"08/04/2026",days:9,note:""},{description:"Deloitte (Turev Urun Degerleme Hizmeti)",usd:0,eur:0,gbp:16200.0,totalUSD:21915.36,person:"Ali Bekce",dueDate:"10/04/2026",days:0,note:""},{description:"Heads Resourcing (Aday Arama Ucreti)",usd:0,eur:0,gbp:12000.0,totalUSD:16233.6,person:"Sevgul Erdogan",dueDate:"07/04/2026",days:7,note:""},{description:"Jupiter Intelligence (Surdurulebilirlik Yazilimi)",usd:15000.0,eur:0,gbp:0,totalUSD:15000.0,person:"Alan Knight",dueDate:"11/03/2026",days:10,note:""},{description:"Morphsites (Website Tasarim ve Gelistirme)",usd:0,eur:0,gbp:9600.0,totalUSD:12986.88,person:"Anita Siddle",dueDate:"23/03/2026",days:37,note:""},{description:"Ponty Butchers (Sponsorluk)",usd:0,eur:0,gbp:8500.0,totalUSD:11498.8,person:"Josh Bassett",dueDate:"19/03/2026",days:0,note:""},{description:"Lumiere (Surdurulebilirlik Zirvesi Mekan Kirasi)",usd:0,eur:0,gbp:8460.0,totalUSD:11444.688,person:"Alan Knight",dueDate:"27/03/2026",days:29,note:""},{description:"Hamden (Ise Alim Hizmetleri - Sales Manager)",usd:0,eur:0,gbp:7599.6,totalUSD:10280.73888,person:"Sevgul Erdogan",dueDate:"23/04/2026",days:21,note:""},{description:"Climateprime (Sera Gazi Emisyonlari Danismanligi)",usd:0,eur:0,gbp:7500.0,totalUSD:10146.0,person:"Alan Knight",dueDate:"07/04/2026",days:0,note:""},{description:"Bonus Turizm (Ucak Biletleri - 450,788 TL)",usd:0,eur:0,gbp:7429.603774115569,totalUSD:10050.76798562354,person:"Jim Yow",dueDate:"14/04/2026",days:10,note:""},{description:"Diger",usd:7702.98,eur:0.0,gbp:25804.87,totalUSD:42611.80813599999,person:"",dueDate:"",days:0,note:""},{description:"Bugun Odemesi Yapilacak Borc",usd:1192210.8199999998,eur:65570.33,gbp:1824123.133774116,totalUSD:3737159.229274624,person:"",dueDate:"",days:0,note:""},{description:"Mart Sonuna Kadar Vadesi Gelecek Ek Borc",usd:1880000.0,eur:135455.17,gbp:859867.3900000001,totalUSD:3202862.5230370006,person:"",dueDate:"",days:0,note:""},{description:"Asesoria (ESG Data Destek ve Kurumsal Raporlama Hizmetleri)",usd:0,eur:0,gbp:6090.0,totalUSD:8238.552,person:"Anita Siddle",dueDate:"10/04/2026",days:7,note:""},{description:"Climateprime (Sera Gazi Emisyonlari Danismanligi)",usd:7702.98,eur:0,gbp:0,totalUSD:7702.98,person:"Alan Knight",dueDate:"23/04/2026",days:0,note:""},{description:"JDC (Yangin Guvenligi Onarim Isleri)",usd:0,eur:0,gbp:5153.4,totalUSD:6971.51952,person:"Christine Webb",dueDate:"21/04/2026",days:0,note:""},{description:"3 Hours West (Logo, Dokuman, Kitapcik Tasarim Hizmetleri)",usd:0,eur:0,gbp:3044.4,totalUSD:4118.46432,person:"Chris Perry",dueDate:"31/03/2026",days:17,note:""},{description:"ConnectWise (IT Uzaktan Cihaz Izleme ve Yonetim Hizmeti)",usd:0,eur:0,gbp:2755.2,totalUSD:3727.23456,person:"Enver Citkin",dueDate:"24/04/2026",days:0,note:""},{description:"JDC (Aydinlatma Isleri)",usd:0,eur:0,gbp:1974.0,totalUSD:2670.4272,person:"Christine Webb",dueDate:"21/04/2026",days:0,note:""},{description:"Jonathan Shopley (Surdurulebilirlik Danismanlik Hizmetleri)",usd:0,eur:0,gbp:1750.0,totalUSD:2367.4,person:"Alan Knight",dueDate:"16/04/2026",days:1,note:""},{description:"Croner (HR Platformu Yillik Uyelik)",usd:0,eur:0,gbp:1440.0,totalUSD:1948.032,person:"Sevgul Erdogan",dueDate:"14/04/2026",days:3,note:""},{description:"AXA Health (Saglik Sigortasi)",usd:0,eur:0,gbp:1334.62,totalUSD:1805.4739359999999,person:"Sevgul Erdogan",dueDate:"21/04/2026",days:0,note:""},{description:"Heads Resourcing (Aday Degerlendirme Ucreti)",usd:0,eur:0,gbp:1080.0,totalUSD:1461.0240000000001,person:"Sevgul Erdogan",dueDate:"13/03/2026",days:35,note:""},{description:"Outstanding Global (Liderlik ve Kariyer Gelisimi Danismanlik)",usd:0,eur:0,gbp:877.25,totalUSD:1186.7438,person:"Alan Knight",dueDate:"12/03/2026",days:36,note:""},{description:"Presentation Graphic Design (Sunum Donusturme Calismasi)",usd:0,eur:0,gbp:306.0,totalUSD:413.9568,person:"Chris Perry",dueDate:"23/03/2026",days:25,note:""},{description:"Serre Solutions (ResponsibleGlass Danismanligi)",usd:0,eur:0,gbp:25000.0,totalUSD:33820.0,person:"",dueDate:"12/03/2026",days:36,note:""},{description:"HN Communications (ResponsibleGlass Iletisim ve Etki Danismanligi)",usd:0,eur:0,gbp:21636.0,totalUSD:29269.1808,person:"",dueDate:"20/03/2026",days:28,note:""},{description:"ResponsibleGlass (Proje Yonetimi - Mart 2026)",usd:0,eur:0,gbp:16459.09,totalUSD:22265.856952000002,person:"",dueDate:"14/04/2026",days:3,note:""},{description:"HN Communications (ResponsibleGlass Iletisim - Mart)",usd:0,eur:0,gbp:10818.0,totalUSD:14634.5904,person:"",dueDate:"31/03/2026",days:17,note:""},{description:"HN Communications (ResponsibleGlass Iletisim - Nisan)",usd:0,eur:0,gbp:10818.0,totalUSD:14634.5904,person:"",dueDate:"01/05/2026",days:0,note:""},{description:"HN Communications (ResponsibleGlass CRM Destegi)",usd:0,eur:0,gbp:5700.0,totalUSD:7710.96,person:"",dueDate:"26/03/2026",days:22,note:""},{description:"OneWorldStandards Ltd (ResponsibleGlass Surdurulebilirlik Danismanligi)",usd:0,eur:0,gbp:5130.0,totalUSD:6939.864,person:"",dueDate:"03/04/2026",days:14,note:""},{description:"Greenall Glass Consulting (Danismanlik Hizmetleri)",usd:0,eur:0,gbp:2700.0,totalUSD:3652.56,person:"",dueDate:"31/03/2026",days:17,note:""},{description:"SJ Barron (ResponsibleGlass Kurumsal Tasarim Hizmeti)",usd:0,eur:0,gbp:965.0,totalUSD:1305.452,person:"",dueDate:"21/05/2026",days:0,note:""},{description:"Loaded Hype (ResponsibleGlass Website Calismasi)",usd:0,eur:0,gbp:900.0,totalUSD:1217.52,person:"",dueDate:"31/03/2026",days:17,note:""},{description:"SJ Barron (ResponsibleGlass Kurumsal Tasarim Hizmeti)",usd:0,eur:0,gbp:750.0,totalUSD:1014.6,person:"",dueDate:"19/03/2026",days:29,note:""},{description:"Berkeley Research Group (Rekabet Hukuku Danismanligi)",usd:0,eur:0,gbp:78000.0,totalUSD:105518.4,person:"",dueDate:"24/04/2026",days:0,note:""},{description:"PAYE",usd:0,eur:0,gbp:20638.0,totalUSD:27919.0864,person:"",dueDate:"20/04/2026",days:0,note:""},{description:"S&P Global (Yillik Veri Aboneligi)",usd:10639.68,eur:0,gbp:0,totalUSD:10639.68,person:"",dueDate:"29/04/2026",days:0,note:""},{description:"Emlak Vergisi Taksiti - Mayis",usd:0,eur:0,gbp:5799.0,totalUSD:7844.8872,person:"",dueDate:"01/05/2026",days:0,note:""},{description:"Kaban (Muhasebe ve Vergi Danismanligi)",usd:0,eur:0,gbp:4176.0,totalUSD:5649.2928,person:"",dueDate:"22/04/2026",days:0,note:""},{description:"City Central (Ofis Temizlik)",usd:0,eur:0,gbp:975.86,totalUSD:1320.143408,person:"",dueDate:"31/03/2026",days:17,note:""},{description:"FIPRA (ETS Hukuki Danismanlik)",usd:0,eur:72607.05,gbp:0,totalUSD:85567.40842500002,person:"M. Ali Erdogan",dueDate:"11/05/2026",days:0,note:""},{description:"Heads Resourcing (Ise Alim Danismanlik - Group Procurement Director)",usd:0,eur:0,gbp:59385.49,totalUSD:80336.69087199999,person:"Sevgul Erdogan",dueDate:"30/04/2026",days:0,note:""},{description:"ERM (ESG Danismanlik Hizmetleri)",usd:0,eur:0,gbp:59133.78,totalUSD:79996.177584,person:"Alan Knight",dueDate:"29/04/2026",days:0,note:""},{description:"Hoyng Rokh (Solvay Dava Ucreti)",usd:0,eur:62848.12,gbp:0,totalUSD:74066.50942000002,person:"M. Ali Erdogan",dueDate:"29/04/2026",days:0,note:""},{description:"Robertsbridge (Sustainability High Ground Projesi Danismanligi)",usd:0,eur:0,gbp:38400.0,totalUSD:51947.52,person:"Alan Knight",dueDate:"30/04/2026",days:0,note:""},{description:"DSS (WPS ve PSM Faz 2)",usd:50000.0,eur:0,gbp:0,totalUSD:50000.0,person:"A. Warren",dueDate:"30/04/2026",days:0,note:""},{description:"City of London (College Hill Mayis 2026 Emlak Vergileri)",usd:0,eur:0,gbp:33688.0,totalUSD:45573.1264,person:"Henry Penn",dueDate:"01/05/2026",days:0,note:""},{description:"Computershare (Yonetim Kurulu Raporlama ve Kayit Hizmetleri)",usd:0,eur:0,gbp:17266.8,totalUSD:23358.52704,person:"Jeremy Small",dueDate:"30/04/2026",days:0,note:""},{description:"GMT Economics (ETS Danismanlik Hizmetleri)",usd:0,eur:0,gbp:17160.0,totalUSD:23214.048,person:"M. Ali Erdogan",dueDate:"30/04/2026",days:0,note:""},{description:"Diger",usd:0.0,eur:0.0,gbp:14897.310000000001,totalUSD:20153.080968000002,person:"",dueDate:"",days:0,note:""},{description:"Climateprime (Sera Gazi Emisyonlari Danismanligi)",usd:0,eur:0,gbp:7500.0,totalUSD:10146.0,person:"Alan Knight",dueDate:"01/05/2026",days:0,note:""},{description:"Lyreco (Ofis Mutfak Masraflari)",usd:0,eur:0,gbp:2519.32,totalUSD:3408.136096,person:"Christine Webb",dueDate:"30/04/2026",days:0,note:""},{description:"Bupa (Saglik Sigortasi)",usd:0,eur:0,gbp:1534.0,totalUSD:2075.1952,person:"Sevgul Erdogan",dueDate:"01/05/2026",days:0,note:""},{description:"Addison Lee (Ulasim)",usd:0,eur:0,gbp:1291.38,totalUSD:1746.9788640000002,person:"",dueDate:"30/04/2026",days:0,note:""},{description:"Cala Construction (Ofis Tadilat)",usd:0,eur:0,gbp:1152.0,totalUSD:1558.4256,person:"Christine Webb",dueDate:"06/05/2026",days:0,note:""},{description:"Tag (Etkinlik Ucreti)",usd:0,eur:0,gbp:450.85,totalUSD:609.90988,person:"",dueDate:"01/05/2026",days:0,note:""},{description:"Big Ocean Data (Gemi Gozlem Programi Aylik Ucreti)",usd:0,eur:0,gbp:449.76,totalUSD:608.435328,person:"",dueDate:"29/04/2026",days:0,note:""}];


type InvEntry={id:string,direction:"in"|"out",sheet:"daily"|"aysonu",description:string,usd:number,eur:number,gbp:number,totalUSD:number,person:string,dueDate:string,note:string,importedAt:number};
type BankAccount={id:string,company:string,bankLabel:string,accountNo:string,country:string,cny:number,try_:number,usd:number,eur:number,gbp:number,totalUSD:number};
type LoanFacility={id:string,bank:string,usd:number,eur:number,totalUSD:number,pct:number};
type Rate={id:string,label:string,value:number};

// Parse DD/MM/YYYY → ms-since-epoch; 0 if invalid
function parseInvDate(s:string):number{if(!s)return 0;const p=s.split(/[\/\-\.]/);if(p.length===3&&p[2]?.length===4)return new Date(+p[2],+p[1]-1,+p[0]).getTime();return 0;}
// Days between today and due date: positive = overdue days, 0 = today, negative = days remaining
function overdueDays(dueDate:string):number{const d=parseInvDate(dueDate);if(!d)return 0;const today=new Date();today.setHours(0,0,0,0);return Math.round((today.getTime()-d)/86400000);}

function InvoicesTab(){
  const[invoices,setInvoices]=useState<InvEntry[]>([]);
  const[accounts,setAccounts]=useState<BankAccount[]>([]);
  const[loans,setLoans]=useState<LoanFacility[]>([]);
  const[rates,setRates]=useState<Rate[]>([]);
  const[activeSheet,setActiveSheet]=useState<"daily"|"aysonu"|"accounts">("daily");
  const[reportDate,setReportDate]=useState<string>(()=>{const d=new Date();return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`;});
  const[pasteMode,setPasteMode]=useState(false);
  const[pasteText,setPasteText]=useState("");
  const[isDraggingFile,setIsDraggingFile]=useState(false);
  const[dropStatus,setDropStatus]=useState<string>("");

  const INV_KEY="invoices_history";
  const PEOPLE_KEY="invoices_people_memory";
  const ACC_KEY="invoices_bank_accounts";
  const LOANS_KEY="invoices_loans";
  const RATES_KEY="invoices_rates";
  const RDATE_KEY="invoices_report_date";

  // === INITIAL LOAD ===
  useEffect(()=>{
    try{
      // Invoices: load from storage or seed from defaults
      const s=localStorage.getItem(INV_KEY);
      if(s){
        const all=JSON.parse(s);
        if(Array.isArray(all)&&all.length>0&&all[0].direction){setInvoices(all);}
        else{seedInvoices();}
      }else{seedInvoices();}
      // Report date
      const rd=localStorage.getItem(RDATE_KEY);if(rd)setReportDate(rd);
      // Bank accounts
      const a=localStorage.getItem(ACC_KEY);
      if(a){setAccounts(JSON.parse(a));}else{seedAccounts();}
      // Loans
      const l=localStorage.getItem(LOANS_KEY);
      if(l){setLoans(JSON.parse(l));}else{seedLoans();}
      // Rates
      const r=localStorage.getItem(RATES_KEY);
      if(r){setRates(JSON.parse(r));}else{seedRates();}
    }catch(e){console.error("InvoicesTab load:",e);}
  },[]);

  function seedInvoices(){
    const now=Date.now();
    // Classify incoming defaults: first 59 are "daily", rest are "aysonu"; none are direction=in (portal only exports outbound invoices, inbound are summary lines we skip)
    const seeded:InvEntry[]=INVOICE_HISTORY_DEFAULT.map((d,i)=>({
      id:"inv_seed_"+i,
      direction:"out",
      sheet:i<59?"daily":"aysonu",
      description:d.description,
      usd:d.usd,eur:d.eur,gbp:d.gbp,totalUSD:d.totalUSD,
      person:d.person,dueDate:d.dueDate,note:d.note,
      importedAt:now
    }));
    setInvoices(seeded);
    __undoEnabled=false;try{localStorage.setItem(INV_KEY,JSON.stringify(seeded));}catch{}__undoEnabled=true;
    const mem:Record<string,string>={};for(const i of INVOICE_HISTORY_DEFAULT){if(i.description&&i.person)mem[i.description.toLowerCase()]=i.person;}
    __undoEnabled=false;try{localStorage.setItem(PEOPLE_KEY,JSON.stringify(mem));}catch{}__undoEnabled=true;
  }

  function seedAccounts(){
    const def:BankAccount[]=[
      // WE SODA CARI HESAPLARI
      {id:"acc_1",company:"WE SODA CARI",bankLabel:"Barclays GBP Cari Hesabi",accountNo:"20-00-00-23128024",country:"UK",cny:0,try_:0,usd:0,eur:0,gbp:0,totalUSD:0},
      {id:"acc_2",company:"WE SODA CARI",bankLabel:"Barclays GBP Payroll Hesabi",accountNo:"",country:"UK",cny:0,try_:0,usd:0,eur:0,gbp:0,totalUSD:0},
      {id:"acc_3",company:"WE SODA CARI",bankLabel:"Barclays USD Cari Hesabi",accountNo:"20-00-00-47283944",country:"UK",cny:0,try_:0,usd:0,eur:0,gbp:0,totalUSD:0},
      {id:"acc_4",company:"WE SODA CARI",bankLabel:"Barclays EUR Cari Hesabi",accountNo:"20-00-00-84374811",country:"UK",cny:0,try_:0,usd:0,eur:0,gbp:0,totalUSD:0},
      {id:"acc_5",company:"WE SODA CARI",bankLabel:"Isbank GBP Cari Hesabi",accountNo:"405111-001",country:"UK",cny:0,try_:0,usd:0,eur:0,gbp:2529,totalUSD:3421.2312},
      // GOLDMAN SACHS DEPOSIT
      {id:"acc_6",company:"GOLDMAN SACHS",bankLabel:"USD",accountNo:"405129-001",country:"UK",cny:0,try_:0,usd:0,eur:0,gbp:0,totalUSD:0},
      {id:"acc_7",company:"GOLDMAN SACHS",bankLabel:"EUR",accountNo:"405129-002",country:"UK",cny:0,try_:0,usd:0,eur:0,gbp:0,totalUSD:0},
      {id:"acc_8",company:"GOLDMAN SACHS",bankLabel:"GBP",accountNo:"405129-003",country:"UK",cny:0,try_:0,usd:0,eur:0,gbp:0,totalUSD:0},
      // JP MORGAN DEPOSIT
      {id:"acc_9",company:"JP MORGAN",bankLabel:"USD",accountNo:"",country:"UK",cny:0,try_:0,usd:49981,eur:0,gbp:0,totalUSD:49981},
      {id:"acc_10",company:"JP MORGAN",bankLabel:"EUR",accountNo:"",country:"UK",cny:0,try_:0,usd:0,eur:2024,gbp:0,totalUSD:2385.28},
      {id:"acc_11",company:"JP MORGAN",bankLabel:"GBP",accountNo:"",country:"UK",cny:0,try_:0,usd:0,eur:0,gbp:17723,totalUSD:23975.68},
      // KEW SODA
      {id:"acc_12",company:"KEW SODA",bankLabel:"Isbank Cari Hesap (USD)",accountNo:"405129-001",country:"UK",cny:0,try_:0,usd:0,eur:0,gbp:0,totalUSD:0},
      {id:"acc_13",company:"KEW SODA",bankLabel:"Isbank Cari Hesap (EUR)",accountNo:"405129-002",country:"UK",cny:0,try_:0,usd:0,eur:0,gbp:0,totalUSD:0},
      // WSIH BONO
      {id:"acc_14",company:"WSIH BONO",bankLabel:"WSIH Bono - 5Y",accountNo:"",country:"UK",cny:0,try_:0,usd:980000000,eur:0,gbp:0,totalUSD:980000000},
      {id:"acc_15",company:"WSIH BONO",bankLabel:"WSIH Bono - 7Y",accountNo:"",country:"UK",cny:0,try_:0,usd:750000000,eur:0,gbp:0,totalUSD:750000000},
    ];
    setAccounts(def);__undoEnabled=false;try{localStorage.setItem(ACC_KEY,JSON.stringify(def));}catch{}__undoEnabled=true;
  }
  function seedLoans(){
    const def:LoanFacility[]=[
      {id:"ln_1",bank:"BNP",usd:21387096.77,eur:7548387.10,totalUSD:30282870.97,pct:0.2516},
      {id:"ln_2",bank:"JP MORGAN",usd:20564516.13,eur:7258064.52,totalUSD:29118145.16,pct:0.2419},
      {id:"ln_3",bank:"MUFG",usd:20564516.13,eur:7258064.52,totalUSD:29118145.16,pct:0.2419},
      {id:"ln_4",bank:"GOLDMAN SACHS",usd:13709677.42,eur:4838709.68,totalUSD:19412096.77,pct:0.1613},
      {id:"ln_5",bank:"ABC",usd:8774193.55,eur:3096774.19,totalUSD:12423741.94,pct:0.1032},
    ];
    setLoans(def);__undoEnabled=false;try{localStorage.setItem(LOANS_KEY,JSON.stringify(def));}catch{}__undoEnabled=true;
  }
  function seedRates(){
    const def:Rate[]=[
      {id:"r1",label:"GBP/EUR",value:1.14789987271956},
      {id:"r2",label:"GBP/USD",value:1.3528},
      {id:"r3",label:"GBP/TRY",value:60.67456808},
      {id:"r4",label:"EUR/USD",value:1.1785},
      {id:"r5",label:"USD/TRY",value:44.8511},
      {id:"r6",label:"USD/CNY",value:7.1925},
    ];
    setRates(def);__undoEnabled=false;try{localStorage.setItem(RATES_KEY,JSON.stringify(def));}catch{}__undoEnabled=true;
  }

  // === SAVE HELPERS ===
  const saveInv=(v:InvEntry[])=>{setInvoices(v);try{trackedSetItem(INV_KEY,JSON.stringify(v));}catch{}};
  const saveAcc=(v:BankAccount[])=>{setAccounts(v);try{trackedSetItem(ACC_KEY,JSON.stringify(v));}catch{}};
  const saveLoan=(v:LoanFacility[])=>{setLoans(v);try{trackedSetItem(LOANS_KEY,JSON.stringify(v));}catch{}};
  const saveRate=(v:Rate[])=>{setRates(v);try{trackedSetItem(RATES_KEY,JSON.stringify(v));}catch{}};
  const saveRDate=(v:string)=>{setReportDate(v);try{trackedSetItem(RDATE_KEY,v);}catch{}};

  // === INVOICE CRUD ===
  const addInv=(direction:"in"|"out",sheet:"daily"|"aysonu")=>{saveInv([...invoices,{id:"inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,7),direction,sheet,description:"",usd:0,eur:0,gbp:0,totalUSD:0,person:"",dueDate:"",note:"",importedAt:Date.now()}]);};
  const updInv=(id:string,f:keyof InvEntry,v:any)=>{saveInv(invoices.map(r=>r.id===id?{...r,[f]:v}:r));};
  const delInv=(id:string)=>{if(!confirm("Sil?"))return;saveInv(invoices.filter(r=>r.id!==id));};

  // === DRAG-DROP EXCEL IMPORT ===
  const ensureSheetJS=async()=>{if((window as any).XLSX)return (window as any).XLSX;return new Promise<any>((resolve,reject)=>{const s=document.createElement("script");s.src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";s.onload=()=>resolve((window as any).XLSX);s.onerror=()=>reject(new Error("XLSX"));document.head.appendChild(s);});};
  const handleFile=async(file:File)=>{
    try{
      setDropStatus("İşleniyor: "+file.name);
      const name=file.name.toLowerCase();
      if(!name.endsWith(".xlsx")&&!name.endsWith(".xls")){setDropStatus("Sadece .xlsx dosyası kabul edilir.");setTimeout(()=>setDropStatus(""),3500);return;}
      const XLSX=await ensureSheetJS();
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:"array",cellDates:true});
      let mem:Record<string,string>={};
      try{mem=JSON.parse(localStorage.getItem(PEOPLE_KEY)||"{}");}catch{}
      const allParsed:InvEntry[]=[];

      const toDateStr=(v:any)=>{
        if(!v)return "";
        if(v instanceof Date)return `${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;
        if(typeof v==="number"&&v>25569){const d=new Date((v-25569)*86400*1000);return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;}
        return String(v).trim();
      };
      const num=(v:any)=>{if(typeof v==="number")return v;const s=String(v||"").replace(/[,\s]/g,"");const n=parseFloat(s);return isNaN(n)?0:n;};

      // Process Daily 1st Page + Ay Sonu (invoices), but skip Bank Accounts page
      for(const sn of wb.SheetNames){
        const isAySonu=/ay\s*sonu/i.test(sn);
        const isDaily=/daily|1st|page/i.test(sn);
        const isAccounts=/bank|account|loan/i.test(sn);
        if(isAccounts){
          // Extract bank accounts, loans, rates
          importAccountsFromSheet(wb.Sheets[sn],XLSX);
          continue;
        }
        if(!isDaily&&!isAySonu)continue;
        const sheetTag:"daily"|"aysonu"=isAySonu?"aysonu":"daily";
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:""}) as any[][];

        for(const row of rows){
          if(!row||row.length<5)continue;
          // INFLOWS: cols B-F (1-5) + person(N=13) + date(P=15) — but from sample it seems inflows are in first 2-3 summary rows, the rest are all outflows in I-T
          // OUTFLOWS: cols I-T (8-19): I=desc(8), J=usd(9), K=eur(10), L=gbp(11), M=total(12), N=person(13), P=date(15), T=note(19)
          // Try inflows first (col B)
          const bDesc=String(row[1]||"").trim();
          if(bDesc&&bDesc.length>3&&!/GIRIS|USD|EUR|GBP|TOPLAM/i.test(bDesc)){
            const bUSD=num(row[2]),bEUR=num(row[3]),bGBP=num(row[4]),bTot=num(row[5]);
            if(bTot>0||bUSD>0||bEUR>0||bGBP>0){
              allParsed.push({id:"inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,8),direction:"in",sheet:sheetTag,description:bDesc,usd:bUSD,eur:bEUR,gbp:bGBP,totalUSD:bTot||Math.max(bUSD,bEUR,bGBP),person:"",dueDate:toDateStr(row[15]),note:"",importedAt:Date.now()});
            }
          }
          // Outflows (I column)
          const iDesc=String(row[8]||"").trim();
          if(!iDesc||iDesc.length<3)continue;
          const du=iDesc.toUpperCase();
          if(du==="USD"||du==="EUR"||du==="GBP"||/CIKIS|TOPLAM/i.test(du))continue;
          const iUSD=num(row[9]),iEUR=num(row[10]),iGBP=num(row[11]),iTot=num(row[12]);
          if(iTot===0&&iUSD===0&&iEUR===0&&iGBP===0)continue;
          const iPerson=String(row[13]||"").trim()||mem[iDesc.toLowerCase()]||"";
          allParsed.push({id:"inv_"+Date.now()+"_"+Math.random().toString(36).slice(2,8),direction:"out",sheet:sheetTag,description:iDesc,usd:iUSD,eur:iEUR,gbp:iGBP,totalUSD:iTot||Math.max(iUSD,iEUR,iGBP),person:iPerson,dueDate:toDateStr(row[15]),note:String(row[19]||"").trim(),importedAt:Date.now()});
        }
      }

      if(allParsed.length===0){setDropStatus("Fatura bulunamadı. Excel sayfa isimleri 'Daily 1st Page' ve 'Ay Sonu' olmalı.");setTimeout(()=>setDropStatus(""),4500);return;}
      // Merge
      const merged:InvEntry[]=[...invoices];
      let added=0,updated=0;
      for(const p of allParsed){
        const idx=merged.findIndex(e=>e.description===p.description&&e.dueDate===p.dueDate&&e.direction===p.direction);
        if(idx>=0){merged[idx]={...merged[idx],...p,id:merged[idx].id,note:p.note||merged[idx].note,person:p.person||merged[idx].person};updated++;}
        else{merged.push(p);added++;}
      }
      for(const p of allParsed){if(p.description&&p.person)mem[p.description.toLowerCase()]=p.person;}
      try{trackedSetItem(PEOPLE_KEY,JSON.stringify(mem));}catch{}
      saveInv(merged);
      // Update report date from filename if possible, e.g. Daily_Reporting_17Apr2026
      const dateMatch=file.name.match(/(\d{1,2})\s*([A-Za-z]{3})\s*(\d{4})/);
      if(dateMatch){
        const months:Record<string,string>={"jan":"01","feb":"02","mar":"03","apr":"04","may":"05","jun":"06","jul":"07","aug":"08","sep":"09","oct":"10","nov":"11","dec":"12"};
        const mm=months[dateMatch[2].toLowerCase().slice(0,3)]||"01";
        saveRDate(`${dateMatch[1].padStart(2,"0")}/${mm}/${dateMatch[3]}`);
      }
      setDropStatus(`✓ ${added} yeni fatura, ${updated} güncelleme`);
      setTimeout(()=>setDropStatus(""),4000);
    }catch(err){setDropStatus("Hata: "+(err instanceof Error?err.message:""));setTimeout(()=>setDropStatus(""),4000);}
  };

  // Import Bank Accounts / Loans / Rates from the third sheet
  const importAccountsFromSheet=(ws:any,XLSX:any)=>{
    try{
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}) as any[][];
      const num=(v:any)=>{if(typeof v==="number")return v;const s=String(v||"").replace(/[,\s]/g,"");const n=parseFloat(s);return isNaN(n)?0:n;};
      // Loans block: look for rows with bank name in col B and amounts in G,H,J
      const newLoans:LoanFacility[]=[];
      const newRates:Rate[]=[];
      let inLoans=false;
      for(let r=0;r<rows.length;r++){
        const row=rows[r];if(!row)continue;
        const b=String(row[1]||"").trim();
        if(/ROTATIF|KREDI/i.test(b))inLoans=true;
        if(/KUR/i.test(b))inLoans=false;
        if(inLoans&&b&&!/ROTATIF|KREDI|USD|TOPLAM|WE SODA/i.test(b)&&/^[A-Z]/.test(b)){
          const usd=num(row[6]),eur=num(row[7]),total=num(row[9]),pct=num(row[11]);
          if(usd>0||total>0)newLoans.push({id:"ln_"+Date.now()+"_"+r,bank:b,usd,eur,totalUSD:total,pct});
        }
        // Rates: GBP/EUR, GBP/USD, EUR/USD, USD/TRY etc
        if(/^[A-Z]{3}\/[A-Z]{3}$/.test(b)){
          const val=num(row[3]);if(val>0)newRates.push({id:"rt_"+Date.now()+"_"+r,label:b,value:val});
          const b2=String(row[7]||"").trim();const val2=num(row[9]);
          if(/^[A-Z]{3}\/[A-Z]{3}$/.test(b2)&&val2>0)newRates.push({id:"rt_"+Date.now()+"_"+r+"_b",label:b2,value:val2});
        }
      }
      if(newLoans.length>0)saveLoan(newLoans);
      if(newRates.length>0)saveRate(newRates);
    }catch(e){console.error("importAccounts:",e);}
  };

  const onDrop=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(false);const files=Array.from(e.dataTransfer.files);if(files.length>0)handleFile(files[0]);};
  const onDragOver=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(true);};
  const onDragLeave=(e:React.DragEvent)=>{if(e.currentTarget===e.target)setIsDraggingFile(false);};

  // === LIVE COMPUTATION ===
  const live=useMemo(()=>invoices.map(inv=>({...inv,overdueDays:overdueDays(inv.dueDate)})),[invoices]);
  const people=[...new Set(live.filter(r=>r.person).map(r=>r.person))].sort();
  const critInvoices=live.filter(r=>r.direction==="out"&&r.overdueDays>30);

  // === RENDER HELPERS ===
  // Excel column widths mimic the original
  // Thin-bordered, soft-colored styles. Numerals use tabular-nums for readability.
  const numStyle={fontFamily:"'Calibri','Segoe UI','Arial',system-ui,sans-serif",fontVariantNumeric:"tabular-nums" as const,fontSize:"13px",letterSpacing:"0.01em"};
  const numStyleBold={...numStyle,fontSize:"14px",fontWeight:600};
  const cellIn="px-2 py-1 border border-slate-200 bg-white";
  const hdrIn="px-2 py-1.5 border-b border-slate-200 bg-slate-50 text-[10px] font-semibold text-slate-600 text-center tracking-wide uppercase";
  const totalRowCls="border-t border-slate-300 bg-slate-50 font-semibold";

  // === FUNDING NEEDS COMPUTATION ===
  // Mevcut Mevduat = sum of Total USD from Bank Accounts (excluding WSIH Bono which is long-term)
  const mevcutMevduat=useMemo(()=>{
    return accounts
      .filter(a=>!/BONO/i.test(a.company))  // exclude WSIH Bono (long-term investment, not liquid deposit)
      .reduce((s,a)=>s+a.totalUSD,0);
  },[accounts]);

  // Outflow totals per sheet (what needs to be paid out)
  const dailyOutTotal=useMemo(()=>live.filter(r=>r.direction==="out"&&r.sheet==="daily").reduce((s,r)=>s+r.totalUSD,0),[live]);
  const monthlyOutTotal=useMemo(()=>live.filter(r=>r.direction==="out"&&r.sheet==="aysonu").reduce((s,r)=>s+r.totalUSD,0),[live]);

  // Funding logic: assume Mevcut Mevduat is drawn down first (across both buckets in order: daily first, then monthly)
  // For each bucket, compute:
  //   Mevduattan Kullanim = min(remaining deposit, bucket outflow)
  //   Fonlama Ihtiyaci = bucket outflow - Mevduattan Kullanim
  const funding=useMemo(()=>{
    let remaining=mevcutMevduat;
    // Daily first (more urgent)
    const dailyFromDeposit=Math.min(remaining,dailyOutTotal);
    const dailyFundingNeed=Math.max(0,dailyOutTotal-dailyFromDeposit);
    remaining=Math.max(0,remaining-dailyFromDeposit);
    // Then monthly
    const monthlyFromDeposit=Math.min(remaining,monthlyOutTotal);
    const monthlyFundingNeed=Math.max(0,monthlyOutTotal-monthlyFromDeposit);
    return{
      dailyOut:dailyOutTotal,dailyFromDeposit,dailyFundingNeed,
      monthlyOut:monthlyOutTotal,monthlyFromDeposit,monthlyFundingNeed,
      totalOut:dailyOutTotal+monthlyOutTotal,
      totalFromDeposit:dailyFromDeposit+monthlyFromDeposit,
      totalFundingNeed:dailyFundingNeed+monthlyFundingNeed,
    };
  },[mevcutMevduat,dailyOutTotal,monthlyOutTotal]);

  // Funding table — rendered at the top of daily/aysonu pages
  const renderFundingTable=(activeSheet:"daily"|"aysonu")=>{
    const f=funding;
    const isDaily=activeSheet==="daily";
    const numCal={fontFamily:"'Calibri','Segoe UI','Arial',system-ui,sans-serif",fontVariantNumeric:"tabular-nums" as const};
    return<div className="mb-5 bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
      <div className="flex items-center gap-2.5 px-5 py-3 border-b border-slate-100">
        <div className="w-1 h-6 bg-amber-400 rounded-full"></div>
        <span className="text-[13px] font-semibold text-slate-800">Fonlama İhtiyacı</span>
        <span className="text-[10px] text-slate-400">— vadesi gelen faturaların mevduat sonrası açık tutarı</span>
      </div>
      <table className="border-collapse w-full">
        <thead>
          <tr className="bg-slate-50/50">
            <th className="px-5 py-2.5 border-b border-slate-100 text-[10px] font-semibold text-slate-500 text-left tracking-wider uppercase" style={{minWidth:240}}>Kategori</th>
            <th className="px-5 py-2.5 border-b border-slate-100 text-[10px] font-semibold text-slate-500 text-right tracking-wider uppercase" style={{minWidth:160}}>Toplam Ödeme</th>
            <th className="px-5 py-2.5 border-b border-slate-100 text-[10px] font-semibold text-emerald-600 text-right tracking-wider uppercase" style={{minWidth:180}}>Mevduattan Kullanım</th>
            <th className="px-5 py-2.5 border-b border-slate-100 text-[10px] font-semibold text-amber-600 text-right tracking-wider uppercase" style={{minWidth:170}}>Fonlama İhtiyacı</th>
          </tr>
        </thead>
        <tbody>
          <tr className={"group "+(isDaily?"bg-sky-50/30":"hover:bg-slate-50/50")}>
            <td className="px-5 py-3 border-b border-slate-100 text-left">
              <div className="flex items-center gap-2">
                <span className="text-[12px] text-slate-800 font-medium">Vadesi Gelenler</span>
                {isDaily&&<span className="text-[9px] text-sky-600 font-medium px-1.5 py-0.5 bg-sky-100/70 rounded">aktif</span>}
              </div>
              <span className="text-[10px] text-slate-400">1 hafta içi</span>
            </td>
            <td className="px-5 py-3 border-b border-slate-100 text-right text-slate-700" style={{...numCal,fontSize:"14px"}}>{fmt(f.dailyOut)}</td>
            <td className="px-5 py-3 border-b border-slate-100 text-right text-emerald-600" style={{...numCal,fontSize:"14px"}}>{fmt(f.dailyFromDeposit)}</td>
            <td className="px-5 py-3 border-b border-slate-100 text-right font-semibold text-amber-700" style={{...numCal,fontSize:"14px"}}>{fmt(f.dailyFundingNeed)}</td>
          </tr>
          <tr className={"group "+(!isDaily?"bg-sky-50/30":"hover:bg-slate-50/50")}>
            <td className="px-5 py-3 border-b border-slate-100 text-left">
              <div className="flex items-center gap-2">
                <span className="text-[12px] text-slate-800 font-medium">Vadesi Yaklaşanlar</span>
                {!isDaily&&<span className="text-[9px] text-sky-600 font-medium px-1.5 py-0.5 bg-sky-100/70 rounded">aktif</span>}
              </div>
              <span className="text-[10px] text-slate-400">1 hafta - 1 ay arası</span>
            </td>
            <td className="px-5 py-3 border-b border-slate-100 text-right text-slate-700" style={{...numCal,fontSize:"14px"}}>{fmt(f.monthlyOut)}</td>
            <td className="px-5 py-3 border-b border-slate-100 text-right text-emerald-600" style={{...numCal,fontSize:"14px"}}>{fmt(f.monthlyFromDeposit)}</td>
            <td className="px-5 py-3 border-b border-slate-100 text-right font-semibold text-amber-700" style={{...numCal,fontSize:"14px"}}>{fmt(f.monthlyFundingNeed)}</td>
          </tr>
          <tr className="bg-slate-50/70">
            <td className="px-5 py-3.5 text-left text-[12px] font-semibold text-slate-800">Toplam</td>
            <td className="px-5 py-3.5 text-right font-semibold text-slate-800" style={{...numCal,fontSize:"15px"}}>{fmt(f.totalOut)}</td>
            <td className="px-5 py-3.5 text-right font-semibold text-emerald-700" style={{...numCal,fontSize:"15px"}}>{fmt(f.totalFromDeposit)}</td>
            <td className="px-5 py-3.5 text-right font-semibold text-amber-800" style={{...numCal,fontSize:"16px"}}>{fmt(f.totalFundingNeed)}</td>
          </tr>
        </tbody>
      </table>
      <div className="px-5 py-2.5 bg-slate-50/30 border-t border-slate-100 text-[10px] text-slate-500 flex items-center gap-1.5">
        <svg className="w-3 h-3 text-slate-400" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
        <span>Mevcut Mevduat: </span>
        <span className="font-semibold text-emerald-700" style={numCal}>{fmt(mevcutMevduat)} <span className="text-[9px] text-slate-400 font-normal">USD</span></span>
        <span className="text-slate-400">— Hesaplar &amp; Krediler sekmesinden (WSIH Bono hariç)</span>
      </div>
    </div>;
  };

  const renderSheet=(sheet:"daily"|"aysonu")=>{
    const outflows=live.filter(r=>r.sheet===sheet&&r.direction==="out").sort((a,b)=>b.totalUSD-a.totalUSD);
    const sumUSD=(arr:InvEntry[])=>arr.reduce((s,r)=>s+r.usd,0);
    const sumEUR=(arr:InvEntry[])=>arr.reduce((s,r)=>s+r.eur,0);
    const sumGBP=(arr:InvEntry[])=>arr.reduce((s,r)=>s+r.gbp,0);
    const sumTotal=(arr:InvEntry[])=>arr.reduce((s,r)=>s+r.totalUSD,0);
    const title=sheet==="daily"?"Günlük Özet Rapor":"Ay Sonu Özet Rapor";

    const fmtAmt=(n:number)=>n===0?"":fmt(n);
    const fmtDays=(d:number)=>{if(d<=0)return "";if(d>30)return<span className="text-red-700 font-bold">{d}</span>;if(d>7)return<span className="text-orange-600 font-semibold">{d}</span>;return<span className="text-amber-600">{d}</span>;};

    return<div className="overflow-x-auto">
      {/* FUNDING NEEDS TABLE — on top of every sheet */}
      {renderFundingTable(sheet)}

      <div className="relative bg-white rounded-xl shadow-sm border border-slate-200 inline-block min-w-full overflow-hidden">
        {/* Report banner */}
        <div className="flex items-center justify-between px-5 py-3 bg-white border-b border-slate-100">
          <div className="flex items-center gap-2.5">
            <div className="w-1 h-8 bg-teal-400 rounded-full"></div>
            <div>
              <div className="text-[10px] text-slate-400 uppercase tracking-wider font-medium">We Soda Ltd</div>
              <div className="text-[13px] font-semibold text-slate-800 leading-tight">{title}</div>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <span className="text-[10px] text-slate-400 uppercase tracking-wider">Rapor Tarihi</span>
            <input value={reportDate} onChange={e=>saveRDate(e.target.value)} className="text-[12px] font-medium px-2.5 py-1 rounded-md border border-slate-200 outline-none text-slate-700 focus:border-teal-400 focus:ring-1 focus:ring-teal-100 bg-white w-[110px] text-center" style={numFont} placeholder="DD/MM/YYYY"/>
          </div>
        </div>

        <table className="border-collapse w-full" style={{minWidth:1100,fontFamily:"'Inter', system-ui, sans-serif"}}>
          <thead>
            {/* Section header — soft rose */}
            <tr>
              <th colSpan={9} className="px-3 py-2 bg-rose-50/60 border-b border-slate-200 text-[11px] font-semibold text-rose-700 text-left tracking-wide uppercase">Çıkışlar</th>
            </tr>
            {/* Column headers */}
            <tr>
              <th className={hdrIn+" text-left"} style={{minWidth:320}}>Açıklama</th>
              <th className={hdrIn} style={{minWidth:100}}>USD</th>
              <th className={hdrIn} style={{minWidth:100}}>EUR</th>
              <th className={hdrIn} style={{minWidth:100}}>GBP</th>
              <th className={hdrIn+" bg-rose-50/40"} style={{minWidth:120}}>Toplam USD</th>
              <th className={hdrIn+" text-left"} style={{minWidth:130}}>Sorumlu</th>
              <th className={hdrIn} style={{minWidth:105}}>Son Ödeme</th>
              <th className={hdrIn} style={{minWidth:90}}>Gecikme (gün)</th>
              <th className="w-8"></th>
            </tr>
          </thead>
          <tbody>
            {outflows.length===0?<tr><td colSpan={9} className="py-12 text-center text-slate-400 text-[11px] border-b border-slate-200">Fatura yok. Excel raporunu yukarıdan yükleyin veya "+ Çıkış Ekle" butonuna tıklayın.</td></tr>:outflows.map((outRow,i)=>{
              const outOverdue=outRow.overdueDays>30;
              return<tr key={outRow.id} className={"hover:bg-slate-50/60 "+(outOverdue?"bg-rose-50/40":"")}>
                <td className={cellIn}><input className="w-full text-[11px] outline-none bg-transparent text-slate-800" value={outRow.description} onChange={e=>updInv(outRow.id,"description",e.target.value)}/></td>
                <td className={cellIn+" text-right"}><input className="w-full text-right outline-none bg-transparent text-emerald-700" style={numStyle} value={fmtAmt(outRow.usd)} onChange={e=>updInv(outRow.id,"usd",parseFloat(stripC(e.target.value))||0)}/></td>
                <td className={cellIn+" text-right"}><input className="w-full text-right outline-none bg-transparent text-sky-700" style={numStyle} value={fmtAmt(outRow.eur)} onChange={e=>updInv(outRow.id,"eur",parseFloat(stripC(e.target.value))||0)}/></td>
                <td className={cellIn+" text-right"}><input className="w-full text-right outline-none bg-transparent text-violet-700" style={numStyle} value={fmtAmt(outRow.gbp)} onChange={e=>updInv(outRow.id,"gbp",parseFloat(stripC(e.target.value))||0)}/></td>
                <td className={cellIn+" text-right bg-rose-50/20"}><input className="w-full text-right outline-none bg-transparent text-slate-800" style={numStyleBold} value={fmtAmt(outRow.totalUSD)} onChange={e=>updInv(outRow.id,"totalUSD",parseFloat(stripC(e.target.value))||0)}/></td>
                <td className={cellIn}><input className="w-full text-[11px] outline-none bg-transparent text-slate-700" value={outRow.person} onChange={e=>updInv(outRow.id,"person",e.target.value)} list="inv-people"/></td>
                <td className={cellIn+" text-center"}><DateCell value={outRow.dueDate} onChange={v=>updInv(outRow.id,"dueDate",v)}/></td>
                <td className={cellIn+" text-center"} style={numStyle}>{fmtDays(outRow.overdueDays)}</td>
                <td className="px-1 border border-slate-200 bg-white text-center"><button onClick={()=>delInv(outRow.id)} className="text-slate-300 hover:text-rose-500 text-[14px] leading-none px-1" title="Sil">×</button></td>
              </tr>;
            })}
            {/* TOPLAM row */}
            <tr className={totalRowCls}>
              <td className="px-3 py-2 border-t border-slate-300 text-right text-[11px] text-slate-700 uppercase tracking-wide">Toplam Çıkış</td>
              <td className="px-2 py-2 border-t border-slate-300 text-right text-slate-700" style={numStyleBold}>{fmt(sumUSD(outflows))}</td>
              <td className="px-2 py-2 border-t border-slate-300 text-right text-slate-700" style={numStyleBold}>{fmt(sumEUR(outflows))}</td>
              <td className="px-2 py-2 border-t border-slate-300 text-right text-slate-700" style={numStyleBold}>{fmt(sumGBP(outflows))}</td>
              <td className="px-2 py-2 border-t border-slate-300 text-right bg-rose-100/50 text-rose-800" style={{...numStyleBold,fontSize:"14px"}}>{fmt(sumTotal(outflows))}</td>
              <td colSpan={4} className="px-3 py-2 border-t border-slate-300 text-right text-[10px] text-slate-500">{outflows.length} fatura</td>
            </tr>
          </tbody>
        </table>

        {/* Add button */}
        <div className="flex gap-2 px-3 py-2 bg-slate-50/50 border-t border-slate-200">
          <button onClick={()=>addInv("out",sheet)} className="px-3 py-1 bg-rose-500 hover:bg-rose-600 text-white rounded-md text-[11px] font-medium transition-colors">+ Çıkış Ekle</button>
          <span className="ml-auto text-[10px] text-slate-500 self-center">{outflows.length} fatura • Sıralama: büyükten küçüğe (USD)</span>
        </div>
      </div>
    </div>;
  };

  const renderAccounts=()=>{
    const sumUSD=accounts.reduce((s,a)=>s+a.totalUSD,0);
    const companies=[...new Set(accounts.map(a=>a.company))];
    return<div className="space-y-4" style={{fontFamily:"'Inter', system-ui, sans-serif"}}>
      {/* Banner */}
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
        <div className="flex items-center justify-between px-5 py-3 bg-white border-b border-slate-100">
          <div className="flex items-center gap-2.5">
            <div className="w-1 h-8 bg-violet-400 rounded-full"></div>
            <div>
              <div className="text-[10px] text-slate-400 uppercase tracking-wider font-medium">We Soda Ltd</div>
              <div className="text-[13px] font-semibold text-slate-800 leading-tight">Hesaplar, Krediler &amp; Kurlar</div>
            </div>
          </div>
          <div className="flex items-center gap-6">
            <div className="text-right">
              <div className="text-[10px] text-slate-400 uppercase tracking-wider font-medium">Toplam Mevduat</div>
              <div className="text-emerald-700 font-semibold" style={{...numStyle,fontSize:"14px"}}>{fmt(sumUSD)} <span className="text-[9px] text-slate-400 font-normal">USD</span></div>
            </div>
            <div className="flex items-center gap-2">
              <span className="text-[10px] text-slate-400 uppercase tracking-wider">Tarih</span>
              <input value={reportDate} onChange={e=>saveRDate(e.target.value)} className="text-[12px] font-medium px-2.5 py-1 rounded-md border border-slate-200 outline-none text-slate-700 focus:border-violet-400 focus:ring-1 focus:ring-violet-100 bg-white w-[110px] text-center" style={numStyle}/>
            </div>
          </div>
        </div>

        {/* Accounts table */}
        {companies.map(comp=>{
          const compAccs=accounts.filter(a=>a.company===comp);
          return<div key={comp} className="border-b border-slate-200 last:border-b-0">
            <div className="px-4 py-1.5 bg-slate-50 border-b border-slate-200 text-[11px] font-semibold text-slate-700 tracking-wide">{comp}</div>
            <table className="border-collapse w-full">
              <thead>
                <tr>
                  <th className={hdrIn+" text-left"} style={{minWidth:200}}>Hesap</th>
                  <th className={hdrIn+" text-left"} style={{minWidth:130}}>Hesap No</th>
                  <th className={hdrIn} style={{minWidth:60}}>Ülke</th>
                  <th className={hdrIn} style={{minWidth:90}}>CNY</th>
                  <th className={hdrIn} style={{minWidth:90}}>TRY</th>
                  <th className={hdrIn} style={{minWidth:90}}>USD</th>
                  <th className={hdrIn} style={{minWidth:90}}>EUR</th>
                  <th className={hdrIn} style={{minWidth:90}}>GBP</th>
                  <th className={hdrIn+" bg-amber-50/40"} style={{minWidth:120}}>Toplam USD</th>
                  <th className="w-6"></th>
                </tr>
              </thead>
              <tbody>{compAccs.map((a,i)=><tr key={a.id} className="hover:bg-slate-50/60">
                <td className={cellIn}><input className="w-full text-[11px] outline-none text-slate-800" value={a.bankLabel} onChange={e=>saveAcc(accounts.map(x=>x.id===a.id?{...x,bankLabel:e.target.value}:x))}/></td>
                <td className={cellIn}><input className="w-full outline-none text-slate-600" style={numStyle} value={a.accountNo} onChange={e=>saveAcc(accounts.map(x=>x.id===a.id?{...x,accountNo:e.target.value}:x))}/></td>
                <td className={cellIn+" text-center"}><input className="w-full text-[11px] text-center outline-none text-slate-600" value={a.country} onChange={e=>saveAcc(accounts.map(x=>x.id===a.id?{...x,country:e.target.value}:x))}/></td>
                <td className={cellIn+" text-right"}><input className="w-full text-right outline-none text-slate-700" style={numStyle} value={a.cny?addC(a.cny.toString()):""} onChange={e=>saveAcc(accounts.map(x=>x.id===a.id?{...x,cny:parseFloat(stripC(e.target.value))||0}:x))}/></td>
                <td className={cellIn+" text-right"}><input className="w-full text-right outline-none text-slate-700" style={numStyle} value={a.try_?addC(a.try_.toString()):""} onChange={e=>saveAcc(accounts.map(x=>x.id===a.id?{...x,try_:parseFloat(stripC(e.target.value))||0}:x))}/></td>
                <td className={cellIn+" text-right"}><input className="w-full text-right outline-none text-emerald-700" style={numStyle} value={a.usd?addC(a.usd.toString()):""} onChange={e=>saveAcc(accounts.map(x=>x.id===a.id?{...x,usd:parseFloat(stripC(e.target.value))||0}:x))}/></td>
                <td className={cellIn+" text-right"}><input className="w-full text-right outline-none text-sky-700" style={numStyle} value={a.eur?addC(a.eur.toString()):""} onChange={e=>saveAcc(accounts.map(x=>x.id===a.id?{...x,eur:parseFloat(stripC(e.target.value))||0}:x))}/></td>
                <td className={cellIn+" text-right"}><input className="w-full text-right outline-none text-violet-700" style={numStyle} value={a.gbp?addC(a.gbp.toString()):""} onChange={e=>saveAcc(accounts.map(x=>x.id===a.id?{...x,gbp:parseFloat(stripC(e.target.value))||0}:x))}/></td>
                <td className={cellIn+" text-right bg-amber-50/20"}><input className="w-full text-right outline-none bg-transparent text-slate-800" style={numStyleBold} value={a.totalUSD?addC(a.totalUSD.toString()):""} onChange={e=>saveAcc(accounts.map(x=>x.id===a.id?{...x,totalUSD:parseFloat(stripC(e.target.value))||0}:x))}/></td>
                <td className="px-1"><button onClick={()=>{if(confirm("Sil?"))saveAcc(accounts.filter(x=>x.id!==a.id));}} className="text-slate-300 hover:text-rose-500 text-[13px] leading-none px-1">×</button></td>
              </tr>)}
              <tr className="bg-slate-50 border-t border-slate-200">
                <td colSpan={8} className="px-3 py-2 text-right text-[11px] font-semibold text-slate-600 uppercase tracking-wide">Toplam — {comp}</td>
                <td className="px-2 py-2 text-right bg-amber-50/40 text-slate-800" style={numStyleBold}>{fmt(compAccs.reduce((s,a)=>s+a.totalUSD,0))}</td>
                <td></td>
              </tr>
              </tbody>
            </table>
          </div>;
        })}
        <div className="px-3 py-2 bg-slate-50/50 border-t border-slate-200">
          <button onClick={()=>saveAcc([...accounts,{id:"acc_"+Date.now(),company:"WE SODA CARI",bankLabel:"",accountNo:"",country:"UK",cny:0,try_:0,usd:0,eur:0,gbp:0,totalUSD:0}])} className="px-3 py-1 bg-sky-500 hover:bg-sky-600 text-white rounded-md text-[11px] font-medium transition-colors">+ Hesap Ekle</button>
        </div>
      </div>

      {/* Loans */}
      <div className="bg-white rounded-lg shadow-sm border border-slate-200 overflow-hidden">
        <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 text-[11px] font-semibold text-slate-700 tracking-wide uppercase">Rotatif Kredi</div>
        <table className="border-collapse w-full">
          <thead><tr>
            <th className={hdrIn+" text-left"} style={{minWidth:150}}>Banka</th>
            <th className={hdrIn} style={{minWidth:120}}>USD</th>
            <th className={hdrIn} style={{minWidth:120}}>EUR</th>
            <th className={hdrIn+" bg-amber-50/40"} style={{minWidth:130}}>Toplam USD</th>
            <th className={hdrIn} style={{minWidth:70}}>Pay (%)</th>
            <th className="w-6"></th>
          </tr></thead>
          <tbody>{loans.map(l=><tr key={l.id} className="hover:bg-slate-50/60">
            <td className={cellIn}><input className="w-full text-[11px] font-medium outline-none text-slate-800" value={l.bank} onChange={e=>saveLoan(loans.map(x=>x.id===l.id?{...x,bank:e.target.value}:x))}/></td>
            <td className={cellIn+" text-right"}><input className="w-full text-right outline-none text-emerald-700" style={numStyle} value={l.usd?addC(l.usd.toString()):""} onChange={e=>saveLoan(loans.map(x=>x.id===l.id?{...x,usd:parseFloat(stripC(e.target.value))||0}:x))}/></td>
            <td className={cellIn+" text-right"}><input className="w-full text-right outline-none text-sky-700" style={numStyle} value={l.eur?addC(l.eur.toString()):""} onChange={e=>saveLoan(loans.map(x=>x.id===l.id?{...x,eur:parseFloat(stripC(e.target.value))||0}:x))}/></td>
            <td className={cellIn+" text-right bg-amber-50/20"}><input className="w-full text-right outline-none bg-transparent text-slate-800" style={numStyleBold} value={l.totalUSD?addC(l.totalUSD.toString()):""} onChange={e=>saveLoan(loans.map(x=>x.id===l.id?{...x,totalUSD:parseFloat(stripC(e.target.value))||0}:x))}/></td>
            <td className={cellIn+" text-right text-slate-600"} style={numStyle}>{(l.pct*100).toFixed(1)}%</td>
            <td className="px-1"><button onClick={()=>{if(confirm("Sil?"))saveLoan(loans.filter(x=>x.id!==l.id));}} className="text-slate-300 hover:text-rose-500 text-[13px] leading-none px-1">×</button></td>
          </tr>)}
          <tr className="bg-slate-50 border-t border-slate-200">
            <td className="px-3 py-2 text-right text-[11px] font-semibold text-slate-600 uppercase tracking-wide">Toplam Rotatif</td>
            <td className="px-2 py-2 text-right text-slate-700" style={numStyleBold}>{fmt(loans.reduce((s,l)=>s+l.usd,0))}</td>
            <td className="px-2 py-2 text-right text-slate-700" style={numStyleBold}>{fmt(loans.reduce((s,l)=>s+l.eur,0))}</td>
            <td className="px-2 py-2 text-right bg-amber-50/40 text-slate-800" style={{...numStyleBold,fontSize:"14px"}}>{fmt(loans.reduce((s,l)=>s+l.totalUSD,0))}</td>
            <td className="px-2 py-2 text-right text-slate-600" style={numStyle}>100.0%</td>
            <td></td>
          </tr>
          </tbody>
        </table>
        <div className="px-3 py-2 bg-slate-50/50 border-t border-slate-200">
          <button onClick={()=>saveLoan([...loans,{id:"ln_"+Date.now(),bank:"",usd:0,eur:0,totalUSD:0,pct:0}])} className="px-3 py-1 bg-violet-500 hover:bg-violet-600 text-white rounded-md text-[11px] font-medium transition-colors">+ Banka Ekle</button>
        </div>
      </div>

      {/* Rates */}
      <div className="bg-white rounded-lg shadow-sm border border-slate-200 overflow-hidden">
        <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 text-[11px] font-semibold text-slate-700 tracking-wide uppercase">Kurlar</div>
        <table className="border-collapse w-full">
          <tbody>{(() => {const pairs=[];for(let i=0;i<rates.length;i+=2)pairs.push([rates[i],rates[i+1]]);return pairs.map((pair,pi)=><tr key={pi} className="hover:bg-slate-50/60 border-b border-slate-100 last:border-b-0">
            {pair.map((r,ri)=>r?<React.Fragment key={r.id}><td className={cellIn} style={{minWidth:110}}><input className="w-full text-[11px] font-medium outline-none text-slate-800" value={r.label} onChange={e=>saveRate(rates.map(x=>x.id===r.id?{...x,label:e.target.value}:x))}/></td>
            <td className={cellIn+" text-right"} style={{minWidth:140}}><input className="w-full text-right outline-none text-slate-700" style={numStyle} value={r.value.toFixed(4)} onChange={e=>saveRate(rates.map(x=>x.id===r.id?{...x,value:parseFloat(e.target.value)||0}:x))}/></td></React.Fragment>:<React.Fragment key={"blank_"+ri}><td className={cellIn}></td><td className={cellIn}></td></React.Fragment>)}
          </tr>);})()}</tbody>
        </table>
      </div>
    </div>;
  };

  return<div className="relative" onDragOver={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop}>
    {isDraggingFile&&<div className="absolute inset-0 z-50 bg-teal-500/20 border-4 border-dashed border-teal-500 rounded-lg flex items-center justify-center pointer-events-none"><div className="bg-white px-6 py-4 rounded-xl shadow-xl border border-teal-200"><div className="text-2xl font-bold text-teal-700 mb-1">📥 Daily Reporting.xlsx'i bırakın</div><div className="text-[11px] text-gray-500 text-center">Tüm 3 sayfa otomatik içe aktarılır</div></div></div>}
    {dropStatus&&<div className={"mb-3 px-3 py-2 rounded-lg text-xs font-semibold "+(dropStatus.startsWith("✓")?"bg-emerald-50 text-emerald-700 border border-emerald-200":dropStatus.startsWith("Hata")||dropStatus.startsWith("Sadece")||dropStatus.startsWith("Fatura")?"bg-red-50 text-red-700 border border-red-200":"bg-blue-50 text-blue-700 border border-blue-200")}>{dropStatus}</div>}

    {/* Header */}
    <div className="flex justify-between items-end mb-5 flex-wrap gap-3 pb-3 border-b border-slate-100">
      <div>
        <h2 className="text-[20px] font-semibold text-slate-800 tracking-tight">Invoices</h2>
        <p className="text-[11px] text-slate-500 mt-0.5">We Soda fatura portali — <span className="font-medium text-slate-600" style={numFont}>{live.length}</span> kayıt, Excel raporu formatıyla</p>
      </div>
      <div className="flex items-center gap-2 flex-wrap">
        <label className="inline-flex items-center gap-2 px-3.5 py-2 bg-white hover:bg-slate-50 text-slate-700 rounded-lg text-[11px] font-medium border border-slate-200 shadow-sm cursor-pointer transition-colors">
          <svg className="w-3.5 h-3.5 text-slate-500" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>
          <span>Excel Raporu Yükle</span>
          <input type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f);e.target.value="";}}/>
        </label>
        <button onClick={()=>{if(!confirm("Tüm fatura geçmişi ve hafıza silinecek. Emin misiniz?"))return;saveInv([]);try{localStorage.removeItem(PEOPLE_KEY);}catch{}}} className="inline-flex items-center gap-1.5 px-3 py-2 bg-white hover:bg-slate-50 text-slate-500 hover:text-rose-600 rounded-lg text-[11px] font-medium border border-slate-200 shadow-sm transition-colors" title="Tümünü Sil">
          <svg className="w-3 h-3" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6M1 7h22M10 7V4a1 1 0 011-1h2a1 1 0 011 1v3"/></svg>
          <span>Temizle</span>
        </button>
      </div>
    </div>

    {/* Sheet tabs */}
    <div className="flex gap-6 border-b border-slate-200 mb-4 px-2">
      {[
        {k:"daily",label:"Vadesi Gelenler",color:"#059669"},
        {k:"aysonu",label:"Vadesi Yaklaşanlar",color:"#d97706"},
        {k:"accounts",label:"Hesaplar & Krediler",color:"#7c3aed"},
      ].map(t=><button key={t.k} onClick={()=>setActiveSheet(t.k as any)} className={"px-1 py-2.5 text-[12px] font-medium border-b-2 transition -mb-px "+(activeSheet===t.k?"text-slate-800":"border-transparent text-slate-500 hover:text-slate-700")} style={activeSheet===t.k?{borderBottomColor:t.color,color:t.color}:{}}>{t.label}</button>)}
    </div>

    {activeSheet==="daily"&&renderSheet("daily")}
    {activeSheet==="aysonu"&&renderSheet("aysonu")}
    {activeSheet==="accounts"&&renderAccounts()}

    <datalist id="inv-people">{people.map(p=><option key={p} value={p}/>)}</datalist>
  </div>;
}



// ═══════════════════════════════════════════════════════════════════════════
// DAILY REPORT — Soda Grubu Özet Bilgileri / We Soda Treasury Daily Report
// 3-page bilingual (TR/EN) report: UK Cash Report, US Cash Report, TR Daily Cash Flow
// Mirrors the Excel template format (16 April 2026 layout).
// Features: drag-and-drop Sources/Uses reordering, Kyriba Excel import,
// PDF export with landscape A4 layout matching original.
// ═══════════════════════════════════════════════════════════════════════════

type DRLang = "tr" | "en";
type DRPage = "uk" | "us" | "tr";
type DRFlow = { id:string; desc:string; amount:number; color?:string; };
type DRLoan = { id:string; bank:string; tl:number; ihrUSD:number; ihrEUR:number; shortUSD:number; midUSD:number; midEUR:number; midTotalUSD:number; totalUSD:number; };
type DRDisc = { id:string; bank:string; usd:number; eur:number; gbp:number; totalUSD:number; };
type DRDeposit = { id:string; group:"tr"|"uk"|"us"; company:string; tl:number; usd:number; eur:number; gbp:number; totalUSD:number; };
type DRAR = { id:string; name:string; amount:number; };
type DRARRegion = "tr_rec"|"tr_pay"|"us_rec"|"us_pay";
type DRLimit = { id:string; bank:string; usd:number; tl:number; totalUSD:number; };
type DRReserveAcc = { id:string; name:string; amount:number; };
type DRTRFlow = { id:string; section:string; counterparty:string; description:string; amount:number; color?:string; };  // TR daily cash flow

// Color palette for per-row text coloring
const DR_COLORS: {label:string, value:string}[] = [
  {label:"Siyah", value:""},
  {label:"Kırmızı", value:"#C00000"},
  {label:"Mavi", value:"#2F5FA8"},
  {label:"Yeşil", value:"#00703C"},
  {label:"Turuncu", value:"#C65911"},
  {label:"Mor", value:"#7030A0"},
];

type DailyReportData = {
  date: string;  // DD/MM/YYYY
  lang: DRLang;
  // UK Page
  ukSources: DRFlow[];
  ukUses: DRFlow[];
  loans: DRLoan[];
  discounted: DRDisc[];
  depositsTR: DRDeposit[];
  depositsUK: DRDeposit[];
  depositsUS: DRDeposit[];
  trReceivables: DRAR[];
  trPayables: DRAR[];
  usReceivables: DRAR[];
  usPayables: DRAR[];
  limits: DRLimit[];
  reserveAccounts: DRReserveAcc[];
  fx: { eur_usd:number; usd_try:number; eur_try:number; gbp_usd:number };
  // US Page
  usSources: DRFlow[];
  usUses: DRFlow[];
  // TR Page (TL denominated daily cash flow)
  trInflows: DRTRFlow[];
  trOutflows: DRTRFlow[];
};

// Complete bilingual label dictionary — all strings used in the report
const DR_LABEL: Record<string, {tr:string, en:string}> = {
  titleMain:     {tr:"Soda Grubu Özet Bilgileri", en:"We Soda Treasury Daily Report"},
  ukReportName:  {tr:"UK Günlük Kasa Raporu", en:"UK Daily Cash Report"},
  usReportName:  {tr:"US Günlük Kasa Raporu", en:"US Daily Cash Report"},
  trReportName:  {tr:"SODA GRUBU TR — GÜNLÜK NAKİT AKIŞ TABLOSU (Kazan, Eti, Denmar, WIDT Konsolide)", en:"SODA GROUP TR — DAILY CASH FLOW (Kazan, Eti, Denmar, WIDT Consolidated)"},
  realized:      {tr:"GERÇEKLEŞMİŞ", en:"REALIZED"},
  sources:       {tr:"Girişler (USD)", en:"Sources (USD)"},
  uses:          {tr:"Çıkışlar (USD)", en:"Uses (USD)"},
  trIn:          {tr:"GİRİŞLER", en:"INFLOWS"},
  trOut:         {tr:"ÇIKIŞLAR", en:"OUTFLOWS"},
  total:         {tr:"Toplam (USD)", en:"Total (USD)"},
  totalTL:       {tr:"Günlük İşlemler Toplamı TL", en:"Daily Transactions Total TL"},
  totalPayTL:    {tr:"Günlük Ödemeler Toplamı TL", en:"Daily Payments Total TL"},
  grandTL:       {tr:"Genel Toplam TL", en:"Grand Total TL"},
  depositAdded:  {tr:"Mevduata Eklenen", en:"Added to Deposit"},
  depositUsed:   {tr:"Mevduattan Kullanım", en:"Withdrawn from Deposit"},
  counterparty:  {tr:"Muhatap", en:"Counterparty"},
  explanation:   {tr:"Açıklama", en:"Description"},
  amountTL:      {tr:"Tutar(TL)", en:"Amount(TL)"},
  bankName:      {tr:"Banka Adı", en:"Bank Name"},
  shortLoan:     {tr:"Kısa Vadeli Krediler", en:"Short Term Loans"},
  shortLoanT:    {tr:"Kısa Vadeli Krediler Toplam (USD)", en:"Short Term Loans Total (USD)"},
  midLoan:       {tr:"Orta Uzun Vadeli Krediler", en:"Mid-Long Term Loans"},
  midLoanT:      {tr:"Orta Uzun Vadeli Krediler Toplam (USD)", en:"Mid-Long Term Loans Total (USD)"},
  grandUSD:      {tr:"Genel Toplam (USD)", en:"Total in (USD)"},
  tl:            {tr:"TL", en:"TL"},
  ihrUSD:        {tr:"İhracat/USD", en:"Export/USD"},
  ihrEUR:        {tr:"İhracat/EUR", en:"Export/EUR"},
  usd:           {tr:"USD", en:"USD"},
  eur:           {tr:"EUR", en:"EUR"},
  gbp:           {tr:"GBP", en:"GBP"},
  sumRow:        {tr:"Toplam", en:"Total"},
  discExp:       {tr:"İskonto Edilen İhracat Bedelleri", en:"Discounted Export Receivables"},
  depStatus:     {tr:"Mevduat Durumu", en:"Deposit Status"},
  company:       {tr:"Mevduat Türü", en:"Company"},
  tlDep:         {tr:"TL Mevduatlar", en:"TL Deposit"},
  usdDep:        {tr:"USD Mevduatlar", en:"USD Deposit"},
  eurDep:        {tr:"EUR Mevduatlar", en:"EUR Deposit"},
  gbpDep:        {tr:"GBP Mevduatlar", en:"GBP Deposit"},
  trDeposit:     {tr:"TR Serbest Mevduatlar", en:"TR Current Accounts"},
  ukDeposit:     {tr:"UK Serbest Mevduatlar", en:"UK Current Accounts"},
  usDeposit:     {tr:"US Serbest Mevduatlar", en:"US Current Accounts"},
  trReceivables: {tr:"TR Alacaklar", en:"TR Receivables"},
  trPayables:    {tr:"TR Borçlar", en:"TR Payables"},
  usReceivables: {tr:"US Alacaklar", en:"US Receivables"},
  usPayables:    {tr:"US Payables", en:"US Payables"},
  balance:       {tr:"Bakiye", en:"Amount"},
  grandTotal:    {tr:"Genel Toplam", en:"Grand Total"},
  limits:        {tr:"Kullanıma Hazır Limitler", en:"Unused Available Credit Limits"},
  reserve:       {tr:"Bloke Mevduatlar", en:"Reserve Accounts"},
  exportsRow:    {tr:"İhracat", en:"Exports"},
  domestic:      {tr:"Yurtiçi Alacaklar", en:"Domestic Receivables"},
  suppliers:     {tr:"Satıcılar", en:"Supplier"},
  pastPeriod:    {tr:"Geçmiş Dönem", en:"Overdue"},
  currentPeriod: {tr:"Güncel Dönem", en:"Due"},
};

// 16 April 2026 template defaults
const DR_LOANS_2026: DRLoan[] = [
  {id:"l1", bank:"We Soda Bond - 5Y", tl:0, ihrUSD:0, ihrEUR:0, shortUSD:0, midUSD:980000000, midEUR:0, midTotalUSD:980000000, totalUSD:980000000},
  {id:"l2", bank:"We Soda Bond - 7Y", tl:0, ihrUSD:0, ihrEUR:0, shortUSD:0, midUSD:750000000, midEUR:0, midTotalUSD:750000000, totalUSD:750000000},
  {id:"l3", bank:"WE Industries Holdings Isbank", tl:0, ihrUSD:0, ihrEUR:377144557, shortUSD:0, midUSD:0, midEUR:0, midTotalUSD:0, totalUSD:444427146},
  {id:"l4", bank:"WE Soda US - Acquisition Loan", tl:0, ihrUSD:0, ihrEUR:0, shortUSD:0, midUSD:409500000, midEUR:0, midTotalUSD:409500000, totalUSD:409500000},
  {id:"l5", bank:"US ORRI Bond", tl:0, ihrUSD:0, ihrEUR:0, shortUSD:0, midUSD:396672544, midEUR:0, midTotalUSD:396672544, totalUSD:396672544},
  {id:"l6", bank:"We Soda RCF", tl:0, ihrUSD:0, ihrEUR:0, shortUSD:0, midUSD:85000000, midEUR:30000000, midTotalUSD:120352000, totalUSD:120352000},
  {id:"l7", bank:"Kazan Denizbank RCF", tl:0, ihrUSD:0, ihrEUR:52500000, shortUSD:0, midUSD:0, midEUR:0, midTotalUSD:0, totalUSD:61866000},
  {id:"l8", bank:"WIDT Eximbank", tl:225000000, ihrUSD:5000000, ihrEUR:0, shortUSD:10929152, midUSD:0, midEUR:0, midTotalUSD:0, totalUSD:10929152},
];

const DR_DISC_2026: DRDisc[] = [
  {id:"d1", bank:"BNPP - Soda World UK", usd:45606354, eur:5837578, gbp:9780103, totalUSD:65045233},
  {id:"d2", bank:"BNPP - Soda World UK (US ANSAC)", usd:54384290, eur:0, gbp:0, totalUSD:54384290},
  {id:"d3", bank:"ABC Bank - We Soda Wyoming", usd:33388594, eur:0, gbp:0, totalUSD:33388594},
  {id:"d4", bank:"ABC Bank - Soda World UK", usd:0, eur:18391875, gbp:0, totalUSD:21672986},
  {id:"d5", bank:"BNPP - Soda World Europe", usd:0, eur:5429229, gbp:0, totalUSD:6397804},
  {id:"d6", bank:"ABC Bank - Soda World Europe", usd:0, eur:3583755, gbp:0, totalUSD:4223097},
  {id:"d7", bank:"Bankinter", usd:0, eur:10601964, gbp:0, totalUSD:12493354},
  {id:"d8", bank:"BBVA", usd:0, eur:10345712, gbp:0, totalUSD:12191387},
  {id:"d9", bank:"Santander", usd:0, eur:9922671, gbp:0, totalUSD:11692876},
  {id:"d10", bank:"Mitsui", usd:173691074, eur:0, gbp:0, totalUSD:173691074},
  {id:"d11", bank:"Traxys", usd:0, eur:62500000, gbp:0, totalUSD:73650000},
];

const DR_DEP_TR_2026: DRDeposit[] = [
  {id:"t1", group:"tr", company:"Eti", tl:1090964918, usd:13388253, eur:22797616, gbp:172, totalUSD:64677001},
  {id:"t2", group:"tr", company:"Kazan", tl:5934870, usd:3474, eur:20865734, gbp:1339, totalUSD:24726336},
  {id:"t3", group:"tr", company:"WIDT", tl:19291789, usd:3283, eur:38228, gbp:0, totalUSD:480223},
  {id:"t4", group:"tr", company:"Denmar", tl:4182603, usd:5615, eur:7611, gbp:15, totalUSD:108241},
];

const DR_DEP_UK_2026: DRDeposit[] = [
  {id:"u1", group:"uk", company:"Soda World UK", tl:0, usd:13579324, eur:1050095, gbp:321174, totalUSD:15252172},
  {id:"u2", group:"uk", company:"Soda World Europe", tl:0, usd:20259, eur:1801241, gbp:4278, totalUSD:2148641},
  {id:"u3", group:"uk", company:"We Soda", tl:0, usd:7623, eur:72174, gbp:16693, totalUSD:115304},
  {id:"u4", group:"uk", company:"WE Industries Holdings", tl:0, usd:12541, eur:6224, gbp:3410, totalUSD:24498},
];

const DR_DEP_US_2026: DRDeposit[] = [
  {id:"us1", group:"us", company:"We Soda Wyoming", tl:0, usd:6234692, eur:0, gbp:0, totalUSD:6234692},
  {id:"us2", group:"us", company:"ANSAC", tl:0, usd:4751212, eur:0, gbp:0, totalUSD:4751212},
  {id:"us3", group:"us", company:"WE Soda Enterprises", tl:0, usd:265182, eur:0, gbp:0, totalUSD:265182},
  {id:"us4", group:"us", company:"We ORRI", tl:0, usd:26098, eur:0, gbp:0, totalUSD:26098},
  {id:"us5", group:"us", company:"West Soda", tl:0, usd:16852, eur:0, gbp:0, totalUSD:16852},
  {id:"us6", group:"us", company:"We Soda US", tl:0, usd:14181, eur:0, gbp:0, totalUSD:14181},
  {id:"us7", group:"us", company:"Soda World US", tl:0, usd:4698, eur:0, gbp:0, totalUSD:4698},
  {id:"us8", group:"us", company:"Imperial", tl:0, usd:3022, eur:0, gbp:0, totalUSD:3022},
];

const DR_TR_REC_2026: DRAR[] = [
  {id:"tr1", name:"İhracat", amount:55888387},
  {id:"tr2", name:"Yurtiçi Alacaklar", amount:56745327},
];

const DR_TR_PAY_2026: DRAR[] = [
  {id:"tp1", name:"Satıcılar", amount:12383990},
  {id:"tp2", name:"Geçmiş Dönem", amount:18268},
  {id:"tp3", name:"Güncel Dönem", amount:12365722},
];

const DR_US_REC_2026: DRAR[] = [
  {id:"ur1", name:"İhracat", amount:10857954},
  {id:"ur2", name:"Yurtiçi Alacaklar", amount:88908382},
];

const DR_US_PAY_2026: DRAR[] = [
  {id:"up1", name:"Satıcılar", amount:22219022},
  {id:"up2", name:"Geçmiş Dönem", amount:11407621},
  {id:"up3", name:"Güncel Dönem", amount:10811401},
];

const DR_LIMITS_2026: DRLimit[] = [
  {id:"lm1", bank:"We Soda RCF", usd:189648000, tl:189648000, totalUSD:189648000},
  {id:"lm2", bank:"Discountable Cheques", usd:2545263, tl:2545263, totalUSD:2545263},
];

const DR_RESERVE_2026: DRReserveAcc[] = [
  {id:"ra1", name:"US ORRI Bond", amount:19290309},
  {id:"ra2", name:"Takasbank Kazan Soda", amount:1419728},
  {id:"ra3", name:"Takasbank Eti Soda", amount:1279134},
];

function DailyReport({fx}:{fx:typeof FX_D}){
  const[data,setData]=useState<DailyReportData>(()=>{
    try{
      const s=localStorage.getItem("daily_report_current");
      if(s){const p=JSON.parse(s);if(p&&p.date&&p.ukSources)return p;}
    }catch{}
    const today=new Date();
    const dd=String(today.getDate()).padStart(2,"0"),mm=String(today.getMonth()+1).padStart(2,"0");
    return {
      date: `${dd}/${mm}/${today.getFullYear()}`,
      lang: "tr",
      ukSources: [], ukUses: [],
      loans: DR_LOANS_2026,
      discounted: DR_DISC_2026,
      depositsTR: DR_DEP_TR_2026,
      depositsUK: DR_DEP_UK_2026,
      depositsUS: DR_DEP_US_2026,
      trReceivables: DR_TR_REC_2026,
      trPayables: DR_TR_PAY_2026,
      usReceivables: DR_US_REC_2026,
      usPayables: DR_US_PAY_2026,
      limits: DR_LIMITS_2026,
      reserveAccounts: DR_RESERVE_2026,
      fx: {eur_usd:fx.eur_usd||1.1784, usd_try:fx.usd_try||44.6681, eur_try:(fx.eur_usd||1.1784)*(fx.usd_try||44.6681), gbp_usd:1.3557},
      usSources: [], usUses: [],
      trInflows: [], trOutflows: [],
    };
  });

  const[activePage,setActivePage]=useState<DRPage>("uk");
  const[dragItem,setDragItem]=useState<{list:string, id:string}|null>(null);
  const[dropStatus,setDropStatus]=useState("");
  const[isDraggingFile,setIsDraggingFile]=useState(false);

  const save=(d:DailyReportData)=>{setData(d);try{trackedSetItem("daily_report_current",JSON.stringify(d));}catch{}};
  const L=(k:string)=>DR_LABEL[k]?.[data.lang]||k;
  const numCal={fontFamily:"'Calibri','Segoe UI','Arial',system-ui,sans-serif",fontVariantNumeric:"tabular-nums" as const,letterSpacing:"0.01em"};

  // ═══ FLOW CRUD (Sources/Uses for UK and US) ═══
  const flowField=(p:"uk"|"us",t:"sources"|"uses")=>(p==="uk"?(t==="sources"?"ukSources":"ukUses"):(t==="sources"?"usSources":"usUses")) as keyof DailyReportData;
  const addFlow=(p:"uk"|"us",t:"sources"|"uses")=>{const f=flowField(p,t);save({...data,[f]:[...(data[f] as DRFlow[]),{id:"fl_"+Date.now()+"_"+Math.random().toString(36).slice(2,6),desc:"",amount:0}]});};
  const updFlow=(p:"uk"|"us",t:"sources"|"uses",id:string,fi:keyof DRFlow,v:any)=>{const f=flowField(p,t);save({...data,[f]:(data[f] as DRFlow[]).map(r=>r.id===id?{...r,[fi]:v}:r)});};
  const delFlow=(p:"uk"|"us",t:"sources"|"uses",id:string)=>{const f=flowField(p,t);save({...data,[f]:(data[f] as DRFlow[]).filter(r=>r.id!==id)});};

  // ═══ DRAG & DROP for Sources/Uses reordering ═══
  const onDragStart=(listKey:string,id:string)=>(e:React.DragEvent)=>{setDragItem({list:listKey,id});e.dataTransfer.effectAllowed="move";};
  const onDragOverRow=(e:React.DragEvent)=>{e.preventDefault();e.dataTransfer.dropEffect="move";};
  const onDropRow=(listKey:string,targetId:string)=>(e:React.DragEvent)=>{
    e.preventDefault();e.stopPropagation();
    if(!dragItem||dragItem.list!==listKey||dragItem.id===targetId){setDragItem(null);return;}
    const [p,t]=listKey.split("_") as ["uk"|"us","sources"|"uses"];
    const f=flowField(p,t);
    const arr=[...(data[f] as DRFlow[])];
    const fromIdx=arr.findIndex(r=>r.id===dragItem.id);
    const toIdx=arr.findIndex(r=>r.id===targetId);
    if(fromIdx<0||toIdx<0){setDragItem(null);return;}
    const [moved]=arr.splice(fromIdx,1);
    arr.splice(toIdx,0,moved);
    save({...data,[f]:arr});
    setDragItem(null);
  };

  // ═══ LOANS / DISC / DEPOSITS / AR / LIMITS / RESERVE CRUD ═══
  const crudAdd=<T extends {id:string}>(field:keyof DailyReportData,newRow:T)=>{save({...data,[field]:[...(data[field] as any[]),newRow]});};
  const crudUpd=(field:keyof DailyReportData,id:string,f:string,v:any)=>{save({...data,[field]:(data[field] as any[]).map(r=>r.id===id?{...r,[f]:v}:r)});};
  const crudDel=(field:keyof DailyReportData,id:string)=>{save({...data,[field]:(data[field] as any[]).filter(r=>r.id!==id)});};

  // ═══ KYRIBA EXCEL IMPORT ═══
  const ensureSheetJS=async()=>{if((window as any).XLSX)return (window as any).XLSX;return new Promise<any>((resolve,reject)=>{const s=document.createElement("script");s.src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";s.onload=()=>resolve((window as any).XLSX);s.onerror=()=>reject(new Error("XLSX"));document.head.appendChild(s);});};
  const handleKyribaFile=async(file:File)=>{
    try{
      setDropStatus("Kyriba dosyası okunuyor...");
      const XLSX=await ensureSheetJS();
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:"array",cellDates:true});
      // Kyriba exports typically have columns: Date, Account, Description, Debit/Credit, Amount, Currency
      // We scan every sheet and classify rows into Sources (credit/inflow) or Uses (debit/outflow) based on amount sign or Debit/Credit column
      const newSources:DRFlow[]=[];
      const newUses:DRFlow[]=[];
      for(const sn of wb.SheetNames){
        const ws=wb.Sheets[sn];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}) as any[][];
        if(!rows||rows.length<2)continue;
        // Find header row — look for a row with "Amount" or "Tutar" or "Description"/"Açıklama"
        let headerIdx=-1;
        let descCol=-1, amtCol=-1, debitCol=-1, creditCol=-1;
        for(let r=0;r<Math.min(10,rows.length);r++){
          const row=rows[r];
          for(let c=0;c<row.length;c++){
            const v=String(row[c]||"").toLowerCase().trim();
            if(v==="description"||v==="açıklama"||v==="libelle"||v==="libellé")descCol=c;
            if(v==="amount"||v==="tutar"||v==="montant"||v==="amount (eur)"||v==="amount (usd)")amtCol=c;
            if(v==="debit"||v==="borç")debitCol=c;
            if(v==="credit"||v==="alacak")creditCol=c;
          }
          if(descCol>=0&&(amtCol>=0||(debitCol>=0&&creditCol>=0))){headerIdx=r;break;}
        }
        if(headerIdx<0)continue;
        // Parse data rows
        for(let r=headerIdx+1;r<rows.length;r++){
          const row=rows[r];if(!row||!row.length)continue;
          const desc=String(row[descCol]||"").trim();
          if(!desc||desc.length<3)continue;
          let amt=0, isCredit=false;
          if(amtCol>=0){
            const raw=row[amtCol];
            if(typeof raw==="number"){amt=raw;isCredit=amt>0;}
            else{const n=parseFloat(String(raw||"").replace(/[,\s]/g,""));if(!isNaN(n)){amt=Math.abs(n);isCredit=n>0;}}
          } else if(debitCol>=0&&creditCol>=0){
            const d=parseFloat(String(row[debitCol]||"0").replace(/[,\s]/g,""))||0;
            const c=parseFloat(String(row[creditCol]||"0").replace(/[,\s]/g,""))||0;
            if(c>0){amt=c;isCredit=true;}else if(d>0){amt=d;isCredit=false;}
          }
          if(amt===0)continue;
          const item={id:"fl_"+Date.now()+"_"+Math.random().toString(36).slice(2,6),desc,amount:Math.abs(amt)};
          if(isCredit)newSources.push(item);else newUses.push(item);
        }
      }
      if(newSources.length===0&&newUses.length===0){
        setDropStatus("Kyriba formatı anlaşılamadı. Kolon başlıkları 'Description/Açıklama', 'Amount/Tutar' veya 'Debit/Credit' olmalı.");
        setTimeout(()=>setDropStatus(""),5000);return;
      }
      // Merge with existing UK Sources/Uses (current active page)
      const mergedSources=[...(data.ukSources),...newSources];
      const mergedUses=[...(data.ukUses),...newUses];
      save({...data,ukSources:mergedSources,ukUses:mergedUses});
      setDropStatus(`✓ Kyriba: ${newSources.length} giriş, ${newUses.length} çıkış eklendi`);
      setTimeout(()=>setDropStatus(""),4000);
    }catch(err){
      setDropStatus("Hata: "+(err instanceof Error?err.message:String(err)));
      setTimeout(()=>setDropStatus(""),5000);
    }
  };
  const onFileDrop=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(false);const f=e.dataTransfer.files[0];if(f)handleKyribaFile(f);};
  const onFileDragOver=(e:React.DragEvent)=>{e.preventDefault();setIsDraggingFile(true);};
  const onFileDragLeave=(e:React.DragEvent)=>{if(e.currentTarget===e.target)setIsDraggingFile(false);};

  // Totals
  const sumFlow=(arr:DRFlow[])=>arr.reduce((s,r)=>s+(r.amount||0),0);
  const loanTot=(f:keyof DRLoan)=>data.loans.reduce((s,l)=>s+(l[f] as number||0),0);
  const discTot=(f:keyof DRDisc)=>data.discounted.reduce((s,d)=>s+(d[f] as number||0),0);
  const depTot=(arr:DRDeposit[],f:keyof DRDeposit)=>arr.reduce((s,d)=>s+(d[f] as number||0),0);

  // ═══ PDF EXPORT — Multi-page, landscape, matches original format ═══
  const exportToPDF=()=>{
    const w=window.open("","_blank","width=1600,height=1000");
    if(!w)return;
    w.document.write(renderPrintableHTML());
    w.document.close();
    setTimeout(()=>{try{w.print();}catch{}},600);
  };

  const renderPrintableHTML=()=>{
    // Color palette mirrors the original Excel report EXACTLY:
    //   #4472C4 = steel blue banner + section headers (white text)
    //   #DDEBF7 = light pastel cyan (sub-banner + source/use rows + general row tint)
    //   #BDD7EE = deeper pastel blue (Short Term Loans totals, Discounted totals)
    //   #9DC3E6 = medium blue (Mid-Long totals)
    //   #F4B183 = orange (Grand Total row)
    //   #C6EFCE = light green (Receivables/Payables totals + Deposit section totals)
    //   #FFF2CC = light yellow (main totals)
    // Row text colors (red/blue) survive the user's per-row styling in HTML `style` attributes.
    const css=`
      @page{size:A4 landscape;margin:5mm;}
      *{box-sizing:border-box;}
      body{font-family:Calibri,'Segoe UI',Arial,sans-serif;font-size:7.5pt;margin:0;color:#1f2937;-webkit-print-color-adjust:exact;print-color-adjust:exact;color-adjust:exact;}
      .page{page-break-after:always;padding:0;}
      .page:last-child{page-break-after:auto;}
      .banner{display:flex;align-items:stretch;background:#4472C4;border-bottom:0;}
      .banner .title{flex:1;background:#4472C4;color:#ffffff;display:flex;align-items:center;justify-content:center;font-size:16pt;font-weight:bold;padding:10pt 0;letter-spacing:0.3pt;}
      .banner .logo{background:#ffffff;display:flex;align-items:center;padding:0 14pt;gap:6pt;min-width:130pt;justify-content:center;}
      .banner .logo .we{background:#1E90FF;color:#ffffff;font-weight:bold;padding:4pt 9pt;font-size:14pt;border-radius:3pt 0 0 3pt;}
      .banner .logo .soda{background:#FFC72C;color:#1f2937;font-weight:bold;padding:4pt 9pt;font-size:14pt;border-radius:0 3pt 3pt 0;}
      .banner .logo .tag{font-size:5.5pt;color:#558A3B;font-weight:bold;letter-spacing:1px;line-height:1.1;}
      .banner .date{background:#4472C4;color:#ffffff;padding:10pt 18pt;display:flex;align-items:center;justify-content:center;font-weight:bold;font-size:11pt;min-width:110pt;border-left:1pt solid #6487CE;}
      .subHdr{background:#DDEBF7;padding:4pt 10pt;font-weight:bold;color:#1f2937;text-align:center;font-size:10pt;border:0.4pt solid #8FB3E0;}
      table{border-collapse:collapse;width:100%;font-size:7.5pt;margin:0;table-layout:auto;}
      td,th{border:0.4pt solid #8A8A8A;padding:1.5pt 4pt;vertical-align:middle;line-height:1.2;}
      th{background:#4472C4;color:#ffffff;font-weight:bold;text-align:center;font-size:7.5pt;padding:3pt 4pt;}
      .num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap;}
      .left{text-align:left;}
      .ct{text-align:center;}
      /* Section background tints — mirror Excel palette */
      .total{background:#FFF2CC;font-weight:bold;}              /* light yellow totals */
      .grand{background:#F4B183;font-weight:bold;}              /* orange grand totals */
      .sec{background:#4472C4;color:#ffffff;font-weight:bold;padding:2.5pt 5pt;font-size:7.5pt;text-align:left;}
      .subsec{background:#305496;color:#ffffff;font-weight:bold;padding:2pt 5pt;font-size:7.5pt;text-align:center;}
      /* Tints that appear on header rows grouping columns */
      .tBlue{background:#DDEBF7;color:#1f2937;}                 /* light pastel cyan (sources/uses) */
      .tBlueMed{background:#BDD7EE;color:#1f2937;}              /* short-term loans + disc totals */
      .tBlueDark{background:#9DC3E6;color:#1f2937;}             /* mid-long loan totals */
      .tOrange{background:#F4B183;color:#1f2937;font-weight:bold;}  /* grand total orange */
      .tGreen{background:#C6EFCE;color:#1f2937;}                /* AR/AP totals green */
      .tYellow{background:#FFF2CC;color:#1f2937;}               /* sub-total yellow */
      /* Legacy aliases — map older class names to new palette */
      .green{background:#C6EFCE;color:#1f2937;}
      .red{background:#DDEBF7;color:#C00000;}                   /* red text on pastel row */
      .orange{background:#F4B183;color:#1f2937;}
      .blue{background:#9DC3E6;color:#1f2937;}
      .yellow{background:#FFF2CC;color:#1f2937;}
      .violet{background:#DDEBF7;color:#1f2937;}
      .cyan{background:#BDD7EE;color:#1f2937;}
      .emerald{background:#C6EFCE;color:#1f2937;}
      .gap{border:0 !important;width:3pt;padding:0;}
      .flex2{display:grid;grid-template-columns:1fr 1fr;gap:6pt;margin-top:4pt;}
      .flex3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:4pt;margin-top:4pt;}
      .mt{margin-top:4pt;}
    `;

    // Logo block — matches We·Soda branding with green "WEST EAST SODA" tag
    const logoBlock=`<div class="logo"><div style="display:flex;"><span class="we">we</span><span class="soda">soda</span></div><div class="tag">WEST EAST<br>SODA</div></div>`;

    const fmtN=(n:number)=>n===0?"":n.toLocaleString("en-US",{maximumFractionDigits:0,useGrouping:true}).replace(/,/g,".");
    const fmtFX=(n:number)=>n===0?"":n.toLocaleString("en-US",{minimumFractionDigits:4,maximumFractionDigits:4}).replace(",","X").replace(".",",").replace("X",".");
    const maxLen=(...arrs:any[])=>Math.max(...arrs.map(a=>a.length),1);

    // UK PAGE
    const ukMax=maxLen(data.ukSources,data.ukUses);
    let ukFlowRows="";
    for(let i=0;i<ukMax;i++){
      const s=data.ukSources[i], u=data.ukUses[i];
      const sStyle=s?.color?`style="color:${s.color};font-weight:600;"`:"";
      const uStyle=u?.color?`style="color:${u.color};font-weight:600;"`:"";
      ukFlowRows+=`<tr>
        <td class="left" ${sStyle}>${s?s.desc:""}</td>
        <td class="num" ${sStyle}>${s?fmtN(s.amount):""}</td>
        <td style="border:0 !important;width:4pt;"></td>
        <td class="left" ${uStyle}>${u?u.desc:""}</td>
        <td class="num" ${uStyle}>${u?fmtN(u.amount):""}</td>
      </tr>`;
    }

    const ukLoanRows=data.loans.map(l=>`<tr>
      <td class="left">${l.bank}</td>
      <td class="num">${fmtN(l.tl)}</td>
      <td class="num">${fmtN(l.ihrUSD)}</td>
      <td class="num">${fmtN(l.ihrEUR)}</td>
      <td class="num">${fmtN(l.shortUSD)}</td>
      <td class="num orange"><b>${fmtN(l.shortUSD)}</b></td>
      <td class="num">${fmtN(l.midUSD)}</td>
      <td class="num">${fmtN(l.midEUR)}</td>
      <td class="num blue"><b>${fmtN(l.midTotalUSD)}</b></td>
      <td class="num yellow"><b>${fmtN(l.totalUSD)}</b></td>
    </tr>`).join("");

    const ukDiscRows=data.discounted.map(d=>`<tr>
      <td class="left">${d.bank}</td>
      <td class="num">${fmtN(d.usd)}</td>
      <td class="num">${fmtN(d.gbp)}</td>
      <td class="num">${fmtN(d.eur)}</td>
      <td class="num yellow"><b>${fmtN(d.totalUSD)}</b></td>
    </tr>`).join("");

    const makeDepRows=(arr:DRDeposit[])=>arr.map(d=>`<tr>
      <td class="left">${d.company}</td>
      <td class="num">${fmtN(d.tl)}</td>
      <td class="num">${fmtN(d.usd)}</td>
      <td class="num">${fmtN(d.eur)}</td>
      <td class="num">${fmtN(d.gbp)}</td>
      <td class="num yellow"><b>${fmtN(d.totalUSD)}</b></td>
    </tr>`).join("");

    // AR tables
    const makeARTable=(title:string,rec:DRAR[],pay:DRAR[])=>{
      const mx=maxLen(rec,pay);
      let rows="";
      for(let i=0;i<mx;i++){
        const r=rec[i],p=pay[i];
        rows+=`<tr><td class="left">${r?r.name:""}</td><td class="num">${r?fmtN(r.amount):""}</td>
          <td style="border:0 !important;width:4pt;"></td>
          <td class="left">${p?p.name:""}</td><td class="num">${p?fmtN(p.amount):""}</td></tr>`;
      }
      return `<table class="mt"><tr><th colspan="2" class="emerald">${title.split("|")[0]}</th><td style="border:0 !important;"></td><th colspan="2" class="red">${title.split("|")[1]}</th></tr>
        <tr><th class="left">${L("exportsRow")}</th><th>${L("balance")}</th><td style="border:0 !important;"></td><th class="left">${L("suppliers")}</th><th>${L("balance")}</th></tr>
        ${rows}
        <tr class="total"><td class="left">${L("total")}</td><td class="num">${fmtN(rec.reduce((s,r)=>s+r.amount,0))}</td><td style="border:0 !important;"></td><td class="left">${L("total")}</td><td class="num">${fmtN(pay.reduce((s,p)=>s+p.amount,0))}</td></tr>
      </table>`;
    };

    const limitsRows=data.limits.map(l=>`<tr>
      <td class="left">${l.bank}</td>
      <td class="num">${fmtN(l.usd)}</td>
      <td class="num">${fmtN(l.tl)}</td>
      <td class="num yellow"><b>${fmtN(l.totalUSD)}</b></td>
    </tr>`).join("");

    const reserveRows=data.reserveAccounts.map(r=>`<tr>
      <td class="left">${r.name}</td>
      <td class="num yellow"><b>${fmtN(r.amount)}</b></td>
    </tr>`).join("");

    const ukPage=`
    <div class="page">
      <div class="banner"><div class="title">${L("titleMain")}</div>${logoBlock}<div class="date">${data.date}</div></div>
      <div class="subHdr">${L("ukReportName")}</div>
      <table>
        <tr><th colspan="2">${L("sources")}</th><td style="border:0 !important;width:4pt;"></td><th colspan="2">${L("uses")}</th></tr>
        ${ukFlowRows}
        <tr class="total"><td class="left">${L("total")}</td><td class="num">${fmtN(sumFlow(data.ukSources))}</td><td style="border:0 !important;"></td><td class="left">${L("total")}</td><td class="num">${fmtN(sumFlow(data.ukUses))}</td></tr>
      </table>

      <table class="mt">
        <tr>
          <th rowspan="2" class="left">${L("bankName")}</th>
          <th colspan="4" class="orange">${L("shortLoan")}</th>
          <th rowspan="2" class="orange">${L("shortLoanT")}</th>
          <th colspan="2" class="blue">${L("midLoan")}</th>
          <th rowspan="2" class="blue">${L("midLoanT")}</th>
          <th rowspan="2" class="yellow">${L("grandUSD")}</th>
        </tr>
        <tr><th>${L("tl")}</th><th>${L("ihrUSD")}</th><th>${L("ihrEUR")}</th><th>${L("usd")}</th><th>${L("usd")}</th><th>${L("eur")}</th></tr>
        ${ukLoanRows}
        <tr class="total"><td class="left">${L("sumRow")}</td><td class="num">${fmtN(loanTot("tl"))}</td><td class="num">${fmtN(loanTot("ihrUSD"))}</td><td class="num">${fmtN(loanTot("ihrEUR"))}</td><td class="num">${fmtN(loanTot("shortUSD"))}</td><td class="num orange"><b>${fmtN(loanTot("shortUSD"))}</b></td><td class="num">${fmtN(loanTot("midUSD"))}</td><td class="num">${fmtN(loanTot("midEUR"))}</td><td class="num blue"><b>${fmtN(loanTot("midTotalUSD"))}</b></td><td class="num yellow"><b>${fmtN(loanTot("totalUSD"))}</b></td></tr>
      </table>

      <div class="flex2">
        <table>
          <tr><th rowspan="2" class="left">${L("bankName")}</th><th colspan="3" class="violet">${L("discExp")}</th><th rowspan="2" class="yellow">${L("total")}</th></tr>
          <tr><th>${L("usd")}</th><th>${L("gbp")}</th><th>${L("eur")}</th></tr>
          ${ukDiscRows}
          <tr class="total"><td class="left">${L("sumRow")}</td><td class="num">${fmtN(discTot("usd"))}</td><td class="num">${fmtN(discTot("gbp"))}</td><td class="num">${fmtN(discTot("eur"))}</td><td class="num yellow"><b>${fmtN(discTot("totalUSD"))}</b></td></tr>
        </table>

        <table>
          <tr><th colspan="6" class="emerald">${L("depStatus")}</th></tr>
          <tr><th class="left">${L("company")}</th><th>${L("tlDep")}</th><th>${L("usdDep")}</th><th>${L("eurDep")}</th><th>${L("gbpDep")}</th><th class="yellow">${L("total")}</th></tr>
          <tr><td colspan="6" class="sec left">${L("trDeposit")}</td></tr>
          ${makeDepRows(data.depositsTR)}
          <tr><td colspan="6" class="sec left">${L("ukDeposit")}</td></tr>
          ${makeDepRows(data.depositsUK)}
          <tr><td colspan="6" class="sec left">${L("usDeposit")}</td></tr>
          ${makeDepRows(data.depositsUS)}
          <tr class="total"><td class="left">${L("grandTotal")}</td>
            <td class="num">${fmtN(depTot([...data.depositsTR,...data.depositsUK,...data.depositsUS],"tl"))}</td>
            <td class="num">${fmtN(depTot([...data.depositsTR,...data.depositsUK,...data.depositsUS],"usd"))}</td>
            <td class="num">${fmtN(depTot([...data.depositsTR,...data.depositsUK,...data.depositsUS],"eur"))}</td>
            <td class="num">${fmtN(depTot([...data.depositsTR,...data.depositsUK,...data.depositsUS],"gbp"))}</td>
            <td class="num yellow"><b>${fmtN(depTot([...data.depositsTR,...data.depositsUK,...data.depositsUS],"totalUSD"))}</b></td>
          </tr>
        </table>
      </div>

      <div class="flex3">
        ${makeARTable(L("trReceivables")+"|"+L("trPayables"),data.trReceivables,data.trPayables)}
        ${makeARTable(L("usReceivables")+"|"+L("usPayables"),data.usReceivables,data.usPayables)}
        <div>
          <table>
            <tr><th colspan="4" class="cyan">${L("limits")}</th></tr>
            <tr><th class="left">${L("bankName")}</th><th>${L("usd")}</th><th>${L("tl")}</th><th class="yellow">${L("total")}</th></tr>
            ${limitsRows}
            <tr class="total"><td class="left">${L("sumRow")}</td><td class="num">${fmtN(data.limits.reduce((s,l)=>s+l.usd,0))}</td><td class="num">${fmtN(data.limits.reduce((s,l)=>s+l.tl,0))}</td><td class="num yellow"><b>${fmtN(data.limits.reduce((s,l)=>s+l.totalUSD,0))}</b></td></tr>
          </table>
          <table class="mt">
            <tr><th colspan="2" class="violet">${L("reserve")}</th></tr>
            ${reserveRows}
            <tr class="total"><td class="left">${L("total")}</td><td class="num yellow"><b>${fmtN(data.reserveAccounts.reduce((s,r)=>s+r.amount,0))}</b></td></tr>
          </table>
          <table class="mt">
            <tr><th colspan="2" class="violet">FX</th></tr>
            <tr><td class="left">EUR/USD :</td><td class="num">${fmtFX(data.fx.eur_usd)}</td></tr>
            <tr><td class="left">USD/TRY :</td><td class="num">${fmtFX(data.fx.usd_try)}</td></tr>
            <tr><td class="left">EUR/TRY :</td><td class="num">${fmtFX(data.fx.eur_try)}</td></tr>
            <tr><td class="left">GBP/USD :</td><td class="num">${fmtFX(data.fx.gbp_usd)}</td></tr>
          </table>
        </div>
      </div>
    </div>`;

    // US PAGE
    const usMax=maxLen(data.usSources,data.usUses);
    let usFlowRows="";
    for(let i=0;i<usMax;i++){
      const s=data.usSources[i], u=data.usUses[i];
      const sStyle=s?.color?`style="color:${s.color};font-weight:600;"`:"";
      const uStyle=u?.color?`style="color:${u.color};font-weight:600;"`:"";
      usFlowRows+=`<tr>
        <td class="left" ${sStyle}>${s?s.desc:""}</td>
        <td class="num" ${sStyle}>${s?fmtN(s.amount):""}</td>
        <td style="border:0 !important;width:4pt;"></td>
        <td class="left" ${uStyle}>${u?u.desc:""}</td>
        <td class="num" ${uStyle}>${u?fmtN(u.amount):""}</td>
      </tr>`;
    }

    const usPage=data.usSources.length>0||data.usUses.length>0?`
    <div class="page">
      <div class="banner"><div class="title">${L("titleMain")}</div>${logoBlock}<div class="date">${data.date}</div></div>
      <div class="subHdr">${L("usReportName")}</div>
      <table>
        <tr><th colspan="2">${L("sources")}</th><td style="border:0 !important;width:4pt;"></td><th colspan="2">${L("uses")}</th></tr>
        ${usFlowRows}
        <tr class="total"><td class="left">${L("total")}</td><td class="num">${fmtN(sumFlow(data.usSources))}</td><td style="border:0 !important;"></td><td class="left">${L("total")}</td><td class="num">${fmtN(sumFlow(data.usUses))}</td></tr>
      </table>
    </div>`:"";

    // TR PAGE (cash flow table in TL)
    const trMax=maxLen(data.trInflows,data.trOutflows);
    let trRows="";
    for(let i=0;i<trMax;i++){
      const inf=data.trInflows[i],out=data.trOutflows[i];
      const iStyle=inf?.color?`style="color:${inf.color};font-weight:600;"`:"";
      const oStyle=out?.color?`style="color:${out.color};font-weight:600;"`:"";
      trRows+=`<tr>
        <td class="left" ${iStyle}>${inf?inf.counterparty:""}</td>
        <td class="left" ${iStyle}>${inf?inf.description:""}</td>
        <td class="num" ${iStyle}>${inf?fmtN(inf.amount):""}</td>
        <td style="border:0 !important;width:4pt;"></td>
        <td class="left" ${oStyle}>${out?out.counterparty:""}</td>
        <td class="left" ${oStyle}>${out?out.description:""}</td>
        <td class="num" ${oStyle}>${out?fmtN(out.amount):""}</td>
      </tr>`;
    }
    const trInSum=data.trInflows.reduce((s,r)=>s+r.amount,0);
    const trOutSum=data.trOutflows.reduce((s,r)=>s+r.amount,0);
    const diff=trInSum-trOutSum;
    const trPage=data.trInflows.length>0||data.trOutflows.length>0?`
    <div class="page">
      <div class="banner"><div class="title">${L("titleMain")}</div>${logoBlock}<div class="date">${data.date}</div></div>
      <div class="subHdr">${L("realized")} — ${L("trReportName")}</div>
      <table>
        <tr><th colspan="3">${L("trIn")}</th><td style="border:0 !important;"></td><th colspan="3">${L("trOut")}</th></tr>
        <tr><th class="left">${L("counterparty")}</th><th class="left">${L("explanation")}</th><th>${L("amountTL")}</th><td style="border:0 !important;"></td><th class="left">${L("counterparty")}</th><th class="left">${L("explanation")}</th><th>${L("amountTL")}</th></tr>
        ${trRows}
        <tr class="total"><td class="left" colspan="2">${L("totalTL")}</td><td class="num">${fmtN(trInSum)}</td><td style="border:0 !important;"></td><td class="left" colspan="2">${L("totalPayTL")}</td><td class="num">${fmtN(trOutSum)}</td></tr>
        ${diff!==0?`<tr><td colspan="6" class="left">${diff>0?L("depositAdded"):L("depositUsed")}</td><td class="num"><b>${fmtN(Math.abs(diff))}</b></td></tr>`:""}
        <tr class="total"><td class="left" colspan="2">${L("grandTL")}</td><td class="num"><b>${fmtN(Math.max(trInSum,trOutSum))}</b></td><td style="border:0 !important;"></td><td class="left" colspan="2">${L("grandTL")}</td><td class="num"><b>${fmtN(Math.max(trInSum,trOutSum))}</b></td></tr>
      </table>
    </div>`:"";

    return `<!DOCTYPE html><html><head><title>${L("titleMain")} ${data.date}</title><style>${css}</style></head><body>
      ${ukPage}${usPage}${trPage}
    </body></html>`;
  };

  // ═══ EXCEL EXPORT ═══
  const exportToExcel=async()=>{
    const XLSX=await ensureSheetJS();
    const wb=XLSX.utils.book_new();
    const dateSheet=data.date.replace(/\//g,".");
    // UK sheet — flat AOA
    const ukRows:any[][]=[];
    ukRows.push([L("titleMain"),"","","","","","","","","","","",data.date]);
    ukRows.push([]);
    ukRows.push([L("ukReportName")]);
    ukRows.push([L("sources"),"","","","","","",L("total"),L("uses"),"","","",L("total")]);
    const mx=Math.max(data.ukSources.length,data.ukUses.length);
    for(let i=0;i<mx;i++){
      const s=data.ukSources[i],u=data.ukUses[i];
      ukRows.push([s?s.desc:"","","","","","","",s?s.amount:"",u?u.desc:"","","","",u?u.amount:""]);
    }
    ukRows.push([L("total"),"","","","","","",sumFlow(data.ukSources),L("total"),"","","",sumFlow(data.ukUses)]);
    ukRows.push([]);
    ukRows.push([L("bankName"),L("shortLoan"),"","","",L("shortLoanT"),"",L("midLoan"),"",L("midLoanT"),L("grandUSD")]);
    ukRows.push(["",L("tl"),L("ihrUSD"),L("ihrEUR"),L("usd"),"","",L("usd"),L("eur"),"",""]);
    data.loans.forEach(l=>{
      ukRows.push([l.bank,l.tl,l.ihrUSD,l.ihrEUR,l.shortUSD,l.shortUSD,"",l.midUSD,l.midEUR,l.midTotalUSD,l.totalUSD]);
    });
    ukRows.push([L("sumRow"),loanTot("tl"),loanTot("ihrUSD"),loanTot("ihrEUR"),loanTot("shortUSD"),loanTot("shortUSD"),"",loanTot("midUSD"),loanTot("midEUR"),loanTot("midTotalUSD"),loanTot("totalUSD")]);
    const ws1=XLSX.utils.aoa_to_sheet(ukRows);
    ws1["!cols"]=Array(14).fill({wch:20});
    ws1["!cols"][0]={wch:45};
    XLSX.utils.book_append_sheet(wb,ws1,dateSheet);
    // US sheet
    if(data.usSources.length>0||data.usUses.length>0){
      const usRows:any[][]=[[L("titleMain"),"","","","","","","","","","","",data.date],[],[L("usReportName")],[L("sources"),"","","","","","",L("total"),L("uses"),"","","",L("total")]];
      const mxu=Math.max(data.usSources.length,data.usUses.length);
      for(let i=0;i<mxu;i++){
        const s=data.usSources[i],u=data.usUses[i];
        usRows.push([s?s.desc:"","","","","","","",s?s.amount:"",u?u.desc:"","","","",u?u.amount:""]);
      }
      usRows.push([L("total"),"","","","","","",sumFlow(data.usSources),L("total"),"","","",sumFlow(data.usUses)]);
      const ws2=XLSX.utils.aoa_to_sheet(usRows);
      ws2["!cols"]=Array(14).fill({wch:20});ws2["!cols"][0]={wch:45};
      XLSX.utils.book_append_sheet(wb,ws2,dateSheet+" US");
    }
    // TR sheet
    if(data.trInflows.length>0||data.trOutflows.length>0){
      const trRows:any[][]=[[L("titleMain"),"","","","","","",data.date],[],[L("realized")],[L("trReportName")],[],[L("trIn"),"","","",L("trOut")],[L("counterparty"),L("explanation"),L("amountTL"),"",L("counterparty"),L("explanation"),L("amountTL")]];
      const mxt=Math.max(data.trInflows.length,data.trOutflows.length);
      for(let i=0;i<mxt;i++){
        const inf=data.trInflows[i],out=data.trOutflows[i];
        trRows.push([inf?inf.counterparty:"",inf?inf.description:"",inf?inf.amount:"","",out?out.counterparty:"",out?out.description:"",out?out.amount:""]);
      }
      const inSum=data.trInflows.reduce((s,r)=>s+r.amount,0),outSum=data.trOutflows.reduce((s,r)=>s+r.amount,0);
      trRows.push([L("totalTL"),"",inSum,"",L("totalPayTL"),"",outSum]);
      const ws3=XLSX.utils.aoa_to_sheet(trRows);
      ws3["!cols"]=[{wch:40},{wch:40},{wch:18},{wch:4},{wch:40},{wch:40},{wch:18}];
      XLSX.utils.book_append_sheet(wb,ws3,dateSheet+" TR");
    }
    const fname=`Soda_${data.lang==="tr"?"Gunluk_Rapor":"Daily_Report"}_${dateSheet}.xlsx`;
    XLSX.writeFile(wb,fname);
  };

  // TR flow CRUD
  const addTRFlow=(kind:"in"|"out")=>{const f=kind==="in"?"trInflows":"trOutflows";save({...data,[f]:[...(data[f] as DRTRFlow[]),{id:"tr_"+Date.now()+"_"+Math.random().toString(36).slice(2,6),section:"",counterparty:"",description:"",amount:0}]});};
  const updTRFlow=(kind:"in"|"out",id:string,fi:keyof DRTRFlow,v:any)=>{const f=kind==="in"?"trInflows":"trOutflows";save({...data,[f]:(data[f] as DRTRFlow[]).map(r=>r.id===id?{...r,[fi]:v}:r)});};
  const delTRFlow=(kind:"in"|"out",id:string)=>{const f=kind==="in"?"trInflows":"trOutflows";save({...data,[f]:(data[f] as DRTRFlow[]).filter(r=>r.id!==id)});};

  // Styles
  const cellSt="px-2 py-1 border border-slate-200 bg-white";
  const hdrSt="px-2 py-1.5 border border-slate-300 bg-slate-100 text-[10px] font-semibold text-slate-700 text-center tracking-wide";
  const sumRowSt="bg-amber-50 font-semibold border-t border-slate-300";
  const numIn="w-full text-right outline-none bg-transparent";
  const txtIn="w-full outline-none bg-transparent";
  const inpVal=(n:number)=>n===0?"":fmt(n);

  // Render a flow row with drag-and-drop
  const flowRow=(p:"uk"|"us",t:"sources"|"uses",r:DRFlow,i:number)=>{
    const listKey=p+"_"+t;
    const isDragging=dragItem?.list===listKey&&dragItem?.id===r.id;
    const rowColor=r.color||"";
    const textColor=rowColor||"#1f2937";
    return<tr key={r.id} draggable onDragStart={onDragStart(listKey,r.id)} onDragOver={onDragOverRow} onDrop={onDropRow(listKey,r.id)} className={"group cursor-move "+(isDragging?"opacity-40":"hover:bg-slate-50/60")}>
      <td className={cellSt+" w-6 text-center text-slate-300 group-hover:text-slate-500 select-none"} title="Sürükle-bırak ile sırala">⋮⋮</td>
      <td className={cellSt}>
        <div className="flex items-center gap-1">
          <input className={txtIn+" text-[11px]"} style={{color:textColor,fontWeight:rowColor?600:400}} value={r.desc} onChange={e=>updFlow(p,t,r.id,"desc",e.target.value)} placeholder="—"/>
          {/* Color palette dropdown (only shows on hover for cleanliness) */}
          <div className="relative opacity-0 group-hover:opacity-100 transition-opacity">
            <select value={rowColor} onChange={e=>updFlow(p,t,r.id,"color",e.target.value||undefined)} className="w-5 h-5 rounded cursor-pointer border border-slate-200 bg-white text-[9px] appearance-none text-center outline-none" title="Yazı rengini değiştir" style={{color:textColor,fontWeight:"bold"}}>
              {DR_COLORS.map(c=><option key={c.label} value={c.value} style={{color:c.value||"#1f2937"}}>{c.label}</option>)}
            </select>
            <span className="absolute -top-0.5 -right-0.5 w-1.5 h-1.5 rounded-full pointer-events-none" style={{background:rowColor||"#94a3b8"}}></span>
          </div>
          <button onClick={()=>delFlow(p,t,r.id)} className="text-slate-300 hover:text-rose-500 text-[12px]">×</button>
        </div>
      </td>
      <td className={cellSt+" text-right"} style={{minWidth:120}}><input className={numIn} style={{...numCal,fontSize:"12px",color:textColor,fontWeight:rowColor?600:400}} value={inpVal(r.amount)} onChange={e=>updFlow(p,t,r.id,"amount",parseFloat(stripC(e.target.value))||0)}/></td>
    </tr>;
  };

  // Active data based on page
  const activeSources=activePage==="uk"?data.ukSources:(activePage==="us"?data.usSources:[]);
  const activeUses=activePage==="uk"?data.ukUses:(activePage==="us"?data.usUses:[]);
  const activeTotalSrc=sumFlow(activeSources);
  const activeTotalUse=sumFlow(activeUses);

  return<div className="space-y-4 relative" style={{fontFamily:"'Inter',system-ui,sans-serif"}} onDragOver={onFileDragOver} onDragLeave={onFileDragLeave} onDrop={onFileDrop}>
    {isDraggingFile&&<div className="fixed inset-0 z-50 bg-cyan-500/20 border-4 border-dashed border-cyan-500 flex items-center justify-center pointer-events-none"><div className="bg-white px-6 py-4 rounded-xl shadow-xl border border-cyan-200"><div className="text-xl font-bold text-cyan-700 mb-1">📥 Kyriba dosyasını bırakın</div><div className="text-[11px] text-slate-500">Excel Girişler/Çıkışlar otomatik eklenecek</div></div></div>}
    {dropStatus&&<div className={"mb-3 px-3 py-2 rounded-lg text-xs font-semibold "+(dropStatus.startsWith("✓")?"bg-emerald-50 text-emerald-700 border border-emerald-200":dropStatus.startsWith("Hata")||dropStatus.startsWith("Kyriba formatı")?"bg-red-50 text-red-700 border border-red-200":"bg-cyan-50 text-cyan-700 border border-cyan-200")}>{dropStatus}</div>}

    {/* Header */}
    <div className="flex items-center justify-between flex-wrap gap-3 pb-3 border-b border-slate-100">
      <div>
        <h2 className="text-[20px] font-semibold text-slate-800 tracking-tight">Daily Report</h2>
        <p className="text-[11px] text-slate-500 mt-0.5">Soda Grubu Özet Bilgileri — 3 sayfa: UK Kasa · US Kasa · TR Nakit Akışı</p>
      </div>
      <div className="flex items-center gap-2 flex-wrap">
        <div className="flex border border-slate-200 rounded-lg overflow-hidden bg-white shadow-sm">
          <button onClick={()=>save({...data,lang:"tr"})} className={"px-3 py-1.5 text-[11px] font-medium "+(data.lang==="tr"?"bg-cyan-50 text-cyan-700":"text-slate-500 hover:bg-slate-50")}>TR</button>
          <button onClick={()=>save({...data,lang:"en"})} className={"px-3 py-1.5 text-[11px] font-medium border-l border-slate-200 "+(data.lang==="en"?"bg-cyan-50 text-cyan-700":"text-slate-500 hover:bg-slate-50")}>EN</button>
        </div>
        <div className="flex items-center gap-1.5">
          <span className="text-[10px] text-slate-400 uppercase tracking-wider">Tarih</span>
          <input value={data.date} onChange={e=>save({...data,date:e.target.value})} className="text-[12px] font-medium px-2.5 py-1 rounded-md border border-slate-200 outline-none text-slate-700 focus:border-cyan-400 bg-white w-[110px] text-center" style={numCal} placeholder="DD/MM/YYYY"/>
        </div>
        <label className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-cyan-50 hover:bg-cyan-100 text-cyan-700 rounded-lg text-[11px] font-medium border border-cyan-200 shadow-sm cursor-pointer transition-colors">
          📥 Kyriba Yükle
          <input type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)handleKyribaFile(f);e.target.value="";}}/>
        </label>
        <button onClick={exportToExcel} className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded-lg text-[11px] font-medium border border-emerald-200 shadow-sm transition-colors">📊 Excel</button>
        <button onClick={exportToPDF} className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-lg text-[11px] font-medium border border-rose-200 shadow-sm transition-colors">📄 PDF</button>
      </div>
    </div>

    {/* Page tabs */}
    <div className="flex gap-6 border-b border-slate-200 px-2">
      {[
        {k:"uk",label:"UK Günlük Kasa Raporu",color:"#0891b2"},
        {k:"us",label:"US Günlük Kasa Raporu",color:"#059669"},
        {k:"tr",label:"TR Günlük Nakit Akışı",color:"#d97706"},
      ].map(t=><button key={t.k} onClick={()=>setActivePage(t.k as DRPage)} className={"px-1 py-2.5 text-[12px] font-medium border-b-2 transition -mb-px "+(activePage===t.k?"text-slate-800":"border-transparent text-slate-500 hover:text-slate-700")} style={activePage===t.k?{borderBottomColor:t.color,color:t.color}:{}}>{t.label}</button>)}
    </div>

    {/* UK PAGE */}
    {activePage==="uk"&&<div className="space-y-4">
      {/* Banner */}
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
        <div className="flex items-center justify-between px-5 py-3 bg-white border-b border-slate-100">
          <div className="flex items-center gap-2.5">
            <div className="w-1 h-8 bg-cyan-400 rounded-full"></div>
            <div>
              <div className="text-[10px] text-slate-400 uppercase tracking-wider font-medium">{L("titleMain")}</div>
              <div className="text-[13px] font-semibold text-slate-800 leading-tight">{L("ukReportName")}</div>
            </div>
          </div>
          <div className="text-right"><div className="text-[11px] text-slate-500">{data.date}</div></div>
        </div>

        {/* Sources/Uses */}
        <div className="p-4">
          <table className="w-full border-collapse" style={{minWidth:1100}}>
            <thead>
              <tr>
                <th className={hdrSt+" bg-emerald-100 text-emerald-800"} style={{width:24}}></th>
                <th className={hdrSt+" bg-emerald-100 text-emerald-800 text-left"}>{L("sources")}</th>
                <th className={hdrSt+" bg-emerald-100 text-emerald-800"}>{L("total")}</th>
                <th className="w-3"></th>
                <th className={hdrSt+" bg-rose-100 text-rose-800"} style={{width:24}}></th>
                <th className={hdrSt+" bg-rose-100 text-rose-800 text-left"}>{L("uses")}</th>
                <th className={hdrSt+" bg-rose-100 text-rose-800"}>{L("total")}</th>
              </tr>
            </thead>
            <tbody>
              {Array.from({length:Math.max(data.ukSources.length,data.ukUses.length,1)}).map((_,i)=>{
                const s=data.ukSources[i], u=data.ukUses[i];
                return<tr key={i}>
                  {s?<>{flowRow("uk","sources",s,i).props.children}</>:<><td className={cellSt+" w-6"}></td><td className={cellSt}></td><td className={cellSt}></td></>}
                  <td className="w-3"></td>
                  {u?<>{flowRow("uk","uses",u,i).props.children}</>:<><td className={cellSt+" w-6"}></td><td className={cellSt}></td><td className={cellSt}></td></>}
                </tr>;
              })}
              <tr className={sumRowSt}>
                <td colSpan={2} className="px-2 py-1.5 border border-slate-300 text-left text-[11px]">{L("total")}</td>
                <td className="px-2 py-1.5 border border-slate-300 text-right text-emerald-800" style={{...numCal,fontSize:"13px",fontWeight:600}}>{fmt(sumFlow(data.ukSources))}</td>
                <td className="w-3"></td>
                <td colSpan={2} className="px-2 py-1.5 border border-slate-300 text-left text-[11px]">{L("total")}</td>
                <td className="px-2 py-1.5 border border-slate-300 text-right text-rose-800" style={{...numCal,fontSize:"13px",fontWeight:600}}>{fmt(sumFlow(data.ukUses))}</td>
              </tr>
            </tbody>
          </table>
          <div className="flex gap-2 mt-2">
            <button onClick={()=>addFlow("uk","sources")} className="px-2 py-0.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded text-[10px] font-medium border border-emerald-200">+ Giriş</button>
            <button onClick={()=>addFlow("uk","uses")} className="px-2 py-0.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded text-[10px] font-medium border border-rose-200">+ Çıkış</button>
            <span className="ml-auto text-[9px] text-slate-400 self-center">⋮⋮ Satırları sürükleyip bırakarak yeniden sıralayabilirsiniz</span>
          </div>
        </div>

        {/* Loans */}
        <div className="px-4 pb-4">
          <div className="text-[10px] font-semibold text-slate-600 uppercase tracking-wider mb-1">{L("bankName")} — Krediler</div>
          <table className="w-full border-collapse" style={{minWidth:1100}}>
            <thead>
              <tr>
                <th className={hdrSt+" text-left"} rowSpan={2} style={{minWidth:180}}>{L("bankName")}</th>
                <th className={hdrSt+" bg-orange-100"} colSpan={4}>{L("shortLoan")}</th>
                <th className={hdrSt+" bg-orange-100"} rowSpan={2} style={{minWidth:110}}>{L("shortLoanT")}</th>
                <th className={hdrSt+" bg-blue-100"} colSpan={2}>{L("midLoan")}</th>
                <th className={hdrSt+" bg-blue-100"} rowSpan={2} style={{minWidth:110}}>{L("midLoanT")}</th>
                <th className={hdrSt+" bg-yellow-100"} rowSpan={2} style={{minWidth:115}}>{L("grandUSD")}</th>
                <th className="w-6"></th>
              </tr>
              <tr>
                <th className={hdrSt}>{L("tl")}</th><th className={hdrSt}>{L("ihrUSD")}</th><th className={hdrSt}>{L("ihrEUR")}</th><th className={hdrSt}>{L("usd")}</th>
                <th className={hdrSt}>{L("usd")}</th><th className={hdrSt}>{L("eur")}</th>
                <th className="w-6"></th>
              </tr>
            </thead>
            <tbody>
              {data.loans.map(l=><tr key={l.id}>
                <td className={cellSt}><input className={txtIn+" text-[11px]"} value={l.bank} onChange={e=>crudUpd("loans",l.id,"bank",e.target.value)}/></td>
                {(["tl","ihrUSD","ihrEUR","shortUSD"] as (keyof DRLoan)[]).map(f=><td key={f} className={cellSt+" text-right"}><input className={numIn} style={{...numCal,fontSize:"12px"}} value={inpVal(l[f] as number)} onChange={e=>crudUpd("loans",l.id,f,parseFloat(stripC(e.target.value))||0)}/></td>)}
                <td className={cellSt+" text-right bg-orange-50/50 font-semibold"} style={{...numCal,fontSize:"12px"}}>{fmt(l.shortUSD)}</td>
                {(["midUSD","midEUR"] as (keyof DRLoan)[]).map(f=><td key={f} className={cellSt+" text-right"}><input className={numIn} style={{...numCal,fontSize:"12px"}} value={inpVal(l[f] as number)} onChange={e=>crudUpd("loans",l.id,f,parseFloat(stripC(e.target.value))||0)}/></td>)}
                <td className={cellSt+" text-right bg-blue-50/50"}><input className={numIn+" font-semibold"} style={{...numCal,fontSize:"12px"}} value={inpVal(l.midTotalUSD)} onChange={e=>crudUpd("loans",l.id,"midTotalUSD",parseFloat(stripC(e.target.value))||0)}/></td>
                <td className={cellSt+" text-right bg-yellow-50/50"}><input className={numIn+" font-bold"} style={{...numCal,fontSize:"13px"}} value={inpVal(l.totalUSD)} onChange={e=>crudUpd("loans",l.id,"totalUSD",parseFloat(stripC(e.target.value))||0)}/></td>
                <td className="px-1"><button onClick={()=>crudDel("loans",l.id)} className="text-slate-300 hover:text-rose-500 text-[12px]">×</button></td>
              </tr>)}
              <tr className={sumRowSt}>
                <td className="px-2 py-1.5 border border-slate-300 text-left text-[11px]">{L("sumRow")}</td>
                {(["tl","ihrUSD","ihrEUR","shortUSD"] as (keyof DRLoan)[]).map(f=><td key={f} className="px-2 py-1.5 border border-slate-300 text-right" style={{...numCal,fontSize:"12px"}}>{fmt(loanTot(f))}</td>)}
                <td className="px-2 py-1.5 border border-slate-300 text-right bg-orange-100 font-bold" style={{...numCal,fontSize:"13px"}}>{fmt(loanTot("shortUSD"))}</td>
                {(["midUSD","midEUR"] as (keyof DRLoan)[]).map(f=><td key={f} className="px-2 py-1.5 border border-slate-300 text-right" style={{...numCal,fontSize:"12px"}}>{fmt(loanTot(f))}</td>)}
                <td className="px-2 py-1.5 border border-slate-300 text-right bg-blue-100 font-bold" style={{...numCal,fontSize:"13px"}}>{fmt(loanTot("midTotalUSD"))}</td>
                <td className="px-2 py-1.5 border border-slate-300 text-right bg-yellow-200 font-bold" style={{...numCal,fontSize:"14px"}}>{fmt(loanTot("totalUSD"))}</td>
                <td></td>
              </tr>
            </tbody>
          </table>
          <button onClick={()=>crudAdd<DRLoan>("loans",{id:"l_"+Date.now(),bank:"",tl:0,ihrUSD:0,ihrEUR:0,shortUSD:0,midUSD:0,midEUR:0,midTotalUSD:0,totalUSD:0})} className="mt-1 px-2 py-0.5 bg-slate-50 hover:bg-slate-100 text-slate-700 rounded text-[10px] font-medium border border-slate-200">+ Kredi</button>
        </div>

        {/* Discounted Exports + Deposits */}
        <div className="grid grid-cols-2 gap-4 p-4 border-t border-slate-100">
          <div>
            <div className="text-[10px] font-semibold text-slate-600 uppercase tracking-wider mb-1">{L("discExp")}</div>
            <table className="w-full border-collapse">
              <thead>
                <tr>
                  <th className={hdrSt+" text-left"} rowSpan={2} style={{minWidth:180}}>{L("bankName")}</th>
                  <th className={hdrSt+" bg-violet-100"} colSpan={3}>{L("discExp")}</th>
                  <th className={hdrSt+" bg-yellow-100"} rowSpan={2}>{L("total")}</th>
                  <th className="w-6"></th>
                </tr>
                <tr><th className={hdrSt}>{L("usd")}</th><th className={hdrSt}>{L("gbp")}</th><th className={hdrSt}>{L("eur")}</th><th className="w-6"></th></tr>
              </thead>
              <tbody>
                {data.discounted.map(d=><tr key={d.id}>
                  <td className={cellSt}><input className={txtIn+" text-[11px]"} value={d.bank} onChange={e=>crudUpd("discounted",d.id,"bank",e.target.value)}/></td>
                  {(["usd","gbp","eur"] as (keyof DRDisc)[]).map(f=><td key={f} className={cellSt+" text-right"}><input className={numIn} style={{...numCal,fontSize:"12px"}} value={inpVal(d[f] as number)} onChange={e=>crudUpd("discounted",d.id,f,parseFloat(stripC(e.target.value))||0)}/></td>)}
                  <td className={cellSt+" text-right bg-yellow-50/50"}><input className={numIn+" font-semibold"} style={{...numCal,fontSize:"12px"}} value={inpVal(d.totalUSD)} onChange={e=>crudUpd("discounted",d.id,"totalUSD",parseFloat(stripC(e.target.value))||0)}/></td>
                  <td className="px-1"><button onClick={()=>crudDel("discounted",d.id)} className="text-slate-300 hover:text-rose-500 text-[12px]">×</button></td>
                </tr>)}
                <tr className={sumRowSt}>
                  <td className="px-2 py-1.5 border border-slate-300 text-left text-[11px]">{L("sumRow")}</td>
                  {(["usd","gbp","eur"] as (keyof DRDisc)[]).map(f=><td key={f} className="px-2 py-1.5 border border-slate-300 text-right" style={{...numCal,fontSize:"12px"}}>{fmt(discTot(f))}</td>)}
                  <td className="px-2 py-1.5 border border-slate-300 text-right bg-yellow-200 font-bold" style={{...numCal,fontSize:"13px"}}>{fmt(discTot("totalUSD"))}</td>
                  <td></td>
                </tr>
              </tbody>
            </table>
            <button onClick={()=>crudAdd<DRDisc>("discounted",{id:"d_"+Date.now(),bank:"",usd:0,eur:0,gbp:0,totalUSD:0})} className="mt-1 px-2 py-0.5 bg-slate-50 hover:bg-slate-100 text-slate-700 rounded text-[10px] font-medium border border-slate-200">+ İskonto</button>
          </div>
          <div>
            <div className="text-[10px] font-semibold text-slate-600 uppercase tracking-wider mb-1">{L("depStatus")}</div>
            <table className="w-full border-collapse">
              <thead>
                <tr><th className={hdrSt+" bg-emerald-100 text-left"} style={{minWidth:130}}>{L("company")}</th><th className={hdrSt}>{L("tlDep")}</th><th className={hdrSt}>{L("usdDep")}</th><th className={hdrSt}>{L("eurDep")}</th><th className={hdrSt}>{L("gbpDep")}</th><th className={hdrSt+" bg-yellow-100"}>{L("total")}</th><th className="w-6"></th></tr>
              </thead>
              <tbody>
                {[["depositsTR",L("trDeposit"),"tr"] as const,["depositsUK",L("ukDeposit"),"uk"] as const,["depositsUS",L("usDeposit"),"us"] as const].map(([field,label,grp])=><React.Fragment key={field}>
                  <tr><td colSpan={7} className="px-2 py-1 bg-slate-100 text-[10px] font-semibold text-slate-700 border border-slate-300">{label}</td></tr>
                  {(data[field] as DRDeposit[]).map(d=><tr key={d.id}>
                    <td className={cellSt}><input className={txtIn+" text-[11px]"} value={d.company} onChange={e=>crudUpd(field,d.id,"company",e.target.value)}/></td>
                    {(["tl","usd","eur","gbp"] as (keyof DRDeposit)[]).map(f=><td key={f} className={cellSt+" text-right"}><input className={numIn} style={{...numCal,fontSize:"11px"}} value={inpVal(d[f] as number)} onChange={e=>crudUpd(field,d.id,f,parseFloat(stripC(e.target.value))||0)}/></td>)}
                    <td className={cellSt+" text-right bg-yellow-50/50"}><input className={numIn+" font-semibold"} style={{...numCal,fontSize:"12px"}} value={inpVal(d.totalUSD)} onChange={e=>crudUpd(field,d.id,"totalUSD",parseFloat(stripC(e.target.value))||0)}/></td>
                    <td className="px-1"><button onClick={()=>crudDel(field,d.id)} className="text-slate-300 hover:text-rose-500 text-[12px]">×</button></td>
                  </tr>)}
                  <tr><td colSpan={7} className="px-2 py-0.5 bg-white"><button onClick={()=>crudAdd<DRDeposit>(field,{id:"dep_"+Date.now(),group:grp,company:"",tl:0,usd:0,eur:0,gbp:0,totalUSD:0})} className="text-[10px] text-slate-500 hover:text-slate-700">+ {label} Ekle</button></td></tr>
                </React.Fragment>)}
                <tr className={sumRowSt}>
                  <td className="px-2 py-1.5 border border-slate-300 text-left text-[11px]">{L("grandTotal")}</td>
                  {(["tl","usd","eur","gbp"] as (keyof DRDeposit)[]).map(f=><td key={f} className="px-2 py-1.5 border border-slate-300 text-right" style={{...numCal,fontSize:"12px"}}>{fmt(depTot([...data.depositsTR,...data.depositsUK,...data.depositsUS],f))}</td>)}
                  <td className="px-2 py-1.5 border border-slate-300 text-right bg-yellow-200 font-bold" style={{...numCal,fontSize:"13px"}}>{fmt(depTot([...data.depositsTR,...data.depositsUK,...data.depositsUS],"totalUSD"))}</td>
                  <td></td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* AR + Limits + FX */}
        <div className="grid grid-cols-3 gap-4 p-4 border-t border-slate-100">
          {([["trReceivables","trPayables",L("trReceivables"),L("trPayables")],["usReceivables","usPayables",L("usReceivables"),L("usPayables")]] as const).map(([recF,payF,recL,payL])=>{
            const rec=(data as any)[recF] as DRAR[];const pay=(data as any)[payF] as DRAR[];
            return<div key={recF}>
              <table className="w-full border-collapse">
                <thead><tr><th className={hdrSt+" bg-emerald-100 text-emerald-800 text-left"} colSpan={2}>{recL}</th><th style={{width:6}}></th><th className={hdrSt+" bg-rose-100 text-rose-800 text-left"} colSpan={2}>{payL}</th></tr>
                  <tr><th className={hdrSt+" text-left"}>{L("exportsRow")}</th><th className={hdrSt}>{L("balance")}</th><th></th><th className={hdrSt+" text-left"}>{L("suppliers")}</th><th className={hdrSt}>{L("balance")}</th></tr>
                </thead>
                <tbody>
                  {Array.from({length:Math.max(rec.length,pay.length,1)}).map((_,i)=>{
                    const r=rec[i],p=pay[i];
                    return<tr key={i}>
                      <td className={cellSt}>{r?<div className="flex items-center gap-1"><input className={txtIn+" text-[11px]"} value={r.name} onChange={e=>crudUpd(recF,r.id,"name",e.target.value)}/><button onClick={()=>crudDel(recF,r.id)} className="text-slate-300 hover:text-rose-500 text-[12px]">×</button></div>:""}</td>
                      <td className={cellSt+" text-right"}>{r?<input className={numIn} style={{...numCal,fontSize:"12px"}} value={inpVal(r.amount)} onChange={e=>crudUpd(recF,r.id,"amount",parseFloat(stripC(e.target.value))||0)}/>:""}</td>
                      <td></td>
                      <td className={cellSt}>{p?<div className="flex items-center gap-1"><input className={txtIn+" text-[11px]"} value={p.name} onChange={e=>crudUpd(payF,p.id,"name",e.target.value)}/><button onClick={()=>crudDel(payF,p.id)} className="text-slate-300 hover:text-rose-500 text-[12px]">×</button></div>:""}</td>
                      <td className={cellSt+" text-right"}>{p?<input className={numIn} style={{...numCal,fontSize:"12px"}} value={inpVal(p.amount)} onChange={e=>crudUpd(payF,p.id,"amount",parseFloat(stripC(e.target.value))||0)}/>:""}</td>
                    </tr>;
                  })}
                  <tr className={sumRowSt}>
                    <td className="px-2 py-1.5 border border-slate-300 text-left text-[10px]">{L("total")}</td>
                    <td className="px-2 py-1.5 border border-slate-300 text-right" style={{...numCal,fontSize:"12px"}}>{fmt(rec.reduce((s,r)=>s+r.amount,0))}</td>
                    <td></td>
                    <td className="px-2 py-1.5 border border-slate-300 text-left text-[10px]">{L("total")}</td>
                    <td className="px-2 py-1.5 border border-slate-300 text-right" style={{...numCal,fontSize:"12px"}}>{fmt(pay.reduce((s,p)=>s+p.amount,0))}</td>
                  </tr>
                </tbody>
              </table>
              <div className="flex gap-1 mt-1">
                <button onClick={()=>crudAdd<DRAR>(recF,{id:"ar_"+Date.now(),name:"",amount:0})} className="px-2 py-0.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded text-[10px] font-medium border border-emerald-200">+ Alacak</button>
                <button onClick={()=>crudAdd<DRAR>(payF,{id:"ap_"+Date.now(),name:"",amount:0})} className="px-2 py-0.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded text-[10px] font-medium border border-rose-200">+ Borç</button>
              </div>
            </div>;
          })}

          <div className="space-y-3">
            <table className="w-full border-collapse">
              <thead>
                <tr><th className={hdrSt+" bg-cyan-100 text-cyan-800"} colSpan={4}>{L("limits")}</th></tr>
                <tr><th className={hdrSt+" text-left"} style={{minWidth:120}}>{L("bankName")}</th><th className={hdrSt}>{L("usd")}</th><th className={hdrSt}>{L("tl")}</th><th className={hdrSt+" bg-yellow-100"}>{L("total")}</th></tr>
              </thead>
              <tbody>
                {data.limits.map(lm=><tr key={lm.id}>
                  <td className={cellSt}><div className="flex items-center gap-1"><input className={txtIn+" text-[11px]"} value={lm.bank} onChange={e=>crudUpd("limits",lm.id,"bank",e.target.value)}/><button onClick={()=>crudDel("limits",lm.id)} className="text-slate-300 hover:text-rose-500 text-[12px]">×</button></div></td>
                  {(["usd","tl"] as (keyof DRLimit)[]).map(f=><td key={f} className={cellSt+" text-right"}><input className={numIn} style={{...numCal,fontSize:"12px"}} value={inpVal(lm[f] as number)} onChange={e=>crudUpd("limits",lm.id,f,parseFloat(stripC(e.target.value))||0)}/></td>)}
                  <td className={cellSt+" text-right bg-yellow-50/50"}><input className={numIn+" font-semibold"} style={{...numCal,fontSize:"12px"}} value={inpVal(lm.totalUSD)} onChange={e=>crudUpd("limits",lm.id,"totalUSD",parseFloat(stripC(e.target.value))||0)}/></td>
                </tr>)}
                <tr className={sumRowSt}>
                  <td className="px-2 py-1.5 border border-slate-300 text-left text-[11px]">{L("sumRow")}</td>
                  <td className="px-2 py-1.5 border border-slate-300 text-right" style={{...numCal,fontSize:"12px"}}>{fmt(data.limits.reduce((s,l)=>s+l.usd,0))}</td>
                  <td className="px-2 py-1.5 border border-slate-300 text-right" style={{...numCal,fontSize:"12px"}}>{fmt(data.limits.reduce((s,l)=>s+l.tl,0))}</td>
                  <td className="px-2 py-1.5 border border-slate-300 text-right bg-yellow-200 font-bold" style={{...numCal,fontSize:"13px"}}>{fmt(data.limits.reduce((s,l)=>s+l.totalUSD,0))}</td>
                </tr>
              </tbody>
            </table>
            <button onClick={()=>crudAdd<DRLimit>("limits",{id:"lm_"+Date.now(),bank:"",usd:0,tl:0,totalUSD:0})} className="px-2 py-0.5 bg-slate-50 hover:bg-slate-100 text-slate-700 rounded text-[10px] font-medium border border-slate-200">+ Limit</button>

            <table className="w-full border-collapse">
              <thead><tr><th className={hdrSt+" bg-violet-100 text-violet-800"} colSpan={2}>{L("reserve")}</th></tr></thead>
              <tbody>
                {data.reserveAccounts.map(r=><tr key={r.id}>
                  <td className={cellSt}><div className="flex items-center gap-1"><input className={txtIn+" text-[11px]"} value={r.name} onChange={e=>crudUpd("reserveAccounts",r.id,"name",e.target.value)}/><button onClick={()=>crudDel("reserveAccounts",r.id)} className="text-slate-300 hover:text-rose-500 text-[12px]">×</button></div></td>
                  <td className={cellSt+" text-right bg-yellow-50/50"}><input className={numIn+" font-semibold"} style={{...numCal,fontSize:"12px"}} value={inpVal(r.amount)} onChange={e=>crudUpd("reserveAccounts",r.id,"amount",parseFloat(stripC(e.target.value))||0)}/></td>
                </tr>)}
                <tr className={sumRowSt}>
                  <td className="px-2 py-1.5 border border-slate-300 text-left text-[11px]">{L("total")}</td>
                  <td className="px-2 py-1.5 border border-slate-300 text-right bg-yellow-200 font-bold" style={{...numCal,fontSize:"13px"}}>{fmt(data.reserveAccounts.reduce((s,r)=>s+r.amount,0))}</td>
                </tr>
              </tbody>
            </table>
            <button onClick={()=>crudAdd<DRReserveAcc>("reserveAccounts",{id:"ra_"+Date.now(),name:"",amount:0})} className="px-2 py-0.5 bg-slate-50 hover:bg-slate-100 text-slate-700 rounded text-[10px] font-medium border border-slate-200">+ Bloke</button>

            <table className="w-full border-collapse">
              <thead><tr><th className={hdrSt+" bg-violet-100 text-violet-800"} colSpan={2}>Kurlar / FX</th></tr></thead>
              <tbody>
                {([["eur_usd","EUR/USD"],["usd_try","USD/TRY"],["eur_try","EUR/TRY"],["gbp_usd","GBP/USD"]] as const).map(([k,lbl])=><tr key={k}>
                  <td className={cellSt+" text-[11px]"}>{lbl}</td>
                  <td className={cellSt+" text-right"}><input className={numIn} style={{...numCal,fontSize:"12px"}} value={(data.fx as any)[k]?.toFixed(4)||""} onChange={e=>save({...data,fx:{...data.fx,[k]:parseFloat(e.target.value)||0}})}/></td>
                </tr>)}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>}

    {/* US PAGE */}
    {activePage==="us"&&<div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
      <div className="flex items-center justify-between px-5 py-3 bg-white border-b border-slate-100">
        <div className="flex items-center gap-2.5"><div className="w-1 h-8 bg-emerald-400 rounded-full"></div><div><div className="text-[10px] text-slate-400 uppercase tracking-wider font-medium">{L("titleMain")}</div><div className="text-[13px] font-semibold text-slate-800 leading-tight">{L("usReportName")}</div></div></div>
        <div className="text-right"><div className="text-[11px] text-slate-500">{data.date}</div></div>
      </div>
      <div className="p-4">
        <table className="w-full border-collapse" style={{minWidth:1100}}>
          <thead>
            <tr><th className={hdrSt+" bg-emerald-100 text-emerald-800"} style={{width:24}}></th><th className={hdrSt+" bg-emerald-100 text-emerald-800 text-left"}>{L("sources")}</th><th className={hdrSt+" bg-emerald-100 text-emerald-800"}>{L("total")}</th><th className="w-3"></th><th className={hdrSt+" bg-rose-100 text-rose-800"} style={{width:24}}></th><th className={hdrSt+" bg-rose-100 text-rose-800 text-left"}>{L("uses")}</th><th className={hdrSt+" bg-rose-100 text-rose-800"}>{L("total")}</th></tr>
          </thead>
          <tbody>
            {Array.from({length:Math.max(data.usSources.length,data.usUses.length,1)}).map((_,i)=>{
              const s=data.usSources[i],u=data.usUses[i];
              return<tr key={i}>
                {s?flowRow("us","sources",s,i).props.children:<><td className={cellSt+" w-6"}></td><td className={cellSt}></td><td className={cellSt}></td></>}
                <td className="w-3"></td>
                {u?flowRow("us","uses",u,i).props.children:<><td className={cellSt+" w-6"}></td><td className={cellSt}></td><td className={cellSt}></td></>}
              </tr>;
            })}
            <tr className={sumRowSt}>
              <td colSpan={2} className="px-2 py-1.5 border border-slate-300 text-left text-[11px]">{L("total")}</td>
              <td className="px-2 py-1.5 border border-slate-300 text-right text-emerald-800" style={{...numCal,fontSize:"13px",fontWeight:600}}>{fmt(sumFlow(data.usSources))}</td>
              <td></td>
              <td colSpan={2} className="px-2 py-1.5 border border-slate-300 text-left text-[11px]">{L("total")}</td>
              <td className="px-2 py-1.5 border border-slate-300 text-right text-rose-800" style={{...numCal,fontSize:"13px",fontWeight:600}}>{fmt(sumFlow(data.usUses))}</td>
            </tr>
          </tbody>
        </table>
        <div className="flex gap-2 mt-2">
          <button onClick={()=>addFlow("us","sources")} className="px-2 py-0.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded text-[10px] font-medium border border-emerald-200">+ Giriş</button>
          <button onClick={()=>addFlow("us","uses")} className="px-2 py-0.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded text-[10px] font-medium border border-rose-200">+ Çıkış</button>
          <span className="ml-auto text-[9px] text-slate-400 self-center">⋮⋮ Sürükle-bırak ile sırala</span>
        </div>
      </div>
    </div>}

    {/* TR PAGE */}
    {activePage==="tr"&&<div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
      <div className="flex items-center justify-between px-5 py-3 bg-white border-b border-slate-100">
        <div className="flex items-center gap-2.5"><div className="w-1 h-8 bg-amber-400 rounded-full"></div><div><div className="text-[10px] text-slate-400 uppercase tracking-wider font-medium">{L("realized")}</div><div className="text-[13px] font-semibold text-slate-800 leading-tight">{L("trReportName")}</div></div></div>
        <div className="text-right"><div className="text-[11px] text-slate-500">{data.date}</div></div>
      </div>
      <div className="p-4">
        <table className="w-full border-collapse" style={{minWidth:1200}}>
          <thead>
            <tr>
              <th className={hdrSt+" bg-emerald-100 text-emerald-800"} colSpan={3}>{L("trIn")}</th>
              <th className="w-3"></th>
              <th className={hdrSt+" bg-rose-100 text-rose-800"} colSpan={3}>{L("trOut")}</th>
            </tr>
            <tr>
              <th className={hdrSt+" text-left"} style={{minWidth:200}}>{L("counterparty")}</th>
              <th className={hdrSt+" text-left"} style={{minWidth:220}}>{L("explanation")}</th>
              <th className={hdrSt} style={{minWidth:130}}>{L("amountTL")}</th>
              <th className="w-3"></th>
              <th className={hdrSt+" text-left"} style={{minWidth:200}}>{L("counterparty")}</th>
              <th className={hdrSt+" text-left"} style={{minWidth:220}}>{L("explanation")}</th>
              <th className={hdrSt} style={{minWidth:130}}>{L("amountTL")}</th>
            </tr>
          </thead>
          <tbody>
            {Array.from({length:Math.max(data.trInflows.length,data.trOutflows.length,1)}).map((_,i)=>{
              const inf=data.trInflows[i],out=data.trOutflows[i];
              const infColor=inf?.color||"", outColor=out?.color||"";
              const infTxt=infColor||"#1f2937", outTxt=outColor||"#1f2937";
              return<tr key={i} className="group">
                {inf?<>
                  <td className={cellSt}>
                    <div className="flex items-center gap-1">
                      <input className={txtIn+" text-[11px]"} style={{color:infTxt,fontWeight:infColor?600:400}} value={inf.counterparty} onChange={e=>updTRFlow("in",inf.id,"counterparty",e.target.value)}/>
                      <div className="relative opacity-0 group-hover:opacity-100 transition-opacity">
                        <select value={infColor} onChange={e=>updTRFlow("in",inf.id,"color",e.target.value||undefined)} className="w-5 h-5 rounded cursor-pointer border border-slate-200 bg-white text-[9px] appearance-none text-center outline-none" title="Yazı rengi" style={{color:infTxt,fontWeight:"bold"}}>
                          {DR_COLORS.map(c=><option key={c.label} value={c.value} style={{color:c.value||"#1f2937"}}>{c.label}</option>)}
                        </select>
                        <span className="absolute -top-0.5 -right-0.5 w-1.5 h-1.5 rounded-full pointer-events-none" style={{background:infColor||"#94a3b8"}}></span>
                      </div>
                      <button onClick={()=>delTRFlow("in",inf.id)} className="text-slate-300 hover:text-rose-500 text-[12px]">×</button>
                    </div>
                  </td>
                  <td className={cellSt}><input className={txtIn+" text-[11px]"} style={{color:infTxt,fontWeight:infColor?600:400}} value={inf.description} onChange={e=>updTRFlow("in",inf.id,"description",e.target.value)}/></td>
                  <td className={cellSt+" text-right"}><input className={numIn} style={{...numCal,fontSize:"12px",color:infTxt,fontWeight:infColor?600:400}} value={inpVal(inf.amount)} onChange={e=>updTRFlow("in",inf.id,"amount",parseFloat(stripC(e.target.value))||0)}/></td>
                </>:<><td className={cellSt}></td><td className={cellSt}></td><td className={cellSt}></td></>}
                <td></td>
                {out?<>
                  <td className={cellSt}>
                    <div className="flex items-center gap-1">
                      <input className={txtIn+" text-[11px]"} style={{color:outTxt,fontWeight:outColor?600:400}} value={out.counterparty} onChange={e=>updTRFlow("out",out.id,"counterparty",e.target.value)}/>
                      <div className="relative opacity-0 group-hover:opacity-100 transition-opacity">
                        <select value={outColor} onChange={e=>updTRFlow("out",out.id,"color",e.target.value||undefined)} className="w-5 h-5 rounded cursor-pointer border border-slate-200 bg-white text-[9px] appearance-none text-center outline-none" title="Yazı rengi" style={{color:outTxt,fontWeight:"bold"}}>
                          {DR_COLORS.map(c=><option key={c.label} value={c.value} style={{color:c.value||"#1f2937"}}>{c.label}</option>)}
                        </select>
                        <span className="absolute -top-0.5 -right-0.5 w-1.5 h-1.5 rounded-full pointer-events-none" style={{background:outColor||"#94a3b8"}}></span>
                      </div>
                      <button onClick={()=>delTRFlow("out",out.id)} className="text-slate-300 hover:text-rose-500 text-[12px]">×</button>
                    </div>
                  </td>
                  <td className={cellSt}><input className={txtIn+" text-[11px]"} style={{color:outTxt,fontWeight:outColor?600:400}} value={out.description} onChange={e=>updTRFlow("out",out.id,"description",e.target.value)}/></td>
                  <td className={cellSt+" text-right"}><input className={numIn} style={{...numCal,fontSize:"12px",color:outTxt,fontWeight:outColor?600:400}} value={inpVal(out.amount)} onChange={e=>updTRFlow("out",out.id,"amount",parseFloat(stripC(e.target.value))||0)}/></td>
                </>:<><td className={cellSt}></td><td className={cellSt}></td><td className={cellSt}></td></>}
              </tr>;
            })}
            <tr className={sumRowSt}>
              <td colSpan={2} className="px-2 py-1.5 border border-slate-300 text-left text-[11px]">{L("totalTL")}</td>
              <td className="px-2 py-1.5 border border-slate-300 text-right text-emerald-800" style={{...numCal,fontSize:"13px",fontWeight:600}}>{fmt(data.trInflows.reduce((s,r)=>s+r.amount,0))}</td>
              <td></td>
              <td colSpan={2} className="px-2 py-1.5 border border-slate-300 text-left text-[11px]">{L("totalPayTL")}</td>
              <td className="px-2 py-1.5 border border-slate-300 text-right text-rose-800" style={{...numCal,fontSize:"13px",fontWeight:600}}>{fmt(data.trOutflows.reduce((s,r)=>s+r.amount,0))}</td>
            </tr>
          </tbody>
        </table>
        <div className="flex gap-2 mt-2">
          <button onClick={()=>addTRFlow("in")} className="px-2 py-0.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded text-[10px] font-medium border border-emerald-200">+ Giriş</button>
          <button onClick={()=>addTRFlow("out")} className="px-2 py-0.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded text-[10px] font-medium border border-rose-200">+ Çıkış</button>
        </div>
      </div>
    </div>}

    {/* Info footer */}
    <div className="bg-sky-50/50 border border-sky-200 rounded-lg p-3 text-[10px] text-slate-600 leading-relaxed">
      <b>Özellikler:</b> ⋮⋮ Sürükle-bırak ile Girişler/Çıkışlar sırasını değiştirin. <b>📥 Kyriba Yükle</b> ile Kyriba Excel dosyasındaki işlemleri otomatik olarak Girişler/Çıkışlar kısmına ekleyin (Debit→Çıkış, Credit→Giriş; kolon başlıkları "Description/Açıklama" ve "Amount/Tutar" veya "Debit/Credit" olmalı). <b>📄 PDF</b> çıktısı 3 sayfa halinde A4 yatay formatında orijinal Excel görünümüne birebir benzer. <b>📊 Excel</b> çıktısı her sayfa için ayrı sheet üretir (DD.MM.YYYY / DD.MM.YYYY US / DD.MM.YYYY TR).
    </div>
  </div>;
}


export default function TreasuryApp(){
  const[user,setUser]=useState<typeof USERS[0]|null>(null);const[tab,setTab]=useState("dashboard");const[fx,setFx]=useState(FX_D);const[sm,setSM]=useState(new Date().getMonth());const[sy,setSY]=useState(new Date().getFullYear());
  const undoCount=useUndoStackSize();
  const chatUnread=useChatUnread(user?.username||"");
  useEffect(()=>{fetchTCMB().then(setFx);},[]);
  // Global Ctrl+Z / Cmd+Z handler — works anywhere in the app
  useEffect(()=>{
    const onKey=(e:KeyboardEvent)=>{
      // Ctrl+Z (Windows/Linux) or Cmd+Z (Mac). Skip if Shift (redo) or Ctrl+Y
      if((e.ctrlKey||e.metaKey)&&!e.shiftKey&&(e.key==="z"||e.key==="Z")){
        // If user is actively typing in a single input, let native undo handle within that input first.
        // But if the stack has entries, intercept — native input undo doesn't help with form state anyway
        // since we store values in localStorage on every change.
        const ok=performUndo();
        if(ok){e.preventDefault();}
      }
    };
    window.addEventListener("keydown",onKey);
    return()=>window.removeEventListener("keydown",onKey);
  },[]);
  if(!user)return<Login onLogin={setUser}/>;const canEdit=user.role!=="viewer";
  return(<div className="min-h-screen bg-gray-50 text-gray-800">
    <style>{`
      .num-cell{font-family:'Calibri','Segoe UI','Arial',system-ui,sans-serif;font-variant-numeric:tabular-nums;letter-spacing:0.01em;}
    `}</style>
    <div className="max-w-7xl mx-auto px-4 py-4">
    <div className="flex justify-between items-start mb-2 flex-wrap gap-2"><div className="flex items-center gap-3"><img src="/wesoda_logo.jpg" alt="We Soda" style={{height:"36px",width:"auto",maxWidth:"110px",objectFit:"contain",display:"block"}} onError={(e)=>{(e.target as HTMLImageElement).style.display="none";}}/><div><h1 className="text-lg font-bold">Soda Grubu Hazine Yönetim Sistemi</h1><p className="text-[11px] text-gray-400">Konsolide nakit yönetimi</p></div></div><div className="flex items-center gap-3"><button onClick={()=>setTab("dashboard")} className={"text-[11px] px-3 py-1 border rounded flex items-center gap-1 transition relative "+(chatUnread>0?"border-red-300 text-red-700 bg-red-50 hover:bg-red-100 animate-pulse":"border-gray-200 text-gray-500 hover:bg-gray-50")} title={chatUnread>0?chatUnread+" okunmamış mesaj — Dashboard'a git":"Takım mesajlaşması"}><span>💬</span>{chatUnread>0&&<span className="px-1.5 py-0 bg-red-500 text-white rounded-full text-[9px] font-bold">{chatUnread}</span>}</button><button onClick={()=>performUndo()} disabled={undoCount===0} className={"text-[11px] px-3 py-1 border rounded flex items-center gap-1 transition "+(undoCount===0?"border-gray-200 text-gray-300 cursor-not-allowed":"border-indigo-300 text-indigo-700 bg-indigo-50 hover:bg-indigo-100")} title={"Son değişikliği geri al (Ctrl+Z) — "+undoCount+" adım mevcut"}><span>↶ Geri Al</span>{undoCount>0&&<span className="px-1.5 py-0 bg-indigo-600 text-white rounded-full text-[9px] font-bold">{undoCount}</span>}</button><div className="text-right"><div className="text-xs font-medium">{user.name}</div><div className="text-[10px] text-gray-400">{user.role==="admin"?"Tam yetki":user.role==="analyst"?"Düzenleme":"Görüntüleme"}</div></div><button onClick={()=>setUser(null)} className="text-[11px] px-2 py-1 border border-gray-300 rounded hover:bg-gray-100">Çıkış</button></div></div>
    <div className="text-[11px] text-gray-400 mb-1"><span className="font-medium">Kurlar ({fx.source}):</span> <span className={mn}>USD/TRY {fx.usd_try.toFixed(4)}</span> | <span className={mn}>EUR/USD {fx.eur_usd.toFixed(4)}</span></div>
    {["cashflow","kasa"].includes(tab)&&<div className="flex items-center gap-2 mb-2"><select value={sm} onChange={e=>setSM(+e.target.value)} className={ddS}>{MONTHS.map((m,i)=><option key={i} value={i}>{m}</option>)}</select><select value={sy} onChange={e=>setSY(+e.target.value)} className={ddS}>{YEARS.map(y=><option key={y} value={y}>{y}</option>)}</select></div>}
    <div className="flex border-b border-gray-200 overflow-x-auto mb-4">
      <Tb active={tab==="dashboard"} color="#7c3aed" onClick={()=>setTab("dashboard")}>Dashboard</Tb>
      <Tb active={tab==="cashflow"} color="#2563eb" onClick={()=>setTab("cashflow")}>Cashflow</Tb>
      <Tb active={tab==="kasa"} color="#059669" onClick={()=>setTab("kasa")}>TR/UK Kasa</Tb>
      <Tb active={tab==="krediler"} color="#be185d" onClick={()=>setTab("krediler")}>Loans</Tb>
      <Tb active={tab==="invoices"} color="#9333ea" onClick={()=>setTab("invoices")}>Invoices</Tb>
      <Tb active={tab==="daily"} color="#0891b2" onClick={()=>setTab("daily")}>Daily Report</Tb>
    </div>
    {tab==="dashboard"&&<Dashboard fx={fx} setFx={setFx} sm={sm} sy={sy} user={user}/>}
    {tab==="cashflow"&&<CashflowWithSubs canEdit={canEdit} sm={sm} sy={sy} fx={fx}/>}
    {tab==="kasa"&&<Kasa canEdit={canEdit} sm={sm} sy={sy} fx={fx}/>}
    {tab==="krediler"&&<Krediler/>}
    {tab==="invoices"&&<InvoicesTab/>}
    {tab==="daily"&&<DailyReport fx={fx}/>}
  </div></div>);
}