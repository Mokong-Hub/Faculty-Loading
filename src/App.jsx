import { useState, useMemo, useEffect } from "react";

// ══════════════════════════════════════════════════════════════════════════════
// COLLEGES & PROGRAMS
// ══════════════════════════════════════════════════════════════════════════════
const COLLEGES = {
  CCJE: { color:"#f97316", bg:"rgba(249,115,22,0.10)",  programs:["BSCRIM"] },
  COED: { color:"#22c55e", bg:"rgba(34,197,94,0.10)",   programs:["BSED-ENG","BSED-FIL","BSED-MATH","BSED-SCI","BEED","BSNED"] },
  CTHM: { color:"#a855f7", bg:"rgba(168,85,247,0.10)",  programs:["BSHM","BSTM"] },
  CBA:  { color:"#3b82f6", bg:"rgba(59,130,246,0.10)",  programs:["BSBA-MM","BSBA-FM","BSBA MM","BSBA FM","BSA","BSAIS"] },
  CAHS: { color:"#ec4899", bg:"rgba(236,72,153,0.10)",  programs:["BSN","ABPSY","ABPSYCH","BSPHARMA"] },
  COE:  { color:"#f59e0b", bg:"rgba(245,158,11,0.10)",  programs:["BSCE","BSEE","BSME"] },
  CITE: { color:"#06b6d4", bg:"rgba(6,182,212,0.10)",   programs:["BSIT"] },
  COME: { color:"#84cc16", bg:"rgba(132,204,22,0.10)",  programs:["BSMARE"] },
};
function getCollege(prog) {
  for (const [k,d] of Object.entries(COLLEGES))
    if (d.programs.includes(prog)) return k;
  return null;
}
function clg(prog) { const k=getCollege(prog); return k?COLLEGES[k]:null; }

const SUBJECT_TITLES = {
  "ART 002":"Art Appreciation","BAM 006":"Business Admin & Mgmt",
  "ENG 189":"Technical Writing","GEN 001":"Understanding the Self",
  "GEN 002":"Readings in Phil. History","GEN 003":"The Contemporary World",
  "GEN 006":"Science, Technology & Society","GEN 010":"Gender and Society",
  "HIS 007":"Philippine History","MAT 152":"Mathematics in the Modern World",
  "NST 021":"NSTP 1","NST 023":"CWTS 1",
  "PED 025":"PE 1","PED 027":"PE 2","PED 030":"PE 3","PED 032":"PE 4",
  "PHI 002":"Ethics","PHI 020":"Logic","SCX 010":"Environmental Science",
};

// ══════════════════════════════════════════════════════════════════════════════
// NORMALIZERS
// ══════════════════════════════════════════════════════════════════════════════
function normDay(d) {
  if (!d) return "";
  d = d.trim().toUpperCase().replace(/\s+/g,"").replace(/&/g,"/");
  if (d==="TH") return "THU";
  if (["TH/FRI","THFRI","THUFRI"].includes(d)) return "THU/FRI";
  if (["TH/SAT","THSAT","THUSAT"].includes(d)) return "THU/SAT";
  if (["MONTUE","MON/TUE","MON&TUE"].includes(d)) return "MON/TUE";
  if (["FRISAT","FRI/SAT","FRI&SAT"].includes(d)) return "FRI/SAT";
  if (["TUEWED","TUE/WED","TUE&WED"].includes(d)) return "TUE/WED";
  return d;
}
function normTime(t) {
  if (!t) return "";
  const s = t.replace(/\s+/g,"").toUpperCase();
  const rules = [
    ["7:30AM-9:00AM",   /^7:?30A?M?-9:?00A?M?$/i],
    ["9:00AM-10:30AM",  /^9:?00A?M?-10:?30A?M?$/i],
    ["10:30AM-12:00NN", /^10:?30A?M?-12:?00(NN|PM|AM)?$/i],
    ["12:00NN-1:30PM",  /^12:?00(NN|PM)?-1:?30P?M?$/i],
    ["1:30PM-3:00PM",   /^(01:30PM-03:00PM|1:?30P?M?-3:?00P?M?)$/i],
    ["3:00PM-4:30PM",   /^3:?00P?M?-4:?30P?M?$/i],
    ["4:30PM-6:00PM",   /^4:?30P?M?-6:?00P?M?$/i],
    ["6:00PM-7:30PM",   /^6:?00P?M?-7:?30P?M?$/i],
  ];
  for (const [c,re] of rules) if (re.test(s)) return c;
  return t.trim();
}

// ══════════════════════════════════════════════════════════════════════════════
// TIME SLOT CONSTANTS
// ══════════════════════════════════════════════════════════════════════════════
const AM_SLOTS = ["7:30AM-9:00AM","9:00AM-10:30AM","10:30AM-12:00NN"];
const PM_SLOTS = ["1:30PM-3:00PM","3:00PM-4:30PM","4:30PM-6:00PM","6:00PM-7:30PM"];
const LUNCH    = "12:00NN-1:30PM";
const ALL_SLOTS = [...AM_SLOTS,LUNCH,...PM_SLOTS];
const WORK_DAYS = ["MON","TUE","WED","THU","FRI","SAT"];
const ALL_DAYS  = [...WORK_DAYS,"MON/TUE","TUE/WED","THU/FRI","FRI/SAT","THU/SAT"];

function slotIdx(t)  { return ALL_SLOTS.indexOf(t); }
function isAM(t)     { return AM_SLOTS.includes(t); }
function isPM(t)     { return PM_SLOTS.includes(t); }

// ══════════════════════════════════════════════════════════════════════════════
// SCHEDULING ENGINE
// ══════════════════════════════════════════════════════════════════════════════

// Rule 2: 10:30 + 1:30 on same day = lunch trapped
function wouldTrapLunch(existingSlots, candidate) {
  const all = [...existingSlots, candidate].map(slotIdx).filter(i=>i>=0).sort((a,b)=>a-b);
  for (let i=0; i<all.length-1; i++)
    if (all[i]===2 && all[i+1]===4) return true;
  return false;
}

// Rule 3: 3 consecutive slots = 4.5h straight
function wouldStack3(existingSlots, candidate) {
  const idx = [...existingSlots, candidate].map(slotIdx).filter(i=>i>=0 && i!==3).sort((a,b)=>a-b);
  let streak=1;
  for (let i=1;i<idx.length;i++) {
    streak = idx[i]===idx[i-1]+1 ? streak+1 : 1;
    if (streak>=3) return true;
  }
  return false;
}

// AM/PM half-split: first ceil(n/2) sections → AM, rest → PM
function halfSplit(posInGroup, total) {
  return posInGroup < Math.ceil(total/2) ? "AM" : "PM";
}

function autoSchedule(rows) {
  const updated = rows.map(r=>({...r}));

  // Occupancy trackers
  const facOcc  = {};
  const roomOcc = {};
  const secSlots = {}; // `id|day` → [times]

  for (const r of updated) {
    if (!r.day||!r.time) continue;
    const d=normDay(r.day), t=normTime(r.time);
    if (r.faculty&&!r.faculty.includes("PLACE")) facOcc[`${r.faculty}|${d}|${t}`]=true;
    if (r.room) roomOcc[`${r.room}|${d}|${t}`]=true;
    const sk=`${r.id}|${d}`;
    secSlots[sk]=[...(secSlots[sk]||[]),t];
  }

  // Group by code+program = parallel group
  const groups = {};
  updated.forEach((r,i)=>{
    const k=`${r.code}||${r.program}`;
    groups[k]=groups[k]||[];
    groups[k].push(i);
  });

  // Rotating cursors per subject
  const amCur={}, pmCur={}, dayCur={};

  for (const [,indices] of Object.entries(groups)) {
    const code = updated[indices[0]].code;
    if (amCur[code]===undefined) amCur[code]=0;
    if (pmCur[code]===undefined) pmCur[code]=0;
    if (dayCur[code]===undefined) dayCur[code]=0;
    const total = indices.length;

    indices.forEach((rowIdx, pos) => {
      const r = updated[rowIdx];
      if (r.day && r.time) return; // already set — honour it

      const half     = halfSplit(pos, total);
      const slotPool = half==="AM" ? AM_SLOTS : PM_SLOTS;
      const cur      = half==="AM" ? amCur    : pmCur;

      // Day order rotated by dayCur
      const ds = dayCur[code] % WORK_DAYS.length;
      const dayOrder = [...WORK_DAYS.slice(ds), ...WORK_DAYS.slice(0,ds)];

      let assigned=false;
      outer:
      for (const day of dayOrder) {
        for (let si=0; si<slotPool.length; si++) {
          const idx  = (cur[code]+si) % slotPool.length;
          const time = slotPool[idx];
          if (time===LUNCH) continue;
          if (r.faculty&&!r.faculty.includes("PLACE") && facOcc[`${r.faculty}|${day}|${time}`]) continue;
          if (r.room && roomOcc[`${r.room}|${day}|${time}`]) continue;
          const sk=`${r.id}|${day}`;
          const existing=secSlots[sk]||[];
          if (wouldTrapLunch(existing,time)) continue;
          if (wouldStack3(existing,time)) continue;

          // ✓ assign
          updated[rowIdx]={...r,day,time,_dirty:true};
          if (r.faculty&&!r.faculty.includes("PLACE")) facOcc[`${r.faculty}|${day}|${time}`]=true;
          if (r.room) roomOcc[`${r.room}|${day}|${time}`]=true;
          secSlots[sk]=[...existing,time];
          cur[code]=(idx+1)%slotPool.length;
          if (pos%2===1) dayCur[code]=(dayCur[code]+1)%WORK_DAYS.length;
          assigned=true;
          break outer;
        }
      }
    });
  }
  return updated;
}

function buildFacultyPool(rows) {
  const map={};
  for (const r of rows) {
    const f=r.faculty;
    if (!f||f.includes("PLACE")) continue;
    if (!map[f]) map[f]={name:f,specialties:new Set(),assignedDays:new Set(),maxHrs:24};
    if (r.code) map[f].specialties.add(r.code);
    if (r.day)  map[f].assignedDays.add(normDay(r.day));
  }
  return Object.values(map).map(f=>({...f,specialties:[...f.specialties],assignedDays:[...f.assignedDays]}));
}

function autoAssignFaculty(rows, pool) {
  const updated = rows.map(r=>({...r}));
  const hrs={}, occ={};
  for (const f of pool) hrs[f.name]=0;
  for (const r of updated) {
    if (r.faculty&&!r.faculty.includes("PLACE")) {
      hrs[r.faculty]=(hrs[r.faculty]||0)+(r.tempHrs||1.5);
      if (r.day&&r.time) occ[`${r.faculty}|${normDay(r.day)}|${normTime(r.time)}`]=true;
    }
  }
  for (let i=0;i<updated.length;i++) {
    const r=updated[i];
    if (r.faculty&&!r.faculty.includes("PLACE")) continue;
    const d=normDay(r.day), t=normTime(r.time);
    const pick = (candidates) => candidates
      .filter(f=>{
        if (f.specialties.length&&!f.specialties.includes(r.code)) return false;
        if ((hrs[f.name]||0)+(r.tempHrs||1.5)>(f.maxHrs||24)) return false;
        if (d&&t&&occ[`${f.name}|${d}|${t}`]) return false;
        return true;
      })
      .sort((a,b)=>(hrs[a.name]||0)-(hrs[b.name]||0));
    const dayPref = pool.filter(f=>!f.assignedDays.length||(d&&f.assignedDays.includes(d)));
    const best = pick(dayPref)[0] || pick(pool)[0];
    if (best) {
      updated[i]={...r,faculty:best.name,_dirty:true};
      hrs[best.name]=(hrs[best.name]||0)+(r.tempHrs||1.5);
      if (d&&t) occ[`${best.name}|${d}|${t}`]=true;
    }
  }
  return updated;
}

// ══════════════════════════════════════════════════════════════════════════════
// CONFLICT DETECTION
// ══════════════════════════════════════════════════════════════════════════════
function detectConflicts(rows) {
  const fac={},room={},out=new Set();
  for (const r of rows) {
    const d=normDay(r.day), t=normTime(r.time);
    if (!d||!t) continue;
    if (r.faculty&&!r.faculty.includes("PLACE")) {
      const k=`${r.faculty}|${d}|${t}`;
      (fac[k]=fac[k]||[]).push(r.id);
    }
    if (r.room) {
      const k=`${r.room}|${d}|${t}`;
      (room[k]=room[k]||[]).push(r.id);
    }
  }
  for (const ids of Object.values(fac))  if(ids.length>1) ids.forEach(id=>out.add(`fac:${id}`));
  for (const ids of Object.values(room)) if(ids.length>1) ids.forEach(id=>out.add(`room:${id}`));
  return out;
}
function calcLoads(rows) {
  const m={};
  for (const r of rows) {
    if (!r.faculty||r.faculty.includes("PLACE")) continue;
    m[r.faculty]=m[r.faculty]||{hrs:0,sections:0};
    m[r.faculty].hrs+=r.tempHrs||0;
    m[r.faculty].sections++;
  }
  return m;
}

// ══════════════════════════════════════════════════════════════════════════════
// EXCEL PARSERS
// ══════════════════════════════════════════════════════════════════════════════
async function parseLoadingFile(file) {
  return new Promise((res,rej)=>{
    if (!window.XLSX){rej("SheetJS not loaded yet — wait a moment and retry.");return;}
    const fr=new FileReader();
    fr.onload=e=>{
      try {
        const wb=window.XLSX.read(e.target.result,{type:"array"});
        const ws=wb.Sheets["Faculty Loading & Class Schedul"]||wb.Sheets[wb.SheetNames[0]];
        const raw=window.XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        const rows=[];
        for (let i=3;i<raw.length;i++) {
          const r=raw[i];
          if (!r[0]) continue;
          const code=String(r[4]||"").trim();
          rows.push({
            id:`r${i}`,
            yearLevel:String(r[0]||"").trim(),
            program:String(r[1]||"").trim(),
            faculty:String(r[2]||"").trim(),
            section:String(r[3]||"").trim(),
            code, title:String(r[5]||"").trim()||SUBJECT_TITLES[code]||code,
            typeClass:String(r[7]||"").trim(),
            enrolled:parseInt(r[8])||0,
            units:parseFloat(r[12])||0,
            tempHrs:parseFloat(r[16])||0,
            day:normDay(String(r[22]||"")),
            time:normTime(String(r[23]||"")),
            room:String(r[25]||"").trim(),
            identifier:String(r[20]||"").trim(),
            _dirty:false,
          });
        }
        res(rows);
      } catch(err){rej(err.message);}
    };
    fr.readAsArrayBuffer(file);
  });
}
async function parseRoomFile(file) {
  return new Promise((res,rej)=>{
    if (!window.XLSX){rej("SheetJS not loaded.");return;}
    const fr=new FileReader();
    fr.onload=e=>{
      try {
        const wb=window.XLSX.read(e.target.result,{type:"array"});
        const sn=wb.SheetNames.find(s=>s.toLowerCase().includes("assigned"))||wb.SheetNames[0];
        const ws=wb.Sheets[sn];
        const raw=window.XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        const rooms=[];let curB="";
        for (let i=2;i<raw.length;i++){
          const r=raw[i];
          if (r[0]) curB=String(r[0]).trim().replace(/\s+/g," ");
          if (!r[2]) continue;
          rooms.push({building:curB,room:String(r[2]).trim(),capacity:parseInt(r[3])||0,
            MON:String(r[4]||"").trim(),TUE:String(r[5]||"").trim(),
            WED:String(r[6]||"").trim(),THU:String(r[7]||"").trim(),
            FRI:String(r[8]||"").trim(),SAT:String(r[9]||"").trim()});
        }
        res(rooms);
      } catch(err){rej(err.message);}
    };
    fr.readAsArrayBuffer(file);
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// UI ATOMS
// ══════════════════════════════════════════════════════════════════════════════
function Btn({children,onClick,disabled,variant="primary",sm}) {
  const base=`font-semibold rounded-lg transition-colors disabled:opacity-40 whitespace-nowrap ${sm?"px-3 py-1.5 text-[11px]":"px-4 py-2 text-xs"}`;
  const v={primary:"bg-indigo-600 hover:bg-indigo-500 text-white",success:"bg-emerald-600 hover:bg-emerald-500 text-white",
    danger:"bg-red-700 hover:bg-red-600 text-white",ghost:"border border-slate-700 text-slate-400 hover:border-slate-500 hover:text-slate-200",
    amber:"bg-amber-600 hover:bg-amber-500 text-white"};
  return <button onClick={onClick} disabled={disabled} className={`${base} ${v[variant]||v.primary}`}>{children}</button>;
}
function CellSel({value,options,onChange,highlight,empty}) {
  return (
    <select value={value||""} onChange={e=>onChange(e.target.value)}
      style={highlight?{borderColor:"#f97316",background:"rgba(249,115,22,0.06)"}:{}}
      className={`w-full bg-transparent border rounded px-1.5 py-1 text-[11px] text-slate-200 focus:outline-none focus:border-indigo-400 cursor-pointer
        ${highlight?"border-orange-500/80":"border-slate-700/60 hover:border-slate-600"}`}>
      <option value="">{empty||"—"}</option>
      {options.map(o=>{const val=typeof o==="string"?o:o.value,lbl=typeof o==="string"?o:o.label;
        return <option key={val} value={val} disabled={o.disabled}>{lbl}</option>;})}
    </select>
  );
}
function CellTxt({value,onChange,placeholder,highlight}) {
  return <input type="text" value={value||""} onChange={e=>onChange(e.target.value)} placeholder={placeholder}
    style={highlight?{borderColor:"#f97316"}:{}}
    className={`w-full bg-transparent border rounded px-1.5 py-1 text-[11px] text-slate-200 focus:outline-none focus:border-indigo-400
      ${highlight?"border-orange-500/80":"border-slate-700/60 hover:border-slate-600"}`}/>;
}

// ══════════════════════════════════════════════════════════════════════════════
// IMPORT MODAL
// ══════════════════════════════════════════════════════════════════════════════
function ImportModal({type,onDone,onClose}) {
  const isRoom=type==="rooms";
  const [tab,setTab]=useState("excel");
  const [url,setUrl]=useState("");
  const [status,setStatus]=useState(null);
  const [busy,setBusy]=useState(false);
  const accent=isRoom?"#f59e0b":"#6366f1";

  const handleFile=async e=>{
    const f=e.target.files?.[0];if(!f)return;
    setBusy(true);setStatus(null);
    try {
      const rows=isRoom?await parseRoomFile(f):await parseLoadingFile(f);
      setStatus({ok:true,rows,msg:`${rows.length} rows from "${f.name}"`});
    } catch(err){setStatus({ok:false,rows:[],msg:String(err)});}
    setBusy(false);
  };

  const handleGSheet=async()=>{
    setBusy(true);setStatus(null);
    try {
      const m=url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
      if (!m) throw new Error("Invalid Google Sheets URL.");
      const gid=(url.match(/gid=(\d+)/)||[])[1]||"0";
      const resp=await fetch(`https://docs.google.com/spreadsheets/d/${m[1]}/export?format=csv&gid=${gid}`);
      if (!resp.ok) throw new Error(`HTTP ${resp.status} — share sheet as "Anyone with link → Viewer".`);
      const text=await resp.text();
      const XLSX=window.XLSX;if(!XLSX)throw new Error("SheetJS not loaded.");
      const wb=XLSX.read(text,{type:"string"});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
      const rows=[];const offset=isRoom?1:3;
      for (let i=offset;i<raw.length;i++){
        const r=raw[i];if(!r[0])continue;
        if (isRoom){rows.push({building:String(r[0]||"").trim(),room:String(r[2]||"").trim(),capacity:parseInt(r[3])||0,
          MON:String(r[4]||"").trim(),TUE:String(r[5]||"").trim(),WED:String(r[6]||"").trim(),
          THU:String(r[7]||"").trim(),FRI:String(r[8]||"").trim(),SAT:String(r[9]||"").trim()});}
        else{const code=String(r[4]||"").trim();rows.push({id:`gs${i}`,yearLevel:String(r[0]||"").trim(),
          program:String(r[1]||"").trim(),faculty:String(r[2]||"").trim(),section:String(r[3]||"").trim(),
          code,title:String(r[5]||"").trim()||SUBJECT_TITLES[code]||code,typeClass:String(r[7]||"").trim(),
          enrolled:parseInt(r[8])||0,units:parseFloat(r[12])||0,tempHrs:parseFloat(r[16])||0,
          day:normDay(String(r[22]||"")),time:normTime(String(r[23]||"")),
          room:String(r[25]||"").trim(),identifier:String(r[20]||"").trim(),_dirty:false});}
      }
      setStatus({ok:true,rows,msg:`${rows.length} rows from Google Sheet`});
    } catch(err){setStatus({ok:false,rows:[],msg:String(err)});}
    setBusy(false);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
      <div className="w-full max-w-lg bg-[#0f1120] border border-slate-700/60 rounded-2xl shadow-2xl flex flex-col max-h-[85vh]">
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-800">
          <div>
            <h2 className="text-sm font-bold text-white">{isRoom?"🏫 Import Classrooms":"📂 Load Faculty Loading Sheet"}</h2>
            <p className="text-[11px] text-slate-600 mt-0.5">{isRoom?"Classroom Assignment file":"CAS Faculty Loading & Class Schedules Excel"}</p>
          </div>
          <button onClick={onClose} className="text-slate-600 hover:text-white text-lg">✕</button>
        </div>
        <div className="flex gap-1 px-6 pt-4">
          {[{id:"excel",l:"Upload Excel",i:"📊"},{id:"gsheet",l:"Google Sheet URL",i:"🔗"}].map(t=>(
            <button key={t.id} onClick={()=>{setTab(t.id);setStatus(null);}}
              style={tab===t.id?{background:accent+"22",borderColor:accent,color:accent}:{}}
              className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-semibold border transition-all
                ${tab!==t.id?"border-slate-700 text-slate-500 hover:text-slate-300":""}`}>
              {t.i} {t.l}
            </button>
          ))}
        </div>
        <div className="flex-1 overflow-y-auto px-6 py-4 space-y-4">
          {tab==="excel"&&(
            <label className="flex flex-col items-center gap-3 border-2 border-dashed border-slate-700 hover:border-indigo-500 rounded-xl p-10 cursor-pointer transition-colors bg-slate-800/20">
              <span className="text-4xl">{isRoom?"🏫":"📊"}</span>
              <span className="text-sm text-slate-300 font-semibold">Click to select .xlsx file</span>
              <span className="text-xs text-slate-600">{isRoom?"1st_Sem_SY_2627_Classroom_Assignment.xlsx":"CAS-1S_2627_Faculty_Loading_Template.xlsx"}</span>
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFile}/>
            </label>
          )}
          {tab==="gsheet"&&(
            <div className="space-y-3">
              <div className="text-xs text-slate-600 bg-slate-800/40 rounded-xl p-3 leading-relaxed">
                Share sheet as <span className="text-amber-400">Anyone with link → Viewer</span>. Copy URL while target tab is active (<span className="font-mono text-sky-400">gid=…</span>).<br/>
                <span className="text-emerald-400">Re-fetch any time</span> for live data.
              </div>
              <input value={url} onChange={e=>{setUrl(e.target.value);setStatus(null);}}
                placeholder="https://docs.google.com/spreadsheets/d/…/edit#gid=…"
                className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-xs text-slate-200 font-mono focus:outline-none focus:border-indigo-500"/>
              <Btn onClick={handleGSheet} disabled={!url.trim()||busy} variant="ghost">{busy?"⏳ Fetching…":"🔗 Fetch Sheet"}</Btn>
            </div>
          )}
          {busy&&<div className="text-xs text-indigo-400 animate-pulse">⏳ Reading file…</div>}
          {status&&(
            <div className={`rounded-xl p-3 text-xs ${status.ok?"bg-emerald-900/20 border border-emerald-800/40 text-emerald-300":"bg-red-900/20 border border-red-800/40 text-red-300"}`}>
              {status.ok?"✓ ":""}{status.msg}
              {status.ok&&!isRoom&&<div className="mt-1 text-emerald-700">
                {status.rows.filter(r=>!r.faculty||r.faculty.includes("PLACE")).length} without faculty ·{" "}
                {status.rows.filter(r=>!r.room).length} without room
              </div>}
            </div>
          )}
        </div>
        <div className="px-6 py-4 border-t border-slate-800 flex gap-3">
          <button onClick={()=>status?.ok&&onDone(status.rows)} disabled={!status?.ok||!status.rows?.length}
            style={{background:accent}} className="flex-1 py-2 text-white text-xs font-bold rounded-lg disabled:opacity-40 hover:opacity-90">
            ✓ Import {status?.rows?.length?`${status.rows.length} rows`:""}
          </button>
          <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// AUTO-SCHEDULE SETTINGS MODAL
// ══════════════════════════════════════════════════════════════════════════════
function AutoModal({rows,onRun,onClose}) {
  const [scope,setScope]=useState("unscheduled");
  const [doFac,setDoFac]=useState(true);
  const [doSched,setDoSched]=useState(true);
  const unsch=rows.filter(r=>!r.day||!r.time).length;
  const unass=rows.filter(r=>!r.faculty||r.faculty.includes("PLACE")).length;
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
      <div className="w-full max-w-md bg-[#0f1120] border border-indigo-700/40 rounded-2xl shadow-2xl">
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-800">
          <h2 className="text-sm font-bold text-white">⚡ Auto-Schedule Settings</h2>
          <button onClick={onClose} className="text-slate-600 hover:text-white text-lg">✕</button>
        </div>
        <div className="px-6 py-5 space-y-5">
          <div>
            <div className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">What to run</div>
            {[{id:"doSched",l:"⏰ Assign Day & Time",s:`${unsch} unscheduled rows`,v:doSched,set:setDoSched},
              {id:"doFac",  l:"👤 Assign Faculty",   s:`${unass} unassigned rows`, v:doFac,  set:setDoFac}
            ].map(o=>(
              <label key={o.id} className="flex items-center gap-3 p-3 mb-2 rounded-xl border border-slate-800 cursor-pointer hover:border-slate-700 bg-slate-900/30">
                <input type="checkbox" checked={o.v} onChange={e=>o.set(e.target.checked)} className="w-4 h-4 accent-indigo-500"/>
                <div><div className="text-xs font-semibold text-slate-200">{o.l}</div>
                <div className="text-[10px] text-slate-600">{o.s}</div></div>
              </label>
            ))}
          </div>
          <div>
            <div className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Scope</div>
            <div className="flex gap-2">
              {[["unscheduled","Only empty cells"],["all","Re-assign everything"]].map(([v,l])=>(
                <button key={v} onClick={()=>setScope(v)}
                  className={`flex-1 py-2 px-3 rounded-xl text-xs font-semibold border transition-all
                    ${scope===v?"bg-indigo-600 border-indigo-500 text-white":"border-slate-700 text-slate-500 hover:border-slate-600"}`}>
                  {l}
                </button>
              ))}
            </div>
            {scope==="all"&&<div className="mt-2 text-[10px] text-amber-400 bg-amber-900/20 rounded-lg px-3 py-1.5 border border-amber-800/30">⚠ Overwrites existing assignments.</div>}
          </div>
          <div className="bg-slate-900/50 rounded-xl p-3 text-[11px] text-slate-600 space-y-1 border border-slate-800">
            <div className="text-slate-400 font-semibold mb-1">Rules enforced:</div>
            <div>🔒 12:00NN–1:30PM = lunch, always protected</div>
            <div>🍱 No lunch trap (10:30 + 1:30 back-to-back)</div>
            <div>⏱ No 3 consecutive slots (4.5h straight)</div>
            <div>⚖ First ½ sections → AM · Second ½ → PM</div>
            <div>🔄 Time cursor rotates per subject code</div>
            <div>📅 Day cursor rotates for week-wide spread</div>
          </div>
        </div>
        <div className="px-6 pb-5 flex gap-3">
          <Btn onClick={()=>onRun({scope,doFac,doSched})} disabled={!doFac&&!doSched} variant="success">⚡ Run Auto-Schedule</Btn>
          <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// LOADING TABLE — main spreadsheet editor
// ══════════════════════════════════════════════════════════════════════════════
function LoadingTable({rows,setRows,rooms,conflicts}) {
  const [f,setF]=useState({college:"All",program:"All",yearLevel:"All",code:"All",show:"All",search:""});
  const [sort,setSort]=useState("section");

  const programs=useMemo(()=>["All",...new Set(rows.map(r=>r.program).filter(Boolean).sort())]          ,[rows]);
  const codes   =useMemo(()=>["All",...new Set(rows.map(r=>r.code).filter(Boolean).sort())]             ,[rows]);
  const roomList=useMemo(()=>rooms.map(r=>r.room).filter(Boolean).sort()                                ,[rooms]);
  const facList =useMemo(()=>{const s=new Set();rows.forEach(r=>{if(r.faculty&&!r.faculty.includes("PLACE"))s.add(r.faculty);});return [...s].sort();},[rows]);
  const facConf =useMemo(()=>new Set([...conflicts].filter(c=>c.startsWith("fac:")).map(c=>c.slice(4))) ,[conflicts]);
  const roomConf=useMemo(()=>new Set([...conflicts].filter(c=>c.startsWith("room:")).map(c=>c.slice(5))),[conflicts]);

  const filtered=useMemo(()=>{
    let l=[...rows];
    if (f.college!=="All")   l=l.filter(r=>getCollege(r.program)===f.college);
    if (f.program!=="All")   l=l.filter(r=>r.program===f.program);
    if (f.yearLevel!=="All") l=l.filter(r=>r.yearLevel===f.yearLevel);
    if (f.code!=="All")      l=l.filter(r=>r.code===f.code);
    if (f.show==="Missing Faculty")  l=l.filter(r=>!r.faculty||r.faculty.includes("PLACE"));
    if (f.show==="Missing Room")     l=l.filter(r=>!r.room);
    if (f.show==="Missing Schedule") l=l.filter(r=>!r.day||!r.time);
    if (f.show==="Conflicts")        l=l.filter(r=>facConf.has(r.id)||roomConf.has(r.id));
    if (f.show==="Changed")          l=l.filter(r=>r._dirty);
    if (f.search){const q=f.search.toLowerCase();l=l.filter(r=>[r.section,r.faculty,r.code,r.program,r.room].some(v=>v?.toLowerCase().includes(q)));}
    const field=sort==="faculty"?"faculty":sort==="program"?"program":"section";
    return l.sort((a,b)=>(a[field]||"").localeCompare(b[field]||""));
  },[rows,f,sort,facConf,roomConf]);

  const upd=(id,field,val)=>setRows(prev=>prev.map(r=>r.id===id?{...r,[field]:val,_dirty:true}:r));
  const sf=(k,v)=>setF(p=>({...p,[k]:v}));
  const PLACEHOLDERS=["SOCSCI PLACE HOLDER 1","LANGUAGES PLACE HOLDER 1","LANGUAGES PLACE HOLDER 2","CCJE PLACEHOLDER 1","CCJE PLACEHOLDER 2"];

  return (
    <div className="flex flex-col gap-3 h-full">
      {/* Filter bar */}
      <div className="flex gap-2 flex-wrap items-center bg-[#0b0d18] border border-slate-800/80 rounded-xl px-3 py-2.5">
        <input value={f.search} onChange={e=>sf("search",e.target.value)}
          placeholder="🔍 section / faculty / code / room"
          className="bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-indigo-500 w-52"/>
        {[{k:"college",opts:["All",...Object.keys(COLLEGES)]},
          {k:"program",opts:programs},{k:"yearLevel",opts:["All","Freshmen","Upperclassmen"]},
          {k:"code",opts:codes},{k:"show",opts:["All","Missing Faculty","Missing Room","Missing Schedule","Conflicts","Changed"]}
        ].map(({k,opts})=>(
          <select key={k} value={f[k]} onChange={e=>sf(k,e.target.value)}
            className="bg-slate-800 border border-slate-700 rounded-lg px-2 py-1.5 text-xs text-slate-300 focus:outline-none focus:border-indigo-500">
            {opts.map(o=><option key={o}>{o}</option>)}
          </select>
        ))}
        <div className="ml-auto text-[11px] text-slate-700">{filtered.length}/{rows.length}</div>
      </div>

      {/* Table */}
      <div className="overflow-auto rounded-xl border border-slate-800 flex-1">
        <table className="w-full text-xs border-collapse">
          <thead className="sticky top-0 z-10">
            <tr className="bg-[#090b14] border-b border-slate-800">
              {[{h:"●",w:"w-10"},{h:"Yr",w:"w-12",s:"yearLevel"},{h:"",w:"w-8"},
                {h:"Program",w:"w-24",s:"program"},{h:"Section",w:"w-32",s:"section"},
                {h:"Subject",w:"w-32"},{h:"Type",w:"w-16"},{h:"Hrs",w:"w-10"},
                {h:"▶ FACULTY (C)",w:"w-52",s:"faculty",a:true},
                {h:"▶ DAY (W)",w:"w-28",a:true},{h:"▶ TIME (X)",w:"w-44",a:true},{h:"▶ ROOM (Z)",w:"w-28",a:true},
              ].map(c=>(
                <th key={c.h} onClick={()=>c.s&&setSort(c.s)}
                  style={c.a?{borderBottom:"2px solid rgba(99,102,241,0.5)"}:{}}
                  className={`text-left px-2 py-2.5 font-semibold whitespace-nowrap select-none
                    ${c.a?"text-indigo-400":"text-slate-600"} ${c.s?"cursor-pointer hover:text-slate-400":""} ${c.w||""}`}>
                  {c.h}{c.s&&sort===c.s?" ↑":""}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((row,idx)=>{
              const mFac=!row.faculty||row.faculty.includes("PLACE");
              const mRoom=!row.room;
              const mTime=!row.day||!row.time;
              const isFc=facConf.has(row.id), isRc=roomConf.has(row.id);
              const col=clg(row.program);
              const nt=normTime(row.time);
              const nonStd=row.time&&nt&&!ALL_SLOTS.includes(nt);
              const bg=isFc||isRc?"rgba(239,68,68,0.05)":row._dirty?"rgba(99,102,241,0.03)":idx%2===0?"transparent":"rgba(255,255,255,0.01)";
              return (
                <tr key={row.id} style={{background:bg}} className="border-b border-slate-900/50 hover:bg-white/[0.02] transition-colors">
                  {/* Status */}
                  <td className="px-2 py-2 text-center">
                    <div className="flex gap-0.5 justify-center">
                      {[{ok:!mFac},{ok:!mTime,warn:nonStd},{ok:!mRoom}].map((d,i)=>(
                        <span key={i} className={`w-2 h-2 rounded-full shrink-0 ${d.warn?"bg-amber-400":d.ok?"bg-emerald-500":"bg-red-500 animate-pulse"}`}/>
                      ))}
                    </div>
                    {(isFc||isRc)&&<div className="text-[8px] text-red-500 font-bold mt-0.5">CONF</div>}
                  </td>
                  {/* Year */}
                  <td className="px-2 py-2">
                    <span className={`text-[10px] font-bold px-1 py-0.5 rounded ${row.yearLevel==="Freshmen"?"bg-blue-900/40 text-blue-300":"bg-violet-900/40 text-violet-300"}`}>
                      {row.yearLevel==="Freshmen"?"Y1":"Y2+"}
                    </span>
                  </td>
                  {/* College dot */}
                  <td className="px-1 py-2 text-center">
                    {col&&<span style={{color:col.color,fontSize:14}}>●</span>}
                  </td>
                  <td className="px-2 py-2 text-slate-300 text-[11px] font-medium">{row.program}</td>
                  <td className="px-2 py-2 font-mono text-slate-500 text-[11px] whitespace-nowrap">{row.section}</td>
                  <td className="px-2 py-2">
                    <div className="text-[11px] font-semibold text-slate-300">{row.code}</div>
                    <div className="text-[10px] text-slate-700 max-w-[140px] truncate" title={row.title}>{row.title}</div>
                  </td>
                  <td className="px-2 py-2 text-slate-600 text-[10px] whitespace-nowrap">{row.typeClass?.replace("Parallel of ","P·")}</td>
                  <td className="px-2 py-2 text-center text-slate-600 font-mono text-[11px]">{row.tempHrs}</td>
                  {/* FACULTY */}
                  <td className="px-1.5 py-1.5" style={{borderLeft:"1px solid rgba(99,102,241,0.2)"}}>
                    <CellSel value={row.faculty} highlight={mFac||isFc} empty="— assign faculty —" onChange={v=>upd(row.id,"faculty",v)}
                      options={[...facList.map(f=>({value:f,label:f})),{value:"",label:"── Placeholders ──",disabled:true},...PLACEHOLDERS.map(p=>({value:p,label:p}))]}/>
                    {isFc&&<div className="text-[9px] text-red-500 mt-0.5">⚠ double-booked</div>}
                  </td>
                  {/* DAY */}
                  <td className="px-1.5 py-1.5">
                    <CellSel value={normDay(row.day)} highlight={mTime} empty="— day —" onChange={v=>upd(row.id,"day",v)} options={ALL_DAYS}/>
                  </td>
                  {/* TIME */}
                  <td className="px-1.5 py-1.5">
                    <CellSel value={nt} highlight={mTime} empty="— time —" onChange={v=>upd(row.id,"time",v)}
                      options={[{value:"",label:"─ AM ─",disabled:true},...AM_SLOTS,
                        {value:LUNCH,label:`${LUNCH} (LUNCH)`},
                        {value:"",label:"─ PM ─",disabled:true},...PM_SLOTS]}/>
                    <div className="flex gap-1 mt-0.5">
                      {isAM(nt)&&<span className="text-[9px] text-sky-600 font-bold">AM</span>}
                      {isPM(nt)&&<span className="text-[9px] text-amber-600 font-bold">PM</span>}
                      {nonStd&&<span className="text-[9px] text-amber-500" title={row.time}>⚠ non-std</span>}
                    </div>
                  </td>
                  {/* ROOM */}
                  <td className="px-1.5 py-1.5" style={{borderRight:"1px solid rgba(99,102,241,0.2)"}}>
                    {roomList.length>0
                      ?<CellSel value={row.room} highlight={mRoom||isRc} empty="— room —" onChange={v=>upd(row.id,"room",v)} options={roomList}/>
                      :<CellTxt value={row.room} highlight={mRoom||isRc} placeholder="BL 301" onChange={v=>upd(row.id,"room",v)}/>}
                    {isRc&&<div className="text-[9px] text-red-500 mt-0.5">⚠ room conflict</div>}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {filtered.length===0&&<div className="text-center py-16 text-slate-700 text-sm">{rows.length===0?"No data loaded.":"No matching rows."}</div>}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// FACULTY SUMMARY
// ══════════════════════════════════════════════════════════════════════════════
function FacultySummary({rows,conflicts}) {
  const loads=useMemo(()=>calcLoads(rows),[rows]);
  const facConf=useMemo(()=>new Set([...conflicts].filter(c=>c.startsWith("fac:")).map(c=>c.slice(4))),[conflicts]);
  const byFac=useMemo(()=>{const m={};rows.forEach(r=>{if(r.faculty&&!r.faculty.includes("PLACE")){m[r.faculty]=m[r.faculty]||[];m[r.faculty].push(r);}});return m;},[rows]);
  const [search,setSearch]=useState("");
  const [open,setOpen]=useState(null);
  const MAX=24;
  const names=Object.keys(loads).sort((a,b)=>(loads[b]?.hrs||0)-(loads[a]?.hrs||0));
  const shown=search?names.filter(n=>n.toLowerCase().includes(search.toLowerCase())):names;
  return (
    <div className="space-y-3">
      <div className="flex items-center gap-3">
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Search faculty…"
          className="bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-indigo-500 w-52"/>
        <span className="text-xs text-slate-600">{shown.length} faculty loaded</span>
      </div>
      <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
        {shown.map(name=>{
          const ld=loads[name]||{hrs:0,sections:0};
          const pct=Math.min((ld.hrs/MAX)*100,100);
          const over=ld.hrs>MAX;
          const hasConf=(byFac[name]||[]).some(r=>facConf.has(r.id));
          const isOpen=open===name;
          const col=clg(byFac[name]?.[0]?.program||"");
          return (
            <div key={name} onClick={()=>setOpen(isOpen?null:name)}
              style={hasConf?{borderColor:"rgba(239,68,68,0.4)"}:over?{borderColor:"rgba(245,158,11,0.4)"}:{}}
              className="bg-[#0d0f1a] border border-slate-800 rounded-xl cursor-pointer hover:border-slate-700 transition-all overflow-hidden">
              <div className="p-3">
                <div className="flex justify-between items-start mb-2">
                  <div>
                    <div className="text-[11px] font-bold text-slate-200 leading-tight" title={name}>{name.length>26?name.slice(0,24)+"…":name}</div>
                    {col&&<div className="text-[10px] font-bold mt-0.5" style={{color:col.color}}>{getCollege(byFac[name]?.[0]?.program||"")}</div>}
                  </div>
                  <div className="text-right">
                    <div className={`text-base font-bold tabular-nums ${over?"text-red-400":"text-emerald-400"}`}>{ld.hrs}<span className="text-[10px] text-slate-600">h</span></div>
                    <div className="text-[10px] text-slate-600">{ld.sections}sec</div>
                  </div>
                </div>
                <div className="h-1.5 bg-slate-800 rounded-full overflow-hidden">
                  <div className={`h-full rounded-full transition-all ${over?"bg-red-500":pct>83?"bg-amber-500":"bg-emerald-500"}`} style={{width:`${pct}%`}}/>
                </div>
                <div className="flex justify-between mt-1 text-[9px] text-slate-700">
                  <span>{pct.toFixed(0)}% of {MAX}h</span>
                  <div className="flex gap-1">{hasConf&&<span className="text-red-500 font-bold">CONFLICT</span>}{over&&<span className="text-amber-500 font-bold">OVERLOAD</span>}</div>
                </div>
              </div>
              {isOpen&&(
                <div className="border-t border-slate-800 bg-slate-900/40 max-h-44 overflow-y-auto">
                  {(byFac[name]||[]).map(r=>(
                    <div key={r.id} className={`flex gap-2 px-3 py-1 text-[10px] border-b border-slate-900/50 ${facConf.has(r.id)?"bg-red-900/10":""}`}>
                      <span className="font-mono text-slate-600 w-24 shrink-0">{r.section}</span>
                      <span className="text-slate-600 w-16 shrink-0">{r.code}</span>
                      <span className={`shrink-0 font-bold ${isAM(normTime(r.time))?"text-sky-500":isPM(normTime(r.time))?"text-amber-500":"text-slate-600"}`}>{normDay(r.day)||"—"}</span>
                      <span className="text-slate-700 truncate">{normTime(r.time)||"—"}</span>
                      {facConf.has(r.id)&&<span className="text-red-500 shrink-0">⚠</span>}
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// SCHEDULE VIEW
// ══════════════════════════════════════════════════════════════════════════════
function ScheduleView({rows}) {
  const [day,setDay]=useState("MON");
  const [grp,setGrp]=useState("program");
  const forDay=useMemo(()=>rows.filter(r=>{const d=normDay(r.day);return d===day||d.includes(day);}),[rows,day]);
  const bySlot=useMemo(()=>{
    const m={};
    ALL_SLOTS.forEach(s=>m[s]=[]);
    for (const r of forDay){
      const t=normTime(r.time);
      if (m[t]) m[t].push(r);
    }
    return m;
  },[forDay]);
  const dayCnt=d=>rows.filter(r=>{const nd=normDay(r.day);return nd===d||nd.includes(d);}).length;
  return (
    <div className="space-y-4">
      <div className="flex gap-1 flex-wrap items-center">
        {WORK_DAYS.map(d=>(
          <button key={d} onClick={()=>setDay(d)}
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold border transition-all
              ${day===d?"bg-indigo-600 border-indigo-500 text-white":"border-slate-700 text-slate-500 hover:text-slate-300"}`}>
            {d} <span className="text-[10px] opacity-50">({dayCnt(d)})</span>
          </button>
        ))}
        <select value={grp} onChange={e=>setGrp(e.target.value)}
          className="ml-auto bg-slate-800 border border-slate-700 rounded-lg px-2 py-1.5 text-xs text-slate-300 focus:outline-none">
          {["program","faculty","code"].map(o=><option key={o} value={o}>Group: {o}</option>)}
        </select>
      </div>
      <div className="space-y-1.5">
        {ALL_SLOTS.map(slot=>{
          const sr=bySlot[slot]||[];
          const isLunch=slot===LUNCH;
          if (isLunch) return (
            <div key={slot} className="flex items-center gap-3">
              <div className="w-36 shrink-0 text-[10px] text-slate-700 font-mono text-right pr-2">{slot}</div>
              <div className="flex-1 h-8 rounded-lg border border-dashed border-slate-800/60 flex items-center justify-center text-[10px] text-slate-800 font-bold tracking-widest">
                LUNCH — PROTECTED
              </div>
            </div>
          );
          if (!sr.length) return (
            <div key={slot} className="flex items-center gap-3">
              <div className={`w-36 shrink-0 text-[10px] font-mono text-right pr-2 ${isAM(slot)?"text-sky-800":"text-amber-900"}`}>{slot}</div>
              <div className={`flex-1 h-6 rounded-lg ${isAM(slot)?"bg-sky-950/20":"bg-amber-950/10"} border ${isAM(slot)?"border-sky-900/20":"border-amber-900/10"}`}/>
            </div>
          );
          const grouped={};
          for (const r of sr){const k=grp==="program"?r.program:grp==="faculty"?r.faculty:r.code;(grouped[k]=grouped[k]||[]).push(r);}
          return (
            <div key={slot} className="flex gap-3 items-start">
              <div className={`w-36 shrink-0 text-[10px] font-mono text-right pr-2 pt-2 font-semibold ${isAM(slot)?"text-sky-500":"text-amber-500"}`}>{slot}</div>
              <div className="flex-1 flex gap-1.5 flex-wrap py-1">
                {Object.entries(grouped).map(([key,rr])=>{
                  const col=clg(rr[0]?.program);
                  return (
                    <div key={key} style={col?{background:col.bg,borderColor:col.color+"30"}:{background:"rgba(99,102,241,0.07)",borderColor:"rgba(99,102,241,0.2)"}}
                      className="border rounded-lg px-2 py-1.5 text-[10px] min-w-[90px] max-w-[160px]">
                      <div className="font-bold truncate" style={{color:col?.color||"#6366f1"}}>{key||"—"}</div>
                      <div className="text-slate-500">{rr.length}sec · {rr[0]?.code}</div>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// CONFLICTS PANEL
// ══════════════════════════════════════════════════════════════════════════════
function ConflictsPanel({rows,conflicts}) {
  const fCnf=useMemo(()=>new Set([...conflicts].filter(c=>c.startsWith("fac:")).map(c=>c.slice(4))),[conflicts]);
  const rCnf=useMemo(()=>new Set([...conflicts].filter(c=>c.startsWith("room:")).map(c=>c.slice(5))),[conflicts]);
  const fg={},rg={};
  for (const r of rows){
    const d=normDay(r.day),t=normTime(r.time);
    if (fCnf.has(r.id)&&r.faculty&&d&&t){const k=`${r.faculty}||${d}||${t}`;(fg[k]=fg[k]||[]).push(r);}
    if (rCnf.has(r.id)&&r.room&&d&&t){const k=`${r.room}||${d}||${t}`;(rg[k]=rg[k]||[]).push(r);}
  }
  const mF=rows.filter(r=>!r.faculty||r.faculty.includes("PLACE"));
  const mR=rows.filter(r=>!r.room);
  const mS=rows.filter(r=>!r.day||!r.time);
  return (
    <div className="space-y-5">
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
        {[{l:"Faculty Conflicts",n:Object.keys(fg).length,c:"#ef4444"},{l:"Room Conflicts",n:Object.keys(rg).length,c:"#f97316"},
          {l:"Missing Faculty",n:mF.length,c:"#f59e0b"},{l:"Missing Room/Sched",n:mR.length+mS.length,c:"#6366f1"}
        ].map(s=>(
          <div key={s.l} style={{borderColor:s.c+"30"}} className="bg-[#0d0f1a] border rounded-xl p-4 text-center">
            <div className="text-2xl font-bold" style={{color:s.c}}>{s.n}</div>
            <div className="text-[11px] text-slate-600 mt-1">{s.l}</div>
          </div>
        ))}
      </div>
      {Object.entries(fg).length>0&&<div>
        <div className="text-xs font-bold text-red-400 uppercase tracking-widest mb-2">⚠ Faculty Double-Bookings</div>
        {Object.entries(fg).map(([k,rr])=>{const [fac,d,t]=k.split("||");return(
          <div key={k} className="bg-red-900/10 border border-red-900/30 rounded-xl p-3 mb-2">
            <div className="text-xs font-bold text-red-300 mb-1.5">{fac} — {d} {t}</div>
            <div className="flex gap-2 flex-wrap">{rr.map(r=><span key={r.id} className="text-[10px] font-mono bg-red-900/20 border border-red-900/30 rounded px-2 py-0.5 text-red-400">{r.section} ({r.code})</span>)}</div>
          </div>
        );})}
      </div>}
      {Object.entries(rg).length>0&&<div>
        <div className="text-xs font-bold text-orange-400 uppercase tracking-widest mb-2">⚠ Room Double-Bookings</div>
        {Object.entries(rg).map(([k,rr])=>{const [rm,d,t]=k.split("||");return(
          <div key={k} className="bg-orange-900/10 border border-orange-900/30 rounded-xl p-3 mb-2">
            <div className="text-xs font-bold text-orange-300 mb-1.5">{rm} — {d} {t}</div>
            <div className="flex gap-2 flex-wrap">{rr.map(r=><span key={r.id} className="text-[10px] font-mono bg-orange-900/20 border border-orange-900/30 rounded px-2 py-0.5 text-orange-400">{r.section}/{r.program}</span>)}</div>
          </div>
        );})}
      </div>}
      {!Object.keys(fg).length&&!Object.keys(rg).length&&<div className="text-center py-10 text-emerald-500 font-semibold">✓ No scheduling conflicts</div>}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// DASHBOARD
// ══════════════════════════════════════════════════════════════════════════════
function Dashboard({rows,conflicts,onAuto}) {
  const total=rows.length;
  const wFac=rows.filter(r=>r.faculty&&!r.faculty.includes("PLACE")).length;
  const wRm=rows.filter(r=>r.room).length;
  const wSch=rows.filter(r=>r.day&&r.time).length;
  const full=rows.filter(r=>r.faculty&&!r.faculty.includes("PLACE")&&r.room&&r.day&&r.time).length;
  const amC=rows.filter(r=>isAM(normTime(r.time))).length;
  const pmC=rows.filter(r=>isPM(normTime(r.time))).length;
  const pct=(n,d)=>d>0?Math.round((n/d)*100):0;
  const byCol={};
  for (const r of rows){const c=getCollege(r.program)||"Other";byCol[c]=byCol[c]||{t:0,f:0,r:0,s:0};byCol[c].t++;
    if(r.faculty&&!r.faculty.includes("PLACE"))byCol[c].f++;if(r.room)byCol[c].r++;if(r.day&&r.time)byCol[c].s++;}
  return (
    <div className="space-y-5">
      <div className="bg-[#0d0f1a] border border-indigo-700/25 rounded-xl p-4">
        <div className="text-[10px] font-bold text-slate-600 uppercase tracking-widest mb-3">⚡ Quick Actions</div>
        <div className="flex gap-2 flex-wrap">
          <Btn variant="success" onClick={onAuto}>⚡ Auto-Schedule + Assign Faculty</Btn>
        </div>
        <div className="mt-2 text-[10px] text-slate-700 leading-relaxed">
          Rules: Lunch protected · No lunch trap · No 3-consecutive · AM/PM half-split · Subject/day cursor rotation
        </div>
      </div>
      <div className="grid grid-cols-3 gap-3 sm:grid-cols-6">
        {[{l:"Total",v:total,c:"#6366f1"},{l:"Complete",v:full,c:"#22c55e",s:`${pct(full,total)}%`},
          {l:"Faculty",v:wFac,c:"#3b82f6",s:`${pct(wFac,total)}%`},{l:"Rooms",v:wRm,c:"#f59e0b",s:`${pct(wRm,total)}%`},
          {l:"Scheduled",v:wSch,c:"#a855f7",s:`${pct(wSch,total)}%`},{l:"Conflicts",v:conflicts.size,c:conflicts.size>0?"#ef4444":"#22c55e"},
        ].map(s=>(
          <div key={s.l} style={{borderColor:s.c+"25"}} className="bg-[#0d0f1a] border rounded-xl p-3 text-center">
            <div className="text-xl font-bold tabular-nums" style={{color:s.c}}>{s.v}</div>
            <div className="text-[10px] text-slate-600 mt-0.5">{s.l}</div>
            {s.s&&<div className="text-[9px] text-slate-700">{s.s}</div>}
          </div>
        ))}
      </div>
      {/* AM/PM bar */}
      <div className="bg-[#0d0f1a] border border-slate-800 rounded-xl p-4">
        <div className="text-[10px] font-bold text-slate-600 uppercase tracking-widest mb-3">AM · PM Split</div>
        <div className="flex gap-4 text-xs mb-2">
          <span className="text-sky-400 font-bold">AM: {amC} ({pct(amC,total)}%)</span>
          <span className="text-amber-400 font-bold">PM: {pmC} ({pct(pmC,total)}%)</span>
          <span className="text-slate-600">Unscheduled: {total-amC-pmC}</span>
        </div>
        <div className="h-4 bg-slate-800 rounded-full overflow-hidden flex">
          <div style={{width:`${pct(amC,total)}%`}} className="bg-sky-600 transition-all"/>
          <div style={{width:`${pct(pmC,total)}%`}} className="bg-amber-600 transition-all"/>
        </div>
      </div>
      {/* Progress */}
      <div className="bg-[#0d0f1a] border border-slate-800 rounded-xl p-4">
        <div className="text-[10px] font-bold text-slate-600 uppercase tracking-widest mb-3">Completion</div>
        {[{l:"Faculty",v:wFac,c:"#3b82f6"},{l:"Rooms",v:wRm,c:"#f59e0b"},{l:"Schedule",v:wSch,c:"#a855f7"},{l:"Complete",v:full,c:"#22c55e"}].map(it=>(
          <div key={it.l} className="mb-2.5">
            <div className="flex justify-between text-[11px] mb-1"><span className="text-slate-600">{it.l}</span>
              <span className="text-slate-700 tabular-nums">{it.v}/{total} ({pct(it.v,total)}%)</span></div>
            <div className="h-1.5 bg-slate-800 rounded-full overflow-hidden">
              <div style={{width:`${pct(it.v,total)}%`,background:it.c}} className="h-full rounded-full transition-all duration-700"/>
            </div>
          </div>
        ))}
      </div>
      {/* By college */}
      <div className="bg-[#0d0f1a] border border-slate-800 rounded-xl overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-800 text-[10px] font-bold text-slate-600 uppercase tracking-widest">By College</div>
        <table className="w-full text-xs"><thead><tr className="border-b border-slate-900">
          {["College","Sections","Faculty","Room","Schedule","Done"].map(h=><th key={h} className="px-4 py-2 text-left text-slate-700 font-semibold">{h}</th>)}
        </tr></thead><tbody>
          {Object.entries(byCol).sort(([a],[b])=>a.localeCompare(b)).map(([col,d])=>{
            const cd=COLLEGES[col];const done=Math.min(d.f,d.r,d.s);
            return (<tr key={col} className="border-b border-slate-900/50 hover:bg-white/[0.015]">
              <td className="px-4 py-2"><span className="font-bold" style={{color:cd?.color||"#94a3b8"}}>{col}</span></td>
              <td className="px-4 py-2 text-slate-600 tabular-nums">{d.t}</td>
              {[d.f,d.r,d.s].map((v,i)=><td key={i} className="px-4 py-2"><span className={v===d.t?"text-emerald-400":"text-amber-400"}>{v}/{d.t}</span></td>)}
              <td className="px-4 py-2"><div className="flex items-center gap-2">
                <div className="h-1.5 w-16 bg-slate-800 rounded-full overflow-hidden">
                  <div className="h-full bg-emerald-500 rounded-full" style={{width:`${pct(done,d.t)}%`}}/></div>
                <span className="text-slate-600 text-[10px]">{pct(done,d.t)}%</span>
              </div></td>
            </tr>);
          })}
        </tbody></table>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// EXPORT
// ══════════════════════════════════════════════════════════════════════════════
function ExportPanel({rows}) {
  const [grp,setGrp]=useState("program");
  const grouped=useMemo(()=>{const g={};for(const r of rows){const k=grp==="college"?(getCollege(r.program)||"Other"):r[grp]||"—";(g[k]=g[k]||[]).push(r);}return g;},[rows,grp]);
  const dl=()=>{
    const h=["YearLevel","Program","Faculty","Section","Code","Title","TypeClass","Units","TempHrs","Day","Time","Room"];
    const csv=[h.join(","),...rows.map(r=>[r.yearLevel,r.program,`"${r.faculty||""}"`,r.section,r.code,`"${r.title||""}"`,r.typeClass,r.units,r.tempHrs,normDay(r.day),normTime(r.time),r.room].join(","))].join("\n");
    const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv"}));a.download="faculty_loading.csv";a.click();
  };
  return (
    <div className="space-y-4">
      <div className="flex gap-3 items-center flex-wrap">
        <select value={grp} onChange={e=>setGrp(e.target.value)}
          className="bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-200 focus:outline-none">
          {["program","college","faculty","code"].map(o=><option key={o} value={o}>Group by {o}</option>)}
        </select>
        <Btn onClick={dl}>⬇ Download CSV</Btn>
        <span className="text-xs text-slate-600">{rows.length} rows</span>
      </div>
      {Object.entries(grouped).sort(([a],[b])=>a.localeCompare(b)).map(([key,rr])=>{
        const col=clg(rr[0]?.program);
        const af=rr.filter(r=>r.faculty&&!r.faculty.includes("PLACE")).length;
        const ar=rr.filter(r=>r.room).length;
        return (
          <div key={key} className="bg-[#0d0f1a] border border-slate-800 rounded-xl overflow-hidden">
            <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-800">
              <span className="font-bold text-sm text-slate-200">{key}</span>
              {col&&<span style={{color:col.color,fontSize:10,fontWeight:700}}>{getCollege(rr[0]?.program)}</span>}
              <div className="ml-auto flex gap-4 text-xs">
                <span className="text-slate-600">{rr.length}sec</span>
                <span className={af===rr.length?"text-emerald-400":"text-amber-400"}>{af}/{rr.length}fac</span>
                <span className={ar===rr.length?"text-emerald-400":"text-orange-400"}>{ar}/{rr.length}rm</span>
              </div>
            </div>
            <div className="overflow-x-auto"><table className="w-full text-[11px]"><tbody>
              {rr.map(r=>(
                <tr key={r.id} className="border-b border-slate-900/40 hover:bg-white/[0.015]">
                  <td className="px-3 py-1.5 font-mono text-slate-600 w-32 whitespace-nowrap">{r.section}</td>
                  <td className="px-3 py-1.5 text-slate-600 w-20">{r.code}</td>
                  <td className="px-3 py-1.5 font-semibold text-slate-300 w-48">{r.faculty||<span className="text-red-500 italic text-[10px]">unassigned</span>}</td>
                  <td className={`px-3 py-1.5 w-16 font-bold text-[10px] ${isAM(normTime(r.time))?"text-sky-500":isPM(normTime(r.time))?"text-amber-500":"text-slate-600"}`}>{normDay(r.day)||"—"}</td>
                  <td className="px-3 py-1.5 text-slate-600 w-36 whitespace-nowrap">{normTime(r.time)||"—"}</td>
                  <td className="px-3 py-1.5 text-slate-600">{r.room||<span className="text-orange-500 italic">—</span>}</td>
                </tr>
              ))}
            </tbody></table></div>
          </div>
        );
      })}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// ROOT
// ══════════════════════════════════════════════════════════════════════════════
const NAV=[{id:"dashboard",l:"Dashboard",i:"⊞"},{id:"loading",l:"Faculty Loading",i:"📋"},
  {id:"faculty",l:"Faculty Summary",i:"👤"},{id:"schedule",l:"Schedule View",i:"🗓"},
  {id:"conflicts",l:"Conflicts",i:"⚠"},{id:"export",l:"Export",i:"📤"}];

export default function App() {
  const [rows,setRows]=useState([]);
  const [rooms,setRooms]=useState([]);
  const [tab,setTab]=useState("dashboard");
  const [modal,setModal]=useState(null);
  const [toast,setToast]=useState(null);

  useEffect(()=>{
    if (!window.XLSX){const s=document.createElement("script");
      s.src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";document.head.appendChild(s);}
  },[]);

  const conflicts=useMemo(()=>detectConflicts(rows),[rows]);

  const flash=(msg,type="ok")=>{setToast({msg,type});setTimeout(()=>setToast(null),3500);};

  const runAuto=({scope,doFac,doSched})=>{
    setModal(null);
    let u=scope==="all"?rows.map(r=>({...r,day:"",time:"",faculty:r.faculty?.includes("PLACE")?"":r.faculty})):[...rows];
    if (doSched) u=autoSchedule(u);
    if (doFac) u=autoAssignFaculty(u,buildFacultyPool(rows));
    setRows(u);
    flash(`Done — ${u.filter(r=>r.day&&r.time).length} scheduled · ${u.filter(r=>r.faculty&&!r.faculty.includes("PLACE")).length} faculty assigned`);
    setTab("dashboard");
  };

  const mFac=rows.filter(r=>!r.faculty||r.faculty.includes("PLACE")).length;
  const mRoom=rows.filter(r=>!r.room).length;

  return (
    <div className="min-h-screen font-mono" style={{background:"linear-gradient(160deg,#07090f 0%,#0c0e1a 60%,#07090f 100%)"}}>
      {modal==="loading"&&<ImportModal type="loading" onDone={r=>{setRows(r);setModal(null);setTab("dashboard");flash(`${r.length} rows loaded`);}} onClose={()=>setModal(null)}/>}
      {modal==="rooms"&&<ImportModal type="rooms" onDone={r=>{setRooms(r);setModal(null);flash(`${r.length} rooms imported`);}} onClose={()=>setModal(null)}/>}
      {modal==="auto"&&<AutoModal rows={rows} onRun={runAuto} onClose={()=>setModal(null)}/>}
      {toast&&(
        <div style={{top:16,right:16,zIndex:100}} className={`fixed px-4 py-3 rounded-xl text-xs font-semibold shadow-2xl
          ${toast.type==="ok"?"bg-emerald-900/90 border border-emerald-700/50 text-emerald-200":"bg-red-900/90 border border-red-700/50 text-red-200"}`}>
          {toast.msg}
        </div>
      )}

      <header className="sticky top-0 z-20 border-b border-slate-800/60 bg-[#07090f]/96 backdrop-blur-md">
        <div className="flex items-center gap-3 px-4 py-3 max-w-screen-2xl mx-auto">
          <div className="flex items-center gap-2 mr-3 shrink-0">
            <div className="w-7 h-7 rounded-lg bg-indigo-600/20 border border-indigo-600/30 flex items-center justify-center text-sm">📋</div>
            <div className="hidden sm:block">
              <div className="text-[11px] font-bold text-white">Faculty Loading</div>
              <div className="text-[9px] text-slate-700">CAS · 1S AY 2026–27</div>
            </div>
          </div>
          <nav className="flex gap-0.5 flex-1 overflow-x-auto">
            {NAV.map(t=>(
              <button key={t.id} onClick={()=>setTab(t.id)}
                className={`flex items-center gap-1 px-3 py-1.5 rounded-lg text-xs whitespace-nowrap font-semibold transition-all
                  ${tab===t.id?"bg-indigo-600 text-white":"text-slate-600 hover:text-slate-300 hover:bg-slate-800/50"}`}>
                <span>{t.i}</span>{t.l}
                {t.id==="conflicts"&&conflicts.size>0&&<span className="bg-red-500 text-white text-[9px] rounded-full px-1.5 py-0.5 font-bold">{conflicts.size}</span>}
              </button>
            ))}
          </nav>
          <div className="flex gap-2 shrink-0 items-center">
            {rows.length>0&&<><Btn sm variant="amber" onClick={()=>setModal("auto")}>⚡ Auto-Schedule</Btn></>}
            <Btn sm variant="ghost" onClick={()=>setModal("rooms")}>🏫 Rooms</Btn>
            <Btn sm variant="primary" onClick={()=>setModal("loading")}>📂 Load Excel</Btn>
          </div>
        </div>
      </header>

      <main className="max-w-screen-2xl mx-auto px-4 py-5">
        {rows.length===0?(
          <div className="flex flex-col items-center justify-center min-h-[68vh] text-center gap-6">
            <div className="w-20 h-20 rounded-2xl bg-indigo-600/10 border border-indigo-600/20 flex items-center justify-center text-4xl">📊</div>
            <div>
              <h1 className="text-lg font-bold text-slate-200 mb-2">Faculty Loading & Class Schedules</h1>
              <p className="text-sm text-slate-600 max-w-lg leading-relaxed">
                Upload your Excel to begin. The app reads the three columns you fill manually:<br/>
                <strong className="text-white">C — Faculty</strong> · <strong className="text-white">W/X — Day/Time</strong> · <strong className="text-white">Z — Room</strong><br/>
                Then hit <strong className="text-indigo-400">⚡ Auto-Schedule</strong> to fill empty cells with smart AM/PM splitting,
                lunch protection, and fair rotation across subjects.
              </p>
            </div>
            <div className="flex gap-3 flex-wrap justify-center">
              <button onClick={()=>setModal("loading")} className="px-6 py-3 bg-indigo-600 hover:bg-indigo-500 text-white text-sm font-bold rounded-xl transition-colors shadow-lg shadow-indigo-900/30">📂 Upload Faculty Loading Excel</button>
              <button onClick={()=>setModal("rooms")} className="px-6 py-3 border border-slate-700 text-slate-400 text-sm font-semibold rounded-xl hover:border-slate-600 hover:text-slate-300 transition-colors">🏫 Upload Classroom Assignment</button>
            </div>
            <div className="flex gap-2 flex-wrap justify-center mt-2 max-w-2xl">
              {Object.entries(COLLEGES).map(([col,d])=>(
                <div key={col} style={{background:d.bg,borderColor:d.color+"30"}} className="border rounded-lg px-3 py-1.5 text-[11px]">
                  <span className="font-bold" style={{color:d.color}}>{col}</span>
                  <span className="text-slate-700 ml-1">{d.programs.filter(p=>!p.includes(" ")).slice(0,3).join(" · ")}</span>
                </div>
              ))}
            </div>
          </div>
        ):(
          <>
            {tab==="dashboard"&&<Dashboard rows={rows} conflicts={conflicts} onAuto={()=>setModal("auto")}/>}
            {tab==="loading"  &&<LoadingTable rows={rows} setRows={setRows} rooms={rooms} conflicts={conflicts}/>}
            {tab==="faculty"  &&<FacultySummary rows={rows} conflicts={conflicts}/>}
            {tab==="schedule" &&<ScheduleView rows={rows}/>}
            {tab==="conflicts"&&<ConflictsPanel rows={rows} conflicts={conflicts}/>}
            {tab==="export"   &&<ExportPanel rows={rows}/>}
          </>
        )}
      </main>
    </div>
  );
}