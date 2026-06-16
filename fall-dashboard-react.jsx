import { useState, useEffect, useCallback, useRef, useMemo } from "react";
import { BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

// ── Colors ────────────────────────────────────────────────────────────
const C = { navy:"#1B2A4A", mid:"#253660", teal:"#0E7C7B", tealLt:"#17A89E", gold:"#C9A84C", goldLt:"#E2C46A", red:"#C0392B", orange:"#E67E22", green:"#27AE60", slate:"#F0F3F8", muted:"#5E6E8C" };
const INJC = { None:C.green, Minor:C.gold, Moderate:C.orange, Major:C.red };
const PAL  = [C.teal,C.gold,C.navy,C.orange,C.red,C.tealLt,C.goldLt];

// ── Supabase ──────────────────────────────────────────────────────────
const SB_URL = "https://xnsdvdfceflmagfhpycw.supabase.co";
const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inhuc2R2ZGZjZWZsbWFnZmhweWN3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3MTI3NjI4NzcsImV4cCI6MjAyODMzODg3N30.hnfGJBDsLJxmBQDxOUfBNiMBxkQPW3eL8KdZAKXdH6k";
const TABLE  = "fall_events";
const sb     = createClient(SB_URL, SB_KEY);

// ── Column map ────────────────────────────────────────────────────────
const COL_MAP = {
  "facility":"facility","unit":"unit","patient_name":"patientName","patient_initials":"patientInitials","initials":"patientInitials",
  "sex":"sex","gender":"sex","age":"age","patient_age":"age","patient_bmi":"bmi",
  "visit_number":"csn","visit_number_/csn":"csn","mr_number":"mrNumber","event_number":"eventNumber","event_year":"eventYear",
  "event_date":"eventDate","event_time":"eventTime","injury_level":"injuryLevel",
  "rl_entered":"rlEntered","rl_event_entered?":"rlEntered","rl_date_entered":"rlEnteredDate","rl_time_entered":"rlEnteredTime",
  "evnet_description_from_rl":"eventDescription","event_description_from_rl":"eventDescription","event_description":"eventDescription",
  "post_falls_assessment_in_emr":"postFallAssessment","post_fall_assessment":"postFallAssessment",
  "post_fall_assessment_time":"postFallAssessmentTime",
  "fall_log_flowsheet_completed_in_emr":"fallLogFlowsheet","fall_log_flowsheet":"fallLogFlowsheet",
  "post_fall_vs":"postFallVitals","post_fall_vitals":"postFallVitals",
  "pain_assessment_post_fall":"painAssessment","pain_assessment":"painAssessment",
  "neuro_assessment_post_fall":"neuroAssessment","neuro_assessment":"neuroAssessment",
  "skin_assessment_post_fall":"skinAssessment","skin_assessment":"skinAssessment",
  "nurse_note_about_event":"nurseNote","provider_note_regarding_the_event":"providerNote","provider_note_about_event":"providerNote",
  "morse_score_pre__fall":"morseScore","morse_score_pre_fall":"morseScore","morse_score":"morseScore",
  "fall_risk_level":"fallRiskLevel","assisted?":"assisted","assisted":"assisted",
  "intoxication_impairment_contributing_factor":"intoxication","intoxication/_impairment":"intoxication",
  "activity_at_time_of_fall":"activityAtFall","activity_at_fall":"activityAtFall",
  "jh_hlm_score":"jhHlmScore","jh_hlm":"jhHlmScore",
  "recommended_device_utilized":"recommendedDevice","cause_of_fall":"causeOfFall",
  "send_for_analysis":"sendForAnalysis","area_leader":"areaLeader",
  "pertient_notes_about_event":"pertinentNotes","pertinent_notes_about_event":"pertinentNotes","pertinent_notes":"pertinentNotes",
  "post_fall_huddle_report_done":"huddleCompleted","huddle_completed":"huddleCompleted",
  "date_post_fall_huddle_report_received":"huddleReportDate",
  "days_to_receive_post_fall_huddle_report":"daysToHuddle","days_to_huddle":"daysToHuddle",
  "post_fall_huddle_manger_f_u_within_48_72_hours":"huddleManagerFU48h",
  "call_light_on_at_time_of_fall":"callLight","how_long_was_call_light_on_before_fall?":"howLong",
  "primary_rn":"primaryRN","primary_uap":"primaryUAP",
  "recommendations_based_on_aca_post_fall_huddle":"recommendations","recommendations_based_on_aca":"recommendations","recommendations":"recommendations",
  "patient's_bmi":"bmi",
};

function nk(s){ return s==null?"":String(s).trim().toLowerCase().replace(/[\s\/\-]+/g,"_"); }
function getInitials(n){ if(!n||n==="[Redacted]")return"—"; return n.trim().split(/\s+/).map(w=>w[0]?.toUpperCase()||"").join("").slice(0,4); }
function fmtVal(v){
  if(v==null||v==="")return"";
  if(typeof v==="boolean")return v?"Yes":"No";
  if(v instanceof Date)return v.toLocaleDateString("en-US",{month:"2-digit",day:"2-digit",year:"numeric"});
  return String(v);
}
function mapRow(raw){
  const o={};
  for(const [k,v] of Object.entries(raw)){
    if(k==null||k===""||typeof k==="number")continue;
    const ks=String(k).trim(); if(!ks)continue;
    const m=COL_MAP[nk(ks)]||nk(ks);
    o[m]=fmtVal(v);
  }
  if(!o.patientInitials&&o.patientName)o.patientInitials=getInitials(o.patientName);
  delete o.patientName;
  const d=new Date(o.eventDate); o._d=isNaN(d)?null:d;
  return o;
}
function rowKey(d){ return [d.facility||"",d.unit||"",String(d.event_date||d.eventDate||""),String(d.event_time||d.eventTime||"")].join("|").toLowerCase().trim(); }
function getMonthKey(d){ if(!d._d)return null; return d._d.toLocaleDateString("en-US",{month:"long",year:"numeric"}); }

// ── DB row → internal ─────────────────────────────────────────────────
function dbToInc(r){
  const o={
    facility:r.facility, unit:r.unit, patientInitials:r.patient_initials, sex:r.sex, age:r.age,
    eventDate:r.event_date, eventTime:r.event_time, injuryLevel:r.injury_level, rlEntered:r.rl_entered,
    fallRiskLevel:r.fall_risk_level, morseScore:r.morse_score, assisted:r.assisted,
    causeOfFall:r.cause_of_fall, activityAtFall:r.activity_at_fall, jhHlmScore:r.jh_hlm_score,
    huddleCompleted:r.huddle_completed, daysToHuddle:r.days_to_huddle,
    primaryRN:r.primary_rn, primaryUAP:r.primary_uap,
    painAssessment:r.pain_assessment, neuroAssessment:r.neuro_assessment,
    skinAssessment:r.skin_assessment, postFallVitals:r.post_fall_vitals,
    fallLogFlowsheet:r.fall_log_flowsheet, recommendations:r.recommendations,
    pertinentNotes:r.pertinent_notes, eventDescription:r.event_description,
    nurseNote:r.nurse_note, providerNote:r.provider_note,
    callLight:r.call_light, howLong:r.how_long, bmi:r.bmi,
  };
  const d=new Date(o.eventDate); o._d=isNaN(d)?null:d;
  return o;
}

// ── Summary generator ─────────────────────────────────────────────────
function generateSummary(inc){
  const ini=inc.patientInitials||"Patient", sex=inc.sex||"", age=inc.age||"?", unit=inc.unit||"?";
  const date=inc.eventDate||"?", time=inc.eventTime?` at ${inc.eventTime}`:"";
  const inj=inc.injuryLevel||"no documented injury", act=inc.activityAtFall||"unknown activity";
  const asst=(inc.assisted||"").toLowerCase();
  const risk=inc.fallRiskLevel||"?";
  const morse=inc.morseScore?`Morse score ${inc.morseScore}`:"";
  const jh=inc.jhHlmScore?`, JH-HLM ${inc.jhHlmScore}`:"";
  const rl=(inc.rlEntered||"").toLowerCase().includes("yes")?"RL event entered.":"RL event NOT entered.";
  const rn=inc.primaryRN?`Primary RN: ${inc.primaryRN}.`:"";
  const uap=inc.primaryUAP?`Primary UAP: ${inc.primaryUAP}.`:"";
  const checks=[["Post-fall vitals",inc.postFallVitals],["Pain assessment",inc.painAssessment],["Neuro assessment",inc.neuroAssessment],["Skin assessment",inc.skinAssessment],["Fall log flowsheet",inc.fallLogFlowsheet]];
  const missed=checks.filter(([,v])=>!v||(!String(v).toLowerCase().includes("complet")&&v!=="Yes"&&v!==true));
  const protocolLine=missed.length===0?"All post-fall protocol elements completed.":missed.length===checks.length?"Post-fall protocol elements were not completed.":`Partially completed — ${missed.map(([k])=>k).join(", ")} not documented.`;
  const huddleDone=(inc.huddleCompleted||"").toLowerCase().includes("yes")||inc.huddleCompleted===true;
  const days=inc.daysToHuddle!=null&&inc.daysToHuddle!==""?inc.daysToHuddle:null;
  const huddleLine=huddleDone?`Post-fall huddle completed${days!==null?` within ${days} day${days==1?"":"s"}`:""}. `:"Post-fall huddle not completed. ";
  const recLine=inc.recommendations?`Recommended: ${String(inc.recommendations).trim().slice(0,180)}.`:inc.pertinentNotes?`Notes: ${String(inc.pertinentNotes).trim().slice(0,180)}.`:"";
  const descLine=inc.eventDescription&&String(inc.eventDescription).length>10?`Event: ${String(inc.eventDescription).trim().slice(0,180)}${String(inc.eventDescription).length>180?"…":""}. `:"";
  return[
    `On ${date}${time}, ${ini} (${sex?sex+", ":""}age ${age}) on Unit ${unit} sustained a${inj.toLowerCase().startsWith("a")||inj.toLowerCase().startsWith("u")?"n":""} ${inj.toLowerCase()} fall while ${act.toLowerCase()}${asst.includes("un")||asst.includes("no")?", unassisted":asst.includes("yes")||asst.includes("asst")?", assisted":""}.`,
    descLine,
    `Classified as ${risk} fall risk (${morse}${jh}). ${rl} ${rn} ${uap}`.trim().replace(/\s+/g," "),
    protocolLine, huddleLine, recLine,
  ].filter(Boolean).join(" ");
}

// ── Demo data ─────────────────────────────────────────────────────────
const DEMO=[
  {facility:"AOMC",unit:"3B",patientInitials:"J.D.",sex:"M",age:"74",eventDate:"2024-01-08",eventTime:"02:15",injuryLevel:"Minor",rlEntered:"Yes",fallRiskLevel:"High",morseScore:"65",assisted:"No",causeOfFall:"Toileting",activityAtFall:"Ambulating to BR",jhHlmScore:"14",huddleCompleted:"Yes",daysToHuddle:"1",primaryRN:"Smith, J.",primaryUAP:"Jones, M.",painAssessment:"Completed",neuroAssessment:"Completed",skinAssessment:"Completed",postFallVitals:"Completed",recommendations:"Bed alarm reinstated; 1:1 for nights",_d:new Date("2024-01-08")},
  {facility:"AOMC",unit:"3B",patientInitials:"M.L.",sex:"F",age:"81",eventDate:"2024-01-22",eventTime:"14:40",injuryLevel:"None",rlEntered:"Yes",fallRiskLevel:"High",morseScore:"55",assisted:"Yes",causeOfFall:"Balance/Gait",activityAtFall:"Standing",jhHlmScore:"11",huddleCompleted:"Yes",daysToHuddle:"1",primaryRN:"Davis, T.",primaryUAP:"Brown, K.",painAssessment:"Completed",neuroAssessment:"Completed",skinAssessment:"Completed",postFallVitals:"Completed",recommendations:"PT consult placed",_d:new Date("2024-01-22")},
  {facility:"AOMC",unit:"3C",patientInitials:"R.K.",sex:"M",age:"68",eventDate:"2024-02-03",eventTime:"22:00",injuryLevel:"Moderate",rlEntered:"Yes",fallRiskLevel:"High",morseScore:"75",assisted:"No",causeOfFall:"Medication Effect",activityAtFall:"Getting OOB",jhHlmScore:"16",huddleCompleted:"Yes",daysToHuddle:"2",primaryRN:"Wilson, R.",primaryUAP:"Taylor, A.",painAssessment:"Completed",neuroAssessment:"Completed",skinAssessment:"Completed",postFallVitals:"Completed",recommendations:"Meds reviewed; low bed ordered",_d:new Date("2024-02-03")},
  {facility:"AOMC",unit:"3B",patientInitials:"B.T.",sex:"F",age:"77",eventDate:"2024-02-17",eventTime:"06:30",injuryLevel:"None",rlEntered:"Yes",fallRiskLevel:"Moderate",morseScore:"45",assisted:"Yes",causeOfFall:"Equipment Issue",activityAtFall:"Chair Transfer",jhHlmScore:"10",huddleCompleted:"No",daysToHuddle:"",primaryRN:"Garcia, L.",primaryUAP:"Martinez, S.",painAssessment:"Completed",neuroAssessment:"Completed",skinAssessment:"Completed",postFallVitals:"Completed",recommendations:"IV tubing rerouted",_d:new Date("2024-02-17")},
  {facility:"AOMC",unit:"3C",patientInitials:"E.W.",sex:"F",age:"89",eventDate:"2024-03-05",eventTime:"10:10",injuryLevel:"Minor",rlEntered:"Yes",fallRiskLevel:"High",morseScore:"60",assisted:"No",causeOfFall:"Confusion/Altered MS",activityAtFall:"Attempting OOB",jhHlmScore:"13",huddleCompleted:"Yes",daysToHuddle:"1",primaryRN:"Adams, B.",primaryUAP:"Clark, D.",painAssessment:"Not Completed",neuroAssessment:"Completed",skinAssessment:"Completed",postFallVitals:"Completed",recommendations:"Sitter ordered; reorientation q1h",_d:new Date("2024-03-05")},
];

// ── Styles ────────────────────────────────────────────────────────────
const S = {
  card:{background:"#fff",borderRadius:8,padding:16,border:"1px solid #E0E6EF",marginBottom:14},
  cardTitle:{fontWeight:700,fontSize:12,color:C.navy,marginBottom:10,textTransform:"uppercase",letterSpacing:".05em"},
  field:{borderBottom:"1px solid #F0F3F8",paddingBottom:4},
  fl:{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:".05em"},
  fv:{fontSize:12,color:C.navy,fontWeight:600},
};

// ── Sub-components ────────────────────────────────────────────────────
function KPI({label,value,sub,color}){
  return <div style={{background:"#fff",borderRadius:8,padding:"14px 16px",border:"1px solid #E0E6EF",borderTop:`4px solid ${color}`,flex:1,minWidth:130}}>
    <div style={{fontSize:10,color:C.muted,fontWeight:700,letterSpacing:".07em",textTransform:"uppercase",marginBottom:4}}>{label}</div>
    <div style={{fontSize:26,fontWeight:800,color,lineHeight:1.1}}>{value}</div>
    <div style={{fontSize:11,color:C.muted,marginTop:3}}>{sub}</div>
  </div>;
}

function Field({label,value}){
  return <div style={S.field}>
    <div style={S.fl}>{label}</div>
    <div style={S.fv}>{value||"—"}</div>
  </div>;
}

function IncidentSlide({inc,idx}){
  const [open,setOpen]=useState(true);
  const ic=INJC[inc.injuryLevel]||C.muted;
  const cf=["postFallVitals","painAssessment","neuroAssessment","skinAssessment","postFallAssessment"];
  const done=cf.filter(f=>inc[f]?.toString().toLowerCase().includes("complet")).length;
  const pct=Math.round(done/cf.length*100);
  const pc=pct===100?C.green:pct>=60?C.gold:C.red;
  const summary=useMemo(()=>generateSummary(inc),[inc]);
  const flds=[
    ["Initials",inc.patientInitials],["Sex",inc.sex],["Age",inc.age],["Unit",inc.unit],
    ["Date",inc.eventDate],["Time",inc.eventTime],["Injury",inc.injuryLevel],["Risk Level",inc.fallRiskLevel],
    ["Morse Score",inc.morseScore],["Cause",inc.causeOfFall],["Activity",inc.activityAtFall],["Assisted?",inc.assisted],
    ["JH-HLM",inc.jhHlmScore],["Primary RN",inc.primaryRN],["Primary UAP",inc.primaryUAP],["RL Entered?",inc.rlEntered],
    ["Huddle Done?",inc.huddleCompleted],["Days to Huddle",inc.daysToHuddle],
    ["Post Vitals",inc.postFallVitals],["Pain Assess.",inc.painAssessment],
    ["Neuro Assess.",inc.neuroAssessment],["Skin Assess.",inc.skinAssessment],
  ];
  return <div style={{background:"#fff",borderRadius:8,border:"1px solid #E0E6EF",overflow:"hidden",marginBottom:14}}>
    {/* Header */}
    <div style={{background:C.navy,padding:"12px 16px",display:"flex",alignItems:"center",gap:12}}>
      <div style={{background:ic,borderRadius:"50%",width:32,height:32,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontWeight:800,fontSize:13,color:"#fff"}}>{idx+1}</div>
      <div style={{flex:1,minWidth:0}}>
        <div style={{color:"#fff",fontWeight:700,fontSize:14}}>{inc.facility||"AOMC"} · Unit {inc.unit||"—"} · {inc.eventDate||"—"}</div>
        <div style={{color:C.goldLt,fontSize:11,marginTop:2}}>{inc.patientInitials||"—"} · {inc.sex||"—"} · Age {inc.age||"—"} · {inc.causeOfFall||"Unknown"}</div>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
        <button onClick={()=>setOpen(o=>!o)} style={{background:C.teal,color:"#fff",border:"none",borderRadius:5,padding:"5px 10px",fontSize:11,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap"}}>
          {open?"▲ Hide Summary":"✦ Summary"}
        </button>
        <div style={{background:`${ic}22`,border:`1px solid ${ic}66`,borderRadius:5,padding:"3px 10px",color:ic,fontWeight:700,fontSize:11}}>{inc.injuryLevel||"Unknown"}</div>
      </div>
    </div>
    {/* Summary panel */}
    {open && <div style={{background:"#EEF6FF",borderBottom:"1px solid #C8DDEF",padding:"12px 16px"}}>
      <div style={{fontSize:9,color:C.mid,fontWeight:800,textTransform:"uppercase",letterSpacing:".07em",marginBottom:6}}>Clinical Summary</div>
      <p style={{fontSize:13,lineHeight:1.65,color:C.navy,margin:0}}>{summary}</p>
    </div>}
    {/* Body */}
    <div style={{padding:"14px 16px",display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"5px 10px"}}>
        {flds.map(([k,v])=><Field key={k} label={k} value={v}/>)}
      </div>
      <div>
        <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:5}}>Post-Fall Assessment Compliance</div>
        <div style={{background:"#E0E6EF",borderRadius:20,height:10,overflow:"hidden",marginBottom:4}}>
          <div style={{width:`${pct}%`,background:pc,height:"100%",borderRadius:20}}/>
        </div>
        <div style={{fontSize:12,fontWeight:700,color:pc,marginBottom:12}}>{pct}% — {done}/{cf.length} Complete</div>
        {inc.recommendations&&<div style={{background:"#F7F9FC",border:"1px solid #D0DAE8",borderRadius:6,padding:10,marginBottom:8}}>
          <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",marginBottom:5}}>Recommendations</div>
          <div style={{fontSize:12,lineHeight:1.5}}>{inc.recommendations}</div>
        </div>}
        {inc.pertinentNotes&&<div style={{background:"#FFFBF0",border:`1px solid ${C.goldLt}`,borderRadius:6,padding:10}}>
          <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",marginBottom:5}}>Pertinent Notes</div>
          <div style={{fontSize:12,lineHeight:1.5}}>{inc.pertinentNotes}</div>
        </div>}
      </div>
    </div>
  </div>;
}

// ── Main App ──────────────────────────────────────────────────────────
export default function App(){
  const [allData,setAllData]=useState(DEMO);
  const [tab,setTab]=useState("overview");
  const [fMonth,setFMonth]=useState("ALL");
  const [fUnit,setFUnit]=useState("ALL");
  const [fInj,setFInj]=useState("ALL");
  const [sbStatus,setSbStatus]=useState("loading");
  const [uploadMsg,setUploadMsg]=useState(null);
  const [isDrag,setIsDrag]=useState(false);

  // Supabase load on mount
  useEffect(()=>{
    (async()=>{
      try{
        const {data,error}=await sb.from(TABLE).select("*").order("event_date",{ascending:true});
        if(error)throw error;
        if(data&&data.length){
          const rows=data.map(dbToInc);
          setAllData(prev=>{
            const existing=new Set(prev.map(rowKey));
            const news=rows.filter(r=>!existing.has(rowKey(r)));
            return [...prev,...news];
          });
        }
        setSbStatus("connected");
      }catch(e){ setSbStatus("error"); console.error("SB load:",e); }
    })();
  },[]);

  // Derived
  const months=useMemo(()=>{
    const m={};
    allData.forEach(d=>{if(d._d){const k=getMonthKey(d);if(k)m[k]=d._d.getTime();}});
    return Object.keys(m).sort((a,b)=>m[a]-m[b]);
  },[allData]);

  const units=useMemo(()=>[...new Set(allData.map(d=>d.unit).filter(Boolean))],[allData]);
  const injuries=useMemo(()=>[...new Set(allData.map(d=>d.injuryLevel).filter(Boolean))],[allData]);

  const filtered=useMemo(()=>allData.filter(d=>
    (fMonth==="ALL"||getMonthKey(d)===fMonth)&&
    (fUnit==="ALL"||d.unit===fUnit)&&
    (fInj==="ALL"||d.injuryLevel===fInj)
  ),[allData,fMonth,fUnit,fInj]);

  // KPIs
  const total=filtered.length;
  const withInj=filtered.filter(d=>["Minor","Moderate","Major"].includes(d.injuryLevel)).length;
  const injRate=total?Math.round(withInj/total*100):0;
  const huddled=filtered.filter(d=>d.huddleCompleted?.toString().toLowerCase().includes("yes")).length;
  const huddleRate=total?Math.round(huddled/total*100):0;
  const avgMorse=total?Math.round(filtered.reduce((s,d)=>s+(parseFloat(d.morseScore)||0),0)/total):0;

  // Chart data
  const byUnit=useMemo(()=>{const m={};filtered.forEach(d=>{const k=d.unit||"Unknown";m[k]=(m[k]||0)+1;});return Object.entries(m).map(([unit,count])=>({unit,count}));},[filtered]);
  const byInj=useMemo(()=>{const m={};filtered.forEach(d=>{const k=d.injuryLevel||"Unknown";m[k]=(m[k]||0)+1;});return Object.entries(m).map(([name,value])=>({name,value}));},[filtered]);
  const byTime=useMemo(()=>{const b={"00-06":0,"06-12":0,"12-18":0,"18-24":0};filtered.forEach(d=>{const h=parseInt((d.eventTime||"").split(":")[0])||0;if(h<6)b["00-06"]++;else if(h<12)b["06-12"]++;else if(h<18)b["12-18"]++;else b["18-24"]++;});return Object.entries(b).map(([period,count])=>({period,count}));},[filtered]);
  const byMonth=useMemo(()=>{const m={};filtered.forEach(d=>{if(!d._d)return;const k=d._d.toLocaleDateString("en-US",{month:"short",year:"2-digit"});m[k]=(m[k]||0)+1;});return Object.entries(m).map(([month,count])=>({month,count}));},[filtered]);
  const byCause=useMemo(()=>{const m={};filtered.forEach(d=>{const k=d.causeOfFall||"Unknown";m[k]=(m[k]||0)+1;});return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,7);},[filtered]);
  const morseL=filtered.filter(d=>(parseFloat(d.morseScore)||0)<25).length;
  const morseM=filtered.filter(d=>{const s=parseFloat(d.morseScore)||0;return s>=25&&s<45;}).length;
  const morseH=filtered.filter(d=>(parseFloat(d.morseScore)||0)>=45).length;

  // Save to Supabase
  const saveRows=useCallback(async(rows)=>{
    setSbStatus("saving");
    try{
      const payload=rows.map(r=>({
        facility:r.facility||null, unit:r.unit||null, patient_initials:r.patientInitials||null,
        sex:r.sex||null, age:r.age||null, event_date:r.eventDate||null,
        event_time:r.eventTime?String(r.eventTime):null, injury_level:r.injuryLevel||null,
        rl_entered:r.rlEntered||null, fall_risk_level:r.fallRiskLevel||null,
        morse_score:r.morseScore||null, assisted:r.assisted||null,
        cause_of_fall:r.causeOfFall||null, activity_at_fall:r.activityAtFall||null,
        jh_hlm_score:r.jhHlmScore||null, huddle_completed:r.huddleCompleted||null,
        days_to_huddle:r.daysToHuddle||null, primary_rn:r.primaryRN||null,
        primary_uap:r.primaryUAP||null, pain_assessment:r.painAssessment||null,
        neuro_assessment:r.neuroAssessment||null, skin_assessment:r.skinAssessment||null,
        post_fall_vitals:r.postFallVitals||null, fall_log_flowsheet:r.fallLogFlowsheet||null,
        recommendations:r.recommendations||null, pertinent_notes:r.pertinentNotes||null,
        event_description:r.eventDescription||null, nurse_note:r.nurseNote||null,
        provider_note:r.providerNote||null, call_light:r.callLight||null,
        bmi:r.bmi||null, event_number:r.eventNumber||null, event_year:r.eventYear||null,
      }));
      const {error}=await sb.from(TABLE).upsert(payload,{onConflict:"facility,unit,event_date,event_time",ignoreDuplicates:false});
      if(error)throw error;
      setSbStatus("connected");
    }catch(e){ setSbStatus("error"); console.error("SB save:",e); }
  },[]);

  // File import
  const processFile=useCallback((file)=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"binary",cellDates:true});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const raw=XLSX.utils.sheet_to_json(ws,{defval:""});
        if(!raw.length){setUploadMsg({text:"⚠️ No data rows found.",type:"warn"});return;}
        const mapped=raw.map(mapRow);
        setAllData(prev=>{
          const existing=new Set(prev.map(rowKey));
          const newRows=mapped.filter(d=>!existing.has(rowKey(d)));
          const dupes=mapped.length-newRows.length;
          if(!newRows.length){setUploadMsg({text:`⚠️ No new records — ${dupes} duplicate${dupes!==1?"s":""} skipped.`,type:"warn"});return prev;}
          setUploadMsg({text:`✅ Added ${newRows.length} record${newRows.length!==1?"s":""}${dupes?` · ${dupes} duplicate${dupes!==1?"s":""} skipped`:""}`,type:"ok"});
          saveRows(newRows);
          return [...prev,...newRows];
        });
      }catch(err){setUploadMsg({text:"❌ Parse error: "+err.message,type:"err"});}
    };
    reader.readAsBinaryString(file);
  },[saveRows]);

  const sbColor=sbStatus==="connected"?C.teal:sbStatus==="saving"||sbStatus==="loading"?C.gold:sbStatus==="error"?C.red:C.muted;
  const sbLabel={"connected":"☁ Supabase Connected","saving":"☁ Saving…","loading":"☁ Loading…","error":"☁ DB Error","disconnected":"☁ No DB"}[sbStatus]||"☁ …";

  const tabStyle=(id)=>({padding:"10px 16px",background:"none",border:"none",borderBottom:tab===id?`3px solid ${C.goldLt}`:"3px solid transparent",color:tab===id?C.goldLt:"#94A8C8",fontWeight:tab===id?700:500,fontSize:13,cursor:"pointer",whiteSpace:"nowrap"});
  const selStyle={padding:"6px 10px",borderRadius:5,border:`1px solid ${C.mid}`,background:C.mid,color:"#fff",fontSize:12,fontWeight:600};

  return <div style={{fontFamily:"'Segoe UI',sans-serif",background:C.slate,minHeight:"100vh",color:C.navy,fontSize:14}}>

    {/* Header */}
    <div style={{background:C.navy,padding:"12px 16px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
      <div>
        <div style={{color:C.gold,fontSize:9,fontWeight:800,letterSpacing:".18em",textTransform:"uppercase"}}>ARNOT HEALTH · AOMC</div>
        <div style={{color:"#fff",fontSize:17,fontWeight:800}}>Fall Event Reporting Tracker</div>
      </div>
      <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
        <div onClick={()=>{setSbStatus("loading");sb.from(TABLE).select("*").order("event_date",{ascending:true}).then(({data,error})=>{if(!error&&data){const rows=data.map(dbToInc);setAllData(rows.map(r=>{const d=new Date(r.eventDate);r._d=isNaN(d)?null:d;return r;}));}setSbStatus(error?"error":"connected");})}}
          style={{background:sbColor,color:"#fff",fontSize:11,fontWeight:700,padding:"6px 10px",borderRadius:5,cursor:"pointer",whiteSpace:"nowrap",userSelect:"none"}}>{sbLabel}</div>
        <span style={{color:C.goldLt,fontSize:11,fontWeight:700,letterSpacing:".06em",textTransform:"uppercase"}}>Review Month</span>
        <select value={fMonth} onChange={e=>setFMonth(e.target.value)} style={{...selStyle,background:C.gold,color:C.navy,borderColor:C.gold,fontWeight:800}}>
          <option value="ALL">ALL Months</option>
          {months.map(m=><option key={m} value={m}>{m}</option>)}
        </select>
        <select value={fUnit} onChange={e=>setFUnit(e.target.value)} style={selStyle}>
          <option value="ALL">ALL Units</option>
          {units.map(u=><option key={u}>{u}</option>)}
        </select>
        <select value={fInj} onChange={e=>setFInj(e.target.value)} style={selStyle}>
          <option value="ALL">ALL Injuries</option>
          {injuries.map(i=><option key={i}>{i}</option>)}
        </select>
      </div>
    </div>

    {/* Tab bar */}
    <div style={{background:C.mid,display:"flex",padding:"0 16px",overflowX:"auto",gap:2}}>
      {[["overview","📊 Overview"],["trends","📈 Trends"],["incidents",`📋 Incidents (${filtered.length})`],["upload","⬆️ Upload"]].map(([id,label])=>
        <button key={id} style={tabStyle(id)} onClick={()=>setTab(id)}>{label}</button>
      )}
    </div>

    {/* Content */}
    <div style={{padding:"16px",maxWidth:1080,margin:"0 auto"}}>

      {/* OVERVIEW */}
      {tab==="overview" && <>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:16}}>
          <KPI label="Total Falls" value={total} sub="Filtered period" color={C.navy}/>
          <KPI label="Injury Rate" value={`${injRate}%`} sub={`${withInj} with injury`} color={injRate>30?C.red:C.gold}/>
          <KPI label="Huddle Compliance" value={`${huddleRate}%`} sub={`${huddled}/${total} completed`} color={huddleRate===100?C.green:huddleRate>=80?C.teal:C.orange}/>
          <KPI label="Avg Morse Score" value={avgMorse} sub="Pre-fall score" color={avgMorse>=60?C.red:C.teal}/>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
          <div style={S.card}><div style={S.cardTitle}>Falls by Unit</div>
            <ResponsiveContainer width="100%" height={180}>
              <BarChart data={byUnit}><CartesianGrid strokeDasharray="3 3" stroke="#E0E6EF"/><XAxis dataKey="unit" tick={{fontSize:11}}/><YAxis tick={{fontSize:11}}/><Tooltip/><Bar dataKey="count" fill={C.teal} radius={[4,4,0,0]}/></BarChart>
            </ResponsiveContainer></div>
          <div style={S.card}><div style={S.cardTitle}>Injury Level Distribution</div>
            <ResponsiveContainer width="100%" height={180}>
              <PieChart><Pie data={byInj} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={70} label={({name,value})=>`${name}: ${value}`} labelLine={false} fontSize={10}>
                {byInj.map(({name},i)=><Cell key={name} fill={INJC[name]||PAL[i%PAL.length]}/>)}
              </Pie><Tooltip/></PieChart>
            </ResponsiveContainer></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          <div style={S.card}><div style={S.cardTitle}>Top Causes of Fall</div>
            {byCause.map(([name,val],i)=><div key={name} style={{display:"flex",alignItems:"center",gap:8,marginBottom:7}}>
              <div style={{width:8,height:8,borderRadius:"50%",background:PAL[i%PAL.length],flexShrink:0}}/>
              <div style={{flex:1,fontSize:12}}>{name}</div>
              <div style={{background:"#F0F3F8",borderRadius:4,padding:"2px 7px",fontSize:11,fontWeight:700}}>{val}</div>
            </div>)}
          </div>
          <div style={S.card}><div style={S.cardTitle}>Falls by Time of Day</div>
            <ResponsiveContainer width="100%" height={180}>
              <BarChart data={byTime}><CartesianGrid strokeDasharray="3 3" stroke="#E0E6EF"/><XAxis dataKey="period" tick={{fontSize:11}}/><YAxis tick={{fontSize:11}}/><Tooltip/><Bar dataKey="count" fill={C.gold} radius={[4,4,0,0]}/></BarChart>
            </ResponsiveContainer></div>
        </div>
      </>}

      {/* TRENDS */}
      {tab==="trends" && <>
        <div style={S.card}><div style={S.cardTitle}>Monthly Fall Volume</div>
          <ResponsiveContainer width="100%" height={200}>
            <LineChart data={byMonth}><CartesianGrid strokeDasharray="3 3" stroke="#E0E6EF"/><XAxis dataKey="month" tick={{fontSize:11}}/><YAxis tick={{fontSize:11}}/><Tooltip/><Line type="monotone" dataKey="count" stroke={C.teal} strokeWidth={2.5} dot={{r:4,fill:C.teal}}/></LineChart>
          </ResponsiveContainer></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          <div style={S.card}><div style={S.cardTitle}>Huddle Compliance</div>
            <ResponsiveContainer width="100%" height={180}>
              <PieChart><Pie data={[{name:"Completed",value:huddled},{name:"Pending",value:total-huddled}]} dataKey="value" cx="50%" cy="50%" outerRadius={70} label={({name,value})=>`${name}: ${value}`} labelLine={false} fontSize={11}>
                <Cell fill={C.green}/><Cell fill="#E0E6EF"/>
              </Pie><Tooltip/></PieChart>
            </ResponsiveContainer>
            <div style={{textAlign:"center",fontSize:13,fontWeight:700,color:huddleRate>=90?C.green:C.orange,marginTop:8}}>{huddleRate}% Completion</div>
          </div>
          <div style={S.card}><div style={S.cardTitle}>Morse Score Distribution</div>
            {[["Low Risk (<25)",morseL,C.green],["Moderate (25–44)",morseM,C.gold],["High Risk (45+)",morseH,C.red]].map(([label,val,color])=><div key={label} style={{marginBottom:10}}>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}><span>{label}</span><span style={{fontWeight:700,color}}>{val}</span></div>
              <div style={{background:"#E0E6EF",borderRadius:20,height:8,overflow:"hidden"}}><div style={{width:`${total?val/total*100:0}%`,background:color,height:"100%",borderRadius:20}}/></div>
            </div>)}
          </div>
        </div>
      </>}

      {/* INCIDENTS */}
      {tab==="incidents" && <>
        <div style={{fontWeight:700,fontSize:14,marginBottom:12}}>Incident Summary — {filtered.length} Event{filtered.length!==1?"s":""}</div>
        {filtered.length===0
          ? <div style={{textAlign:"center",padding:32,color:C.muted}}>No incidents match current filters.</div>
          : filtered.map((inc,i)=><IncidentSlide key={rowKey(inc)||i} inc={inc} idx={i}/>)
        }
      </>}

      {/* UPLOAD */}
      {tab==="upload" && <div style={{maxWidth:580}}>
        <div style={{fontWeight:700,fontSize:15,marginBottom:8}}>Import Fall Event Data</div>
        <p style={{fontSize:13,color:C.muted,marginBottom:14}}>Upload your Excel export. Columns are auto-mapped. Only new records are added.</p>
        <div
          onDrop={e=>{e.preventDefault();setIsDrag(false);if(e.dataTransfer.files[0])processFile(e.dataTransfer.files[0]);}}
          onDragOver={e=>{e.preventDefault();setIsDrag(true);}}
          onDragLeave={()=>setIsDrag(false)}
          style={{border:`2px dashed ${isDrag?C.teal:"#C0CDD8"}`,borderRadius:8,background:isDrag?"#EAF8F7":"#fff",padding:"32px 20px",textAlign:"center",transition:".2s"}}>
          <div style={{fontSize:32,marginBottom:8}}>📂</div>
          <div style={{fontWeight:700,marginBottom:4}}>Drag & drop your file here</div>
          <div style={{fontSize:12,color:C.muted,marginBottom:14}}>Supports .xlsx and .csv</div>
          <label style={{background:C.teal,color:"#fff",padding:"8px 18px",borderRadius:6,fontWeight:700,fontSize:13,cursor:"pointer"}}>
            Browse File<input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>{if(e.target.files[0])processFile(e.target.files[0]);}}/>
          </label>
        </div>
        {uploadMsg && <div style={{marginTop:12,padding:"9px 14px",borderRadius:6,fontSize:13,fontWeight:600,background:uploadMsg.type==="ok"?"#EAF8EF":uploadMsg.type==="warn"?"#FFFBF0":"#FEF0F0",border:`1px solid ${uploadMsg.type==="ok"?C.green:uploadMsg.type==="warn"?C.gold:C.red}`}}>{uploadMsg.text}</div>}
        <div style={{...S.card,marginTop:16}}>
          <div style={{fontWeight:700,fontSize:12,marginBottom:8}}>Expected Column Headers</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
            {["Facility","Unit","Patient Name","Event Date","Event Time","Injury Level","RL Entered","Morse Score Pre- Fall","Fall Risk Level","Assisted?","Cause of Fall","Activity at Time of Fall","JH-HLM Score","Primary RN","Primary UAP","Post Fall Huddle Report Done","Days to Receive Post Fall Huddle Report","Recommendations Based on ACA Post Fall Huddle"].map(t=><span key={t} style={{background:"#F0F3F8",border:"1px solid #D0DAE8",borderRadius:3,padding:"2px 6px",fontSize:10,color:C.muted,fontFamily:"monospace"}}>{t}</span>)}
          </div>
        </div>
        <div style={{display:"flex",gap:10,marginTop:12}}>
          <button onClick={()=>{setAllData([]);setUploadMsg({text:"Data cleared.",type:"warn"});}} style={{padding:"7px 14px",borderRadius:6,border:`1px solid ${C.red}`,background:"#fff",color:C.red,fontWeight:700,fontSize:12,cursor:"pointer"}}>Clear All Data</button>
          <button onClick={()=>{setAllData(DEMO);setUploadMsg({text:"Demo data restored.",type:"ok"});}} style={{padding:"7px 14px",borderRadius:6,border:`1px solid ${C.teal}`,background:"#fff",color:C.teal,fontWeight:700,fontSize:12,cursor:"pointer"}}>Restore Demo Data</button>
        </div>
      </div>}

    </div>
  </div>;
}
