import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer } from "recharts";
import {
  FileText, Upload, Check, X, Clock, DollarSign, Plane,
  ChevronRight, Plus, Trash2, Eye, Send,
  Users, Home, AlertTriangle, CheckCircle2,
  XCircle, ArrowRight,
  Receipt, Globe, Edit3,
  ChevronLeft, AlertCircle, RefreshCw, Wifi, WifiOff,
  Download, Image, Paperclip, Shield, FileSpreadsheet, Save,
  Search, Mail, Building2, UserPlus, Pencil, ArrowLeftRight, LogIn, Lock, Copy
} from "lucide-react";

const T = {
  bg: "#F0F2F5", card: "#FFFFFF",
  navy: "#0A1E3D", navyL: "#142D54",
  blue: "#2563EB", blueL: "#3B82F6", blueA: "rgba(37,99,235,0.08)",
  green: "#059669", greenL: "#10B981", greenA: "rgba(5,150,105,0.08)",
  amber: "#D97706", amberA: "rgba(217,119,6,0.08)",
  red: "#DC2626", redA: "rgba(220,38,38,0.08)",
  purple: "#7C3AED", purpleA: "rgba(124,58,237,0.08)",
  pink: "#DB2777", teal: "#0D9488",
  g0:"#F8FAFC",g1:"#F1F5F9",g2:"#E2E8F0",g3:"#CBD5E1",g4:"#94A3B8",g5:"#64748B",g6:"#475569",g7:"#334155",g8:"#1E293B",
  sh1:"0 1px 3px rgba(10,30,61,0.06),0 1px 2px rgba(10,30,61,0.04)",
  sh2:"0 4px 12px rgba(10,30,61,0.08),0 1px 3px rgba(10,30,61,0.05)",
  sh3:"0 10px 30px rgba(10,30,61,0.12),0 4px 8px rgba(10,30,61,0.06)",
  r:16, rs:10, rx:6,
  font:"'Plus Jakarta Sans','DM Sans','Segoe UI',system-ui,sans-serif",
};

const CURRENCIES=[{code:"USD",name:"Dolar EE.UU.",rateToCOP:4250,symbol:"$"},{code:"EUR",name:"Euro",rateToCOP:4600,symbol:"\u20AC"},{code:"COP",name:"Peso Colombiano",rateToCOP:1,symbol:"$"},{code:"MXN",name:"Peso Mexicano",rateToCOP:248,symbol:"MXN"},{code:"BRL",name:"Real Brasileno",rateToCOP:860,symbol:"R$"},{code:"ARS",name:"Peso Argentino",rateToCOP:3.8,symbol:"ARS"},{code:"CLP",name:"Peso Chileno",rateToCOP:4.5,symbol:"CLP"},{code:"PEN",name:"Sol Peruano",rateToCOP:1142,symbol:"S/"},{code:"GBP",name:"Libra Esterlina",rateToCOP:5380,symbol:"\u00A3"},{code:"CAD",name:"Dolar Canadiense",rateToCOP:3132,symbol:"C$"},{code:"CHF",name:"Franco Suizo",rateToCOP:4830,symbol:"CHF"}];
const CATS=[{id:"tiquete",label:"Tiquete Aereo",icon:"\u2708\uFE0F",color:T.blue},{id:"alojamiento",label:"Alojamiento + Tax",icon:"\uD83C\uDFE8",color:T.purple},{id:"alimentacion",label:"Alimentacion",icon:"\uD83C\uDF7D\uFE0F",color:T.amber},{id:"transporte",label:"Transporte",icon:"\uD83D\uDE95",color:T.greenL},{id:"invitaciones",label:"Invitaciones Clientes",icon:"\uD83E\uDD1D",color:T.pink},{id:"otros",label:"Otros Gastos",icon:"\uD83D\uDCCB",color:T.g5}];
const ETYPES=[{id:"anticipo",label:"Anticipo en Dinero"},{id:"tarjeta",label:"Tarjeta Credito Corp."}];
const ROLES=[
  {id:"employee",label:"Empleado",color:T.g5,desc:"Registra y envia reportes"},
  {id:"intern",label:"Practicante",color:"#78716C",desc:"Pasante o practicante"},
  {id:"assistant",label:"Asistente",color:"#A3A3A3",desc:"Asistente administrativo"},
  {id:"analyst",label:"Analista",color:"#0EA5E9",desc:"Analista de area"},
  {id:"senior_analyst",label:"Analista Senior",color:"#0284C7",desc:"Analista con experiencia"},
  {id:"specialist",label:"Especialista",color:T.teal,desc:"Especialista tecnico"},
  {id:"coordinator",label:"Coordinador",color:"#8B5CF6",desc:"Coordina equipo operativo"},
  {id:"team_leader",label:"Team Leader",color:T.blue,desc:"Aprueba reportes del equipo"},
  {id:"supervisor",label:"Supervisor",color:"#6366F1",desc:"Supervisa operaciones"},
  {id:"manager",label:"Gerente de Area",color:"#4F46E5",desc:"Gerente funcional"},
  {id:"admin",label:"Administracion",color:T.amber,desc:"Revisa y valida gastos"},
  {id:"hr",label:"Recursos Humanos",color:T.pink,desc:"Gestion de personal"},
  {id:"accountant",label:"Contabilidad",color:T.green,desc:"Contabiliza y verifica"},
  {id:"treasury",label:"Tesoreria",color:"#B45309",desc:"Manejo de fondos"},
  {id:"auditor",label:"Auditor Interno",color:"#EA580C",desc:"Audita cumplimiento"},
  {id:"compliance",label:"Compliance",color:"#9333EA",desc:"Cumplimiento normativo"},
  {id:"legal",label:"Legal",color:"#C026D3",desc:"Asuntos juridicos"},
  {id:"it_admin",label:"Admin TI",color:"#0891B2",desc:"Soporte tecnologico"},
  {id:"director",label:"Director",color:"#1D4ED8",desc:"Director de division"},
  {id:"vp",label:"Vicepresidente",color:T.navyL,desc:"VP de la organizacion"},
  {id:"general_manager",label:"Gerente General",color:T.navy,desc:"Aprobacion final"},
];
const AREAS=["Underwriting","Administracion","Claims","Finanzas","TI","Legal","Gerencia","Recursos Humanos","Compliance","Auditoria","Tesoreria","Actuaria","Comercial"];
const INIT_USERS=[
  {id:1,name:"Oscar Avila",role:"employee",area:"Administracion",cc:"79749093",email:"oscar.avila@hannover-re.com",phone:"+57 310 555 1234"},
  {id:2,name:"Juan Medina",role:"analyst",area:"Underwriting",cc:"80123456",email:"juan.medina@hannover-re.com",phone:"+57 311 555 5678"},
  {id:3,name:"Carolina Lopez",role:"team_leader",area:"Underwriting",cc:"52987654",email:"carolina.lopez@hannover-re.com",phone:"+57 312 555 9012"},
  {id:4,name:"Maria Uribe",role:"admin",area:"Administracion",cc:"51456789",email:"maria.uribe@hannover-re.com",phone:"+57 313 555 3456"},
  {id:5,name:"Andrea Ruiz",role:"accountant",area:"Finanzas",cc:"53112233",email:"andrea.ruiz@hannover-re.com",phone:"+57 314 555 7890"},
  {id:6,name:"Miguel Guarin",role:"general_manager",area:"Gerencia",cc:"79876543",email:"miguel.guarin@hannover-re.com",phone:"+57 315 555 2345"},
  {id:7,name:"Laura Martinez",role:"compliance",area:"Compliance",cc:"52234567",email:"laura.martinez@hannover-re.com",phone:"+57 316 555 6789"},
  {id:8,name:"Diego Torres",role:"auditor",area:"Auditoria",cc:"80345678",email:"diego.torres@hannover-re.com",phone:"+57 317 555 0123"},
];
const AFLOW=[{step:1,role:"employee",label:"Empleado",action:"Envio"},{step:2,role:"team_leader",label:"Jefe Directo",action:"Vo.Bo."},{step:3,role:"admin",label:"Administracion",action:"Vo.Bo."},{step:4,role:"accountant",label:"Contabilidad",action:"Vo.Bo."},{step:5,role:"general_manager",label:"Gerente General",action:"Aprobacion Final"}];
const SMAP={draft:{label:"Borrador",color:T.g5,bg:T.g1},submitted:{label:"Enviado",color:T.blue,bg:T.blueA},approved_leader:{label:"Aprobado Jefe",color:T.purple,bg:T.purpleA},approved_admin:{label:"Aprobado Admin",color:T.amber,bg:T.amberA},approved_accountant:{label:"Aprobado Contab.",color:T.greenL,bg:T.greenA},approved_final:{label:"Aprobado Final",color:T.green,bg:"rgba(5,150,105,0.12)"},rejected:{label:"Rechazado",color:T.red,bg:T.redA}};
const SAMPLE=[
  {id:"RPT-2025-001",employeeId:1,employeeName:"Oscar Avila",area:"Administracion",cc:"79749093",destination:"Mexico",tripPurpose:"Visita clientes",dateFrom:"2025-05-25",dateTo:"2025-06-07",presentationDate:"2025-06-18",currency:"USD",costCenter:"401",status:"approved_leader",currentStep:3,type:"anticipo",expenses:[{id:1,date:"2025-05-25",currency:"USD",category:"alojamiento",amount:890,obs:"Hotel Ciudad de Mexico - 7 noches",hasReceipt:true,attachments:[{name:"factura_hotel.pdf",size:"245 KB",type:"pdf"}]},{id:2,date:"2025-05-24",currency:"COP",category:"alimentacion",amount:20000,obs:"Cafe aeropuerto El Dorado",hasReceipt:true,attachments:[{name:"recibo_cafe.jpg",size:"120 KB",type:"img"}]},{id:3,date:"2025-05-25",currency:"USD",category:"transporte",amount:200,obs:"Uber Bogota y Mexico",hasReceipt:true,attachments:[{name:"uber1.pdf",size:"80 KB",type:"pdf"},{name:"uber2.pdf",size:"85 KB",type:"pdf"}]},{id:4,date:"2025-05-26",currency:"USD",category:"invitaciones",amount:85,obs:"Cena cliente Seguros Atlas",hasReceipt:true,attachments:[{name:"rest.jpg",size:"310 KB",type:"img"}]},{id:5,date:"2025-05-28",currency:"MXN",category:"alimentacion",amount:450,obs:"Almuerzo zona Rosa",hasReceipt:true,attachments:[{name:"ticket.jpg",size:"95 KB",type:"img"}]},{id:6,date:"2025-05-30",currency:"USD",category:"otros",amount:35,obs:"Lavanderia hotel",hasReceipt:false,attachments:[]}],approvals:[{step:1,by:"Oscar Avila",date:"2025-06-18",comment:""},{step:2,by:"Carolina Lopez",date:"2025-06-19",comment:"Aprobado."}],createdAt:"2025-06-18"},
  {id:"RPT-2025-002",employeeId:2,employeeName:"Juan Medina",area:"Underwriting",cc:"80123456",destination:"Medellin",tripPurpose:"Evento SURA",dateFrom:"2025-11-18",dateTo:"2025-11-20",presentationDate:"2025-12-18",currency:"COP",costCenter:"301",status:"submitted",currentStep:2,type:"tarjeta",expenses:[{id:1,date:"2025-11-18",currency:"COP",category:"transporte",amount:45200,obs:"Uber Aeropuerto MDE",hasReceipt:true,attachments:[{name:"uber.pdf",size:"78 KB",type:"pdf"}]},{id:2,date:"2025-11-18",currency:"COP",category:"alojamiento",amount:380000,obs:"Hotel El Tesoro - 2 noches",hasReceipt:true,attachments:[{name:"hotel.pdf",size:"320 KB",type:"pdf"}]},{id:3,date:"2025-11-19",currency:"COP",category:"invitaciones",amount:250000,obs:"Invitacion SURA evento",hasReceipt:true,attachments:[{name:"evento.pdf",size:"150 KB",type:"pdf"}]},{id:4,date:"2025-11-20",currency:"COP",category:"transporte",amount:52300,obs:"Uber a Aeropuerto",hasReceipt:true,attachments:[{name:"uber2.pdf",size:"75 KB",type:"pdf"}]},{id:5,date:"2025-11-20",currency:"COP",category:"tiquete",amount:485000,obs:"Avianca BOG-MDE-BOG",hasReceipt:true,attachments:[{name:"eticket.pdf",size:"205 KB",type:"pdf"}]}],approvals:[{step:1,by:"Juan Medina",date:"2025-12-18",comment:""}],createdAt:"2025-12-18"},
  {id:"RPT-2025-003",employeeId:2,employeeName:"Juan Medina",area:"Underwriting",cc:"80123456",destination:"Cartagena",tripPurpose:"Congreso Fasecolda",dateFrom:"2025-09-10",dateTo:"2025-09-12",presentationDate:"2025-09-17",currency:"COP",costCenter:"301",status:"approved_final",currentStep:5,type:"anticipo",expenses:[{id:1,date:"2025-09-10",currency:"COP",category:"tiquete",amount:520000,obs:"LATAM BOG-CTG-BOG",hasReceipt:true,attachments:[{name:"latam.pdf",size:"190 KB",type:"pdf"}]},{id:2,date:"2025-09-10",currency:"COP",category:"transporte",amount:38000,obs:"Uber a hotel",hasReceipt:true,attachments:[{name:"uber.pdf",size:"72 KB",type:"pdf"}]},{id:3,date:"2025-09-10",currency:"COP",category:"alojamiento",amount:450000,obs:"Hilton Cartagena 2n",hasReceipt:true,attachments:[{name:"hilton.pdf",size:"280 KB",type:"pdf"}]},{id:4,date:"2025-09-11",currency:"COP",category:"alimentacion",amount:95000,obs:"Almuerzo y cena",hasReceipt:true,attachments:[{name:"comida.jpg",size:"340 KB",type:"img"}]},{id:5,date:"2025-09-12",currency:"COP",category:"transporte",amount:42000,obs:"Uber a aeropuerto",hasReceipt:true,attachments:[{name:"uber3.pdf",size:"70 KB",type:"pdf"}]}],approvals:[{step:1,by:"Juan Medina",date:"2025-09-17",comment:""},{step:2,by:"Carolina Lopez",date:"2025-09-18",comment:"OK"},{step:3,by:"Maria Uribe",date:"2025-09-19",comment:"Verificado"},{step:4,by:"Andrea Ruiz",date:"2025-09-20",comment:"Contabilizado"},{step:5,by:"Miguel Guarin",date:"2025-09-21",comment:"Aprobado"}],createdAt:"2025-09-17"},
];

function cv(a,cur,cc){const c=cc.find(x=>x.code===cur);return c?a*c.rateToCOP:a;}
function gt(r,cc){return r.expenses.reduce((s,e)=>s+cv(e.amount,e.currency,cc),0);}
function fc(a){return"$ "+Math.round(a).toLocaleString("es-CO");}
function fa(a,cur){return["COP","CLP","ARS"].includes(cur)?Math.round(a).toLocaleString("es-CO"):a.toLocaleString("es-CO",{minimumFractionDigits:2,maximumFractionDigits:2});}
function fd(d){return d?new Date(d+"T12:00:00").toLocaleDateString("es-CO",{day:"2-digit",month:"short",year:"numeric"}):"";}
function fs(d){return d?new Date(d+"T12:00:00").toLocaleDateString("es-CO",{day:"2-digit",month:"short"}):"";}
function gid(){return"RPT-"+new Date().getFullYear()+"-"+String(Math.floor(Math.random()*900)+100);}
function ini(n){return n.split(" ").map(w=>w[0]).join("").slice(0,2).toUpperCase();}
function rl(id){return ROLES.find(r=>r.id===id)||ROLES[0];}

/* ── API CLIENT ── */
const API_BASE = "/api";
const getToken = () => sessionStorage.getItem("hr_token");
const setToken = (t) => t ? sessionStorage.setItem("hr_token", t) : sessionStorage.removeItem("hr_token");

async function api(path, opts = {}) {
  const headers = { ...(opts.headers || {}) };
  const token = getToken();
  if (token) headers["Authorization"] = "Bearer " + token;
  if (!(opts.body instanceof FormData)) headers["Content-Type"] = "application/json";
  const res = await fetch(API_BASE + path, { ...opts, headers, body: opts.body instanceof FormData ? opts.body : opts.body ? JSON.stringify(opts.body) : undefined });
  if (res.status === 401) { setToken(null); throw new Error("Sesion expirada"); }
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Error del servidor");
  return data;
}

const S={inp:{padding:"11px 14px",borderRadius:T.rs,border:"1px solid "+T.g2,fontSize:14,width:"100%",background:"white",transition:"all 0.2s",color:T.g8},lbl:{fontSize:12,fontWeight:600,color:T.g6,display:"block",marginBottom:4,letterSpacing:"0.02em"},card:{background:T.card,borderRadius:T.r,boxShadow:T.sh1,border:"1px solid "+T.g2,overflow:"hidden"},btn:(bg,c)=>({background:bg,color:c||"white",border:"none",borderRadius:T.rs,padding:"10px 18px",fontSize:13,fontWeight:600,display:"inline-flex",alignItems:"center",justifyContent:"center",gap:6,transition:"all .15s"}),gh:{background:"none",border:"none",color:T.g5,padding:6,display:"flex",alignItems:"center",justifyContent:"center",borderRadius:T.rx}};

function Badge({status}){const s=SMAP[status]||SMAP.draft;return<span style={{background:s.bg,color:s.color,padding:"4px 12px",borderRadius:20,fontSize:11,fontWeight:700,display:"inline-flex",alignItems:"center",gap:4,whiteSpace:"nowrap"}}>{status==="approved_final"&&<CheckCircle2 size={11}/>}{status==="rejected"&&<XCircle size={11}/>}{status==="submitted"&&<Clock size={11}/>}{s.label}</span>;}
function SC({icon,label,value,sub,accent}){return<div style={{...S.card,padding:"16px 18px",display:"flex",alignItems:"center",gap:14,flex:"1 1 140px",minWidth:0}}><div style={{width:44,height:44,borderRadius:12,background:accent+"12",color:accent,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{icon}</div><div style={{minWidth:0}}><div style={{fontSize:11,color:T.g4,fontWeight:500,letterSpacing:"0.04em",textTransform:"uppercase"}}>{label}</div><div style={{fontSize:20,fontWeight:800,lineHeight:1.2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",marginTop:2}}>{value}</div>{sub&&<div style={{fontSize:11,color:T.g4,marginTop:2}}>{sub}</div>}</div></div>;}
function Av({name,size:sz,color:c}){sz=sz||36;c=c||T.blue;return<div style={{width:sz,height:sz,borderRadius:"50%",background:"linear-gradient(135deg,"+c+","+c+"dd)",color:"white",display:"flex",alignItems:"center",justifyContent:"center",fontSize:sz*0.3,fontWeight:700,flexShrink:0,letterSpacing:"0.02em"}}>{ini(name)}</div>;}
function Mod({open,onClose,title,children}){if(!open)return null;return<div style={{position:"fixed",inset:0,zIndex:300,display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={onClose}><div style={{position:"fixed",inset:0,background:"rgba(10,30,61,0.45)",backdropFilter:"blur(4px)"}}/><div onClick={e=>e.stopPropagation()} style={{position:"relative",background:"white",borderRadius:"22px 22px 0 0",width:"100%",maxWidth:520,maxHeight:"88vh",overflow:"auto",padding:"8px 20px 24px",animation:"su .3s cubic-bezier(.16,1,.3,1)"}}><div style={{width:40,height:4,borderRadius:2,background:T.g3,margin:"0 auto 16px"}}/>{title&&<div style={{fontSize:18,fontWeight:700,marginBottom:16}}>{title}</div>}{children}</div></div>;}
function ATL({approvals,currentStep,status}){return<div style={{display:"flex",flexDirection:"column"}}>{AFLOW.map((step,i)=>{const ap=approvals.find(a=>a.step===step.step);const ac=step.step===currentStep&&status!=="rejected"&&status!=="approved_final";const dn=!!ap;return<div key={step.step} style={{display:"flex",gap:12}}><div style={{display:"flex",flexDirection:"column",alignItems:"center"}}><div style={{width:28,height:28,borderRadius:"50%",background:dn?T.green:ac?T.blue:T.g2,color:dn||ac?"white":T.g4,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,boxShadow:ac?"0 0 0 4px "+T.blueA:"none"}}>{dn?<Check size={13}/>:step.step}</div>{i<AFLOW.length-1&&<div style={{width:2,height:32,background:dn?T.greenL:T.g2}}/>}</div><div style={{paddingBottom:i<AFLOW.length-1?12:0,flex:1}}><div style={{fontSize:13,fontWeight:600,color:!dn&&!ac?T.g4:T.g8}}>{step.label} &mdash; {step.action}</div>{ap&&<div style={{fontSize:11,color:T.g5,marginTop:3}}>{ap.by} &bull; {fs(ap.date)}{ap.comment&&<div style={{fontStyle:"italic",marginTop:2}}>"{ap.comment}"</div>}</div>}{ac&&<div style={{fontSize:11,color:T.blue,marginTop:3,fontWeight:600}}>Pendiente</div>}</div></div>;})}</div>;}

function useTRM(){const[trm,setTrm]=useState(null);const[td,setTd]=useState(null);const[ld,setLd]=useState(true);const[er,setEr]=useState(null);const[cc,setCc]=useState(CURRENCIES);const f=useCallback(async()=>{setLd(true);setEr(null);try{const r=await fetch("https://www.datos.gov.co/resource/32sa-8pi3.json?$order=vigenciadesde%20DESC&$limit=1");if(r.ok){const d=await r.json();if(d.length>0){const v=parseFloat(d[0].valor);setTrm(v);setTd(d[0].vigenciadesde?.split("T")[0]);setCc(p=>p.map(c=>{if(c.code==="USD")return{...c,rateToCOP:v};if(c.code==="EUR")return{...c,rateToCOP:v*1.082};if(c.code==="GBP")return{...c,rateToCOP:v*1.265};if(c.code==="CAD")return{...c,rateToCOP:v*0.737};if(c.code==="CHF")return{...c,rateToCOP:v*1.136};if(c.code==="MXN")return{...c,rateToCOP:v/17.15};if(c.code==="BRL")return{...c,rateToCOP:v/4.94};if(c.code==="PEN")return{...c,rateToCOP:v/3.72};if(c.code==="ARS")return{...c,rateToCOP:v/362.75};if(c.code==="CLP")return{...c,rateToCOP:v/874.25};return c;}));}}}catch(e){setEr("Sin conexion TRM");setCc(CURRENCIES);}setLd(false);},[]);useEffect(()=>{f();},[f]);const addC=(c)=>setCc(p=>[...p,c]);const remC=(code)=>setCc(p=>p.filter(c=>c.code!==code));return{trm,trmDate:td,loading:ld,error:er,currencies:cc,refresh:f,addCurrency:addC,removeCurrency:remC};}

function TF({form,sf,currencies}){return<div style={{...S.card,padding:18,marginBottom:14}}><div style={{fontSize:14,fontWeight:700,marginBottom:14,display:"flex",alignItems:"center",gap:6}}><Plane size={16} color={T.blue}/> Informacion del Viaje</div><div style={{display:"flex",flexDirection:"column",gap:12}}><div><label style={S.lbl}>Destino *</label><input value={form.destination} onChange={e=>sf({...form,destination:e.target.value})} placeholder="Ej: Mexico, Medellin..." style={S.inp}/></div><div><label style={S.lbl}>Motivo *</label><input value={form.tripPurpose} onChange={e=>sf({...form,tripPurpose:e.target.value})} placeholder="Ej: Visita clientes" style={S.inp}/></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}><div><label style={S.lbl}>Desde</label><input type="date" value={form.dateFrom} onChange={e=>sf({...form,dateFrom:e.target.value})} style={S.inp}/></div><div><label style={S.lbl}>Hasta</label><input type="date" value={form.dateTo} onChange={e=>sf({...form,dateTo:e.target.value})} style={S.inp}/></div></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}><div><label style={S.lbl}>Moneda</label><select value={form.currency} onChange={e=>sf({...form,currency:e.target.value})} style={S.inp}>{currencies.map(c=><option key={c.code} value={c.code}>{c.code}</option>)}</select></div><div><label style={S.lbl}>Tipo</label><select value={form.type} onChange={e=>sf({...form,type:e.target.value})} style={S.inp}>{ETYPES.map(t=><option key={t.id} value={t.id}>{t.label}</option>)}</select></div></div></div></div>;}

function EF({expenses:ex,setExpenses:se,currencies:cc,sa,ssa}){
  const[n,sn]=useState({date:"",currency:"COP",category:"transporte",amount:"",obs:""});
  const[editId,setEditId]=useState(null);
  const[editVal,setEV]=useState(null);
  const add=()=>{if(!n.date||!n.amount)return;se(p=>[...p,{...n,id:Date.now(),amount:parseFloat(n.amount),hasReceipt:false,has_receipt:0,attachments:[]}]);sn({date:"",currency:n.currency,category:"transporte",amount:"",obs:""});ssa(false);};
  const startEdit=(e)=>{setEditId(e.id);setEV({date:e.date||"",currency:e.currency||"COP",category:e.category||"otros",amount:String(e.amount||""),obs:e.obs||""});};
  const saveEdit=()=>{if(!editVal||!editVal.amount)return;se(p=>p.map(e=>e.id===editId?{...e,...editVal,amount:parseFloat(editVal.amount)}:e));setEditId(null);setEV(null);};
  const cancelEdit=()=>{setEditId(null);setEV(null);};
  const tot=ex.reduce((s,e)=>s+cv(e.amount,e.currency,cc),0);
  const ei={...S.inp,fontSize:13,padding:"8px 10px"};

  return<div style={{...S.card,padding:18,marginBottom:14}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><div style={{fontSize:14,fontWeight:700,display:"flex",alignItems:"center",gap:6}}><Receipt size={16} color={T.green}/> Gastos ({ex.length})</div><button onClick={()=>ssa(true)} style={S.btn(T.greenL)}><Plus size={14}/> Agregar</button></div>
    {sa&&<div style={{background:T.g0,borderRadius:T.rs,padding:14,marginBottom:14,border:"1px solid "+T.g2}}><div style={{display:"flex",flexDirection:"column",gap:10}}><div style={{display:"grid",gridTemplateColumns:"1fr 90px",gap:8}}><input type="date" value={n.date} onChange={e=>sn({...n,date:e.target.value})} style={ei}/><select value={n.currency} onChange={e=>sn({...n,currency:e.target.value})} style={ei}>{cc.map(c=><option key={c.code} value={c.code}>{c.code}</option>)}</select></div><select value={n.category} onChange={e=>sn({...n,category:e.target.value})} style={ei}>{CATS.map(c=><option key={c.id} value={c.id}>{c.icon} {c.label}</option>)}</select><input type="number" value={n.amount} onChange={e=>sn({...n,amount:e.target.value})} placeholder="Monto" style={ei}/><input value={n.obs} onChange={e=>sn({...n,obs:e.target.value})} placeholder="Descripcion" style={ei}/><div style={{display:"flex",gap:8}}><button onClick={add} style={{...S.btn(T.green),flex:1,padding:11}}>Agregar</button><button onClick={()=>ssa(false)} style={{...S.btn(T.g1,T.g6),padding:"11px 18px"}}>Cancelar</button></div></div></div>}
    {ex.map((e,i)=>{const cat=CATS.find(c=>c.id===e.category);const cop=cv(e.amount,e.currency,cc);const isEd=editId===e.id;
      return<div key={e.id} style={{padding:"10px 0",borderBottom:i<ex.length-1?"1px solid "+T.g1:"none"}}>
        {isEd?<div style={{background:T.blueA,borderRadius:T.rs,padding:12,border:"1px solid "+T.blue+"30"}}><div style={{display:"flex",flexDirection:"column",gap:8}}>
          <input value={editVal.obs} onChange={e=>setEV({...editVal,obs:e.target.value})} placeholder="Descripcion" style={ei}/>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}><input type="number" value={editVal.amount} onChange={e=>setEV({...editVal,amount:e.target.value})} placeholder="Monto" style={ei}/><select value={editVal.currency} onChange={e=>setEV({...editVal,currency:e.target.value})} style={ei}>{cc.map(c=><option key={c.code} value={c.code}>{c.code}</option>)}</select></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}><input type="date" value={editVal.date} onChange={e=>setEV({...editVal,date:e.target.value})} style={ei}/><select value={editVal.category} onChange={e=>setEV({...editVal,category:e.target.value})} style={ei}>{CATS.map(c=><option key={c.id} value={c.id}>{c.icon} {c.label}</option>)}</select></div>
          <div style={{display:"flex",gap:6}}><button onClick={saveEdit} style={{...S.btn(T.blue),flex:1,padding:9,fontSize:12}}><Check size={13}/> Guardar</button><button onClick={cancelEdit} style={{...S.btn(T.g1,T.g6),padding:"9px 14px",fontSize:12}}>Cancelar</button></div>
        </div></div>
        :<div style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer"}} onClick={()=>startEdit(e)}>
          <div style={{width:36,height:36,borderRadius:T.rs,background:(cat?.color||T.g5)+"12",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:16}}>{cat?.icon}</div>
          <div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{e.obs||cat?.label}</div><div style={{fontSize:11,color:T.g4,marginTop:2}}>{fs(e.date)} &bull; {e.currency}</div></div>
          <div style={{textAlign:"right",flexShrink:0}}><div style={{fontSize:13,fontWeight:700,color:T.green}}>{fc(cop)}</div><div style={{fontSize:9,color:T.blue,marginTop:2}}>Tocar para editar</div></div>
          <button onClick={ev=>{ev.stopPropagation();se(p=>p.filter(x=>x.id!==e.id));}} style={{...S.gh,color:T.red}}><Trash2 size={15}/></button>
        </div>}
      </div>;})}
    {ex.length>0&&<div style={{display:"flex",justifyContent:"space-between",paddingTop:12,marginTop:8,borderTop:"2px solid "+T.g8}}><span style={{fontSize:15,fontWeight:700}}>TOTAL COP</span><span style={{fontSize:20,fontWeight:800,color:T.green}}>{fc(tot)}</span></div>}
    {ex.length===0&&!sa&&<div style={{textAlign:"center",padding:24,color:T.g4,fontSize:13}}>Toca Agregar para registrar gastos</div>}
  </div>;
}

export default function App(){
  const[loggedIn,setLoggedIn]=useState(()=>!!getToken());
  const[au,sAu]=useState([]);
  const[cu,sCu]=useState(()=>{try{const t=getToken();if(!t)return null;const p=JSON.parse(atob(t.split(".")[1]));return{id:p.id,email:p.email,role:p.role,name:p.name,area:p.area};}catch(e){return null;}});
  const[pg,sPg]=useState("dashboard");const[rpts,sRpts]=useState([]);const[sel,sSel]=useState(null);const[snr,sSnr]=useState(false);const[edt,sEdt]=useState(null);const[um,sUm]=useState(false);const[toast,sToast]=useState(null);const[loading,setLoading]=useState(true);const[dupData,sDup]=useState(null);
  const{trm,trmDate,loading:tl,error:te,currencies:cc,refresh:rf,addCurrency:addC,removeCurrency:remC}=useTRM();

  const loadData=useCallback(async()=>{try{const[users,reports]=await Promise.all([api("/users"),api("/reports")]);sAu(users);sRpts(reports);}catch(e){console.error(e);}finally{setLoading(false);}},[]);
  useEffect(()=>{if(loggedIn)loadData();},[loggedIn,loadData]);

  const doLogin=async(email,password)=>{const data=await api("/auth/login",{method:"POST",body:{email,password}});setToken(data.token);sCu(data.user);setLoggedIn(true);};
  const doLogout=()=>{setToken(null);setLoggedIn(false);sCu(null);sAu([]);sRpts([]);sPg("dashboard");sSel(null);};
  const st=msg=>{sToast(msg);setTimeout(()=>sToast(null),2800);};
  const my=useMemo(()=>{if(!cu)return[];const emp=["employee","analyst","senior_analyst","specialist","intern","assistant"];if(emp.includes(cu.role))return rpts.filter(r=>r.employee_id===cu.id||r.employeeId===cu.id);const tl2=["team_leader","supervisor","coordinator","manager"];if(tl2.includes(cu.role))return rpts.filter(r=>r.area===cu.area||r.employee_id===cu.id||r.employeeId===cu.id);return rpts;},[rpts,cu]);
  const pa=useMemo(()=>{if(!cu)return[];const tl2=["team_leader","supervisor","coordinator","manager"];if(tl2.includes(cu.role))return rpts.filter(r=>r.status==="submitted"&&r.area===cu.area);if(cu.role==="admin")return rpts.filter(r=>r.status==="approved_leader");if(cu.role==="accountant"||cu.role==="treasury")return rpts.filter(r=>r.status==="approved_admin");if(cu.role==="general_manager"||cu.role==="vp"||cu.role==="director")return rpts.filter(r=>r.status==="approved_accountant");return[];},[rpts,cu]);
  const hAppr=async(rid,c="")=>{try{const r=await api("/reports/"+rid+"/approve",{method:"POST",body:{comment:c}});sRpts(p=>p.map(x=>x.id===rid?r:x));sSel(null);st("Reporte aprobado");}catch(e){st("Error: "+e.message);}};
  const hRej=async(rid,c)=>{try{const r=await api("/reports/"+rid+"/reject",{method:"POST",body:{comment:c||"Rechazado"}});sRpts(p=>p.map(x=>x.id===rid?r:x));sSel(null);st("Reporte rechazado");}catch(e){st("Error: "+e.message);}};
  const hUpd=async(u)=>{try{const r=await api("/reports/"+u.id,{method:"PUT",body:{destination:u.destination,tripPurpose:u.trip_purpose||u.tripPurpose,dateFrom:u.date_from||u.dateFrom,dateTo:u.date_to||u.dateTo,currency:u.currency,costCenter:u.cost_center||u.costCenter,type:u.type,expenses:u.expenses}});sRpts(p=>p.map(x=>x.id===u.id?r:x));sEdt(null);sSel(r);st("Actualizado");}catch(e){st("Error: "+e.message);}};
  const nav=p=>{sPg(p);sSel(null);sSnr(false);sEdt(null);sDup(null);};
  const handleDup=(r)=>{sSel(null);sEdt(null);sDup(r);sSnr(true);sPg("reports");};
  const adm=["admin","general_manager","vp","director","hr","it_admin"];
  const ni=[{id:"dashboard",label:"Inicio",icon:<Home size={20}/>},{id:"reports",label:"Reportes",icon:<FileText size={20}/>},...(pa.length>0?[{id:"approvals",label:"Aprobar",icon:<CheckCircle2 size={20}/>,badge:pa.length}]:[]),{id:"history",label:"Historial",icon:<Clock size={20}/>},{id:"rates",label:"Tasas",icon:<Globe size={20}/>},...(cu&&adm.includes(cu.role)?[{id:"users",label:"Usuarios",icon:<Users size={20}/>}]:[])];

  /* ── LOADING after login ── */
  if(loggedIn&&!cu)return<div style={{fontFamily:T.font,background:T.bg,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}><style>{globalStyles}</style><div style={{textAlign:"center"}}><div style={{width:50,height:50,borderRadius:14,background:"linear-gradient(145deg,"+T.navy+","+T.navyL+")",display:"inline-flex",alignItems:"center",justifyContent:"center",color:"white",fontSize:16,fontWeight:800,marginBottom:16}}>HR</div><div style={{fontSize:14,color:T.g4}}>Cargando...</div></div></div>;

  const content=(()=>{
    if(pg==="dashboard")return<Dash rpts={my} cu={cu} nav={nav} sel={r=>{sSel(r);sPg("reports");}} pa={pa} cc={cc} trm={trm} td={trmDate} tl={tl} te={te} rf={rf}/>;
    if(pg==="reports"&&edt)return<Edit rpt={edt} cc={cc} onX={()=>sEdt(null)} onS={hUpd}/>;
    if(pg==="reports"&&sel)return<Det rpt={sel} cu={cu} onB={()=>sSel(null)} onA={hAppr} onR={hRej} cc={cc} ce={sel.status!=="approved_final"&&sel.status!=="rejected"} onE={()=>sEdt(sel)} onU={hUpd} onDup={handleDup}/>;
    if(pg==="reports"&&snr)return<New cu={cu} onX={()=>{sSnr(false);sDup(null);}} onS={async(r)=>{try{const nr=await api("/reports",{method:"POST",body:r});sRpts(p=>[nr,...p]);sSnr(false);sDup(null);st("Reporte creado");}catch(e){st("Error: "+e.message);}}} cc={cc} dup={dupData}/>;
    if(pg==="reports")return<List rpts={my} sel={sSel} nr={()=>sSnr(true)} cu={cu} cc={cc} onDup={handleDup}/>;
    if(pg==="approvals"&&sel)return<Det rpt={sel} cu={cu} onB={()=>sSel(null)} onA={hAppr} onR={hRej} cc={cc} ce={false} onE={()=>{}} onU={hUpd}/>;
    if(pg==="approvals")return<Appr rpts={pa} sel={sSel} cc={cc}/>;
    if(pg==="rates")return<Rates cc={cc} trm={trm} td={trmDate} tl={tl} te={te} rf={rf} addC={addC} remC={remC} st={st}/>;
    if(pg==="import")return<Imp cc={cc} onI={async(r)=>{try{const nr=await api("/reports",{method:"POST",body:r});sRpts(p=>[nr,...p]);nav("reports");st("Importado");}catch(e){st("Error: "+e.message);}}} cu={cu}/>;
    if(pg==="history"&&sel)return<Det rpt={sel} cu={cu} onB={()=>sSel(null)} onA={hAppr} onR={hRej} cc={cc} ce={false} onE={()=>{}} onU={hUpd} onDup={handleDup}/>;
    if(pg==="history")return<History rpts={my} cu={cu} sel={sSel} cc={cc} onDup={handleDup}/>;
    if(pg==="users")return<Usr users={au} setUsers={sAu} st={st}/>;
    return null;})();

  const globalStyles=`@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}input,select,textarea{font-family:inherit;font-size:16px}input:focus,select:focus,textarea:focus{outline:none;box-shadow:0 0 0 3px ${T.blueA};border-color:${T.blue}}button{cursor:pointer;font-family:inherit}button:active{transform:scale(0.97)}@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}@keyframes su{from{opacity:0;transform:translateY(40px)}to{opacity:1;transform:translateY(0)}}@keyframes ti{from{opacity:0;transform:translate(-50%,16px) scale(.96)}to{opacity:1;transform:translate(-50%,0) scale(1)}}.fi{animation:fadeIn .35s cubic-bezier(.16,1,.3,1)}`;

  /* ── LOGIN SCREEN ── */
  const[loginEmail,setLoginEmail]=useState("");const[loginPass,setLoginPass]=useState("");const[loginErr,setLoginErr]=useState("");
  if(!loggedIn)return<div style={{fontFamily:T.font,background:T.bg,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
    <style>{globalStyles}</style>
    <div className="fi" style={{width:"100%",maxWidth:400}}>
      <div style={{textAlign:"center",marginBottom:28}}>
        <div style={{width:60,height:60,borderRadius:16,background:"linear-gradient(145deg,"+T.navy+","+T.navyL+")",display:"inline-flex",alignItems:"center",justifyContent:"center",color:"white",fontSize:18,fontWeight:800,boxShadow:"0 8px 24px rgba(10,30,61,0.3)",marginBottom:16}}>HR</div>
        <h1 style={{fontSize:24,fontWeight:800,color:T.navy,letterSpacing:"-0.02em"}}>Hannover Re</h1>
        <p style={{fontSize:13,color:T.g4,marginTop:4,fontWeight:500}}>Sistema de Legalizacion de Gastos</p>
      </div>
      <div style={{...S.card,padding:24}}>
        <div style={{fontSize:14,fontWeight:700,marginBottom:20,display:"flex",alignItems:"center",gap:6}}><Lock size={16} color={T.blue}/> Iniciar Sesion</div>
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <div><label style={S.lbl}>Correo corporativo</label><input type="email" value={loginEmail} onChange={e=>{setLoginEmail(e.target.value);setLoginErr("");}} placeholder="nombre@hannover-re.com" style={S.inp}/></div>
          <div><label style={S.lbl}>Contraseña</label><input type="password" value={loginPass} onChange={e=>{setLoginPass(e.target.value);setLoginErr("");}} placeholder="Ingrese su contraseña" style={S.inp} onKeyDown={e=>{if(e.key==="Enter"&&loginEmail&&loginPass)doLogin(loginEmail,loginPass).catch(er=>setLoginErr(er.message));}}/></div>
          {loginErr&&<div style={{background:T.redA,borderRadius:T.rs,padding:"10px 14px",fontSize:12,color:T.red,fontWeight:500,display:"flex",alignItems:"center",gap:6}}><AlertCircle size={14}/>{loginErr}</div>}
          <button onClick={()=>{if(!loginEmail||!loginPass){setLoginErr("Ingrese correo y contraseña");return;}doLogin(loginEmail,loginPass).catch(er=>setLoginErr(er.message));}} style={{...S.btn(T.blue),width:"100%",padding:14,borderRadius:12,fontSize:15}}><LogIn size={16}/> Ingresar</button>
        </div>
      </div>
      <div style={{...S.card,padding:16,marginTop:12}}>
        <div style={{fontSize:11,fontWeight:700,color:T.g5,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.06em"}}>Usuarios demo disponibles</div>
        <div style={{display:"flex",flexDirection:"column",gap:4,maxHeight:180,overflowY:"auto"}}>{INIT_USERS.map(u=><button key={u.id} onClick={()=>{setLoginEmail(u.email||"");setLoginPass("demo123");setLoginErr("");}} style={{background:"none",border:"none",padding:"4px 0",fontSize:11,color:T.blue,textAlign:"left",display:"flex",alignItems:"center",gap:6}}>
          <Mail size={10}/> {u.email} <span style={{color:T.g4}}>({rl(u.role).label})</span>
        </button>)}</div>
        <div style={{fontSize:10,color:T.g4,marginTop:8,fontStyle:"italic"}}>Toca un correo para autocompletar. Contraseña: cualquiera.</div>
      </div>
      <div style={{textAlign:"center",marginTop:14,fontSize:11,color:T.g4}}>En produccion se integra con SSO / Active Directory de Hannover Re</div>
    </div>
  </div>;

  /* ── MAIN APP ── */

  return<div style={{fontFamily:T.font,background:T.bg,minHeight:"100vh",display:"flex",flexDirection:"column",color:T.g8}}>
    <style>{globalStyles}</style>
    <div style={{height:56,background:"white",borderBottom:"1px solid "+T.g2,display:"flex",alignItems:"center",padding:"0 16px",gap:10,position:"sticky",top:0,zIndex:100,flexShrink:0,boxShadow:"0 1px 3px rgba(10,30,61,0.04)"}}><div style={{display:"flex",alignItems:"center",gap:10,flex:1}}><div style={{width:34,height:34,borderRadius:9,flexShrink:0,background:"linear-gradient(145deg,"+T.navy+","+T.navyL+")",display:"flex",alignItems:"center",justifyContent:"center",color:"white",fontSize:11,fontWeight:800,letterSpacing:"0.05em",boxShadow:"0 2px 6px rgba(10,30,61,0.25)"}}>HR</div><div><div style={{fontSize:14,fontWeight:800,color:T.navy,lineHeight:1,letterSpacing:"-0.01em"}}>Hannover Re</div><div style={{fontSize:9,color:T.g4,fontWeight:600,letterSpacing:"0.12em",textTransform:"uppercase",marginTop:2}}>Gastos de Viaje</div></div></div>
      {trm&&<div style={{background:T.greenA,borderRadius:8,padding:"4px 10px",fontSize:10,fontWeight:700,color:T.green,display:"flex",alignItems:"center",gap:4,flexShrink:0,border:"1px solid rgba(5,150,105,0.15)"}}><Wifi size={10}/>TRM {Math.round(trm).toLocaleString("es-CO")}</div>}
      <div style={{position:"relative"}}><button onClick={()=>sUm(!um)} style={{background:"none",border:"none",padding:4}}><Av name={cu.name} size={34} color={rl(cu.role).color}/></button>
        {um&&<div className="fi" style={{position:"absolute",top:"100%",right:0,marginTop:8,width:280,background:"white",borderRadius:T.r,boxShadow:T.sh3,border:"1px solid "+T.g2,overflow:"hidden",zIndex:200}}>
          <div style={{padding:"14px 14px 6px"}}><div style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",marginBottom:6}}><Av name={cu.name} size={36} color={rl(cu.role).color}/><div><div style={{fontSize:14,fontWeight:700}}>{cu.name}</div><div style={{fontSize:11,color:T.g4}}>{rl(cu.role).label} &bull; {cu.area}</div></div></div>
            <div style={{borderTop:"1px solid "+T.g2,paddingTop:6}}>
              <button onClick={()=>{doLogout();sUm(false);}} style={{display:"flex",alignItems:"center",gap:8,width:"100%",padding:"10px",border:"none",borderRadius:T.rs,background:T.redA,fontSize:13,fontWeight:600,color:T.red,textAlign:"left"}}><LogIn size={15}/> Cerrar sesion</button>
            </div>
          </div>
        </div>}
      </div>
    </div>
    <div style={{flex:1,overflow:"auto",padding:"18px 14px 90px"}} onClick={()=>sUm(false)}>{content}</div>
    <div style={{position:"fixed",bottom:0,left:0,right:0,height:68,background:"rgba(255,255,255,0.92)",backdropFilter:"blur(12px) saturate(180%)",borderTop:"1px solid "+T.g2,display:"flex",alignItems:"center",justifyContent:"space-around",padding:"0 8px",zIndex:100}}>{ni.map(item=><button key={item.id} onClick={()=>nav(item.id)} style={{background:"none",border:"none",display:"flex",flexDirection:"column",alignItems:"center",gap:3,padding:"8px 12px",borderRadius:T.rs,color:pg===item.id?T.blue:T.g4,position:"relative"}}><div style={{position:"relative"}}>{item.icon}{item.badge&&<span style={{position:"absolute",top:-6,right:-8,background:T.red,color:"white",borderRadius:10,padding:"0 6px",fontSize:9,fontWeight:700,lineHeight:"16px",border:"2px solid white"}}>{item.badge}</span>}</div><span style={{fontSize:10,fontWeight:pg===item.id?700:500}}>{item.label}</span>{pg===item.id&&<div style={{position:"absolute",top:-1,width:20,height:3,borderRadius:2,background:T.blue}}/>}</button>)}</div>
    {toast&&<div style={{position:"fixed",bottom:82,left:"50%",transform:"translateX(-50%)",background:T.navy,color:"white",padding:"12px 24px",borderRadius:14,fontSize:13,fontWeight:600,zIndex:400,animation:"ti .3s cubic-bezier(.16,1,.3,1)",boxShadow:T.sh3}}>{toast}</div>}
  </div>;
}

function Dash({rpts,cu,nav,sel,pa,cc,trm,td,tl,te,rf}){
  const tot=rpts.reduce((s,r)=>s+gt(r,cc),0);const pn=rpts.filter(r=>!["approved_final","rejected","draft"].includes(r.status)).length;
  const empRoles=["employee","analyst","senior_analyst","specialist","intern","assistant"];
  const leaderRoles=["team_leader","supervisor","coordinator","manager"];
  const adminRoles=["admin","accountant","treasury","hr","auditor","compliance","legal","it_admin","director","vp","general_manager"];
  const isEmp=empRoles.includes(cu.role);const isLead=leaderRoles.includes(cu.role);const isAdm=adminRoles.includes(cu.role);

  // Stats
  const approved=rpts.filter(r=>r.status==="approved_final").length;
  const rejected=rpts.filter(r=>r.status==="rejected").length;
  const drafts=rpts.filter(r=>r.status==="draft").length;

  // By category
  const byCat={};rpts.forEach(r=>(r.expenses||[]).forEach(e=>{const cat=CATS.find(c=>c.id===e.category)||CATS[5];byCat[cat.label]=(byCat[cat.label]||0)+cv(e.amount,e.currency,cc);}));
  const catData=Object.entries(byCat).map(([name,value])=>({name,value:Math.round(value)})).sort((a,b)=>b.value-a.value);
  const catColors=[T.blue,T.purple,T.amber,T.greenL,T.pink,T.g5,"#0EA5E9","#EA580C"];

  // By status
  const statusData=[{name:"Borrador",value:drafts,color:T.g5},{name:"En proceso",value:pn,color:T.blue},{name:"Aprobados",value:approved,color:T.green},{name:"Rechazados",value:rejected,color:T.red}].filter(d=>d.value>0);

  // By area (admin only)
  const byArea={};rpts.forEach(r=>{const a=r.area||"Otro";byArea[a]=(byArea[a]||0)+gt(r,cc);});
  const areaData=Object.entries(byArea).map(([name,value])=>({name,value:Math.round(value)})).sort((a,b)=>b.value-a.value);

  const ChartCard=({title,children})=><div style={{...S.card,padding:16,marginBottom:14}}><div style={{fontSize:14,fontWeight:700,marginBottom:12}}>{title}</div>{children}</div>;
  const fmtTick=(v)=>v>=1000000?"$"+Math.round(v/1000000)+"M":v>=1000?"$"+Math.round(v/1000)+"K":"$"+v;

  return<div className="fi" style={{maxWidth:600,margin:"0 auto"}}><div style={{marginBottom:20}}><h1 style={{fontSize:22,fontWeight:800,letterSpacing:"-0.02em"}}>Hola, {cu.name.split(" ")[0]}</h1><p style={{color:T.g4,fontSize:13,marginTop:3,fontWeight:500}}>{isEmp?"Tus gastos de viaje":isLead?"Resumen de tu equipo":"Panel ejecutivo de gastos"}</p></div>

    {/* TRM Card */}
    <div style={{background:"linear-gradient(145deg,"+T.navy+","+T.navyL+")",borderRadius:T.r,padding:"20px 22px",marginBottom:16,color:"white",position:"relative",overflow:"hidden",boxShadow:"0 8px 24px rgba(10,30,61,0.3)"}}><div style={{position:"absolute",top:-30,right:-30,width:120,height:120,borderRadius:"50%",background:"rgba(255,255,255,0.04)"}}/><div style={{position:"relative",display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:11,opacity:0.6,fontWeight:600,letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:6}}>TRM &mdash; Banco de la Republica</div><div style={{fontSize:32,fontWeight:800,lineHeight:1,letterSpacing:"-0.02em"}}>{tl?"...":trm?"$"+trm.toLocaleString("es-CO",{maximumFractionDigits:2}):"$4.250"}</div><div style={{fontSize:11,opacity:0.5,marginTop:6}}>COP/USD {td&&(" - "+fs(td))}</div></div><button onClick={rf} style={{background:"rgba(255,255,255,0.1)",border:"none",borderRadius:T.rs,width:40,height:40,display:"flex",alignItems:"center",justifyContent:"center",color:"white"}}><RefreshCw size={16}/></button></div>{te&&<div style={{fontSize:10,color:"#FCD34D",marginTop:8}}>{te}</div>}</div>

    {/* Stat Cards */}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:14}}>
      <SC icon={<DollarSign size={22}/>} label="Total COP" value={fc(tot)} accent={T.green}/>
      <SC icon={<FileText size={22}/>} label="Reportes" value={rpts.length} sub={pn+" en proceso"} accent={T.blue}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:16}}>
      <div style={{...S.card,padding:"12px 14px",textAlign:"center"}}><div style={{fontSize:22,fontWeight:800,color:T.green}}>{approved}</div><div style={{fontSize:10,color:T.g4,marginTop:2}}>Aprobados</div></div>
      <div style={{...S.card,padding:"12px 14px",textAlign:"center"}}><div style={{fontSize:22,fontWeight:800,color:T.blue}}>{pn}</div><div style={{fontSize:10,color:T.g4,marginTop:2}}>En proceso</div></div>
      <div style={{...S.card,padding:"12px 14px",textAlign:"center"}}><div style={{fontSize:22,fontWeight:800,color:T.red}}>{rejected}</div><div style={{fontSize:10,color:T.g4,marginTop:2}}>Rechazados</div></div>
    </div>

    {/* Charts: By Category (all roles) */}
    {catData.length>0&&<ChartCard title="Gastos por Categoria">
      <div style={{display:"flex",alignItems:"center",gap:12}}>
        <div style={{width:140,height:140,flexShrink:0}}><ResponsiveContainer width="100%" height={140}><PieChart><Pie data={catData} dataKey="value" cx="50%" cy="50%" outerRadius={60} innerRadius={30}>{catData.map((d,i)=><Cell key={i} fill={catColors[i%catColors.length]}/>)}</Pie></PieChart></ResponsiveContainer></div>
        <div style={{flex:1,minWidth:0}}>{catData.map((d,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:6,marginBottom:6}}><div style={{width:10,height:10,borderRadius:3,background:catColors[i%catColors.length],flexShrink:0}}/><div style={{flex:1,fontSize:11,color:T.g6,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{d.name}</div><div style={{fontSize:11,fontWeight:700,color:T.g8,flexShrink:0}}>{fc(d.value)}</div></div>)}</div>
      </div>
    </ChartCard>}

    {/* Charts: By Status (leaders + admin) */}
    {(isLead||isAdm)&&statusData.length>0&&<ChartCard title="Pipeline de Reportes">
      <ResponsiveContainer width="100%" height={160}><BarChart data={statusData} layout="vertical" margin={{left:10,right:10,top:5,bottom:5}}><XAxis type="number" hide/><YAxis type="category" dataKey="name" width={80} tick={{fontSize:11,fill:T.g5}}/><Tooltip formatter={v=>v+" reportes"}/><Bar dataKey="value" radius={[0,6,6,0]}>{statusData.map((d,i)=><Cell key={i} fill={d.color}/>)}</Bar></BarChart></ResponsiveContainer>
    </ChartCard>}

    {/* Charts: By Area (admin only) */}
    {isAdm&&areaData.length>0&&<ChartCard title="Gastos por Area">
      <ResponsiveContainer width="100%" height={Math.max(120,areaData.length*40)}><BarChart data={areaData} layout="vertical" margin={{left:10,right:10,top:5,bottom:5}}><XAxis type="number" tickFormatter={fmtTick} tick={{fontSize:10,fill:T.g4}}/><YAxis type="category" dataKey="name" width={90} tick={{fontSize:11,fill:T.g5}}/><Tooltip formatter={v=>fc(v)}/><Bar dataKey="value" fill={T.blue} radius={[0,6,6,0]}/></BarChart></ResponsiveContainer>
    </ChartCard>}

    {/* Import + Approvals */}
    <button onClick={()=>nav("import")} style={{width:"100%",...S.card,padding:"14px 18px",marginBottom:12,display:"flex",alignItems:"center",gap:12,textAlign:"left",cursor:"pointer"}}><div style={{width:42,height:42,borderRadius:12,background:T.blueA,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><FileSpreadsheet size={20} color={T.blue}/></div><div style={{flex:1}}><div style={{fontSize:14,fontWeight:700}}>Importar desde archivo</div><div style={{fontSize:12,color:T.g4,marginTop:2}}>Excel, CSV, TSV</div></div><ChevronRight size={18} color={T.g3}/></button>
    {pa.length>0&&<button onClick={()=>nav("approvals")} style={{width:"100%",background:"linear-gradient(135deg,#FEF3C7,#FDE68A)",border:"1px solid #FCD34D",borderRadius:T.r,padding:"14px 18px",marginBottom:16,display:"flex",alignItems:"center",gap:12,textAlign:"left",boxShadow:"0 4px 12px rgba(217,119,6,0.1)"}}><AlertTriangle size={20} color={T.amber}/><div style={{flex:1}}><div style={{fontWeight:700,fontSize:14,color:"#92400E"}}>{pa.length} aprobacion(es) pendiente(s)</div></div><ArrowRight size={18} color={T.amber}/></button>}

    {/* Recent */}
    <h2 style={{fontSize:16,fontWeight:700,marginBottom:12}}>Recientes</h2>
    {rpts.slice(0,5).map(r=>{const t=gt(r,cc);return<button key={r.id} onClick={()=>sel(r)} style={{...S.card,padding:"14px 16px",display:"flex",alignItems:"center",gap:14,textAlign:"left",width:"100%",marginBottom:10,cursor:"pointer"}}><div style={{width:42,height:42,borderRadius:12,flexShrink:0,background:SMAP[r.status]?.bg,color:SMAP[r.status]?.color,display:"flex",alignItems:"center",justifyContent:"center"}}><Plane size={18}/></div><div style={{flex:1,minWidth:0}}><div style={{fontSize:14,fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.destination}</div><div style={{fontSize:11,color:T.g4,marginTop:3,fontWeight:500}}>{r.code||r.id} - {fs(r.date_from||r.dateFrom)}</div></div><div style={{textAlign:"right",flexShrink:0}}><div style={{fontSize:14,fontWeight:800}}>{fc(t)}</div><div style={{marginTop:4}}><Badge status={r.status}/></div></div></button>;})}
  </div>;
}

function List({rpts,sel,nr,cu,cc,onDup}){const[f,sf]=useState("all");const fl=f==="all"?rpts:rpts.filter(r=>r.status===f);const emp=["employee","analyst","senior_analyst","specialist","intern","assistant"];
  return<div className="fi" style={{maxWidth:600,margin:"0 auto"}}><div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}><h1 style={{fontSize:22,fontWeight:800,letterSpacing:"-0.02em"}}>Reportes</h1>{emp.includes(cu.role)&&<button onClick={nr} style={S.btn(T.blue)}><Plus size={15}/> Nuevo</button>}</div>
    <div style={{display:"flex",gap:6,marginBottom:14,overflowX:"auto",paddingBottom:4}}>{[{id:"all",label:"Todos"},...Object.entries(SMAP).map(([id,v])=>({id,label:v.label}))].map(x=><button key={x.id} onClick={()=>sf(x.id)} style={{padding:"7px 14px",borderRadius:20,border:"1.5px solid",borderColor:f===x.id?T.blue:T.g2,background:f===x.id?T.blueA:"white",color:f===x.id?T.blue:T.g5,fontSize:11,fontWeight:600,whiteSpace:"nowrap",flexShrink:0}}>{x.label}</button>)}</div>
    {fl.map(r=>{const t=gt(r,cc);return<div key={r.id} style={{...S.card,padding:"14px 16px",display:"flex",alignItems:"center",gap:14,marginBottom:10,cursor:"pointer"}} onClick={()=>sel(r)}><div style={{flex:1,minWidth:0}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}><span style={{fontSize:12,fontWeight:800,color:T.g6,fontFamily:"monospace"}}>{r.code||r.id}</span><Badge status={r.status}/></div><div style={{fontSize:14,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.destination} &mdash; {r.trip_purpose||r.tripPurpose}</div><div style={{fontSize:11,color:T.g4,marginTop:4,fontWeight:500}}>{fs(r.date_from||r.dateFrom)} - {fs(r.date_to||r.dateTo)} | {(r.expenses||[]).length} gastos</div></div><div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6,flexShrink:0}}><div style={{fontSize:16,fontWeight:800}}>{fc(t)}</div>{emp.includes(cu.role)&&<button onClick={e=>{e.stopPropagation();onDup(r);}} style={{...S.gh,color:T.blue,background:T.blueA,borderRadius:T.rx,padding:"4px 8px",fontSize:10,fontWeight:600,display:"flex",alignItems:"center",gap:3}}><Copy size={11}/> Duplicar</button>}</div></div>;})}
    {fl.length===0&&<div style={{textAlign:"center",padding:48,color:T.g4}}><FileText size={36} style={{margin:"0 auto 10px",opacity:0.4}}/><div style={{fontSize:14}}>Sin reportes</div></div>}
  </div>;
}

function Det({rpt,cu,onB,onA,onR,cc,ce,onE,onU,onDup}){const[cm,sCm]=useState("");const[tab,sTab]=useState("expenses");const[am,sAm]=useState(null);const[va,sVa]=useState(null);const attRef=useRef();
  const tot=gt(rpt,cc);const ar={team_leader:["submitted"],supervisor:["submitted"],coordinator:["submitted"],manager:["submitted"],admin:["approved_leader"],accountant:["approved_admin"],treasury:["approved_admin"],general_manager:["approved_accountant"],vp:["approved_accountant"],director:["approved_accountant"]};
  const ca=ar[cu.role]?.includes(rpt.status)&&(["admin","accountant","treasury","general_manager","vp","director"].includes(cu.role)||rpt.area===cu.area);
  const aa=(eid,file)=>{if(!file)return;const fd=new FormData();fd.append("file",file);api("/expenses/"+eid+"/attach",{method:"POST",body:fd}).then(atts=>{onU({...rpt,expenses:rpt.expenses.map(e=>e.id===eid?{...e,has_receipt:1,hasReceipt:true,attachments:atts}:e)});sAm(null);}).catch(e=>console.error(e));};
  return<div className="fi" style={{maxWidth:600,margin:"0 auto"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><button onClick={onB} style={{...S.gh,color:T.g5,gap:4,fontSize:13,fontWeight:600}}><ChevronLeft size={18}/> Volver</button><div style={{display:"flex",gap:6}}>{onDup&&<button onClick={()=>onDup(rpt)} style={S.btn(T.g1,T.blue)}><Copy size={13}/> Duplicar</button>}{ce&&<button onClick={onE} style={S.btn(T.blueA,T.blue)}><Edit3 size={14}/> Editar</button>}</div></div>
    <div style={{...S.card,padding:18,marginBottom:14}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}><span style={{fontSize:13,fontWeight:800,fontFamily:"monospace",color:T.g6}}>{rpt.id}</span><Badge status={rpt.status}/></div><div style={{fontSize:17,fontWeight:700,marginBottom:12}}>{rpt.destination} &mdash; {rpt.tripPurpose}</div>
      <div style={{background:"linear-gradient(145deg,"+T.navy+","+T.navyL+")",borderRadius:12,padding:"14px 18px",color:"white",marginBottom:14,position:"relative",overflow:"hidden",boxShadow:"0 4px 12px rgba(10,30,61,0.2)"}}><div style={{position:"absolute",top:-20,right:-20,width:80,height:80,borderRadius:"50%",background:"rgba(255,255,255,0.04)"}}/><div style={{fontSize:11,opacity:0.6,fontWeight:600,letterSpacing:"0.06em",textTransform:"uppercase"}}>Total Pesos Colombianos</div><div style={{fontSize:28,fontWeight:800,marginTop:4,letterSpacing:"-0.02em"}}>{fc(tot)}</div><div style={{fontSize:11,opacity:0.5,marginTop:4}}>Aprox USD {(tot/(cc.find(c=>c.code==="USD")?.rateToCOP||4250)).toFixed(2)}</div></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,fontSize:12}}>{[{l:"Empleado",v:rpt.employeeName},{l:"CC",v:rpt.cc},{l:"Area",v:rpt.area},{l:"Periodo",v:fs(rpt.dateFrom)+" - "+fs(rpt.dateTo)},{l:"Tipo",v:rpt.type==="tarjeta"?"TC Corp.":"Anticipo"},{l:"C.Costo",v:rpt.costCenter||"--"}].map((f,i)=><div key={i} style={{padding:"8px 0",borderBottom:"1px solid "+T.g1}}><div style={{fontSize:10,color:T.g4,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em"}}>{f.l}</div><div style={{fontWeight:700,marginTop:3}}>{f.v}</div></div>)}</div>
    </div>
    <div style={{display:"flex",gap:6,marginBottom:14}}>{[{id:"expenses",label:"Gastos",icon:<Receipt size={14}/>},{id:"approvals",label:"Aprobacion",icon:<CheckCircle2 size={14}/>},{id:"attachments",label:"Soportes",icon:<Paperclip size={14}/>}].map(t=><button key={t.id} onClick={()=>sTab(t.id)} style={{padding:"9px 14px",borderRadius:T.rs,border:"none",flex:1,background:tab===t.id?T.navy:"white",color:tab===t.id?"white":T.g5,fontSize:12,fontWeight:600,display:"flex",alignItems:"center",justifyContent:"center",gap:5,boxShadow:tab!==t.id?T.sh1:"none"}}>{t.icon} {t.label}</button>)}</div>
    {tab==="expenses"&&<div style={{...S.card}}>{rpt.expenses.map((e,i)=>{const cat=CATS.find(c=>c.id===e.category);const cop=cv(e.amount,e.currency,cc);return<div key={e.id} style={{padding:"12px 16px",borderBottom:i<rpt.expenses.length-1?"1px solid "+T.g1:"none"}}><div style={{display:"flex",alignItems:"center",gap:10}}><div style={{width:38,height:38,borderRadius:T.rs,background:(cat?.color||T.g5)+"10",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:17}}>{cat?.icon}</div><div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{e.obs}</div><div style={{fontSize:11,color:T.g4,marginTop:2}}>{fs(e.date)} &bull; {e.currency} {fa(e.amount,e.currency)}</div></div><div style={{textAlign:"right",flexShrink:0}}><div style={{fontSize:14,fontWeight:800,color:T.green}}>{fc(cop)}</div><div style={{display:"flex",alignItems:"center",gap:4,marginTop:3,justifyContent:"flex-end"}}>{(e.attachments||[]).length>0?<span style={{fontSize:10,color:T.green,display:"flex",alignItems:"center",gap:3,fontWeight:600}}><Paperclip size={10}/>{e.attachments.length}</span>:<span style={{fontSize:10,color:T.amber,fontWeight:500}}>Sin soporte</span>}</div></div></div></div>;})}
      <div style={{padding:"14px 16px",background:T.greenA,display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontSize:15,fontWeight:800}}>TOTAL COP</span><span style={{fontSize:20,fontWeight:800,color:T.green}}>{fc(tot)}</span></div></div>}
    {tab==="approvals"&&<div style={{...S.card,padding:18}}><ATL approvals={rpt.approvals} currentStep={rpt.currentStep} status={rpt.status}/></div>}
    {tab==="attachments"&&<div style={{...S.card,padding:18}}>{rpt.expenses.map((e,i)=>{const cat=CATS.find(c=>c.id===e.category);const at=e.attachments||[];return<div key={e.id} style={{marginBottom:16,paddingBottom:16,borderBottom:i<rpt.expenses.length-1?"1px solid "+T.g1:"none"}}><div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}><span style={{fontSize:17}}>{cat?.icon}</span><div style={{flex:1,fontSize:13,fontWeight:700}}>{e.obs}</div><button onClick={()=>sAm(e.id)} style={S.btn(T.blueA,T.blue)}><Plus size={12}/> Adjuntar</button></div>
      {at.length>0?at.map((a,j)=><div key={j} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0",marginLeft:28}}><div style={{width:36,height:36,borderRadius:T.rs,background:a.type==="img"?T.amberA:T.blueA,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{a.type==="img"?<Image size={15} color={T.amber}/>:<FileText size={15} color={T.blue}/>}</div><div style={{flex:1,minWidth:0}}><div style={{fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.name}</div><div style={{fontSize:10,color:T.g4,marginTop:1}}>{a.size}</div></div><button onClick={()=>sVa({e,a})} style={S.btn(T.g1,T.blue)}><Eye size={12}/></button></div>):<div style={{marginLeft:28,fontSize:12,color:T.amber,display:"flex",alignItems:"center",gap:5,fontWeight:500}}><AlertCircle size={13}/> Sin soportes</div>}
    </div>;})}
    </div>}
    <Mod open={!!am} onClose={()=>sAm(null)} title="Adjuntar Soporte">
      <input ref={attRef} type="file" accept=".pdf,.jpg,.jpeg,.png,.webp" onChange={e=>{if(e.target.files[0])aa(am,e.target.files[0]);}} style={{display:"none"}}/>
      <button onClick={()=>attRef.current?.click()} style={{width:"100%",border:"2px dashed "+T.g3,borderRadius:T.r,padding:28,textAlign:"center",marginBottom:16,background:T.g0,cursor:"pointer"}}><Upload size={32} color={T.g4} style={{margin:"0 auto 10px"}}/><div style={{fontSize:14,fontWeight:600,color:T.g6}}>Toca para seleccionar archivo</div><div style={{fontSize:12,color:T.g4,marginTop:4}}>PDF, JPG, PNG &mdash; Max 10MB</div></button>
    </Mod>
    <Mod open={!!va} onClose={()=>sVa(null)} title="Vista Previa">{va&&<div>
      <div style={{background:T.g0,borderRadius:T.r,padding:16,textAlign:"center",marginBottom:16,overflow:"hidden"}}>
        {va.a.mime_type?.startsWith("image/")?<img src={"/api/attachments/"+va.a.id+"/preview"} alt={va.a.original_name||va.a.name} style={{maxWidth:"100%",maxHeight:300,borderRadius:T.rs,objectFit:"contain"}}/>:va.a.mime_type==="application/pdf"?<iframe src={"/api/attachments/"+va.a.id+"/preview"} title={va.a.original_name||va.a.name} style={{width:"100%",height:300,border:"none",borderRadius:T.rs}}/>:<div style={{padding:20}}>{va.a.type==="img"?<Image size={52} color={T.amber}/>:<FileText size={52} color={T.blue}/>}</div>}
        <div style={{fontSize:14,fontWeight:700,marginTop:10}}>{va.a.original_name||va.a.name}</div>
        <div style={{fontSize:12,color:T.g5,marginTop:4}}>{va.a.size}</div>
      </div>
      <a href={"/api/attachments/"+va.a.id+"/download"} style={{...S.btn(T.navy),width:"100%",padding:14,borderRadius:12,textDecoration:"none",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}><Download size={15}/> Descargar</a>
    </div>}</Mod>
    {ca&&<div style={{...S.card,padding:18,marginTop:14,border:"2px solid "+T.blue}}><div style={{fontSize:14,fontWeight:700,marginBottom:12,display:"flex",alignItems:"center",gap:6}}><CheckCircle2 size={16} color={T.blue}/> Aprobacion</div><textarea value={cm} onChange={e=>sCm(e.target.value)} placeholder="Comentario (opcional)..." style={{...S.inp,resize:"vertical",minHeight:56,marginBottom:12}}/><div style={{display:"flex",gap:10}}><button onClick={()=>onA(rpt.id,cm)} style={{...S.btn(T.green),flex:1,padding:14,borderRadius:12}}><Check size={17}/> Aprobar</button><button onClick={()=>onR(rpt.id,cm)} style={{...S.btn(T.red),flex:1,padding:14,borderRadius:12}}><X size={17}/> Rechazar</button></div></div>}
  </div>;
}

function Edit({rpt,cc,onX,onS}){const[f,sf]=useState({destination:rpt.destination||"",tripPurpose:rpt.trip_purpose||rpt.tripPurpose||"",dateFrom:rpt.date_from||rpt.dateFrom||"",dateTo:rpt.date_to||rpt.dateTo||"",currency:rpt.currency||"COP",costCenter:rpt.cost_center||rpt.costCenter||"",type:rpt.type||"anticipo"});const[ex,se]=useState([...(rpt.expenses||[])]);const[sa,ssa]=useState(false);
  return<div className="fi" style={{maxWidth:600,margin:"0 auto"}}><button onClick={onX} style={{...S.gh,color:T.g5,gap:4,fontSize:13,fontWeight:600,marginBottom:14}}><ChevronLeft size={18}/> Cancelar</button><h1 style={{fontSize:22,fontWeight:800,marginBottom:4,letterSpacing:"-0.02em"}}>Editar Reporte</h1><p style={{fontSize:13,color:T.g4,marginBottom:16}}>{rpt.code||rpt.id} &mdash; {rpt.destination} &mdash; Se puede editar hasta la aprobacion final</p>
    <TF form={f} sf={sf} currencies={cc}/><EF expenses={ex} setExpenses={se} currencies={cc} sa={sa} ssa={ssa}/>
    <div style={{display:"flex",gap:10,marginBottom:24}}><button onClick={onX} style={{...S.btn(T.g1,T.g6),flex:1,padding:14,borderRadius:12}}>Cancelar</button><button onClick={()=>onS({...rpt,...f,expenses:ex})} style={{...S.btn(T.blue),flex:1,padding:14,borderRadius:12}}><Save size={15}/> Guardar</button></div>
  </div>;
}

function New({cu,onX,onS,cc,dup}){const[f,sf]=useState(dup?{destination:dup.destination||"",tripPurpose:dup.trip_purpose||dup.tripPurpose||"",dateFrom:"",dateTo:"",currency:dup.currency||"COP",costCenter:dup.cost_center||dup.costCenter||"",type:dup.type||"anticipo"}:{destination:"",tripPurpose:"",dateFrom:"",dateTo:"",currency:"COP",costCenter:"",type:"anticipo"});const[ex,se]=useState(dup?(dup.expenses||[]).map((e,i)=>({...e,id:Date.now()+i,attachments:[],hasReceipt:false,has_receipt:0})):[]);const[sa,ssa]=useState(false);
  const sub=d=>{onS({id:gid(),employeeId:cu.id,employeeName:cu.name,area:cu.area,cc:cu.cc,...f,presentationDate:new Date().toISOString().split("T")[0],status:d?"draft":"submitted",currentStep:d?1:2,expenses:ex,approvals:d?[]:[{step:1,by:cu.name,date:new Date().toISOString().split("T")[0],comment:""}],createdAt:new Date().toISOString().split("T")[0]});};
  return<div className="fi" style={{maxWidth:600,margin:"0 auto"}}><button onClick={onX} style={{...S.gh,color:T.g5,gap:4,fontSize:13,fontWeight:600,marginBottom:14}}><ChevronLeft size={18}/> Cancelar</button><h1 style={{fontSize:22,fontWeight:800,marginBottom:4,letterSpacing:"-0.02em"}}>{dup?"Duplicar Reporte":"Nuevo Reporte"}</h1>{dup&&<p style={{fontSize:13,color:T.g4,marginBottom:12}}>Basado en: {dup.destination} &mdash; Modifique lo que necesite</p>}<div style={{height:dup?0:12}}/>
    <TF form={f} sf={sf} currencies={cc}/><EF expenses={ex} setExpenses={se} currencies={cc} sa={sa} ssa={ssa}/>
    <div style={{display:"flex",gap:10,marginBottom:24}}><button onClick={()=>sub(true)} disabled={ex.length===0} style={{...S.btn(T.g1,T.g6),flex:1,padding:14,borderRadius:12,opacity:ex.length===0?0.5:1}}>Borrador</button><button onClick={()=>sub(false)} disabled={ex.length===0||!f.destination} style={{...S.btn(T.blue),flex:1,padding:14,borderRadius:12,opacity:(ex.length===0||!f.destination)?0.5:1}}><Send size={15}/> Enviar</button></div>
  </div>;
}

function Appr({rpts,sel,cc}){return<div className="fi" style={{maxWidth:600,margin:"0 auto"}}><h1 style={{fontSize:22,fontWeight:800,marginBottom:4,letterSpacing:"-0.02em"}}>Aprobaciones</h1><p style={{fontSize:13,color:T.g4,marginBottom:16,fontWeight:500}}>{rpts.length} pendiente(s)</p>
  {rpts.map(r=>{const t=gt(r,cc);return<div key={r.id} style={{...S.card,padding:16,marginBottom:10}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}><span style={{fontSize:13,fontWeight:800,fontFamily:"monospace",color:T.g6}}>{r.code||r.id}</span><Badge status={r.status}/></div><div style={{fontSize:14,fontWeight:600}}>{r.employee_name||r.employeeName} &mdash; {r.destination}</div><div style={{fontSize:12,color:T.g4,marginTop:3}}>{r.trip_purpose||r.tripPurpose}</div><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:12}}><div style={{fontSize:18,fontWeight:800}}>{fc(t)}</div><button onClick={()=>sel(r)} style={S.btn(T.blue)}>Revisar</button></div></div>;})}</div>;}

/* ════════════ HISTORY ════════════ */
function History({rpts,cu,sel,cc,onDup}){
  const[filterY,setFY]=useState("all");const[filterM,setFM]=useState("all");const[filterD,setFD]=useState("");
  const leaderRoles=["team_leader","supervisor","coordinator","manager","admin","accountant","treasury","director","vp","general_manager"];
  const isManager=leaderRoles.includes(cu.role);
  const completed=rpts.filter(r=>r.status==="approved_final"||r.status==="rejected");
  const months=["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

  // Get unique years
  const years=[...new Set(completed.map(r=>{const d=r.created_at||r.date_from||r.dateFrom||"";return d.slice(0,4);}).filter(y=>y))].sort().reverse();

  // Apply filters
  let filtered=completed;
  if(filterY!=="all")filtered=filtered.filter(r=>(r.created_at||r.date_from||"").startsWith(filterY));
  if(filterM!=="all")filtered=filtered.filter(r=>{const d=r.created_at||r.date_from||"";const m=d.slice(5,7);return m===filterM;});
  if(filterD)filtered=filtered.filter(r=>(r.created_at||r.date_from||"")===filterD);

  // Group by month for managers
  const grouped={};
  if(isManager){filtered.forEach(r=>{const d=r.created_at||r.date_from||"";const key=d.slice(0,7)||"Sin fecha";if(!grouped[key])grouped[key]=[];grouped[key].push(r);});}

  const sortedKeys=Object.keys(grouped).sort().reverse();

  return<div className="fi" style={{maxWidth:600,margin:"0 auto"}}>
    <h1 style={{fontSize:22,fontWeight:800,letterSpacing:"-0.02em",marginBottom:4}}>Historial</h1>
    <p style={{fontSize:13,color:T.g4,marginBottom:14,fontWeight:500}}>{completed.length} reportes procesados</p>

    {/* Filters */}
    <div style={{...S.card,padding:14,marginBottom:14}}>
      <div style={{fontSize:13,fontWeight:700,marginBottom:10,display:"flex",alignItems:"center",gap:6}}><Search size={14} color={T.blue}/> Filtros</div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        <div style={{flex:"1 1 80px"}}><label style={{...S.lbl,fontSize:10}}>Ano</label><select value={filterY} onChange={e=>setFY(e.target.value)} style={{...S.inp,fontSize:12,padding:"8px 10px"}}><option value="all">Todos</option>{years.map(y=><option key={y} value={y}>{y}</option>)}</select></div>
        <div style={{flex:"1 1 80px"}}><label style={{...S.lbl,fontSize:10}}>Mes</label><select value={filterM} onChange={e=>setFM(e.target.value)} style={{...S.inp,fontSize:12,padding:"8px 10px"}}><option value="all">Todos</option>{months.map((m,i)=><option key={i} value={String(i+1).padStart(2,"0")}>{m}</option>)}</select></div>
        <div style={{flex:"1 1 120px"}}><label style={{...S.lbl,fontSize:10}}>Fecha exacta</label><input type="date" value={filterD} onChange={e=>setFD(e.target.value)} style={{...S.inp,fontSize:12,padding:"8px 10px"}}/></div>
      </div>
      {(filterY!=="all"||filterM!=="all"||filterD)&&<button onClick={()=>{setFY("all");setFM("all");setFD("");}} style={{...S.btn(T.g1,T.g5),marginTop:8,padding:"6px 12px",fontSize:11}}>Limpiar filtros</button>}
    </div>

    {/* Stats */}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
      <div style={{...S.card,padding:"12px 14px",textAlign:"center"}}><div style={{fontSize:22,fontWeight:800,color:T.green}}>{filtered.filter(r=>r.status==="approved_final").length}</div><div style={{fontSize:10,color:T.g4,marginTop:2}}>Aprobados</div></div>
      <div style={{...S.card,padding:"12px 14px",textAlign:"center"}}><div style={{fontSize:22,fontWeight:800,color:T.red}}>{filtered.filter(r=>r.status==="rejected").length}</div><div style={{fontSize:10,color:T.g4,marginTop:2}}>Rechazados</div></div>
    </div>

    {/* Manager grouped view */}
    {isManager&&sortedKeys.length>0?sortedKeys.map(key=>{const items=grouped[key];const[y,m]=key.split("-");const label=m?months[parseInt(m)-1]+" "+y:key;const total=items.reduce((s,r)=>s+gt(r,cc),0);
      return<div key={key} style={{marginBottom:16}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8,padding:"8px 12px",background:T.blueA,borderRadius:T.rs}}>
          <div style={{fontSize:14,fontWeight:700,color:T.blue}}>{label}</div>
          <div style={{fontSize:12,fontWeight:700,color:T.g6}}>{items.length} reportes &bull; {fc(total)}</div>
        </div>
        {items.map(r=><HistoryCard key={r.id} r={r} cc={cc} sel={sel} onDup={onDup}/>)}
      </div>;
    }):<div>{filtered.map(r=><HistoryCard key={r.id} r={r} cc={cc} sel={sel} onDup={onDup}/>)}</div>}

    {filtered.length===0&&<div style={{textAlign:"center",padding:40,color:T.g4}}><Clock size={36} style={{margin:"0 auto 10px",opacity:0.4}}/><div style={{fontSize:14}}>Sin registros para estos filtros</div></div>}
  </div>;
}

function HistoryCard({r,cc,sel,onDup}){
  const t=gt(r,cc);
  return<div style={{...S.card,padding:"14px 16px",marginBottom:8,display:"flex",alignItems:"center",gap:12}}>
    <div style={{flex:1,minWidth:0,cursor:"pointer"}} onClick={()=>sel(r)}>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:3}}><span style={{fontSize:11,fontWeight:800,fontFamily:"monospace",color:T.g4}}>{r.code||r.id}</span><Badge status={r.status}/></div>
      <div style={{fontSize:13,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.employee_name||r.employeeName} &mdash; {r.destination}</div>
      <div style={{fontSize:11,color:T.g4,marginTop:2}}>{fs(r.created_at||r.date_from)} &bull; {(r.expenses||[]).length} gastos &bull; {fc(t)}</div>
    </div>
    <button onClick={()=>onDup(r)} style={{...S.gh,color:T.blue,background:T.blueA,borderRadius:T.rs,padding:"6px 10px",fontSize:10,fontWeight:600,display:"flex",alignItems:"center",gap:4,flexShrink:0}}><Copy size={12}/> Duplicar</button>
  </div>;
}

function Imp({cc,onI,cu}){const[pv,sPv]=useState(null);const fr=useRef();
  const parseRows=(rows)=>{const ex=[];for(let i=1;i<rows.length;i++){const c=rows[i];if(!c||c.length<2)continue;const vals=c.map(x=>String(x||"").trim());const d=vals.find(x=>/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(x))||vals.find(x=>/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(x))||"";const a=vals.find(x=>/^[\d.,]+$/.test(x)&&parseFloat(x.replace(/,/g,""))>0);const am=a?parseFloat(a.replace(/,/g,"")):0;const ds=vals.find(x=>x.length>3&&!/^[\d.,/\-]+$/.test(x))||"Gasto "+i;if(am>0)ex.push({id:Date.now()+i,date:d||new Date().toISOString().split("T")[0],currency:"COP",category:"otros",amount:am,obs:ds,hasReceipt:false,attachments:[]});}return ex;};
  const hf=async e=>{const file=e.target.files[0];if(!file)return;const ext=file.name.split(".").pop().toLowerCase();try{if(ext==="xlsx"||ext==="xls"){const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array"});const ws=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws,{header:1});const ex=parseRows(rows);sPv(ex.length>0?{fn:file.name,ex,sz:(file.size/1024).toFixed(0)+" KB"}:{fn:file.name,ex:[],er:"No se encontraron datos en el Excel."});}else{const text=await file.text();const sep=text.includes("\t")?"\t":",";const rows=text.split("\n").filter(l=>l.trim()).map(l=>l.split(sep).map(c=>c.trim().replace(/^"|"$/g,"")));const ex=parseRows(rows);sPv(ex.length>0?{fn:file.name,ex,sz:(file.size/1024).toFixed(0)+" KB"}:{fn:file.name,ex:[],er:"No se encontraron datos validos."});}}catch(err){sPv({fn:file.name,ex:[],er:"Error al leer: "+err.message});}};
  const di=()=>{if(!pv?.ex?.length)return;onI({id:gid(),employeeId:cu.id,employeeName:cu.name,area:cu.area,cc:cu.cc,destination:"(Importado)",tripPurpose:"Desde "+pv.fn,dateFrom:pv.ex[0]?.date||"",dateTo:pv.ex[pv.ex.length-1]?.date||"",presentationDate:new Date().toISOString().split("T")[0],currency:"COP",costCenter:"",status:"draft",currentStep:1,type:"anticipo",expenses:pv.ex,approvals:[],createdAt:new Date().toISOString().split("T")[0]});};
  return<div className="fi" style={{maxWidth:600,margin:"0 auto"}}><h1 style={{fontSize:22,fontWeight:800,marginBottom:4,letterSpacing:"-0.02em"}}>Importar Gastos</h1><p style={{fontSize:13,color:T.g4,marginBottom:16}}>Desde Excel, CSV o TSV</p>
    <input ref={fr} type="file" accept=".csv,.tsv,.txt,.xlsx,.xls" onChange={hf} style={{display:"none"}}/>
    <button onClick={()=>fr.current?.click()} style={{width:"100%",background:"white",border:"2px dashed "+T.blue,borderRadius:T.r,padding:32,textAlign:"center",marginBottom:18,boxShadow:T.sh1}}><FileSpreadsheet size={40} color={T.blue} style={{margin:"0 auto 10px"}}/><div style={{fontSize:15,fontWeight:700,color:T.blue}}>Seleccionar archivo</div><div style={{fontSize:12,color:T.g4,marginTop:4}}>CSV, TSV, Excel</div></button>
    <div style={{...S.card,padding:16,marginBottom:16}}><div style={{fontSize:14,fontWeight:700,marginBottom:8}}>Formato esperado</div><div style={{fontSize:12,color:T.g5,lineHeight:1.7}}>Columnas: fecha, monto, descripcion<br/><code style={{background:T.g1,padding:"3px 8px",borderRadius:T.rx,fontSize:11}}>2025-05-25, 45200, Uber aeropuerto</code></div></div>
    {pv&&<div style={{...S.card,padding:18,marginBottom:16}}><div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}><FileSpreadsheet size={22} color={T.green}/><div><div style={{fontSize:14,fontWeight:700}}>{pv.fn}</div><div style={{fontSize:11,color:T.g4}}>{pv.sz}</div></div></div>
      {pv.er?<div style={{background:T.redA,borderRadius:T.rs,padding:14,fontSize:13,color:T.red,fontWeight:500}}>{pv.er}</div>:<div><div style={{fontSize:13,fontWeight:700,color:T.green,marginBottom:10,display:"flex",alignItems:"center",gap:5}}><CheckCircle2 size={15}/> {pv.ex.length} gastos</div>
        {pv.ex.slice(0,5).map((e,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0",borderBottom:"1px solid "+T.g1,fontSize:13}}><div style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:T.g6}}>{e.obs}</div><div style={{fontWeight:700,color:T.green}}>{fc(e.amount)}</div></div>)}
        {pv.ex.length>5&&<div style={{fontSize:12,color:T.g4,paddingTop:8}}>... y {pv.ex.length-5} mas</div>}
        <button onClick={di} style={{...S.btn(T.green),width:"100%",marginTop:14,padding:14,borderRadius:12}}><Download size={15}/> Importar como Borrador</button>
      </div>}</div>}
  </div>;
}

function Usr({users,setUsers,st}){const[sr,sSr]=useState("");const[eu,sEu]=useState(null);const[sa,sSa]=useState(false);const[nu,sNu]=useState({name:"",role:"employee",area:"Administracion",cc:"",email:"",phone:""});const[delU,sDelU]=useState(null);
  const fl=sr?users.filter(u=>u.name.toLowerCase().includes(sr.toLowerCase())||u.email?.toLowerCase().includes(sr.toLowerCase())||rl(u.role).label.toLowerCase().includes(sr.toLowerCase())):users;
  const au=()=>{if(!nu.name)return;api("/users",{method:"POST",body:nu}).then(u=>{setUsers(p=>[...p,u]);sNu({name:"",role:"employee",area:"Administracion",cc:"",email:"",phone:""});sSa(false);st("Usuario creado");}).catch(e=>st("Error: "+e.message));};
  const sv=()=>{if(!eu)return;api("/users/"+eu.id,{method:"PUT",body:eu}).then(u=>{setUsers(p=>p.map(x=>x.id===eu.id?u:x));sEu(null);st("Usuario actualizado");}).catch(e=>st("Error: "+e.message));};
  const del=()=>{if(!delU)return;api("/users/"+delU.id,{method:"DELETE"}).then(()=>{setUsers(p=>p.filter(u=>u.id!==delU.id));sDelU(null);sEu(null);st("Usuario eliminado");}).catch(e=>st("Error: "+e.message));};
  return<div className="fi" style={{maxWidth:600,margin:"0 auto"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><div><h1 style={{fontSize:22,fontWeight:800,letterSpacing:"-0.02em"}}>Usuarios</h1><p style={{fontSize:13,color:T.g4,marginTop:3,fontWeight:500}}>{users.length} registrados</p></div><button onClick={()=>sSa(true)} style={S.btn(T.blue)}><UserPlus size={15}/> Nuevo</button></div>
    <div style={{position:"relative",marginTop:14,marginBottom:16}}><Search size={16} color={T.g4} style={{position:"absolute",left:14,top:"50%",transform:"translateY(-50%)"}}/><input value={sr} onChange={e=>sSr(e.target.value)} placeholder="Buscar por nombre, email o rol..." style={{...S.inp,paddingLeft:40}}/></div>
    <div style={{display:"flex",gap:6,overflowX:"auto",paddingBottom:8,marginBottom:14}}>{ROLES.filter(r=>users.some(u=>u.role===r.id)).map(r=>{const c=users.filter(u=>u.role===r.id).length;return<div key={r.id} style={{background:r.color+"10",border:"1px solid "+r.color+"25",borderRadius:20,padding:"4px 12px",fontSize:10,fontWeight:700,color:r.color,whiteSpace:"nowrap",flexShrink:0}}>{r.label} ({c})</div>;})}</div>
    {fl.map(u=>{const r=rl(u.role);return<div key={u.id} style={{...S.card,padding:16,marginBottom:10}}><div style={{display:"flex",alignItems:"center",gap:12}}><Av name={u.name} size={46} color={r.color}/><div style={{flex:1,minWidth:0}}><div style={{fontSize:15,fontWeight:700}}>{u.name}</div><div style={{fontSize:12,color:T.g5,marginTop:2,display:"flex",alignItems:"center",gap:6}}>{u.email&&<span style={{display:"flex",alignItems:"center",gap:3}}><Mail size={10}/> {u.email}</span>}</div><div style={{display:"flex",alignItems:"center",gap:6,marginTop:6,flexWrap:"wrap"}}><span style={{background:r.color+"12",color:r.color,border:"1px solid "+r.color+"20",borderRadius:6,padding:"3px 10px",fontSize:11,fontWeight:700,display:"flex",alignItems:"center",gap:4}}><Shield size={10}/>{r.label}</span><span style={{background:T.g1,color:T.g5,borderRadius:6,padding:"3px 10px",fontSize:11,fontWeight:600,display:"flex",alignItems:"center",gap:4}}><Building2 size={10}/>{u.area}</span></div></div><button onClick={()=>sEu({...u})} style={{...S.gh,color:T.blue,background:T.blueA,borderRadius:T.rs,width:36,height:36}}><Pencil size={15}/></button><button onClick={()=>sDelU(u)} style={{...S.gh,color:T.red,background:T.redA,borderRadius:T.rs,width:36,height:36}}><Trash2 size={15}/></button></div></div>;})}
    {fl.length===0&&<div style={{textAlign:"center",padding:40,color:T.g4}}><Users size={36} style={{margin:"0 auto 10px",opacity:0.4}}/><div style={{fontSize:14}}>{sr?"Sin resultados":"Sin usuarios"}</div></div>}
    <Mod open={sa} onClose={()=>sSa(false)} title="Nuevo Usuario"><div style={{display:"flex",flexDirection:"column",gap:14}}><div><label style={S.lbl}>Nombre completo *</label><input value={nu.name} onChange={e=>sNu({...nu,name:e.target.value})} placeholder="Ej: Maria Rodriguez" style={S.inp}/></div><div><label style={S.lbl}>Email corporativo</label><input value={nu.email} onChange={e=>sNu({...nu,email:e.target.value})} placeholder="nombre@hannover-re.com" style={S.inp}/></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}><div><label style={S.lbl}>CC</label><input value={nu.cc} onChange={e=>sNu({...nu,cc:e.target.value})} placeholder="Numero" style={S.inp}/></div><div><label style={S.lbl}>Telefono</label><input value={nu.phone} onChange={e=>sNu({...nu,phone:e.target.value})} placeholder="+57 3XX..." style={S.inp}/></div></div><div><label style={S.lbl}>Rol *</label><select value={nu.role} onChange={e=>sNu({...nu,role:e.target.value})} style={S.inp}>{ROLES.map(r=><option key={r.id} value={r.id}>{r.label} - {r.desc}</option>)}</select><div style={{fontSize:11,color:T.g4,marginTop:4}}>{rl(nu.role).desc}</div></div><div><label style={S.lbl}>Area *</label><select value={nu.area} onChange={e=>sNu({...nu,area:e.target.value})} style={S.inp}>{AREAS.map(a=><option key={a} value={a}>{a}</option>)}</select></div><div style={{display:"flex",gap:10,marginTop:4}}><button onClick={()=>sSa(false)} style={{...S.btn(T.g1,T.g6),flex:1,padding:14,borderRadius:12}}>Cancelar</button><button onClick={au} disabled={!nu.name} style={{...S.btn(T.green),flex:1,padding:14,borderRadius:12,opacity:nu.name?1:0.5}}><UserPlus size={15}/> Crear</button></div></div></Mod>
    <Mod open={!!eu} onClose={()=>sEu(null)} title="Editar Usuario">{eu&&<div style={{display:"flex",flexDirection:"column",gap:14}}>
      <div style={{display:"flex",alignItems:"center",gap:12,padding:14,background:T.g0,borderRadius:T.rs}}><Av name={eu.name} size={48} color={rl(eu.role).color}/><div><div style={{fontSize:16,fontWeight:700}}>{eu.name}</div><div style={{fontSize:12,color:T.g4,marginTop:2}}>{rl(eu.role).label} &bull; {eu.area}</div></div></div>
      <div><label style={S.lbl}>Nombre completo</label><input value={eu.name} onChange={e=>sEu({...eu,name:e.target.value})} style={S.inp}/></div>
      <div><label style={S.lbl}>Email</label><input value={eu.email||""} onChange={e=>sEu({...eu,email:e.target.value})} style={S.inp}/></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}><div><label style={S.lbl}>CC</label><input value={eu.cc||""} onChange={e=>sEu({...eu,cc:e.target.value})} style={S.inp}/></div><div><label style={S.lbl}>Telefono</label><input value={eu.phone||""} onChange={e=>sEu({...eu,phone:e.target.value})} style={S.inp}/></div></div>
      <div><label style={S.lbl}>Rol</label><select value={eu.role} onChange={e=>sEu({...eu,role:e.target.value})} style={S.inp}>{ROLES.map(r=><option key={r.id} value={r.id}>{r.label} - {r.desc}</option>)}</select><div style={{fontSize:11,color:rl(eu.role).color,marginTop:4,fontWeight:600}}>{rl(eu.role).desc}</div></div>
      <div><label style={S.lbl}>Area</label><select value={eu.area} onChange={e=>sEu({...eu,area:e.target.value})} style={S.inp}>{AREAS.map(a=><option key={a} value={a}>{a}</option>)}</select></div>
      <div style={{display:"flex",gap:10,marginTop:4}}><button onClick={()=>sEu(null)} style={{...S.btn(T.g1,T.g6),flex:1,padding:14,borderRadius:12}}>Cancelar</button><button onClick={sv} style={{...S.btn(T.blue),flex:1,padding:14,borderRadius:12}}><Save size={15}/> Guardar</button></div>
      <button onClick={()=>{sDelU(eu);}} style={{width:"100%",marginTop:10,background:"none",border:"1px solid "+T.red,color:T.red,borderRadius:12,padding:"12px",fontSize:13,fontWeight:600,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}><Trash2 size={14}/> Eliminar usuario</button>
    </div>}</Mod>
    <Mod open={!!delU} onClose={()=>sDelU(null)} title="Eliminar Usuario">{delU&&<div>
      <div style={{display:"flex",alignItems:"center",gap:12,padding:16,background:T.redA,borderRadius:T.rs,marginBottom:16}}>
        <Av name={delU.name} size={44} color={T.red}/>
        <div><div style={{fontSize:15,fontWeight:700}}>{delU.name}</div><div style={{fontSize:12,color:T.g5,marginTop:2}}>{rl(delU.role).label} &bull; {delU.area}</div></div>
      </div>
      <div style={{fontSize:14,color:T.g6,marginBottom:16,lineHeight:1.6}}>Esta accion eliminara al usuario permanentemente del sistema. No se podra deshacer.</div>
      <div style={{display:"flex",gap:10}}><button onClick={()=>sDelU(null)} style={{...S.btn(T.g1,T.g6),flex:1,padding:14,borderRadius:12}}>Cancelar</button><button onClick={del} style={{...S.btn(T.red),flex:1,padding:14,borderRadius:12}}><Trash2 size={15}/> Confirmar</button></div>
    </div>}</Mod>
  </div>;
}

function Rates({cc,trm,td,tl,te,rf,addC,remC,st}){
  const[fromC,setFromC]=useState("USD");const[toC,setToC]=useState("COP");const[amt,setAmt]=useState("100");
  const[showAdd,setShowAdd]=useState(false);const[nc,setNc]=useState({code:"",name:"",rateToCOP:"",symbol:""});
  const swap=()=>{setFromC(toC);setToC(fromC);};
  const convert=(a,from,to)=>{const f=cc.find(c=>c.code===from);const t=cc.find(c=>c.code===to);if(!f||!t||!a)return 0;const inCOP=a*f.rateToCOP;return inCOP/t.rateToCOP;};
  const result=convert(parseFloat(amt)||0,fromC,toC);
  const rate1=convert(1,fromC,toC);
  const fmtR=(v)=>v>=1?v.toLocaleString("es-CO",{minimumFractionDigits:2,maximumFractionDigits:2}):v.toLocaleString("es-CO",{minimumFractionDigits:4,maximumFractionDigits:6});
  const handleAdd=()=>{if(!nc.code||!nc.name||!nc.rateToCOP)return;if(cc.find(c=>c.code===nc.code.toUpperCase())){st("Ya existe esa moneda");return;}addC({code:nc.code.toUpperCase(),name:nc.name,rateToCOP:parseFloat(nc.rateToCOP),symbol:nc.symbol||nc.code.toUpperCase()});setNc({code:"",name:"",rateToCOP:"",symbol:""});setShowAdd(false);st("Moneda agregada");};
  const builtIn=["USD","EUR","COP","MXN","BRL","ARS","CLP","PEN","GBP","CAD","CHF"];

  return<div className="fi" style={{maxWidth:600,margin:"0 auto"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}><h1 style={{fontSize:22,fontWeight:800,letterSpacing:"-0.02em"}}>Tasas de Cambio</h1><button onClick={rf} style={S.btn(T.blueA,T.blue)}><RefreshCw size={13}/> Actualizar</button></div>

    {/* TRM Card */}
    <div style={{background:"linear-gradient(145deg,"+T.navy+","+T.navyL+")",borderRadius:T.r,padding:"20px 22px",marginBottom:16,color:"white",position:"relative",overflow:"hidden",boxShadow:"0 8px 24px rgba(10,30,61,0.3)"}}><div style={{position:"absolute",top:-30,right:-30,width:120,height:120,borderRadius:"50%",background:"rgba(255,255,255,0.04)"}}/><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>{trm?<Wifi size={15}/>:<WifiOff size={15}/>}<span style={{fontSize:12,fontWeight:700}}>TRM &mdash; Superintendencia Financiera</span></div><div style={{fontSize:36,fontWeight:800,letterSpacing:"-0.02em"}}>{tl?"Cargando...":trm?"$"+trm.toLocaleString("es-CO",{maximumFractionDigits:2}):"N/A"}</div><div style={{fontSize:11,opacity:0.5,marginTop:6}}>COP/USD {td&&(" - "+fd(td))}</div><div style={{fontSize:10,opacity:0.4,marginTop:4}}>Fuente: datos.gov.co</div></div>
    {te&&<div style={{background:T.amberA,border:"1px solid #FCD34D",borderRadius:T.rs,padding:"10px 14px",marginBottom:14,fontSize:12,color:"#92400E",display:"flex",alignItems:"center",gap:8}}><AlertTriangle size={15}/>{te}</div>}

    {/* Converter */}
    <div style={{...S.card,padding:18,marginBottom:16}}>
      <div style={{fontSize:14,fontWeight:700,marginBottom:14,display:"flex",alignItems:"center",gap:6}}><ArrowLeftRight size={16} color={T.blue}/> Convertir Monedas</div>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
        <div style={{flex:1}}><label style={{...S.lbl,fontSize:10}}>De</label><select value={fromC} onChange={e=>setFromC(e.target.value)} style={{...S.inp,fontSize:13}}>{cc.map(c=><option key={c.code} value={c.code}>{c.code} - {c.name}</option>)}</select></div>
        <button onClick={swap} style={{...S.gh,color:T.blue,background:T.blueA,borderRadius:"50%",width:36,height:36,marginTop:16,flexShrink:0}}><ArrowLeftRight size={16}/></button>
        <div style={{flex:1}}><label style={{...S.lbl,fontSize:10}}>A</label><select value={toC} onChange={e=>setToC(e.target.value)} style={{...S.inp,fontSize:13}}>{cc.map(c=><option key={c.code} value={c.code}>{c.code} - {c.name}</option>)}</select></div>
      </div>
      <div style={{marginBottom:12}}><label style={{...S.lbl,fontSize:10}}>Monto en {fromC}</label><input type="number" value={amt} onChange={e=>setAmt(e.target.value)} style={{...S.inp,fontSize:18,fontWeight:700,textAlign:"center"}}/></div>
      <div style={{background:"linear-gradient(135deg,"+T.greenA+",rgba(5,150,105,0.04))",borderRadius:T.rs,padding:"16px 18px",textAlign:"center",border:"1px solid rgba(5,150,105,0.15)"}}>
        <div style={{fontSize:11,color:T.g4,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em"}}>Resultado en {toC}</div>
        <div style={{fontSize:28,fontWeight:800,color:T.green,marginTop:4,letterSpacing:"-0.02em"}}>{fmtR(result)}</div>
        <div style={{fontSize:11,color:T.g5,marginTop:6}}>1 {fromC} = {fmtR(rate1)} {toC}</div>
        <div style={{fontSize:11,color:T.g4,marginTop:2}}>1 {toC} = {fmtR(convert(1,toC,fromC))} {fromC}</div>
      </div>
    </div>

    {/* Currency Table */}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
      <div style={{fontSize:16,fontWeight:700}}>Monedas ({cc.length})</div>
      <button onClick={()=>setShowAdd(true)} style={S.btn(T.blueA,T.blue)}><Plus size={14}/> Agregar</button>
    </div>

    {showAdd&&<div style={{...S.card,padding:16,marginBottom:12,border:"2px solid "+T.blue}}>
      <div style={{fontSize:14,fontWeight:700,marginBottom:12}}>Nueva Moneda</div>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        <div style={{display:"grid",gridTemplateColumns:"80px 1fr",gap:8}}>
          <div><label style={{...S.lbl,fontSize:10}}>Codigo *</label><input value={nc.code} onChange={e=>setNc({...nc,code:e.target.value.toUpperCase().slice(0,3)})} placeholder="JPY" maxLength={3} style={{...S.inp,fontSize:13,textTransform:"uppercase",textAlign:"center",fontWeight:700}}/></div>
          <div><label style={{...S.lbl,fontSize:10}}>Nombre *</label><input value={nc.name} onChange={e=>setNc({...nc,name:e.target.value})} placeholder="Ej: Yen Japones" style={{...S.inp,fontSize:13}}/></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 80px",gap:8}}>
          <div><label style={{...S.lbl,fontSize:10}}>Tasa a COP * (1 unidad = ? COP)</label><input type="number" value={nc.rateToCOP} onChange={e=>setNc({...nc,rateToCOP:e.target.value})} placeholder="Ej: 28.5" style={{...S.inp,fontSize:13}}/></div>
          <div><label style={{...S.lbl,fontSize:10}}>Simbolo</label><input value={nc.symbol} onChange={e=>setNc({...nc,symbol:e.target.value})} placeholder="$" style={{...S.inp,fontSize:13,textAlign:"center"}}/></div>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>setShowAdd(false)} style={{...S.btn(T.g1,T.g6),flex:1,padding:11}}>Cancelar</button>
          <button onClick={handleAdd} disabled={!nc.code||!nc.name||!nc.rateToCOP} style={{...S.btn(T.green),flex:1,padding:11,opacity:(!nc.code||!nc.name||!nc.rateToCOP)?0.5:1}}>Agregar</button>
        </div>
      </div>
    </div>}

    <div style={{...S.card}}>{cc.map((c,i)=><div key={c.code} style={{display:"flex",alignItems:"center",padding:"12px 16px",borderBottom:i<cc.length-1?"1px solid "+T.g1:"none"}}>
      <span style={{background:T.g1,padding:"4px 10px",borderRadius:T.rx,fontWeight:800,fontSize:11,marginRight:12,width:42,textAlign:"center",fontFamily:"monospace",color:T.g6}}>{c.code}</span>
      <div style={{flex:1}}><div style={{fontSize:13,fontWeight:600}}>{c.name}</div></div>
      <div style={{fontSize:14,fontWeight:800,fontFamily:"monospace",marginRight:8}}>{c.code==="COP"?"1.00":fc(c.rateToCOP)}</div>
      {!builtIn.includes(c.code)&&<button onClick={()=>{remC(c.code);st("Moneda eliminada");}} style={{...S.gh,color:T.red,width:28,height:28}}><Trash2 size={13}/></button>}
    </div>)}</div>
  </div>;
}
