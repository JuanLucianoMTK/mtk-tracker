import { useState, useMemo, useRef, useEffect } from "react";
import { ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";
import { supabase } from "./supabase";

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const ESTATUSES  = ["activo","vacante","incapacidad","vacaciones"];
const CATEGORIAS = ["A","B","C"];
const MESES_FULL = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const MESES      = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
const ANIOS      = [2024,2025,2026,2027,2028];
const FONT       = "-apple-system,BlinkMacSystemFont,'SF Pro Display','Segoe UI',Helvetica,sans-serif";

const STATUS_CFG = {
  activo:      { label:"Activo",      bg:"#DCFCE7", text:"#14532D", dot:"#22C55E" },
  vacante:     { label:"Vacante",     bg:"#F1F5F9", text:"#475569", dot:"#94A3B8" },
  incapacidad: { label:"Incapacidad", bg:"#FFFBEB", text:"#78350F", dot:"#F59E0B" },
  vacaciones:  { label:"Vacaciones",  bg:"#EFF6FF", text:"#1E40AF", dot:"#3B82F6" },
};

const ROL_CFG = {
  admin:       { label:"Administrador",  bg:"#EDE9FE", text:"#5B21B6" },
  encargado:   { label:"Encargado",      bg:"#DBEAFE", text:"#1E40AF" },
  jefe_taller: { label:"Jefe de taller", bg:"#DCFCE7", text:"#14532D" },
};

const pk = (id,m,a) => `${id}-${m}-${a}`;
function tsNow() {
  const n = new Date();
  const fecha = n.toLocaleDateString("es-MX",{day:"2-digit",month:"2-digit",year:"numeric"});
  const hora  = n.toLocaleTimeString("es-MX",{hour:"2-digit",minute:"2-digit"});
  return `${fecha}, ${hora}`;
}
function fechaHora(ts) { const p=ts.split(","); return { fecha:p[0]?.trim()||ts, hora:p[1]?.trim()||"" }; }

function genPosicionesVacias(tallerId, count, mes, anio) {
  return Array.from({ length:count }, (_,i) => ({
    id: `${tallerId}-${mes}-${anio}-${i+1}`,
    taller_id: tallerId, mes, anio,
    numero: i+1,
    categoria: CATEGORIAS[i%3],
    num_socio: "",
    nombre_tecnico: "",
    estatus: "vacante",
    comentario: "",
  }));
}

// ─── SUPABASE HELPERS ────────────────────────────────────────────────────────
async function dbLoadTalleres() {
  const { data, error } = await supabase.from("talleres").select("*").order("id");
  if (error) { console.error(error); return []; }
  return data;
}

async function dbLoadUsuarios() {
  const { data, error } = await supabase.from("usuarios").select("*").order("id");
  if (error) { console.error(error); return []; }
  return data;
}

async function dbLoadNotifConfig() {
  const { data, error } = await supabase.from("notificaciones_config").select("*").eq("id",1).single();
  if (error) return { activo:false, correo_destino:"" };
  return data;
}

async function dbSaveNotifConfig(cfg) {
  await supabase.from("notificaciones_config").upsert({ id:1, ...cfg });
}

async function dbLoadPosiciones(tallerId, mes, anio) {
  const { data, error } = await supabase.from("posiciones")
    .select("*")
    .eq("taller_id", tallerId)
    .eq("mes", mes)
    .eq("anio", anio)
    .order("numero");
  if (error) { console.error(error); return null; }
  return data;
}

async function dbLoadPosicionesPrevMonth(tallerId, mes, anio) {
  const prevM = mes===0?11:mes-1;
  const prevA = mes===0?anio-1:anio;
  const { data } = await supabase.from("posiciones")
    .select("*")
    .eq("taller_id", tallerId)
    .eq("mes", prevM)
    .eq("anio", prevA)
    .order("numero");
  return data||[];
}

async function dbUpsertPosiciones(rows) {
  const { error } = await supabase.from("posiciones").upsert(rows, { onConflict:"taller_id,mes,anio,numero" });
  if (error) console.error("upsert posiciones:", error);
}

async function dbUpdatePosicion(pos) {
  const { error } = await supabase.from("posiciones")
    .update({ num_socio:pos.num_socio, nombre_tecnico:pos.nombre_tecnico, estatus:pos.estatus, comentario:pos.comentario, updated_at:new Date().toISOString(), updated_by:pos.updated_by })
    .eq("taller_id", pos.taller_id)
    .eq("mes", pos.mes)
    .eq("anio", pos.anio)
    .eq("numero", pos.numero);
  if (error) console.error("update posicion:", error);
}

async function dbSaveTaller(taller) {
  const { error } = await supabase.from("talleres").update({
    posiciones_autorizadas: taller.posiciones_autorizadas,
    jefe_nombre: taller.jefe_nombre,
    jefe_email: taller.jefe_email,
  }).eq("id", taller.id);
  if (error) console.error("update taller:", error);
}

async function dbAddUsuario(u) {
  const { data, error } = await supabase.from("usuarios").insert({
    nombre: u.nombre, email: u.email, password: u.password,
    rol: u.rol, talleres_ids: u.talleres_ids, activo: u.activo,
  }).select().single();
  if (error) { console.error(error); return null; }
  return data;
}

async function dbToggleUsuario(id, activo) {
  await supabase.from("usuarios").update({ activo }).eq("id", id);
}

async function dbDeleteUsuario(id) {
  await supabase.from("usuarios").delete().eq("id", id);
}

async function dbAddBitacora(entry) {
  const { fecha, hora } = fechaHora(entry.ts);
  await supabase.from("bitacora_tecnicos").insert({ ...entry, fecha, hora });
}

async function dbAddBitacoraAcceso(entry) {
  const { fecha, hora } = fechaHora(entry.ts);
  await supabase.from("bitacora_accesos").insert({ ...entry, fecha, hora });
}

async function dbLoadBitacora() {
  const { data } = await supabase.from("bitacora_tecnicos").select("*").order("id", { ascending:false }).limit(200);
  return data||[];
}

async function dbLoadBitacoraAccesos() {
  const { data } = await supabase.from("bitacora_accesos").select("*").order("id", { ascending:false }).limit(200);
  return data||[];
}

function getMockChartData(anio, total) {
  const seeds = [0.88,0.91,0.87,0.93,0.89,0.92,0.90,0.85,0.94,0.91,0.88,0.86];
  return MESES.map((mes,i)=>({ mes, autorizados:total, activos:Math.round(total*(seeds[(i+anio)%12])) }));
}

// ─── EXCEL EXPORTS ────────────────────────────────────────────────────────────
function exportBitacora(bitacora) {
  const rows = bitacora.map((r,i)=>({ "ID":i+1,"Fecha":r.fecha||r.ts,"Hora":r.hora||"","Fecha/Hora":r.ts,"Usuario":r.usuario,"Taller":r.taller||"—","Tipo":r.tipo||"—","Posición":r.posicion||"—","Campo":r.campo||"—","Valor anterior":r.anterior||"—","Valor nuevo":r.nuevo||"—","Comentario":r.comentario||"—" }));
  const ws=XLSX.utils.json_to_sheet(rows); const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Bitácora");
  XLSX.writeFile(wb,`MTK_Bitacora_${Date.now()}.xlsx`);
}

function exportAccesos(bitacora) {
  const rows = bitacora.map((r,i)=>({ "ID":i+1,"Fecha":r.fecha||r.ts,"Hora":r.hora||"","Usuario":r.usuario,"Acción":r.accion||"—","Afectado":r.afectado||"—","Detalle":r.detalle||"—" }));
  const ws=XLSX.utils.json_to_sheet(rows); const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Accesos");
  XLSX.writeFile(wb,`MTK_Accesos_${Date.now()}.xlsx`);
}

function exportDashboardExcel(talleres, posData, mes, anio) {
  const rows = talleres.map(t=>{ const pos=posData[pk(t.id,mes,anio)]||[]; const activos=pos.filter(p=>p.estatus==="activo").length; const vacantes=pos.filter(p=>p.estatus==="vacante").length; const pct=t.posiciones_autorizadas>0?Math.round(vacantes/t.posiciones_autorizadas*100):0; return { "Taller":t.nombre,"Periodo":`${MESES_FULL[mes]} ${anio}`,"Pos. aut.":t.posiciones_autorizadas,"Activos":activos||"—","Vacantes":vacantes||"—","% Vacantes":pos.length>0?`${pct}%`:"—","Incapacidad":pos.filter(p=>p.estatus==="incapacidad").length||"—","Vacaciones":pos.filter(p=>p.estatus==="vacaciones").length||"—","Jefe":t.jefe_nombre||"—" }; });
  const ws=XLSX.utils.json_to_sheet(rows); const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Dashboard");
  XLSX.writeFile(wb,`MTK_Dashboard_${MESES_FULL[mes]}_${anio}.xlsx`);
}

// ─── UI HELPERS ───────────────────────────────────────────────────────────────
const iBase = { background:"#F2F2F7",border:"1.5px solid transparent",borderRadius:10,padding:"10px 14px",fontSize:14,color:"#1C1C1E",outline:"none",fontFamily:FONT,boxSizing:"border-box",width:"100%",transition:"border-color 0.15s" };

function Chip({ estatus }) {
  const c=STATUS_CFG[estatus]||STATUS_CFG.vacante;
  return <span style={{ display:"inline-flex",alignItems:"center",gap:5,background:c.bg,color:c.text,padding:"3px 10px",borderRadius:999,fontSize:11,fontWeight:600 }}><span style={{ width:6,height:6,borderRadius:"50%",background:c.dot,flexShrink:0 }}/>{c.label}</span>;
}
function RolBadge({ rol }) {
  const c=ROL_CFG[rol]||{label:rol,bg:"#F1F5F9",text:"#374151"};
  return <span style={{ fontSize:11,fontWeight:700,padding:"3px 9px",borderRadius:999,background:c.bg,color:c.text }}>{c.label}</span>;
}
function PeriodSelector({ mes,setMes,anio,setAnio }) {
  const s={border:"1.5px solid #E2E8F0",borderRadius:9,padding:"7px 12px",fontSize:13,outline:"none",background:"#fff",fontFamily:FONT,fontWeight:600,color:"#374151",cursor:"pointer"};
  return <div style={{ display:"flex",gap:8,alignItems:"center" }}><select value={mes} onChange={e=>setMes(Number(e.target.value))} style={s}>{MESES_FULL.map((m,i)=><option key={i} value={i}>{m}</option>)}</select><select value={anio} onChange={e=>setAnio(Number(e.target.value))} style={s}>{ANIOS.map(a=><option key={a} value={a}>{a}</option>)}</select></div>;
}
function Toast({ msg }) {
  if (!msg) return null;
  return <div style={{ position:"fixed",bottom:28,left:"50%",transform:"translateX(-50%)",background:"#0B2267",color:"#fff",padding:"12px 22px",borderRadius:14,fontSize:13,fontWeight:600,fontFamily:FONT,boxShadow:"0 8px 32px rgba(11,34,103,0.35)",zIndex:9999,display:"flex",alignItems:"center",gap:10,whiteSpace:"nowrap" }}>✓ {msg}</div>;
}
function Loader({ text="Cargando..." }) {
  return <div style={{ minHeight:"60vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:12,fontFamily:FONT }}>
    <div style={{ width:36,height:36,borderRadius:"50%",border:"3px solid #E2E8F0",borderTop:"3px solid #2563EB",animation:"spin 0.8s linear infinite" }}/>
    <p style={{ margin:0,fontSize:13,color:"#94A3B8" }}>{text}</p>
    <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
  </div>;
}
function ExportBtn({ onClick,label="Exportar Excel" }) {
  return <button onClick={onClick} style={{ display:"flex",alignItems:"center",gap:6,background:"#DCFCE7",color:"#14532D",border:"none",borderRadius:9,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:FONT }}>↓ {label}</button>;
}

// ─── LOGIN ────────────────────────────────────────────────────────────────────
function LoginView({ onLogin }) {
  const [email,setEmail]=useState(""); const [pass,setPass]=useState(""); const [err,setErr]=useState(""); const [loading,setLoading]=useState(false);
  const attempt = async () => {
    setLoading(true); setErr("");
    const { data, error } = await supabase.from("usuarios").select("*").eq("email",email).eq("password",pass).eq("activo",true).single();
    setLoading(false);
    if (error||!data) { setErr("Correo o contraseña incorrectos."); return; }
    onLogin(data);
  };
  return (
    <div style={{ minHeight:"100vh",background:"linear-gradient(155deg,#071A52 0%,#0B2267 40%,#1641A3 75%,#2563EB 100%)",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:FONT }}>
      <div style={{ width:380 }}>
        <div style={{ textAlign:"center",marginBottom:28 }}>
          <div style={{ width:60,height:60,background:"rgba(255,255,255,0.12)",backdropFilter:"blur(20px)",borderRadius:18,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 14px",border:"1px solid rgba(255,255,255,0.2)" }}><span style={{ fontSize:28 }}>⚙️</span></div>
          <p style={{ margin:0,fontSize:24,fontWeight:700,color:"#fff",letterSpacing:-0.6 }}>Mecánica TEK</p>
          <p style={{ margin:"5px 0 0",fontSize:13,color:"rgba(255,255,255,0.5)" }}>Control de Técnicos KOF T1</p>
        </div>
        <div style={{ background:"rgba(255,255,255,0.97)",borderRadius:22,padding:"28px 28px 22px",boxShadow:"0 24px 80px rgba(0,0,0,0.3)" }}>
          {[["Correo electrónico","email","correo@mecanicatek.com",email,v=>setEmail(v)],["Contraseña","password","••••••••",pass,v=>setPass(v)]].map(([lbl,type,ph,val,fn])=>(
            <div key={lbl} style={{ marginBottom:14 }}>
              <label style={{ display:"block",fontSize:11,fontWeight:700,color:"#8E8E93",marginBottom:6,textTransform:"uppercase",letterSpacing:0.6 }}>{lbl}</label>
              <input type={type} value={val} placeholder={ph} onChange={e=>fn(e.target.value)} onKeyDown={e=>e.key==="Enter"&&attempt()} style={iBase} onFocus={e=>e.target.style.borderColor="#2563EB"} onBlur={e=>e.target.style.borderColor="transparent"}/>
            </div>
          ))}
          {err&&<p style={{ color:"#DC2626",fontSize:13,marginBottom:10,fontWeight:500 }}>{err}</p>}
          <button onClick={attempt} disabled={loading} style={{ width:"100%",background:"linear-gradient(135deg,#0B2267 0%,#2563EB 100%)",color:"#fff",border:"none",borderRadius:12,padding:"13px",fontSize:15,fontWeight:700,cursor:"pointer",fontFamily:FONT,opacity:loading?0.7:1 }}>
            {loading?"Verificando...":"Iniciar sesión"}
          </button>
        </div>
        <p style={{ textAlign:"center",marginTop:14,fontSize:11,color:"rgba(255,255,255,0.3)" }}>Mecánica TEK · Control de Técnicos KOF T1</p>
      </div>
    </div>
  );
}

// ─── NAV ─────────────────────────────────────────────────────────────────────
function Nav({ user,view,setView,onLogout }) {
  const canConfig=user.rol==="admin"||user.rol==="encargado";
  return (
    <nav style={{ background:"rgba(255,255,255,0.88)",backdropFilter:"blur(24px)",borderBottom:"1px solid rgba(0,0,0,0.07)",padding:"0 24px",height:52,display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100,fontFamily:FONT }}>
      <div style={{ display:"flex",alignItems:"center",gap:10 }}>
        <div style={{ width:28,height:28,background:"linear-gradient(135deg,#0B2267,#2563EB)",borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center" }}><span style={{ fontSize:14 }}>⚙️</span></div>
        <span style={{ fontWeight:700,fontSize:15,color:"#0B2267",letterSpacing:-0.3 }}>Mecánica TEK</span>
        <span style={{ fontSize:11,color:"#94A3B8" }}>KOF T1</span>
      </div>
      <div style={{ display:"flex",alignItems:"center",gap:6 }}>
        {view!=="dashboard"&&user.rol!=="jefe_taller"&&<button onClick={()=>setView("dashboard")} style={{ fontSize:13,color:"#2563EB",background:"none",border:"none",cursor:"pointer",fontWeight:600,padding:"4px 10px",borderRadius:7,fontFamily:FONT }}>← Dashboard</button>}
        {canConfig&&<button onClick={()=>setView(view==="config"?"dashboard":"config")} style={{ fontSize:13,padding:"5px 14px",borderRadius:8,cursor:"pointer",fontFamily:FONT,fontWeight:600,border:"none",background:view==="config"?"#0B2267":"#EFF6FF",color:view==="config"?"#fff":"#1E40AF" }}>Configuración</button>}
        <div style={{ display:"flex",alignItems:"center",gap:8,padding:"4px 12px 4px 5px",background:"#F1F5F9",borderRadius:20 }}>
          <div style={{ width:24,height:24,background:"linear-gradient(135deg,#0B2267,#2563EB)",borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center" }}><span style={{ color:"#fff",fontSize:11,fontWeight:700 }}>{user.nombre[0]}</span></div>
          <span style={{ fontSize:13,color:"#374151",fontWeight:500 }}>{user.nombre}</span>
          <RolBadge rol={user.rol}/>
        </div>
        <button onClick={onLogout} style={{ fontSize:13,color:"#DC2626",background:"none",border:"none",cursor:"pointer",padding:"4px 8px",fontFamily:FONT }}>Salir</button>
      </div>
    </nav>
  );
}

// ─── CHART TOOLTIP ───────────────────────────────────────────────────────────
function ChartTooltip({ active,payload,label }) {
  if (!active||!payload?.length) return null;
  return <div style={{ background:"#fff",border:"1px solid #E2E8F0",borderRadius:10,padding:"10px 14px",fontFamily:FONT,boxShadow:"0 4px 16px rgba(0,0,0,0.1)" }}><p style={{ margin:"0 0 6px",fontSize:12,fontWeight:700,color:"#374151" }}>{label}</p>{payload.map(p=><p key={p.name} style={{ margin:"3px 0",fontSize:12,color:p.color,fontWeight:600 }}>{p.name}: <span style={{ color:"#1C1C1E" }}>{p.value}</span></p>)}</div>;
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({ talleres,posData,onSelect,mes,setMes,anio,setAnio }) {
  const total=talleres.reduce((s,t)=>s+t.posiciones_autorizadas,0);
  const completados=talleres.filter(t=>(posData[pk(t.id,mes,anio)]||[]).length>0).length;
  const chartData=useMemo(()=>getMockChartData(anio,total),[anio,total]);
  const totalVacantes=talleres.reduce((s,t)=>{ const p=posData[pk(t.id,mes,anio)]||[]; return s+p.filter(x=>x.estatus==="vacante").length; },0);
  const pctVacantes=total>0?Math.round(totalVacantes/total*100):0;
  const exportPDF=()=>{ const s=document.createElement("style"); s.innerHTML=`@media print{nav{display:none!important}body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}`; document.head.appendChild(s); window.print(); setTimeout(()=>document.head.removeChild(s),1000); };
  const cardS={background:"#fff",borderRadius:16,padding:"18px 22px",boxShadow:"0 1px 4px rgba(0,0,0,0.05),0 0 0 1px rgba(0,0,0,0.04)"};
  return (
    <div style={{ padding:"28px 24px",fontFamily:FONT,maxWidth:1040,margin:"0 auto" }}>
      <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:24 }}>
        <div><h1 style={{ margin:0,fontSize:26,fontWeight:700,color:"#0B2267",letterSpacing:-0.9 }}>Dashboard Nacional</h1><p style={{ margin:"4px 0 0",fontSize:13,color:"#94A3B8" }}>Control de Técnicos KOF T1</p></div>
        <div style={{ display:"flex",gap:8,alignItems:"center" }}>
          <PeriodSelector mes={mes} setMes={setMes} anio={anio} setAnio={setAnio}/>
          <button onClick={exportPDF} style={{ display:"flex",alignItems:"center",gap:6,background:"#0B2267",color:"#fff",border:"none",borderRadius:9,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:FONT }}>⬇ PDF</button>
          <ExportBtn onClick={()=>exportDashboardExcel(talleres,posData,mes,anio)} label="Excel"/>
        </div>
      </div>
      <div style={{ display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:22 }}>
        {[{label:"Posiciones autorizadas",val:total,sub:"en 11 talleres",color:"#0B2267"},{label:"Talleres con captura",val:`${completados} / 11`,sub:`${MESES_FULL[mes]} ${anio}`,color:"#2563EB"},{label:"Pendiente de captura",val:11-completados,sub:"talleres sin iniciar",color:11-completados>0?"#F59E0B":"#22C55E"},{label:"Vacantes nacionales",val:`${pctVacantes}%`,sub:`${totalVacantes} posiciones`,color:pctVacantes>10?"#DC2626":"#94A3B8"}].map(c=><div key={c.label} style={cardS}><p style={{ margin:0,fontSize:11,color:"#94A3B8",textTransform:"uppercase",letterSpacing:0.7,fontWeight:700 }}>{c.label}</p><p style={{ margin:"6px 0 2px",fontSize:28,fontWeight:800,color:c.color,letterSpacing:-1 }}>{c.val}</p><p style={{ margin:0,fontSize:11,color:"#94A3B8" }}>{c.sub}</p></div>)}
      </div>
      <div style={{ ...cardS,padding:"22px 24px 16px",marginBottom:22 }}>
        <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:18 }}>
          <div><p style={{ margin:0,fontSize:15,fontWeight:700,color:"#0B2267" }}>Técnicos por mes · {anio}</p><p style={{ margin:"3px 0 0",fontSize:12,color:"#94A3B8" }}>Activos mensuales vs plantilla autorizada</p></div>
          <div style={{ display:"flex",gap:16 }}><div style={{ display:"flex",alignItems:"center",gap:6 }}><span style={{ width:12,height:12,borderRadius:3,background:"#2563EB",display:"inline-block" }}/><span style={{ fontSize:12,color:"#374151",fontWeight:500 }}>Activos</span></div><div style={{ display:"flex",alignItems:"center",gap:6 }}><span style={{ width:22,height:2,background:"#0B2267",display:"inline-block",borderRadius:2 }}/><span style={{ fontSize:12,color:"#374151",fontWeight:500 }}>Autorizados</span></div></div>
        </div>
        <ResponsiveContainer width="100%" height={200}>
          <ComposedChart data={chartData} margin={{ top:4,right:10,left:-10,bottom:0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9" vertical={false}/>
            <XAxis dataKey="mes" tick={{ fontSize:11,fill:"#94A3B8",fontFamily:FONT }} axisLine={false} tickLine={false}/>
            <YAxis tick={{ fontSize:11,fill:"#94A3B8",fontFamily:FONT }} axisLine={false} tickLine={false}/>
            <Tooltip content={<ChartTooltip/>}/>
            <Bar dataKey="activos" name="Activos" fill="#2563EB" radius={[5,5,0,0]} maxBarSize={28}/>
            <Line dataKey="autorizados" name="Autorizados" stroke="#0B2267" strokeWidth={2.5} dot={{ fill:"#0B2267",r:3 }} type="monotone"/>
          </ComposedChart>
        </ResponsiveContainer>
      </div>
      <div style={{ display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10 }}>
        {talleres.map(t=>{ const pos=posData[pk(t.id,mes,anio)]||[]; const activos=pos.length?pos.filter(p=>p.estatus==="activo").length:null; const vacantes=pos.length?pos.filter(p=>p.estatus==="vacante").length:null; const pctVac=t.posiciones_autorizadas>0&&vacantes!=null?Math.round(vacantes/t.posiciones_autorizadas*100):null;
          return <button key={t.id} onClick={()=>onSelect(t)} style={{ background:"#fff",border:"1.5px solid rgba(0,0,0,0.05)",borderRadius:16,padding:"16px 18px",textAlign:"left",cursor:"pointer",boxShadow:"0 1px 3px rgba(0,0,0,0.04)",transition:"all 0.18s",fontFamily:FONT }} onMouseEnter={e=>{e.currentTarget.style.boxShadow="0 6px 24px rgba(11,34,103,0.12)";e.currentTarget.style.borderColor="#BFDBFE";}} onMouseLeave={e=>{e.currentTarget.style.boxShadow="0 1px 3px rgba(0,0,0,0.04)";e.currentTarget.style.borderColor="rgba(0,0,0,0.05)";}}>
            <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12 }}>
              <div><p style={{ margin:0,fontSize:15,fontWeight:700,color:"#0B2267",letterSpacing:-0.4 }}>{t.nombre}</p>{t.jefe_nombre&&<p style={{ margin:"2px 0 0",fontSize:11,color:"#94A3B8" }}>{t.jefe_nombre}</p>}</div>
              <div style={{ display:"flex",gap:6,alignItems:"center" }}>{pctVac!=null&&<span style={{ fontSize:11,fontWeight:700,padding:"3px 8px",borderRadius:6,background:pctVac>10?"#FEE2E2":"#F1F5F9",color:pctVac>10?"#DC2626":"#64748B" }}>{pctVac}% vac.</span>}<div style={{ background:"#EFF6FF",width:34,height:34,borderRadius:9,display:"flex",alignItems:"center",justifyContent:"center" }}><span style={{ color:"#2563EB",fontWeight:800,fontSize:14 }}>{t.posiciones_autorizadas}</span></div></div>
            </div>
            <div style={{ display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6 }}>
              {[{label:"Aut.",val:t.posiciones_autorizadas,color:"#374151"},{label:"Activos",val:activos??"—",color:"#22C55E"},{label:"Vacantes",val:vacantes??"—",color:"#94A3B8"},{label:"Inact.",val:activos!=null?t.posiciones_autorizadas-activos:"—",color:"#F59E0B"}].map(s=><div key={s.label} style={{ background:"#F8FAFC",borderRadius:8,padding:"6px",textAlign:"center" }}><p style={{ margin:0,fontSize:9,color:"#94A3B8",fontWeight:700,textTransform:"uppercase" }}>{s.label}</p><p style={{ margin:"2px 0 0",fontSize:17,fontWeight:800,color:s.color,letterSpacing:-0.5 }}>{s.val}</p></div>)}
            </div>
          </button>; })}
      </div>
    </div>
  );
}

// ─── TALLER VIEW ──────────────────────────────────────────────────────────────
function TallerView({ taller,posiciones,onUpdate,mes,setMes,anio,setAnio,notifConfig,user,onToast }) {
  const [editId,setEditId]=useState(null); const [form,setForm]=useState({}); const [saving,setSaving]=useState(false);
  const activos=posiciones.filter(p=>p.estatus==="activo").length;
  const startEdit=pos=>{setEditId(pos.id);setForm({num_socio:pos.num_socio||"",nombre_tecnico:pos.nombre_tecnico,estatus:pos.estatus,comentario:pos.comentario||""});};
  const cancelEdit=()=>setEditId(null);
  const saveEdit=async(pos)=>{
    setSaving(true);
    const updated={...pos,...form,updated_by:user.nombre};
    await dbUpdatePosicion(updated);
    const changes=[];
    if(form.num_socio!==pos.num_socio) changes.push({campo:"# Socio",anterior:pos.num_socio||"—",nuevo:form.num_socio||"—"});
    if(form.nombre_tecnico!==pos.nombre_tecnico) changes.push({campo:"Nombre técnico",anterior:pos.nombre_tecnico||"—",nuevo:form.nombre_tecnico||"—"});
    if(form.estatus!==pos.estatus) changes.push({campo:"Estatus",anterior:STATUS_CFG[pos.estatus]?.label,nuevo:STATUS_CFG[form.estatus]?.label});
    if(form.comentario!==pos.comentario) changes.push({campo:"Comentario",anterior:pos.comentario||"—",nuevo:form.comentario||"—"});
    for(const c of changes) await dbAddBitacora({ts:tsNow(),usuario:user.nombre,taller:taller.nombre,tipo:"Edición técnico",posicion:`#${pos.numero} Cat.${pos.categoria}`,campo:c.campo,anterior:c.anterior,nuevo:c.nuevo,comentario:form.comentario});
    onUpdate(taller.id,pos.id,form);
    setSaving(false); setEditId(null);
    if(notifConfig.activo&&notifConfig.correo_destino) onToast(`Aviso enviado a ${notifConfig.correo_destino}`);
  };
  const exportExcel=()=>{ const rows=posiciones.map(p=>({ "Periodo":`${MESES_FULL[mes]} ${anio}`,"Taller":taller.nombre,"Posición":p.numero,"Categoría":p.categoria,"# Socio":p.num_socio||"—","Técnico":p.nombre_tecnico||"Sin asignar","Estatus":STATUS_CFG[p.estatus]?.label,"Comentario":p.comentario||"—" })); const ws=XLSX.utils.json_to_sheet(rows); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,taller.nombre); XLSX.writeFile(wb,`MTK_${taller.nombre}_${MESES_FULL[mes]}_${anio}.xlsx`); };
  return (
    <div style={{ padding:"28px 24px",fontFamily:FONT,maxWidth:1040,margin:"0 auto" }}>
      <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:22 }}>
        <div><h1 style={{ margin:0,fontSize:26,fontWeight:700,color:"#0B2267",letterSpacing:-0.9 }}>{taller.nombre}</h1><p style={{ margin:"4px 0 0",fontSize:13,color:"#94A3B8" }}>{taller.posiciones_autorizadas} posiciones{taller.jefe_nombre&&` · ${taller.jefe_nombre}`}</p></div>
        <div style={{ display:"flex",gap:8,alignItems:"center" }}><PeriodSelector mes={mes} setMes={setMes} anio={anio} setAnio={setAnio}/><ExportBtn onClick={exportExcel}/></div>
      </div>
      {notifConfig.activo&&notifConfig.correo_destino&&<div style={{ background:"#EFF6FF",border:"1px solid #BFDBFE",borderRadius:10,padding:"9px 14px",marginBottom:16,display:"flex",alignItems:"center",gap:8 }}><span>📧</span><span style={{ fontSize:12,color:"#1E40AF",fontWeight:500 }}>Notificaciones activas · <strong>{notifConfig.correo_destino}</strong></span></div>}
      <div style={{ display:"flex",gap:8,marginBottom:20,flexWrap:"wrap",alignItems:"center" }}>
        {Object.entries(STATUS_CFG).map(([s,cfg])=>{ const count=posiciones.filter(p=>p.estatus===s).length; return <div key={s} style={{ display:"flex",alignItems:"center",gap:6,background:cfg.bg,padding:"6px 14px",borderRadius:20 }}><span style={{ width:7,height:7,borderRadius:"50%",background:cfg.dot }}/><span style={{ fontSize:12,color:cfg.text,fontWeight:600 }}>{cfg.label}</span><span style={{ fontSize:14,fontWeight:800,color:cfg.text }}>{count}</span></div>; })}
        <span style={{ marginLeft:"auto",fontSize:13,fontWeight:700,padding:"6px 14px",borderRadius:20,background:activos===taller.posiciones_autorizadas?"#DCFCE7":"#FFFBEB",color:activos===taller.posiciones_autorizadas?"#14532D":"#78350F" }}>{activos}/{taller.posiciones_autorizadas} activos</span>
      </div>
      <div style={{ background:"#fff",borderRadius:16,boxShadow:"0 1px 4px rgba(0,0,0,0.05),0 0 0 1px rgba(0,0,0,0.04)",overflow:"hidden" }}>
        <div style={{ display:"grid",gridTemplateColumns:"44px 54px 90px 1fr 150px 1fr 88px",padding:"10px 20px",background:"#F8FAFC",borderBottom:"1px solid #F1F5F9" }}>
          {["#","Cat.","# Socio","Técnico","Estatus","Comentario",""].map((h,i)=><span key={i} style={{ fontSize:11,fontWeight:700,color:"#94A3B8",textTransform:"uppercase",letterSpacing:0.7 }}>{h}</span>)}
        </div>
        {posiciones.map((pos,idx)=>{ const isEditing=editId===pos.id; return (
          <div key={pos.id} style={{ display:"grid",gridTemplateColumns:"44px 54px 90px 1fr 150px 1fr 88px",alignItems:"center",padding:"11px 20px",background:isEditing?"#EFF6FF":idx%2===0?"#fff":"#FAFBFC",borderBottom:"1px solid #F1F5F9" }}>
            <span style={{ fontSize:12,color:"#CBD5E1",fontWeight:700 }}>{String(pos.numero).padStart(2,"0")}</span>
            <span style={{ display:"inline-flex",alignItems:"center",justifyContent:"center",width:28,height:28,background:"#0B2267",color:"#fff",borderRadius:7,fontSize:12,fontWeight:800 }}>{pos.categoria}</span>
            <div style={{ paddingRight:8 }}>{isEditing?<input value={form.num_socio} onChange={e=>setForm(p=>({...p,num_socio:e.target.value}))} placeholder="# Socio" style={{...iBase,fontSize:12,padding:"7px 8px"}} onFocus={e=>e.target.style.borderColor="#2563EB"} onBlur={e=>e.target.style.borderColor="transparent"}/>:<span style={{ fontSize:12,color:pos.num_socio?"#1C1C1E":"#CBD5E1",fontWeight:pos.num_socio?600:400 }}>{pos.num_socio||"—"}</span>}</div>
            <div style={{ paddingRight:12 }}>{isEditing?<input value={form.nombre_tecnico} onChange={e=>setForm(p=>({...p,nombre_tecnico:e.target.value}))} placeholder="Nombre del técnico" style={{...iBase,fontSize:13,padding:"7px 10px"}} onFocus={e=>e.target.style.borderColor="#2563EB"} onBlur={e=>e.target.style.borderColor="transparent"}/>:<span style={{ fontSize:13,color:pos.nombre_tecnico?"#1C1C1E":"#CBD5E1",fontStyle:pos.nombre_tecnico?"normal":"italic" }}>{pos.nombre_tecnico||"Sin asignar"}</span>}</div>
            <div>{isEditing?<select value={form.estatus} onChange={e=>setForm(p=>({...p,estatus:e.target.value}))} style={{ border:"1.5px solid #E2E8F0",borderRadius:8,padding:"6px 10px",fontSize:13,outline:"none",background:"#fff",fontFamily:FONT,cursor:"pointer" }}>{ESTATUSES.map(s=><option key={s} value={s}>{STATUS_CFG[s].label}</option>)}</select>:<Chip estatus={pos.estatus}/>}</div>
            <div style={{ paddingRight:12 }}>{isEditing?<input value={form.comentario} onChange={e=>setForm(p=>({...p,comentario:e.target.value}))} placeholder="Comentario opcional..." style={{...iBase,fontSize:12,padding:"7px 10px"}} onFocus={e=>e.target.style.borderColor="#2563EB"} onBlur={e=>e.target.style.borderColor="transparent"}/>:<span style={{ fontSize:12,color:pos.comentario?"#374151":"#CBD5E1",fontStyle:pos.comentario?"normal":"italic" }}>{pos.comentario||"—"}</span>}</div>
            <div style={{ display:"flex",gap:4,justifyContent:"flex-end" }}>
              {isEditing?<><button onClick={()=>saveEdit(pos)} disabled={saving} style={{ background:"#DCFCE7",color:"#14532D",border:"none",borderRadius:7,padding:"5px 10px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:FONT,opacity:saving?0.6:1 }}>{saving?"...":"✓"}</button><button onClick={cancelEdit} style={{ background:"#F1F5F9",color:"#64748B",border:"none",borderRadius:7,padding:"5px 8px",fontSize:12,cursor:"pointer",fontFamily:FONT }}>✕</button></>
              :<button onClick={()=>startEdit(pos)} style={{ fontSize:12,color:"#2563EB",background:"#EFF6FF",border:"none",borderRadius:7,padding:"5px 12px",cursor:"pointer",fontWeight:600,fontFamily:FONT }}>Editar</button>}
            </div>
          </div>); })}
      </div>
    </div>
  );
}

// ─── CONFIG ───────────────────────────────────────────────────────────────────
function ConfigView({ talleres,setTalleres,notifConfig,setNotifConfig,usuarios,setUsuarios,user }) {
  const [tab,setTab]=useState("plantilla");
  const [editPosId,setEditPosId]=useState(null); const [posVal,setPosVal]=useState("");
  const [editJefeId,setEditJefeId]=useState(null); const [jefeForm,setJefeForm]=useState({jefe_nombre:"",jefe_email:""});
  const [notifForm,setNotifForm]=useState({...notifConfig});
  const [showNewUser,setShowNewUser]=useState(false);
  const [userForm,setUserForm]=useState({nombre:"",email:"",password:"",rol:"jefe_taller",talleres_ids:[],activo:true});
  const [bitacora,setBitacora]=useState([]); const [bitAccesos,setBitAccesos]=useState([]); const [loadingBit,setLoadingBit]=useState(false);
  const total=talleres.reduce((s,t)=>s+t.posiciones_autorizadas,0);
  const cardStyle={background:"#fff",borderRadius:16,overflow:"hidden",boxShadow:"0 1px 4px rgba(0,0,0,0.05),0 0 0 1px rgba(0,0,0,0.04)"};

  const savePos=async id=>{ const t=talleres.find(x=>x.id===id); if(!t) return; const updated={...t,posiciones_autorizadas:Number(posVal)}; await dbSaveTaller(updated); setTalleres(p=>p.map(x=>x.id===id?updated:x)); setEditPosId(null); };
  const saveJefe=async id=>{ const t=talleres.find(x=>x.id===id); if(!t) return; const updated={...t,...jefeForm}; await dbSaveTaller(updated); setTalleres(p=>p.map(x=>x.id===id?updated:x)); setEditJefeId(null); };
  const saveNotif=async()=>{ await dbSaveNotifConfig(notifForm); setNotifConfig({...notifForm}); };
  const addUser=async()=>{ const nu=await dbAddUsuario(userForm); if(nu){ setUsuarios(p=>[...p,nu]); await dbAddBitacoraAcceso({ts:tsNow(),usuario:user.nombre,accion:"Alta de usuario",afectado:userForm.nombre,detalle:`Rol: ${userForm.rol}`}); } setShowNewUser(false); setUserForm({nombre:"",email:"",password:"",rol:"jefe_taller",talleres_ids:[],activo:true}); };
  const toggleUser=async(id,activo)=>{ await dbToggleUsuario(id,!activo); setUsuarios(p=>p.map(u=>u.id===id?{...u,activo:!activo}:u)); };
  const deleteUser=async id=>{ await dbDeleteUsuario(id); setUsuarios(p=>p.filter(u=>u.id!==id)); };
  const toggleTallerUser=tid=>setUserForm(p=>({...p,talleres_ids:p.talleres_ids.includes(tid)?p.talleres_ids.filter(x=>x!==tid):[...p.talleres_ids,tid]}));

  useEffect(()=>{ if(tab==="bitacora"){ setLoadingBit(true); Promise.all([dbLoadBitacora(),dbLoadBitacoraAccesos()]).then(([b,a])=>{ setBitacora(b); setBitAccesos(a); setLoadingBit(false); }); } },[tab]);

  const TABS=[{id:"plantilla",label:"Plantilla"},{id:"jefes",label:"Jefes de taller"},{id:"usuarios",label:"Usuarios"},{id:"notificaciones",label:"Notificaciones"},{id:"bitacora",label:"Bitácora"}];

  return (
    <div style={{ padding:"28px 24px",maxWidth:820,margin:"0 auto",fontFamily:FONT }}>
      <div style={{ marginBottom:24 }}><h1 style={{ margin:0,fontSize:26,fontWeight:700,color:"#0B2267",letterSpacing:-0.9 }}>Configuración</h1><p style={{ margin:"4px 0 0",fontSize:13,color:"#94A3B8" }}>Parámetros del sistema · solo admin y encargado</p></div>
      <div style={{ display:"flex",gap:4,marginBottom:22,background:"#F1F5F9",padding:4,borderRadius:12,flexWrap:"wrap" }}>
        {TABS.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{ padding:"7px 16px",borderRadius:9,border:"none",cursor:"pointer",fontFamily:FONT,fontSize:13,fontWeight:600,transition:"all 0.15s",background:tab===t.id?"#fff":"transparent",color:tab===t.id?"#0B2267":"#94A3B8",boxShadow:tab===t.id?"0 1px 4px rgba(0,0,0,0.1)":"none" }}>{t.label}</button>)}
      </div>

      {tab==="plantilla"&&<>
        <div style={{ background:"#FFFBEB",border:"1px solid #FDE68A",borderRadius:12,padding:"12px 16px",marginBottom:18,display:"flex",gap:10 }}><span>⚠️</span><div><p style={{ margin:0,fontSize:13,fontWeight:700,color:"#78350F" }}>Zona restringida</p><p style={{ margin:"3px 0 0",fontSize:12,color:"#92400E" }}>Modifica solo cuando Coca-Cola autorice una variación oficial en la tarifa.</p></div></div>
        <div style={cardStyle}>
          <div style={{ padding:"12px 20px",borderBottom:"1px solid #F1F5F9",display:"flex",justifyContent:"space-between" }}><span style={{ fontSize:13,fontWeight:700,color:"#374151" }}>Posiciones por taller</span><span style={{ fontSize:12,color:"#94A3B8" }}>Total: {total}</span></div>
          {talleres.map((t,idx)=><div key={t.id} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"14px 20px",borderBottom:idx<talleres.length-1?"1px solid #F1F5F9":"none" }}>
            <p style={{ margin:0,fontSize:14,fontWeight:600,color:"#0B2267" }}>{t.nombre}</p>
            <div style={{ display:"flex",alignItems:"center",gap:12 }}>
              {editPosId===t.id?<><input type="number" min="1" value={posVal} onChange={e=>setPosVal(e.target.value)} style={{ width:70,border:"1.5px solid #2563EB",borderRadius:8,padding:"7px 10px",fontSize:15,fontWeight:700,textAlign:"center",outline:"none",fontFamily:FONT }}/><button onClick={()=>savePos(t.id)} style={{ background:"#DCFCE7",color:"#14532D",border:"none",borderRadius:8,padding:"7px 16px",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:FONT }}>Guardar</button><button onClick={()=>setEditPosId(null)} style={{ background:"#F1F5F9",color:"#64748B",border:"none",borderRadius:8,padding:"7px 12px",fontSize:13,cursor:"pointer",fontFamily:FONT }}>Cancelar</button></>
              :<><div style={{ textAlign:"center" }}><p style={{ margin:0,fontSize:22,fontWeight:800,color:"#0B2267" }}>{t.posiciones_autorizadas}</p><p style={{ margin:0,fontSize:10,color:"#94A3B8" }}>posiciones</p></div><button onClick={()=>{setEditPosId(t.id);setPosVal(t.posiciones_autorizadas);}} style={{ border:"1.5px solid #E2E8F0",background:"#fff",borderRadius:9,padding:"7px 16px",fontSize:13,color:"#374151",cursor:"pointer",fontFamily:FONT }}>Modificar</button></>}
            </div>
          </div>)}
        </div>
      </>}

      {tab==="jefes"&&<div style={cardStyle}>
        <div style={{ padding:"12px 20px",borderBottom:"1px solid #F1F5F9" }}><span style={{ fontSize:13,fontWeight:700,color:"#374151" }}>Jefes de taller</span></div>
        {talleres.map((t,idx)=><div key={t.id} style={{ borderBottom:idx<talleres.length-1?"1px solid #F1F5F9":"none" }}>
          {editJefeId===t.id?<div style={{ padding:"16px 20px",background:"#EFF6FF" }}>
            <p style={{ margin:"0 0 12px",fontSize:14,fontWeight:700,color:"#0B2267" }}>{t.nombre}</p>
            <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12 }}>
              <div><label style={{ display:"block",fontSize:11,fontWeight:700,color:"#8E8E93",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5 }}>Nombre</label><input value={jefeForm.jefe_nombre} onChange={e=>setJefeForm(p=>({...p,jefe_nombre:e.target.value}))} placeholder="Nombre completo" style={{...iBase,fontSize:13,padding:"8px 12px"}} onFocus={e=>e.target.style.borderColor="#2563EB"} onBlur={e=>e.target.style.borderColor="transparent"}/></div>
              <div><label style={{ display:"block",fontSize:11,fontWeight:700,color:"#8E8E93",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5 }}>Correo</label><input value={jefeForm.jefe_email} onChange={e=>setJefeForm(p=>({...p,jefe_email:e.target.value}))} placeholder="correo@empresa.com" style={{...iBase,fontSize:13,padding:"8px 12px"}} onFocus={e=>e.target.style.borderColor="#2563EB"} onBlur={e=>e.target.style.borderColor="transparent"}/></div>
            </div>
            <div style={{ display:"flex",gap:8 }}><button onClick={()=>saveJefe(t.id)} style={{ background:"linear-gradient(135deg,#0B2267,#2563EB)",color:"#fff",border:"none",borderRadius:9,padding:"8px 20px",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:FONT }}>Guardar</button><button onClick={()=>setEditJefeId(null)} style={{ background:"#F1F5F9",color:"#64748B",border:"none",borderRadius:9,padding:"8px 14px",fontSize:13,cursor:"pointer",fontFamily:FONT }}>Cancelar</button></div>
          </div>:<div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"14px 20px" }}>
            <div><p style={{ margin:0,fontSize:14,fontWeight:600,color:"#0B2267" }}>{t.nombre}</p>{t.jefe_nombre?<p style={{ margin:"3px 0 0",fontSize:13,color:"#374151" }}>{t.jefe_nombre} <span style={{ color:"#2563EB" }}>· {t.jefe_email}</span></p>:<p style={{ margin:"3px 0 0",fontSize:12,color:"#CBD5E1",fontStyle:"italic" }}>Sin jefe registrado</p>}</div>
            <button onClick={()=>{setEditJefeId(t.id);setJefeForm({jefe_nombre:t.jefe_nombre||"",jefe_email:t.jefe_email||""});}} style={{ border:"1.5px solid #E2E8F0",background:"#fff",borderRadius:9,padding:"7px 16px",fontSize:13,color:"#374151",cursor:"pointer",fontFamily:FONT }}>{t.jefe_nombre?"Editar":"+ Agregar"}</button>
          </div>}
        </div>)}
      </div>}

      {tab==="usuarios"&&<>
        <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16 }}>
          <p style={{ margin:0,fontSize:13,color:"#64748B" }}>Gestiona accesos a la plataforma por taller.</p>
          <button onClick={()=>setShowNewUser(true)} style={{ background:"linear-gradient(135deg,#0B2267,#2563EB)",color:"#fff",border:"none",borderRadius:9,padding:"8px 18px",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:FONT }}>+ Nuevo usuario</button>
        </div>
        {showNewUser&&<div style={{ background:"#EFF6FF",border:"1px solid #BFDBFE",borderRadius:14,padding:"20px",marginBottom:16 }}>
          <p style={{ margin:"0 0 14px",fontSize:15,fontWeight:700,color:"#0B2267" }}>Nuevo usuario</p>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10 }}>
            <div><label style={{ display:"block",fontSize:11,fontWeight:700,color:"#8E8E93",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5 }}>Nombre</label><input value={userForm.nombre} onChange={e=>setUserForm(p=>({...p,nombre:e.target.value}))} style={{...iBase,fontSize:13,padding:"8px 12px"}} onFocus={e=>e.target.style.borderColor="#2563EB"} onBlur={e=>e.target.style.borderColor="transparent"}/></div>
            <div><label style={{ display:"block",fontSize:11,fontWeight:700,color:"#8E8E93",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5 }}>Correo</label><input value={userForm.email} onChange={e=>setUserForm(p=>({...p,email:e.target.value}))} style={{...iBase,fontSize:13,padding:"8px 12px"}} onFocus={e=>e.target.style.borderColor="#2563EB"} onBlur={e=>e.target.style.borderColor="transparent"}/></div>
            <div><label style={{ display:"block",fontSize:11,fontWeight:700,color:"#8E8E93",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5 }}>Contraseña</label><input value={userForm.password} onChange={e=>setUserForm(p=>({...p,password:e.target.value}))} style={{...iBase,fontSize:13,padding:"8px 12px"}} onFocus={e=>e.target.style.borderColor="#2563EB"} onBlur={e=>e.target.style.borderColor="transparent"}/></div>
            <div><label style={{ display:"block",fontSize:11,fontWeight:700,color:"#8E8E93",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5 }}>Rol</label><select value={userForm.rol} onChange={e=>setUserForm(p=>({...p,rol:e.target.value}))} style={{ border:"1.5px solid #E2E8F0",borderRadius:10,padding:"10px 14px",fontSize:13,outline:"none",background:"#fff",fontFamily:FONT,width:"100%",cursor:"pointer" }}><option value="jefe_taller">Jefe de taller</option><option value="encargado">Encargado</option><option value="admin">Administrador</option></select></div>
          </div>
          {userForm.rol==="jefe_taller"&&<div style={{ marginBottom:12 }}>
            <label style={{ display:"block",fontSize:11,fontWeight:700,color:"#8E8E93",marginBottom:8,textTransform:"uppercase",letterSpacing:0.5 }}>Talleres asignados</label>
            <div style={{ display:"flex",flexWrap:"wrap",gap:8 }}>{talleres.map(t=><button key={t.id} onClick={()=>toggleTallerUser(t.id)} style={{ padding:"6px 12px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontSize:12,fontWeight:600,fontFamily:FONT,background:userForm.talleres_ids.includes(t.id)?"#0B2267":"#fff",color:userForm.talleres_ids.includes(t.id)?"#fff":"#374151",borderColor:userForm.talleres_ids.includes(t.id)?"#0B2267":"#E2E8F0" }}>{t.nombre}</button>)}</div>
          </div>}
          <div style={{ display:"flex",gap:8 }}><button onClick={addUser} disabled={!userForm.nombre||!userForm.email||!userForm.password} style={{ background:"linear-gradient(135deg,#0B2267,#2563EB)",color:"#fff",border:"none",borderRadius:9,padding:"8px 20px",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:FONT,opacity:(!userForm.nombre||!userForm.email||!userForm.password)?0.5:1 }}>Crear usuario</button><button onClick={()=>setShowNewUser(false)} style={{ background:"#F1F5F9",color:"#64748B",border:"none",borderRadius:9,padding:"8px 14px",fontSize:13,cursor:"pointer",fontFamily:FONT }}>Cancelar</button></div>
        </div>}
        <div style={cardStyle}>
          <div style={{ padding:"12px 20px",borderBottom:"1px solid #F1F5F9",display:"flex",justifyContent:"space-between" }}><span style={{ fontSize:13,fontWeight:700,color:"#374151" }}>Usuarios registrados</span><span style={{ fontSize:12,color:"#94A3B8" }}>{usuarios.length} usuarios</span></div>
          {usuarios.map((u,idx)=><div key={u.id} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"14px 20px",borderBottom:idx<usuarios.length-1?"1px solid #F1F5F9":"none",opacity:u.activo?1:0.5 }}>
            <div style={{ display:"flex",alignItems:"center",gap:12 }}>
              <div style={{ width:36,height:36,borderRadius:"50%",background:u.activo?"linear-gradient(135deg,#0B2267,#2563EB)":"#E2E8F0",display:"flex",alignItems:"center",justifyContent:"center" }}><span style={{ color:u.activo?"#fff":"#94A3B8",fontWeight:700,fontSize:14 }}>{u.nombre[0]}</span></div>
              <div><p style={{ margin:0,fontSize:14,fontWeight:600,color:"#1C1C1E" }}>{u.nombre} <RolBadge rol={u.rol}/></p><p style={{ margin:"2px 0 0",fontSize:12,color:"#94A3B8" }}>{u.email}</p></div>
            </div>
            <div style={{ display:"flex",gap:6 }}>
              <button onClick={()=>toggleUser(u.id,u.activo)} style={{ fontSize:12,padding:"5px 12px",borderRadius:7,border:"1.5px solid",cursor:"pointer",fontFamily:FONT,fontWeight:600,background:u.activo?"#FFFBEB":"#DCFCE7",color:u.activo?"#78350F":"#14532D",borderColor:u.activo?"#FDE68A":"#86EFAC" }}>{u.activo?"Desactivar":"Activar"}</button>
              {u.id>3&&<button onClick={()=>deleteUser(u.id)} style={{ fontSize:12,padding:"5px 10px",borderRadius:7,border:"1.5px solid #FEE2E2",cursor:"pointer",fontFamily:FONT,background:"#FFF1F2",color:"#DC2626" }}>✕</button>}
            </div>
          </div>)}
        </div>
      </>}

      {tab==="notificaciones"&&<div style={cardStyle}>
        <div style={{ padding:"12px 20px",borderBottom:"1px solid #F1F5F9" }}><span style={{ fontSize:13,fontWeight:700,color:"#374151" }}>Avisos por correo</span></div>
        <div style={{ padding:"16px 20px",borderBottom:"1px solid #F1F5F9",display:"flex",justifyContent:"space-between",alignItems:"center" }}>
          <div><p style={{ margin:0,fontSize:14,fontWeight:600,color:"#1C1C1E" }}>Activar notificaciones</p><p style={{ margin:"3px 0 0",fontSize:12,color:"#94A3B8" }}>Aviso automático cada vez que un jefe edite un técnico</p></div>
          <div onClick={()=>setNotifForm(p=>({...p,activo:!p.activo}))} style={{ width:50,height:28,borderRadius:14,background:notifForm.activo?"#2563EB":"#D1D5DB",cursor:"pointer",position:"relative",transition:"background 0.2s",flexShrink:0 }}><div style={{ width:22,height:22,borderRadius:"50%",background:"#fff",position:"absolute",top:3,left:notifForm.activo?25:3,transition:"left 0.2s",boxShadow:"0 1px 4px rgba(0,0,0,0.2)" }}/></div>
        </div>
        <div style={{ padding:"16px 20px",borderBottom:"1px solid #F1F5F9",opacity:notifForm.activo?1:0.4 }}>
          <label style={{ display:"block",fontSize:11,fontWeight:700,color:"#8E8E93",marginBottom:8,textTransform:"uppercase",letterSpacing:0.6 }}>Correo destino</label>
          <input value={notifForm.correo_destino||""} onChange={e=>setNotifForm(p=>({...p,correo_destino:e.target.value}))} disabled={!notifForm.activo} placeholder="admin@mecanicatek.com" style={{...iBase,maxWidth:380}} onFocus={e=>e.target.style.borderColor="#2563EB"} onBlur={e=>e.target.style.borderColor="transparent"}/>
          <p style={{ margin:"6px 0 0",fontSize:12,color:"#94A3B8" }}>Separa múltiples correos con comas</p>
        </div>
        <div style={{ padding:"16px 20px",display:"flex",gap:10,alignItems:"center" }}>
          <button onClick={saveNotif} style={{ background:"linear-gradient(135deg,#0B2267,#2563EB)",color:"#fff",border:"none",borderRadius:10,padding:"9px 22px",fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:FONT }}>Guardar</button>
          {notifConfig.activo!==notifForm.activo||notifConfig.correo_destino!==notifForm.correo_destino?<span style={{ fontSize:12,color:"#F59E0B",fontWeight:500 }}>· Cambios sin guardar</span>:<span style={{ fontSize:12,color:"#22C55E",fontWeight:500 }}>· Guardado</span>}
        </div>
      </div>}

      {tab==="bitacora"&&<>
        {loadingBit?<Loader text="Cargando bitácora..."/>:<>
          <div style={{ marginBottom:20 }}>
            <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10 }}>
              <div><p style={{ margin:0,fontSize:15,fontWeight:700,color:"#0B2267" }}>Bitácora de técnicos</p><p style={{ margin:"2px 0 0",fontSize:12,color:"#94A3B8" }}>Últimos 200 movimientos</p></div>
              <ExportBtn onClick={()=>exportBitacora(bitacora)} label="Exportar Excel"/>
            </div>
            <div style={cardStyle}>
              {bitacora.length===0?<div style={{ padding:"28px",textAlign:"center",color:"#94A3B8",fontSize:13 }}>Sin movimientos aún</div>:<>
                <div style={{ display:"grid",gridTemplateColumns:"40px 110px 110px 90px 90px 1fr 1fr",padding:"9px 16px",background:"#F8FAFC",borderBottom:"1px solid #F1F5F9" }}>{["ID","Fecha","Usuario","Taller","Posición","Campo","Cambio"].map((h,i)=><span key={i} style={{ fontSize:10,fontWeight:700,color:"#94A3B8",textTransform:"uppercase",letterSpacing:0.5 }}>{h}</span>)}</div>
                {bitacora.map((r,i)=><div key={r.id} style={{ display:"grid",gridTemplateColumns:"40px 110px 110px 90px 90px 1fr 1fr",alignItems:"center",padding:"9px 16px",borderBottom:"1px solid #F1F5F9",background:i%2===0?"#fff":"#FAFBFC" }}>
                  <span style={{ fontSize:11,color:"#CBD5E1",fontWeight:700 }}>{r.id}</span>
                  <span style={{ fontSize:11,color:"#94A3B8" }}>{r.fecha||r.ts}</span>
                  <span style={{ fontSize:12,fontWeight:600,color:"#374151" }}>{r.usuario}</span>
                  <span style={{ fontSize:12,color:"#374151" }}>{r.taller}</span>
                  <span style={{ fontSize:12,color:"#374151" }}>{r.posicion}</span>
                  <span style={{ fontSize:12,color:"#374151" }}>{r.campo}</span>
                  <span style={{ fontSize:12 }}><span style={{ color:"#DC2626",textDecoration:"line-through",marginRight:4 }}>{r.anterior}</span>→ <span style={{ color:"#22C55E" }}>{r.nuevo}</span></span>
                </div>)}
              </>}
            </div>
          </div>
          <div>
            <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10 }}>
              <div><p style={{ margin:0,fontSize:15,fontWeight:700,color:"#0B2267" }}>Bitácora de accesos</p><p style={{ margin:"2px 0 0",fontSize:12,color:"#94A3B8" }}>Cambios en configuración de usuarios</p></div>
              <ExportBtn onClick={()=>exportAccesos(bitAccesos)} label="Exportar Excel"/>
            </div>
            <div style={cardStyle}>
              {bitAccesos.length===0?<div style={{ padding:"28px",textAlign:"center",color:"#94A3B8",fontSize:13 }}>Sin eventos aún</div>:<>
                <div style={{ display:"grid",gridTemplateColumns:"40px 110px 130px 1fr 1fr",padding:"9px 16px",background:"#F8FAFC",borderBottom:"1px solid #F1F5F9" }}>{["ID","Fecha","Realizado por","Acción","Detalle"].map((h,i)=><span key={i} style={{ fontSize:10,fontWeight:700,color:"#94A3B8",textTransform:"uppercase",letterSpacing:0.5 }}>{h}</span>)}</div>
                {bitAccesos.map((r,i)=><div key={r.id} style={{ display:"grid",gridTemplateColumns:"40px 110px 130px 1fr 1fr",alignItems:"center",padding:"9px 16px",borderBottom:"1px solid #F1F5F9",background:i%2===0?"#fff":"#FAFBFC" }}>
                  <span style={{ fontSize:11,color:"#CBD5E1",fontWeight:700 }}>{r.id}</span>
                  <span style={{ fontSize:11,color:"#94A3B8" }}>{r.fecha||r.ts}</span>
                  <span style={{ fontSize:12,fontWeight:600,color:"#374151" }}>{r.usuario}</span>
                  <span style={{ fontSize:12,color:"#374151" }}>{r.accion}</span>
                  <span style={{ fontSize:12,color:"#374151" }}>{r.detalle}</span>
                </div>)}
              </>}
            </div>
          </div>
        </>}
      </>}
    </div>
  );
}

// ─── ROOT ─────────────────────────────────────────────────────────────────────
export default function App() {
  const [user,         setUser]         = useState(null);
  const [view,         setView]         = useState("login");
  const [talleres,     setTalleres]     = useState([]);
  const [usuarios,     setUsuarios]     = useState([]);
  const [selected,     setSelected]     = useState(null);
  const [posData,      setPosData]      = useState({});
  const [mes,          setMes]          = useState(new Date().getMonth());
  const [anio,         setAnio]         = useState(new Date().getFullYear());
  const [toast,        setToast]        = useState(null);
  const [notifConfig,  setNotifConfig]  = useState({ activo:false, correo_destino:"" });
  const [appLoading,   setAppLoading]   = useState(true);

  useEffect(()=>{
    Promise.all([dbLoadTalleres(), dbLoadUsuarios(), dbLoadNotifConfig()]).then(([t,u,n])=>{
      setTalleres(t); setUsuarios(u); setNotifConfig(n); setAppLoading(false);
    });
  },[]);

  const showToast = msg => { setToast(msg); setTimeout(()=>setToast(null),3200); };

  const login = u => {
    setUser(u);
    if (u.rol==="jefe_taller") { const t=talleres.find(x=>x.id===(u.talleres_ids?.[0])); if(t) doSelect(t,mes,anio); else setView("dashboard"); }
    else setView("dashboard");
  };

  const doSelect = async (t, m=mes, a=anio) => {
    setSelected(t); setView("taller");
    const key = pk(t.id,m,a);
    if (posData[key]) return;
    // load from DB
    let rows = await dbLoadPosiciones(t.id, m, a);
    if (!rows || rows.length===0) {
      // try previous month
      const prev = await dbLoadPosicionesPrevMonth(t.id, m, a);
      if (prev.length>0) {
        rows = prev.map(p=>({ ...p, id:`${t.id}-${m}-${a}-${p.numero}`, mes:m, anio:a, updated_at:new Date().toISOString(), updated_by:"" }));
      } else {
        rows = genPosicionesVacias(t.id, t.posiciones_autorizadas, m, a);
      }
      await dbUpsertPosiciones(rows);
    }
    setPosData(prev=>({...prev,[key]:rows}));
  };

  const updatePos = (tid, posId, form) => {
    const key = pk(tid,mes,anio);
    setPosData(prev=>({...prev,[key]:(prev[key]||[]).map(p=>p.id===posId?{...p,...form}:p)}));
  };

  const logout = () => { setUser(null); setView("login"); setSelected(null); };

  if (appLoading) return <div style={{ minHeight:"100vh",background:"linear-gradient(155deg,#071A52,#2563EB)",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:FONT }}><div style={{ textAlign:"center" }}><div style={{ width:40,height:40,borderRadius:"50%",border:"3px solid rgba(255,255,255,0.3)",borderTop:"3px solid #fff",animation:"spin 0.8s linear infinite",margin:"0 auto 14px" }}/><p style={{ color:"rgba(255,255,255,0.7)",fontSize:14,margin:0 }}>Conectando con Supabase...</p><style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style></div></div>;

  if (!user) return <LoginView onLogin={login}/>;

  return (
    <div style={{ minHeight:"100vh",background:"#F5F5F7",fontFamily:FONT }}>
      <Nav user={user} view={view} setView={setView} onLogout={logout}/>
      {view==="dashboard"&&<Dashboard talleres={talleres} posData={posData} onSelect={(t)=>doSelect(t,mes,anio)} mes={mes} setMes={(m)=>{setMes(m);if(selected)doSelect(selected,m,anio);}} anio={anio} setAnio={(a)=>{setAnio(a);if(selected)doSelect(selected,mes,a);}}/>}
      {view==="taller"&&selected&&<TallerView taller={selected} posiciones={posData[pk(selected.id,mes,anio)]||[]} onUpdate={updatePos} mes={mes} setMes={(m)=>{setMes(m);doSelect(selected,m,anio);}} anio={anio} setAnio={(a)=>{setAnio(a);doSelect(selected,mes,a);}} notifConfig={notifConfig} user={user} onToast={showToast}/>}
      {view==="config"&&<ConfigView talleres={talleres} setTalleres={setTalleres} notifConfig={notifConfig} setNotifConfig={setNotifConfig} usuarios={usuarios} setUsuarios={setUsuarios} user={user}/>}
      <Toast msg={toast}/>
    </div>
  );
}
