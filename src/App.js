import { useState, useMemo, useRef } from "react";

const TAX_RATE = 0.33;
const INDIVIDUAL_RATE = 0.29;
const HISTORY_KEY = "maguje_history";
const APP_PASSWORD = "maguje2026";
const SESSION_KEY = "maguje_auth";

function getWorkingDays(year, month) {
  const days = [];
  const total = new Date(year, month, 0).getDate();
  for (let d = 1; d <= total; d++) {
    if (new Date(year, month - 1, d).getDay() !== 1) days.push(d);
  }
  return days;
}

const DOW_LABELS = ["Dom","","Ter","Qua","Qui","Sex","SÃ¡b"];
const fmt = (v) => (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const fmtShort = (v) => (v||0).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2});

// Parse XLSX attendance file â†’ { "NOME COMPLETO": [day, day, ...] }
async function parseAttendanceFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        // We'll parse using a simple CSV-like approach via SheetJS loaded from CDN
        // Since we can't import SheetJS here, we use a manual binary parse trick
        // Instead, we encode file as base64 and parse in-browser via script tag
        resolve({ raw: e.target.result, name: file.name });
      } catch(err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

const INITIAL_EMPLOYEES = [
  { id: 1,  name: "Douglas Pereira Lima",              role: "GarÃ§om",               sector: "SalÃ£o",   type: "individual", points: 0,  mei: false },
  { id: 2,  name: "Gabriel de Farias Pereira",         role: "GarÃ§om",               sector: "SalÃ£o",   type: "individual", points: 0,  mei: false },
  { id: 3,  name: "Gustavo Fabricio Rodrigues Freire", role: "GarÃ§om",               sector: "SalÃ£o",   type: "individual", points: 0,  mei: false },
  { id: 4,  name: "Antonia Erineuda",                  role: "GarÃ§onete",            sector: "SalÃ£o",   type: "individual", points: 0,  mei: false },
  { id: 5,  name: "Paulo Alves de Almeida",            role: "GarÃ§om",               sector: "SalÃ£o",   type: "individual", points: 0,  mei: false },
  { id: 6,  name: "Reinaldo Alves de Oliveira",        role: "Chefe de Fila Junior", sector: "SalÃ£o",   type: "individual", points: 0,  mei: false },
  { id: 7,  name: "Jeane Rodrigues",                   role: "Chefe de Fila Junior", sector: "SalÃ£o",   type: "individual", points: 0,  mei: false },
  { id: 8,  name: "Verinaldo Gabriel da Rocha",        role: "Chefe de Fila Junior", sector: "SalÃ£o",   type: "individual", points: 0,  mei: false },
  { id: 9,  name: "Jean Carlos Fidelis",               role: "Cumim",                sector: "SalÃ£o",   type: "global",     points: 15, mei: false },
  { id: 10, name: "Claudia Elisabete ConceiÃ§Ã£o",       role: "Cumim",                sector: "SalÃ£o",   type: "global",     points: 15, mei: false },
  { id: 11, name: "Marcos Vinicius Henrique de Souza", role: "Cumim",                sector: "SalÃ£o",   type: "global",     points: 15, mei: false },
  { id: 12, name: "Maria Elenice Ferreira",            role: "Cumim",                sector: "SalÃ£o",   type: "global",     points: 15, mei: false },
  { id: 13, name: "Rodrigo de Pinho Ribeiro",          role: "Cumim",                sector: "SalÃ£o",   type: "global",     points: 15, mei: false },
  { id: 14, name: "Felipe Costa de Abreu",             role: "Suiteiro",             sector: "SalÃ£o",   type: "global",     points: 20, mei: false },
  { id: 15, name: "Crislandia Moura de Lima",          role: "Chefe de Fila Pleno",  sector: "SalÃ£o",   type: "global",     points: 26, mei: false },
  { id: 16, name: "Elizangelo Araujo Miranda",         role: "MaÃ®tre",               sector: "SalÃ£o",   type: "global",     points: 30, mei: false },
  { id: 17, name: "Joaquim Fernandes Gomes",           role: "Assistente Gerente",   sector: "SalÃ£o",   type: "global",     points: 25, mei: false },
  { id: 18, name: "Jose Edilson Pereira Nogueira",     role: "Sub Gerente / MaÃ®tre", sector: "SalÃ£o",   type: "global",     points: 30, mei: false },
  { id: 19, name: "Rodrigo Florentino Fonseca",        role: "Gerente",              sector: "SalÃ£o",   type: "global",     points: 35, mei: false },
  { id: 20, name: "Fabio da Silva Miguel",             role: "Assistente MKT",       sector: "SalÃ£o",   type: "global",     points: 15, mei: false },
  { id: 21, name: "Kayllana Vitoria de Oliveira",      role: "Hostess",              sector: "SalÃ£o",   type: "global",     points: 15, mei: false },
  { id: 22, name: "Suzana Radai Estrela Souza",        role: "Hostess",              sector: "SalÃ£o",   type: "global",     points: 15, mei: false },
  { id: 23, name: "Romenia Fernades Jorge",            role: "Hostess",              sector: "SalÃ£o",   type: "global",     points: 20, mei: false },
  { id: 24, name: "Danilo Silva Gomes",                role: "Barback",              sector: "Bar",     type: "global",     points: 17, mei: false },
  { id: 25, name: "Luan Chrystyan dos Santos",         role: "Barback",              sector: "Bar",     type: "global",     points: 17, mei: false },
  { id: 26, name: "Francisco Tome da Silva",           role: "Copeiro II",           sector: "Bar",     type: "global",     points: 10, mei: false },
  { id: 27, name: "Antonio Mauricio Santos Soares",    role: "Bartender",            sector: "Bar",     type: "global",     points: 20, mei: false },
  { id: 28, name: "Rafael da Silva Romualdo",          role: "Bartender",            sector: "Bar",     type: "global",     points: 20, mei: false },
  { id: 29, name: "Caio Henriques Rodrigues",          role: "Bartender",            sector: "Bar",     type: "global",     points: 20, mei: false },
  { id: 30, name: "Gabriel Paulino Barbosa",           role: "Bartender",            sector: "Bar",     type: "global",     points: 20, mei: false },
  { id: 31, name: "Gabriel Soares Grativol",           role: "Bartender",            sector: "Bar",     type: "global",     points: 20, mei: false },
  { id: 32, name: "Gabriel de Oliveira Fernandes",     role: "Sub Chefe de Bar",     sector: "Bar",     type: "global",     points: 22, mei: false },
  { id: 33, name: "Luiz Gustavo Mesquita Soares",      role: "Chefe de Bar",         sector: "Bar",     type: "global",     points: 25, mei: false },
  { id: 34, name: "Antonia Jacilane de Sousa Costa",   role: "Caixa",                sector: "Caixa",   type: "global",     points: 15, mei: false },
  { id: 35, name: "Antonio Gomes de Sousa",            role: "Copeiro",              sector: "Cozinha", type: "global",     points: 15, mei: false },
  { id: 36, name: "Douglas Leite GonÃ§alves",           role: "Copeiro",              sector: "Cozinha", type: "global",     points: 15, mei: false },
  { id: 37, name: "John Victor Santos do Nascimento",  role: "Copeiro",              sector: "Cozinha", type: "global",     points: 15, mei: false },
  { id: 38, name: "Rosangela Costa Rodrigues",         role: "Copeiro",              sector: "Cozinha", type: "global",     points: 15, mei: false },
  { id: 39, name: "Robert Gustavo Santos de Souza",    role: "Copeiro",              sector: "Cozinha", type: "global",     points: 15, mei: false },
  { id: 40, name: "Daniel Pereira do Sacramento",      role: "Padeiro",              sector: "Cozinha", type: "global",     points: 15, mei: false },
  { id: 41, name: "Rosinaldo Pedro Soares",            role: "Ajudante de Cozinha",  sector: "Cozinha", type: "global",     points: 15, mei: false },
  { id: 42, name: "Thaynara Tonelle Costa",            role: "Cozinheiro I",         sector: "Cozinha", type: "global",     points: 15, mei: false },
  { id: 43, name: "Andre Felizardo Verissimo",         role: "Cozinheiro I",         sector: "Cozinha", type: "global",     points: 15, mei: false },
  { id: 44, name: "Dayveson Rafael da Silva",          role: "Cozinheiro I",         sector: "Cozinha", type: "global",     points: 15, mei: false },
  { id: 45, name: "Andriely Firmino da Silva",         role: "Cozinheiro I",         sector: "Cozinha", type: "global",     points: 15, mei: false },
  { id: 46, name: "Vitor Faria de Oliveira Aguilera",  role: "Cozinheiro I",         sector: "Cozinha", type: "global",     points: 15, mei: false },
  { id: 47, name: "Lucas Barbosa Ribeiro Borges",      role: "Cozinheiro II",        sector: "Cozinha", type: "global",     points: 19, mei: false },
  { id: 48, name: "Robson Roberto da Silva",           role: "Cozinheiro II",        sector: "Cozinha", type: "global",     points: 19, mei: false },
  { id: 49, name: "Wagner Pinto",                      role: "Cozinheiro III",       sector: "Cozinha", type: "global",     points: 21, mei: false },
  { id: 50, name: "Francisco Dalvan Bezerra Gomes",    role: "Cozinheiro III",       sector: "Cozinha", type: "global",     points: 21, mei: false },
  { id: 51, name: "Valdemir Galdino de Oliveira",      role: "Cozinheiro III",       sector: "Cozinha", type: "global",     points: 21, mei: false },
  { id: 52, name: "Luis Augusto Souza da Costa",       role: "Cozinheiro LÃ­der",     sector: "Cozinha", type: "global",     points: 23, mei: false },
  { id: 53, name: "Jaqueline de Souza Galvao",         role: "Sub Chefe Cozinha",    sector: "Cozinha", type: "global",     points: 25, mei: false },
  { id: 54, name: "Eduardo",                           role: "Chef ProduÃ§Ãµes Gast.", sector: "Cozinha", type: "global",     points: 30, mei: false },
  { id: 55, name: "Alex dos Santos",                   role: "ASG",                  sector: "Limpeza", type: "global",     points: 10, mei: false },
  { id: 56, name: "Carlos Daniel Alves de Lima",       role: "ASG",                  sector: "Limpeza", type: "global",     points: 10, mei: false },
  { id: 57, name: "Marlucia Santana Rodrigues",        role: "LÃ­der de ASG",         sector: "Limpeza", type: "global",     points: 20, mei: false },
  { id: 58, name: "JosÃ© Roberto InÃ¡cio da Silva",      role: "Estoquista",           sector: "Limpeza", type: "global",     points: 10, mei: false },
];

const SECTORS = ["Todos","SalÃ£o","Bar","Caixa","Cozinha","Limpeza"];
const SECTOR_COLORS = { SalÃ£o:"#2D6A4F", Bar:"#1B4332", Caixa:"#40916C", Cozinha:"#B5450B", Limpeza:"#7B5EA7" };

function loadHistory() {
  try { return JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]"); } catch { return []; }
}
function saveHistory(records) { localStorage.setItem(HISTORY_KEY, JSON.stringify(records)); }

function normalize(str) {
  return str.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim();
}

// Match attendance file name to employee list
function matchEmployee(attendanceName, employees) {
  const normAttendance = normalize(attendanceName);
  return employees.find(e => {
    const normEmp = normalize(e.name);
    // exact match
    if (normEmp === normAttendance) return true;
    // check if all words of employee name appear in attendance name
    const empWords = normEmp.split(" ").filter(w => w.length > 2);
    const matchCount = empWords.filter(w => normAttendance.includes(w)).length;
    return matchCount >= Math.min(2, empWords.length);
  });
}

// Parse day string "02/04 QUI" â†’ day number (2)
function parseDay(diaStr) {
  const match = diaStr.match(/^(\d{2})\/(\d{2})/);
  if (match) return parseInt(match[1]);
  return null;
}

function calcResults(employees, workDays, dailyRevenue, globalMonthlyRevenue, addicionais, absences) {
  const indivEmployees = employees.filter(e => e.type === "individual");
  const globalEmployees = employees.filter(e => e.type === "global");
  const isAbsent = (empId, day) => !!(absences[empId]||{})[day];
  const getDR = (day) => dailyRevenue[day] || { individual: {} };

  const empTotals = {};
  employees.forEach(e => empTotals[e.id] = 0);
  let totalBruto = 0, totalIndivComm = 0, totalGlobalPool = 0;

  // Global pool: monthly value after 33% tax
  const globalBruto = parseFloat(globalMonthlyRevenue) || 0;
  const globalNet = globalBruto * (1 - TAX_RATE);
  totalBruto += globalBruto;
  totalGlobalPool += globalNet;

  // Adicionais: no tax deduction, goes straight into global pool
  const adicionaisVal = parseFloat(addicionais) || 0;
  totalBruto += adicionaisVal; // shown in bruto for transparency
  totalGlobalPool += adicionaisVal; // full value enters pool

  const totalGlobalNetForDist = globalNet + adicionaisVal;
  const daysCount = workDays.length;
  const globalPerDay = daysCount > 0 ? totalGlobalNetForDist / daysCount : 0;

  // Individual commissions â€” per day per employee
  workDays.forEach(day => {
    const dr = getDR(day);
    indivEmployees.forEach(emp => {
      const sale = parseFloat((dr.individual||{})[emp.id]) || 0;
      totalBruto += sale;
      if (isAbsent(emp.id, day)) return;
      const net = emp.mei ? sale : sale * (1 - TAX_RATE);
      const comm = net * INDIVIDUAL_RATE;
      empTotals[emp.id] = (empTotals[emp.id]||0) + comm;
      totalIndivComm += comm;
    });
  });

  // Global pool distributed per day by attendance
  workDays.forEach(day => {
    const g1Pool = globalPerDay * 0.73;
    const g2Pool = globalPerDay * 0.27;
    const g1Present = globalEmployees.filter(e => ["SalÃ£o","Bar","Caixa"].includes(e.sector) && !isAbsent(e.id,day));
    const g2Present = globalEmployees.filter(e => ["Cozinha","Limpeza"].includes(e.sector) && !isAbsent(e.id,day));
    const g1Pts = g1Present.reduce((s,e)=>s+e.points,0);
    const g2Pts = g2Present.reduce((s,e)=>s+e.points,0);
    g1Present.forEach(e => { if(g1Pts>0) empTotals[e.id]=(empTotals[e.id]||0)+(e.points/g1Pts)*g1Pool; });
    g2Present.forEach(e => { if(g2Pts>0) empTotals[e.id]=(empTotals[e.id]||0)+(e.points/g2Pts)*g2Pool; });
  });

  return { empTotals, totalBruto, totalIndivComm, totalGlobalPool };
}

// â”€â”€ LOGIN â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function LoginScreen({ onLogin }) {
  const [pwd, setPwd] = useState(""); const [error, setError] = useState(false); const [shake, setShake] = useState(false);
  const handleSubmit = (e) => { e.preventDefault(); if (pwd===APP_PASSWORD){sessionStorage.setItem(SESSION_KEY,"1");onLogin();}else{setError(true);setShake(true);setPwd("");setTimeout(()=>setShake(false),500);} };
  return (
    <div style={{fontFamily:"'DM Mono','Courier New',monospace",background:"#F5F0E8",minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Space+Grotesk:wght@500;700&display=swap');*{box-sizing:border-box}@keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-8px)}40%,80%{transform:translateX(8px)}}.shake{animation:shake 0.4s ease}input:focus{outline:none;border-color:#2D6A4F!important}`}</style>
      <div style={{width:"100%",maxWidth:380,padding:"0 20px"}}>
        <div style={{textAlign:"center",marginBottom:36}}>
          <div style={{width:56,height:56,background:"#1B4332",borderRadius:12,margin:"0 auto 16px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:26}}>ðŸŒ¿</div>
          <div style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:22,color:"#1B4332"}}>Maguje</div>
          <div style={{fontSize:12,color:"#888",marginTop:4}}>Sistema de ComissÃµes</div>
        </div>
        <div className={shake?"shake":""} style={{background:"#fff",border:"1.5px solid #D4CFC4",borderRadius:6,padding:"28px 28px 24px"}}>
          <div style={{fontSize:13,fontWeight:500,color:"#333",marginBottom:20}}>Acesso restrito</div>
          <form onSubmit={handleSubmit}>
            <div style={{marginBottom:16}}>
              <div style={{fontSize:11,color:"#666",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:6}}>Senha</div>
              <input type="password" placeholder="â€¢â€¢â€¢â€¢â€¢â€¢â€¢â€¢" value={pwd} autoFocus onChange={e=>{setPwd(e.target.value);setError(false);}}
                style={{background:"#F5F0E8",border:`1.5px solid ${error?"#c0392b":"#ccc"}`,borderRadius:3,padding:"10px 12px",fontFamily:"inherit",fontSize:14,width:"100%",letterSpacing:"0.1em"}}/>
              {error&&<div style={{fontSize:11,color:"#c0392b",marginTop:6}}>Senha incorreta.</div>}
            </div>
            <button type="submit" style={{width:"100%",border:"none",background:"#1B4332",color:"#fff",padding:"11px",borderRadius:3,cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:500}}>Entrar â†’</button>
          </form>
        </div>
      </div>
    </div>
  );
}

// â”€â”€ EMPLOYEE FORM â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function EmployeeForm({ initial, onSave, onCancel }) {
  const [emp, setEmp] = useState({...initial});
  const f = (k,v) => setEmp(p=>({...p,[k]:v}));
  const inp = {background:"#F5F0E8",border:"1px solid #ccc",borderRadius:3,padding:"6px 8px",fontFamily:"inherit",fontSize:12,width:"100%"};
  return (
    <div style={{background:"#fff",border:"1.5px solid #52B788",borderRadius:4,padding:"16px 18px",marginBottom:16}}>
      <div style={{fontSize:11,color:"#1B4332",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:12}}>{initial.id?"Editar FuncionÃ¡rio":"Novo FuncionÃ¡rio"}</div>
      <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end"}}>
        <div style={{flex:"2 1 160px"}}><div style={{fontSize:11,color:"#555",marginBottom:4}}>Nome</div><input type="text" value={emp.name} onChange={e=>f("name",e.target.value)} placeholder="Nome completo" style={inp}/></div>
        <div style={{flex:"1 1 120px"}}><div style={{fontSize:11,color:"#555",marginBottom:4}}>Cargo</div><input type="text" value={emp.role} onChange={e=>f("role",e.target.value)} placeholder="Cargo" style={inp}/></div>
        <div style={{flex:"1 1 100px"}}><div style={{fontSize:11,color:"#555",marginBottom:4}}>Setor</div>
          <select value={emp.sector} onChange={e=>f("sector",e.target.value)} style={inp}>{["SalÃ£o","Bar","Caixa","Cozinha","Limpeza"].map(x=><option key={x}>{x}</option>)}</select></div>
        <div style={{flex:"1 1 130px"}}><div style={{fontSize:11,color:"#555",marginBottom:4}}>Tipo</div>
          <select value={emp.type} onChange={e=>f("type",e.target.value)} style={inp}><option value="individual">Individual (GarÃ§om/Chefe)</option><option value="global">Global (Pool por pontos)</option></select></div>
        {emp.type==="global"&&<div style={{flex:"0 1 80px"}}><div style={{fontSize:11,color:"#555",marginBottom:4}}>Pontos</div><input type="number" min="1" max="50" value={emp.points} onChange={e=>f("points",parseInt(e.target.value)||1)} style={inp}/></div>}
        <div style={{flex:"0 1 80px"}}>
          <div style={{fontSize:11,color:"#555",marginBottom:4}}>MEI</div>
          <label style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer",padding:"6px 0"}}>
            <div onClick={()=>f("mei",!emp.mei)} style={{width:36,height:20,borderRadius:10,background:emp.mei?"#2D6A4F":"#ccc",position:"relative",transition:"background 0.2s",cursor:"pointer",flexShrink:0}}>
              <div style={{position:"absolute",top:2,left:emp.mei?18:2,width:16,height:16,borderRadius:"50%",background:"#fff",transition:"left 0.2s"}}/>
            </div>
            <span style={{fontSize:12,color:emp.mei?"#2D6A4F":"#888"}}>{emp.mei?"Sim":"NÃ£o"}</span>
          </label>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>onSave(emp)} style={{border:"1.5px solid #1B4332",background:"#1B4332",color:"#fff",padding:"7px 16px",borderRadius:3,cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:500}}>{initial.id?"Salvar":"Adicionar"}</button>
          <button onClick={onCancel} style={{border:"1.5px solid #ccc",background:"transparent",color:"#666",padding:"7px 12px",borderRadius:3,cursor:"pointer",fontFamily:"inherit",fontSize:13}}>Cancelar</button>
        </div>
      </div>
    </div>
  );
}

// â”€â”€ ROOT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
export default function App() {
  const [authed, setAuthed] = useState(()=>sessionStorage.getItem(SESSION_KEY)==="1");
  if (!authed) return <LoginScreen onLogin={()=>setAuthed(true)} />;
  return <MainApp />;
}

function MainApp() {
  const now = new Date();
  const [month, setMonth] = useState(`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`);
  const [employees, setEmployees] = useState(INITIAL_EMPLOYEES);
  const [step, setStep] = useState("revenue");
  const [sector, setSector] = useState("Todos");
  const [dailyRevenue, setDailyRevenue] = useState({});
  const [globalMonthlyRevenue, setGlobalMonthlyRevenue] = useState("");
  const [adicionais, setAdicionais] = useState("");
  const [absences, setAbsences] = useState({});
  const [showAdd, setShowAdd] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [history, setHistory] = useState(loadHistory);
  const [viewingHistory, setViewingHistory] = useState(null);
  const [showHistory, setShowHistory] = useState(false);
  const [attendanceFile, setAttendanceFile] = useState(null); // { name, unmatched }
  const [importLog, setImportLog] = useState(null);
  const printRef = useRef();
  const fileInputRef = useRef();

  const [year, mon] = month.split("-").map(Number);
  const workDays = getWorkingDays(year, mon);
  const monthLabel = new Date(year,mon-1,2).toLocaleString("pt-BR",{month:"long",year:"numeric"});

  const getDR = (day) => dailyRevenue[day]||{individual:{}};
  const setIR = (day,eid,val) => setDailyRevenue(p=>({...p,[day]:{...getDR(day),individual:{...getDR(day).individual,[eid]:val}}}));
  const toggleAbs = (eid,day) => setAbsences(p=>({...p,[eid]:{...(p[eid]||{}),[day]:!(p[eid]||{})[day]}}));
  const isAbsent = (eid,day) => !!(absences[eid]||{})[day];
  const absCount = (eid) => workDays.filter(d=>isAbsent(eid,d)).length;
  const indivEmps = employees.filter(e=>e.type==="individual");

  const results = useMemo(()=>calcResults(employees,workDays,dailyRevenue,globalMonthlyRevenue,adicionais,absences),[dailyRevenue,globalMonthlyRevenue,adicionais,absences,employees,month]);

  const activeHistory = viewingHistory ? history.find(h=>h.monthKey===viewingHistory) : null;
  const displayResults = activeHistory ? activeHistory.results : results;
  const displayEmployees = activeHistory ? activeHistory.employees : employees;
  const displayMonthLabel = activeHistory ? new Date(activeHistory.year,activeHistory.mon-1,2).toLocaleString("pt-BR",{month:"long",year:"numeric"}) : monthLabel;
  const filteredEmps = sector==="Todos" ? displayEmployees : displayEmployees.filter(e=>e.sector===sector);

  // â”€â”€ IMPORT ATTENDANCE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  const handleAttendanceImport = (file) => {
    if (!file) return;
    // Use SheetJS from CDN (loaded via script in index.html)
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const XLSX = window.XLSX;
        if (!XLSX) { alert("Biblioteca de leitura nÃ£o carregada. Recarregue a pÃ¡gina."); return; }
        const wb = XLSX.read(e.target.result, { type: "array" });
        const matched = [];
        const unmatched = [];
        const newAbsences = {...absences};

        wb.SheetNames.forEach(sheetName => {
          const ws = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
          rows.forEach((row, idx) => {
            if (idx === 0) return; // skip header
            const empName = String(row[0]||"").trim();
            const diaStr = String(row[3]||"").trim();
            const tipo = String(row[4]||"").trim();
            if (!empName || !diaStr || !tipo || tipo === "undefined") return;
            const day = parseDay(diaStr);
            if (!day) return;
            // ANY value in TIPO column = falta
            const emp = matchEmployee(empName, employees);
            if (emp) {
              matched.push({ empName, day, tipo });
              if (!newAbsences[emp.id]) newAbsences[emp.id] = {};
              newAbsences[emp.id][day] = true;
            } else {
              if (!unmatched.find(u=>u===empName)) unmatched.push(empName);
            }
          });
        });

        setAbsences(newAbsences);
        setAttendanceFile({ name: file.name });
        setImportLog({ matched: matched.length, unmatched, total: matched.length + unmatched.length });
      } catch(err) {
        alert("Erro ao ler o arquivo: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleSaveHistory = () => {
    const rec = {monthKey:month,year,mon,monthLabel,savedAt:new Date().toISOString(),employees,results,totalBruto:results.totalBruto};
    const updated = [rec,...history.filter(h=>h.monthKey!==month)].slice(0,24);
    setHistory(updated); saveHistory(updated);
    alert(`ComissÃµes de ${monthLabel} salvas!`);
  };

  const handleAddEmployee = (emp) => {
    const id = Date.now();
    setEmployees(p=>[...p,{...emp,id,points:emp.type==="global"?(parseInt(emp.points)||15):0}]);
    setShowAdd(false);
  };

  const handleEditEmployee = (emp) => {
    setEmployees(p=>p.map(e=>e.id===emp.id?{...emp,points:emp.type==="global"?(parseInt(emp.points)||15):0}:e));
    setEditingId(null);
  };

  const toggleMei = (id) => setEmployees(p=>p.map(e=>e.id===id?{...e,mei:!e.mei}:e));

  const handlePrint = () => {
    const content = printRef.current;
    const win = window.open("","_blank");
    win.document.write(`<html><head><title>ComissÃµes ${displayMonthLabel} â€” Maguje</title>
    <style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;font-size:11px;color:#111;padding:20px}h1{font-size:16px;color:#1B4332;margin-bottom:4px}.sub{font-size:11px;color:#666;margin-bottom:16px}.summary{display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap}.sum-card{border:1px solid #ccc;border-radius:4px;padding:8px 14px;min-width:130px}.sum-val{font-size:14px;font-weight:700;color:#1B4332}.sum-lbl{font-size:9px;color:#888;text-transform:uppercase;letter-spacing:.06em;margin-top:2px}table{width:100%;border-collapse:collapse;margin-bottom:20px}th{background:#1B4332;color:#fff;padding:7px 10px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.05em}td{padding:6px 10px;border-bottom:1px solid #eee;font-size:11px}tr:nth-child(even) td{background:#fafaf8}.comm{font-weight:700;text-align:right;color:#1B4332}.mei-tag{background:#fff3cd;color:#856404;border:1px solid #ffc107;border-radius:10px;padding:1px 6px;font-size:9px;font-weight:600}tfoot td{background:#e8f0eb!important;font-weight:700;border-top:2px solid #1B4332}@media print{body{padding:10px}}</style>
    </head><body>${content.innerHTML}</body></html>`);
    win.document.close();
    setTimeout(()=>{win.focus();win.print();},400);
  };

  const S = {
    th:{padding:"9px 12px",textAlign:"left",fontSize:11,textTransform:"uppercase",letterSpacing:"0.06em",background:"#1B4332",color:"#fff",whiteSpace:"nowrap"},
    td:{padding:"8px 12px",fontSize:13,borderBottom:"1px solid #eee",verticalAlign:"middle"},
    input:{background:"#F5F0E8",border:"1px solid #ccc",borderRadius:3,padding:"5px 8px",fontFamily:"inherit",fontSize:12,width:"100%"},
    btn:{border:"1.5px solid #1B4332",background:"#1B4332",color:"#fff",padding:"8px 16px",borderRadius:3,cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:500},
    btnOut:{border:"1.5px solid #1B4332",background:"transparent",color:"#1B4332",padding:"8px 14px",borderRadius:3,cursor:"pointer",fontFamily:"inherit",fontSize:13},
    btnGreen:{border:"1.5px solid #40916C",background:"#40916C",color:"#fff",padding:"8px 14px",borderRadius:3,cursor:"pointer",fontFamily:"inherit",fontSize:13},
    btnSm:(c)=>({border:`1px solid ${c}`,background:"transparent",color:c,padding:"3px 9px",borderRadius:3,cursor:"pointer",fontFamily:"inherit",fontSize:11}),
    tab:{padding:"6px 14px",border:"1px solid #ccc",borderRadius:20,cursor:"pointer",fontSize:12,background:"transparent",fontFamily:"inherit"},
    tabA:{padding:"6px 14px",border:"1px solid #1B4332",borderRadius:20,cursor:"pointer",fontSize:12,background:"#1B4332",color:"#fff",fontFamily:"inherit"},
    stepBtn:(a)=>({padding:"8px 20px",border:"none",borderBottom:a?"2px solid #1B4332":"2px solid transparent",background:"transparent",cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:a?600:400,color:a?"#1B4332":"#888"}),
  };

  const MeiTag = ({active}) => active ? <span style={{fontSize:10,padding:"1px 7px",borderRadius:20,background:"#fff3cd",color:"#856404",border:"1px solid #ffc107",fontWeight:600}}>MEI</span> : null;

  // Print content
  const PrintContent = () => {
    const sectors = ["SalÃ£o","Bar","Caixa","Cozinha","Limpeza"];
    const totalComm = displayEmployees.reduce((s,e)=>s+(displayResults.empTotals[e.id]||0),0);
    return (
      <div>
        <h1>Restaurante Maguje â€” ComissÃµes</h1>
        <div className="sub">ReferÃªncia: {displayMonthLabel} Â· Gerado em {new Date().toLocaleDateString("pt-BR")}</div>
        <div className="summary">
          {[{label:"Total Bruto",val:fmt(displayResults.totalBruto)},{label:"Pool Global LÃ­q.",val:fmt(displayResults.totalGlobalPool)},{label:"ComissÃµes Indiv.",val:fmt(displayResults.totalIndivComm)},{label:"Total DistribuÃ­do",val:fmt(totalComm)}].map(m=>(
            <div key={m.label} className="sum-card"><div className="sum-val">{m.val}</div><div className="sum-lbl">{m.label}</div></div>
          ))}
        </div>
        {sectors.map(sec=>{
          const secEmps=displayEmployees.filter(e=>e.sector===sec);
          if(!secEmps.length) return null;
          const secTotal=secEmps.reduce((s,e)=>s+(displayResults.empTotals[e.id]||0),0);
          return (
            <table key={sec}>
              <thead><tr><th colSpan={5} style={{background:"#2D6A4F"}}>{sec}</th></tr>
              <tr><th>FuncionÃ¡rio</th><th>Cargo</th><th>Tipo</th><th>Faltas</th><th style={{textAlign:"right"}}>ComissÃ£o</th></tr></thead>
              <tbody>{secEmps.sort((a,b)=>(displayResults.empTotals[b.id]||0)-(displayResults.empTotals[a.id]||0)).map(emp=>(
                <tr key={emp.id}><td>{emp.name}{emp.mei?" MEI":""}</td><td>{emp.role}</td>
                  <td>{emp.type==="individual"?"Individual (29%)":"Global ("+emp.points+"pts)"}</td>
                  <td>{absCount(emp.id)>0?absCount(emp.id)+" dia(s)":"â€”"}</td>
                  <td className="comm">{fmt(displayResults.empTotals[emp.id]||0)}</td></tr>
              ))}</tbody>
              <tfoot><tr><td colSpan={4}>Total {sec}</td><td className="comm">{fmt(secTotal)}</td></tr></tfoot>
            </table>
          );
        })}
        <table><tfoot><tr style={{background:"#1B4332",color:"#fff"}}><td colSpan={4} style={{fontWeight:700,fontSize:13,padding:"8px 10px"}}>TOTAL GERAL</td><td style={{fontWeight:700,fontSize:14,textAlign:"right",padding:"8px 10px"}}>{fmt(totalComm)}</td></tr></tfoot></table>
      </div>
    );
  };

  return (
    <div style={{fontFamily:"'DM Mono','Courier New',monospace",background:"#F5F0E8",minHeight:"100vh"}}>
      <style>{`*{box-sizing:border-box}input:focus{outline:none;border-color:#2D6A4F!important}.absent-btn{width:28px;height:28px;border-radius:4px;border:1.5px solid #ddd;background:#fff;cursor:pointer;font-size:12px;display:inline-flex;align-items:center;justify-content:center;transition:all 0.1s}.absent-btn.marked{background:#c0392b;border-color:#c0392b;color:#fff}.absent-btn:hover{border-color:#c0392b}.row-hover:hover td{background:#fafaf7!important}::-webkit-scrollbar{height:5px;width:5px}::-webkit-scrollbar-thumb{background:#ccc;border-radius:3px}.print-hidden{display:none}`}</style>
      <div ref={printRef} className="print-hidden"><PrintContent /></div>

      {/* Header */}
      <div style={{background:"#1B4332",padding:"16px 24px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
        <div>
          <div style={{color:"#52B788",fontSize:10,letterSpacing:"0.15em",textTransform:"uppercase",marginBottom:3}}>Restaurante Maguje</div>
          <div style={{color:"#fff",fontSize:20,fontFamily:"'Space Grotesk',sans-serif",fontWeight:700}}>Calculadora de ComissÃµes</div>
          <div style={{color:"#95D5B2",fontSize:12,marginTop:2,textTransform:"capitalize"}}>{monthLabel}</div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <input type="month" value={month} onChange={e=>{setMonth(e.target.value);setViewingHistory(null);}}
            style={{...S.input,width:150,background:"#2D6A4F",color:"#fff",border:"1px solid #52B788",colorScheme:"dark"}}/>
          <button style={{...S.btnOut,color:"#95D5B2",borderColor:"#52B788",fontSize:12,padding:"7px 14px"}} onClick={()=>setShowHistory(!showHistory)}>
            ðŸ—‚ HistÃ³rico{history.length>0?` (${history.length})`:""}
          </button>
          <button style={{...S.btnOut,color:"#f28b82",borderColor:"#f28b82",fontSize:12,padding:"7px 14px"}}
            onClick={()=>{sessionStorage.removeItem(SESSION_KEY);window.location.reload();}}>Sair</button>
        </div>
      </div>

      {/* History panel */}
      {showHistory && (
        <div style={{background:"#fff",borderBottom:"1px solid #D4CFC4",padding:"16px 24px"}}>
          <div style={{fontSize:11,color:"#666",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:12}}>HistÃ³rico de ComissÃµes Salvas</div>
          {history.length===0 ? <div style={{fontSize:13,color:"#aaa"}}>Nenhum mÃªs salvo ainda.</div>
            : <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                {history.map(h=>(
                  <div key={h.monthKey} onClick={()=>{setViewingHistory(h.monthKey===viewingHistory?null:h.monthKey);setStep("results");setShowHistory(false);}}
                    style={{border:`1.5px solid ${viewingHistory===h.monthKey?"#1B4332":"#D4CFC4"}`,borderRadius:4,padding:"10px 16px",cursor:"pointer",
                      background:viewingHistory===h.monthKey?"#1B4332":"#fff",color:viewingHistory===h.monthKey?"#fff":"#333",minWidth:140,transition:"all 0.15s"}}>
                    <div style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:13,textTransform:"capitalize"}}>{h.monthLabel}</div>
                    <div style={{fontSize:11,marginTop:3,opacity:0.7}}>Total: {fmt(h.totalBruto)}</div>
                    <div style={{fontSize:10,marginTop:2,opacity:0.5}}>Salvo em {new Date(h.savedAt).toLocaleDateString("pt-BR")}</div>
                  </div>
                ))}
              </div>}
          {viewingHistory && (
            <div style={{marginTop:12,padding:"10px 14px",background:"#fff8e6",border:"1px solid #f0c040",borderRadius:4,fontSize:12,color:"#7a5c00"}}>
              Visualizando: <strong style={{textTransform:"capitalize"}}>{displayMonthLabel}</strong>
              {" Â· "}<span style={{cursor:"pointer",textDecoration:"underline"}} onClick={()=>setViewingHistory(null)}>Voltar ao mÃªs atual</span>
            </div>
          )}
        </div>
      )}

      {/* Step tabs */}
      {!viewingHistory && (
        <div style={{background:"#fff",borderBottom:"1px solid #e0dbd0",padding:"0 24px",display:"flex",gap:0}}>
          {[["revenue","1. Faturamento"],["absences","2. Faltas"],["results","3. Resultado"]].map(([k,l])=>(
            <button key={k} style={S.stepBtn(step===k)} onClick={()=>setStep(k)}>{l}</button>
          ))}
        </div>
      )}

      <div style={{padding:"20px 24px"}}>
        {viewingHistory && (
          <div style={{background:"#fffdf0",border:"1.5px solid #f0c040",borderRadius:4,padding:"10px 16px",marginBottom:16,fontSize:12,color:"#7a5c00"}}>
            ðŸ“‹ Visualizando histÃ³rico de <strong style={{textTransform:"capitalize"}}>{displayMonthLabel}</strong>
            {" Â· "}<span style={{cursor:"pointer",textDecoration:"underline"}} onClick={()=>setViewingHistory(null)}>Voltar ao mÃªs atual</span>
          </div>
        )}

        {/* â”€â”€ STEP 1: FATURAMENTO â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */}
        {!viewingHistory && step==="revenue" && (
          <>
            {/* Global + Adicionais cards */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:16,marginBottom:20}}>
              {/* Global */}
              <div style={{background:"#fff",border:"1.5px solid #40916C",borderRadius:4,padding:"18px 20px"}}>
                <div style={{fontSize:11,color:"#40916C",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:8,fontWeight:600}}>Faturamento Global do MÃªs</div>
                <div style={{fontSize:12,color:"#555",marginBottom:12,lineHeight:1.6}}>Vendas sem comissÃ£o individual. Desconto de 33% aplicado.</div>
                <input type="number" min="0" placeholder="Ex: 53152.69" value={globalMonthlyRevenue}
                  onChange={e=>setGlobalMonthlyRevenue(e.target.value)}
                  style={{...S.input,fontSize:15,padding:"10px 12px",borderColor:"#40916C",fontWeight:500,marginBottom:10}}/>
                {globalMonthlyRevenue && (
                  <div style={{display:"flex",gap:16,flexWrap:"wrap"}}>
                    <div><div style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:14,color:"#1B4332"}}>{fmt(parseFloat(globalMonthlyRevenue)||0)}</div><div style={{fontSize:10,color:"#888",textTransform:"uppercase",marginTop:1}}>Bruto</div></div>
                    <div><div style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:14,color:"#40916C"}}>{fmt((parseFloat(globalMonthlyRevenue)||0)*(1-TAX_RATE))}</div><div style={{fontSize:10,color:"#888",textTransform:"uppercase",marginTop:1}}>LÃ­quido (âˆ’33%)</div></div>
                  </div>
                )}
              </div>
              {/* Adicionais */}
              <div style={{background:"#fff",border:"1.5px solid #7B5EA7",borderRadius:4,padding:"18px 20px"}}>
                <div style={{fontSize:11,color:"#7B5EA7",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:8,fontWeight:600}}>Adicionais do MÃªs</div>
                <div style={{fontSize:12,color:"#555",marginBottom:12,lineHeight:1.6}}>Entra no pool global. <strong>Sem desconto de 33%</strong> â€” valor integral distribuÃ­do.</div>
                <input type="number" min="0" placeholder="Ex: 5000.00" value={adicionais}
                  onChange={e=>setAdicionais(e.target.value)}
                  style={{...S.input,fontSize:15,padding:"10px 12px",borderColor:"#7B5EA7",fontWeight:500,marginBottom:10}}/>
                {adicionais && (
                  <div style={{display:"flex",gap:16}}>
                    <div><div style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:14,color:"#7B5EA7"}}>{fmt(parseFloat(adicionais)||0)}</div><div style={{fontSize:10,color:"#888",textTransform:"uppercase",marginTop:1}}>Integral ao pool</div></div>
                  </div>
                )}
              </div>
            </div>

            {/* Summary row */}
            {(globalMonthlyRevenue || adicionais) && (
              <div style={{background:"#1B433210",border:"1px solid #1B433230",borderRadius:4,padding:"12px 16px",marginBottom:20,display:"flex",gap:24,flexWrap:"wrap"}}>
                <div><div style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:16,color:"#1B4332"}}>{fmt(((parseFloat(globalMonthlyRevenue)||0)*(1-TAX_RATE))+(parseFloat(adicionais)||0))}</div><div style={{fontSize:10,color:"#666",textTransform:"uppercase",marginTop:2}}>Pool Global Total</div></div>
                <div><div style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:14,color:"#555"}}>{workDays.length} dias Ãºteis</div><div style={{fontSize:10,color:"#666",textTransform:"uppercase",marginTop:2}}>DistribuÃ­dos por dia</div></div>
              </div>
            )}

            {/* Individual daily */}
            <div style={{fontSize:11,color:"#666",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:12,fontWeight:600}}>
              Venda Individual DiÃ¡ria â€” GarÃ§ons e Chefes de Fila Junior (29%)
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{borderCollapse:"collapse",width:"100%",minWidth:200+workDays.length*60}}>
                <thead>
                  <tr>
                    <th style={{...S.th,position:"sticky",left:0,zIndex:2,minWidth:200}}>FuncionÃ¡rio</th>
                    {workDays.map(d=>{const dow=new Date(year,mon-1,d).getDay();return <th key={d} style={{...S.th,textAlign:"center",minWidth:60}}><div>{d}</div><div style={{fontWeight:300,fontSize:9,opacity:0.7}}>{DOW_LABELS[dow]}</div></th>;})}
                    <th style={{...S.th,textAlign:"right",minWidth:110}}>Total Bruto</th>
                  </tr>
                </thead>
                <tbody>
                  {indivEmps.map((emp,idx)=>{
                    const total=workDays.reduce((s,d)=>s+(parseFloat(getDR(d).individual[emp.id])||0),0);
                    return (
                      <tr key={emp.id} className="row-hover">
                        <td style={{...S.td,position:"sticky",left:0,background:idx%2===0?"#fff":"#fafaf8",zIndex:1,fontWeight:500}}>
                          <div style={{display:"flex",alignItems:"center",gap:6}}><span style={{fontSize:12}}>{emp.name}</span><MeiTag active={emp.mei}/></div>
                          <div style={{fontSize:10,color:SECTOR_COLORS[emp.sector]||"#888",marginTop:1}}>{emp.role}</div>
                        </td>
                        {workDays.map(d=>(
                          <td key={d} style={{...S.td,padding:"5px 6px",background:idx%2===0?"#fff":"#fafaf8"}}>
                            <input type="number" min="0" placeholder="0" value={getDR(d).individual[emp.id]||""}
                              onChange={e=>setIR(d,emp.id,e.target.value)}
                              style={{...S.input,width:52,textAlign:"right",padding:"4px 5px"}}/>
                          </td>
                        ))}
                        <td style={{...S.td,textAlign:"right",fontFamily:"'Space Grotesk',sans-serif",fontWeight:600,color:"#1B4332",background:idx%2===0?"#fff":"#fafaf8"}}>{fmt(total)}</td>
                      </tr>
                    );
                  })}
                  <tr style={{background:"#F5F0E8"}}>
                    <td style={{...S.td,position:"sticky",left:0,background:"#F5F0E8",zIndex:1,fontWeight:600,fontSize:11,textTransform:"uppercase",letterSpacing:"0.05em"}}>Total Individual do Dia</td>
                    {workDays.map(d=>{const tot=indivEmps.reduce((s,e)=>s+(parseFloat(getDR(d).individual[e.id])||0),0);return <td key={d} style={{...S.td,textAlign:"right",fontWeight:600,fontSize:11,background:"#F5F0E8",padding:"7px 5px",color:"#1B4332"}}>{tot>0?fmtShort(tot):"â€”"}</td>;})}
                    <td style={{...S.td,textAlign:"right",fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,color:"#1B4332",background:"#F5F0E8"}}>{fmt(workDays.reduce((s,d)=>s+indivEmps.reduce((ss,e)=>ss+(parseFloat(getDR(d).individual[e.id])||0),0),0))}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div style={{marginTop:20,textAlign:"right"}}>
              <button style={S.btn} onClick={()=>setStep("absences")}>PrÃ³ximo: Faltas â†’</button>
            </div>
          </>
        )}

        {/* â”€â”€ STEP 2: FALTAS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */}
        {!viewingHistory && step==="absences" && (
          <>
            {/* Attendance import box */}
            <div style={{background:"#fff",border:"1.5px solid #D4CFC4",borderRadius:4,padding:"18px 20px",marginBottom:16}}>
              <div style={{fontSize:11,color:"#666",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:10,fontWeight:600}}>ðŸ“Ž Importar Planilha de Ponto</div>
              <div style={{fontSize:12,color:"#555",marginBottom:14,lineHeight:1.6}}>
                Importe o arquivo exportado do sistema de ponto (.xlsx). Qualquer marcaÃ§Ã£o na coluna <strong>TIPO</strong> serÃ¡ registrada como falta.
                ApÃ³s importar, vocÃª pode editar manualmente as faltas na tabela abaixo.
              </div>
              <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
                <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{display:"none"}}
                  onChange={e=>{ if(e.target.files[0]) handleAttendanceImport(e.target.files[0]); e.target.value=""; }}/>
                <button style={{...S.btn,background:"#40916C",borderColor:"#40916C"}} onClick={()=>fileInputRef.current.click()}>
                  â¬† Carregar arquivo .xlsx
                </button>
                {attendanceFile && (
                  <span style={{fontSize:12,color:"#2D6A4F",fontWeight:500}}>âœ“ {attendanceFile.name}</span>
                )}
                {attendanceFile && (
                  <button style={S.btnSm("#c0392b")} onClick={()=>{setAttendanceFile(null);setImportLog(null);}}>Remover</button>
                )}
              </div>

              {/* Import log */}
              {importLog && (
                <div style={{marginTop:14,padding:"12px 14px",background:"#f0faf4",border:"1px solid #52B78860",borderRadius:4}}>
                  <div style={{fontSize:12,color:"#1B4332",fontWeight:600,marginBottom:6}}>
                    âœ“ ImportaÃ§Ã£o concluÃ­da â€” {importLog.matched} registros aplicados
                  </div>
                  {importLog.unmatched.length > 0 && (
                    <>
                      <div style={{fontSize:11,color:"#856404",marginBottom:4,fontWeight:500}}>
                        âš  {importLog.unmatched.length} nome(s) nÃ£o encontrado(s) na lista de funcionÃ¡rios:
                      </div>
                      <div style={{fontSize:11,color:"#666",lineHeight:1.8}}>
                        {importLog.unmatched.map(n=><div key={n} style={{padding:"1px 0"}}>â€¢ {n}</div>)}
                      </div>
                    </>
                  )}
                </div>
              )}
            </div>

            {/* Sector filter + add employee */}
            <div style={{display:"flex",gap:6,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
              {SECTORS.map(s2=><button key={s2} style={sector===s2?S.tabA:S.tab} onClick={()=>setSector(s2)}>{s2}</button>)}
              <div style={{marginLeft:"auto",display:"flex",gap:8}}>
                <button style={{...S.btnOut,fontSize:12,padding:"6px 14px"}} onClick={()=>{setShowAdd(!showAdd);setEditingId(null);}}>
                  {showAdd?"âœ• Cancelar":"+ FuncionÃ¡rio"}
                </button>
              </div>
            </div>

            {showAdd && !editingId && (
              <EmployeeForm initial={{name:"",role:"",sector:"SalÃ£o",type:"individual",points:15,mei:false}} onSave={handleAddEmployee} onCancel={()=>setShowAdd(false)}/>
            )}

            <div style={{overflowX:"auto"}}>
              <table style={{borderCollapse:"collapse",width:"100%",minWidth:260+workDays.length*42}}>
                <thead>
                  <tr>
                    <th style={{...S.th,position:"sticky",left:0,zIndex:2,minWidth:200}}>FuncionÃ¡rio</th>
                    <th style={{...S.th,textAlign:"center",minWidth:55}}>MEI</th>
                    {workDays.map(d=>{const dow=new Date(year,mon-1,d).getDay();return <th key={d} style={{...S.th,textAlign:"center",minWidth:42,padding:"8px 4px"}}><div>{d}</div><div style={{fontWeight:300,fontSize:9,opacity:0.7}}>{DOW_LABELS[dow]}</div></th>;})}
                    <th style={{...S.th,textAlign:"center",minWidth:60}}>Faltas</th>
                    <th style={{...S.th,textAlign:"center",minWidth:70}}>AÃ§Ãµes</th>
                  </tr>
                </thead>
                <tbody>
                  {[
                    {label:"GarÃ§ons e Chefes de Fila Junior",emps:filteredEmps.filter(e=>e.type==="individual")},
                    {label:"Equipe Global",emps:filteredEmps.filter(e=>e.type==="global")},
                  ].map(group=>group.emps.length===0?null:(
                    <>{<tr key={group.label}><td colSpan={workDays.length+5} style={{...S.td,background:"#1B433210",fontWeight:600,fontSize:11,textTransform:"uppercase",letterSpacing:"0.06em",color:"#1B4332",padding:"10px 12px"}}>{group.label}</td></tr>}
                    {group.emps.map((emp,idx)=>{
                      const ac=absCount(emp.id);
                      if (editingId===emp.id) return (
                        <tr key={emp.id}><td colSpan={workDays.length+5} style={{padding:"8px 12px",background:"#f9fff9"}}>
                          <EmployeeForm initial={emp} onSave={handleEditEmployee} onCancel={()=>setEditingId(null)}/>
                        </td></tr>
                      );
                      return (
                        <tr key={emp.id} className="row-hover">
                          <td style={{...S.td,position:"sticky",left:0,background:idx%2===0?"#fff":"#fafaf8",zIndex:1}}>
                            <div style={{fontSize:12,fontWeight:500,display:"flex",alignItems:"center",gap:5}}>{emp.name}<MeiTag active={emp.mei}/></div>
                            <div style={{fontSize:10,color:SECTOR_COLORS[emp.sector]||"#888",marginTop:1}}>{emp.role} Â· {emp.sector}{emp.type==="global"?` Â· ${emp.points}pts`:""}</div>
                          </td>
                          <td style={{...S.td,textAlign:"center",background:idx%2===0?"#fff":"#fafaf8"}}>
                            <div onClick={()=>toggleMei(emp.id)} style={{width:34,height:18,borderRadius:9,background:emp.mei?"#2D6A4F":"#ddd",position:"relative",transition:"background 0.2s",cursor:"pointer",margin:"0 auto"}}>
                              <div style={{position:"absolute",top:2,left:emp.mei?16:2,width:14,height:14,borderRadius:"50%",background:"#fff",transition:"left 0.2s"}}/>
                            </div>
                          </td>
                          {workDays.map(d=>{
                            const absent=isAbsent(emp.id,d);
                            return <td key={d} style={{...S.td,textAlign:"center",padding:"5px 4px",background:absent?"#fdecea":idx%2===0?"#fff":"#fafaf8"}}>
                              <button className={`absent-btn${absent?" marked":""}`} onClick={()=>toggleAbs(emp.id,d)}>{absent?"F":"Â·"}</button>
                            </td>;
                          })}
                          <td style={{...S.td,textAlign:"center",background:idx%2===0?"#fff":"#fafaf8"}}>
                            {ac>0?<span style={{background:"#fdecea",color:"#c0392b",borderRadius:10,padding:"2px 8px",fontSize:12,fontWeight:600}}>{ac}</span>:<span style={{color:"#ccc",fontSize:12}}>â€”</span>}
                          </td>
                          <td style={{...S.td,textAlign:"center",background:idx%2===0?"#fff":"#fafaf8"}}>
                            <button style={S.btnSm("#1B4332")} onClick={()=>{setEditingId(emp.id);setShowAdd(false);}}>âœ Editar</button>
                          </td>
                        </tr>
                      );
                    })}</>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{marginTop:20,display:"flex",justifyContent:"space-between"}}>
              <button style={S.btnOut} onClick={()=>setStep("revenue")}>â† Voltar</button>
              <button style={S.btn} onClick={()=>setStep("results")}>Calcular ComissÃµes â†’</button>
            </div>
          </>
        )}

        {/* â”€â”€ STEP 3 + HISTORY: RESULTADO â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */}
        {((!viewingHistory && step==="results") || viewingHistory) && (
          <>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:12,marginBottom:20}}>
              {[
                {label:"Total Bruto",val:fmt(displayResults.totalBruto),color:"#1B4332"},
                {label:"Desconto 33%",val:fmt(displayResults.totalBruto*TAX_RATE),color:"#c0392b"},
                {label:"Pool Global LÃ­q.",val:fmt(displayResults.totalGlobalPool),color:"#40916C"},
                {label:"ComissÃµes Indiv.",val:fmt(displayResults.totalIndivComm),color:"#7B5EA7"},
              ].map(m=>(
                <div key={m.label} style={{background:"#fff",border:"1.5px solid #D4CFC4",borderRadius:4,padding:"13px 16px",textAlign:"center"}}>
                  <div style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:17,color:m.color}}>{m.val}</div>
                  <div style={{fontSize:10,color:"#888",textTransform:"uppercase",letterSpacing:"0.07em",marginTop:3}}>{m.label}</div>
                </div>
              ))}
            </div>
            <div style={{display:"flex",gap:6,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
              {SECTORS.map(s2=><button key={s2} style={sector===s2?S.tabA:S.tab} onClick={()=>setSector(s2)}>{s2}</button>)}
              <div style={{marginLeft:"auto",display:"flex",gap:8}}>
                {!viewingHistory&&<button style={S.btnGreen} onClick={handleSaveHistory}>ðŸ’¾ Salvar MÃªs</button>}
                <button style={S.btn} onClick={handlePrint}>ðŸ–¨ Exportar PDF</button>
              </div>
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{borderCollapse:"collapse",width:"100%"}}>
                <thead>
                  <tr>
                    <th style={{...S.th,minWidth:220}}>FuncionÃ¡rio</th>
                    <th style={S.th}>Cargo</th>
                    <th style={{...S.th,textAlign:"center"}}>Tipo</th>
                    <th style={{...S.th,textAlign:"center"}}>MEI</th>
                    <th style={{...S.th,textAlign:"center"}}>Faltas</th>
                    <th style={{...S.th,textAlign:"right"}}>ComissÃ£o</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredEmps.slice().sort((a,b)=>(displayResults.empTotals[b.id]||0)-(displayResults.empTotals[a.id]||0)).map((emp,idx)=>{
                    const comm=displayResults.empTotals[emp.id]||0;
                    const ac=absCount(emp.id);
                    const color=SECTOR_COLORS[emp.sector]||"#555";
                    return (
                      <tr key={emp.id} className="row-hover">
                        <td style={{...S.td,background:idx%2===0?"#fff":"#fafaf8",fontWeight:500}}>
                          <div style={{fontSize:13,display:"flex",alignItems:"center",gap:6}}>{emp.name}<MeiTag active={emp.mei}/></div>
                          <span style={{display:"inline-block",fontSize:10,padding:"1px 7px",borderRadius:20,marginTop:2,background:color+"18",color,border:`1px solid ${color}40`,textTransform:"uppercase",letterSpacing:"0.05em"}}>{emp.sector}</span>
                        </td>
                        <td style={{...S.td,fontSize:12,color:"#444",background:idx%2===0?"#fff":"#fafaf8"}}>{emp.role}</td>
                        <td style={{...S.td,textAlign:"center",background:idx%2===0?"#fff":"#fafaf8"}}>
                          {emp.type==="individual"?<span style={{fontSize:11,padding:"2px 8px",borderRadius:20,background:"#7B5EA720",color:"#7B5EA7",border:"1px solid #7B5EA740"}}>Individual</span>:<span style={{fontSize:11,padding:"2px 8px",borderRadius:20,background:"#2D6A4F20",color:"#2D6A4F",border:"1px solid #2D6A4F40"}}>Global Â· {emp.points}pts</span>}
                        </td>
                        <td style={{...S.td,textAlign:"center",background:idx%2===0?"#fff":"#fafaf8"}}>
                          {emp.mei?<span style={{fontSize:10,padding:"2px 8px",borderRadius:20,background:"#fff3cd",color:"#856404",border:"1px solid #ffc107",fontWeight:600}}>MEI</span>:<span style={{color:"#ccc",fontSize:11}}>â€”</span>}
                        </td>
                        <td style={{...S.td,textAlign:"center",background:idx%2===0?"#fff":"#fafaf8"}}>
                          {ac>0?<span style={{background:"#fdecea",color:"#c0392b",borderRadius:10,padding:"2px 8px",fontSize:12,fontWeight:600}}>{ac} dia{ac>1?"s":""}</span>:<span style={{color:"#ccc",fontSize:12}}>â€”</span>}
                        </td>
                        <td style={{...S.td,textAlign:"right",background:idx%2===0?"#fff":"#fafaf8"}}>
                          <span style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:15,color:comm>0?"#1B4332":"#bbb"}}>{fmt(comm)}</span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                <tfoot>
                  <tr style={{background:"#f0f5f0",borderTop:"2px solid #1B4332"}}>
                    <td colSpan={5} style={{...S.td,fontWeight:600,color:"#1B4332",fontSize:13}}>Total {sector!=="Todos"?`â€” ${sector}`:""}</td>
                    <td style={{...S.td,textAlign:"right"}}>
                      <span style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:17,color:"#1B4332"}}>{fmt(filteredEmps.reduce((s,e)=>s+(displayResults.empTotals[e.id]||0),0))}</span>
                    </td>
                  </tr>
                </tfoot>
              </table>
            </div>
            {!viewingHistory&&<div style={{marginTop:20}}><button style={S.btnOut} onClick={()=>setStep("absences")}>â† Voltar</button></div>}
          </>
        )}
      </div>

      <div style={{textAlign:"center",padding:"14px",fontSize:11,color:"#aaa",borderTop:"1px solid #e5e0d6",marginTop:20}}>
        Maguje Â· ComissÃµes Â· {monthLabel}
      </div>
    </div>
  );
}
