import React, { useState, useMemo, useRef, useEffect } from "react";

const TAX_RATE = 0.33;
const INDIVIDUAL_RATE = 0.29;
const HISTORY_KEY = "maguje_history";
const APP_PASSWORD = "maguje2026";
const SESSION_KEY = "maguje_auth";

// ── Persistência por mês ──────────────────────────────────────
function loadMonthData(month) {
  try {
    const rev = JSON.parse(localStorage.getItem(`maguje_revenue_v2_${month}`) || "{}");
    const abs = JSON.parse(localStorage.getItem(`maguje_absences_v2_${month}`) || "{}");
    return { rev, abs };
  } catch { return { rev: {}, abs: {} }; }
}
function saveRevenueForMonth(month, revenue) {
  localStorage.setItem(`maguje_revenue_v2_${month}`, JSON.stringify(revenue));
}
function saveAbsencesForMonth(month, absences) {
  localStorage.setItem(`maguje_absences_v2_${month}`, JSON.stringify(absences));
}
function loadEmployees() {
  try { const s = localStorage.getItem("maguje_employees_v2"); return s ? JSON.parse(s) : null; } catch { return null; }
}
function saveEmployees(emps) { localStorage.setItem("maguje_employees_v2", JSON.stringify(emps)); }
function loadHistory() { try { return JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]"); } catch { return []; } }
function saveHistory(records) { localStorage.setItem(HISTORY_KEY, JSON.stringify(records)); }

// ── Fuzzy match com prefixo ───────────────────────────────────
function fuzzyMatchEmployee(rawName, employees) {
  const normalize = s =>
    s.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "")
     .replace(/[0-9]/g, " ").replace(/[^a-z\s]/g, " ").trim();
  const exWords = normalize(rawName).split(/\s+/).filter(w => w.length > 2);
  if (!exWords.length) return null;
  let best = null, bestScore = 0;
  for (const emp of employees) {
    const appWords = normalize(emp.name).split(/\s+/).filter(w => w.length > 2);
    const overlap = exWords.filter(ex =>
      appWords.some(app => app === ex || app.startsWith(ex) || ex.startsWith(app))
    ).length;
    if (overlap < 2) continue;
    const score = overlap / Math.max(exWords.length, appWords.length, 1);
    if (score > bestScore) { bestScore = score; best = emp; }
  }
  return best && bestScore > 0.2 ? best : null;
}

function toTitleCase(str) {
  return str.toLowerCase().replace(/(?:^|\s)\S/g, a => a.toUpperCase());
}

// ── Ler cores da coluna A via JSZip ──────────────────────────
async function getColumnAColors(arrayBuffer) {
  if (!window.JSZip) return {};
  try {
    const zip = await window.JSZip.loadAsync(arrayBuffer);
    const stylesXml = await zip.file("xl/styles.xml").async("string");
    const fills = [];
    const fillRx = /<fill>([\s\S]*?)<\/fill>/g;
    let fm;
    while ((fm = fillRx.exec(stylesXml)) !== null) {
      const inner = fm[1];
      const ptm = inner.match(/patternType="([^"]+)"/);
      const pt = ptm ? ptm[1] : "";
      if (!pt || pt === "none" || pt === "gray125") { fills.push(null); continue; }
      const rgbm = inner.match(/<fgColor[^>]*rgb="([^"]+)"/);
      fills.push(rgbm ? rgbm[1].toUpperCase() : null);
    }
    const xfToColor = [];
    const cxm = stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/);
    if (cxm) {
      const xfRx = /<xf [^>]*fillId="(\d+)"[^/]*/g;
      let xm;
      while ((xm = xfRx.exec(cxm[1])) !== null) {
        const fid = parseInt(xm[1]);
        xfToColor.push(fills[fid] !== undefined ? fills[fid] : null);
      }
    }
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml").async("string");
    const rowColors = {};
    const cRx = /<c r="A(\d+)"([^>]*?)(?:\/>|>)/g;
    let cm;
    while ((cm = cRx.exec(sheetXml)) !== null) {
      const rowNum = parseInt(cm[1]);
      const sm = cm[2].match(/\bs="(\d+)"/);
      const styleIdx = sm ? parseInt(sm[1]) : 0;
      rowColors[rowNum] = xfToColor[styleIdx] || null;
    }
    return rowColors;
  } catch (e) { console.warn("getColumnAColors:", e); return {}; }
}

// ── Importar planilha de faturamento ──────────────────────────
function parseExcelImport(file, employees, month, onComplete, onError) {
  const XLSX = window.XLSX;
  if (!XLSX) { onError("Biblioteca de Excel não carregada. Aguarde e tente novamente."); return; }
  const reader = new FileReader();
  reader.onload = async function(e) {
    try {
      const arrayBuffer = e.target.result;
      const colorMap = await getColumnAColors(arrayBuffer);
      const data = new Uint8Array(arrayBuffer);
      const wb = XLSX.read(data, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const range = XLSX.utils.decode_range(ws["!ref"]);
      const getCell = (r, c) => ws[XLSX.utils.encode_cell({ r, c })];
      const getVal  = (r, c) => { const cell = getCell(r, c); return cell ? cell.v : null; };
      const dailyData = {};
      const newEmps = [];
      const log = [];
      let curEmp = null, curIsIndividual = false, fileMonth = null;
      for (let r = range.s.r; r <= range.e.r; r++) {
        const v0 = getVal(r, 0);
        if (v0 && String(v0).includes("Atendente:")) {
          const rawName = String(v0).replace("Atendente:", "")
            .replace(/\s*-\s*(CONTRATO|GARCOM|EXTRA|GAVETA).*/i, "").trim();
          const matched = fuzzyMatchEmployee(rawName, employees);
          if (matched) {
            curEmp = matched; curIsIndividual = matched.type === "individual";
            log.push({ name: rawName, matched: matched.name, type: matched.type });
          } else {
            const excelRow = r + 1;
            const color = colorMap[excelRow];
            const isGray = color && color.toUpperCase() === "FFD3D3D3";
            curIsIndividual = !isGray;
            if (curIsIndividual) {
              const newId = "imp_" + Date.now() + "_" + r;
              const newEmp = { id: newId, name: toTitleCase(rawName), role: "Garçom", sector: "Salão", type: "individual", points: 0 };
              newEmps.push(newEmp);
              curEmp = newEmp;
              log.push({ name: rawName, isNew: true, type: "individual" });
            } else {
              curEmp = { type: "global_pool" }; curIsIndividual = false;
              log.push({ name: rawName, type: "global" });
            }
          }
        } else if (!v0) {
          const v1 = getVal(r, 1), v2 = getVal(r, 2), v8 = getVal(r, 8);
          if (v1 !== null && v2 !== null && v8 !== null && !isNaN(parseFloat(v8))) {
            let day, rowMonth;
            if (v2 instanceof Date) {
              day = v2.getDate();
              rowMonth = v2.getFullYear() + "-" + String(v2.getMonth() + 1).padStart(2, "0");
            } else if (typeof v2 === "number") {
              try { const d = XLSX.SSF.parse_date_code(v2); day = d.d; rowMonth = d.y + "-" + String(d.m).padStart(2, "0"); }
              catch { continue; }
            } else { continue; }
            if (!fileMonth && rowMonth) fileMonth = rowMonth;
            const total = parseFloat(v8) || 0;
            if (total <= 0 || !day) continue;
            if (!dailyData[day]) dailyData[day] = { individual: {}, globalSum: 0 };
            if (curIsIndividual && curEmp && curEmp.id)
              dailyData[day].individual[curEmp.id] = (dailyData[day].individual[curEmp.id] || 0) + total;
            else if (curEmp) dailyData[day].globalSum += total;
          }
        }
      }
      onComplete({ dailyData, newEmps, log, fileMonth });
    } catch (err) { onError("Erro ao processar planilha: " + err.message); }
  };
  reader.readAsArrayBuffer(file);
}

// ── Importar planilha de ponto ────────────────────────────────
// Regras de classificação pelo MOTIVO:
//   ESQUECIMENTO* → E (50% da comissão)
//   ATESTADO*     → A (sem comissão, igual a falta)
//   FALTA         → F (sem comissão)
//   Tudo mais (FOLGA, DOMINGO, CERTIDÃO, LICENÇA, BANCO...)  → ignorar (100%)
function parsePontoImport(file, employees, month, onComplete, onError) {
  const XLSX = window.XLSX;
  if (!XLSX) { onError("Biblioteca não carregada. Aguarde e tente novamente."); return; }
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const absToAdd = {};
      const log = [];
      let fileMonth = null;

      for (const sheetName of wb.SheetNames) {
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || !row[0]) continue;
          const name   = String(row[0]).trim();
          const dayStr = row[3] != null ? String(row[3]).trim() : "";
          const motivo = row[5] != null ? String(row[5]).trim().toUpperCase().replace(/\s+/g, " ") : "";
          if (!name || !dayStr || !motivo) continue;

          // Dia: "02/04 QUI" → day=2, month=4
          const dayMatch = dayStr.match(/^(\d{1,2})\/(\d{1,2})/);
          if (!dayMatch) continue;
          const dayNum   = parseInt(dayMatch[1]);
          const monthNum = parseInt(dayMatch[2]);

          if (!fileMonth) {
            const yr = parseInt((month || "").split("-")[0]) || new Date().getFullYear();
            fileMonth = yr + "-" + String(monthNum).padStart(2, "0");
          }

          // Classificação
          let status = null;
          if (motivo.includes("ESQUECIMENTO")) status = "E";
          else if (motivo.includes("ATESTADO"))   status = "A";
          else if (motivo.trim() === "FALTA")      status = "F";
          // FOLGA, DOMINGO DO MÊS, CERTIDÃO, LICENÇA, BANCO DE HORAS, HOME OFFICE → null (ignora)

          if (!status) continue;

          const emp = fuzzyMatchEmployee(name, employees);
          if (!emp) { log.push({ name, status, matched: false }); continue; }

          if (!absToAdd[emp.id]) absToAdd[emp.id] = {};
          if (!absToAdd[emp.id][dayNum]) {
            absToAdd[emp.id][dayNum] = status;
            log.push({ name, matched: emp.name, day: dayNum, status });
          }
        }
      }
      onComplete({ absToAdd, log, fileMonth });
    } catch (err) { onError("Erro ao processar ponto: " + err.message); }
  };
  reader.readAsArrayBuffer(file);
}

// ── Login ─────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [pwd, setPwd] = useState("");
  const [error, setError] = useState(false);
  const [shake, setShake] = useState(false);
  const handleSubmit = e => {
    e.preventDefault();
    if (pwd === APP_PASSWORD) { sessionStorage.setItem(SESSION_KEY, "1"); onLogin(); }
    else { setError(true); setShake(true); setPwd(""); setTimeout(() => setShake(false), 500); }
  };
  return (
    <div style={{ fontFamily: "'DM Mono','Courier New',monospace", background: "#F5F0E8", minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Space+Grotesk:wght@500;700&display=swap');
        * { box-sizing: border-box; }
        @keyframes shake { 0%,100%{transform:translateX(0)} 20%,60%{transform:translateX(-8px)} 40%,80%{transform:translateX(8px)} }
        .shake { animation: shake 0.4s ease; }
        input:focus { outline: none; border-color: #2D6A4F !important; }
      `}</style>
      <div style={{ width: "100%", maxWidth: 380, padding: "0 20px" }}>
        <div style={{ textAlign: "center", marginBottom: 36 }}>
          <div style={{ width: 56, height: 56, background: "#1B4332", borderRadius: 12, margin: "0 auto 16px", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <span style={{ fontSize: 26 }}>🌿</span>
          </div>
          <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 22, color: "#1B4332" }}>Maguje</div>
          <div style={{ fontSize: 12, color: "#888", marginTop: 4 }}>Sistema de Comissões</div>
        </div>
        <div className={shake ? "shake" : ""} style={{ background: "#fff", border: "1.5px solid #D4CFC4", borderRadius: 6, padding: "28px 28px 24px" }}>
          <div style={{ fontSize: 13, fontWeight: 500, color: "#333", marginBottom: 20 }}>Acesso restrito</div>
          <form onSubmit={handleSubmit}>
            <div style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 11, color: "#666", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Senha</div>
              <input type="password" placeholder="••••••••" value={pwd} onChange={e => { setPwd(e.target.value); setError(false); }} autoFocus
                style={{ background: "#F5F0E8", border: `1.5px solid ${error ? "#c0392b" : "#ccc"}`, borderRadius: 3, padding: "10px 12px", fontFamily: "inherit", fontSize: 14, width: "100%", letterSpacing: "0.1em" }} />
              {error && <div style={{ fontSize: 11, color: "#c0392b", marginTop: 6 }}>Senha incorreta. Tente novamente.</div>}
            </div>
            <button type="submit" style={{ width: "100%", border: "none", background: "#1B4332", color: "#fff", padding: "11px", borderRadius: 3, cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 500 }}>Entrar →</button>
          </form>
        </div>
      </div>
    </div>
  );
}

function getWorkingDays(year, month) {
  const days = [];
  const total = new Date(year, month, 0).getDate();
  for (let d = 1; d <= total; d++) {
    if (new Date(year, month - 1, d).getDay() !== 1) days.push(d);
  }
  return days;
}

const DOW_LABELS = ["Dom", "", "Ter", "Qua", "Qui", "Sex", "Sáb"];
const fmt = v => (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const fmtShort = v => (v || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const INITIAL_EMPLOYEES = [
  { id: 1,  name: "Douglas Pereira Lima",              role: "Garçom",              sector: "Salão",   type: "individual", points: 0 },
  { id: 2,  name: "Gabriel de Farias Pereira",         role: "Garçom",              sector: "Salão",   type: "individual", points: 0 },
  { id: 3,  name: "Gustavo Fabricio Rodrigues Freire", role: "Garçom",              sector: "Salão",   type: "individual", points: 0 },
  { id: 4,  name: "Antonia Erineuda",                  role: "Garçonete",           sector: "Salão",   type: "individual", points: 0 },
  { id: 5,  name: "Paulo Alves de Almeida",            role: "Garçom",              sector: "Salão",   type: "individual", points: 0 },
  { id: 6,  name: "Reinaldo Alves de Oliveira",        role: "Chefe de Fila Junior",sector: "Salão",   type: "individual", points: 0 },
  { id: 7,  name: "Jeane Rodrigues",                   role: "Chefe de Fila Junior",sector: "Salão",   type: "individual", points: 0 },
  { id: 8,  name: "Verinaldo Gabriel da Rocha",        role: "Chefe de Fila Junior",sector: "Salão",   type: "individual", points: 0 },
  { id: 9,  name: "Jean Carlos Fidelis",               role: "Cumim",               sector: "Salão",   type: "global", points: 15 },
  { id: 10, name: "Claudia Elisabete Conceição",       role: "Cumim",               sector: "Salão",   type: "global", points: 15 },
  { id: 11, name: "Marcos Vinicius Henrique de Souza", role: "Cumim",               sector: "Salão",   type: "global", points: 15 },
  { id: 12, name: "Maria Elenice Ferreira",            role: "Cumim",               sector: "Salão",   type: "global", points: 15 },
  { id: 13, name: "Rodrigo de Pinho Ribeiro",          role: "Cumim",               sector: "Salão",   type: "global", points: 15 },
  { id: 14, name: "Felipe Costa de Abreu",             role: "Suiteiro",            sector: "Salão",   type: "global", points: 20 },
  { id: 15, name: "Crislandia Moura de Lima",          role: "Chefe de Fila Pleno", sector: "Salão",   type: "global", points: 26 },
  { id: 16, name: "Elizangelo Araujo Miranda",         role: "Maître",              sector: "Salão",   type: "global", points: 30 },
  { id: 17, name: "Joaquim Fernandes Gomes",           role: "Assistente Gerente",  sector: "Salão",   type: "global", points: 25 },
  { id: 18, name: "Jose Edilson Pereira Nogueira",     role: "Sub Gerente / Maître",sector: "Salão",   type: "global", points: 30 },
  { id: 19, name: "Rodrigo Florentino Fonseca",        role: "Gerente",             sector: "Salão",   type: "global", points: 35 },
  { id: 20, name: "Fabio da Silva Miguel",             role: "Assistente MKT",      sector: "Salão",   type: "global", points: 15 },
  { id: 21, name: "Kayllana Vitoria de Oliveira",      role: "Hostess",             sector: "Salão",   type: "global", points: 15 },
  { id: 22, name: "Suzana Radai Estrela Souza",        role: "Hostess",             sector: "Salão",   type: "global", points: 15 },
  { id: 23, name: "Romenia Fernades Jorge",            role: "Hostess",             sector: "Salão",   type: "global", points: 20 },
  { id: 24, name: "Danilo Silva Gomes",                role: "Barback",             sector: "Bar",     type: "global", points: 17 },
  { id: 25, name: "Luan Chrystyan dos Santos",         role: "Barback",             sector: "Bar",     type: "global", points: 17 },
  { id: 26, name: "Francisco Tome da Silva",           role: "Copeiro II",          sector: "Bar",     type: "global", points: 10 },
  { id: 27, name: "Antonio Mauricio Santos Soares",    role: "Bartender",           sector: "Bar",     type: "global", points: 20 },
  { id: 28, name: "Rafael da Silva Romualdo",          role: "Bartender",           sector: "Bar",     type: "global", points: 20 },
  { id: 29, name: "Caio Henriques Rodrigues",          role: "Bartender",           sector: "Bar",     type: "global", points: 20 },
  { id: 30, name: "Gabriel Paulino Barbosa",           role: "Bartender",           sector: "Bar",     type: "global", points: 20 },
  { id: 31, name: "Gabriel Soares Grativol",           role: "Bartender",           sector: "Bar",     type: "global", points: 20 },
  { id: 32, name: "Gabriel de Oliveira Fernandes",     role: "Sub Chefe de Bar",    sector: "Bar",     type: "global", points: 22 },
  { id: 33, name: "Luiz Gustavo Mesquita Soares",      role: "Chefe de Bar",        sector: "Bar",     type: "global", points: 25 },
  { id: 34, name: "Antonia Jacilane de Sousa Costa",   role: "Caixa",               sector: "Caixa",   type: "global", points: 15 },
  { id: 35, name: "Antonio Gomes de Sousa",            role: "Copeiro",             sector: "Cozinha", type: "global", points: 15 },
  { id: 36, name: "Douglas Leite Gonçalves",           role: "Copeiro",             sector: "Cozinha", type: "global", points: 15 },
  { id: 37, name: "John Victor Santos do Nascimento",  role: "Copeiro",             sector: "Cozinha", type: "global", points: 15 },
  { id: 38, name: "Rosangela Costa Rodrigues",         role: "Copeiro",             sector: "Cozinha", type: "global", points: 15 },
  { id: 39, name: "Robert Gustavo Santos de Souza",    role: "Copeiro",             sector: "Cozinha", type: "global", points: 15 },
  { id: 40, name: "Daniel Pereira do Sacramento",      role: "Padeiro",             sector: "Cozinha", type: "global", points: 15 },
  { id: 41, name: "Rosinaldo Pedro Soares",            role: "Ajudante de Cozinha", sector: "Cozinha", type: "global", points: 15 },
  { id: 42, name: "Thaynara Tonelle Costa",            role: "Cozinheiro I",        sector: "Cozinha", type: "global", points: 15 },
  { id: 43, name: "Andre Felizardo Verissimo",         role: "Cozinheiro I",        sector: "Cozinha", type: "global", points: 15 },
  { id: 44, name: "Dayveson Rafael da Silva",          role: "Cozinheiro I",        sector: "Cozinha", type: "global", points: 15 },
  { id: 45, name: "Andriely Firmino da Silva",         role: "Cozinheiro I",        sector: "Cozinha", type: "global", points: 15 },
  { id: 46, name: "Vitor Faria de Oliveira Aguilera",  role: "Cozinheiro I",        sector: "Cozinha", type: "global", points: 15 },
  { id: 47, name: "Lucas Barbosa Ribeiro Borges",      role: "Cozinheiro II",       sector: "Cozinha", type: "global", points: 19 },
  { id: 48, name: "Robson Roberto da Silva",           role: "Cozinheiro II",       sector: "Cozinha", type: "global", points: 19 },
  { id: 49, name: "Wagner Pinto",                      role: "Cozinheiro III",      sector: "Cozinha", type: "global", points: 21 },
  { id: 50, name: "Francisco Dalvan Bezerra Gomes",    role: "Cozinheiro III",      sector: "Cozinha", type: "global", points: 21 },
  { id: 51, name: "Valdemir Galdino de Oliveira",      role: "Cozinheiro III",      sector: "Cozinha", type: "global", points: 21 },
  { id: 52, name: "Luis Augusto Souza da Costa",       role: "Cozinheiro Líder",    sector: "Cozinha", type: "global", points: 23 },
  { id: 53, name: "Jaqueline de Souza Galvao",         role: "Sub Chefe Cozinha",   sector: "Cozinha", type: "global", points: 25 },
  { id: 54, name: "Eduardo",                           role: "Chef Produções Gast.",sector: "Cozinha", type: "global", points: 30 },
  { id: 55, name: "Alex dos Santos",                   role: "ASG",                 sector: "Limpeza", type: "global", points: 10 },
  { id: 56, name: "Carlos Daniel Alves de Lima",       role: "ASG",                 sector: "Limpeza", type: "global", points: 10 },
  { id: 57, name: "Marlucia Santana Rodrigues",        role: "Líder de ASG",        sector: "Limpeza", type: "global", points: 20 },
  { id: 58, name: "José Roberto Inácio da Silva",      role: "Estoquista",          sector: "Limpeza", type: "global", points: 10 },
];

const SECTORS = ["Todos", "Salão", "Bar", "Caixa", "Cozinha", "Limpeza"];
const SECTOR_COLORS = { Salão: "#2D6A4F", Bar: "#1B4332", Caixa: "#40916C", Cozinha: "#B5450B", Limpeza: "#7B5EA7" };

// ── Cálculo de comissões ──────────────────────────────────────
// F e A = sem comissão no dia · E = 50% · MEI = sem desconto dos 33%
function calcResults(employees, workDays, dailyRevenue, absences) {
  const indivEmployees = employees.filter(e => e.type === "individual");
  const globalEmployees = employees.filter(e => e.type === "global");
  const getStatus = (empId, day) => (absences[empId] || {})[day] || null;
  const getDayRevenue = day => dailyRevenue[day] || { global: "", individual: {} };
  const empTotals = {};
  employees.forEach(e => (empTotals[e.id] = 0));
  let totalBruto = 0, totalIndivComm = 0, totalGlobalPool = 0;

  workDays.forEach(day => {
    const dr = getDayRevenue(day);
    const globalBruto = parseFloat(dr.global) || 0;
    const globalNet = globalBruto * (1 - TAX_RATE);
    totalBruto += globalBruto;
    let indivContribToGlobal = 0;
    indivEmployees.forEach(emp => {
      const sale = parseFloat((dr.individual || {})[emp.id]) || 0;
      totalBruto += sale;
      const net = sale * (1 - TAX_RATE);
      indivContribToGlobal += net * (1 - INDIVIDUAL_RATE);
      const status = getStatus(emp.id, day);
      if (status === "F" || status === "A") return;
      const factor = status === "E" ? 0.5 : 1;
      const garcomComm = net * INDIVIDUAL_RATE * factor;
      empTotals[emp.id] = (empTotals[emp.id] || 0) + garcomComm;
      totalIndivComm += garcomComm;
    });
    const totalDayGlobalPool = globalNet + indivContribToGlobal;
    totalGlobalPool += totalDayGlobalPool;
    const getEffPts = emp => {
      const s = getStatus(emp.id, day);
      if (s === "F" || s === "A") return 0;
      if (s === "E") return emp.points * 0.5;
      return emp.points;
    };
    const g1Pool = totalDayGlobalPool * 0.73;
    const g2Pool = totalDayGlobalPool * 0.27;
    const g1 = globalEmployees.filter(e => ["Salão", "Bar", "Caixa"].includes(e.sector));
    const g2 = globalEmployees.filter(e => ["Cozinha", "Limpeza"].includes(e.sector));
    const g1Pts = g1.reduce((s, e) => s + getEffPts(e), 0);
    const g2Pts = g2.reduce((s, e) => s + getEffPts(e), 0);
    g1.forEach(e => { const p = getEffPts(e); if (g1Pts > 0 && p > 0) empTotals[e.id] = (empTotals[e.id] || 0) + (p / g1Pts) * g1Pool; });
    g2.forEach(e => { const p = getEffPts(e); if (g2Pts > 0 && p > 0) empTotals[e.id] = (empTotals[e.id] || 0) + (p / g2Pts) * g2Pool; });
  });

  // MEI: undo the 33% deduction (paga comissão sobre valor bruto)
  employees.forEach(emp => {
    if (emp.mei) empTotals[emp.id] = empTotals[emp.id] / (1 - TAX_RATE);
  });

  return { empTotals, totalBruto, totalIndivComm, totalGlobalPool };
}

// ── App root ──────────────────────────────────────────────────
export default function App() {
  const [authed, setAuthed] = useState(() => sessionStorage.getItem(SESSION_KEY) === "1");
  if (!authed) return <LoginScreen onLogin={() => setAuthed(true)} />;
  return <MainApp />;
}

// ── Main App ──────────────────────────────────────────────────
function MainApp() {
  const now = new Date();
  const initialMonth = now.getFullYear() + "-" + String(now.getMonth() + 1).padStart(2, "0");
  const [month, setMonth] = useState(initialMonth);
  const [employees, setEmployees] = useState(() => loadEmployees() || INITIAL_EMPLOYEES);
  const [step, setStep] = useState("revenue");
  const [sector, setSector] = useState("Todos");
  const [showAdd, setShowAdd] = useState(false);
  const [newEmp, setNewEmp] = useState({ name: "", role: "", sector: "Salão", type: "individual", points: 15 });
  const [editingEmpId, setEditingEmpId] = useState(null);
  const [editingEmpData, setEditingEmpData] = useState(null);
  const [history, setHistory] = useState(loadHistory);
  const [viewingHistory, setViewingHistory] = useState(null);
  const [showHistory, setShowHistory] = useState(false);
  const [halfEmpIds, setHalfEmpIds] = useState({});
  const [importMsg, setImportMsg] = useState(null);
  const [pontoMsg, setPontoMsg] = useState(null);
  const printRef = useRef();
  const importInputRef = useRef();
  const pontoInputRef = useRef();

  const [dailyRevenue, setDailyRevenue] = useState(() => loadMonthData(initialMonth).rev);
  const [absences, setAbsences]         = useState(() => loadMonthData(initialMonth).abs);

  useEffect(() => { saveRevenueForMonth(month, dailyRevenue); }, [dailyRevenue, month]);
  useEffect(() => { saveAbsencesForMonth(month, absences); }, [absences, month]);
  useEffect(() => { saveEmployees(employees); }, [employees]);

  const handleMonthChange = newMonth => {
    setMonth(newMonth);
    setViewingHistory(null);
    const { rev, abs } = loadMonthData(newMonth);
    setDailyRevenue(rev);
    setAbsences(abs);
    setHalfEmpIds({});
  };

  const [year, mon] = month.split("-").map(Number);
  const workDays = getWorkingDays(year, mon);
  const monthLabel = new Date(year, mon - 1, 2).toLocaleString("pt-BR", { month: "long", year: "numeric" });

  const getDayRevenue = day => dailyRevenue[day] || { global: "", individual: {} };
  const setGlobalRevenue = (day, val) =>
    setDailyRevenue(p => ({ ...p, [day]: { ...getDayRevenue(day), global: val } }));
  const setIndivRevenue = (day, empId, val) =>
    setDailyRevenue(p => ({
      ...p, [day]: { ...getDayRevenue(day), individual: { ...getDayRevenue(day).individual, [empId]: val } }
    }));

  const getStatus = (empId, day) => (absences[empId] || {})[day] || null;

  // Ciclo: · → F → E → A → ·
  const toggleAbsence = (empId, day) =>
    setAbsences(p => {
      const cur = (p[empId] || {})[day] || null;
      const next = cur === null ? "F" : cur === "F" ? "E" : cur === "E" ? "A" : null;
      return { ...p, [empId]: { ...(p[empId] || {}), [day]: next } };
    });

  const setAbsenceStatus = (empId, day, status) =>
    setAbsences(p => ({ ...p, [empId]: { ...(p[empId] || {}), [day]: status } }));

  const faltaCountByEmp = empId => workDays.filter(d => getStatus(empId, d) === "F").length;
  const esqCountByEmp   = empId => workDays.filter(d => getStatus(empId, d) === "E").length;
  const atestCountByEmp = empId => workDays.filter(d => getStatus(empId, d) === "A").length;

  const toggleHalf = empId => setHalfEmpIds(p => ({ ...p, [empId]: !p[empId] }));

  // MEI toggle (persiste no perfil do funcionário)
  const toggleEmpMei = empId =>
    setEmployees(p => p.map(e => e.id === empId ? { ...e, mei: !e.mei } : e));

  // Editar funcionário
  const startEditEmp = emp => {
    setEditingEmpId(emp.id);
    setEditingEmpData({ ...emp });
    setShowAdd(false);
  };
  const saveEditEmp = () => {
    if (!editingEmpData?.name?.trim()) return;
    setEmployees(p => p.map(e => e.id === editingEmpId
      ? { ...editingEmpData, points: parseInt(editingEmpData.points) || 15 }
      : e));
    setEditingEmpId(null); setEditingEmpData(null);
  };
  const cancelEditEmp = () => { setEditingEmpId(null); setEditingEmpData(null); };

  // Deletar funcionário
  const deleteEmp = empId => {
    if (window.confirm("Remover este funcionário da lista?"))
      setEmployees(p => p.filter(e => e.id !== empId));
  };

  // Reset por aba
  const handleResetRevenue = () => {
    if (window.confirm("Zerar todo o faturamento de " + monthLabel + "?\nEsta ação não pode ser desfeita."))
      setDailyRevenue({});
  };
  const handleResetAbsences = () => {
    if (window.confirm("Zerar todas as marcações de faltas de " + monthLabel + "?\nEsta ação não pode ser desfeita."))
      setAbsences({});
  };

  const indivEmployees = employees.filter(e => e.type === "individual");
  const results = useMemo(
    () => calcResults(employees, workDays, dailyRevenue, absences),
    [dailyRevenue, absences, employees, month]
  );

  const activeHistory = viewingHistory ? history.find(h => h.monthKey === viewingHistory) : null;
  const displayResults   = activeHistory ? activeHistory.results   : results;
  const displayEmployees = activeHistory ? activeHistory.employees : employees;
  const displayMonthLabel = activeHistory
    ? new Date(activeHistory.year, activeHistory.mon - 1, 2).toLocaleString("pt-BR", { month: "long", year: "numeric" })
    : monthLabel;
  const filteredEmps = sector === "Todos" ? displayEmployees : displayEmployees.filter(e => e.sector === sector);

  const handleSaveHistory = () => {
    const record = { monthKey: month, year, mon, monthLabel, savedAt: new Date().toISOString(), employees, results, totalBruto: results.totalBruto };
    const updated = [record, ...history.filter(h => h.monthKey !== month)].slice(0, 24);
    setHistory(updated); saveHistory(updated);
    alert("Comissões de " + monthLabel + " salvas!");
  };

  // ── Importar faturamento ─────────────────────────────────────
  const handleImportFile = e => {
    const file = e.target.files[0];
    if (!file) return;
    setImportMsg({ type: "loading", text: "Processando planilha..." });
    parseExcelImport(file, employees, month,
      ({ dailyData, newEmps, log, fileMonth }) => {
        if (fileMonth && fileMonth !== month) {
          const ok = window.confirm("A planilha é do mês " + fileMonth + ", mas você está no mês " + month + ".\nDeseja importar mesmo assim?");
          if (!ok) { setImportMsg(null); return; }
        }
        if (newEmps.length > 0) setEmployees(prev => [...prev, ...newEmps]);
        setDailyRevenue(prev => {
          const next = { ...prev };
          for (const [dayStr, data] of Object.entries(dailyData)) {
            const day = Number(dayStr);
            const ex = next[day] || { global: "", individual: {} };
            const newIndividual = { ...(ex.individual || {}), ...data.individual };
            const existingGlobal = parseFloat(ex.global) || 0;
            const newGlobal = existingGlobal + data.globalSum;
            next[day] = { global: newGlobal > 0 ? String(newGlobal.toFixed(2)) : ex.global, individual: newIndividual };
          }
          return next;
        });
        const indC = log.filter(l => l.type === "individual").length;
        const gloC = log.filter(l => l.type === "global").length;
        const newC = newEmps.length;
        setImportMsg({ type: "success", text: "✓ " + indC + " individuais, " + gloC + " globais" + (newC > 0 ? ", " + newC + " novo(s)" : ".") });
        setTimeout(() => setImportMsg(null), 6000);
      },
      err => { setImportMsg({ type: "error", text: err }); setTimeout(() => setImportMsg(null), 8000); }
    );
    e.target.value = "";
  };

  // ── Importar ponto ───────────────────────────────────────────
  const handlePontoImport = e => {
    const file = e.target.files[0];
    if (!file) return;
    setPontoMsg({ type: "loading", text: "Processando planilha de ponto..." });
    parsePontoImport(file, employees, month,
      ({ absToAdd, log, fileMonth }) => {
        if (fileMonth && fileMonth !== month) {
          const ok = window.confirm("A planilha de ponto é do mês " + fileMonth + ", mas você está no mês " + month + ".\nDeseja importar mesmo assim?");
          if (!ok) { setPontoMsg(null); return; }
        }
        setAbsences(prev => {
          const next = { ...prev };
          for (const [empId, days] of Object.entries(absToAdd)) {
            next[empId] = { ...(next[empId] || {}), ...days };
          }
          return next;
        });
        const counts = { F: 0, E: 0, A: 0 };
        log.filter(l => l.matched !== false).forEach(l => { counts[l.status] = (counts[l.status] || 0) + 1; });
        const unmatched = log.filter(l => l.matched === false).length;
        setPontoMsg({
          type: "success",
          text: "✓ " + (counts.F || 0) + "F · " + (counts.E || 0) + "E · " + (counts.A || 0) + "A importados" +
                (unmatched > 0 ? " · " + unmatched + " não encontrados" : "")
        });
        setTimeout(() => setPontoMsg(null), 8000);
      },
      err => { setPontoMsg({ type: "error", text: err }); setTimeout(() => setPontoMsg(null), 8000); }
    );
    e.target.value = "";
  };

  // ── PDF ──────────────────────────────────────────────────────
  const handlePrint = () => {
    const content = printRef.current;
    const win = window.open("", "_blank");
    win.document.write(`<html><head><title>Comissões ${displayMonthLabel} — Maguje</title>
      <style>
        *{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;font-size:11px;color:#111;padding:20px}
        h1{font-size:16px;color:#1B4332;margin-bottom:4px}.sub{font-size:11px;color:#666;margin-bottom:16px}
        .half-badge{display:inline-block;background:#f39c12;color:#fff;padding:2px 8px;border-radius:10px;font-size:9px;margin-left:6px}
        .mei-badge{display:inline-block;background:#27ae60;color:#fff;padding:2px 6px;border-radius:10px;font-size:9px;margin-left:4px}
        .summary{display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap}
        .sum-card{border:1px solid #ccc;border-radius:4px;padding:8px 14px;min-width:130px}
        .sum-val{font-size:14px;font-weight:700;color:#1B4332}.sum-lbl{font-size:9px;color:#888;text-transform:uppercase;letter-spacing:.06em;margin-top:2px}
        table{width:100%;border-collapse:collapse;margin-bottom:20px}
        th{background:#1B4332;color:#fff;padding:7px 10px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.05em}
        td{padding:6px 10px;border-bottom:1px solid #eee;font-size:11px}tr:nth-child(even) td{background:#fafaf8}
        .comm{font-weight:700;text-align:right;color:#1B4332}.comm-half{font-weight:700;text-align:right;color:#e67e22}
        .absent{color:#c0392b;font-weight:600}.forget{color:#f39c12;font-weight:600}.attest{color:#2980b9;font-weight:600}
        tfoot td{background:#e8f0eb!important;font-weight:700;border-top:2px solid #1B4332}
        @media print{body{padding:10px}}
      </style></head><body>${content.innerHTML}</body></html>`);
    win.document.close();
    setTimeout(() => { win.focus(); win.print(); }, 400);
  };

  const addEmployee = () => {
    if (!newEmp.name.trim()) return;
    setEmployees(p => [...p, { ...newEmp, id: Date.now(), points: parseInt(newEmp.points) || 15 }]);
    setNewEmp({ name: "", role: "", sector: "Salão", type: "individual", points: 15 });
    setShowAdd(false);
  };

  const S = {
    wrap: { fontFamily: "'DM Mono','Courier New',monospace", background: "#F5F0E8", minHeight: "100vh" },
    header: { background: "#1B4332", padding: "16px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 },
    card: { background: "#fff", border: "1.5px solid #D4CFC4", borderRadius: 4, padding: "16px 20px" },
    th: { padding: "9px 12px", textAlign: "left", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", background: "#1B4332", color: "#fff", whiteSpace: "nowrap" },
    td: { padding: "8px 12px", fontSize: 13, borderBottom: "1px solid #eee", verticalAlign: "middle" },
    input: { background: "#F5F0E8", border: "1px solid #ccc", borderRadius: 3, padding: "5px 8px", fontFamily: "inherit", fontSize: 12, width: "100%" },
    btn: { border: "1.5px solid #1B4332", background: "#1B4332", color: "#fff", padding: "8px 18px", borderRadius: 3, cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 500 },
    btnOut: { border: "1.5px solid #1B4332", background: "transparent", color: "#1B4332", padding: "8px 16px", borderRadius: 3, cursor: "pointer", fontFamily: "inherit", fontSize: 13 },
    btnGreen: { border: "1.5px solid #40916C", background: "#40916C", color: "#fff", padding: "8px 16px", borderRadius: 3, cursor: "pointer", fontFamily: "inherit", fontSize: 13 },
    btnOrange: { border: "1.5px solid #e67e22", background: "#e67e22", color: "#fff", padding: "8px 16px", borderRadius: 3, cursor: "pointer", fontFamily: "inherit", fontSize: 13 },
    btnRed: { border: "1.5px solid #c0392b", background: "transparent", color: "#c0392b", padding: "7px 14px", borderRadius: 3, cursor: "pointer", fontFamily: "inherit", fontSize: 12 },
    btnAmber: { border: "1.5px solid #f39c12", background: "#f39c12", color: "#fff", padding: "8px 16px", borderRadius: 3, cursor: "pointer", fontFamily: "inherit", fontSize: 13 },
    tab: { padding: "6px 14px", border: "1px solid #ccc", borderRadius: 20, cursor: "pointer", fontSize: 12, background: "transparent", fontFamily: "inherit" },
    tabActive: { padding: "6px 14px", border: "1px solid #1B4332", borderRadius: 20, cursor: "pointer", fontSize: 12, background: "#1B4332", color: "#fff", fontFamily: "inherit" },
    stepBtn: active => ({ padding: "8px 20px", border: "none", borderBottom: active ? "2px solid #1B4332" : "2px solid transparent", background: "transparent", cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: active ? 600 : 400, color: active ? "#1B4332" : "#888" }),
  };

  // Painel de edição de funcionário (reutilizado nas 3 abas)
  const EmpFormPanel = ({ data, setData, onSave, onCancel, title, borderColor }) => (
    <div style={{ ...S.card, marginBottom: 16, borderColor: borderColor || "#f39c12" }}>
      <div style={{ fontSize: 11, color: borderColor || "#f39c12", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 12 }}>{title}</div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
        <div style={{ flex: 2, minWidth: 150 }}>
          <div style={{ fontSize: 11, color: "#555", marginBottom: 4 }}>Nome</div>
          <input type="text" value={data.name} onChange={e => setData(p => ({ ...p, name: e.target.value }))} placeholder="Nome completo" style={S.input} />
        </div>
        <div style={{ flex: 1, minWidth: 120 }}>
          <div style={{ fontSize: 11, color: "#555", marginBottom: 4 }}>Cargo</div>
          <input type="text" value={data.role} onChange={e => setData(p => ({ ...p, role: e.target.value }))} placeholder="Cargo" style={S.input} />
        </div>
        <div style={{ flex: 1, minWidth: 100 }}>
          <div style={{ fontSize: 11, color: "#555", marginBottom: 4 }}>Setor</div>
          <select value={data.sector} onChange={e => setData(p => ({ ...p, sector: e.target.value }))} style={S.input}>
            {["Salão", "Bar", "Caixa", "Cozinha", "Limpeza"].map(x => <option key={x}>{x}</option>)}
          </select>
        </div>
        <div style={{ flex: 1, minWidth: 110 }}>
          <div style={{ fontSize: 11, color: "#555", marginBottom: 4 }}>Tipo</div>
          <select value={data.type} onChange={e => setData(p => ({ ...p, type: e.target.value }))} style={S.input}>
            <option value="individual">Individual (Garçom)</option>
            <option value="global">Global (Pool por pontos)</option>
          </select>
        </div>
        {data.type === "global" && (
          <div style={{ flex: 1, minWidth: 80 }}>
            <div style={{ fontSize: 11, color: "#555", marginBottom: 4 }}>Pontos</div>
            <input type="number" min="1" max="50" value={data.points} onChange={e => setData(p => ({ ...p, points: e.target.value }))} style={S.input} />
          </div>
        )}
        <div style={{ flex: 1, minWidth: 80 }}>
          <div style={{ fontSize: 11, color: "#555", marginBottom: 4 }}>MEI</div>
          <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", marginTop: 4 }}>
            <input type="checkbox" checked={!!data.mei} onChange={e => setData(p => ({ ...p, mei: e.target.checked }))} />
            <span style={{ fontSize: 12 }}>MEI</span>
          </label>
        </div>
        <div style={{ display: "flex", alignItems: "flex-end", gap: 8 }}>
          <button style={S.btn} onClick={onSave}>Salvar</button>
          <button style={S.btnOut} onClick={onCancel}>Cancelar</button>
        </div>
      </div>
    </div>
  );

  // ── PDF Content ───────────────────────────────────────────────
  const PrintContent = ({ emps, res, label, halfIds }) => {
    const totalComm = emps.reduce((s, e) => s + (res.empTotals[e.id] || 0) * (halfIds[e.id] ? 0.5 : 1), 0);
    const hasHalf = emps.some(e => halfIds[e.id]);
    return (
      <div>
        <h1>Restaurante Maguje — Comissões{hasHalf && <span className="half-badge">50% parcial</span>}</h1>
        <div className="sub">Referência: {label} · Gerado em {new Date().toLocaleDateString("pt-BR")}</div>
        <div className="summary">
          {[
            { label: "Total Bruto", val: fmt(res.totalBruto) },
            { label: "Pool Global Líq.", val: fmt(res.totalGlobalPool) },
            { label: "Comissões Indiv.", val: fmt(res.totalIndivComm) },
            { label: "Total Distribuído", val: fmt(totalComm) },
          ].map(m => (
            <div key={m.label} className="sum-card">
              <div className="sum-val">{m.val}</div>
              <div className="sum-lbl">{m.label}</div>
            </div>
          ))}
        </div>
        {["Salão", "Bar", "Caixa", "Cozinha", "Limpeza"].map(sec => {
          const secEmps = emps.filter(e => e.sector === sec);
          if (!secEmps.length) return null;
          const secTotal = secEmps.reduce((s, e) => s + (res.empTotals[e.id] || 0) * (halfIds[e.id] ? 0.5 : 1), 0);
          return (
            <table key={sec}>
              <thead>
                <tr><th colSpan={5} style={{ background: "#2D6A4F" }}>{sec}</th></tr>
                <tr><th>Funcionário</th><th>Cargo</th><th>Tipo</th><th style={{ textAlign: "center" }}>F/E/A</th><th style={{ textAlign: "right" }}>Comissão</th></tr>
              </thead>
              <tbody>
                {secEmps.sort((a, b) => (res.empTotals[b.id] || 0) - (res.empTotals[a.id] || 0)).map(emp => {
                  const comm = (res.empTotals[emp.id] || 0) * (halfIds[emp.id] ? 0.5 : 1);
                  const fC = workDays.filter(d => getStatus(emp.id, d) === "F").length;
                  const eC = workDays.filter(d => getStatus(emp.id, d) === "E").length;
                  const aC = workDays.filter(d => getStatus(emp.id, d) === "A").length;
                  return (
                    <tr key={emp.id}>
                      <td>{emp.name}{halfIds[emp.id] && <span className="half-badge">50%</span>}{emp.mei && <span className="mei-badge">MEI</span>}</td>
                      <td>{emp.role}</td>
                      <td>{emp.type === "individual" ? "Individual (29%)" : "Global (" + emp.points + "pts)"}</td>
                      <td style={{ textAlign: "center" }}>
                        {fC > 0 && <span className="absent">{fC}F </span>}
                        {eC > 0 && <span className="forget">{eC}E </span>}
                        {aC > 0 && <span className="attest">{aC}A</span>}
                        {fC === 0 && eC === 0 && aC === 0 && "—"}
                      </td>
                      <td className={halfIds[emp.id] ? "comm-half" : "comm"}>{fmt(comm)}</td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot><tr><td colSpan={4}>Total {sec}</td><td className="comm">{fmt(secTotal)}</td></tr></tfoot>
            </table>
          );
        })}
        <table>
          <tfoot>
            <tr style={{ background: "#1B4332", color: "#fff" }}>
              <td colSpan={4} style={{ fontWeight: 700, fontSize: 13 }}>TOTAL GERAL</td>
              <td style={{ fontWeight: 700, fontSize: 14, textAlign: "right" }}>{fmt(totalComm)}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    );
  };

  // Badges de falta para reuso
  const AbsBadges = ({ empId, small }) => {
    const fC = faltaCountByEmp(empId), eC = esqCountByEmp(empId), aC = atestCountByEmp(empId);
    const sz = small ? { fontSize: 10, padding: "1px 5px" } : { fontSize: 11, padding: "2px 7px" };
    return (
      <>
        {fC > 0 && <span style={{ ...sz, background: "#fdecea", color: "#c0392b", borderRadius: 10, fontWeight: 600, marginRight: 2 }}>{fC}F</span>}
        {eC > 0 && <span style={{ ...sz, background: "#fff3cd", color: "#7a5c00", borderRadius: 10, fontWeight: 600, marginRight: 2 }}>{eC}E</span>}
        {aC > 0 && <span style={{ ...sz, background: "#dbeafe", color: "#1a6fa0", borderRadius: 10, fontWeight: 600 }}>{aC}A</span>}
        {fC === 0 && eC === 0 && aC === 0 && <span style={{ color: "#ccc", fontSize: 12 }}>—</span>}
      </>
    );
  };

  // Botões editar/deletar por linha
  const EmpActions = ({ emp, small }) => (
    <span style={{ display: "inline-flex", gap: 4, marginLeft: 6 }}>
      <button onClick={() => startEditEmp(emp)}
        title="Editar"
        style={{ background: "none", border: "1px solid #ccc", borderRadius: 3, cursor: "pointer", fontSize: small ? 10 : 11, padding: "1px 5px", color: "#666" }}>✏</button>
      <button onClick={() => deleteEmp(emp.id)}
        title="Remover"
        style={{ background: "none", border: "1px solid #fbc8c8", borderRadius: 3, cursor: "pointer", fontSize: small ? 10 : 11, padding: "1px 5px", color: "#c0392b" }}>🗑</button>
    </span>
  );

  // Msg helper
  const MsgBox = ({ msg }) => !msg ? null : (
    <div style={{ fontSize: 12, padding: "7px 12px", borderRadius: 4,
      background: msg.type === "success" ? "#d4edda" : msg.type === "error" ? "#fdecea" : "#fff3cd",
      color: msg.type === "success" ? "#155724" : msg.type === "error" ? "#721c24" : "#856404",
      border: "1px solid " + (msg.type === "success" ? "#c3e6cb" : msg.type === "error" ? "#f5c6cb" : "#ffeeba"),
      maxWidth: 360 }}>{msg.text}</div>
  );

  return (
    <div style={S.wrap}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Space+Grotesk:wght@500;700&display=swap');
        * { box-sizing: border-box; }
        input:focus { outline: none; border-color: #2D6A4F !important; }
        .absent-btn { width:28px; height:28px; border-radius:4px; border:1.5px solid #ddd; background:#fff; cursor:pointer; font-size:12px; font-weight:700; display:flex; align-items:center; justify-content:center; transition:all 0.1s; font-family:inherit; }
        .absent-btn.marked-f { background:#c0392b; border-color:#c0392b; color:#fff; }
        .absent-btn.marked-e { background:#f39c12; border-color:#e67e22; color:#fff; }
        .absent-btn.marked-a { background:#2980b9; border-color:#1a6fa0; color:#fff; }
        .absent-btn:hover { border-color:#888; }
        .row-hover:hover td { background:#fafaf7 !important; }
        ::-webkit-scrollbar { height:5px; width:5px; } ::-webkit-scrollbar-thumb { background:#ccc; border-radius:3px; }
        .print-hidden { display:none; }
      `}</style>

      <div ref={printRef} className="print-hidden">
        <PrintContent emps={displayEmployees} res={displayResults} label={displayMonthLabel} halfIds={halfEmpIds} />
      </div>

      {/* Header */}
      <div style={S.header}>
        <div>
          <div style={{ color: "#52B788", fontSize: 10, letterSpacing: "0.15em", textTransform: "uppercase", marginBottom: 3 }}>Restaurante Maguje</div>
          <div style={{ color: "#fff", fontSize: 20, fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700 }}>Calculadora de Comissões</div>
          <div style={{ color: "#95D5B2", fontSize: 12, marginTop: 2, textTransform: "capitalize" }}>{monthLabel}</div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <input type="month" value={month} onChange={e => handleMonthChange(e.target.value)}
            style={{ ...S.input, width: 150, background: "#2D6A4F", color: "#fff", border: "1px solid #52B788", colorScheme: "dark" }} />
          <button style={{ ...S.btnOut, color: "#95D5B2", borderColor: "#52B788", fontSize: 12, padding: "7px 14px" }}
            onClick={() => setShowHistory(!showHistory)}>
            🗂 Histórico {history.length > 0 && "(" + history.length + ")"}
          </button>
          <button style={{ ...S.btnOut, color: "#f28b82", borderColor: "#f28b82", fontSize: 12, padding: "7px 14px" }}
            onClick={() => { sessionStorage.removeItem(SESSION_KEY); window.location.reload(); }}>
            Sair
          </button>
        </div>
      </div>

      {/* Histórico */}
      {showHistory && (
        <div style={{ background: "#fff", borderBottom: "1px solid #D4CFC4", padding: "16px 24px" }}>
          <div style={{ fontSize: 11, color: "#666", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 12 }}>Histórico</div>
          {history.length === 0
            ? <div style={{ fontSize: 13, color: "#aaa" }}>Nenhum mês salvo ainda.</div>
            : <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
              {history.map(h => (
                <div key={h.monthKey}
                  onClick={() => { setViewingHistory(h.monthKey === viewingHistory ? null : h.monthKey); setStep("results"); setShowHistory(false); }}
                  style={{ border: "1.5px solid " + (viewingHistory === h.monthKey ? "#1B4332" : "#D4CFC4"), borderRadius: 4, padding: "10px 16px", cursor: "pointer", background: viewingHistory === h.monthKey ? "#1B4332" : "#fff", color: viewingHistory === h.monthKey ? "#fff" : "#333", minWidth: 140 }}>
                  <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 13, textTransform: "capitalize" }}>{h.monthLabel}</div>
                  <div style={{ fontSize: 11, marginTop: 3, opacity: 0.7 }}>Total: {fmt(h.totalBruto)}</div>
                  <div style={{ fontSize: 10, marginTop: 2, opacity: 0.5 }}>Salvo {new Date(h.savedAt).toLocaleDateString("pt-BR")}</div>
                </div>
              ))}
            </div>}
          {viewingHistory && (
            <div style={{ marginTop: 12, padding: "10px 14px", background: "#fff8e6", border: "1px solid #f0c040", borderRadius: 4, fontSize: 12, color: "#7a5c00" }}>
              Visualizando <strong style={{ textTransform: "capitalize" }}>{displayMonthLabel}</strong>.{" "}
              <span style={{ cursor: "pointer", textDecoration: "underline" }} onClick={() => setViewingHistory(null)}>Voltar ao mês atual</span>
            </div>
          )}
        </div>
      )}

      {/* Step tabs */}
      {!viewingHistory && (
        <div style={{ background: "#fff", borderBottom: "1px solid #e0dbd0", padding: "0 24px", display: "flex", gap: 0 }}>
          {[["revenue", "1. Faturamento Diário"], ["absences", "2. Faltas por Dia"], ["results", "3. Resultado"]].map(([key, label]) => (
            <button key={key} style={S.stepBtn(step === key)} onClick={() => setStep(key)}>{label}</button>
          ))}
        </div>
      )}

      <div style={{ padding: "20px 24px" }}>

        {/* Painel de edição (aparece em qualquer aba quando editingEmpId está ativo) */}
        {!viewingHistory && editingEmpId && editingEmpData && (
          <EmpFormPanel
            data={editingEmpData} setData={setEditingEmpData}
            onSave={saveEditEmp} onCancel={cancelEditEmp}
            title={"✏ Editar: " + (editingEmpData.name || "")}
            borderColor="#f39c12"
          />
        )}

        {/* ── HISTÓRICO ── */}
        {viewingHistory && (
          <>
            <div style={{ ...S.card, marginBottom: 16, borderColor: "#f0c040", background: "#fffdf0" }}>
              <div style={{ fontSize: 12, color: "#7a5c00" }}>
                📋 Visualizando histórico de <strong style={{ textTransform: "capitalize" }}>{displayMonthLabel}</strong>
                {" · "}<span style={{ cursor: "pointer", textDecoration: "underline" }} onClick={() => setViewingHistory(null)}>Voltar ao mês atual</span>
              </div>
            </div>
            <ResultsTable emps={displayEmployees} res={displayResults} sector={sector} setSector={setSector}
              S={S} SECTORS={SECTORS} SECTOR_COLORS={SECTOR_COLORS} fmt={fmt}
              onPrint={handlePrint} showSave={false} onSave={null}
              halfEmpIds={halfEmpIds} toggleHalf={toggleHalf}
              absCountFns={{ falta: () => 0, esq: () => 0, atest: () => 0 }}
              isHistory={true} onToggleMei={null} onEditEmp={null} onDeleteEmp={null} />
          </>
        )}

        {/* ── STEP 1: FATURAMENTO ── */}
        {!viewingHistory && step === "revenue" && (
          <>
            {/* Importar planilha */}
            <div style={{ ...S.card, marginBottom: 16 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
                <div>
                  <div style={{ fontSize: 11, color: "#666", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 4 }}>Importar planilha do sistema</div>
                  <div style={{ fontSize: 12, color: "#555" }}>
                    Garçons <strong>sem fundo</strong> = venda individual · Fundo cinza = pool global
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                  <MsgBox msg={importMsg} />
                  <input ref={importInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImportFile} />
                  <button style={S.btnOrange} onClick={() => importInputRef.current && importInputRef.current.click()}>
                    📂 Importar Planilha
                  </button>
                  <button style={S.btnRed} onClick={handleResetRevenue} title="Zerar faturamento do mês">
                    🗑 Zerar Faturamento
                  </button>
                </div>
              </div>
            </div>

            <div style={{ ...S.card, marginBottom: 16 }}>
              <div style={{ fontSize: 11, color: "#666", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Regra de comissão</div>
              <div style={{ fontSize: 12, color: "#555", lineHeight: 1.7 }}>
                Venda bruta − 33% = líquido · <strong>29%</strong> do líquido → garçom · <strong>71%</strong> do líquido → pool global. Valores <strong>salvos automaticamente</strong>.
              </div>
            </div>

            <div style={{ overflowX: "auto" }}>
              <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 200 + workDays.length * 60 }}>
                <thead>
                  <tr>
                    <th style={{ ...S.th, position: "sticky", left: 0, zIndex: 2, minWidth: 220 }}>Funcionário</th>
                    {workDays.map(d => {
                      const dow = new Date(year, mon - 1, d).getDay();
                      return <th key={d} style={{ ...S.th, textAlign: "center", minWidth: 60 }}>
                        <div>{d}</div><div style={{ fontWeight: 300, fontSize: 9, opacity: 0.7 }}>{DOW_LABELS[dow]}</div>
                      </th>;
                    })}
                    <th style={{ ...S.th, textAlign: "right", minWidth: 110 }}>Total Bruto</th>
                  </tr>
                </thead>
                <tbody>
                  <tr><td colSpan={workDays.length + 2} style={{ ...S.td, background: "#1B433210", fontWeight: 600, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", color: "#1B4332", padding: "10px 12px" }}>
                    Venda Individual — 29% líq. ao garçom, 71% líq. ao pool global
                  </td></tr>
                  {indivEmployees.map((emp, idx) => {
                    const total = workDays.reduce((s, d) => s + (parseFloat(getDayRevenue(d).individual[emp.id]) || 0), 0);
                    return (
                      <tr key={emp.id} className="row-hover">
                        <td style={{ ...S.td, position: "sticky", left: 0, background: idx % 2 === 0 ? "#fff" : "#fafaf8", zIndex: 1, fontWeight: 500 }}>
                          <div style={{ fontSize: 12, display: "flex", alignItems: "center" }}>
                            {emp.name}
                            <EmpActions emp={emp} small />
                          </div>
                          <div style={{ fontSize: 10, color: SECTOR_COLORS[emp.sector] || "#888", marginTop: 1 }}>{emp.role}</div>
                        </td>
                        {workDays.map(d => (
                          <td key={d} style={{ ...S.td, padding: "5px 6px", background: idx % 2 === 0 ? "#fff" : "#fafaf8" }}>
                            <input type="number" min="0" placeholder="0"
                              value={getDayRevenue(d).individual[emp.id] || ""}
                              onChange={ev => setIndivRevenue(d, emp.id, ev.target.value)}
                              style={{ ...S.input, width: 52, textAlign: "right", padding: "4px 5px" }} />
                          </td>
                        ))}
                        <td style={{ ...S.td, textAlign: "right", fontFamily: "'Space Grotesk',sans-serif", fontWeight: 600, color: "#1B4332", background: idx % 2 === 0 ? "#fff" : "#fafaf8" }}>
                          {fmt(total)}
                        </td>
                      </tr>
                    );
                  })}
                  <tr><td colSpan={workDays.length + 2} style={{ ...S.td, background: "#40916C18", fontWeight: 600, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", color: "#40916C", padding: "10px 12px" }}>
                    Faturamento Global Adicional do Dia
                  </td></tr>
                  <tr className="row-hover">
                    <td style={{ ...S.td, position: "sticky", left: 0, background: "#fff", zIndex: 1, fontWeight: 500 }}>
                      <div style={{ fontSize: 12 }}>Outras vendas (sem comissão individual)</div>
                    </td>
                    {workDays.map(d => (
                      <td key={d} style={{ ...S.td, padding: "5px 6px", background: "#fff" }}>
                        <input type="number" min="0" placeholder="0"
                          value={getDayRevenue(d).global || ""}
                          onChange={ev => setGlobalRevenue(d, ev.target.value)}
                          style={{ ...S.input, width: 52, textAlign: "right", padding: "4px 5px", borderColor: "#40916C60" }} />
                      </td>
                    ))}
                    <td style={{ ...S.td, textAlign: "right", fontFamily: "'Space Grotesk',sans-serif", fontWeight: 600, color: "#40916C" }}>
                      {fmt(workDays.reduce((s, d) => s + (parseFloat(getDayRevenue(d).global) || 0), 0))}
                    </td>
                  </tr>
                  <tr style={{ background: "#F5F0E8" }}>
                    <td style={{ ...S.td, position: "sticky", left: 0, background: "#F5F0E8", zIndex: 1, fontWeight: 600, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em" }}>Total do Dia</td>
                    {workDays.map(d => {
                      const dr = getDayRevenue(d);
                      const tot = indivEmployees.reduce((s, e) => s + (parseFloat(dr.individual[e.id]) || 0), 0) + (parseFloat(dr.global) || 0);
                      return <td key={d} style={{ ...S.td, textAlign: "right", fontWeight: 600, fontSize: 11, background: "#F5F0E8", padding: "7px 5px", color: "#1B4332" }}>
                        {tot > 0 ? fmtShort(tot) : "—"}
                      </td>;
                    })}
                    <td style={{ ...S.td, textAlign: "right", fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, color: "#1B4332", background: "#F5F0E8" }}>
                      {fmt(workDays.reduce((s, d) => {
                        const dr = getDayRevenue(d);
                        return s + indivEmployees.reduce((ss, e) => ss + (parseFloat(dr.individual[e.id]) || 0), 0) + (parseFloat(dr.global) || 0);
                      }, 0))}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div style={{ marginTop: 20, textAlign: "right" }}>
              <button style={S.btn} onClick={() => setStep("absences")}>Próximo: Faltas →</button>
            </div>
          </>
        )}

        {/* ── STEP 2: FALTAS ── */}
        {!viewingHistory && step === "absences" && (
          <>
            {/* Legenda */}
            <div style={{ ...S.card, marginBottom: 16 }}>
              <div style={{ fontSize: 11, color: "#666", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 8 }}>Tipos de marcação · Clique para ciclar: · → F → E → A → ·</div>
              <div style={{ display: "flex", gap: 16, flexWrap: "wrap", fontSize: 12, color: "#555" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span style={{ width: 26, height: 26, borderRadius: 4, background: "#c0392b", color: "#fff", display: "inline-flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 12 }}>F</span>
                  <span><strong>Falta</strong> — sem comissão</span>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span style={{ width: 26, height: 26, borderRadius: 4, background: "#f39c12", color: "#fff", display: "inline-flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 12 }}>E</span>
                  <span><strong>Esquecimento de batida</strong> — 50% da comissão</span>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span style={{ width: 26, height: 26, borderRadius: 4, background: "#2980b9", color: "#fff", display: "inline-flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 12 }}>A</span>
                  <span><strong>Atestado médico</strong> — sem comissão</span>
                </div>
              </div>
            </div>

            {/* Importar ponto + controles */}
            <div style={{ ...S.card, marginBottom: 16 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
                <div>
                  <div style={{ fontSize: 11, color: "#666", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 4 }}>Importar planilha de ponto</div>
                  <div style={{ fontSize: 12, color: "#555" }}>
                    Classifica automaticamente: <strong>Falta→F</strong> · <strong>Esquecimento→E</strong> · <strong>Atestado→A</strong>. Folgas e domingos são ignorados.
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                  <MsgBox msg={pontoMsg} />
                  <input ref={pontoInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handlePontoImport} />
                  <button style={S.btnAmber} onClick={() => pontoInputRef.current && pontoInputRef.current.click()}>
                    📋 Importar Ponto
                  </button>
                  <button style={S.btnRed} onClick={handleResetAbsences} title="Zerar todas as marcações">
                    🗑 Zerar Faltas
                  </button>
                </div>
              </div>
            </div>

            <div style={{ display: "flex", gap: 6, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
              {SECTORS.map(s2 => (
                <button key={s2} style={sector === s2 ? S.tabActive : S.tab} onClick={() => setSector(s2)}>{s2}</button>
              ))}
              <div style={{ marginLeft: "auto" }}>
                <button style={{ ...S.btnOut, fontSize: 12, padding: "6px 14px" }} onClick={() => { setShowAdd(!showAdd); setEditingEmpId(null); }}>
                  {showAdd ? "✕ Cancelar" : "+ Funcionário"}
                </button>
              </div>
            </div>

            {showAdd && !editingEmpId && (
              <EmpFormPanel
                data={newEmp} setData={setNewEmp}
                onSave={addEmployee}
                onCancel={() => setShowAdd(false)}
                title="Novo Funcionário"
                borderColor="#52B788"
              />
            )}

            <div style={{ overflowX: "auto" }}>
              <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 220 + workDays.length * 42 }}>
                <thead>
                  <tr>
                    <th style={{ ...S.th, position: "sticky", left: 0, zIndex: 2, minWidth: 230 }}>Funcionário</th>
                    {workDays.map(d => {
                      const dow = new Date(year, mon - 1, d).getDay();
                      return <th key={d} style={{ ...S.th, textAlign: "center", minWidth: 42, padding: "8px 4px" }}>
                        <div>{d}</div><div style={{ fontWeight: 300, fontSize: 9, opacity: 0.7 }}>{DOW_LABELS[dow]}</div>
                      </th>;
                    })}
                    <th style={{ ...S.th, textAlign: "center", minWidth: 90 }}>F / E / A</th>
                  </tr>
                </thead>
                <tbody>
                  {[
                    { label: "Garçons e Chefes de Fila Junior", emps: filteredEmps.filter(e => e.type === "individual") },
                    { label: "Equipe Global", emps: filteredEmps.filter(e => e.type === "global") },
                  ].map(group => group.emps.length === 0 ? null : (
                    <React.Fragment key={group.label}>
                      <tr><td colSpan={workDays.length + 2} style={{ ...S.td, background: "#1B433210", fontWeight: 600, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", color: "#1B4332", padding: "10px 12px" }}>{group.label}</td></tr>
                      {group.emps.map((emp, idx) => (
                        <tr key={emp.id} className="row-hover">
                          <td style={{ ...S.td, position: "sticky", left: 0, background: idx % 2 === 0 ? "#fff" : "#fafaf8", zIndex: 1 }}>
                            <div style={{ fontSize: 12, fontWeight: 500, display: "flex", alignItems: "center" }}>
                              {emp.name}
                              <EmpActions emp={emp} small />
                            </div>
                            <div style={{ fontSize: 10, color: SECTOR_COLORS[emp.sector] || "#888", marginTop: 1 }}>{emp.role} · {emp.sector}</div>
                          </td>
                          {workDays.map(d => {
                            const status = getStatus(emp.id, d);
                            const bgMap = { F: "#fdecea", E: "#fff8ee", A: "#dbeafe" };
                            return (
                              <td key={d} style={{ ...S.td, textAlign: "center", padding: "5px 4px", background: bgMap[status] || (idx % 2 === 0 ? "#fff" : "#fafaf8") }}>
                                <button className={"absent-btn" + (status === "F" ? " marked-f" : status === "E" ? " marked-e" : status === "A" ? " marked-a" : "")}
                                  onClick={() => toggleAbsence(emp.id, d)}
                                  title={status === "F" ? "Falta" : status === "E" ? "Esquecimento (50%)" : status === "A" ? "Atestado" : "Presente"}>
                                  {status || "·"}
                                </button>
                              </td>
                            );
                          })}
                          <td style={{ ...S.td, textAlign: "center", background: idx % 2 === 0 ? "#fff" : "#fafaf8" }}>
                            <AbsBadges empId={emp.id} />
                          </td>
                        </tr>
                      ))}
                    </React.Fragment>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ marginTop: 20, display: "flex", justifyContent: "space-between" }}>
              <button style={S.btnOut} onClick={() => setStep("revenue")}>← Voltar</button>
              <button style={S.btn} onClick={() => setStep("results")}>Calcular Comissões →</button>
            </div>
          </>
        )}

        {/* ── STEP 3: RESULTADO ── */}
        {!viewingHistory && step === "results" && (
          <ResultsTable emps={employees} res={results} sector={sector} setSector={setSector}
            S={S} SECTORS={SECTORS} SECTOR_COLORS={SECTOR_COLORS} fmt={fmt}
            onPrint={handlePrint} onSave={handleSaveHistory} showSave={true}
            halfEmpIds={halfEmpIds} toggleHalf={toggleHalf}
            absCountFns={{ falta: faltaCountByEmp, esq: esqCountByEmp, atest: atestCountByEmp }}
            isHistory={false}
            onToggleMei={toggleEmpMei}
            onEditEmp={startEditEmp}
            onDeleteEmp={deleteEmp}
            onBack={() => setStep("absences")} />
        )}
      </div>

      <div style={{ textAlign: "center", padding: "14px", fontSize: 11, color: "#aaa", borderTop: "1px solid #e5e0d6", marginTop: 20 }}>
        Maguje · Comissões · {monthLabel} · Dados salvos automaticamente
      </div>
    </div>
  );
}

// ── Tabela de Resultados ──────────────────────────────────────
function ResultsTable({ emps, res, sector, setSector, S, SECTORS, SECTOR_COLORS, fmt, onPrint, onSave, showSave,
  halfEmpIds, toggleHalf, absCountFns, isHistory, onToggleMei, onEditEmp, onDeleteEmp, onBack }) {
  const filteredEmps = sector === "Todos" ? emps : emps.filter(e => e.sector === sector);
  const getComm = emp => (res.empTotals[emp.id] || 0) * (halfEmpIds[emp.id] ? 0.5 : 1);
  const totalComm = filteredEmps.reduce((s, e) => s + getComm(e), 0);
  const { falta: faltaFn, esq: esqFn, atest: atestFn } = absCountFns || { falta: () => 0, esq: () => 0, atest: () => 0 };

  return (
    <>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 12, marginBottom: 20 }}>
        {[
          { label: "Total Bruto", val: fmt(res.totalBruto), color: "#1B4332" },
          { label: "Pool Global Líq.", val: fmt(res.totalGlobalPool), color: "#40916C" },
          { label: "Comissões Indiv.", val: fmt(res.totalIndivComm), color: "#7B5EA7" },
          { label: "Total Distribuído", val: fmt(filteredEmps.reduce((s,e)=>s+getComm(e),0)), color: "#B5450B" },
        ].map(m => (
          <div key={m.label} style={{ background: "#fff", border: "1.5px solid #D4CFC4", borderRadius: 4, padding: "13px 16px", textAlign: "center" }}>
            <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 17, color: m.color }}>{m.val}</div>
            <div style={{ fontSize: 10, color: "#888", textTransform: "uppercase", letterSpacing: "0.07em", marginTop: 3 }}>{m.label}</div>
          </div>
        ))}
      </div>

      <div style={{ ...S.card, marginBottom: 16, background: "#f8f9ff", borderColor: "#c5cae9" }}>
        <div style={{ fontSize: 11, color: "#555", display: "flex", gap: 16, flexWrap: "wrap", alignItems: "center" }}>
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{ background: "#f39c12", color: "#fff", padding: "2px 8px", borderRadius: 10, fontSize: 11, fontWeight: 600 }}>½</span>
            Clique para aplicar 50% individualmente
          </span>
          {!isHistory && (
            <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ background: "#27ae60", color: "#fff", padding: "2px 8px", borderRadius: 10, fontSize: 11, fontWeight: 600 }}>MEI</span>
              Clique para marcar MEI (sem desconto dos 33%)
            </span>
          )}
        </div>
      </div>

      <div style={{ display: "flex", gap: 6, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
        {SECTORS.map(s2 => (
          <button key={s2} style={sector === s2 ? S.tabActive : S.tab} onClick={() => setSector(s2)}>{s2}</button>
        ))}
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          {showSave && <button style={S.btnGreen} onClick={onSave}>💾 Salvar Mês</button>}
          <button style={S.btn} onClick={onPrint}>🖨 Exportar PDF</button>
        </div>
      </div>

      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "100%" }}>
          <thead>
            <tr>
              <th style={{ ...S.th, minWidth: 220 }}>Funcionário</th>
              <th style={S.th}>Cargo</th>
              <th style={{ ...S.th, textAlign: "center" }}>Tipo</th>
              <th style={{ ...S.th, textAlign: "center" }}>F/E/A</th>
              <th style={{ ...S.th, textAlign: "center", width: 50 }}>½</th>
              {!isHistory && <th style={{ ...S.th, textAlign: "center", width: 60 }}>MEI</th>}
              <th style={{ ...S.th, textAlign: "right" }}>Comissão</th>
            </tr>
          </thead>
          <tbody>
            {filteredEmps.slice().sort((a, b) => (res.empTotals[b.id] || 0) - (res.empTotals[a.id] || 0)).map((emp, idx) => {
              const comm = getComm(emp);
              const isHalf = !!halfEmpIds[emp.id];
              const fC = faltaFn(emp.id), eC = esqFn(emp.id), aC = atestFn(emp.id);
              const color = SECTOR_COLORS[emp.sector] || "#555";
              return (
                <tr key={emp.id} className="row-hover">
                  <td style={{ ...S.td, background: idx % 2 === 0 ? "#fff" : "#fafaf8", fontWeight: 500 }}>
                    <div style={{ fontSize: 13, display: "flex", alignItems: "center", gap: 4 }}>
                      {emp.name}
                      {emp.mei && <span style={{ background: "#27ae60", color: "#fff", padding: "1px 6px", borderRadius: 10, fontSize: 9, fontWeight: 700 }}>MEI</span>}
                      {!isHistory && onEditEmp && (
                        <span style={{ display: "inline-flex", gap: 3, marginLeft: 2 }}>
                          <button onClick={() => onEditEmp(emp)} title="Editar"
                            style={{ background: "none", border: "1px solid #ccc", borderRadius: 3, cursor: "pointer", fontSize: 10, padding: "1px 4px", color: "#666" }}>✏</button>
                          <button onClick={() => onDeleteEmp(emp.id)} title="Remover"
                            style={{ background: "none", border: "1px solid #fbc8c8", borderRadius: 3, cursor: "pointer", fontSize: 10, padding: "1px 4px", color: "#c0392b" }}>🗑</button>
                        </span>
                      )}
                    </div>
                    <span style={{ display: "inline-block", fontSize: 10, padding: "1px 7px", borderRadius: 20, marginTop: 2, background: color + "18", color, border: "1px solid " + color + "40", textTransform: "uppercase", letterSpacing: "0.05em" }}>{emp.sector}</span>
                  </td>
                  <td style={{ ...S.td, fontSize: 12, color: "#444", background: idx % 2 === 0 ? "#fff" : "#fafaf8" }}>{emp.role}</td>
                  <td style={{ ...S.td, textAlign: "center", background: idx % 2 === 0 ? "#fff" : "#fafaf8" }}>
                    {emp.type === "individual"
                      ? <span style={{ fontSize: 11, padding: "2px 8px", borderRadius: 20, background: "#7B5EA720", color: "#7B5EA7", border: "1px solid #7B5EA740" }}>Individual</span>
                      : <span style={{ fontSize: 11, padding: "2px 8px", borderRadius: 20, background: "#2D6A4F20", color: "#2D6A4F", border: "1px solid #2D6A4F40" }}>Global · {emp.points}pts</span>}
                  </td>
                  <td style={{ ...S.td, textAlign: "center", background: idx % 2 === 0 ? "#fff" : "#fafaf8" }}>
                    {fC > 0 && <span style={{ background: "#fdecea", color: "#c0392b", borderRadius: 10, padding: "2px 6px", fontSize: 11, fontWeight: 600, marginRight: 2 }}>{fC}F</span>}
                    {eC > 0 && <span style={{ background: "#fff3cd", color: "#7a5c00", borderRadius: 10, padding: "2px 6px", fontSize: 11, fontWeight: 600, marginRight: 2 }}>{eC}E</span>}
                    {aC > 0 && <span style={{ background: "#dbeafe", color: "#1a6fa0", borderRadius: 10, padding: "2px 6px", fontSize: 11, fontWeight: 600 }}>{aC}A</span>}
                    {fC === 0 && eC === 0 && aC === 0 && <span style={{ color: "#ccc", fontSize: 12 }}>—</span>}
                  </td>
                  <td style={{ ...S.td, textAlign: "center", background: idx % 2 === 0 ? "#fff" : "#fafaf8" }}>
                    <button onClick={() => toggleHalf(emp.id)}
                      title={isHalf ? "50% ativo — clique para remover" : "Clique para aplicar 50%"}
                      style={{ width: 30, height: 24, borderRadius: 12, border: "none", cursor: "pointer",
                        background: isHalf ? "#f39c12" : "#e9ecef", color: isHalf ? "#fff" : "#999",
                        fontSize: 12, fontWeight: 700, fontFamily: "inherit", transition: "all 0.15s" }}>
                      ½
                    </button>
                  </td>
                  {!isHistory && (
                    <td style={{ ...S.td, textAlign: "center", background: idx % 2 === 0 ? "#fff" : "#fafaf8" }}>
                      <button onClick={() => onToggleMei && onToggleMei(emp.id)}
                        title={emp.mei ? "MEI ativo — sem desconto 33%" : "Marcar como MEI"}
                        style={{ padding: "3px 8px", borderRadius: 4, border: "none", cursor: "pointer",
                          background: emp.mei ? "#27ae60" : "#e9ecef", color: emp.mei ? "#fff" : "#999",
                          fontSize: 11, fontWeight: emp.mei ? 700 : 400, fontFamily: "inherit", transition: "all 0.15s" }}>
                        {emp.mei ? "MEI ✓" : "MEI"}
                      </button>
                    </td>
                  )}
                  <td style={{ ...S.td, textAlign: "right", background: idx % 2 === 0 ? "#fff" : "#fafaf8" }}>
                    <span style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 15,
                      color: isHalf ? "#e67e22" : comm > 0 ? "#1B4332" : "#bbb" }}>
                      {fmt(comm)}
                      {isHalf && <span style={{ fontSize: 10, marginLeft: 4, color: "#f39c12" }}>50%</span>}
                      {emp.mei && <span style={{ fontSize: 9, marginLeft: 4, color: "#27ae60" }}>MEI</span>}
                    </span>
                  </td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr style={{ background: "#f0f5f0", borderTop: "2px solid #1B4332" }}>
              <td colSpan={isHistory ? 5 : 6} style={{ ...S.td, fontWeight: 600, color: "#1B4332", fontSize: 13 }}>Total {sector !== "Todos" ? "— " + sector : ""}</td>
              <td style={{ ...S.td, textAlign: "right" }}>
                <span style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 17, color: "#1B4332" }}>{fmt(totalComm)}</span>
              </td>
            </tr>
          </tfoot>
        </table>
      </div>

      {onBack && <div style={{ marginTop: 20 }}><button style={S.btnOut} onClick={onBack}>← Voltar</button></div>}
    </>
  );
}
