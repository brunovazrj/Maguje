import { useState, useMemo } from "react";

const TAX_RATE = 0.33;
const GARCOM_RATE = 0.29;

const INITIAL_EMPLOYEES = [
  // SALÃO - Equipe Global (Cumins/Suiteiro)
  { id: 1, name: "Jean Carlos Fidelis", role: "Cumim", sector: "Salão", type: "global", points: 15 },
  { id: 2, name: "Claudia Elisabete Conceição", role: "Cumim", sector: "Salão", type: "global", points: 15 },
  { id: 3, name: "Marcos Vinicius Henrique de Souza", role: "Cumim", sector: "Salão", type: "global", points: 15 },
  { id: 4, name: "Maria Elenice Ferreira", role: "Cumim", sector: "Salão", type: "global", points: 15 },
  { id: 5, name: "Rodrigo de Pinho Ribeiro", role: "Cumim", sector: "Salão", type: "global", points: 15 },
  { id: 6, name: "Felipe Costa de Abreu", role: "Suiteiro", sector: "Salão", type: "global", points: 20 },
  // SALÃO - Garçons (comissão individual)
  { id: 7, name: "Douglas Pereira Lima", role: "Garçom", sector: "Salão", type: "garcom", points: 0 },
  { id: 8, name: "Gabriel de Farias Pereira", role: "Garçom", sector: "Salão", type: "garcom", points: 0 },
  { id: 9, name: "Gustavo Fabricio Rodrigues Freire", role: "Garçom", sector: "Salão", type: "garcom", points: 0 },
  { id: 10, name: "Antonia Erineuda", role: "Garçonete", sector: "Salão", type: "garcom", points: 0 },
  { id: 11, name: "Paulo Alves de Almeida", role: "Garçom", sector: "Salão", type: "garcom", points: 0 },
  // SALÃO - Chefes e Gestão (global)
  { id: 12, name: "Reinaldo Alves de Oliveira", role: "Chefe de Fila Junior", sector: "Salão", type: "global", points: 24 },
  { id: 13, name: "Jeane Rodrigues", role: "Chefe de Fila Junior", sector: "Salão", type: "global", points: 24 },
  { id: 14, name: "Verinaldo Gabriel da Rocha", role: "Chefe de Fila Junior", sector: "Salão", type: "global", points: 24 },
  { id: 15, name: "Crislandia Moura de Lima", role: "Chefe de Fila Pleno", sector: "Salão", type: "global", points: 26 },
  { id: 16, name: "Elizangelo Araujo Miranda", role: "Maître", sector: "Salão", type: "global", points: 30 },
  { id: 17, name: "Joaquim Fernandes Gomes", role: "Assistente Gerente", sector: "Salão", type: "global", points: 25 },
  { id: 18, name: "Jose Edilson Pereira Nogueira", role: "Sub Gerente / Maître", sector: "Salão", type: "global", points: 30 },
  { id: 19, name: "Rodrigo Florentino Fonseca", role: "Gerente", sector: "Salão", type: "global", points: 35 },
  { id: 20, name: "Fabio da Silva Miguel", role: "Assistente MKT", sector: "Salão", type: "global", points: 15 },
  { id: 21, name: "Kayllana Vitoria de Oliveira Donato", role: "Hostess", sector: "Salão", type: "global", points: 15 },
  { id: 22, name: "Suzana Radai Estrela Souza", role: "Hostess", sector: "Salão", type: "global", points: 15 },
  { id: 23, name: "Romenia Fernades Jorge", role: "Hostess", sector: "Salão", type: "global", points: 20 },
  // BAR
  { id: 24, name: "Danilo Silva Gomes", role: "Barback", sector: "Bar", type: "global", points: 17 },
  { id: 25, name: "Luan Chrystyan dos Santos", role: "Barback", sector: "Bar", type: "global", points: 17 },
  { id: 26, name: "Francisco Tome da Silva", role: "Copeiro II", sector: "Bar", type: "global", points: 10 },
  { id: 27, name: "Antonio Mauricio Santos Soares", role: "Bartender", sector: "Bar", type: "global", points: 20 },
  { id: 28, name: "Rafael da Silva Romualdo", role: "Bartender", sector: "Bar", type: "global", points: 20 },
  { id: 29, name: "Caio Henriques Rodrigues", role: "Bartender", sector: "Bar", type: "global", points: 20 },
  { id: 30, name: "Gabriel Paulino Barbosa", role: "Bartender", sector: "Bar", type: "global", points: 20 },
  { id: 31, name: "Gabriel Soares Grativol", role: "Bartender", sector: "Bar", type: "global", points: 20 },
  { id: 32, name: "Gabriel de Oliveira Fernandes", role: "Sub Chefe de Bar", sector: "Bar", type: "global", points: 22 },
  { id: 33, name: "Luiz Gustavo Mesquita Soares", role: "Chefe de Bar", sector: "Bar", type: "global", points: 25 },
  // CAIXA
  { id: 34, name: "Antonia Jacilane de Sousa Costa", role: "Caixa", sector: "Caixa", type: "global", points: 15 },
  // COZINHA
  { id: 35, name: "Antonio Gomes de Sousa", role: "Copeiro", sector: "Cozinha", type: "global", points: 15 },
  { id: 36, name: "Douglas Leite Gonçalves", role: "Copeiro", sector: "Cozinha", type: "global", points: 15 },
  { id: 37, name: "John Victor Santos do Nascimento", role: "Copeiro", sector: "Cozinha", type: "global", points: 15 },
  { id: 38, name: "Rosangela Costa Rodrigues", role: "Copeiro", sector: "Cozinha", type: "global", points: 15 },
  { id: 39, name: "Robert Gustavo Santos de Souza", role: "Copeiro", sector: "Cozinha", type: "global", points: 15 },
  { id: 40, name: "Daniel Pereira do Sacramento", role: "Padeiro", sector: "Cozinha", type: "global", points: 15 },
  { id: 41, name: "Rosinaldo Pedro Soares", role: "Ajudante de Cozinha", sector: "Cozinha", type: "global", points: 15 },
  { id: 42, name: "Thaynara Tonelle Costa", role: "Cozinheiro I", sector: "Cozinha", type: "global", points: 15 },
  { id: 43, name: "Andre Felizardo Verissimo", role: "Cozinheiro I", sector: "Cozinha", type: "global", points: 15 },
  { id: 44, name: "Dayveson Rafael da Silva", role: "Cozinheiro I", sector: "Cozinha", type: "global", points: 15 },
  { id: 45, name: "Andriely Firmino da Silva", role: "Cozinheiro I", sector: "Cozinha", type: "global", points: 15 },
  { id: 46, name: "Vitor Faria de Oliveira Aguilera", role: "Cozinheiro I", sector: "Cozinha", type: "global", points: 15 },
  { id: 47, name: "Lucas Barbosa Ribeiro Borges", role: "Cozinheiro II", sector: "Cozinha", type: "global", points: 19 },
  { id: 48, name: "Robson Roberto da Silva", role: "Cozinheiro II", sector: "Cozinha", type: "global", points: 19 },
  { id: 49, name: "Wagner Pinto", role: "Cozinheiro III", sector: "Cozinha", type: "global", points: 21 },
  { id: 50, name: "Francisco Dalvan Bezerra Gomes", role: "Cozinheiro III", sector: "Cozinha", type: "global", points: 21 },
  { id: 51, name: "Valdemir Galdino de Oliveira", role: "Cozinheiro III", sector: "Cozinha", type: "global", points: 21 },
  { id: 52, name: "Luis Augusto Souza da Costa", role: "Cozinheiro Líder", sector: "Cozinha", type: "global", points: 23 },
  { id: 53, name: "Jaqueline de Souza Galvao", role: "Sub Chefe Cozinha", sector: "Cozinha", type: "global", points: 25 },
  { id: 54, name: "Eduardo", role: "Chef Produções Gast.", sector: "Cozinha", type: "global", points: 30 },
  // LIMPEZA
  { id: 55, name: "Alex dos Santos", role: "ASG", sector: "Limpeza", type: "global", points: 10 },
  { id: 56, name: "Carlos Daniel Alves de Lima", role: "ASG", sector: "Limpeza", type: "global", points: 10 },
  { id: 57, name: "Marlucia Santana Rodrigues", role: "Líder de ASG", sector: "Limpeza", type: "global", points: 20 },
  { id: 58, name: "José Roberto Inácio da Silva", role: "Estoquista", sector: "Limpeza", type: "global", points: 10 },
];

const SECTORS = ["Todos", "Salão", "Bar", "Caixa", "Cozinha", "Limpeza"];

const fmt = (v) => v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });

function initState(employees) {
  const sales = {};
  const absences = {};
  employees.forEach((e) => {
    sales[e.id] = e.type === "garcom" ? "" : "";
    absences[e.id] = 0;
  });
  return { sales, absences };
}

export default function App() {
  const [employees, setEmployees] = useState(INITIAL_EMPLOYEES);
  const [totalPool, setTotalPool] = useState("");
  const [inputs, setInputs] = useState(() => initState(INITIAL_EMPLOYEES));
  const [sector, setSector] = useState("Todos");
  const [step, setStep] = useState("input"); // "input" | "results"
  const [showAdd, setShowAdd] = useState(false);
  const [newEmp, setNewEmp] = useState({ name: "", role: "", sector: "Salão", type: "global", points: 15 });
  const [month, setMonth] = useState(() => {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  });

  const results = useMemo(() => {
    const pool = parseFloat(totalPool) || 0;
    const netPool = pool * (1 - TAX_RATE);

    // Garçons individual commission
    const garcomSales = {};
    let totalGarcomComm = 0;
    employees.forEach((e) => {
      if (e.type === "garcom") {
        const sale = parseFloat(inputs.sales[e.id]) || 0;
        const net = sale * (1 - TAX_RATE);
        const comm = net * GARCOM_RATE;
        garcomSales[e.id] = { sale, net, rawComm: comm };
        totalGarcomComm += comm;
      }
    });

    // Remaining for global pool
    const globalPool = netPool - totalGarcomComm;

    // Global pool: Grupo1 (Salão+Bar+Caixa) = 73%, Grupo2 (Cozinha+Limpeza) = 27%
    const GRUPO1_PCT = 0.73;
    const GRUPO2_PCT = 0.27;

    const grupo1Employees = employees.filter(
      (e) => e.type === "global" && ["Salão", "Bar", "Caixa"].includes(e.sector)
    );
    const grupo2Employees = employees.filter(
      (e) => e.type === "global" && ["Cozinha", "Limpeza"].includes(e.sector)
    );

    const grupo1Pool = globalPool * GRUPO1_PCT;
    const grupo2Pool = globalPool * GRUPO2_PCT;

    // Calculate effective points (deduct absences)
    const getEffectivePoints = (emp) => {
      const abs = parseInt(inputs.absences[emp.id]) || 0;
      if (abs <= 0) return emp.points;
      // Each absence deducts proportional share of their points
      const daysInMonth = 26; // working days approx
      const fraction = Math.max(0, 1 - abs / daysInMonth);
      return emp.points * fraction;
    };

    const calcGroupComm = (groupEmps, groupPool) => {
      const totalPts = groupEmps.reduce((s, e) => s + getEffectivePoints(e), 0);
      if (totalPts === 0) return {};
      const perPoint = groupPool / totalPts;
      const result = {};
      groupEmps.forEach((e) => {
        result[e.id] = getEffectivePoints(e) * perPoint;
      });
      return result;
    };

    const grupo1Comm = calcGroupComm(grupo1Employees, grupo1Pool);
    const grupo2Comm = calcGroupComm(grupo2Employees, grupo2Pool);

    const totals = {};
    employees.forEach((e) => {
      let comm = 0;
      if (e.type === "garcom") {
        const abs = parseInt(inputs.absences[e.id]) || 0;
        const daysInMonth = 26;
        const fraction = Math.max(0, 1 - abs / daysInMonth);
        comm = (garcomSales[e.id]?.rawComm || 0) * fraction;
      } else if (grupo1Comm[e.id] !== undefined) {
        comm = grupo1Comm[e.id];
      } else if (grupo2Comm[e.id] !== undefined) {
        comm = grupo2Comm[e.id];
      }
      totals[e.id] = comm;
    });

    const g1TotalPts = grupo1Employees.reduce((s, e) => s + getEffectivePoints(e), 0);
    const g2TotalPts = grupo2Employees.reduce((s, e) => s + getEffectivePoints(e), 0);
    const g1PerPt = g1TotalPts > 0 ? grupo1Pool / g1TotalPts : 0;
    const g2PerPt = g2TotalPts > 0 ? grupo2Pool / g2TotalPts : 0;

    return {
      pool,
      netPool,
      totalGarcomComm,
      globalPool,
      grupo1Pool,
      grupo2Pool,
      g1PerPt,
      g2PerPt,
      totals,
      garcomSales,
    };
  }, [employees, totalPool, inputs]);

  const filteredEmployees = sector === "Todos"
    ? employees
    : employees.filter((e) => e.sector === sector);

  const handleAddEmployee = () => {
    if (!newEmp.name.trim()) return;
    const id = Date.now();
    setEmployees((prev) => [...prev, { ...newEmp, id, points: parseInt(newEmp.points) || 15 }]);
    setInputs((prev) => ({
      sales: { ...prev.sales, [id]: "" },
      absences: { ...prev.absences, [id]: 0 },
    }));
    setNewEmp({ name: "", role: "", sector: "Salão", type: "global", points: 15 });
    setShowAdd(false);
  };

  const handleRemoveEmployee = (id) => {
    setEmployees((prev) => prev.filter((e) => e.id !== id));
  };

  const sectorColors = {
    Salão: "#2D6A4F",
    Bar: "#1B4332",
    Caixa: "#40916C",
    Cozinha: "#B5450B",
    Limpeza: "#7B5EA7",
  };

  const [d] = month.split("-");
  const monthLabel = new Date(month + "-02").toLocaleString("pt-BR", { month: "long", year: "numeric" });

  return (
    <div style={{ fontFamily: "'DM Mono', 'Courier New', monospace", background: "#F5F0E8", minHeight: "100vh", padding: "0" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Space+Grotesk:wght@400;500;700&display=swap');
        * { box-sizing: border-box; }
        .card { background: #fff; border: 1.5px solid #1B4332; border-radius: 4px; }
        input[type=number], input[type=text], select { 
          background: #F5F0E8; border: 1px solid #ccc; border-radius: 3px;
          padding: 6px 10px; font-family: inherit; font-size: 13px; width: 100%;
          transition: border-color 0.15s;
        }
        input[type=number]:focus, input[type=text]:focus, select:focus {
          outline: none; border-color: #2D6A4F;
        }
        .btn { 
          border: 1.5px solid #1B4332; background: #1B4332; color: #fff;
          padding: 8px 18px; border-radius: 3px; cursor: pointer; font-family: inherit;
          font-size: 13px; font-weight: 500; transition: all 0.15s; letter-spacing: 0.02em;
        }
        .btn:hover { background: #2D6A4F; }
        .btn-outline { background: transparent; color: #1B4332; }
        .btn-outline:hover { background: #e8f0eb; }
        .btn-danger { border-color: #c0392b; background: transparent; color: #c0392b; }
        .btn-danger:hover { background: #fdf2f2; }
        .tab { padding: 7px 16px; border: 1px solid #ccc; border-radius: 20px; cursor: pointer;
          font-size: 12px; background: transparent; font-family: inherit; transition: all 0.15s; }
        .tab.active { background: #1B4332; color: #fff; border-color: #1B4332; }
        .tag { display: inline-block; font-size: 10px; padding: 2px 8px; border-radius: 20px; 
          font-weight: 500; letter-spacing: 0.05em; text-transform: uppercase; }
        .row-hover:hover { background: #f9f7f2; }
        .metric { text-align: center; }
        .metric-val { font-size: 22px; font-weight: 500; font-family: 'Space Grotesk', sans-serif; }
        .metric-lbl { font-size: 10px; color: #666; text-transform: uppercase; letter-spacing: 0.08em; margin-top: 2px; }
      `}</style>

      {/* Header */}
      <div style={{ background: "#1B4332", padding: "20px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ color: "#52B788", fontSize: 10, letterSpacing: "0.15em", textTransform: "uppercase", marginBottom: 4 }}>Restaurante Maguje</div>
          <div style={{ color: "#fff", fontSize: 22, fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700 }}>
            Calculadora de Comissões
          </div>
          <div style={{ color: "#95D5B2", fontSize: 12, marginTop: 2, textTransform: "capitalize" }}>{monthLabel}</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button className="btn btn-outline" style={{ color: "#95D5B2", borderColor: "#52B788" }}
            onClick={() => setStep(step === "input" ? "results" : "input")}>
            {step === "input" ? "Ver Resultado →" : "← Editar Dados"}
          </button>
        </div>
      </div>

      {step === "input" && (
        <div style={{ padding: "20px 28px" }}>
          {/* Month + Pool config */}
          <div className="card" style={{ padding: "18px 20px", marginBottom: 20 }}>
            <div style={{ fontSize: 11, color: "#666", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 14 }}>Configuração do Mês</div>
            <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
              <div style={{ flex: 1, minWidth: 160 }}>
                <div style={{ fontSize: 11, color: "#555", marginBottom: 5 }}>Mês de referência</div>
                <input type="month" value={month} onChange={(e) => setMonth(e.target.value)} style={{ background: "#F5F0E8", border: "1px solid #ccc", borderRadius: 3, padding: "7px 10px", fontFamily: "inherit", fontSize: 13, width: "100%" }} />
              </div>
              <div style={{ flex: 1, minWidth: 220 }}>
                <div style={{ fontSize: 11, color: "#555", marginBottom: 5 }}>Gorjeta Total Bruta (pool geral — equipe não-garçons)</div>
                <input type="number" placeholder="Ex: 53152.69" value={totalPool}
                  onChange={(e) => setTotalPool(e.target.value)}
                  style={{ background: "#F5F0E8", border: "1px solid #ccc", borderRadius: 3, padding: "7px 10px", fontFamily: "inherit", fontSize: 13, width: "100%" }} />
              </div>
            </div>
            {totalPool && (
              <div style={{ marginTop: 14, display: "flex", gap: 24, flexWrap: "wrap" }}>
                <div className="metric"><div className="metric-val" style={{ color: "#1B4332" }}>{fmt(parseFloat(totalPool) || 0)}</div><div className="metric-lbl">Bruto Total</div></div>
                <div className="metric"><div className="metric-val" style={{ color: "#40916C" }}>{fmt(results.netPool)}</div><div className="metric-lbl">Líquido (−33%)</div></div>
                <div className="metric"><div className="metric-val" style={{ color: "#52B788" }}>{fmt(results.grupo1Pool)}</div><div className="metric-lbl">Pool Salão/Bar/Cx (73%)</div></div>
                <div className="metric"><div className="metric-val" style={{ color: "#B5450B" }}>{fmt(results.grupo2Pool)}</div><div className="metric-lbl">Pool Cozinha/Limp. (27%)</div></div>
              </div>
            )}
          </div>

          {/* Sector filter */}
          <div style={{ display: "flex", gap: 6, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
            {SECTORS.map((s) => (
              <button key={s} className={`tab${sector === s ? " active" : ""}`} onClick={() => setSector(s)}>{s}</button>
            ))}
            <div style={{ marginLeft: "auto" }}>
              <button className="btn" style={{ fontSize: 12, padding: "6px 14px" }} onClick={() => setShowAdd(!showAdd)}>
                {showAdd ? "✕ Cancelar" : "+ Funcionário"}
              </button>
            </div>
          </div>

          {/* Add employee form */}
          {showAdd && (
            <div className="card" style={{ padding: "16px 20px", marginBottom: 16, borderColor: "#52B788" }}>
              <div style={{ fontSize: 11, color: "#1B4332", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 12 }}>Novo Funcionário</div>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <div style={{ flex: 2, minWidth: 160 }}>
                  <div style={{ fontSize: 11, color: "#555", marginBottom: 4 }}>Nome</div>
                  <input type="text" placeholder="Nome completo" value={newEmp.name} onChange={(e) => setNewEmp((p) => ({ ...p, name: e.target.value }))} />
                </div>
                <div style={{ flex: 1, minWidth: 120 }}>
                  <div style={{ fontSize: 11, color: "#555", marginBottom: 4 }}>Cargo</div>
                  <input type="text" placeholder="Cargo" value={newEmp.role} onChange={(e) => setNewEmp((p) => ({ ...p, role: e.target.value }))} />
                </div>
                <div style={{ flex: 1, minWidth: 100 }}>
                  <div style={{ fontSize: 11, color: "#555", marginBottom: 4 }}>Setor</div>
                  <select value={newEmp.sector} onChange={(e) => setNewEmp((p) => ({ ...p, sector: e.target.value }))}>
                    {["Salão", "Bar", "Caixa", "Cozinha", "Limpeza"].map((s) => <option key={s}>{s}</option>)}
                  </select>
                </div>
                <div style={{ flex: 1, minWidth: 110 }}>
                  <div style={{ fontSize: 11, color: "#555", marginBottom: 4 }}>Tipo</div>
                  <select value={newEmp.type} onChange={(e) => setNewEmp((p) => ({ ...p, type: e.target.value }))}>
                    <option value="global">Equipe Global</option>
                    <option value="garcom">Garçom (individual)</option>
                  </select>
                </div>
                {newEmp.type === "global" && (
                  <div style={{ flex: 1, minWidth: 80 }}>
                    <div style={{ fontSize: 11, color: "#555", marginBottom: 4 }}>Pontos</div>
                    <input type="number" min="1" max="50" value={newEmp.points} onChange={(e) => setNewEmp((p) => ({ ...p, points: e.target.value }))} />
                  </div>
                )}
                <div style={{ display: "flex", alignItems: "flex-end" }}>
                  <button className="btn" onClick={handleAddEmployee}>Adicionar</button>
                </div>
              </div>
            </div>
          )}

          {/* Employee table */}
          <div className="card" style={{ overflow: "hidden" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr style={{ background: "#1B4332", color: "#fff" }}>
                  <th style={{ padding: "10px 14px", textAlign: "left", fontSize: 11, letterSpacing: "0.06em", textTransform: "uppercase" }}>Funcionário</th>
                  <th style={{ padding: "10px 14px", textAlign: "left", fontSize: 11, letterSpacing: "0.06em", textTransform: "uppercase" }}>Cargo / Pontos</th>
                  <th style={{ padding: "10px 14px", textAlign: "left", fontSize: 11, letterSpacing: "0.06em", textTransform: "uppercase", minWidth: 150 }}>Venda (Garçons)</th>
                  <th style={{ padding: "10px 14px", textAlign: "center", fontSize: 11, letterSpacing: "0.06em", textTransform: "uppercase", minWidth: 80 }}>Faltas</th>
                  <th style={{ padding: "10px 6px", width: 36 }}></th>
                </tr>
              </thead>
              <tbody>
                {filteredEmployees.map((emp, idx) => {
                  const color = sectorColors[emp.sector] || "#555";
                  return (
                    <tr key={emp.id} className="row-hover" style={{ borderBottom: "1px solid #eee", background: idx % 2 === 0 ? "#fff" : "#fafaf8" }}>
                      <td style={{ padding: "10px 14px" }}>
                        <div style={{ fontSize: 13, fontWeight: 500 }}>{emp.name}</div>
                        <span className="tag" style={{ background: color + "18", color, border: `1px solid ${color}40`, marginTop: 3 }}>
                          {emp.sector}
                        </span>
                      </td>
                      <td style={{ padding: "10px 14px" }}>
                        <div style={{ fontSize: 12, color: "#444" }}>{emp.role}</div>
                        {emp.type === "global" ? (
                          <div style={{ fontSize: 11, color: "#888", marginTop: 2 }}>{emp.points} pts</div>
                        ) : (
                          <div style={{ fontSize: 11, color: "#B5450B", marginTop: 2 }}>29% da venda</div>
                        )}
                      </td>
                      <td style={{ padding: "8px 14px" }}>
                        {emp.type === "garcom" ? (
                          <input type="number" min="0" placeholder="R$ 0,00"
                            value={inputs.sales[emp.id]}
                            onChange={(e) => setInputs((p) => ({ ...p, sales: { ...p.sales, [emp.id]: e.target.value } }))}
                          />
                        ) : (
                          <span style={{ fontSize: 12, color: "#aaa" }}>—</span>
                        )}
                      </td>
                      <td style={{ padding: "8px 14px", textAlign: "center" }}>
                        <input type="number" min="0" max="31"
                          value={inputs.absences[emp.id]}
                          onChange={(e) => setInputs((p) => ({ ...p, absences: { ...p.absences, [emp.id]: e.target.value } }))}
                          style={{ width: 60, textAlign: "center" }}
                        />
                      </td>
                      <td style={{ padding: "8px 6px", textAlign: "center" }}>
                        <button className="btn btn-danger" style={{ padding: "3px 8px", fontSize: 11, border: "none", background: "transparent", color: "#ccc", cursor: "pointer" }}
                          onClick={() => handleRemoveEmployee(emp.id)} title="Remover">✕</button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div style={{ marginTop: 20, textAlign: "right" }}>
            <button className="btn" style={{ fontSize: 14, padding: "10px 28px" }} onClick={() => setStep("results")}>
              Calcular Comissões →
            </button>
          </div>
        </div>
      )}

      {step === "results" && (
        <div style={{ padding: "20px 28px" }}>
          {/* Summary cards */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px,1fr))", gap: 12, marginBottom: 24 }}>
            {[
              { label: "Total Bruto", val: results.pool, color: "#1B4332" },
              { label: "Desconto 33%", val: results.pool - results.netPool, color: "#c0392b" },
              { label: "Total Líquido", val: results.netPool, color: "#2D6A4F" },
              { label: "Pool Salão/Bar/Cx", val: results.grupo1Pool, color: "#40916C" },
              { label: "Pool Cozinha/Limp.", val: results.grupo2Pool, color: "#B5450B" },
              { label: "Garçons (individual)", val: results.totalGarcomComm, color: "#7B5EA7" },
            ].map((m) => (
              <div key={m.label} className="card" style={{ padding: "14px 16px", textAlign: "center" }}>
                <div style={{ fontSize: 18, fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, color: m.color }}>{fmt(m.val)}</div>
                <div style={{ fontSize: 10, color: "#888", textTransform: "uppercase", letterSpacing: "0.07em", marginTop: 3 }}>{m.label}</div>
              </div>
            ))}
          </div>

          {/* Sector filter */}
          <div style={{ display: "flex", gap: 6, marginBottom: 16, flexWrap: "wrap" }}>
            {SECTORS.map((s) => (
              <button key={s} className={`tab${sector === s ? " active" : ""}`} onClick={() => setSector(s)}>{s}</button>
            ))}
          </div>

          {/* Results table */}
          <div className="card" style={{ overflow: "hidden" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr style={{ background: "#1B4332", color: "#fff" }}>
                  <th style={{ padding: "10px 14px", textAlign: "left", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em" }}>Funcionário</th>
                  <th style={{ padding: "10px 14px", textAlign: "left", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em" }}>Cargo</th>
                  <th style={{ padding: "10px 14px", textAlign: "center", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em" }}>Tipo</th>
                  <th style={{ padding: "10px 14px", textAlign: "center", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em" }}>Pts / Venda</th>
                  <th style={{ padding: "10px 14px", textAlign: "center", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em" }}>Faltas</th>
                  <th style={{ padding: "10px 14px", textAlign: "right", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em" }}>Comissão</th>
                </tr>
              </thead>
              <tbody>
                {filteredEmployees
                  .slice()
                  .sort((a, b) => (results.totals[b.id] || 0) - (results.totals[a.id] || 0))
                  .map((emp, idx) => {
                    const comm = results.totals[emp.id] || 0;
                    const color = sectorColors[emp.sector] || "#555";
                    const abs = parseInt(inputs.absences[emp.id]) || 0;
                    return (
                      <tr key={emp.id} className="row-hover" style={{ borderBottom: "1px solid #eee", background: idx % 2 === 0 ? "#fff" : "#fafaf8" }}>
                        <td style={{ padding: "10px 14px" }}>
                          <div style={{ fontSize: 13, fontWeight: 500 }}>{emp.name}</div>
                          <span className="tag" style={{ background: color + "18", color, border: `1px solid ${color}40`, marginTop: 3 }}>
                            {emp.sector}
                          </span>
                        </td>
                        <td style={{ padding: "10px 14px", fontSize: 12, color: "#444" }}>{emp.role}</td>
                        <td style={{ padding: "10px 14px", textAlign: "center" }}>
                          {emp.type === "garcom"
                            ? <span className="tag" style={{ background: "#7B5EA720", color: "#7B5EA7", border: "1px solid #7B5EA740" }}>Individual</span>
                            : <span className="tag" style={{ background: "#2D6A4F20", color: "#2D6A4F", border: "1px solid #2D6A4F40" }}>Global</span>}
                        </td>
                        <td style={{ padding: "10px 14px", textAlign: "center", fontSize: 12, color: "#555" }}>
                          {emp.type === "garcom"
                            ? fmt(parseFloat(inputs.sales[emp.id]) || 0)
                            : `${emp.points} pts`}
                        </td>
                        <td style={{ padding: "10px 14px", textAlign: "center" }}>
                          {abs > 0
                            ? <span style={{ background: "#fdecea", color: "#c0392b", borderRadius: 10, padding: "2px 8px", fontSize: 12 }}>{abs} falta{abs > 1 ? "s" : ""}</span>
                            : <span style={{ color: "#aaa", fontSize: 12 }}>—</span>}
                        </td>
                        <td style={{ padding: "10px 14px", textAlign: "right" }}>
                          <span style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 15, color: comm > 0 ? "#1B4332" : "#bbb" }}>
                            {fmt(comm)}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
              </tbody>
              <tfoot>
                <tr style={{ background: "#f0f5f0", borderTop: "2px solid #1B4332" }}>
                  <td colSpan={5} style={{ padding: "12px 14px", fontSize: 13, fontWeight: 600, color: "#1B4332" }}>
                    Total {sector !== "Todos" ? `— ${sector}` : ""}
                  </td>
                  <td style={{ padding: "12px 14px", textAlign: "right" }}>
                    <span style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 17, color: "#1B4332" }}>
                      {fmt(filteredEmployees.reduce((s, e) => s + (results.totals[e.id] || 0), 0))}
                    </span>
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>

          {/* Point value reference */}
          <div style={{ marginTop: 20, display: "flex", gap: 12, flexWrap: "wrap" }}>
            <div className="card" style={{ padding: "12px 18px", flex: 1, minWidth: 200 }}>
              <div style={{ fontSize: 10, color: "#666", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Valor do Ponto — Salão / Bar / Caixa</div>
              <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 20, fontWeight: 700, color: "#2D6A4F" }}>{fmt(results.g1PerPt)}</div>
              <div style={{ fontSize: 11, color: "#888", marginTop: 2 }}>por ponto</div>
            </div>
            <div className="card" style={{ padding: "12px 18px", flex: 1, minWidth: 200 }}>
              <div style={{ fontSize: 10, color: "#666", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Valor do Ponto — Cozinha / Limpeza</div>
              <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 20, fontWeight: 700, color: "#B5450B" }}>{fmt(results.g2PerPt)}</div>
              <div style={{ fontSize: 11, color: "#888", marginTop: 2 }}>por ponto</div>
            </div>
          </div>

          <div style={{ marginTop: 20, textAlign: "right" }}>
            <button className="btn btn-outline" onClick={() => setStep("input")}>← Voltar e Editar</button>
          </div>
        </div>
      )}

      <div style={{ textAlign: "center", padding: "16px", fontSize: 11, color: "#aaa", borderTop: "1px solid #e5e0d6", marginTop: 20 }}>
        Maguje · Sistema de Comissões · {monthLabel}
      </div>
    </div>
  );
}
