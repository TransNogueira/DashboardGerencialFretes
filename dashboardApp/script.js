// Cole aqui a URL da implantação do Web App do Google Apps Script
const API_URL = "https://script.google.com/macros/s/AKfycbxC0mKHIy5hprRsxMc0UyjyiXOV-vIwPVv4OmonpzlbCgE4b1cezCV_f2HRXbqdFDvTig/exec";
const STORAGE_KEY = "tn_dashboard_token";

let rawData = [];
let filteredData = [];
let charts = {};

const brl = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });

// Elementos da tela de autenticação
const authOverlay = document.getElementById("authOverlay");
const authForm = document.getElementById("authForm");
const tokenInput = document.getElementById("tokenInput");
const authBtn = document.getElementById("authBtn");
const authError = document.getElementById("authError");
const logoutBtn = document.getElementById("logoutBtn");

// Inicialização ao abrir a página
document.addEventListener("DOMContentLoaded", () => {
  const savedToken = sessionStorage.getItem(STORAGE_KEY);
  if (savedToken) {
    carregarDados(savedToken);
  }

  authForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const token = tokenInput.value.trim();
    if (token) carregarDados(token);
  });

  if (logoutBtn) {
    logoutBtn.addEventListener("click", () => {
      sessionStorage.removeItem(STORAGE_KEY);
      rawData = [];
      filteredData = [];
      tokenInput.value = "";
      authError.textContent = "";
      authOverlay.classList.remove("hidden");
      logoutBtn.style.display = "none";
    });
  }

  document.getElementById("apply").addEventListener("click", aplicarFiltros);
  document.getElementById("reset").addEventListener("click", limparFiltros);
  document.getElementById("export").addEventListener("click", exportarCSV);
});

async function carregarDados(token) {
  authBtn.disabled = true;
  authBtn.textContent = "Validando acesso...";
  authError.textContent = "";

  try {
    // Não envie headers customizados para evitar preflight CORS no Apps Script
    const response = await fetch(`${API_URL}?token=${encodeURIComponent(token)}`);
    const payload = await response.json();

    if (payload.erro) {
      sessionStorage.removeItem(STORAGE_KEY);
      authError.textContent = payload.erro;
      authOverlay.classList.remove("hidden");
      return;
    }

    sessionStorage.setItem(STORAGE_KEY, token);
    authOverlay.classList.add("hidden");
    if (logoutBtn) logoutBtn.style.display = "inline-block";

    rawData = payload.dados || [];
    popularFiltros(rawData);
    aplicarFiltros();
  } catch (err) {
    authError.textContent = "Falha de conexão com a planilha. Verifique a URL da API.";
    console.error(err);
  } finally {
    authBtn.disabled = false;
    authBtn.textContent = "Acessar Dashboard";
  }
}

function popularFiltros(dados) {
  const driverSelect = document.getElementById("driver");
  const typeSelect = document.getElementById("type");

  const motoristas = [...new Set(dados.map(d => d.motorista).filter(Boolean))].sort();
  const tipos = [...new Set(dados.map(d => d.tipoCte).filter(Boolean))].sort();

  driverSelect.innerHTML = `<option value="">Todos os motoristas</option>` +
    motoristas.map(m => `<option value="${m}">${m}</option>`).join("");

  typeSelect.innerHTML = `<option value="">Todos os tipos</option>` +
    tipos.map(t => `<option value="${t}">${t}</option>`).join("");
}

function aplicarFiltros() {
  const start = document.getElementById("start").value;
  const end = document.getElementById("end").value;
  const driver = document.getElementById("driver").value;
  const type = document.getElementById("type").value;

  filteredData = rawData.filter(item => {
    if (start && item.data < start) return false;
    if (end && item.data > end) return false;
    if (driver && item.motorista !== driver) return false;
    if (type && item.tipoCte !== type) return false;
    return true;
  });

  atualizarKPIs(filteredData);
  atualizarTabela(filteredData);
  atualizarGraficos(filteredData);
}

function limparFiltros() {
  document.getElementById("start").value = "";
  document.getElementById("end").value = "";
  document.getElementById("driver").value = "";
  document.getElementById("type").value = "";
  aplicarFiltros();
}

function atualizarKPIs(dados) {
  const totalFrete = dados.reduce((acc, d) => acc + d.freteCebrace, 0);
  const totalMotorista = dados.reduce((acc, d) => acc + d.freteMotorista, 0);
  const totalEnlonamento = dados.reduce((acc, d) => acc + d.enlonamento, 0);
  const resultado = totalFrete - totalMotorista - totalEnlonamento;

  document.getElementById("kFreight").textContent = brl.format(totalFrete);
  document.getElementById("kPay").textContent = brl.format(totalMotorista);
  document.getElementById("kTarp").textContent = brl.format(totalEnlonamento);
  document.getElementById("kResult").textContent = brl.format(resultado);
  document.getElementById("kDocs").textContent = dados.length.toLocaleString("pt-BR");
}

function agruparPorMotorista(dados) {
  const mapa = {};
  dados.forEach(d => {
    if (!mapa[d.motorista]) {
      mapa[d.motorista] = { motorista: d.motorista, docs: 0, frete: 0, enlonamento: 0, pago: 0 };
    }
    mapa[d.motorista].docs += 1;
    mapa[d.motorista].frete += d.freteCebrace;
    mapa[d.motorista].enlonamento += d.enlonamento;
    mapa[d.motorista].pago += d.freteMotorista;
  });

  return Object.values(mapa).map(m => {
    const resultado = m.frete - m.pago - m.enlonamento;
    const margem = m.frete > 0 ? (resultado / m.frete) * 100 : 0;
    return { ...m, resultado, margem };
  }).sort((a, b) => b.frete - a.frete);
}

function atualizarTabela(dados) {
  const resumo = agruparPorMotorista(dados);
  const tbody = document.getElementById("tbody");
  document.getElementById("rowCount").textContent = `${resumo.length} motorista(s) • ${dados.length} registro(s)`;

  tbody.innerHTML = resumo.map(r => `
    <tr>
      <td>${r.motorista}</td>
      <td class="num">${r.docs}</td>
      <td class="num">${brl.format(r.frete)}</td>
      <td class="num">${brl.format(r.enlonamento)}</td>
      <td class="num">${brl.format(r.pago)}</td>
      <td class="num">${brl.format(r.resultado)}</td>
      <td class="num">${r.margem.toFixed(1)}%</td>
    </tr>
  `).join("");
}

function renderChart(id, config) {
  if (charts[id]) charts[id].destroy();
  const ctx = document.getElementById(id).getContext("2d");
  charts[id] = new Chart(ctx, config);
}

function atualizarGraficos(dados) {
  const porMotorista = agruparPorMotorista(dados);

  // 1. Fretes e custos por motorista
  renderChart("driverChart", {
    type: "bar",
    data: {
      labels: porMotorista.map(m => m.motorista),
      datasets: [
        { label: "Frete Cebrace", data: porMotorista.map(m => m.frete), backgroundColor: "#2878d0" },
        { label: "Frete Motorista", data: porMotorista.map(m => m.pago), backgroundColor: "#e89a35" },
        { label: "Enlonamento", data: porMotorista.map(m => m.enlonamento), backgroundColor: "#19a58c" }
      ]
    },
    options: { responsive: true, maintainAspectRatio: false }
  });

  // 2. Composição financeira
  const totMotorista = dados.reduce((a, d) => a + d.freteMotorista, 0);
  const totEnlon = dados.reduce((a, d) => a + d.enlonamento, 0);
  const totResultado = Math.max(0, dados.reduce((a, d) => a + d.freteCebrace, 0) - totMotorista - totEnlon);

  renderChart("mixChart", {
    type: "doughnut",
    data: {
      labels: ["Frete Motorista", "Enlonamento", "Resultado Estimado"],
      datasets: [{
        data: [totMotorista, totEnlon, totResultado],
        backgroundColor: ["#e89a35", "#19a58c", "#142b4a"]
      }]
    },
    options: { responsive: true, maintainAspectRatio: false }
  });

  // 3. Evolução mensal
  const mesesMap = {};
  dados.forEach(d => {
    const mes = d.data ? d.data.slice(0, 7) : "Sem data";
    if (!mesesMap[mes]) mesesMap[mes] = { frete: 0, pago: 0, enlon: 0 };
    mesesMap[mes].frete += d.freteCebrace;
    mesesMap[mes].pago += d.freteMotorista;
    mesesMap[mes].enlon += d.enlonamento;
  });
  const meses = Object.keys(mesesMap).sort();

  renderChart("monthChart", {
    type: "line",
    data: {
      labels: meses,
      datasets: [
        { label: "Frete Cebrace", data: meses.map(m => mesesMap[m].frete), borderColor: "#2878d0", tension: 0.25 },
        { label: "Frete Motorista", data: meses.map(m => mesesMap[m].pago), borderColor: "#e89a35", tension: 0.25 },
        { label: "Enlonamento", data: meses.map(m => mesesMap[m].enlon), borderColor: "#19a58c", tension: 0.25 }
      ]
    },
    options: { responsive: true, maintainAspectRatio: false }
  });

  // 4. Ranking de resultado por motorista
  const ranking = [...porMotorista].sort((a, b) => b.resultado - a.resultado);
  renderChart("rankChart", {
    type: "bar",
    data: {
      labels: ranking.map(m => m.motorista),
      datasets: [{
        label: "Resultado Estimado (R$)",
        data: ranking.map(m => m.resultado),
        backgroundColor: "#142b4a"
      }]
    },
    options: { indexAxis: "y", responsive: true, maintainAspectRatio: false }
  });
}

function exportarCSV() {
  if (!filteredData.length) return;
  const cabecalho = ["Data", "Motorista", "Tipo CTE", "CTE", "Placa", "Frete Cebrace", "Enlonamento", "Frete Motorista"];
  const linhas = filteredData.map(d => [
    d.data, d.motorista, d.tipoCte, d.cte, d.placa,
    d.freteCebrace.toFixed(2).replace(".", ","),
    d.enlonamento.toFixed(2).replace(".", ","),
    d.freteMotorista.toFixed(2).replace(".", ",")
  ].join(";"));

  const csv = "\uFEFF" + [cabecalho.join(";"), ...linhas].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `fretes_transnogueira_${new Date().toISOString().slice(0, 10)}.csv`;
  link.click();
}