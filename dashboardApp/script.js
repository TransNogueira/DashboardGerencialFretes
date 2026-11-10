// ============================================================================
// 1. CONFIGURAÇÃO DO SUPABASE NO FRONT-END
// ============================================================================
const SUPABASE_URL = 'https://jydebhtmfzfpudgsvwka.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_5wbjgQTWi_U7jKmJhKI57w_VdXuYoeE'; 

const _supabase = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

let DATA = [];
let charts = {};
const formatadorMoeda = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
const formatadorNum = new Intl.NumberFormat('pt-BR');

// ============================================================================
// 2. INICIALIZAÇÃO AUTOMÁTICA DA PÁGINA
// ============================================================================
document.addEventListener('DOMContentLoaded', async () => {
    document.getElementById('apply').addEventListener('click', atualizarDashboard);
    document.getElementById('reset').addEventListener('click', limparFiltros);
    document.getElementById('export').addEventListener('click', exportarCSV);

    await buscarDadosDoBanco();
});

async function buscarDadosDoBanco() {
    console.log("🔄 Buscando dados da tabela 'fretes' no Supabase...");
    
    const { data, error } = await _supabase.from('fretes').select('*');

    if (error) {
        console.error("❌ Erro ao carregar fretes:", error.message);
        return;
    }

    if (data && data.length > 0) {
        console.log(`✅ ${data.length} registros carregados com sucesso!`);
        
        DATA = data.map(item => {
            const frete = Number(item.frete_cebrace) || 0;
            const motoristaPay = Number(item.frete_motorista) || 0;
            const enlonamento = Number(item.enlonamento) || 0;
            const resultadoEst = frete - motoristaPay - enlonamento;

            return {
                date: item.data_registro || '',
                driver: item.motorista || 'NÃO INFORMADO',
                type: item.tipo_cte || 'EMBARQUE',
                freight: frete,
                tarp: enlonamento,
                driverPay: motoristaPay,
                result: resultadoEst
            };
        });

        inicializarFiltros();
        atualizarDashboard();
    } else {
        console.warn("⚠️ A tabela 'fretes' retornou vazia.");
    }
}

// ============================================================================
// 3. IMPORTAÇÃO INTELIGENTE DE 2 PLANILHAS SIMULTÂNEAS
// ============================================================================
async function importarPlanilhasExcel(event) {
    const arquivos = Array.from(event.target.files);
    if (arquivos.length === 0) return;

    alert(`Carregando ${arquivos.length} arquivo(s). Aguarde o processamento...`);

    try {
        let dadosBrutosBruta = [];
        let dadosBrutosMotorista = [];

        // Lê os arquivos selecionados e classifica cada um
        for (const arq of arquivos) {
            const buffer = await arq.arrayBuffer();
            const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
            const primeiraAba = workbook.SheetNames[0];
            const linhas = XLSX.utils.sheet_to_json(workbook.Sheets[primeiraAba]);

            if (linhas.length === 0) continue;

            const primeiraLinha = linhas[0];
            
            // Verifica se é a planilha de frete de motorista (possui 'Valor Agregado Bruto')
            if ('Valor Agregado Bruto' in primeiraLinha || 'valor_agregado_bruto' in primeiraLinha) {
                dadosBrutosMotorista = linhas;
            } else {
                // Caso contrário, trata como a planilha bruta (contém 'Tipo Movimento' / 'Número Manifesto')
                dadosBrutosBruta = linhas;
            }
        }

        // Se o usuário selecionou apenas 1 planilha e ela era a de motorista, processamos com padrão
        if (dadosBrutosMotorista.length === 0 && dadosBrutosBruta.length > 0) {
            alert("Selecione também o arquivo 'frete_motorista.xls' para podermos processar os valores dos agregados.");
            return;
        }

        // 1. Mapeia a planilha bruta (se fornecida) por Número Manifesto para buscar o Tipo de CT-e
        const mapaTipoCte = {};
        if (dadosBrutosBruta.length > 0) {
            dadosBrutosBruta.forEach(b => {
                const numManifesto = String(b['Número Manifesto'] || '').trim();
                const tipoMovimento = b['Tipo Movimento'] || b['Tipo Serviço'] || 'EMBARQUE';
                if (numManifesto) {
                    mapaTipoCte[numManifesto] = tipoMovimento;
                }
            });
        }

        // 2. Processa os dados da planilha de motoristas
        const dadosLimpos = [];

        for (const linha of dadosBrutosMotorista) {
            const motorista = obterValor(linha, ['Motorista', 'MOTORISTA', 'motorista']);
            if (!motorista) continue;

            const nomeMotoristaStr = String(motorista).trim().toUpperCase();

            // Filtro de agregados (Valor Agregado Bruto > 0)
            const valorAgregadoBrutoRaw = obterValor(linha, ['Valor Agregado Bruto', 'valor_agregado_bruto'], 0.0);
            const freteMotorista = limparMoeda(valorAgregadoBrutoRaw);
            if (freteMotorista <= 0) continue; // Pula frota própria

            // Chave de cruzamento pelo número do manifesto/contrato
            const numManifestoStr = String(obterValor(linha, ['Número', 'Contrato', 'Pré Manifesto'], '')).trim();
            const tipoCte = mapaTipoCte[numManifestoStr] || obterValor(linha, ['Tipo Movimento', 'Tipo Serviço'], 'EMBARQUE');

            const rawData = obterValor(linha, ['Data', 'Emissão CT-e', 'Coleta', 'Data Fechamento']);
            const cte = obterValor(linha, ['Contrato', 'Número', 'CTE'], '');
            const placa = obterValor(linha, ['Veículo', 'Placa Veículo', 'Placa'], '');
            
            // Frete Cebrace exclusivamente de Total Frete
            const freteCebraceRaw = obterValor(linha, ['Total Frete', 'total_frete'], 0.0);

            // Regra de enlonamento: Fixo 118,44, exceto Julio Edson Cavinatti (0)
            let enlonamento = 118.44;
            if (nomeMotoristaStr.includes('JULIO EDSON CAVINATTI') || nomeMotoristaStr.includes('JULIO EDSON CAVINATI')) {
                enlonamento = 0.0;
            }

            dadosLimpos.push({
                data_registro: formatarDataExcel(rawData),
                motorista: nomeMotoristaStr,
                tipo_cte: String(tipoCte).trim(),
                cte: String(cte).trim(),
                placa: String(placa).trim().toUpperCase(),
                frete_cebrace: limparMoeda(freteCebraceRaw),
                enlonamento: enlonamento,
                frete_motorista: freteMotorista
            });
        }

        if (dadosLimpos.length === 0) {
            alert("Nenhum motorista agregado válido encontrado nas planilhas.");
            return;
        }

        console.log(`🚀 Sincronizando ${dadosLimpos.length} registros no Supabase...`);

        // Limpa registros anteriores e insere os novos consolidados
        await _supabase.from('fretes').delete().neq('id', 0);
        const { error } = await _supabase.from('fretes').insert(dadosLimpos);

        if (error) {
            alert("Erro ao salvar dados no Supabase: " + error.message);
            return;
        }

        alert(`Sucesso! ${dadosLimpos.length} registros de agregados foram importados e sincronizados.`);
        await buscarDadosDoBanco(); // Recarrega os gráficos na tela automaticamente

    } catch (err) {
        console.error("Erro na leitura das planilhas:", err);
        alert("Erro ao processar as planilhas. Verifique se selecionou os arquivos corretos.");
    }
}

// Funções Auxiliares
function obterValor(linha, chavesPossiveis, valorPadrao = null) {
    for (const chave of chavesPossiveis) {
        if (linha[chave] !== undefined && linha[chave] !== null && String(linha[chave]).trim() !== '') {
            return linha[chave];
        }
    }
    return valorPadrao;
}

function limparMoeda(valor) {
    if (valor === null || valor === undefined) return 0.0;
    if (typeof valor === 'number') return valor;
    let strValor = String(valor).replace(/R\$/g, '').replace(/\./g, '').replace(/,/g, '.').trim();
    return parseFloat(strValor) || 0.0;
}

function formatarDataExcel(valorData) {
    if (!valorData) return null;
    if (typeof valorData === 'string') {
        const partes = valorData.trim().split('/');
        if (partes.length === 3) {
            return `${partes[2]}-${partes[1].padStart(2, '0')}-${partes[0].padStart(2, '0')}`;
        }
    } else if (typeof valorData === 'number') {
        const dataObj = XLSX.SSF.parse_date_code(valorData);
        if (dataObj) {
            return `${dataObj.y}-${String(dataObj.m).padStart(2, '0')}-${String(dataObj.d).padStart(2, '0')}`;
        }
    }
    return null;
}

// ============================================================================
// 4. LÓGICA DE FILTROS E PREENCHIMENTO DOS SELECTS
// ============================================================================
function inicializarFiltros() {
    const drivers = [...new Set(DATA.map(d => d.driver))].sort();
    const types = [...new Set(DATA.map(d => d.type))].sort();

    const driverSelect = document.getElementById('driver');
    driverSelect.innerHTML = '<option value="">Todos os motoristas</option>';
    drivers.forEach(d => driverSelect.add(new Option(d, d)));

    const typeSelect = document.getElementById('type');
    typeSelect.innerHTML = '<option value="">Todos os tipos</option>';
    types.forEach(t => typeSelect.add(new Option(t, t)));
}

function limparFiltros() {
    document.getElementById('start').value = '';
    document.getElementById('end').value = '';
    document.getElementById('driver').value = '';
    document.getElementById('type').value = '';
    atualizarDashboard();
}

function getDadosFiltrados() {
    const start = document.getElementById('start').value;
    const end = document.getElementById('end').value;
    const driver = document.getElementById('driver').value;
    const type = document.getElementById('type').value;

    return DATA.filter(d => {
        let ok = true;
        if (start) ok = ok && d.date >= start;
        if (end) ok = ok && d.date <= end;
        if (driver) ok = ok && d.driver === driver;
        if (type) ok = ok && d.type === type;
        return ok;
    });
}

// ============================================================================
// 5. ATUALIZAÇÃO GERAL DO DASHBOARD (MÉTRICAS, TABELA E GRÁFICOS)
// ============================================================================
function atualizarDashboard() {
    const filtrados = getDadosFiltrados();

    let totalFreight = 0, totalPay = 0, totalTarp = 0, totalResult = 0;
    
    filtrados.forEach(d => {
        totalFreight += d.freight;
        totalPay += d.driverPay;
        totalTarp += d.tarp;
        totalResult += d.result;
    });

    document.getElementById('kFreight').innerText = formatadorMoeda.format(totalFreight);
    document.getElementById('kPay').innerText = formatadorMoeda.format(totalPay);
    document.getElementById('kTarp').innerText = formatadorMoeda.format(totalTarp);
    document.getElementById('kResult').innerText = formatadorMoeda.format(totalResult);
    document.getElementById('kDocs').innerText = formatadorNum.format(filtrados.length);

    const resumoMotoristas = agruparPorMotorista(filtrados);
    const evolucaoMensal = agruparPorMes(filtrados);

    renderizarTabela(resumoMotoristas, filtrados.length);
    renderizarGraficos(resumoMotoristas, evolucaoMensal, totalPay, totalTarp, totalResult);
}

// ============================================================================
// 6. AGRUPAMENTO DE DADOS (CÁLCULOS)
// ============================================================================
function agruparPorMotorista(dados) {
    const map = {};
    dados.forEach(d => {
        if (!map[d.driver]) {
            map[d.driver] = { driver: d.driver, docs: 0, freight: 0, tarp: 0, pay: 0, result: 0 };
        }
        map[d.driver].docs += 1;
        map[d.driver].freight += d.freight;
        map[d.driver].tarp += d.tarp;
        map[d.driver].pay += d.driverPay;
        map[d.driver].result += d.result;
    });

    return Object.values(map).sort((a, b) => b.result - a.result);
}

function agruparPorMes(dados) {
    const map = {};
    dados.forEach(d => {
        if (!d.date) return;
        const mes = d.date.substring(0, 7);
        if (!map[mes]) {
            map[mes] = { mes: mes, freight: 0, pay: 0, tarp: 0, result: 0 };
        }
        map[mes].freight += d.freight;
        map[mes].pay += d.driverPay;
        map[mes].tarp += d.tarp;
        map[mes].result += d.result;
    });

    return Object.values(map).sort((a, b) => a.mes.localeCompare(b.mes));
}

// ============================================================================
// 7. RENDERIZAÇÃO DA TABELA
// ============================================================================
function renderizarTabela(resumoMotoristas, totalRegistros) {
    const tbody = document.getElementById('tbody');
    tbody.innerHTML = '';
    
    document.getElementById('rowCount').innerText = `${formatadorNum.format(totalRegistros)} registros • ${resumoMotoristas.length} motoristas`;

    resumoMotoristas.forEach(r => {
        const margem = r.freight > 0 ? ((r.result / r.freight) * 100).toFixed(1) + '%' : '0%';
        const corResultado = r.result >= 0 ? 'positive' : 'negative';

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${r.driver}</td>
            <td class="num">${r.docs}</td>
            <td class="num">${formatadorMoeda.format(r.freight)}</td>
            <td class="num">${formatadorMoeda.format(r.tarp)}</td>
            <td class="num">${formatadorMoeda.format(r.pay)}</td>
            <td class="num ${corResultado}">${formatadorMoeda.format(r.result)}</td>
            <td class="num">${margem}</td>
        `;
        tbody.appendChild(tr);
    });
}

// ============================================================================
// 8. RENDERIZAÇÃO DOS GRÁFICOS (CHART.JS)
// ============================================================================
function destruirGrafico(id) {
    if (charts[id]) {
        charts[id].destroy();
    }
}

function renderizarGraficos(resumoMotoristas, evolucaoMensal, totalPay, totalTarp, totalResult) {
    if (typeof Chart === 'undefined') return;

    Chart.defaults.font.family = 'Inter, sans-serif';
    Chart.defaults.color = '#6b7788';

    const top10 = resumoMotoristas.slice(0, 10);
    const driverLabels = top10.map(r => r.driver ? r.driver.split(' ').slice(0, 2).join(' ') : '');

    // --- GRÁFICO 1: Fretes e Custos por Motorista (Barra) ---
    destruirGrafico('driverChart');
    const ctxDriver = document.getElementById('driverChart').getContext('2d');
    charts['driverChart'] = new Chart(ctxDriver, {
        type: 'bar',
        data: {
            labels: driverLabels,
            datasets: [
                { label: 'Frete Cebrace', data: top10.map(r => r.freight), backgroundColor: '#2878d0', borderRadius: 4 },
                { label: 'Frete Motorista', data: top10.map(r => r.pay), backgroundColor: '#e89a35', borderRadius: 4 }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { position: 'top' } },
            scales: { y: { beginAtZero: true } }
        }
    });

    // --- GRÁFICO 2: Composição Financeira (Doughnut) ---
    destruirGrafico('mixChart');
    const ctxMix = document.getElementById('mixChart').getContext('2d');
    charts['mixChart'] = new Chart(ctxMix, {
        type: 'doughnut',
        data: {
            labels: ['Frete Motorista', 'Enlonamento', 'Resultado Estimado'],
            datasets: [{
                data: [totalPay, totalTarp, Math.max(0, totalResult)],
                backgroundColor: ['#e89a35', '#19a58c', '#142b4a'],
                borderWidth: 0,
                cutout: '65%'
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { position: 'right' } }
        }
    });

    // --- GRÁFICO 3: Evolução Mensal (Linha) ---
    destruirGrafico('monthChart');
    const ctxMonth = document.getElementById('monthChart').getContext('2d');
    const mesesLabels = evolucaoMensal.map(r => r.mes.split('-').reverse().join('/'));
    
    charts['monthChart'] = new Chart(ctxMonth, {
        type: 'line',
        data: {
            labels: mesesLabels,
            datasets: [
                { label: 'Frete Cebrace', data: evolucaoMensal.map(r => r.freight), borderColor: '#2878d0', backgroundColor: '#2878d0', tension: 0.3 },
                { label: 'Resultado', data: evolucaoMensal.map(r => r.result), borderColor: '#142b4a', backgroundColor: '#142b4a', tension: 0.3, borderDash: [5, 5] }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { position: 'top' } },
            scales: { y: { beginAtZero: true } }
        }
    });

    // --- GRÁFICO 4: Ranking de Resultado (Barra Horizontal) ---
    destruirGrafico('rankChart');
    const ctxRank = document.getElementById('rankChart').getContext('2d');
    charts['rankChart'] = new Chart(ctxRank, {
        type: 'bar',
        data: {
            labels: driverLabels,
            datasets: [{
                label: 'Resultado Líquido',
                data: top10.map(r => r.result),
                backgroundColor: top10.map(r => r.result >= 0 ? '#13856f' : '#c54e4e'),
                borderRadius: 4
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: { x: { beginAtZero: true } }
        }
    });
}

// ============================================================================
// 9. EXPORTAÇÃO PARA CSV
// ============================================================================
function exportarCSV() {
    const dados = getDadosFiltrados();
    if (dados.length === 0) {
        alert("Não há dados para exportar com os filtros atuais.");
        return;
    }

    const cabecalho = "Data,Motorista,Tipo,Frete_Cebrace,Frete_Motorista,Enlonamento,Resultado\n";
    const linhas = dados.map(d => 
        `${d.date},"${d.driver}","${d.type}",${d.freight},${d.driverPay},${d.tarp},${d.result}`
    ).join("\n");

    const blob = new Blob([cabecalho + linhas], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "relatorio_fretes.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}