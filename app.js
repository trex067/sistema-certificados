const API_URL = "API_URL_SECRET";
let listaFuncionarios = [];
let listaTSTs = [];
let assinaturaDataUrl = null;

// Puxar Logo
window.addEventListener('DOMContentLoaded', async () => {
    if (typeof API_URL !== 'undefined' && !API_URL.includes("SECRET")) {
        try {
            const response = await fetch(API_URL);
            const data = await response.json();
            if (data.sucesso && data.logoId) {
                const imgLogo = document.getElementById('logoEmpresa');
                imgLogo.src = `https://drive.google.com/uc?id=${data.logoId}`;
                imgLogo.style.display = "inline-block";
            }
        } catch (e) { console.log("Aviso: Logo dinâmica não carregada."); }
    }
});

// Tema
const themeToggleBtn = document.getElementById('themeToggle');
const currentTheme = localStorage.getItem('theme') || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
document.documentElement.setAttribute('data-theme', currentTheme);
function atualizarIconeTema(theme) { themeToggleBtn.innerHTML = theme === 'dark' ? '☀️ Claro' : '🌙 Escuro'; }
atualizarIconeTema(currentTheme);
themeToggleBtn.addEventListener('click', () => {
    let theme = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
    atualizarIconeTema(theme);
});

// Excel e Sheets
let filesTemporarios = null;
document.getElementById('fileInput').addEventListener('change', e => filesTemporarios = e.target.files);

function processarExcelVisual() {
    if (!filesTemporarios || filesTemporarios.length === 0) return alert("⚠️ Selecione arquivos Excel primeiro.");
    let processados = 0;
    for (let f = 0; f < filesTemporarios.length; f++) {
        const reader = new FileReader();
        reader.onload = function(evt) {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            processarArrayMatriz(XLSX.utils.sheet_to_json(sheet, {header: 1}));
        };
        reader.readAsArrayBuffer(filesTemporarios[f]);
        processados++;
    }
    document.getElementById('fileInput').value = "";
    filesTemporarios = null;
    if(processados > 0) alert("✅ Arquivos Excel enviados para processamento!");
}

async function importarDoSheets() {
    const link = document.getElementById('linkSheets').value.trim();
    if (!link) return alert("Por favor, cole o link da planilha.");
    const match = link.match(/\/d\/(.*?)(\/|$)/);
    if (!match) return alert("Link inválido.");
    const id = match[1];
    
    try {
        document.getElementById('linkSheets').value = "Processando...";
        const response = await fetch(`https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv`);
        if (!response.ok) throw new Error("Planilha bloqueada. Altere para 'Qualquer pessoa com o link'.");
        const csvText = await response.text();
        const linhas = csvText.split('\n').map(row => row.split(',').map(cell => cell.replace(/(^"|"$)/g, '')));
        processarArrayMatriz(linhas);
        document.getElementById('linkSheets').value = "";
        alert("✅ Dados puxados da nuvem!");
    } catch (e) {
        document.getElementById('linkSheets').value = "";
        alert("❌ Erro: " + e.message);
    }
}

// CTRL+V Mágico
document.addEventListener('paste', (e) => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
    let pastedData = (e.clipboardData || window.clipboardData).getData('Text');
    if (pastedData) processarTextoBaguncado(pastedData);
});

function processarTextoBaguncado(texto) {
    let textoNormalizado = texto.replace(/\r/g, '').replace(/\n\s*(==>|Colaborador|CPF|Admissão|Documento)/gi, ' $1');
    let linhas = textoNormalizado.split('\n');
    let adicionados = 0, duplicados = 0;
    const regexCpf = /(?:\d{3}[\.\s,\-\/]*){3}\d{2}/; 
    const regexData = /\d{2}[\/\.-]\d{2}[\/\.-]\d{2,4}|\d{4}-\d{2}-\d{2}|\d{2}\s+de\s+[a-zA-Z]{3,9}\s+de\s+\d{2,4}/i;
    const regexCodMarkers = /(?:c[óo]d(?:igo)?|id)\s*[:=\-]?\s*#?\s*([0-9a-z\-]+)/i;

    linhas.forEach(linha => {
        let textoLinha = linha.trim();
        if (!textoLinha || textoLinha.includes('---')) return;
        let matchCpf = textoLinha.match(regexCpf);
        if (matchCpf) {
            let cpfValidado = formatarEValidarCPF(matchCpf[0]);
            if (cpfValidado.valido) {
                let matchData = textoLinha.match(regexData);
                let dataAdm = converterData(matchData ? matchData[0] : null);
                let cod = "";
                let matchCod = textoLinha.match(regexCodMarkers);
                if (matchCod && matchCod[1]) cod = matchCod[1];
                else {
                    let matchNumInicio = textoLinha.match(/^\[?#?(\d+)\]?[\s;\|,-]/);
                    if (matchNumInicio) cod = matchNumInicio[1];
                }
                if (!isNaN(cod) && cod !== "") cod = parseInt(cod, 10).toString();
                if (cod.toLowerCase().startsWith('cod-')) cod = cod.substring(4); 
                
                let lixos = [matchCpf[0], matchData ? matchData[0] : "", matchCod ? matchCod[0] : "", /Nome do Colaborador/gi, /Colaborador/gi, /Documento/gi, /Entrou na obra em/gi, /Admissão/gi, /adm/gi, /INICIO/gi, /data/gi, /cpf/gi, /erro de leitura/gi, /c[óo]digo/gi, /cod/gi, /id/gi, /==>/g, /\*+/g, /\[|\]/g];
                let nomeSujo = textoLinha;
                lixos.forEach(lixo => { if (lixo) nomeSujo = nomeSujo.replace(lixo, ' '); });
                let nomeLimpo = nomeSujo.replace(/[^a-zA-ZáéíóúãõçÁÉÍÓÚÃÕÇ\s]/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase();
                nomeLimpo = nomeLimpo.split(' ').filter(p => p.length > 1 || p === 'E').join(' ') || "NOME NÃO ENCONTRADO";

                if (listaFuncionarios.find(f => f.cpf === cpfValidado.cpfFormatado)) duplicados++;
                else {
                    listaFuncionarios.push({ id: Date.now() + Math.floor(Math.random()*1000), codigo: cod, nome: nomeLimpo, cpf: cpfValidado.cpfFormatado, admissao: dataAdm });
                    adicionados++;
                }
            }
        }
    });
    if (adicionados > 0 || duplicados > 0) {
        alert(`✨ MÁGICA CONCLUÍDA!\n\n✅ ${adicionados} inseridos.\n⚠️ ${duplicados} ignorados (duplicados).`);
        renderizarTabela();
    }
}

function processarArrayMatriz(matriz) {
    let adicionados = 0, duplicados = 0;
    for (let i = 1; i < matriz.length; i++) {
        if (!matriz[i] || matriz[i].length < 3) continue;
        let cpfValidado = formatarEValidarCPF(matriz[i][2] ? matriz[i][2].toString().trim() : "");
        if (!cpfValidado.valido) continue;
        if (listaFuncionarios.find(f => f.cpf === cpfValidado.cpfFormatado)) { duplicados++; continue; }
        
        listaFuncionarios.push({
            id: Date.now() + Math.floor(Math.random()*1000), 
            codigo: matriz[i][0] ? matriz[i][0].toString().trim() : "",
            nome: matriz[i][1] ? matriz[i][1].toString().trim().toUpperCase() : "",
            cpf: cpfValidado.cpfFormatado,
            admissao: converterData(matriz[i][3])
        });
        adicionados++;
    }
    if (duplicados > 0) alert(`⚠️ ${duplicados} funcionário(s) ignorado(s) (duplicados).`);
    renderizarTabela();
}

// Tabela
function adicionarLinhaManual() { listaFuncionarios.push({ id: Date.now(), codigo: "", nome: "", cpf: "", admissao: new Date() }); renderizarTabela(); }
function excluirSelecionados() {
    const idsParaExcluir = Array.from(document.querySelectorAll('.chk-row:checked')).map(cb => parseInt(cb.value));
    listaFuncionarios = listaFuncionarios.filter(f => !idsParaExcluir.includes(f.id));
    document.getElementById('chkTodos').checked = false;
    renderizarTabela();
}
function limparTabela() { if(confirm("Apagar todos da lista?")) { listaFuncionarios = []; document.getElementById('chkTodos').checked = false; renderizarTabela(); } }
function toggleSelecionarTodos(chk) { document.querySelectorAll('.chk-row').forEach(cb => cb.checked = chk.checked); }
function atualizarDadoManual(id, campo, valor) {
    const func = listaFuncionarios.find(f => f.id === id);
    if (func) {
        if (campo === 'nome') func.nome = valor.toUpperCase();
        if (campo === 'codigo') func.codigo = valor;
        if (campo === 'cpf') func.cpf = formatarEValidarCPF(valor).cpfFormatado || valor;
        if (campo === 'admissao') func.admissao = new Date(valor + "T00:00:00");
    }
}

function formataDataInput(data) { return (!(data instanceof Date) || isNaN(data)) ? "" : data.toISOString().split('T')[0]; }
function toggleColunaNR(nr, isChecked) {
    listaFuncionarios.forEach((f, i) => {
        const checkbox = document.getElementById(`nr${nr}chk${i}`);
        if (checkbox) { checkbox.checked = isChecked; atualizarVisualData(nr, i); }
    });
}
function atualizarVisualData(nr, index) {
    const check = document.getElementById(`nr${nr}chk${index}`), inputData = document.getElementById(`nr${nr}dt${index}`);
    inputData.className = check.checked ? "form-control form-control-sm nr-ativada" : "form-control form-control-sm nr-desativada";
}

function renderizarTabela() {
    const tbody = document.getElementById('funcTableBody');
    tbody.innerHTML = '';
    listaFuncionarios.forEach((f, index) => {
        const dAdm = formataDataInput(f.admissao);
        tbody.innerHTML += `
            <tr>
                <td><input type="checkbox" class="chk-row form-check-input" value="${f.id}"></td>
                <td class="col-codigo"><input type="text" class="form-control form-control-sm text-center fw-bold" value="${f.codigo}" onchange="atualizarDadoManual(${f.id}, 'codigo', this.value)"></td>
                <td class="col-nome"><input type="text" class="form-control form-control-sm text-uppercase" value="${f.nome}" onchange="atualizarDadoManual(${f.id}, 'nome', this.value)"></td>
                <td class="col-cpf"><input type="text" class="form-control form-control-sm text-center" value="${f.cpf}" onchange="atualizarDadoManual(${f.id}, 'cpf', this.value)"></td>
                <td class="col-data"><input type="date" class="form-control form-control-sm" value="${dAdm}" onchange="atualizarDadoManual(${f.id}, 'admissao', this.value)"></td>
                ${['06', '12', '18', '35'].map(nr => `
                <td>
                    <div class="form-check form-switch d-flex justify-content-center mb-1"><input class="form-check-input" type="checkbox" id="nr${nr}chk${index}" onchange="atualizarVisualData('${nr}', ${index})"></div>
                    <input type="date" id="nr${nr}dt${index}" class="form-control form-control-sm nr-desativada" value="${dAdm}">
                </td>`).join('')}
            </tr>`;
    });
}

// TST (AGORA COM JUNÇÃO DO NÚMERO + UF)
function adicionarTST() {
    const nome = document.getElementById('tstNome').value.trim().toUpperCase();
    const regNum = document.getElementById('tstRegNum').value.trim().toUpperCase();
    const regUF = document.getElementById('tstRegUF').value;
    const nrs = Array.from(document.querySelectorAll('.tst-nr:checked')).map(cb => cb.value);
    
    if (!nome || !regNum) return alert("Preencha o Nome e o Número de Registro do TST");
    if (nrs.length === 0) return alert("Marque ao menos uma NR para este TST");

    for (let nr of nrs) {
        if (listaTSTs.find(t => t.nrs.includes(nr))) return alert(`A ${nr} já está associada a outro TST.`);
    }
    
    listaTSTs.push({ nome, registro: `${regNum}/${regUF}`, nrs });
    
    document.getElementById('tstNome').value = '';
    document.getElementById('tstRegNum').value = '';
    document.querySelectorAll('.tst-nr').forEach(cb => cb.checked = false);
    renderizarTSTs();
}
function excluirTST(index) { listaTSTs.splice(index, 1); renderizarTSTs(); }
function limparTSTs() { if(confirm("Apagar todos os Instrutores?")) { listaTSTs = []; renderizarTSTs(); } }
function renderizarTSTs() {
    document.getElementById('tstList').innerHTML = listaTSTs.map((tst, i) => `
        <div class="alert alert-secondary d-flex justify-content-between align-items-center p-2 mb-2" style="background-color: var(--cor-input-bg); border-color: var(--cor-borda);">
            <div style="color: var(--cor-texto);"><strong>${tst.nome}</strong> (${tst.registro}) - <span class="badge bg-primary">${tst.nrs.join(', ')}</span></div>
            <button class="btn btn-sm btn-outline-danger" onclick="excluirTST(${i})">🗑️ Remover</button>
        </div>`).join('');
}

// Assinatura
const canvas = document.getElementById('canvasAssinatura'), ctx = canvas.getContext('2d');
let desenhando = false, canvasTocado = false;
function iniciarDesenho(e) { desenhando = true; desenhar(e); canvasTocado = true; }
function pararDesenho() { desenhando = false; ctx.beginPath(); }
function desenhar(e) {
    if (!desenhando) return;
    e.preventDefault();
    const rect = canvas.getBoundingClientRect(), x = (e.clientX || (e.touches && e.touches[0].clientX)) - rect.left, y = (e.clientY || (e.touches && e.touches[0].clientY)) - rect.top;
    ctx.strokeStyle = document.documentElement.getAttribute('data-theme') === 'dark' ? '#ffffff' : '#000000';
    ctx.lineWidth = 2; ctx.lineCap = 'round'; ctx.lineTo(x, y); ctx.stroke(); ctx.beginPath(); ctx.moveTo(x, y);
}
canvas.addEventListener('mousedown', iniciarDesenho); canvas.addEventListener('mousemove', desenhar);
canvas.addEventListener('mouseup', pararDesenho); canvas.addEventListener('mouseout', pararDesenho);
canvas.addEventListener('touchstart', iniciarDesenho, {passive: false}); canvas.addEventListener('touchmove', desenhar, {passive: false}); canvas.addEventListener('touchend', pararDesenho);
function limparAssinatura() { ctx.clearRect(0, 0, canvas.width, canvas.height); canvasTocado = false; }

// ==========================================
// VALIDADOR MATEMÁTICO REAL DO CPF (CORRIGIDO)
// ==========================================
function formatarEValidarCPF(valor) {
    if (!valor) return { valido: false };
    let cpf = valor.replace(/\D/g, '').padStart(11, '0');
    
    // Barrar CPFs vazios ou com números repetidos (111.111.111-11)
    if (cpf.length !== 11 || /^(\d)\1{10}$/.test(cpf)) return { valido: false };
    
    let soma = 0, resto;
    
    // Validação Matemática do 1º Dígito (Peso 10 a 2)
    for (let i = 1; i <= 9; i++) soma += parseInt(cpf.substring(i - 1, i)) * (11 - i);
    resto = (soma * 10) % 11;
    if (resto === 10 || resto === 11) resto = 0;
    if (resto !== parseInt(cpf.substring(9, 10))) return { valido: false };
    
    // Validação Matemática do 2º Dígito (Peso 11 a 2)
    soma = 0;
    for (let i = 1; i <= 10; i++) soma += parseInt(cpf.substring(i - 1, i)) * (12 - i);
    resto = (soma * 10) % 11;
    if (resto === 10 || resto === 11) resto = 0;
    if (resto !== parseInt(cpf.substring(10, 11))) return { valido: false }; // Aqui a mentira cai!
    
    return { valido: true, cpfFormatado: cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4") };
}

function converterData(valor) {
    if (!valor) return new Date();
    if (typeof valor === 'number') return new Date(Math.round((valor - 25569) * 86400 * 1000));
    let valorStr = valor.toString().trim().toLowerCase();
    if (valorStr.includes(' de ')) {
        const meses = { 'jan': 0, 'feb': 1, 'fev': 1, 'mar': 2, 'apr': 3, 'abr': 3, 'may': 4, 'mai': 4, 'jun': 5, 'jul': 6, 'aug': 7, 'ago': 7, 'sep': 8, 'set': 8, 'oct': 9, 'out': 9, 'nov': 10, 'dec': 11, 'dez': 11 };
        let partes = valorStr.split(' de ');
        if (partes.length === 3) return new Date(parseInt(partes[2], 10), (meses[partes[1].substring(0, 3)] || 0), parseInt(partes[0], 10));
    }
    let partes = valorStr.replace(/[\.\-]/g, '/').split('/');
    if (partes.length === 3) {
        if (partes[0].length === 4) return new Date(parseInt(partes[0]), parseInt(partes[1])-1, parseInt(partes[2]));
        let ano = parseInt(partes[2]); if (ano < 100) ano += 2000;
        return new Date(ano, parseInt(partes[1])-1, parseInt(partes[0]));
    }
    return new Date();
}

async function gerarCertificados(destino) {
    const footer = document.getElementById('creditos-dev');
    if (!footer || !footer.innerHTML.includes('DDFR LTDA')) return;
    if (listaFuncionarios.length === 0) return alert("Importe ou adicione funcionários na tabela.");
    
    // A NOVA TRAVA DO E-MAIL VAZIO
    const emailsInput = document.getElementById('emailsDestino').value;
    const emailsDestino = emailsInput.split(',').map(e => e.trim()).filter(e => e);
    
    if (destino === 'EMAIL' && emailsDestino.length === 0) {
        return alert("⚠️ ERRO: Digite pelo menos um endereço de e-mail na caixa acima para enviar.");
    }
    
    if (typeof API_URL === 'undefined' || API_URL.includes("SECRET")) return alert("Erro: Conecte a API via GitHub Secrets.");
    if (canvasTocado) assinaturaDataUrl = canvas.toDataURL("image/png");

    const pacote = { destino: destino, emails: emailsDestino, tsts: listaTSTs, funcionarios: [] };

    listaFuncionarios.forEach((f, i) => {
        let funcReq = { codigo: f.codigo, nome: f.nome, cpf: f.cpf, nrs: [], assinatura: assinaturaDataUrl };
        ['06', '12', '18', '35'].forEach(nr => {
            if (document.getElementById(`nr${nr}chk${i}`).checked) {
                const dataBr = document.getElementById(`nr${nr}dt${i}`).value.split('-').reverse().join('/');
                if (!listaTSTs.find(t => t.nrs.includes(`NR ${nr}`))) throw new Error(`Falta cadastrar TST para a NR ${nr}.`);
                funcReq.nrs.push({ tipo: `NR ${nr}`, data: dataBr });
            }
        });
        if (funcReq.nrs.length > 0) pacote.funcionarios.push(funcReq);
    });

    if (pacote.funcionarios.length === 0) return alert("Nenhum certificado foi marcado para geração (ative as chaves verdes).");

    const status = document.getElementById('statusProcessamento');
    status.innerHTML = `⌛ Processando no Google Drive... (Isso pode levar alguns segundos)`;
    status.className = "mt-3 fw-bold text-center text-primary";

    try {
        const response = await fetch(API_URL, { method: 'POST', body: JSON.stringify(pacote) });
        const result = await response.json();
        if (result.sucesso) {
            status.className = "mt-3 fw-bold text-center text-success";
            status.innerHTML = `✅ Sucesso! ${result.total} certificados processados.`;
            if (destino === 'ZIP' && result.link) status.innerHTML += `<br>🔗 <a href="${result.link}" target="_blank">BAIXAR ARQUIVO .ZIP</a>`;
        } else throw new Error(result.erro);
    } catch (error) {
        status.className = "mt-3 fw-bold text-center text-danger";
        status.innerHTML = `❌ Falha: ${error.message}`;
    }
}
