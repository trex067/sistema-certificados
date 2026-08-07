let listaFuncionarios = [];
let listaTSTs = [];
let assinaturaDataUrl = null;

// ==========================================
// MODO ESCURO / CLARO (THEME TOGGLE)
// ==========================================
const themeToggleBtn = document.getElementById('themeToggle');
const currentTheme = localStorage.getItem('theme') || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
document.documentElement.setAttribute('data-theme', currentTheme);

themeToggleBtn.addEventListener('click', () => {
    let theme = document.documentElement.getAttribute('data-theme');
    theme = theme === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
});

// ==========================================
// 1. LEITURA DE EXCEL, SHEETS E CTRL+V
// ==========================================

// Via Arquivo Excel
document.getElementById('fileInput').addEventListener('change', function(e) {
    const files = e.target.files;
    for (let f = 0; f < files.length; f++) {
        const reader = new FileReader();
        reader.onload = function(evt) {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet, {header: 1});
            processarArrayMatriz(json);
        };
        reader.readAsArrayBuffer(files[f]);
    }
});

// Via Link do Google Sheets
async function importarDoSheets() {
    const link = document.getElementById('linkSheets').value.trim();
    if (!link) return alert("Por favor, cole o link da planilha.");
    
    // Extrai o ID da URL do Google Sheets
    const match = link.match(/\/d\/(.*?)(\/|$)/);
    if (!match) return alert("Link inválido. Certifique-se de colar a URL completa da planilha.");
    
    const id = match[1];
    // URL nativa do Google para exportar a planilha pública como CSV
    const csvUrl = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv`;

    try {
        document.getElementById('linkSheets').value = "Processando...";
        const response = await fetch(csvUrl);
        if (!response.ok) throw new Error("A planilha está bloqueada. Altere o compartilhamento para 'Qualquer pessoa com o link'.");
        
        const csvText = await response.text();
        
        // Separa as linhas e limpa as aspas do CSV
        const linhas = csvText.split('\n').map(row => row.split(',').map(cell => cell.replace(/(^"|"$)/g, '')));
        
        processarArrayMatriz(linhas);
        document.getElementById('linkSheets').value = "";
        alert("Dados puxados com sucesso!");
    } catch (e) {
        document.getElementById('linkSheets').value = "";
        alert("Erro ao importar: " + e.message);
    }
}

// Via CTRL+V (Paste) - Global
document.addEventListener('paste', (e) => {
    // Ignora se o usuário estiver colando texto dentro de um campo de digitação (input)
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;

    let clipboardData = e.clipboardData || window.clipboardData;
    let pastedData = clipboardData.getData('Text');
    if (!pastedData) return;

    const linhas = pastedData.split('\n').map(linha => linha.split('\t'));
    processarArrayMatriz(linhas);
});

function processarArrayMatriz(matriz) {
    let adicionados = 0;
    let duplicados = 0;

    for (let i = 1; i < matriz.length; i++) {
        const linha = matriz[i];
        if (!linha || linha.length < 3) continue;
        
        let cod = linha[0] ? linha[0].toString().trim() : "";
        let nome = linha[1] ? linha[1].toString().trim().toUpperCase() : "";
        let cpfBruto = linha[2] ? linha[2].toString().trim() : "";
        let cpfValidado = formatarEValidarCPF(cpfBruto);

        if (!cpfValidado.valido) continue;

        // Bloqueia duplicados verificando o CPF na lista existente
        let duplicado = listaFuncionarios.find(f => f.cpf === cpfValidado.cpfFormatado);
        if (duplicado) {
            duplicados++;
            continue;
        }

        let dataAdm = converterData(linha[3]);
        
        listaFuncionarios.push({
            id: Date.now() + i, // ID único para a linha
            codigo: cod,
            nome: nome,
            cpf: cpfValidado.cpfFormatado,
            admissao: dataAdm
        });
        adicionados++;
    }
    
    if (duplicados > 0) alert(`${duplicados} funcionário(s) ignorado(s) pois já estavam na lista.`);
    renderizarTabela();
}

// ==========================================
// 2. FUNÇÕES DA TABELA (Adicionar, Excluir, Selecionar)
// ==========================================
function adicionarLinhaManual() {
    listaFuncionarios.push({
        id: Date.now(),
        codigo: "",
        nome: "",
        cpf: "",
        admissao: new Date()
    });
    renderizarTabela();
}

function excluirSelecionados() {
    const checkboxes = document.querySelectorAll('.chk-row:checked');
    const idsParaExcluir = Array.from(checkboxes).map(cb => parseInt(cb.value));
    
    listaFuncionarios = listaFuncionarios.filter(f => !idsParaExcluir.includes(f.id));
    document.getElementById('chkTodos').checked = false;
    renderizarTabela();
}

function limparTabela() {
    if(confirm("Tem certeza que deseja apagar todos os funcionários da lista?")) {
        listaFuncionarios = [];
        document.getElementById('chkTodos').checked = false;
        renderizarTabela();
    }
}

function toggleSelecionarTodos(masterCheckbox) {
    const checkboxes = document.querySelectorAll('.chk-row');
    checkboxes.forEach(cb => cb.checked = masterCheckbox.checked);
}

function atualizarDadoManual(id, campo, valor) {
    const func = listaFuncionarios.find(f => f.id === id);
    if (func) {
        if (campo === 'nome') func.nome = valor.toUpperCase();
        if (campo === 'codigo') func.codigo = valor;
        if (campo === 'cpf') func.cpf = formatarEValidarCPF(valor).cpfFormatado || valor;
        if (campo === 'admissao') func.admissao = new Date(valor + "T00:00:00");
    }
}

// ==========================================
// 3. TOGGLE DE NRS E RENDERIZAÇÃO
// ==========================================
function formataDataInput(data) {
    if (!(data instanceof Date) || isNaN(data)) return "";
    return data.toISOString().split('T')[0];
}

function toggleColunaNR(nr, isChecked) {
    listaFuncionarios.forEach((f, i) => {
        const checkbox = document.getElementById(`nr${nr}chk${i}`);
        if (checkbox) {
            checkbox.checked = isChecked;
            atualizarVisualData(nr, i);
        }
    });
}

function atualizarVisualData(nr, index) {
    const check = document.getElementById(`nr${nr}chk${index}`);
    const inputData = document.getElementById(`nr${nr}dt${index}`);
    if (check.checked) {
        inputData.classList.remove('nr-desativada');
        inputData.classList.add('nr-ativada');
    } else {
        inputData.classList.remove('nr-ativada');
        inputData.classList.add('nr-desativada');
    }
}

function renderizarTabela() {
    const tbody = document.getElementById('funcTableBody');
    tbody.innerHTML = '';
    
    listaFuncionarios.forEach((f, index) => {
        const dataAdmFormatada = formataDataInput(f.admissao);
        
        tbody.innerHTML += `
            <tr>
                <td><input type="checkbox" class="chk-row form-check-input" value="${f.id}"></td>
                <td class="col-codigo"><input type="text" class="form-control form-control-sm text-center fw-bold" value="${f.codigo}" onchange="atualizarDadoManual(${f.id}, 'codigo', this.value)"></td>
                <td class="col-nome"><input type="text" class="form-control form-control-sm text-uppercase" value="${f.nome}" onchange="atualizarDadoManual(${f.id}, 'nome', this.value)"></td>
                <td class="col-cpf"><input type="text" class="form-control form-control-sm text-center" value="${f.cpf}" onchange="atualizarDadoManual(${f.id}, 'cpf', this.value)"></td>
                <td class="col-data"><input type="date" class="form-control form-control-sm" value="${dataAdmFormatada}" onchange="atualizarDadoManual(${f.id}, 'admissao', this.value)"></td>
                
                ${['06', '12', '18', '35'].map(nr => `
                <td>
                    <div class="form-check form-switch d-flex justify-content-center mb-1">
                        <input class="form-check-input" type="checkbox" id="nr${nr}chk${index}" onchange="atualizarVisualData('${nr}', ${index})">
                    </div>
                    <input type="date" id="nr${nr}dt${index}" class="form-control form-control-sm nr-desativada" value="${dataAdmFormatada}">
                </td>`).join('')}
            </tr>
        `;
    });
}

// ==========================================
// 4. GESTÃO DE TST E CANVAS
// ==========================================
function adicionarTST() {
    const nome = document.getElementById('tstNome').value.trim().toUpperCase();
    const reg = document.getElementById('tstReg').value.trim().toUpperCase();
    const nrs = Array.from(document.querySelectorAll('.tst-nr:checked')).map(cb => cb.value);
    
    if (!nome || !reg) return alert("Preencha Nome e Registro do TST");
    if (!reg.includes('/')) return alert("O Registro deve conter a UF (Ex: 0000000/PR)");
    if (nrs.length === 0) return alert("Marque ao menos uma NR para este TST");

    for (let nr of nrs) {
        if (listaTSTs.find(t => t.nrs.includes(nr))) {
            return alert(`A ${nr} já está associada a outro TST na lista.`);
        }
    }
    
    listaTSTs.push({ nome, registro: reg, nrs });
    document.getElementById('tstNome').value = '';
    document.getElementById('tstReg').value = '';
    document.querySelectorAll('.tst-nr').forEach(cb => cb.checked = false);
    renderizarTSTs();
}

function excluirTST(index) {
    listaTSTs.splice(index, 1);
    renderizarTSTs();
}

function limparTSTs() {
    if(confirm("Apagar todos os Instrutores?")) {
        listaTSTs = [];
        renderizarTSTs();
    }
}

function renderizarTSTs() {
    const div = document.getElementById('tstList');
    div.innerHTML = listaTSTs.map((tst, i) => `
        <div class="alert alert-secondary d-flex justify-content-between align-items-center p-2 mb-2" style="background-color: var(--cor-input-bg); border-color: var(--cor-borda);">
            <div style="color: var(--cor-texto);"><strong>${tst.nome}</strong> (${tst.registro}) - Autorizado: <span class="badge bg-primary">${tst.nrs.join(', ')}</span></div>
            <button class="btn btn-sm btn-outline-danger" onclick="excluirTST(${i})">🗑️ Remover</button>
        </div>
    `).join('');
}

// Canvas
const canvas = document.getElementById('canvasAssinatura');
const ctx = canvas.getContext('2d');
let desenhando = false;
let canvasTocado = false;

function iniciarDesenho(e) { desenhando = true; desenhar(e); canvasTocado = true; }
function pararDesenho() { desenhando = false; ctx.beginPath(); }
function desenhar(e) {
    if (!desenhando) return;
    e.preventDefault();
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX || e.touches && e.touches[0].clientX) - rect.left;
    const y = (e.clientY || e.touches && e.touches[0].clientY) - rect.top;
    
    // Cor da caneta dependendo do tema escuro/claro
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    ctx.strokeStyle = isDark ? '#ffffff' : '#000000';
    
    ctx.lineWidth = 2; 
    ctx.lineCap = 'round';
    ctx.lineTo(x, y); 
    ctx.stroke(); 
    ctx.beginPath(); 
    ctx.moveTo(x, y);
}
canvas.addEventListener('mousedown', iniciarDesenho); canvas.addEventListener('mousemove', desenhar);
canvas.addEventListener('mouseup', pararDesenho); canvas.addEventListener('mouseout', pararDesenho);
canvas.addEventListener('touchstart', iniciarDesenho, {passive: false}); 
canvas.addEventListener('touchmove', desenhar, {passive: false});
canvas.addEventListener('touchend', pararDesenho);

function limparAssinatura() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    canvasTocado = false;
}

// ==========================================
// 5. VALIDAÇÕES (CPF/DATA) E ENVIO (API)
// ==========================================
function formatarEValidarCPF(valor) {
    if (!valor) return { valido: false };
    let cpf = valor.replace(/\D/g, '').padStart(11, '0');
    if (cpf.length !== 11 || /^(\d)\1{10}$/.test(cpf)) return { valido: false };
    let soma = 0, resto;
    for (let i = 1; i <= 9; i++) soma += parseInt(cpf.substring(i - 1, i)) * (11 - i);
    resto = (soma * 10) % 11;
    if (resto === 10 || resto === 11) resto = 0;
    if (resto !== parseInt(cpf.substring(9, 10))) return { valido: false };
    return { valido: true, cpfFormatado: cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4") };
}

function converterData(valor) {
    if (!valor) return new Date();
    if (typeof valor === 'number') return new Date(Math.round((valor - 25569) * 86400 * 1000));
    const partes = valor.toString().split('/');
    if (partes.length === 3) return new Date(partes[2], partes[1]-1, partes[0]);
    return new Date(valor);
}

async function gerarCertificados(destino) {
    const footer = document.getElementById('creditos-dev');
    if (!footer || !footer.innerHTML.includes('DDFR LTDA') || !footer.innerHTML.includes('C.A Segurança do Trabalho LTDA')) {
        document.body.innerHTML = `<h2 style='color:red; text-align:center; margin-top: 50px;'>⚠️ ACESSO BLOQUEADO - Direitos Autorais Violados.</h2>`;
        return;
    }

    if (listaFuncionarios.length === 0) return alert("Importe ou adicione funcionários na tabela.");
    
    // NOTA: Se você não configurou o API_GAS_URL no GitHub Secrets e tentar rodar isso na sua máquina, vai dar erro "API_URL is not defined" ou "Failed to fetch".
    if (typeof API_URL === 'undefined') {
        return alert("O site não está conectado ao servidor do Google. Você precisa subir os arquivos para o GitHub para que o sistema injete a chave de acesso (API_GAS_URL).");
    }

    if (canvasTocado) assinaturaDataUrl = canvas.toDataURL("image/png");

    const pacote = {
        destino: destino,
        emails: document.getElementById('emailsDestino').value.split(',').map(e => e.trim()).filter(e => e),
        tsts: listaTSTs,
        funcionarios: []
    };

    listaFuncionarios.forEach((f, i) => {
        let funcReq = { codigo: f.codigo, nome: f.nome, cpf: f.cpf, nrs: [], assinatura: assinaturaDataUrl };
        ['06', '12', '18', '35'].forEach(nr => {
            if (document.getElementById(`nr${nr}chk${i}`).checked) {
                const dataBr = document.getElementById(`nr${nr}dt${i}`).value.split('-').reverse().join('/');
                if (!listaTSTs.find(t => t.nrs.includes(`NR ${nr}`))) {
                    throw new Error(`Falta cadastrar um TST habilitado para a NR ${nr}.`);
                }
                funcReq.nrs.push({ tipo: `NR ${nr}`, data: dataBr });
            }
        });
        if (funcReq.nrs.length > 0) pacote.funcionarios.push(funcReq);
    });

    if (pacote.funcionarios.length === 0) return alert("Nenhum certificado foi marcado para geração (ative as chaves verdes).");

    const status = document.getElementById('statusProcessamento');
    status.innerHTML = `⌛ Processando via EstagIArio, enviando dados para o Google...`;
    status.className = "mt-3 fw-bold text-center text-primary";

    try {
        const response = await fetch(API_URL, { method: 'POST', body: JSON.stringify(pacote) });
        const result = await response.json();
        
        if (result.sucesso) {
            status.className = "mt-3 fw-bold text-center text-success";
            status.innerHTML = `✅ Sucesso! ${result.total} certificados processados.`;
            if (destino === 'ZIP' && result.link) status.innerHTML += `<br>🔗 <a href="${result.link}" target="_blank">BAIXAR ARQUIVO .ZIP</a>`;
        } else {
            throw new Error(result.erro);
        }
    } catch (error) {
        status.className = "mt-3 fw-bold text-center text-danger";
        status.innerHTML = `❌ Falha na conexão: ${error.message}`;
    }
}
