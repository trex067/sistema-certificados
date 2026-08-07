let listaFuncionarios = [];
let listaTSTs = [];

// ==========================================
// 1. LEITURA DE EXCEL E TEXTO
// ==========================================
document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, {header: 1});
        processarArrayMatriz(json);
    };
    reader.readAsArrayBuffer(file);
});

function processarTextoColado() {
    const texto = document.getElementById('pasteInput').value;
    const linhas = texto.split('\n').map(linha => linha.split('\t'));
    processarArrayMatriz(linhas);
}

function processarArrayMatriz(matriz) {
    listaFuncionarios = [];
    // Pula o cabeçalho
    for (let i = 1; i < matriz.length; i++) {
        const linha = matriz[i];
        if (!linha || linha.length < 3) continue;
        
        let cpfBruto = linha[2] ? linha[2].toString().trim() : "";
        let cpfValidado = formatarEValidarCPF(cpfBruto);
        
        if (!cpfValidado.valido) {
            alert(`⚠️ Atenção: O CPF de ${linha[1]} está incorreto ou vazio e não foi importado.`);
            continue;
        }
        
        // Conversão de Data Excel ou Texto
        let dataAdm = converterData(linha[3]);

        listaFuncionarios.push({
            codigo: linha[0],
            nome: linha[1],
            cpf: cpfValidado.cpfFormatado,
            admissao: dataAdm
        });
    }
    renderizarTabela();
}

// ==========================================
// 2. VALIDAÇÃO DE CPF (Adiciona zeros à esquerda)
// ==========================================
function formatarEValidarCPF(valor) {
    if (!valor) return { valido: false };
    let cpf = valor.replace(/\D/g, '').padStart(11, '0');
    if (cpf.length !== 11) return { valido: false };
    
    // Validação matemática básica
    if (/^(\d)\1{10}$/.test(cpf)) return { valido: false };
    let soma = 0, resto;
    for (let i = 1; i <= 9; i++) soma += parseInt(cpf.substring(i - 1, i)) * (11 - i);
    resto = (soma * 10) % 11;
    if (resto === 10 || resto === 11) resto = 0;
    if (resto !== parseInt(cpf.substring(9, 10))) return { valido: false };
    
    return { 
        valido: true, 
        cpfFormatado: cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4") 
    };
}

function converterData(valor) {
    if (!valor) return new Date();
    // Se vier do Excel formato número
    if (typeof valor === 'number') {
        return new Date(Math.round((valor - 25569) * 86400 * 1000));
    }
    // Formato string DD/MM/YYYY
    const partes = valor.split('/');
    if (partes.length === 3) return new Date(partes[2], partes[1]-1, partes[0]);
    return new Date(valor);
}

// ==========================================
// 3. DATAS AUTOMÁTICAS (Lógica original)
// ==========================================
function addDiasUteis(data, dias) {
    let d = new Date(data);
    let count = 0;
    while (count < dias) {
        d.setDate(d.getDate() + 1);
        if (d.getDay() !== 0 && d.getDay() !== 6) count++; // Ignorando fds simples no frontend
    }
    return d;
}
function formataDataInput(data) {
    return data.toISOString().split('T')[0];
}

// ==========================================
// 4. RENDERIZAÇÃO DA TABELA
// ==========================================
function renderizarTabela() {
    const tbody = document.getElementById('funcTableBody');
    tbody.innerHTML = '';
    
    listaFuncionarios.forEach((f, index) => {
        const d0618 = addDiasUteis(f.admissao, 1); // Dia seguinte útil
        const d12 = addDiasUteis(d0618, 1);
        const d35 = addDiasUteis(f.admissao, 10); // Simplificação, usuário pode alterar
        
        tbody.innerHTML += `
            <tr>
                <td>${f.codigo}</td>
                <td>${f.nome}</td>
                <td>${f.cpf}</td>
                <td>${f.admissao.toLocaleDateString('pt-BR')}</td>
                <td><input type="checkbox" id="nr06_chk_${index}"> <input type="date" id="nr06_dt_${index}" value="${formataDataInput(d0618)}" class="form-control form-control-sm mt-1" min="${formataDataInput(f.admissao)}"></td>
                <td><input type="checkbox" id="nr12_chk_${index}"> <input type="date" id="nr12_dt_${index}" value="${formataDataInput(d12)}" class="form-control form-control-sm mt-1" min="${formataDataInput(f.admissao)}"></td>
                <td><input type="checkbox" id="nr18_chk_${index}"> <input type="date" id="nr18_dt_${index}" value="${formataDataInput(d0618)}" class="form-control form-control-sm mt-1" min="${formataDataInput(f.admissao)}"></td>
                <td><input type="checkbox" id="nr35_chk_${index}"> <input type="date" id="nr35_dt_${index}" value="${formataDataInput(d35)}" class="form-control form-control-sm mt-1" min="${formataDataInput(f.admissao)}"></td>
            </tr>
        `;
    });
}

// ==========================================
// 5. GESTÃO DE TSTs
// ==========================================
function adicionarTST() {
    const nome = document.getElementById('tstNome').value;
    const reg = document.getElementById('tstReg').value;
    const nrs = Array.from(document.querySelectorAll('.tst-nr:checked')).map(cb => cb.value);
    
    if (!nome || !reg) { alert("Preencha Nome e Registro do TST"); return; }
    
    listaTSTs.push({ nome, registro: reg, nrs });
    
    document.getElementById('tstNome').value = '';
    document.getElementById('tstReg').value = '';
    document.querySelectorAll('.tst-nr').forEach(cb => cb.checked = false);
    
    atualizarListaTST();
}

function atualizarListaTST() {
    const div = document.getElementById('tstList');
    div.innerHTML = listaTSTs.map((tst, i) => 
        `<span class="badge bg-secondary me-2 p-2">
            ${tst.nome} (${tst.nrs.join(', ')})
            <button class="btn-close btn-close-white ms-2" style="font-size: 10px;" onclick="listaTSTs.splice(${i}, 1); atualizarListaTST();"></button>
        </span>`
    ).join('');
}

// ==========================================
// 6. ENVIO PARA O BACKEND (API)
// ==========================================
async function gerarCertificados(destino) {
// ==========================================
// SISTEMA ANTI-FRAUDE DE CRÉDITOS
// ==========================================
const footer = document.getElementById('creditos-dev');
if (!footer || !footer.innerHTML.includes('Daniel Filipe Rosa') || !footer.innerHTML.includes('Caroline Ávila')) {
    document.body.innerHTML = `
        <div style="display:flex; height:100vh; align-items:center; justify-content:center; background-color:#f8d7da; color:#721c24; flex-direction:column; text-align:center; font-family:sans-serif;">
            <h1 style="font-size: 50px;">⚠️ ACESSO BLOQUEADO</h1>
            <h3>Violação de Direitos Autorais.</h3>
            <p>Os créditos dos desenvolvedores originais foram removidos ou alterados.<br>O sistema não funcionará até que a autoria seja restaurada.</p>
        </div>`;
    return; // Para a execução do código na hora
}
    if (listaFuncionarios.length === 0) return alert("Importe os funcionários primeiro.");
    if (API_URL.includes("SECRET")) return alert("Aviso: Configuração do Servidor incompleta (Erro de Secret).");

    const btnZip = document.getElementById('btnZip');
    const btnEmail = document.getElementById('btnEmail');
    const status = document.getElementById('statusProcessamento');
    
    // Travar botões
    btnZip.disabled = true; btnEmail.disabled = true;
    status.innerHTML = `<span class="spinner-border spinner-border-sm text-primary"></span> Processando via EstagIArio, aguarde... (Isso pode levar alguns minutos dependendo da quantidade)`;
    status.className = "mt-3 fw-bold text-center text-primary";

    // Montar Pacote (Payload)
    const pacote = {
        destino: destino,
        emails: document.getElementById('emailsDestino').value.split(',').map(e => e.trim()).filter(e => e),
        tsts: listaTSTs,
        funcionarios: []
    };

    // Coletar NRs marcadas da tabela
    listaFuncionarios.forEach((f, i) => {
        let funcReq = { codigo: f.codigo, nome: f.nome, cpf: f.cpf, nrs: [] };
        
        ['06', '12', '18', '35'].forEach(nr => {
            const isChecked = document.getElementById(`nr${nr}_chk_${i}`).checked;
            if (isChecked) {
                const dataRaw = document.getElementById(`nr${nr}_dt_${i}`).value;
                const dataBr = dataRaw.split('-').reverse().join('/'); // Formata para DD/MM/YYYY para o PDF
                funcReq.nrs.push({ tipo: `NR ${nr}`, data: dataBr });
            }
        });
        
        if (funcReq.nrs.length > 0) pacote.funcionarios.push(funcReq);
    });

    if (pacote.funcionarios.length === 0) {
        status.innerHTML = "⚠️ Nenhuma NR foi marcada na tabela.";
        btnZip.disabled = false; btnEmail.disabled = false;
        return;
    }

    try {
        const response = await fetch(API_URL, {
            method: 'POST',
            body: JSON.stringify(pacote)
        });
        const result = await response.json();
        
        if (result.sucesso) {
            status.className = "mt-3 fw-bold text-center text-success";
            status.innerHTML = `✅ Sucesso! ${result.total} certificados gerados.`;
            if (destino === 'ZIP' && result.link) {
                status.innerHTML += `<br><br><a href="${result.link}" target="_blank" class="btn btn-success btn-lg">📥 CLIQUE AQUI PARA BAIXAR O ZIP</a>`;
            }
        } else {
            throw new Error(result.erro || "Erro desconhecido");
        }
    } catch (error) {
        status.className = "mt-3 fw-bold text-center text-danger";
        status.innerHTML = `❌ Erro de comunicação com o servidor: ${error.message}`;
    }

    btnZip.disabled = false; btnEmail.disabled = false;
}
