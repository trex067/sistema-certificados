let listaFuncionarios = [];
let listaTSTs = [];

// ==========================================
// MÓDULO DE ASSINATURA DIGITAL (CANVAS)
// ==========================================
const canvas = document.getElementById('canvasAssinatura');
const ctx = canvas ? canvas.getContext('2d') : null;
let desenhando = false;

if (canvas) {
    if(ctx) { ctx.lineWidth = 3; ctx.lineCap = 'round'; ctx.strokeStyle = '#000000'; }

    canvas.addEventListener('mousedown', (e) => { desenhando = true; ctx.beginPath(); ctx.moveTo(e.offsetX, e.offsetY); });
    canvas.addEventListener('mousemove', (e) => { if (desenhando) { ctx.lineTo(e.offsetX, e.offsetY); ctx.stroke(); } });
    canvas.addEventListener('mouseup', () => desenhando = false);
    canvas.addEventListener('mouseout', () => desenhando = false);
    
    // Tratamento responsivo para Touch (Mobile)
    canvas.addEventListener('touchstart', (e) => {
        e.preventDefault(); desenhando = true;
        const rect = canvas.getBoundingClientRect();
        // Corrige o scale caso a tela seja redimensionada
        const scaleX = canvas.width / rect.width;
        const scaleY = canvas.height / rect.height;
        ctx.beginPath(); 
        ctx.moveTo((e.touches[0].clientX - rect.left) * scaleX, (e.touches[0].clientY - rect.top) * scaleY);
    }, {passive: false});
    
    canvas.addEventListener('touchmove', (e) => {
        e.preventDefault();
        if (desenhando) {
            const rect = canvas.getBoundingClientRect();
            const scaleX = canvas.width / rect.width;
            const scaleY = canvas.height / rect.height;
            ctx.lineTo((e.touches[0].clientX - rect.left) * scaleX, (e.touches[0].clientY - rect.top) * scaleY);
            ctx.stroke();
        }
    }, {passive: false});
    canvas.addEventListener('touchend', () => desenhando = false);
}

function toggleAssinatura() { document.getElementById('boxAssinatura').style.display = document.getElementById('chkAssinatura').checked ? 'block' : 'none'; }
function limparCanvas() { if(ctx) ctx.clearRect(0, 0, canvas.width, canvas.height); }

// ==========================================
// 1. IMPORTAÇÃO DE DADOS
// ==========================================
async function importarArquivosExcel() {
    const input = document.getElementById('fileInput');
    if (input.files.length === 0) return alert("Selecione pelo menos um arquivo Excel.");
    let totalAdicionados = 0, totalDuplicados = 0, totalIgnorados = 0;
    for (let i = 0; i < input.files.length; i++) {
        await new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = function(evt) {
                const workbook = XLSX.read(new Uint8Array(evt.target.result), {type: 'array'});
                const res = processarDados(XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1}), false);
                totalAdicionados += res.adicionados; totalDuplicados += res.duplicados; totalIgnorados += res.ignorados; resolve();
            };
            reader.readAsArrayBuffer(input.files[i]);
        });
    }
    input.value = '';
    alert(`Leitura concluída!\n✅ ${totalAdicionados} adicionados\n🔁 ${totalDuplicados} duplicados\n⚠️ ${totalIgnorados} inválidos`);
}

async function importarGoogleSheets() {
    const match = document.getElementById('sheetsLink').value.match(/\/d\/(.*?)(\/|$)/);
    if (!match) return alert("Link inválido.");
    const btn = document.querySelector('#tab-sheets button');
    const txtOriginal = btn.innerHTML; btn.innerHTML = "⏳ Baixando..."; btn.disabled = true;
    try {
        const response = await fetch(`https://docs.google.com/spreadsheets/d/${match[1]}/export?format=xlsx`);
        if (!response.ok) throw new Error("A planilha não está pública.");
        const workbook = XLSX.read(await response.arrayBuffer(), { type: 'array' });
        const res = processarDados(XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1}), false);
        document.getElementById('sheetsLink').value = '';
        alert(`Sucesso!\n✅ ${res.adicionados} adicionados\n🔁 ${res.duplicados} duplicados ignorados`);
    } catch(error) { alert(`⚠️ Erro: ${error.message}`); } 
    finally { btn.innerHTML = txtOriginal; btn.disabled = false; }
}

document.addEventListener('paste', function(e) {
    if (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA') return;
    const pastedData = (e.clipboardData || window.clipboardData).getData('Text');
    if (pastedData && pastedData.includes('\t')) {
        e.preventDefault();
        const res = processarDados(pastedData.split('\n').map(linha => linha.split('\t')), false);
        alert(`Dados Colados!\n✅ ${res.adicionados} adicionados\n🔁 ${res.duplicados} duplicados`);
    }
});

function processarDados(matriz, exibirAlerta = true) {
    let adicionados = 0, duplicados = 0, ignorados = 0;
    for (let i = 0; i < matriz.length; i++) {
        const linha = matriz[i];
        if (!linha || linha.length < 3) continue;
        let cpfBruto = linha[2] ? linha[2].toString().trim() : "";
        if (cpfBruto.toLowerCase() === "cpf" || (linha[1] && linha[1].toString().toLowerCase() === "nome")) continue;
        let cpfValidado = formatarEValidarCPF(cpfBruto);
        if (!cpfValidado.valido) { ignorados++; continue; }
        if (listaFuncionarios.find(f => f.cpf === cpfValidado.cpfFormatado)) { duplicados++; continue; }
        listaFuncionarios.push({ codigo: linha[0] || "", nome: linha[1] ? linha[1].toString().toUpperCase().trim() : "", cpf: cpfValidado.cpfFormatado, admissao: converterData(linha[3]) });
        adicionados++;
    }
    renderizarTabela(); return { adicionados, duplicados, ignorados };
}

function limparLista() {
    if (listaFuncionarios.length > 0 && confirm("🚨 Apagar todos os funcionários atuais da tela?")) { listaFuncionarios = []; renderizarTabela(); }
}

// ==========================================
// 2. FUNÇÕES ÚTEIS
// ==========================================
function formatarEValidarCPF(valor) {
    if (!valor) return { valido: false };
    let cpf = valor.replace(/\D/g, '').padStart(11, '0');
    if (cpf.length !== 11 || /^(\d)\1{10}$/.test(cpf)) return { valido: false };
    let soma = 0, resto;
    for (let i = 1; i <= 9; i++) soma += parseInt(cpf.substring(i - 1, i)) * (11 - i);
    resto = (soma * 10) % 11; if (resto >= 10) resto = 0;
    if (resto !== parseInt(cpf.substring(9, 10))) return { valido: false };
    return { valido: true, cpfFormatado: cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4") };
}

function converterData(v) {
    if (!v) return new Date();
    if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400 * 1000));
    const s = v.toString().trim();
    if (s.includes('/')) {
        const p = s.split('/');
        if (p.length === 3) return new Date(p[2].split(' ')[0], p[1] - 1, p[0]);
    }
    const d = new Date(s); return isNaN(d.getTime()) ? new Date() : d;
}
function addDiasUteis(data, dias) {
    let d = new Date(data); let count = 0;
    while (count < dias) { d.setDate(d.getDate() + 1); if (d.getDay() !== 0 && d.getDay() !== 6) count++; }
    return d;
}
function formataDataInput(data) { return data.toISOString().split('T')[0]; }

// ==========================================
// 3. RENDERIZAÇÃO
// ==========================================
function renderizarTabela() {
    const tbody = document.getElementById('funcTableBody');
    if (listaFuncionarios.length === 0) { tbody.innerHTML = '<tr><td colspan="8" class="text-center text-muted py-5 fs-5">Nenhum dado importado.</td></tr>'; return; }
    tbody.innerHTML = '';
    listaFuncionarios.forEach((f, i) => {
        const d0618 = addDiasUteis(f.admissao, 1), d12 = addDiasUteis(d0618, 1), d35 = addDiasUteis(f.admissao, 10);
        tbody.innerHTML += `<tr>
            <td class="fw-bold">${f.codigo}</td>
            <td class="text-nowrap fw-bold">${f.nome}</td>
            <td class="text-nowrap">${f.cpf}</td>
            <td class="text-nowrap">${f.admissao.toLocaleDateString('pt-BR')}</td>
            <td><div class="form-check form-switch"><input class="form-check-input" type="checkbox" id="nr06_chk_${i}"></div><input type="date" id="nr06_dt_${i}" value="${formataDataInput(d0618)}" class="form-control mt-1" min="${formataDataInput(f.admissao)}"></td>
            <td><div class="form-check form-switch"><input class="form-check-input" type="checkbox" id="nr12_chk_${i}"></div><input type="date" id="nr12_dt_${i}" value="${formataDataInput(d12)}" class="form-control mt-1" min="${formataDataInput(f.admissao)}"></td>
            <td><div class="form-check form-switch"><input class="form-check-input" type="checkbox" id="nr18_chk_${i}"></div><input type="date" id="nr18_dt_${i}" value="${formataDataInput(d0618)}" class="form-control mt-1" min="${formataDataInput(f.admissao)}"></td>
            <td><div class="form-check form-switch"><input class="form-check-input" type="checkbox" id="nr35_chk_${i}"></div><input type="date" id="nr35_dt_${i}" value="${formataDataInput(d35)}" class="form-control mt-1" min="${formataDataInput(f.admissao)}"></td>
        </tr>`;
    });
}

// ==========================================
// 4. GESTÃO DE TSTs
// ==========================================
function adicionarTST() {
    const nome = document.getElementById('tstNome').value.toUpperCase().trim();
    const reg = document.getElementById('tstReg').value.trim();
    const nrs = Array.from(document.querySelectorAll('.tst-nr:checked')).map(cb => cb.value);
    
    if (!nome) return alert("❌ Preencha o Nome Completo do TST.");
    if (!reg) return alert("❌ Preencha o Nº de Registro do TST.");
    if (nrs.length === 0) return alert("❌ Marque pelo menos 1 NR habilitada para este TST.");
    
    let assinatura = null;
    if (document.getElementById('chkAssinatura').checked && canvas) assinatura = canvas.toDataURL('image/png');

    listaTSTs.push({ nome, registro: reg, nrs, assinatura });
    
    document.getElementById('tstNome').value = ''; document.getElementById('tstReg').value = '';
    document.querySelectorAll('.tst-nr').forEach(cb => cb.checked = false);
    document.getElementById('chkAssinatura').checked = false; toggleAssinatura(); limparCanvas();
    atualizarListaTST();
}

function editarTST(index) {
    const t = listaTSTs[index];
    document.getElementById('tstNome').value = t.nome;
    document.getElementById('tstReg').value = t.registro;
    document.querySelectorAll('.tst-nr').forEach(cb => { cb.checked = t.nrs.includes(cb.value); });
    listaTSTs.splice(index, 1); atualizarListaTST(); document.getElementById('tstNome').focus();
}

function atualizarListaTST() {
    document.getElementById('tstList').innerHTML = listaTSTs.map((t, i) => `
        <div class="badge bg-secondary p-3 d-flex align-items-center flex-column flex-md-row fs-6">
            <span>👤 ${t.nome} (${t.registro}) <br class="d-md-none"> 📌 NRs: ${t.nrs.join(', ')} ${t.assinatura ? '🖋️' : ''}</span>
            <div class="mt-2 mt-md-0 ms-md-3">
                <button class="btn btn-sm btn-light text-primary py-0" onclick="editarTST(${i})">✏️ Editar</button>
                <button class="btn btn-sm btn-light text-danger py-0" onclick="listaTSTs.splice(${i}, 1); atualizarListaTST();">✖ Remover</button>
            </div>
        </div>`).join('');
}

// ==========================================
// 5. INTEGRAÇÃO E ENVIO (Sem Template Secreto)
// ==========================================
async function gerarCertificados(destino) {
    const footer = document.getElementById('creditos-dev');
    if (!footer || !footer.innerHTML.includes('Daniel Filipe Rosa') || !footer.innerHTML.includes('Caroline Ávila')) {
        document.body.innerHTML = `<div style="display:flex; height:100vh; align-items:center; justify-content:center; background-color:#111; color:#FF8C00; flex-direction:column; text-align:center;">
            <h1>⚠️ ACESSO BLOQUEADO</h1><h3>Violação de Autoria.</h3></div>`;
        return; 
    }

    if (listaFuncionarios.length === 0) return alert("Importe os funcionários primeiro.");
    
    // O Template não é mais enviado no pacote. O Google Apps Script vai ler direto da planilha!
    const pacote = { destino: destino, emails: document.getElementById('emailsDestino').value.split(',').map(e=>e.trim()).filter(e=>e), tsts: listaTSTs, funcionarios: [] };

    let nrsRequisitadas = new Set();
    listaFuncionarios.forEach((f, i) => {
        let funcReq = { codigo: f.codigo, nome: f.nome, cpf: f.cpf, nrs: [] };
        ['06', '12', '18', '35'].forEach(nr => {
            if (document.getElementById(`nr${nr}_chk_${i}`).checked) {
                funcReq.nrs.push({ tipo: `NR ${nr}`, data: document.getElementById(`nr${nr}_dt_${i}`).value.split('-').reverse().join('/') });
                nrsRequisitadas.add(`NR ${nr}`);
            }
        });
        if (funcReq.nrs.length > 0) pacote.funcionarios.push(funcReq);
    });

    if (pacote.funcionarios.length === 0) return alert("⚠️ Nenhuma NR foi marcada para nenhum funcionário.");

    let nrsCobertas = new Set();
    listaTSTs.forEach(tst => tst.nrs.forEach(nr => nrsCobertas.add(nr)));
    for (let nrNecessaria of nrsRequisitadas) {
        if (!nrsCobertas.has(nrNecessaria)) return alert(`🚨 GERAÇÃO BLOQUEADA!\nFalta um TST habilitado para a norma: ${nrNecessaria}`);
    }

    const btnZip = document.getElementById('btnZip'); const btnEmail = document.getElementById('btnEmail'); const status = document.getElementById('statusProcessamento');
    btnZip.disabled = true; btnEmail.disabled = true;
    status.innerHTML = `<span class="spinner-border spinner-border-sm" style="color: var(--piacentini-orange);"></span> Gerando documentos na Nuvem...`;
    status.style.color = 'var(--piacentini-orange)';

    try {
        const response = await fetch(API_URL, { method: 'POST', body: JSON.stringify(pacote) });
        const result = await response.json();
        if (result.sucesso) {
            status.className = "mt-4 fs-5 fw-bold text-center text-success";
            status.innerHTML = `✅ ${result.total} certificados gerados.`;
            if (destino === 'ZIP' && result.link) status.innerHTML += `<br><br><a href="${result.link}" target="_blank" class="btn btn-success btn-lg fw-bold">📥 BAIXAR ARQUIVO ZIP</a>`;
        } else throw new Error(result.erro);
    } catch (error) {
        status.className = "mt-4 fs-5 fw-bold text-center text-danger";
        status.innerHTML = `❌ Erro do Servidor: ${error.message}`;
    }
    btnZip.disabled = false; btnEmail.disabled = false;
}
