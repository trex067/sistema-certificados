let listaFuncionarios = [];
let listaTSTs = [];
let assinaturaPreenchida = false; 

// ==========================================
// MÓDULO DE ASSINATURA DIGITAL
// ==========================================
const canvas = document.getElementById('canvasAssinatura');
const ctx = canvas ? canvas.getContext('2d') : null;
let desenhando = false;

if (canvas && ctx) {
    ctx.lineWidth = 2.5; ctx.lineCap = 'round'; ctx.strokeStyle = '#000000';
    const iniciarTraco = (x, y) => { desenhando = true; assinaturaPreenchida = true; ctx.beginPath(); ctx.moveTo(x, y); };
    const fazerTraco = (x, y) => { if (desenhando) { ctx.lineTo(x, y); ctx.stroke(); } };
    const pararTraco = () => { desenhando = false; };
    canvas.addEventListener('mousedown', (e) => iniciarTraco(e.offsetX, e.offsetY));
    canvas.addEventListener('mousemove', (e) => fazerTraco(e.offsetX, e.offsetY));
    canvas.addEventListener('mouseup', pararTraco); canvas.addEventListener('mouseout', pararTraco);
    canvas.addEventListener('touchstart', (e) => {
        e.preventDefault(); const rect = canvas.getBoundingClientRect();
        iniciarTraco((e.touches[0].clientX - rect.left) * (canvas.width / rect.width), (e.touches[0].clientY - rect.top) * (canvas.height / rect.height));
    }, {passive: false});
    canvas.addEventListener('touchmove', (e) => {
        e.preventDefault(); const rect = canvas.getBoundingClientRect();
        fazerTraco((e.touches[0].clientX - rect.left) * (canvas.width / rect.width), (e.touches[0].clientY - rect.top) * (canvas.height / rect.height));
    }, {passive: false});
    canvas.addEventListener('touchend', pararTraco);
}
function limparCanvas() { if(ctx) { ctx.clearRect(0, 0, canvas.width, canvas.height); assinaturaPreenchida = false; } }

// ==========================================
// MÓDULO DE DADOS E TABELA (INCLUI MANUAL)
// ==========================================
function adicionarLinhaManual() {
    listaFuncionarios.push({ codigo: "", nome: "", cpf: "", admissao: new Date() });
    renderizarTabela();
}

// Funções para atualizar os dados editados diretamente na tabela
function atualizarDado(index, campo, valor) {
    if (campo === 'admissao') {
        const d = new Date(valor + 'T00:00:00'); 
        listaFuncionarios[index][campo] = isNaN(d.getTime()) ? new Date() : d;
        renderizarTabela(); // Re-renderiza para atualizar as datas dos cursos baseados na nova admissão
    } else {
        listaFuncionarios[index][campo] = valor.toUpperCase();
    }
}

function processarDados(matriz) {
    let adc = 0, dup = 0, ign = 0;
    for (let i = 0; i < matriz.length; i++) {
        const l = matriz[i]; if (!l || l.length < 3) continue;
        let cBruto = l[2] ? l[2].toString().trim() : "";
        if (cBruto.toLowerCase() === "cpf" || (l[1] && l[1].toString().toLowerCase() === "nome")) continue;
        let cValido = formatarEValidarCPF(cBruto);
        if (!cValido.valido) { ign++; continue; }
        if (listaFuncionarios.find(f => f.cpf === cValido.cpfFormatado)) { dup++; continue; }
        listaFuncionarios.push({ codigo: l[0] || "", nome: l[1] ? l[1].toString().toUpperCase().trim() : "", cpf: cValido.cpfFormatado, admissao: converterData(l[3]) });
        adc++;
    }
    return { adicionados: adc, duplicados: dup, ignorados: ign };
}

// Lógica de "Esbranquiçar" e Ativar/Desativar Datas
function toggleNrData(index, nr) {
    const isChecked = document.getElementById(`nr${nr}_chk_${index}`).checked;
    const inputData = document.getElementById(`nr${nr}_dt_${index}`);
    inputData.disabled = !isChecked;
    if (isChecked) {
        inputData.classList.remove('data-desativada');
        inputData.classList.add('data-ativada');
    } else {
        inputData.classList.remove('data-ativada');
        inputData.classList.add('data-desativada');
    }
}

function renderizarTabela() {
    const tb = document.getElementById('funcTableBody');
    if (listaFuncionarios.length === 0) { tb.innerHTML = '<tr><td colspan="8" class="text-center text-muted py-5 fs-5">Nenhum dado inserido.</td></tr>'; return; }
    tb.innerHTML = '';
    listaFuncionarios.forEach((f, i) => {
        const d06 = addDiasUteis(f.admissao, 1), d12 = addDiasUteis(d06, 1), d35 = addDiasUteis(f.admissao, 10);
        
        // Os campos de nome, cpf, etc. agora são INPUTS transparentes para permitir digitação livre!
        tb.innerHTML += `<tr>
            <td><input type="text" class="input-tabela fw-bold" value="${f.codigo}" onchange="atualizarDado(${i}, 'codigo', this.value)" placeholder="Cód"></td>
            <td class="text-nowrap"><input type="text" class="input-tabela fw-bold" value="${f.nome}" onchange="atualizarDado(${i}, 'nome', this.value)" placeholder="Nome Completo"></td>
            <td class="text-nowrap"><input type="text" class="input-tabela" value="${f.cpf}" onchange="atualizarDado(${i}, 'cpf', this.value)" placeholder="000.000.000-00"></td>
            <td class="text-nowrap"><input type="date" class="input-tabela" value="${fData(f.admissao)}" onchange="atualizarDado(${i}, 'admissao', this.value)"></td>
            
            <td class="col-nr"><div class="d-flex align-items-center"><div class="form-check form-switch me-2"><input class="form-check-input" type="checkbox" id="nr06_chk_${i}" onchange="toggleNrData(${i}, '06')"></div><input type="date" id="nr06_dt_${i}" value="${fData(d06)}" class="form-control form-control-sm data-desativada" disabled></div></td>
            <td class="col-nr"><div class="d-flex align-items-center"><div class="form-check form-switch me-2"><input class="form-check-input" type="checkbox" id="nr12_chk_${i}" onchange="toggleNrData(${i}, '12')"></div><input type="date" id="nr12_dt_${i}" value="${fData(d12)}" class="form-control form-control-sm data-desativada" disabled></div></td>
            <td class="col-nr"><div class="d-flex align-items-center"><div class="form-check form-switch me-2"><input class="form-check-input" type="checkbox" id="nr18_chk_${i}" onchange="toggleNrData(${i}, '18')"></div><input type="date" id="nr18_dt_${i}" value="${fData(d06)}" class="form-control form-control-sm data-desativada" disabled></div></td>
            <td class="col-nr"><div class="d-flex align-items-center"><div class="form-check form-switch me-2"><input class="form-check-input" type="checkbox" id="nr35_chk_${i}" onchange="toggleNrData(${i}, '35')"></div><input type="date" id="nr35_dt_${i}" value="${fData(d35)}" class="form-control form-control-sm data-desativada" disabled></div></td>
        </tr>`;
    });
}

// Outros Utilitários
async function importarArquivosExcel() {
    const input = document.getElementById('fileInput'); if (input.files.length === 0) return alert("Selecione um arquivo Excel.");
    let adc = 0, dup = 0;
    for (let i = 0; i < input.files.length; i++) {
        await new Promise((res) => { const r = new FileReader(); r.onload = function(e) { const wb = XLSX.read(new Uint8Array(e.target.result), {type: 'array'}); const rData = processarDados(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1})); adc += rData.adicionados; dup += rData.duplicados; res(); }; r.readAsArrayBuffer(input.files[i]); });
    }
    input.value = ''; renderizarTabela(); alert(`✅ ${adc} importados\n🔁 ${dup} duplicados ignorados`);
}
async function importarGoogleSheets() {
    const match = document.getElementById('sheetsLink').value.match(/\/d\/(.*?)(\/|$)/); if (!match) return alert("Link inválido.");
    const btn = document.querySelector('#tab-sheets button'); const txt = btn.innerHTML; btn.innerHTML = "⏳ Baixando..."; btn.disabled = true;
    try {
        const res = await fetch(`https://docs.google.com/spreadsheets/d/${match[1]}/export?format=xlsx`);
        if (!res.ok) throw new Error("Planilha não está pública.");
        const wb = XLSX.read(await res.arrayBuffer(), { type: 'array' }); const rData = processarDados(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1}));
        document.getElementById('sheetsLink').value = ''; renderizarTabela(); alert(`✅ ${rData.adicionados} importados`);
    } catch(err) { alert(`⚠️ Erro: ${err.message}`); } finally { btn.innerHTML = txt; btn.disabled = false; }
}
document.addEventListener('paste', function(e) {
    if (['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) return;
    const txt = (e.clipboardData || window.clipboardData).getData('Text');
    if (txt && txt.includes('\t')) { e.preventDefault(); const r = processarDados(txt.split('\n').map(l => l.split('\t'))); renderizarTabela(); alert(`✅ ${r.adicionados} colados`); }
});
function limparLista() { if (listaFuncionarios.length > 0 && confirm("🚨 Apagar todos da tela?")) { listaFuncionarios = []; renderizarTabela(); } }
function formatarEValidarCPF(v) {
    if (!v) return { valido: false }; let c = v.replace(/\D/g, '').padStart(11, '0');
    if (c.length !== 11 || /^(\d)\1{10}$/.test(c)) return { valido: false };
    let s = 0, r; for (let i = 1; i <= 9; i++) s += parseInt(c.substring(i - 1, i)) * (11 - i); r = (s * 10) % 11; if (r >= 10) r = 0; if (r !== parseInt(c.substring(9, 10))) return { valido: false };
    return { valido: true, cpfFormatado: c.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4") };
}
function converterData(v) {
    if (!v) return new Date(); if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400 * 1000));
    const s = v.toString().trim(); if (s.includes('/')) { const p = s.split('/'); if (p.length === 3) return new Date(p[2].split(' ')[0], p[1] - 1, p[0]); }
    const d = new Date(s); return isNaN(d.getTime()) ? new Date() : d;
}
function addDiasUteis(d, dQnt) { let nd = new Date(d); let c = 0; while (c < dQnt) { nd.setDate(nd.getDate() + 1); if (nd.getDay() !== 0 && nd.getDay() !== 6) c++; } return nd; }
function fData(d) { return d.toISOString().split('T')[0]; }

// ==========================================
// MÓDULO DOS TSTs (Com travas de Regras)
// ==========================================
function adicionarTST() {
    const nome = document.getElementById('tstNome').value.toUpperCase().trim();
    const reg = document.getElementById('tstReg').value.trim();
    const nrs = Array.from(document.querySelectorAll('.tst-nr:checked')).map(cb => cb.value);
    
    if (!nome) return alert("❌ Preencha o Nome do TST.");
    
    // VALIDAÇÃO 1: Formato do Registro 000000/UF
    if (!/^[A-Z0-9-]+\/[A-Z]{2}$/.test(reg)) {
        return alert("❌ O Registro deve terminar com a barra e a sigla do estado.\n\nExemplos válidos:\n12345/PR\nTE-999/SP");
    }

    if (nrs.length === 0) return alert("❌ Marque pelo menos 1 NR habilitada para este TST.");
    
    // VALIDAÇÃO 2: Apenas UM instrutor por NR
    for (let nrNova of nrs) {
        for (let tstExistente of listaTSTs) {
            if (tstExistente.nrs.includes(nrNova)) {
                return alert(`⚠️ CONFLITO DE NORMAS!\n\nA norma ${nrNova} já possui um instrutor cadastrado (${tstExistente.nome}).\n\nO sistema permite apenas 1 instrutor assinando por norma. Desmarque a ${nrNova} ou remova o TST anterior.`);
            }
        }
    }
    
    let assinatura = null;
    if (assinaturaPreenchida && canvas) assinatura = canvas.toDataURL('image/png');

    listaTSTs.push({ nome, registro: reg, nrs, assinatura });
    
    document.getElementById('tstNome').value = ''; document.getElementById('tstReg').value = '';
    document.querySelectorAll('.tst-nr').forEach(cb => cb.checked = false);
    limparCanvas(); atualizarListaTST();
}

function atualizarListaTST() {
    document.getElementById('tstList').innerHTML = listaTSTs.map((t, i) => `
        <div class="badge bg-secondary p-3 d-flex align-items-center flex-column flex-md-row fs-6 border border-warning text-white">
            <span>👤 ${t.nome} (${t.registro}) <br class="d-md-none"> 📌 NRs: ${t.nrs.join(', ')} ${t.assinatura ? '🖋️ (Com Assinatura)' : ''}</span>
            <div class="mt-2 mt-md-0 ms-md-3">
                <button class="btn btn-sm btn-light text-danger py-0 fw-bold" onclick="listaTSTs.splice(${i}, 1); atualizarListaTST();">✖ Remover</button>
            </div>
        </div>`).join('');
}

// ==========================================
// MÓDULO DE GERAÇÃO E ENVIO
// ==========================================
async function gerarCertificados(destino) {
    const fHtml = document.getElementById('creditos-dev')?.innerHTML || "";
    if (!fHtml.includes('DDFR LTDA') || !fHtml.includes('C.A Segurança do Trabalho LTDA') || !fHtml.includes('danielf.r@hotmail.com')) {
        document.body.innerHTML = `<div style="display:flex;height:100vh;align-items:center;justify-content:center;background:#111;color:#FF8C00;flex-direction:column; text-align:center;"><h1>⚠️ SISTEMA BLOQUEADO</h1><h3>Violação de Direitos Autorais. Os créditos originais foram alterados.</h3></div>`;
        return; 
    }

    if (listaFuncionarios.length === 0) return alert("A tabela de funcionários está vazia.");
    
    const pacote = { destino, emails: document.getElementById('emailsDestino').value.split(',').map(e=>e.trim()).filter(e=>e), tsts: listaTSTs, funcionarios: [] };

    let nrsReq = new Set();
    listaFuncionarios.forEach((f, i) => {
        if(!f.nome || !f.cpf) return; // ignora linhas em branco deixadas manualmente
        let req = { codigo: f.codigo, nome: f.nome, cpf: f.cpf, nrs: [] };
        ['06', '12', '18', '35'].forEach(nr => {
            const elChk = document.getElementById(`nr${nr}_chk_${i}`);
            if (elChk && elChk.checked) {
                req.nrs.push({ tipo: `NR ${nr}`, data: document.getElementById(`nr${nr}_dt_${i}`).value.split('-').reverse().join('/') });
                nrsReq.add(`NR ${nr}`);
            }
        });
        if (req.nrs.length > 0) pacote.funcionarios.push(req);
    });

    if (pacote.funcionarios.length === 0) return alert("⚠️ Você não marcou NENHUMA norma na tabela (Ligue as chaves verdes).");

    let nrsCob = new Set(); listaTSTs.forEach(t => t.nrs.forEach(n => nrsCob.add(n)));
    for (let nr of nrsReq) if (!nrsCob.has(nr)) return alert(`🚨 GERAÇÃO BLOQUEADA!\nFalta TST para a norma: ${nr}`);

    const s = document.getElementById('statusProcessamento');
    document.getElementById('btnZip').disabled = true; document.getElementById('btnEmail').disabled = true;
    s.innerHTML = `⏳ Gerando documentos na Nuvem...`;

    try {
        const response = await fetch(API_URL, { method: 'POST', body: JSON.stringify(pacote) });
        if (!response.ok) throw new Error("A API não respondeu. Verifique se o Google Script está configurado como 'Qualquer Pessoa'.");
        const result = await response.json();
        if (result.sucesso) {
            s.className = "mt-4 fs-5 fw-bold text-center text-success"; s.innerHTML = `✅ ${result.total} certificados gerados com sucesso.`;
            if (destino === 'ZIP' && result.link) s.innerHTML += `<br><br><a href="${result.link}" target="_blank" class="btn btn-success btn-lg fw-bold">📥 BAIXAR ARQUIVO ZIP</a>`;
        } else throw new Error(result.erro);
    } catch (error) {
        s.className = "mt-4 fs-5 fw-bold text-center text-danger"; s.innerHTML = `❌ Erro: ${error.message}`;
    }
    document.getElementById('btnZip').disabled = false; document.getElementById('btnEmail').disabled = false;
}
