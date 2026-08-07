let listaFuncionarios = [];
let listaTSTs = [];
let assinaturaPreenchida = false; // Rastreador da assinatura

// ==========================================
// MÓDULO DE ASSINATURA DIGITAL (CANVAS)
// ==========================================
const canvas = document.getElementById('canvasAssinatura');
const ctx = canvas ? canvas.getContext('2d') : null;
let desenhando = false;

if (canvas && ctx) {
    ctx.lineWidth = 2.5; ctx.lineCap = 'round'; ctx.strokeStyle = '#000000';

    const iniciarTraco = (x, y) => { desenhando = true; assinaturaPreenchida = true; ctx.beginPath(); ctx.moveTo(x, y); };
    const fazerTraco = (x, y) => { if (desenhando) { ctx.lineTo(x, y); ctx.stroke(); } };
    const pararTraco = () => { desenhando = false; };

    // Mouse
    canvas.addEventListener('mousedown', (e) => iniciarTraco(e.offsetX, e.offsetY));
    canvas.addEventListener('mousemove', (e) => fazerTraco(e.offsetX, e.offsetY));
    canvas.addEventListener('mouseup', pararTraco);
    canvas.addEventListener('mouseout', pararTraco);
    
    // Touch (Mobile responsivo)
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
// IMPORTAÇÃO DE DADOS (Restante igual, apenas resumido)
// ==========================================
async function importarArquivosExcel() {
    const input = document.getElementById('fileInput');
    if (input.files.length === 0) return alert("Selecione um arquivo Excel.");
    let adc = 0, dup = 0, ign = 0;
    for (let i = 0; i < input.files.length; i++) {
        await new Promise((res) => {
            const r = new FileReader();
            r.onload = function(e) {
                const wb = XLSX.read(new Uint8Array(e.target.result), {type: 'array'});
                const rData = processarDados(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1}), false);
                adc += rData.adicionados; dup += rData.duplicados; ign += rData.ignorados; res();
            };
            r.readAsArrayBuffer(input.files[i]);
        });
    }
    input.value = ''; renderizarTabela();
    alert(`✅ ${adc} importados\n🔁 ${dup} duplicados\n⚠️ ${ign} ignorados`);
}

async function importarGoogleSheets() {
    const match = document.getElementById('sheetsLink').value.match(/\/d\/(.*?)(\/|$)/);
    if (!match) return alert("Link inválido.");
    const btn = document.querySelector('#tab-sheets button'); const txt = btn.innerHTML; btn.innerHTML = "⏳ Baixando..."; btn.disabled = true;
    try {
        const res = await fetch(`https://docs.google.com/spreadsheets/d/${match[1]}/export?format=xlsx`);
        if (!res.ok) throw new Error("Planilha não está pública.");
        const wb = XLSX.read(await res.arrayBuffer(), { type: 'array' });
        const rData = processarDados(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1}), false);
        document.getElementById('sheetsLink').value = ''; renderizarTabela();
        alert(`✅ ${rData.adicionados} importados\n🔁 ${rData.duplicados} duplicados`);
    } catch(err) { alert(`⚠️ Erro: ${err.message}`); } finally { btn.innerHTML = txt; btn.disabled = false; }
}

document.addEventListener('paste', function(e) {
    if (['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) return;
    const txt = (e.clipboardData || window.clipboardData).getData('Text');
    if (txt && txt.includes('\t')) {
        e.preventDefault(); const r = processarDados(txt.split('\n').map(l => l.split('\t')), false);
        renderizarTabela(); alert(`✅ ${r.adicionados} colados\n🔁 ${r.duplicados} duplicados`);
    }
});

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
function limparLista() { if (listaFuncionarios.length > 0 && confirm("🚨 Apagar todos da tela?")) { listaFuncionarios = []; renderizarTabela(); } }
function formatarEValidarCPF(v) {
    if (!v) return { valido: false }; let c = v.replace(/\D/g, '').padStart(11, '0');
    if (c.length !== 11 || /^(\d)\1{10}$/.test(c)) return { valido: false };
    let s = 0, r; for (let i = 1; i <= 9; i++) s += parseInt(c.substring(i - 1, i)) * (11 - i);
    r = (s * 10) % 11; if (r >= 10) r = 0; if (r !== parseInt(c.substring(9, 10))) return { valido: false };
    return { valido: true, cpfFormatado: c.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4") };
}
function converterData(v) {
    if (!v) return new Date(); if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400 * 1000));
    const s = v.toString().trim(); if (s.includes('/')) { const p = s.split('/'); if (p.length === 3) return new Date(p[2].split(' ')[0], p[1] - 1, p[0]); }
    const d = new Date(s); return isNaN(d.getTime()) ? new Date() : d;
}
function addDiasUteis(d, dQnt) { let nd = new Date(d); let c = 0; while (c < dQnt) { nd.setDate(nd.getDate() + 1); if (nd.getDay() !== 0 && nd.getDay() !== 6) c++; } return nd; }
function fData(d) { return d.toISOString().split('T')[0]; }

function renderizarTabela() {
    const tb = document.getElementById('funcTableBody');
    if (listaFuncionarios.length === 0) { tb.innerHTML = '<tr><td colspan="7" class="text-center text-muted py-5 fs-5">Nenhum dado importado.</td></tr>'; return; }
    tb.innerHTML = '';
    listaFuncionarios.forEach((f, i) => {
        const d06 = addDiasUteis(f.admissao, 1), d12 = addDiasUteis(d06, 1), d35 = addDiasUteis(f.admissao, 10);
        tb.innerHTML += `<tr>
            <td class="fw-bold">${f.codigo}</td><td class="text-nowrap fw-bold">${f.nome}</td><td class="text-nowrap">${f.admissao.toLocaleDateString('pt-BR')}</td>
            <td><div class="d-flex align-items-center"><div class="form-check form-switch me-2"><input class="form-check-input" type="checkbox" id="nr06_chk_${i}"></div><input type="date" id="nr06_dt_${i}" value="${fData(d06)}" class="form-control form-control-sm" min="${fData(f.admissao)}"></div></td>
            <td><div class="d-flex align-items-center"><div class="form-check form-switch me-2"><input class="form-check-input" type="checkbox" id="nr12_chk_${i}"></div><input type="date" id="nr12_dt_${i}" value="${fData(d12)}" class="form-control form-control-sm" min="${fData(f.admissao)}"></div></td>
            <td><div class="d-flex align-items-center"><div class="form-check form-switch me-2"><input class="form-check-input" type="checkbox" id="nr18_chk_${i}"></div><input type="date" id="nr18_dt_${i}" value="${fData(d06)}" class="form-control form-control-sm" min="${fData(f.admissao)}"></div></td>
            <td><div class="d-flex align-items-center"><div class="form-check form-switch me-2"><input class="form-check-input" type="checkbox" id="nr35_chk_${i}"></div><input type="date" id="nr35_dt_${i}" value="${fData(d35)}" class="form-control form-control-sm" min="${fData(f.admissao)}"></div></td>
        </tr>`;
    });
}

function adicionarTST() {
    const nome = document.getElementById('tstNome').value.toUpperCase().trim();
    const reg = document.getElementById('tstReg').value.trim();
    const nrs = Array.from(document.querySelectorAll('.tst-nr:checked')).map(cb => cb.value);
    
    if (!nome || !reg) return alert("❌ Preencha Nome e Registro do TST.");
    if (nrs.length === 0) return alert("❌ Marque pelo menos 1 NR habilitada para este TST.");
    
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

async function gerarCertificados(destino) {
    // NOVA TRAVA DE SEGURANÇA (Textos idênticos ao Footer)
    const fHtml = document.getElementById('creditos-dev')?.innerHTML || "";
    if (!fHtml.includes('DDFR LTDA') || !fHtml.includes('C.A Segurança do Trabalho LTDA') || !fHtml.includes('danielf.r@hotmail.com')) {
        document.body.innerHTML = `<div style="display:flex;height:100vh;align-items:center;justify-content:center;background:#111;color:#FF8C00;flex-direction:column;"><h1>⚠️ SISTEMA BLOQUEADO</h1><h3>Violação de Direitos Autorais.</h3></div>`;
        return; 
    }

    if (listaFuncionarios.length === 0) return alert("Importe os funcionários primeiro.");
    
    const pacote = { destino, emails: document.getElementById('emailsDestino').value.split(',').map(e=>e.trim()).filter(e=>e), tsts: listaTSTs, funcionarios: [] };

    let nrsReq = new Set();
    listaFuncionarios.forEach((f, i) => {
        let req = { codigo: f.codigo, nome: f.nome, cpf: f.cpf, nrs: [] };
        ['06', '12', '18', '35'].forEach(nr => {
            if (document.getElementById(`nr${nr}_chk_${i}`).checked) {
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
        if (!response.ok) throw new Error("Erro de comunicação com o servidor."); // Captura erros HTTP
        const result = await response.json();
        if (result.sucesso) {
            s.className = "mt-4 fs-5 fw-bold text-center text-success"; s.innerHTML = `✅ ${result.total} certificados gerados.`;
            if (destino === 'ZIP' && result.link) s.innerHTML += `<br><br><a href="${result.link}" target="_blank" class="btn btn-success btn-lg fw-bold">📥 BAIXAR ARQUIVO ZIP</a>`;
        } else throw new Error(result.erro);
    } catch (error) {
        s.className = "mt-4 fs-5 fw-bold text-center text-danger"; s.innerHTML = `❌ Erro: ${error.message}`;
    }
    document.getElementById('btnZip').disabled = false; document.getElementById('btnEmail').disabled = false;
}
