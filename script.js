// ══════════════════════════════════════════
// SeторMap — script.js
// ══════════════════════════════════════════

// ── ESTADO GLOBAL ──
let DATA = [];
let SETOR_COLORS = {};
let allBricks = [];
let brickColorMap = {};
let brickSetorMap = {};
let mode = 'setor';
let hiddenSetores = new Set();
let hiddenBricks = new Set();
let failed = 0;
let failedList = [];
let markerStore = {};
let satOn = false;
let layerGroup, map, tileOSM, tileSat;
let mapInitialized = false;

const BRICK_PALETTE = [
  '#ef4444','#f97316','#f59e0b','#eab308','#84cc16','#22c55e','#10b981','#14b8a6',
  '#06b6d4','#0ea5e9','#3b82f6','#6366f1','#8b5cf6','#a855f7','#d946ef','#ec4899',
  '#f43f5e','#dc2626','#ea580c','#d97706','#ca8a04','#65a30d','#16a34a','#059669',
  '#0d9488','#0891b2','#0284c7','#2563eb','#4f46e5','#7c3aed','#9333ea','#c026d3',
  '#db2777','#e11d48','#b91c1c','#c2410c','#b45309','#a16207','#4d7c0f','#15803d',
  '#047857','#0f766e','#0e7490','#0369a1','#1d4ed8','#4338ca','#6d28d9','#7e22ce',
  '#a21caf','#be185d','#be123c','#991b1b','#9a3412','#92400e','#854d0e','#3f6212',
  '#166534','#065f46','#115e59','#155e75','#075985','#1e40af','#3730a3','#5b21b6',
  '#fb923c','#fbbf24','#fde047','#bef264','#86efac','#6ee7b7','#5eead4','#67e8f9',
  '#7dd3fc','#93c5fd','#a5b4fc','#c4b5fd','#d8b4fe','#f0abfc','#f9a8d4','#fda4af',
  '#fca5a5','#fdba74','#fcd34d','#fef08a','#d9f99d','#bbf7d0','#a7f3d0','#99f6e4',
  '#a5f3fc','#bae6fd','#bfdbfe','#c7d2fe','#ddd6fe','#e9d5ff','#f5d0fe','#fbcfe8',
  '#fecaca','#fed7aa','#fef3c7','#fef9c3','#ecfccb','#dcfce7','#d1fae5','#ccfbf1',
  '#cffafe','#e0f2fe','#dbeafe','#e0e7ff','#ede9fe','#f3e8ff','#fae8ff','#fce7f3',
  '#fee2e2','#ffedd5','#fef3c7','#0c4a6e','#1e3a8a','#312e81','#4c1d95','#5b21b6',
  '#701a75','#831843','#881337','#7f1d1d','#7c2d12','#78350f','#713f12','#365314',
];

const SETOR_PALETTE = [
  '#34d399','#60a5fa','#f472b6','#a78bfa','#fb923c','#facc15',
  '#4ade80','#38bdf8','#f87171','#c084fc','#86efac','#67e8f9',
];

// ══════════════════════════════════════════
// HOME TAB SWITCH
// ══════════════════════════════════════════
function switchHomeTab(tab) {
  const isEnr = tab === 'enr';
  document.getElementById('tabMapa').classList.toggle('on', !isEnr);
  document.getElementById('tabEnr').classList.toggle('on', isEnr);
  document.getElementById('home-mapa-panel').style.display = isEnr ? 'none' : '';
  document.getElementById('home-enr-panel').style.display  = isEnr ? 'block' : 'none';
}

// ══════════════════════════════════════════
// HOME ENRICHER
// ══════════════════════════════════════════
let hePdvMap = {}, heQueue = [], heIdx = 0, heRunning = false, heStopped = false;

(function heInit() {
  const drop = document.getElementById('heDropZone');
  const fi   = document.getElementById('heFileInput');
  if (!drop || !fi) return;
  drop.addEventListener('dragover',  e => { e.preventDefault(); drop.classList.add('drag'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('drag'));
  drop.addEventListener('drop', e => {
    e.preventDefault(); drop.classList.remove('drag');
    if (e.dataTransfer.files[0]) heLoad(e.dataTransfer.files[0]);
  });
  fi.addEventListener('change', () => { if (fi.files[0]) heLoad(fi.files[0]); fi.value = ''; });
})();

function heLoad(file) {
  heMsg('');
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!rows.length) { heMsg('Planilha vazia.', true); return; }

      const keys = Object.keys(rows[0]);
      const fc   = (...c) => c.map(x => keys.find(y => y.trim().toLowerCase() === x.toLowerCase())).find(Boolean);
      const cS   = fc('Setor','setor');
      const cB   = fc('Brick','brick');
      const cP   = fc('PDV','pdv');
      if (!cS || !cB || !cP) { heMsg('Colunas obrigatórias não encontradas: Setor, Brick, PDV.', true); return; }

      hePdvMap = {};
      rows.forEach(row => {
        const raw   = String(row[cP] || '').trim();
        const setor = String(row[cS] || '').trim();
        const brick = String(row[cB] || '').trim();
        if (!raw || !setor || !brick) return;

        let cnpj = '', nome = '';
        if (raw.includes('|')) {
          const parts = raw.split('|').map(s => s.trim());
          cnpj = parts[0].replace(/\D/g, '');
          nome = parts[1] || '';
        } else {
          cnpj = raw.replace(/\D/g, '').substring(0, 14);
          nome = raw;
        }
        if (!cnpj || cnpj.length < 7) return;

        const key = cnpj + '||' + brick;
        if (!hePdvMap[key]) {
          hePdvMap[key] = { setor, brick, nome, cnpj, cidade: '', telefone: '', status: 'pendente' };
        }
      });

      const uniq = Object.keys(hePdvMap);
      if (!uniq.length) { heMsg('Nenhum PDV válido encontrado.', true); return; }

      document.getElementById('heBadgeFile').textContent  = file.name;
      document.getElementById('heBadgeRows').textContent  = rows.length + ' linhas';
      document.getElementById('heBadgeUniq').textContent  = uniq.length + ' PDVs únicos';
      document.getElementById('heInfo').style.display     = 'flex';
      document.getElementById('heActions').style.display  = 'flex';
      document.getElementById('heStats').style.display    = 'grid';
      document.getElementById('heTableWrap').style.display = 'block';
      document.getElementById('heTotal').textContent      = uniq.length;
      document.getElementById('hePend').textContent       = uniq.length;
      document.getElementById('heOk').textContent         = '0';
      document.getElementById('heErr').textContent        = '0';
      heQueue = [...uniq]; heIdx = 0;
      heRenderTable();
    } catch(err) { heMsg('Erro ao ler arquivo: ' + err.message, true); }
  };
  reader.readAsArrayBuffer(file);
}

async function heFetchCNPJ(cnpj) {
  // Tenta BrasilAPI com retry e backoff
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      if (attempt > 0) await sleep(1500 * attempt);
      const r = await fetch('https://brasilapi.com.br/api/cnpj/v1/' + cnpj);
      if (r.status === 429 || r.status === 503) { await sleep(3000 + attempt * 2000); continue; }
      if (r.ok) {
        const d   = await r.json();
        const tel = d.ddd_telefone_1 ? d.ddd_telefone_1.replace(/\D/g, '') : '';
        return {
          nome:     d.razao_social || d.nome_fantasia || '',
          cidade:   (d.municipio || '') + (d.uf ? ' / ' + d.uf : ''),
          telefone: tel.length >= 10
            ? '(' + tel.slice(0,2) + ') ' + (tel.length === 11 ? tel.slice(2,7)+'-'+tel.slice(7) : tel.slice(2,6)+'-'+tel.slice(6))
            : tel
        };
      }
      break;
    } catch { /* retry */ }
  }
  // Fallback: cnpj.ws
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      if (attempt > 0) await sleep(1500 * attempt);
      const r = await fetch('https://publica.cnpj.ws/cnpj/' + cnpj);
      if (r.status === 429 || r.status === 503) { await sleep(3000 + attempt * 2000); continue; }
      if (!r.ok) break;
      const d   = await r.json();
      const est = d.estabelecimento;
      if (!est) break;
      const tel = est.telefone1 ? est.telefone1.replace(/\D/g, '') : '';
      return {
        nome:     d.razao_social || '',
        cidade:   (est.cidade ? est.cidade.nome : '') + (est.estado ? ' / ' + est.estado.sigla : ''),
        telefone: tel.length >= 10
          ? '(' + tel.slice(0,2) + ') ' + (tel.length === 11 ? tel.slice(2,7)+'-'+tel.slice(7) : tel.slice(2,6)+'-'+tel.slice(6))
          : tel
      };
    } catch { /* retry */ }
  }
  return null;
}

async function heStart() {
  if (heRunning) return;
  heRunning = true; heStopped = false;
  document.getElementById('heBtnStart').style.display  = 'none';
  document.getElementById('heBtnStop').style.display   = '';
  document.getElementById('heBtnResume').style.display = 'none';
  document.getElementById('heProg').style.display      = 'block';
  document.getElementById('heBtnExport').disabled      = false;
  heMsg('');
  await heRunQueue();
}

function heStop() {
  heStopped = true; heRunning = false;
  document.getElementById('heBtnStop').style.display   = 'none';
  document.getElementById('heBtnResume').style.display = '';
  heMsg('Pausado em ' + heIdx + ' de ' + heQueue.length + '. Clique em Retomar.');
}

async function heResume() {
  if (heRunning) return;
  heRunning = true; heStopped = false;
  document.getElementById('heBtnResume').style.display = 'none';
  document.getElementById('heBtnStop').style.display   = '';
  heMsg('');
  await heRunQueue();
}

async function heRunQueue() {
  const total = heQueue.length;
  let consecutiveErrors = 0;

  for (; heIdx < total; heIdx++) {
    if (heStopped) return;
    const key = heQueue[heIdx];
    const pdv = hePdvMap[key];
    if (!pdv || pdv.status !== 'pendente') continue;

    if (!pdv.cnpj || pdv.cnpj.length < 11) {
      pdv.status = 'sem cnpj';
      consecutiveErrors = 0;
    } else {
      const res = await heFetchCNPJ(pdv.cnpj);
      if (res) {
        if (res.nome)     pdv.nome     = res.nome;
        pdv.cidade        = res.cidade   || '';
        pdv.telefone      = res.telefone || '';
        pdv.status        = 'ok';
        consecutiveErrors = 0;
      } else {
        pdv.status = 'erro';
        consecutiveErrors++;
        // 5 erros seguidos = possível rate limit → pausa 10s e retenta
        if (consecutiveErrors >= 5) {
          heMsg('⚠️ Muitos erros seguidos — aguardando 10s para evitar bloqueio da API…');
          await sleep(10000);
          consecutiveErrors = 0;
          const recent = heQueue.slice(Math.max(0, heIdx - 4), heIdx + 1);
          recent.forEach(k => { if (hePdvMap[k] && hePdvMap[k].status === 'erro') hePdvMap[k].status = 'pendente'; });
          heIdx = Math.max(0, heIdx - 4) - 1;
          heMsg('');
          continue;
        }
      }
    }

    const pct = Math.round((heIdx + 1) / total * 100);
    document.getElementById('heProgFill').style.width = pct + '%';
    document.getElementById('heProgPct').textContent  = pct + '%';
    document.getElementById('heProgTxt').textContent  = 'Buscando… ' + (heIdx + 1) + ' / ' + total;
    heRenderTable();
    heUpdateStats();
    await sleep(600);
  }

  heRunning = false;
  document.getElementById('heBtnStop').style.display   = 'none';
  document.getElementById('heBtnStart').style.display  = 'none';
  document.getElementById('heBtnResume').style.display = 'none';
  const ok = Object.values(hePdvMap).filter(p => p.status === 'ok').length;
  heMsg('✓ Concluído! ' + ok + ' de ' + total + ' PDVs encontrados. Clique em Baixar planilha.');
}

function heExport() {
  const pdvs = Object.values(hePdvMap);
  const wb   = XLSX.utils.book_new();
  // Colunas: Setor · Brick · Nome do PDV · CNPJ · Cidade · Telefone
  const wsData = [['Setor', 'Brick', 'Nome do PDV', 'CNPJ', 'Cidade', 'Telefone']];
  pdvs.forEach(p => {
    const cnpjFmt = p.cnpj.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
    wsData.push([p.setor, p.brick, p.nome, cnpjFmt, p.cidade, p.telefone]);
  });
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [32, 30, 40, 20, 24, 16].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws, 'PDVs Enriquecidos');
  // Aba não encontrados
  const ws2 = [['CNPJ', 'Nome', 'Setor', 'Brick', 'Status']];
  pdvs.filter(p => p.status !== 'ok').forEach(p => ws2.push([p.cnpj, p.nome, p.setor, p.brick, p.status]));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ws2), 'Não encontrados');
  XLSX.writeFile(wb, 'PDVs_enriquecidos.xlsx');
}

function heRenderTable() {
  const tbody   = document.getElementById('heTbody');
  const entries = Object.values(hePdvMap).slice(0, 150);
  const pill    = s =>
    s === 'ok'       ? '<span style="color:#34d399;font-size:9px;font-weight:700">✓ ok</span>'      :
    s === 'erro'     ? '<span style="color:#f87171;font-size:9px;font-weight:700">✗ erro</span>'    :
    s === 'sem cnpj' ? '<span style="color:#64748b;font-size:9px">s/cnpj</span>'                    :
                       '<span style="color:#fbbf24;font-size:9px">…</span>';
  tbody.innerHTML = entries.map(p => {
    const setorShort = p.setor.replace(/^\d+\s*-\s*/, '');
    const brickShort = p.brick.replace(/^\d+\s*-\s*[^-]+-/i, '');
    const cnpjFmt   = p.cnpj.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
    return '<tr>'
      + '<td title="' + p.setor + '">' + setorShort + '</td>'
      + '<td title="' + p.brick + '">' + brickShort + '</td>'
      + '<td title="' + p.nome  + '">' + (p.nome     || '—') + '</td>'
      + '<td>' + cnpjFmt + '</td>'
      + '<td>' + (p.cidade   || '—') + '</td>'
      + '<td>' + (p.telefone || '—') + '</td>'
      + '<td>' + pill(p.status) + '</td>'
      + '</tr>';
  }).join('');
  const total = Object.keys(hePdvMap).length;
  if (total > 150) tbody.innerHTML += '<tr><td colspan="7" style="text-align:center;color:#475569;font-size:10px;padding:8px">… mais ' + (total - 150) + ' PDVs na planilha final</td></tr>';
}

function heUpdateStats() {
  const v = Object.values(hePdvMap);
  document.getElementById('heOk').textContent   = v.filter(p => p.status === 'ok').length;
  document.getElementById('heErr').textContent  = v.filter(p => p.status === 'erro' || p.status === 'sem cnpj').length;
  document.getElementById('hePend').textContent = v.filter(p => p.status === 'pendente').length;
}

function heReset() {
  heStopped = true; heRunning = false;
  hePdvMap = {}; heQueue = []; heIdx = 0;
  ['heInfo','heActions','heStats'].forEach(id => document.getElementById(id).style.display = 'none');
  document.getElementById('heProg').style.display      = 'none';
  document.getElementById('heTableWrap').style.display = 'none';
  document.getElementById('heBtnStart').style.display  = '';
  document.getElementById('heBtnStop').style.display   = 'none';
  document.getElementById('heBtnResume').style.display = 'none';
  document.getElementById('heBtnExport').disabled      = true;
  document.getElementById('heTbody').innerHTML         = '';
  document.getElementById('heFileInput').value         = '';
  heMsg('');
}

function heMsg(txt, isErr) {
  const el = document.getElementById('heMsg');
  el.textContent = txt;
  el.className   = 'he-msg' + (isErr ? ' err' : '');
}

// ══════════════════════════════════════════
// UPLOAD / FILE PROCESSING
// ══════════════════════════════════════════
const dropZone  = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const errEl     = document.getElementById('upload-error');
const loadEl    = document.getElementById('upload-loading');

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('drag');
  if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0], 'upload');
});
fileInput.addEventListener('change', () => { if (fileInput.files[0]) processFile(fileInput.files[0], 'upload'); });

const modalDrop      = document.getElementById('modal-drop');
const modalFileInput = document.getElementById('modal-file-input');
modalDrop.addEventListener('dragover', e => { e.preventDefault(); modalDrop.classList.add('drag'); });
modalDrop.addEventListener('dragleave', () => modalDrop.classList.remove('drag'));
modalDrop.addEventListener('drop', e => {
  e.preventDefault(); modalDrop.classList.remove('drag');
  if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0], 'modal');
});
modalFileInput.addEventListener('change', () => { if (modalFileInput.files[0]) processFile(modalFileInput.files[0], 'modal'); });

function processFile(file, ctx) {
  ctx = ctx || 'upload';
  const errElCtx  = ctx === 'modal' ? document.getElementById('modal-error')  : errEl;
  const loadElCtx = ctx === 'modal' ? document.getElementById('modal-loading') : loadEl;
  errElCtx.style.display  = 'none';
  loadElCtx.style.display = 'block';

  function showErr(msg) {
    errElCtx.textContent    = msg;
    errElCtx.style.display  = 'block';
    loadElCtx.style.display = 'none';
  }

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!rows.length) { showErr('Planilha vazia ou sem dados.'); return; }

      const keys = Object.keys(rows[0]);
      function findCol(...candidates) {
        const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
        for (const c of candidates) {
          const found = keys.find(k => norm(k) === norm(c));
          if (found) return found;
        }
        return null;
      }

      const colSetor = findCol('Setor','setor','SETOR');
      const colBrick = findCol('Brick','brick','BRICK');
      const colPDV   = findCol('PDV','pdv','Pdv');
      const colEnd   = findCol('Endereço','endereco','Endereco','ENDEREÇO','ENDERECO','endereço');
      const colCNPJ  = findCol('CNPJ (só dígitos)','CNPJ (so digitos)','cnpj (só dígitos)');
      const colBand  = findCol('Bandeira','bandeira','BANDEIRA');

      if (!colSetor || !colBrick || !colPDV) {
        showErr('Colunas obrigatórias não encontradas: Setor, Brick, PDV.\n\nColunas encontradas: ' + keys.join(', '));
        return;
      }

      const pdvMap = {};
      for (const row of rows) {
        const setor    = String(row[colSetor] || '').trim();
        const brick    = String(row[colBrick] || '').trim();
        const pdvRaw   = String(row[colPDV]   || '').trim();
        const bandeira = colBand ? String(row[colBand] || '').trim() : 'N/I';
        if (!setor || !brick || !pdvRaw) continue;

        let cnpj = colCNPJ ? String(row[colCNPJ] || '').trim().replace(/\D/g, '') : '';
        let nome = '', bairro = '', cidade = '', uf = '';

        if (pdvRaw.includes('|')) {
          const parts = pdvRaw.split('|').map(s => s.trim());
          if (!cnpj) cnpj = parts[0].replace(/\D/g, '');
          nome = parts[1] || '';
          const loc = parts[2] || '';
          const di  = loc.lastIndexOf(' - ');
          if (di >= 0) {
            bairro = loc.substring(0, di).trim();
            const cu = loc.substring(di + 3).trim();
            const si = cu.indexOf('/');
            if (si >= 0) { cidade = cu.substring(0, si).trim(); uf = cu.substring(si + 1).trim(); }
            else cidade = cu;
          } else bairro = loc;
        } else {
          if (!cnpj) cnpj = pdvRaw.replace(/\D/g, '').substring(0, 14);
          nome = pdvRaw;
        }

        if (!cnpj || cnpj.length < 7) continue;
        const endereco = colEnd ? String(row[colEnd] || '').trim() : '';

        // Extrai cidade do brick (fonte de verdade)
        let cidadeBrick = cidade, ufBrick = uf;
        const bm = brick.match(/^\d+\s*-\s*([^-]+)/);
        if (bm) {
          const raw = bm[1].trim().toUpperCase();
          const MAP = {
            'PORTO ALEGRE': ['Porto Alegre','RS'], 'SAO PAULO': ['São Paulo','SP'],
            'CURITIBA': ['Curitiba','PR'], 'FLORIANOPOLIS': ['Florianópolis','SC'],
            'CAXIAS DO SUL': ['Caxias do Sul','RS'], 'CANOAS': ['Canoas','RS'],
            'NOVO HAMBURGO': ['Novo Hamburgo','RS'], 'SAO LEOPOLDO': ['São Leopoldo','RS'],
            'PELOTAS': ['Pelotas','RS'], 'SANTA MARIA': ['Santa Maria','RS'],
            'GRAVATAI': ['Gravataí','RS'], 'VIAMAO': ['Viamão','RS'],
            'ALVORADA': ['Alvorada','RS'], 'CACHOEIRINHA': ['Cachoeirinha','RS'],
            'ESTEIO': ['Esteio','RS'], 'SAPUCAIA DO SUL': ['Sapucaia do Sul','RS'],
          };
          if (MAP[raw]) { cidadeBrick = MAP[raw][0]; ufBrick = MAP[raw][1]; }
          else { cidadeBrick = raw.split(' ').map(w => w.charAt(0)+w.slice(1).toLowerCase()).join(' '); ufBrick = uf || 'RS'; }
        }

        const key = cnpj + '||' + brick;
        if (pdvMap[key]) {
          if (endereco && !pdvMap[key].endereco) pdvMap[key].endereco = endereco;
        } else {
          pdvMap[key] = { cnpj, nome, bandeira, endereco, bairro, cidade: cidadeBrick, uf: ufBrick, brick, setor };
        }
      }

      const parsed = Object.values(pdvMap);
      if (!parsed.length) { showErr('Nenhuma linha válida encontrada.'); return; }

      const cidades = parsed.map(p => p.cidade).filter(Boolean);
      const cidadeMaisComum = cidades.length
        ? cidades.sort((a,b) => cidades.filter(x=>x===b).length - cidades.filter(x=>x===a).length)[0]
        : '';

      const setoresUnicos = [...new Set(parsed.map(p => p.setor))].sort();
      const setorColors = {};
      setoresUnicos.forEach((s, i) => {
        const isVago  = s.toUpperCase().includes('VAGO');
        const color   = isVago ? '#94a3b8' : SETOR_PALETTE[i % SETOR_PALETTE.length];
        const dashIdx = s.indexOf(' - ');
        const label   = dashIdx >= 0 ? s.substring(dashIdx + 3) : s;
        setorColors[s] = { color, label };
      });

      loadElCtx.style.display = 'none';
      salvarEIniciar(parsed, setorColors, cidadeMaisComum, file.name);
    } catch(err) { showErr('Erro ao ler o arquivo: ' + err.message); }
  };
  reader.onerror = () => showErr('Erro ao ler o arquivo.');
  reader.readAsArrayBuffer(file);
}

function salvarEIniciar(parsed, setorColors, cidade, nomeArquivo) {
  try {
    const payload = JSON.stringify({ parsed, setorColors, cidade, nomeArquivo, savedAt: Date.now() });
    localStorage.setItem('setormap_data', payload);
  } catch(e) { console.warn('localStorage cheio:', e); }
  fecharModal();
  iniciarApp(parsed, setorColors, cidade);
}

// ══════════════════════════════════════════
// APP INIT
// ══════════════════════════════════════════
async function iniciarApp(dados, setorColors, cidade) {
  DATA = dados;
  SETOR_COLORS = setorColors;
  failed = 0; failedList = []; markerStore = {};
  hiddenSetores = new Set(); hiddenBricks = new Set();

  allBricks = [...new Set(DATA.map(p => p.brick))].sort();
  brickColorMap = {};
  allBricks.forEach((b, i) => { brickColorMap[b] = BRICK_PALETTE[i % BRICK_PALETTE.length]; });

  const brickSetorCount = {};
  DATA.forEach(p => {
    if (!brickSetorCount[p.brick]) brickSetorCount[p.brick] = {};
    brickSetorCount[p.brick][p.setor] = (brickSetorCount[p.brick][p.setor] || 0) + 1;
  });
  brickSetorMap = {};
  Object.keys(brickSetorCount).forEach(b => {
    const top = Object.entries(brickSetorCount[b]).sort((a,b)=>b[1]-a[1])[0][0];
    brickSetorMap[b] = top;
  });

  const nBricks  = allBricks.length;
  const nSetores = Object.keys(SETOR_COLORS).length;
  document.getElementById('hdr-info').innerHTML =
    (cidade || 'PDVs') + ' &nbsp;|&nbsp; <b>' + DATA.length + '</b> PDVs · <b>' + nBricks + '</b> bricks · <b>' + nSetores + '</b> setores';
  document.getElementById('sTotal').textContent = DATA.length;
  document.getElementById('sPend').textContent  = DATA.length;

  document.getElementById('upload-screen').style.display = 'none';
  document.getElementById('app').style.display = 'flex';

  if (!mapInitialized) {
    map = L.map('map', { zoomControl: false }).setView([-15.78, -47.93], 4);
    L.control.zoom({ position: 'bottomright' }).addTo(map);
    tileOSM = L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', { attribution: '&copy; OSM &copy; CARTO', maxZoom: 19 });
    tileSat = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', { attribution: '&copy; Esri', maxZoom: 19 });
    tileOSM.addTo(map);
    layerGroup = L.layerGroup().addTo(map);
    mapInitialized = true;
  } else {
    layerGroup.clearLayers();
  }

  const temCache = DATA.length > 0 && DATA.every(p => p._lat && p._lng);
  document.getElementById('status').className  = '';
  document.getElementById('stxt').textContent  = temCache ? 'Carregando mapa salvo…' : 'Iniciando geocodificação…';
  document.getElementById('prog-fill').style.width = '0';
  document.getElementById('sLoc').textContent  = '0';
  document.getElementById('sFail').textContent = '0';

  buildSetorLegend();
  await runGeo(cidade);
}

// ══════════════════════════════════════════
// MODAL
// ══════════════════════════════════════════
function abrirModalCarregar() {
  const saved   = getSaved();
  const savedEl = document.getElementById('modalSaved');
  const divider = document.getElementById('modalDivider');
  document.getElementById('modal-error').style.display   = 'none';
  document.getElementById('modal-loading').style.display = 'none';
  document.getElementById('modal-file-input').value = '';

  if (saved) {
    const d     = new Date(saved.savedAt);
    const dtStr = d.toLocaleDateString('pt-BR') + ' às ' + d.toLocaleTimeString('pt-BR', {hour:'2-digit',minute:'2-digit'});
    document.getElementById('modalSavedInfo').innerHTML =
      '<b>' + (saved.nomeArquivo || 'planilha') + '</b><br>' +
      saved.parsed.length + ' PDVs · ' + Object.keys(saved.setorColors).length + ' setores · salvo em ' + dtStr;
    savedEl.style.display = 'block'; divider.style.display = 'block';
  } else {
    savedEl.style.display = 'none'; divider.style.display = 'none';
  }
  document.getElementById('modal-overlay').classList.add('open');
}
function fecharModal() { document.getElementById('modal-overlay').classList.remove('open'); }
function fecharModalSeOverlay(e) { if (e.target === document.getElementById('modal-overlay')) fecharModal(); }

function getSaved() {
  try { const r = localStorage.getItem('setormap_data'); return r ? JSON.parse(r) : null; }
  catch { return null; }
}

function persistirCoordenadas() {
  try {
    const existing = getSaved();
    if (!existing) return;
    localStorage.setItem('setormap_data', JSON.stringify({ ...existing, parsed: DATA, coordsSavedAt: Date.now() }));
  } catch(e) { console.warn('Erro ao persistir coords:', e); }
}

function recarregarSalvo() {
  const saved = getSaved();
  if (!saved) return;
  fecharModal();
  iniciarApp(saved.parsed, saved.setorColors, saved.cidade);
}

function limparSalvo() {
  if (!confirm('Apagar dados salvos? Será necessário carregar uma nova planilha.')) return;
  localStorage.removeItem('setormap_data');
  fecharModal();
  document.getElementById('app').style.display = 'none';
  document.getElementById('upload-screen').style.display = 'flex';
  document.getElementById('upload-error').style.display  = 'none';
  document.getElementById('upload-loading').style.display = 'none';
}

// Startup automático
(function startup() {
  const saved = getSaved();
  if (saved && saved.parsed && saved.parsed.length) {
    document.getElementById('upload-screen').style.display = 'none';
    iniciarApp(saved.parsed, saved.setorColors, saved.cidade);
  }
})();

// ══════════════════════════════════════════
// MAPA HELPERS
// ══════════════════════════════════════════
function formatCNPJ(c) { return c.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5'); }
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function getColor(pdv) {
  return mode === 'setor'
    ? (SETOR_COLORS[pdv.setor] || { color: '#64748b' }).color
    : (brickColorMap[pdv.brick] || '#64748b');
}
function isConflict(pdv) {
  if (pdv.setor.toUpperCase().includes('VAGO')) return false;
  const dom = brickSetorMap[pdv.brick];
  return dom && dom !== pdv.setor;
}
function isHidden(pdv) {
  if (hiddenSetores.has(pdv.setor)) return true;
  if (mode === 'brick' && hiddenBricks.has(pdv.brick)) return true;
  return false;
}
function makeIcon(color, conflict) {
  const ring = conflict ? 'stroke="#fbbf24" stroke-width="3"' : 'stroke="rgba(0,0,0,.4)" stroke-width="1.5"';
  const svg  = '<svg width="16" height="16" viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg">'
    + '<circle cx="8" cy="8" r="6.5" fill="' + color + '" ' + ring + '/>'
    + (conflict ? '' : '<circle cx="8" cy="8" r="2.5" fill="rgba(255,255,255,.65)"/>')
    + '</svg>';
  return L.divIcon({ html: svg, className: '', iconSize: [16,16], iconAnchor: [8,8], popupAnchor: [0,-10] });
}

// ── LEGENDA SETORES ──
function getSetorCheckState() {
  const total  = Object.keys(SETOR_COLORS).length;
  const hidden = Object.keys(SETOR_COLORS).filter(k => hiddenSetores.has(k)).length;
  if (hidden === 0) return 'all';
  if (hidden === total) return 'none';
  return 'some';
}
function buildSetorLegend() {
  const container = document.getElementById('setorLegend');
  container.innerHTML = '';
  const state  = getSetorCheckState();
  const header = document.createElement('div');
  header.className = 'setor-header';
  const hCb = document.createElement('div');
  hCb.className = 'xl-cb' + (state === 'all' ? ' checked' : state === 'some' ? ' indeterminate' : '');
  const hLbl = document.createElement('div');
  hLbl.className = 'setor-header-lbl';
  hLbl.textContent = state === 'all' ? 'Todos visíveis' : state === 'none' ? 'Todos ocultos' : 'Seleção parcial';
  header.appendChild(hCb); header.appendChild(hLbl);
  header.addEventListener('click', () => {
    if (getSetorCheckState() === 'all') Object.keys(SETOR_COLORS).forEach(k => hiddenSetores.add(k));
    else hiddenSetores.clear();
    applyVisibility(); buildSetorLegend();
  });
  container.appendChild(header);
  Object.entries(SETOR_COLORS).forEach(([k, v]) => {
    const cnt   = DATA.filter(p => p.setor === k).length;
    const isOff = hiddenSetores.has(k);
    const row   = document.createElement('div');
    row.className = 'setor-item' + (isOff ? ' off' : '');
    const cb    = document.createElement('div'); cb.className = 'xl-cb' + (isOff ? '' : ' checked');
    const dot   = document.createElement('div'); dot.className = 'setor-dot'; dot.style.background = v.color;
    const name  = document.createElement('div'); name.className = 'setor-name'; name.textContent = v.label;
    const count = document.createElement('div'); count.className = 'setor-count'; count.textContent = cnt;
    row.appendChild(cb); row.appendChild(dot); row.appendChild(name); row.appendChild(count);
    row.addEventListener('click', () => toggleSetor(k));
    container.appendChild(row);
  });
}

// ── LEGENDA BRICKS ──
function getBrickCheckState(visibleBricks) {
  const hidden = visibleBricks.filter(b => hiddenBricks.has(b)).length;
  if (hidden === 0) return 'all';
  if (hidden === visibleBricks.length) return 'none';
  return 'some';
}
function buildBrickLegend(filter) {
  filter = (filter || '').toLowerCase();
  const container = document.getElementById('brickLegend');
  container.innerHTML = '';
  const filtered = allBricks.filter(b => {
    const lbl = b.replace(/^\d+\s*-\s*[^-]+-/i, '');
    return !filter || lbl.toLowerCase().includes(filter) || b.toLowerCase().includes(filter);
  });
  if (!filtered.length) {
    container.innerHTML = '<div style="font-size:10px;color:#475569;padding:8px 0">Nenhum brick encontrado</div>';
    return;
  }
  const state  = getBrickCheckState(filtered);
  const header = document.createElement('div'); header.className = 'brick-header';
  const hCb    = document.createElement('div'); hCb.className = 'xl-cb' + (state === 'all' ? ' checked' : state === 'some' ? ' indeterminate' : '');
  const hLbl   = document.createElement('div'); hLbl.className = 'brick-header-lbl';
  hLbl.textContent = state === 'all' ? 'Todos visíveis' : state === 'none' ? 'Todos ocultos' : 'Seleção parcial';
  header.appendChild(hCb); header.appendChild(hLbl);
  header.addEventListener('click', () => {
    if (getBrickCheckState(filtered) === 'all') filtered.forEach(b => hiddenBricks.add(b));
    else filtered.forEach(b => hiddenBricks.delete(b));
    applyVisibility(); buildBrickLegend(document.getElementById('brickSearch').value);
  });
  container.appendChild(header);
  const list = document.createElement('div'); list.className = 'brick-list';
  filtered.forEach(b => {
    const cnt   = DATA.filter(p => p.brick === b).length;
    const isOff = hiddenBricks.has(b);
    const lbl   = b.replace(/^(\d+)\s*-\s*[^-]+-/i, '$1 · ') || b;
    const row   = document.createElement('div'); row.className = 'brick-item';
    const cb    = document.createElement('div'); cb.className = 'xl-cb' + (isOff ? '' : ' checked'); cb.style.cssText = 'width:10px;height:10px';
    const dot   = document.createElement('div'); dot.className = 'brick-dot' + (isOff ? ' off' : ''); dot.style.background = brickColorMap[b];
    const name  = document.createElement('div'); name.className = 'brick-name' + (isOff ? ' off' : ''); name.textContent = lbl;
    const count = document.createElement('div'); count.className = 'brick-count'; count.textContent = cnt;
    row.appendChild(cb); row.appendChild(dot); row.appendChild(name); row.appendChild(count);
    row.addEventListener('click', () => toggleBrick(b));
    list.appendChild(row);
  });
  container.appendChild(list);
}
function filterBricks(q) { buildBrickLegend(q); }

function setMode(m) {
  mode = m;
  document.getElementById('mSetor').classList.toggle('on', m === 'setor');
  document.getElementById('mBrick').classList.toggle('on', m === 'brick');
  document.getElementById('setorBlock').style.display = m === 'setor' ? '' : 'none';
  document.getElementById('brickBlock').style.display = m === 'brick' ? '' : 'none';
  if (m === 'brick') buildBrickLegend();
  Object.values(markerStore).forEach(item => item.marker.setIcon(makeIcon(getColor(item.pdv), isConflict(item.pdv))));
  applyVisibility();
}
function applyVisibility() {
  Object.values(markerStore).forEach(item => {
    if (isHidden(item.pdv)) layerGroup.removeLayer(item.marker);
    else layerGroup.addLayer(item.marker);
  });
}
function toggleSetor(s) { if (hiddenSetores.has(s)) hiddenSetores.delete(s); else hiddenSetores.add(s); applyVisibility(); buildSetorLegend(); }
function toggleBrick(b) { if (hiddenBricks.has(b)) hiddenBricks.delete(b); else hiddenBricks.add(b); applyVisibility(); buildBrickLegend(document.getElementById('brickSearch').value); }
function toggleSat() {
  satOn = !satOn;
  document.getElementById('btnSat').classList.toggle('on', satOn);
  if (satOn) { map.removeLayer(tileOSM); tileSat.addTo(map); }
  else { map.removeLayer(tileSat); tileOSM.addTo(map); }
}
function centerMap() {
  const markers = Object.values(markerStore);
  if (markers.length) map.fitBounds(L.latLngBounds(markers.map(m => m.marker.getLatLng())), { padding: [40,40] });
}
function doSearch(q) {
  const input = (q || '').trim();
  if (!input) { Object.values(markerStore).forEach(item => item.marker.setOpacity(1)); return; }
  const lq          = input.toLowerCase();
  const cnpjDigits  = input.replace(/\D/g, '');
  const matches     = Object.values(markerStore).filter(item => {
    const pdv = item.pdv;
    return pdv.nome.toLowerCase().includes(lq)
      || (cnpjDigits.length >= 7 && pdv.cnpj.includes(cnpjDigits))
      || (pdv.endereco || '').toLowerCase().includes(lq)
      || pdv.brick.toLowerCase().includes(lq)
      || pdv.setor.toLowerCase().includes(lq)
      || (pdv.bairro || '').toLowerCase().includes(lq);
  });
  Object.values(markerStore).forEach(item => item.marker.setOpacity(matches.length === 0 || matches.includes(item) ? 1 : 0.15));
  if (!matches.length) return;
  if (matches.length === 1) { map.setView(matches[0].marker.getLatLng(), 17); matches[0].marker.openPopup(); }
  else map.fitBounds(L.latLngBounds(matches.map(m => m.marker.getLatLng())), { padding: [40,40] });
}

function buildTooltipHTML(pdv) {
  const si         = SETOR_COLORS[pdv.setor] || { color: '#64748b', label: pdv.setor };
  const brickShort = pdv.brick.replace(/^\d+\s*-\s*[^-]+-/i, '');
  const brickCode  = (pdv.brick.match(/^(\d+)/) || [])[1] || '';
  let endStr;
  if (pdv.endereco)  endStr = pdv.endereco + (pdv.bairro ? ' — ' + pdv.bairro : '');
  else if (pdv.bairro) endStr = '<span style="color:#94a3b8;font-style:italic">Bairro: ' + pdv.bairro + (pdv.cidade ? ' · ' + pdv.cidade : '') + '</span>';
  else endStr = '<span style="color:#64748b;font-style:italic">Sem endereço cadastrado</span>';
  return '<div class="tip-name">🏥 ' + pdv.nome + '</div>'
    + '<div class="tip-row"><b>Endereço:</b> ' + endStr + '</div>'
    + '<div class="tip-row"><b>Brick:</b> ' + (brickCode ? brickCode + ' · ' : '') + brickShort + '</div>'
    + '<span class="tip-setor" style="background:' + si.color + '22;color:' + si.color + ';border:1px solid ' + si.color + '55">' + si.label + '</span>';
}

function buildPopupHTML(pdv) {
  const si         = SETOR_COLORS[pdv.setor] || { color: '#64748b', label: pdv.setor };
  const di         = SETOR_COLORS[brickSetorMap[pdv.brick]] || { color: '#64748b', label: '' };
  const brickShort = pdv.brick.replace(/^\d+\s*-\s*[^-]+-/i, '');
  const conf       = isConflict(pdv);
  const warn       = conf
    ? '<div class="pop-warn" style="background:#451a03;border-color:#92400e;color:#fbbf24">⚠ Brick pertence ao setor <b style="color:' + di.color + '">' + di.label + '</b></div>'
    : (pdv.setor.toUpperCase().includes('VAGO') ? '<div class="pop-warn">📋 Sem setor responsável (VAGO)</div>' : '');
  return '<div class="pop-h">🏥 ' + pdv.nome + '</div>'
    + '<div class="pop-cnpj">' + formatCNPJ(pdv.cnpj) + '</div>'
    + '<div class="pop-row"><b>Endereço:</b> ' + (pdv.endereco || '—') + '</div>'
    + '<div class="pop-row"><b>Bairro:</b> '   + (pdv.bairro   || '—') + '</div>'
    + '<div class="pop-row"><b>Cidade:</b> '   + (pdv.cidade   || '—') + (pdv.uf ? ' / ' + pdv.uf : '') + '</div>'
    + '<div class="pop-row"><b>Bandeira:</b> ' + (pdv.bandeira || '—') + '</div>'
    + '<div class="pop-row"><b>Brick:</b> '    + brickShort + '</div>'
    + '<div class="pop-row"><b>Setor:</b> <span style="color:' + si.color + ';font-weight:600">' + si.label + '</span></div>'
    + warn;
}

function addMarker(pdv, lat, lng) {
  pdv._lat = lat; pdv._lng = lng;
  const marker = L.marker([lat,lng], { icon: makeIcon(getColor(pdv), isConflict(pdv)) });
  marker.bindTooltip(buildTooltipHTML(pdv), { className: 'pdv-tip', direction: 'top', offset: [0,-8], sticky: false });
  marker.bindPopup(buildPopupHTML(pdv), { maxWidth: 290 });
  markerStore[pdv.cnpj] = { marker, pdv };
  if (!isHidden(pdv)) layerGroup.addLayer(marker);
}

function updateStats(done, total) {
  const pct = Math.round(done / total * 100);
  document.getElementById('sLoc').textContent  = done - failed;
  document.getElementById('sPend').textContent = total - done;
  document.getElementById('sFail').textContent = failed;
  document.getElementById('sFailCard').style.borderColor = failed > 0 ? '#7f1d1d' : '';
  document.getElementById('prog-fill').style.width = pct + '%';
  document.getElementById('stxt').textContent = 'Geocodificando… ' + done + '/' + total + ' — ' + (done - failed) + ' encontradas';
  if (done === total) {
    document.getElementById('status').className = 'done';
    document.getElementById('stxt').textContent = '✓ ' + (done - failed) + ' PDVs localizados · ' + failed + ' não encontrados';
    const markers = Object.values(markerStore);
    if (markers.length) map.fitBounds(L.latLngBounds(markers.map(m => m.marker.getLatLng())), { padding: [40,40] });
    persistirCoordenadas();
  }
}

// ══════════════════════════════════════════
// GEOCODIFICAÇÃO
// ══════════════════════════════════════════
const geoCache = {};

async function nominatimStructured(street, city, state) {
  const key = 'struct|' + street + '|' + city + '|' + state;
  if (geoCache[key] !== undefined) return geoCache[key];
  try {
    const params = new URLSearchParams({ format:'json', limit:'1', countrycodes:'br', street, city, state: state||'RS' });
    const r = await fetch('https://nominatim.openstreetmap.org/search?' + params, { headers: { 'Accept-Language': 'pt-BR' } });
    const d = await r.json();
    if (d && d[0] && !['suburb','neighbourhood','quarter'].includes(d[0].type)) {
      const pos = { lat: parseFloat(d[0].lat), lng: parseFloat(d[0].lon) };
      geoCache[key] = pos; return pos;
    }
    geoCache[key] = null;
  } catch { geoCache[key] = null; }
  await sleep(1100);
  return null;
}

async function nominatimFree(q, cidadeValidar) {
  if (geoCache[q] !== undefined) return geoCache[q];
  try {
    const r = await fetch(
      'https://nominatim.openstreetmap.org/search?format=json&q=' + encodeURIComponent(q) + '&limit=5&countrycodes=br&addressdetails=1',
      { headers: { 'Accept-Language': 'pt-BR' } }
    );
    const d = await r.json();
    if (d && d.length) {
      const nonSub = d.filter(x => !['suburb','neighbourhood','quarter','village','town','city_block','city'].includes(x.type));
      const cands  = nonSub.length ? nonSub : d;
      if (cidadeValidar) {
        const cidLower = cidadeValidar.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
        const inCity   = cands.find(x => {
          const addr = x.address || {};
          const city = (addr.city||addr.town||addr.municipality||addr.county||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
          return city.includes(cidLower) || cidLower.includes(city.split(' ')[0]);
        });
        if (inCity) { const pos = { lat: parseFloat(inCity.lat), lng: parseFloat(inCity.lon) }; geoCache[q] = pos; return pos; }
        geoCache[q] = null;
      } else {
        const pos = { lat: parseFloat(cands[0].lat), lng: parseFloat(cands[0].lon) };
        geoCache[q] = pos; return pos;
      }
    }
    geoCache[q] = null;
  } catch { geoCache[q] = null; }
  await sleep(1100);
  return null;
}

async function geocodePDV(pdv) {
  const cidade    = pdv.cidade || 'Porto Alegre';
  const uf        = pdv.uf     || 'RS';
  const cidadeStr = cidade + ', ' + uf + ', Brasil';

  if (pdv.endereco) {
    const m = pdv.endereco.match(/^(.+?)\s*,\s*(\d+)/);
    if (m) {
      const rua = m[1].trim(), num = m[2];
      const p1  = await nominatimStructured(rua + ' ' + num, cidade, uf); if (p1) return p1; await sleep(200);
      const p2  = await nominatimStructured(rua, cidade, uf);             if (p2) return p2; await sleep(200);
    } else {
      const p = await nominatimStructured(pdv.endereco.trim(), cidade, uf); if (p) return p; await sleep(200);
    }
  }
  if (pdv.endereco) {
    const m = pdv.endereco.match(/^(.+?)\s*,\s*(\d+)/);
    if (m) { const p = await nominatimFree(m[1].trim()+', '+m[2]+', '+cidadeStr, cidade); if (p) return p; await sleep(200); }
    const p = await nominatimFree(pdv.endereco + ', ' + cidadeStr, cidade); if (p) return p; await sleep(200);
  }
  if (pdv.brick) {
    const body     = pdv.brick.replace(/^\d+\s*-\s*/, '').replace(/^[A-ZÁÉÍÓÚ\s]+-/i, '');
    const numMatch = body.match(/^(.+?)-(\d+)\s*$/);
    if (numMatch) {
      const p1 = await nominatimStructured(numMatch[1].trim()+' '+numMatch[2], cidade, uf); if (p1) return p1; await sleep(200);
      const p2 = await nominatimFree(numMatch[1].trim()+', '+numMatch[2]+', '+cidadeStr, cidade); if (p2) return p2; await sleep(200);
    } else if (body.length > 3) {
      const p = await nominatimStructured(body.trim(), cidade, uf); if (p) return p; await sleep(200);
    }
  }
  if (pdv.bairro) { const p = await nominatimFree(pdv.bairro + ', ' + cidadeStr, cidade); if (p) return p; }
  return null;
}

async function fetchEnderecoViaCNPJ(cnpj) {
  if (!cnpj || cnpj.length !== 14) return null;
  try {
    const r = await fetch('https://brasilapi.com.br/api/cnpj/v1/' + cnpj);
    if (r.ok) {
      const d    = await r.json();
      const logr = [(d.descricao_tipo_de_logradouro||''), (d.logradouro||'')].filter(Boolean).join(' ');
      const num  = d.numero || 'S/N';
      if (logr) return { endereco: (logr+', '+num).trim(), bairro: d.bairro||'', cidade: d.municipio||'', uf: d.uf||'' };
    }
  } catch {}
  try {
    const r = await fetch('https://publica.cnpj.ws/cnpj/' + cnpj);
    if (!r.ok) return null;
    const d   = await r.json();
    const end = d.estabelecimento;
    if (!end) return null;
    const logr = [end.tipo_logradouro, end.logradouro].filter(Boolean).join(' ');
    const num  = end.numero || 'S/N';
    return { endereco: (logr+(num?', '+num:'')).trim(), bairro: end.bairro||'', cidade: end.cidade?end.cidade.nome:'', uf: end.estado?end.estado.sigla:'' };
  } catch { return null; }
}

async function runGeo(cidade) {
  const total = DATA.length;
  let done = 0;
  const jaGeo = DATA.filter(p => p._lat && p._lng);
  if (jaGeo.length === total) {
    document.getElementById('stxt').textContent = 'Carregando mapa salvo…';
    for (const pdv of DATA) { addMarker(pdv, pdv._lat, pdv._lng); done++; }
    updateStats(done, total); return;
  }
  for (const pdv of DATA) {
    if (pdv._lat && pdv._lng) { addMarker(pdv, pdv._lat, pdv._lng); done++; updateStats(done, total); continue; }
    if (!pdv.endereco && pdv.cnpj && pdv.cnpj.length === 14) {
      const info = await fetchEnderecoViaCNPJ(pdv.cnpj);
      if (info && info.endereco) { pdv.endereco = info.endereco; if (!pdv.bairro && info.bairro) pdv.bairro = info.bairro; }
      await sleep(300);
    }
    const pos = await geocodePDV(pdv);
    await sleep(300);
    if (pos) addMarker(pdv, pos.lat, pos.lng);
    else { failed++; failedList.push(pdv); }
    done++; updateStats(done, total);
  }
}

// ══════════════════════════════════════════
// FAIL PANEL & ENRICHER (no app)
// ══════════════════════════════════════════
function toggleFailPanel() {
  const panel  = document.getElementById('fail-panel');
  const isOpen = panel.style.display === 'flex';
  if (isOpen) { panel.style.display = 'none'; return; }
  if (!failedList.length) return;
  document.getElementById('failPanelSub').textContent = failedList.length + ' PDV' + (failedList.length > 1 ? 's' : '') + ' sem coordenada';
  document.getElementById('failPanelList').innerHTML  = failedList.map(p => {
    const bs  = p.brick.replace(/^(\d+)\s*-\s*[^-]+-/i, '$1 · ');
    const cf  = p.cnpj.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
    const mot = !p.endereco && !p.bairro ? 'sem endereço' : p.endereco ? 'endereço não encontrado' : 'bairro não encontrado';
    return '<div class="fail-item"><div class="fail-item-name" title="'+p.nome+'">'+p.nome+'</div><div class="fail-item-brick" title="'+p.brick+'">'+bs+'</div><div class="fail-item-cnpj">'+cf+'</div><div style="font-size:10px;color:#475569;flex-shrink:0">'+mot+'</div></div>';
  }).join('');
  panel.style.display = 'flex';
  setTimeout(() => map && map.invalidateSize(), 50);
}

function toggleEnricher() {
  const enr     = document.getElementById('enricher-screen');
  const layout  = document.querySelector('.layout');
  const statusBar = document.getElementById('status');
  const progWrap  = document.getElementById('prog-wrap');
  const btn       = document.getElementById('btnEnr');
  const isOpen    = enr.style.display === 'flex';
  enr.style.display     = isOpen ? 'none' : 'flex';
  layout.style.display  = isOpen ? 'flex' : 'none';
  statusBar.style.display = isOpen ? '' : 'none';
  progWrap.style.display  = isOpen ? '' : 'none';
  btn.classList.toggle('on', !isOpen);
  if (isOpen) setTimeout(() => map && map.invalidateSize(), 50);
}

// Enricher (inside app) — mantido para retrocompatibilidade
let enrPdvMap = {}, enrQueue = [], enrIdx = 0, enrRunning = false, enrStopped = false;

(function enrInit() {
  const drop = document.getElementById('enrDrop');
  const fi   = document.getElementById('enrFileInput');
  if (!drop || !fi) return;
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('drag'));
  drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('drag'); if (e.dataTransfer.files[0]) enrLoad(e.dataTransfer.files[0]); });
  fi.addEventListener('change', () => { if (fi.files[0]) enrLoad(fi.files[0]); });
})();

function enrLoad(file) {
  enrMsg('');
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!rows.length) { enrMsg('Planilha vazia.', true); return; }
      const k  = Object.keys(rows[0]);
      const fc = (...c) => c.map(x => k.find(y => y.trim().toLowerCase() === x.toLowerCase())).find(Boolean);
      const cS = fc('Setor','setor'), cB = fc('Brick','brick'), cP = fc('PDV','pdv'), cBa = fc('Bandeira','bandeira');
      if (!cS || !cB || !cP) { enrMsg('Colunas obrigatórias não encontradas: Setor, Brick, PDV.', true); return; }
      enrPdvMap = {};
      rows.forEach(row => {
        const raw = String(row[cP]||'').trim(), setor = String(row[cS]||'').trim(), brick = String(row[cB]||'').trim(), banda = cBa ? String(row[cBa]||'').trim() : 'N/I';
        if (!raw||!setor||!brick) return;
        let cnpj='', nome='', bairro='', cidade='', uf='';
        if (raw.includes('|')) {
          const p = raw.split('|').map(s=>s.trim()); cnpj=p[0].replace(/\D/g,''); nome=p[1]||'';
          const loc=p[2]||'', di=loc.lastIndexOf(' - ');
          if (di>=0) { bairro=loc.substring(0,di).trim(); const cu=loc.substring(di+3).trim(), si=cu.indexOf('/'); if (si>=0){cidade=cu.substring(0,si).trim();uf=cu.substring(si+1).trim();}else cidade=cu; } else bairro=loc;
        } else { cnpj=raw.replace(/\D/g,'').substring(0,14); nome=raw; }
        if (!cnpj||cnpj.length<7) return;
        const key=cnpj+'||'+brick;
        if (!enrPdvMap[key]) enrPdvMap[key]={cnpj,nome,bandeira:banda,setor,brick,bairro,cidade,uf,endereco:'',status:'pendente'};
      });
      const uniq = Object.keys(enrPdvMap);
      if (!uniq.length) { enrMsg('Nenhum PDV válido encontrado.', true); return; }
      document.getElementById('enrBadgeFile').textContent = file.name;
      document.getElementById('enrBadgeRows').textContent = rows.length + ' linhas';
      document.getElementById('enrBadgeUniq').textContent = uniq.length + ' PDVs únicos';
      document.getElementById('enrInfo').style.display = 'flex';
      document.getElementById('enrActions').style.display = 'flex';
      document.getElementById('enrStatRow').style.display = 'grid';
      document.getElementById('enrTableWrap').style.display = 'block';
      document.getElementById('enrTotal').textContent = uniq.length;
      document.getElementById('enrPend').textContent = uniq.length;
      document.getElementById('enrOk').textContent = '0';
      document.getElementById('enrErr').textContent = '0';
      enrQueue = [...uniq]; enrIdx = 0;
      enrRenderTable();
    } catch(err) { enrMsg('Erro: ' + err.message, true); }
  };
  reader.readAsArrayBuffer(file);
}

function enrRenderTable() {
  const tbody   = document.getElementById('enrTbody');
  const entries = Object.values(enrPdvMap).slice(0, 120);
  tbody.innerHTML = entries.map(p => {
    const pill = p.status==='ok' ? '<span class="enr-pill enr-ok">ok</span>' : p.status==='erro' ? '<span class="enr-pill enr-err">erro</span>' : p.status==='sem cnpj' ? '<span class="enr-pill enr-skip">s/cnpj</span>' : '<span class="enr-pill enr-pend">pendente</span>';
    const cf = p.cnpj.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/,'$1.$2.$3/$4-$5');
    const br = p.brick.replace(/^(\d+)\s*-\s*[^-]+-/i,'$1 · ');
    return '<tr><td title="'+p.cnpj+'">'+cf+'</td><td title="'+p.nome+'">'+p.nome+'</td><td title="'+p.brick+'">'+br+'</td><td title="'+p.endereco+'">'+(p.endereco||'—')+'</td><td>'+(p.cidade||'')+(p.uf?' / '+p.uf:'')+'</td><td>'+pill+'</td></tr>';
  }).join('');
  const total = Object.keys(enrPdvMap).length;
  if (total > 120) tbody.innerHTML += '<tr><td colspan="6" style="text-align:center;color:#475569;font-size:10px;padding:8px">… mais '+(total-120)+' PDVs</td></tr>';
}
function enrUpdateStats() {
  const v = Object.values(enrPdvMap);
  document.getElementById('enrOk').textContent  = v.filter(p=>p.status==='ok').length;
  document.getElementById('enrErr').textContent = v.filter(p=>p.status==='erro'||p.status==='sem cnpj').length;
  document.getElementById('enrPend').textContent= v.filter(p=>p.status==='pendente').length;
}
async function enrStart() { if (enrRunning) return; enrRunning=true; enrStopped=false; document.getElementById('enrBtnStart').style.display='none'; document.getElementById('enrBtnStop').style.display=''; document.getElementById('enrBtnResume').style.display='none'; document.getElementById('enrProg').style.display='block'; document.getElementById('enrBtnExport').disabled=false; enrMsg(''); await enrRunQueue(); }
function enrStop() { enrStopped=true; enrRunning=false; document.getElementById('enrBtnStop').style.display='none'; document.getElementById('enrBtnResume').style.display=''; enrMsg('Pausado em '+enrIdx+' de '+enrQueue.length+'. Clique em Retomar para continuar.'); }
async function enrResume() { if (enrRunning) return; enrRunning=true; enrStopped=false; document.getElementById('enrBtnResume').style.display='none'; document.getElementById('enrBtnStop').style.display=''; enrMsg(''); await enrRunQueue(); }
async function enrRunQueue() {
  const total = enrQueue.length; let consecutiveErrors = 0;
  for (; enrIdx < total; enrIdx++) {
    if (enrStopped) return;
    const key = enrQueue[enrIdx], pdv = enrPdvMap[key];
    if (!pdv || pdv.status !== 'pendente') continue;
    if (!pdv.cnpj || pdv.cnpj.length < 11) { pdv.status = 'sem cnpj'; consecutiveErrors = 0; }
    else {
      const res = await enrFetchCNPJ(pdv.cnpj);
      if (res && res.endereco) { pdv.endereco=res.endereco; if (!pdv.bairro&&res.bairro)pdv.bairro=res.bairro; if (!pdv.cidade&&res.cidade)pdv.cidade=res.cidade; if (!pdv.uf&&res.uf)pdv.uf=res.uf; if (res.telefone)pdv.telefone=res.telefone; pdv.status='ok'; consecutiveErrors=0; }
      else { pdv.status='erro'; consecutiveErrors++; if (consecutiveErrors>=5) { enrMsg('⚠️ Muitos erros — aguardando 10s…'); await sleep(10000); consecutiveErrors=0; const rk=enrQueue.slice(Math.max(0,enrIdx-4),enrIdx+1); rk.forEach(k=>{if(enrPdvMap[k]&&enrPdvMap[k].status==='erro')enrPdvMap[k].status='pendente';}); enrIdx=Math.max(0,enrIdx-4)-1; enrMsg(''); continue; } }
    }
    const pct = Math.round((enrIdx+1)/total*100);
    document.getElementById('enrProgFill').style.width = pct+'%';
    document.getElementById('enrProgPct').textContent  = pct+'%';
    document.getElementById('enrProgTxt').textContent  = 'Buscando… '+(enrIdx+1)+' / '+total;
    enrRenderTable(); enrUpdateStats(); await sleep(600);
  }
  enrRunning=false; document.getElementById('enrBtnStop').style.display='none'; document.getElementById('enrBtnStart').style.display='none'; document.getElementById('enrBtnResume').style.display='none';
  const ok = Object.values(enrPdvMap).filter(p=>p.status==='ok').length;
  enrMsg('✓ Concluído! '+ok+' de '+total+' endereços encontrados. Baixe a planilha.');
}

async function enrFetchCNPJ(cnpj) {
  for (let a = 0; a < 3; a++) {
    try {
      if (a > 0) await sleep(1500 * a);
      const r = await fetch('https://brasilapi.com.br/api/cnpj/v1/' + cnpj);
      if (r.status===429||r.status===503){await sleep(3000+a*2000);continue;}
      if (r.ok) {
        const d=await r.json(), logr=[(d.descricao_tipo_de_logradouro||''),(d.logradouro||'')].filter(Boolean).join(' '), num=d.numero||'S/N';
        const tel=d.ddd_telefone_1?d.ddd_telefone_1.replace(/\D/g,''):'';
        const telFmt=tel.length>=10?'('+tel.slice(0,2)+') '+(tel.length===11?tel.slice(2,7)+'-'+tel.slice(7):tel.slice(2,6)+'-'+tel.slice(6)):tel;
        if (logr) return {endereco:(logr+', '+num).trim(),bairro:d.bairro||'',cidade:d.municipio||'',uf:d.uf||'',telefone:telFmt};
      } break;
    } catch {}
  }
  for (let a = 0; a < 3; a++) {
    try {
      if (a > 0) await sleep(1500 * a);
      const r = await fetch('https://publica.cnpj.ws/cnpj/' + cnpj);
      if (r.status===429||r.status===503){await sleep(3000+a*2000);continue;}
      if (!r.ok) break;
      const d=await r.json(), end=d.estabelecimento;
      if (!end) break;
      const logr=[end.tipo_logradouro,end.logradouro].filter(Boolean).join(' '), num=end.numero||'S/N';
      const tel=end.telefone1?end.telefone1.replace(/\D/g,''):'';
      const telFmt=tel.length>=10?'('+tel.slice(0,2)+') '+(tel.length===11?tel.slice(2,7)+'-'+tel.slice(7):tel.slice(2,6)+'-'+tel.slice(6)):tel;
      return {endereco:(logr+(num?', '+num:'')).trim(),bairro:end.bairro||'',cidade:end.cidade?end.cidade.nome:'',uf:end.estado?end.estado.sigla:'',telefone:telFmt};
    } catch {}
  }
  return null;
}

function enrExport() {
  const pdvs = Object.values(enrPdvMap), wb = XLSX.utils.book_new();
  const wsData = [['Setor','Brick','Nome do PDV','CNPJ','Cidade','Telefone']];
  pdvs.forEach(p => {
    const cf = p.cnpj.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/,'$1.$2.$3/$4-$5');
    wsData.push([p.setor, p.brick, p.nome, cf, (p.cidade||'')+(p.uf?' / '+p.uf:''), p.telefone||'']);
  });
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [32,30,40,20,22,16].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws, 'PDVs Enriquecidos');
  const ws2=[['CNPJ','Nome','Setor','Brick','Status']];
  pdvs.filter(p=>p.status!=='ok').forEach(p=>ws2.push([p.cnpj,p.nome,p.setor,p.brick,p.status]));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ws2), 'Não encontrados');
  XLSX.writeFile(wb, 'PDVs_enriquecidos.xlsx');
}

function enrReset() {
  enrStopped=true; enrRunning=false; enrPdvMap={}; enrQueue=[]; enrIdx=0;
  ['enrInfo','enrActions','enrStatRow'].forEach(id=>document.getElementById(id).style.display='none');
  document.getElementById('enrProg').style.display='none'; document.getElementById('enrTableWrap').style.display='none';
  document.getElementById('enrBtnStart').style.display=''; document.getElementById('enrBtnStop').style.display='none';
  document.getElementById('enrBtnResume').style.display='none'; document.getElementById('enrBtnExport').disabled=true;
  document.getElementById('enrTbody').innerHTML=''; document.getElementById('enrFileInput').value=''; enrMsg('');
}
function enrMsg(txt, isErr) { const el=document.getElementById('enrMsg'); el.textContent=txt; el.className='enr-msg'+(isErr?' err':''); }
