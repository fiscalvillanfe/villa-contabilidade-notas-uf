// ======= NF-e no navegador: filtro por UF, agrupamento e exportações =======

function sanitizeFolder(name) {
  return (name || 'DESCONHECIDO')
    .replace(/[^\w\s\-\u00C0-\u024F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim() || 'DESCONHECIDO';
}
function onlyDigits(s) { return (s||'').replace(/\D+/g,''); }
function canonicalFolder(name, doc) {
  const docDigits = onlyDigits(doc);
  const baseName = sanitizeFolder(name||'DESCONHECIDO');
  return docDigits ? `${baseName} - ${docDigits}` : baseName;
}
function pickBetterName(a,b){ if(!a) return b; if(!b) return a; return (b.length>a.length)?b:a; }

// nome seguro para arquivo zip (Windows não aceita : * ? " < > | / \)
function safeZipName(name){
  return (name||'empresa').replace(/[\\/:*?"<>|]/g,' ').replace(/\s+/g,' ').trim() || 'empresa';
}

// Busca por caminho usando localName (independe de namespace)
function findFirstLocalPath(root, pathList) {
  for (const path of pathList) {
    let cur = root; let ok = true;
    for (const seg of path) {
      let found = null;
      for (const ch of cur.children) { if ((ch.localName||ch.nodeName)===seg) { found=ch; break; } }
      if (!found) { ok=false; break; } cur = found;
    }
    if (ok && cur && cur.textContent) return cur.textContent.trim();
  }
  return null;
}
function extractUF(doc, criteria) {
  let node = doc.getElementsByTagName('infNFe')[0]
         || (doc.getElementsByTagName('NFe')[0]?.getElementsByTagName('infNFe')[0])
         || doc.documentElement;

  if (criteria==='emit') { const v=findFirstLocalPath(node,[["emit","enderEmit","UF"]]); if(v) return v; }
  else { const v=findFirstLocalPath(node,[["dest","enderDest","UF"]]); if(v) return v; }

  const map={"11":"RO","12":"AC","13":"AM","14":"RR","15":"PA","16":"AP","17":"TO","21":"MA","22":"PI","23":"CE","24":"RN","25":"PB","26":"PE","27":"AL","28":"SE","29":"BA","31":"MG","32":"ES","33":"RJ","35":"SP","41":"PR","42":"SC","43":"RS","50":"MS","51":"MT","52":"GO","53":"DF"};
  const cuf=findFirstLocalPath(node,[["ide","cUF"]]); if(cuf&&map[cuf.padStart(2,'0')]) return map[cuf.padStart(2,'0')];

  if (criteria==='emit') { const v=findFirstLocalPath(node,[["emit","UF"]]); if(v) return v; }
  else { const v=findFirstLocalPath(node,[["dest","UF"]]); if(v) return v; }
  return null;
}
function extractParty(doc, which) {
  let node = doc.getElementsByTagName('infNFe')[0]
         || (doc.getElementsByTagName('NFe')[0]?.getElementsByTagName('infNFe')[0])
         || doc.documentElement;
  const base = which==='dest' ? ['dest'] : ['emit'];
  const name = findFirstLocalPath(node, [base.concat(['xNome'])]);
  let docnum = findFirstLocalPath(node, [base.concat(['CNPJ'])]);
  if (!docnum) docnum = findFirstLocalPath(node, [base.concat(['CPF'])]);
  return { name: name || 'DESCONHECIDO', doc: docnum || '' };
}

function buildCSV(rows, sep=';') {
  return rows.map(r => r.map(v => {
    const s = String(v ?? '');
    if (s.includes('"') || s.includes(sep) || s.includes('\n')) return '"' + s.replace(/"/g,'""') + '"';
    return s;
  }).join(sep)).join('\n');
}

// === elementos da página
const byId = id => document.getElementById(id);
const fileInput = byId('fileInput');
const processBtn = byId('processBtn');
const criteriaSel = byId('criteria');
const companyKeySel = byId('companyKey');
const targetUFIn = byId('targetUF');
const reportArea = byId('reportArea');
const statsGrid = byId('statsGrid');
const companiesBox = byId('companiesBox');
const tableBox = byId('tableBox');
const dlAllBtn = byId('dlAll');
const dlNonBtn = byId('dlNon');

function deriveBaseName(fileList){
  if (!fileList || fileList.length===0) return 'resultado';
  let name = fileList[0].name || 'resultado';
  if (name.includes('.')) name = name.substring(0, name.lastIndexOf('.'));
  return sanitizeFolder(name) || 'resultado';
}

async function readAllXMLs(fileList) {
  const xmlItems = [];
  for (const file of fileList) {
    const low = file.name.toLowerCase();
    if (low.endsWith('.xml')) {
      xmlItems.push({name:file.name, text: await file.text()});
    } else if (low.endsWith('.zip')) {
      const zip = await JSZip.loadAsync(file);
      for (const entry of Object.values(zip.files)) {
        if (!entry.dir && entry.name.toLowerCase().endsWith('.xml')) {
          const txt = await entry.async('string');
          const base = entry.name.split('/').pop();
          xmlItems.push({name: base, text: txt});
        }
      }
    }
  }
  return xmlItems;
}

async function processFiles(fileList) {
  const criteria = criteriaSel.value;        // 'emit' ou 'dest'
  const companyKey = companyKeySel.value;    // 'dest' ou 'emit'
  const targetUF = (targetUFIn.value.trim().toUpperCase() || 'MG');
  const baseName = deriveBaseName(fileList);

  const xmls = await readAllXMLs(fileList);
  if (xmls.length === 0) throw new Error('Nenhum XML encontrado.');

  const stats = { total:0, mg:0, nao_mg:0, unknown:0, errors:0 };
  const rows = [['arquivo','status','uf','empresa','cnpj_cpf','grupo']];
  const nonRows = [];
  const groups = {};     // pasta -> itens
  const placed = [];     // itens temporários
  const docToBestName = {}; // "digits" -> melhor display name

  for (const item of xmls) {
    stats.total++;
    try {
      const docXml = new DOMParser().parseFromString(item.text, 'text/xml');
      const uf = extractUF(docXml, criteria);
      const dest = extractParty(docXml, 'dest');
      const emit = extractParty(docXml, 'emit');
      const comp = (companyKey === 'dest') ? dest : emit;

      const digits = onlyDigits(comp.doc);
      if (digits) docToBestName[digits] = pickBetterName(docToBestName[digits], sanitizeFolder(comp.name));

      let bucket = 'UF_desconhecida';
      let status = 'unknown';
      let ufShow = '';
      if (uf) {
        ufShow = uf.toUpperCase();
        if (ufShow === targetUF) { bucket = targetUF; status = 'ok'; stats.mg++; }
        else { bucket = 'nao_' + targetUF; status = 'ok'; stats.nao_mg++; nonRows.push([item.name, ufShow, '']); }
      } else { stats.unknown++; }

      rows.push([item.name, status, ufShow, comp.name, comp.doc, bucket]);
      placed.push({comp, bucket, filename:item.name, text:item.text});
    } catch(e) {
      stats.errors++;
      rows.push([item.name, 'parse_error', '', '', '', 'ERROS']);
      placed.push({comp:{name:'ERROS',doc:''}, bucket:'ERROS', filename:item.name, text:item.text});
    }
  }

  // Pasta canônica por CNPJ/CPF
  const finalItems = [];
  for (const rec of placed) {
    const digits = onlyDigits(rec.comp.doc);
    let displayName = sanitizeFolder(rec.comp.name);
    if (digits && docToBestName[digits]) displayName = docToBestName[digits];
    const folder = canonicalFolder(displayName, rec.comp.doc);
    const fullPath = `${folder}/${rec.bucket}`;
    finalItems.push({folder, fullPath, filename:rec.filename, text:rec.text, ufBucket:rec.bucket});
    if (!groups[folder]) groups[folder] = [];
    groups[folder].push({fullPath, filename:rec.filename, text:rec.text});
  }

  // Preenche empresa no CSV "fora do estado" com a pasta canônica
  for (const row of nonRows) {
    const found = finalItems.find(p => p.filename === row[0] && p.ufBucket.startsWith('nao_'));
    if (found) row[2] = found.folder;
  }

  async function buildAllZip() {
    const zip = new JSZip();
    for (const rec of finalItems) zip.file(`${rec.fullPath}/${rec.filename}`, rec.text);
    zip.file('resultado_filtragem.csv', buildCSV(rows));
    return await zip.generateAsync({type:'blob'});
  }

  // Gera sub-ZIPs por empresa, com XMLs soltos dentro de cada sub-ZIP
  async function buildNonZip() {
    const nonByCompany = {};
    for (const rec of finalItems) {
      if (rec.ufBucket && rec.ufBucket.startsWith('nao_')) {
        (nonByCompany[rec.folder] ??= []).push(rec);
      }
    }

    const zip = new JSZip();
    for (const [company, items] of Object.entries(nonByCompany)) {
      const sub = new JSZip();
      for (const rec of items) {
        sub.file(rec.filename, rec.text);
      }
      const subContent = await sub.generateAsync({ type: 'uint8array' });
      zip.file(`${safeZipName(company)}.zip`, subContent);
    }
    zip.file('fora_do_estado.csv', buildCSV([['arquivo','uf','empresa'], ...nonRows]));
    return await zip.generateAsync({ type:'blob' });
  }

  async function buildCompanyZip(compName) {
    const items = groups[compName] || [];
    const zip = new JSZip();
    for (const rec of items) {
      const rel = rec.fullPath.substring(compName.length + 1);
      zip.file(`${compName}/${rel}/${rec.filename}`, rec.text);
    }
    return await zip.generateAsync({type:'blob'});
  }

  return {baseName, targetUF, stats, groups, nonRows, buildAllZip, buildNonZip, buildCompanyZip};
}

function renderReport(state) {
  const statsGrid = document.getElementById('statsGrid');
  const companiesBox = document.getElementById('companiesBox');
  const tableBox = document.getElementById('tableBox');
  const dlAllBtn = document.getElementById('dlAll');
  const dlNonBtn = document.getElementById('dlNon');
  const reportArea = document.getElementById('reportArea');

  statsGrid.innerHTML = `
    <div class="stat">Total<br><strong>${state.stats.total}</strong></div>
    <div class="stat">${state.targetUF}<br><strong>${state.stats.mg}</strong></div>
    <div class="stat">Não ${state.targetUF}<br><strong>${state.stats.nao_mg}</strong></div>
    <div class="stat">UF desconhecida<br><strong>${state.stats.unknown}</strong></div>
    <div class="stat">Erros<br><strong>${state.stats.errors}</strong></div>
  `;

  const names = Object.keys(state.groups).sort((a,b)=>a.localeCompare(b));
  if (names.length === 0) {
    companiesBox.innerHTML = '<p class="msg">Nenhuma empresa detectada.</p>';
  } else {
    companiesBox.innerHTML = names.map(n => `
      <div class="company">
        <div><code>${n}</code> <span class="badge">${state.groups[n].length} xml</span></div>
        <button class="btn primary" data-company="${encodeURIComponent(n)}">Baixar ZIP</button>
      </div>
    `).join('');
    companiesBox.querySelectorAll('button[data-company]').forEach(btn => {
      btn.addEventListener('click', async () => {
        const comp = decodeURIComponent(btn.getAttribute('data-company'));
        btn.disabled = true; btn.textContent = 'Gerando...';
        try { saveAs(await state.buildCompanyZip(comp), `${comp}.zip`); }
        finally { btn.disabled = false; btn.textContent = 'Baixar ZIP'; }
      });
    });
  }

  if (!state.nonRows || state.nonRows.length === 0) {
    tableBox.innerHTML = '<p class="msg">Nenhuma nota fora do estado encontrada.</p>';
  } else {
    const rows = state.nonRows.map(r => `<tr><td>${r[0]}</td><td>${r[1]}</td><td>${r[2]}</td></tr>`).join('');
    tableBox.innerHTML = `
      <div style="overflow:auto;">
        <table>
          <thead><tr><th>Arquivo</th><th>UF</th><th>Empresa</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
  }

  dlAllBtn.onclick = async () => {
    dlAllBtn.disabled = true; dlAllBtn.textContent = 'Gerando...';
    try { saveAs(await state.buildAllZip(), `${state.baseName}-resultado.zip`); }
    finally { dlAllBtn.disabled = false; dlAllBtn.textContent = 'Baixar tudo (ZIP)'; }
  };
  dlNonBtn.onclick = async () => {
    dlNonBtn.disabled = true; dlNonBtn.textContent = 'Gerando...';
    try { saveAs(await state.buildNonZip(), `${state.baseName}-notas_fora_do_estado.zip`); }
    finally { dlNonBtn.disabled = false; dlNonBtn.textContent = 'Baixar só fora do estado (ZIP)'; }
  };

  reportArea.classList.remove('hidden');
}

document.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('fileInput');
  const processBtn = document.getElementById('processBtn');
  const reportArea = document.getElementById('reportArea');

  processBtn.addEventListener('click', async () => {
    const files = fileInput.files;
    if (!files || files.length === 0) { alert('Selecione pelo menos um ZIP ou XML.'); return; }
    processBtn.disabled = true; processBtn.textContent = 'Processando...';
    reportArea.classList.add('hidden');
    try { renderReport(await processFiles(files)); }
    catch(e) { console.error(e); alert(e.message || 'Erro ao processar.'); }
    finally { processBtn.disabled = false; processBtn.textContent = 'Processar'; }
  });
});
