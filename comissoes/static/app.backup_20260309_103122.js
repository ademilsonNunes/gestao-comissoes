/* ============================================================
   app.js  Gestão de Comissões | UI rebuild
   ============================================================ */

const STATE = {
  reps: [],
  repByCodvend: {},
  regras: [],
  comissoes: [],
  lancamentos: [],
  filteredLanc: [],
  selectedRepIds: new Set(),
  sortLanc: { key: null, dir: 1 },
  sortReps: { key: null, dir: 1 },
  currentComissaoId: null,
  currentComissaoStatus: '',
};
const SELECTED_REP_IDS_KEY = 'gc:selectedRepIds';

/*  helpers  */
const $ = id => document.getElementById(id);
const exists = id => !!$(id);
const money = v => Number(v || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const num = v => Number(v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 4 });
const pct2 = v => `${Number(v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}%`;
const numInput = v => Number(v || 0).toFixed(2);
const toFloatOrNull = v => {
  if (v === null || v === undefined) return null;
  const s = String(v).replace(',', '.').trim();
  if (!s) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
};

function hydrateSelectedRepIds() {
  try {
    const raw = localStorage.getItem(SELECTED_REP_IDS_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return;
    STATE.selectedRepIds = new Set(parsed.map(x => Number(x)).filter(Number.isFinite));
  } catch {}
}

function persistSelectedRepIds() {
  try {
    localStorage.setItem(SELECTED_REP_IDS_KEY, JSON.stringify(Array.from(STATE.selectedRepIds)));
  } catch {}
}

function syncEnvioIdsField() {
  if (!exists('email-ids')) return;
  const ids = Array.from(STATE.selectedRepIds).sort((a, b) => a - b);
  $('email-ids').value = ids.join(', ');
}

function toast(msg, type = 'default') {
  const t = $('toast');
  if (!t) return;
  t.textContent = msg;
  t.className = type === 'error' ? 'toast-error' : type === 'success' ? 'toast-success' : '';
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2800);
}

async function api(url, options = {}) {
  const r = await fetch(url, options);
  const txt = await r.text();
  let body = {};
  try { body = JSON.parse(txt); } catch { body = { raw: txt }; }
  if (!r.ok) throw new Error(body.error || body.motivo || `Erro ${r.status}`);
  return body;
}

function formData(formEl) {
  const d = {};
  new FormData(formEl).forEach((v, k) => { d[k] = v; });
  return d;
}

function badgeHtml(status) {
  const s = String(status || 'rascunho').toLowerCase();
  const labels = { rascunho: 'Rascunho', aprovado: 'Aprovado', enviado: 'Enviado', erro: 'Erro', ok: 'OK' };
  return `<span class="badge badge-${s}">${labels[s] || status}</span>`;
}

function alertHtml(type, title, text) {
  const icons = { info: 'i', success: 'OK', warn: '!', error: 'x' };
  return `<div class="alert alert-${type}">
    <span class="alert-icon">${icons[type] || '*'}</span>
    <div class="alert-body">
      ${title ? `<div class="alert-title">${title}</div>` : ''}
      <div>${text}</div>
    </div>
  </div>`;
}

function emptyState(icon, title, text = '') {
  return `<div class="empty-state">
    <span class="empty-icon">${icon}</span>
    <div class="empty-title">${title}</div>
    ${text ? `<div class="empty-text">${text}</div>` : ''}
  </div>`;
}

function confirmDialog({
  title = 'Confirmacao',
  message = 'Deseja continuar?',
  confirmText = 'Confirmar',
  cancelText = 'Cancelar',
  danger = false,
} = {}) {
  return new Promise(resolve => {
    const backdrop = document.createElement('div');
    backdrop.className = 'modal-backdrop open';
    backdrop.innerHTML = `
      <div class="modal" role="dialog" aria-modal="true" aria-label="${title}">
        <div class="modal-header">
          <div class="modal-title">${title}</div>
          <button class="modal-close" type="button" aria-label="Fechar">×</button>
        </div>
        <div class="modal-body">
          <div class="text-sm">${message}</div>
        </div>
        <div class="modal-footer">
          <button class="btn btn-ghost" type="button" data-cancel>${cancelText}</button>
          <button class="btn ${danger ? 'btn-danger' : 'btn-primary'}" type="button" data-confirm>${confirmText}</button>
        </div>
      </div>`;

    const close = result => {
      backdrop.classList.remove('open');
      setTimeout(() => backdrop.remove(), 120);
      resolve(result);
    };
    backdrop.addEventListener('click', e => {
      if (e.target === backdrop) close(false);
    });
    backdrop.querySelector('[data-cancel]')?.addEventListener('click', () => close(false));
    backdrop.querySelector('.modal-close')?.addEventListener('click', () => close(false));
    backdrop.querySelector('[data-confirm]')?.addEventListener('click', () => close(true));
    const esc = e => {
      if (e.key === 'Escape') {
        document.removeEventListener('keydown', esc);
        close(false);
      }
    };
    document.addEventListener('keydown', esc);
    document.body.appendChild(backdrop);
  });
}

/* 
   DASHBOARD
 */
async function loadDashboard() {
  if (!exists('kpi-grid')) return;
  try {
    const d = await api('/api/auditoria/cadastro');
    // KPIs
    $('kpi-grid').innerHTML = `
      <div class="kpi-card kpi-blue">
        <div class="kpi-label">Representantes</div>
        <div class="kpi-value">${d.representantes_total ?? 0}</div>
      </div>
      <div class="kpi-card kpi-green">
        <div class="kpi-label">Ativos</div>
        <div class="kpi-value">${d.representantes_ativos ?? 0}</div>
      </div>
      <div class="kpi-card ${(d.representantes_sem_email ?? 0) > 0 ? 'kpi-amber' : 'kpi-green'}">
        <div class="kpi-label">Sem e-mail</div>
        <div class="kpi-value">${d.representantes_sem_email ?? 0}</div>
      </div>
      <div class="kpi-card kpi-blue">
        <div class="kpi-label">Regras ativas</div>
        <div class="kpi-value">${d.regras_ativas ?? 0}</div>
      </div>
      <div class="kpi-card ${(d.codvend_sem_cadastro ?? 0) > 0 ? 'kpi-red' : 'kpi-green'}">
        <div class="kpi-label">CODVEND sem cadastro</div>
        <div class="kpi-value">${d.codvend_sem_cadastro ?? 0}</div>
      </div>
      <div class="kpi-card ${(d.representantes_sem_regra ?? 0) > 0 ? 'kpi-amber' : 'kpi-green'}">
        <div class="kpi-label">Ativos sem regra</div>
        <div class="kpi-value">${d.representantes_sem_regra ?? 0}</div>
      </div>
    `;
    // Audit alerts
    const alerts = [];
    if ((d.codvend_sem_cadastro ?? 0) > 0)
      alerts.push(alertHtml('error', `${d.codvend_sem_cadastro} CODVEND sem cadastro`, 'Vá para Representantes e cadastre os vendedores faltantes para que as comissões sejam calculadas corretamente.'));
    if ((d.representantes_sem_email ?? 0) > 0)
      alerts.push(alertHtml('warn', `${d.representantes_sem_email} representante(s) sem e-mail`, 'Adicione o e-mail para poder enviar os relatórios.'));
    if ((d.representantes_sem_regra ?? 0) > 0)
      alerts.push(alertHtml('warn', `${d.representantes_sem_regra} representante(s) sem regra`, 'Configure as regras de comissão em Regras de Comissão.'));
    if (!alerts.length)
      alerts.push(alertHtml('success', 'Cadastro OK', 'Todos os representantes estão cadastrados com e-mail e regras configuradas.'));
    $('audit-alerts').innerHTML = alerts.join('');

    // Workflow steps
    updateWorkflowStep('ws-import', true, (d.representantes_total ?? 0) > 0);
    updateWorkflowStep('ws-reps', (d.representantes_total ?? 0) > 0, (d.representantes_ativos ?? 0) > 0 && (d.codvend_sem_cadastro ?? 0) === 0);
    updateWorkflowStep('ws-regras', (d.regras_ativas ?? 0) > 0, (d.representantes_sem_regra ?? 0) === 0);
    updateWorkflowStep('ws-calc', (d.regras_ativas ?? 0) > 0, false);
    updateWorkflowStep('ws-envio', false, false);
  } catch (e) {
    if (exists('audit-alerts')) $('audit-alerts').innerHTML = alertHtml('error', 'Erro ao carregar dashboard', e.message);
  }
}

function updateWorkflowStep(id, active, done) {
  const el = $(id);
  if (!el) return;
  el.classList.remove('active', 'done');
  if (done) el.classList.add('done');
  else if (active) el.classList.add('active');
}

/* 
   IMPORTAO
 */
function initImportacao() {
  if (!exists('upload-zone')) return;
  const zone = $('upload-zone');
  const fileInput = $('upload-file');
  const btn = $('btn-upload');
  const fnLabel = $('upload-filename');

  zone.addEventListener('click', () => fileInput.click());
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) { fileInput.files = e.dataTransfer.files; setFile(file); }
  });
  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) setFile(fileInput.files[0]);
  });

  function setFile(file) {
    fnLabel.textContent = ` ${file.name}`;
    fnLabel.classList.remove('hidden');
    btn.disabled = false;
  }

  btn.addEventListener('click', async () => {
    const file = fileInput.files[0];
    if (!file) return toast('Selecione um arquivo primeiro.', 'error');
    const fd = new FormData();
    fd.append('arquivo', file);
    const mes = $('imp-mes').value;
    const ano = $('imp-ano').value;
    if (mes) fd.append('mes', mes);
    if (ano) fd.append('ano', ano);

    btn.disabled = true;
    btn.innerHTML = '<span class="spinner"></span> Importando...';
    $('imp-progress').classList.remove('hidden');
    animateProgress('imp-progress-fill', 'imp-progress-text');

    try {
      const r = await fetch('/api/importacao/upload', { method: 'POST', body: fd });
      const body = await r.json();
      if (!r.ok) throw new Error(body.error || 'Falha na importação');
      $('imp-result-area').innerHTML = alertHtml('success', 'Importação concluída!',
        `Importados: <strong>${body.importados ?? 0}</strong> lançamentos · Período: ${body.mes ?? '?'}/${body.ano ?? '?'}`);
      toast('Importação concluída!', 'success');
    } catch (e) {
      $('imp-result-area').innerHTML = alertHtml('error', 'Erro na importação', e.message);
      toast(e.message, 'error');
    } finally {
      btn.disabled = false;
      btn.innerHTML = '<span></span> Importar arquivo';
      $('imp-progress').classList.add('hidden');
    }
  });

  if (exists('btn-ver-pendencias')) {
    $('btn-ver-pendencias').addEventListener('click', async () => {
      try {
        const r = await api('/api/importacao/1/pendencias');
        const pendentes = r.pendentes || [];
        if (!pendentes.length) {
          $('pendencias-body').innerHTML = alertHtml('success', 'Nenhuma pendência', 'Todos os CODVEND possuem representante cadastrado.');
        } else {
          const tags = pendentes.map(c => `<span class="badge badge-inativo" style="margin:2px">${c}</span>`).join('');
          $('pendencias-body').innerHTML = alertHtml('error', `${pendentes.length} CODVEND sem cadastro`, `${tags}<div class="mt-12"><a href="/representantes" class="btn btn-primary btn-sm">Ir para Representantes</a></div>`);
        }
      } catch (e) { toast(e.message, 'error'); }
    });
  }
}

function animateProgress(fillId, textId) {
  let p = 0;
  const fill = $(fillId);
  const text = $(textId);
  const iv = setInterval(() => {
    p = Math.min(p + Math.random() * 15, 90);
    if (fill) fill.style.width = p + '%';
    if (text) text.textContent = `Processando... ${Math.round(p)}%`;
  }, 300);
  setTimeout(() => {
    clearInterval(iv);
    if (fill) fill.style.width = '100%';
    if (text) text.textContent = 'Concluído.';
  }, 3000);
}

/* 
   REPRESENTANTES
 */
async function loadReps() {
  if (!exists('reps-list') && !exists('comissoes-list')) return;
  STATE.reps = await api('/api/representantes');
  STATE.repByCodvend = {};
  for (const r of STATE.reps) STATE.repByCodvend[String(r.codvend || '')] = r;
  if (exists('reps-list')) renderReps();
  if (exists('reps-count-label')) $('reps-count-label').textContent = `${STATE.reps.length} representante(s) cadastrado(s)`;
}

function renderReps() {
  const container = $('reps-list');
  if (!container) return;
  const search = ($('rep-search')?.value || '').toLowerCase();
  const statusFilter = $('rep-filter-status')?.value ?? '';
  let list = STATE.reps.filter(r => {
    const matchSearch = !search || `${r.codvend} ${r.nome} ${r.email}`.toLowerCase().includes(search);
    const matchStatus = statusFilter === '' || String(r.ativo) === statusFilter;
    return matchSearch && matchStatus;
  });

  if ($('reps-count-label')) $('reps-count-label').textContent = `${list.length} de ${STATE.reps.length} representante(s)`;

  if (!list.length) {
    container.innerHTML = emptyState('', 'Nenhum representante encontrado');
    return;
  }

  const rows = list.map(r => `
    <tr data-rep-row="${r.id}">
      <td class="mono">${r.id}</td>
      <td><span class="badge badge-${r.ativo ? 'ativo' : 'inativo'}">${r.ativo ? 'Ativo' : 'Inativo'}</span></td>
      <td class="mono fw-bold">${r.codvend || ''}</td>
      <td><input class="inline-input w120 rep-nome" value="${esc(r.nome || '')}"></td>
      <td><input class="inline-input w180 rep-email" type="email" value="${esc(r.email || '')}"></td>
      <td><input class="inline-input w180 rep-corpo" value="${esc(r.corpo_email || '')}"></td>
      <td>
        <div class="table-actions">
          <button class="btn btn-ghost btn-sm" data-rep-save="${r.id}"> Salvar</button>
          <button class="btn btn-danger btn-sm" data-rep-del="${r.id}">${r.ativo ? 'Inativar' : 'Já inativo'}</button>
        </div>
      </td>
    </tr>`).join('');

  container.innerHTML = `<div class="table-wrapper"><table>
    <thead><tr>
      <th>ID</th><th>Status</th><th>Acoes</th><th>Nome</th><th>E-mail</th><th>Corpo e-mail</th><th>Ações</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table></div>`;

  container.querySelectorAll('button[data-rep-save]').forEach(b => {
    b.addEventListener('click', async () => {
      const id = Number(b.dataset.repSave);
      const row = container.querySelector(`tr[data-rep-row="${id}"]`);
      b.disabled = true; b.innerHTML = '<span class="spinner"></span>';
      try {
        await api(`/api/representantes/${id}`, {
          method: 'PUT', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            nome: row.querySelector('.rep-nome').value,
            email: row.querySelector('.rep-email').value,
            corpo_email: row.querySelector('.rep-corpo').value,
          }),
        });
        toast('Representante atualizado.', 'success');
        await loadReps();
      } catch (e) { toast(e.message, 'error'); }
      finally { b.disabled = false; b.innerHTML = ' Salvar'; }
    });
  });

  container.querySelectorAll('button[data-rep-del]').forEach(b => {
    b.addEventListener('click', async () => {
      const ok = await confirmDialog({ title: 'Inativar Representante', message: 'Deseja inativar este representante?', confirmText: 'Inativar', danger: true });
      if (!ok) return;
      try {
        await api(`/api/representantes/${Number(b.dataset.repDel)}`, { method: 'DELETE' });
        toast('Representante inativado.', 'success');
        await loadReps(); await loadDashboard();
      } catch (e) { toast(e.message, 'error'); }
    });
  });
}

function initRepresentantes() {
  if (!exists('rep-form')) return;
  $('rep-form').addEventListener('submit', async e => {
    e.preventDefault();
    const btn = e.target.querySelector('[type=submit]');
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Salvando...';
    try {
      await api('/api/representantes', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(formData(e.target)),
      });
      e.target.reset();
      await loadReps(); await loadDashboard();
      toast('Representante salvo!', 'success');
      if (exists('rep-result-area')) $('rep-result-area').innerHTML = alertHtml('success', '', 'Representante cadastrado com sucesso.');
    } catch (err) { toast(err.message, 'error'); }
    finally { btn.disabled = false; btn.innerHTML = ' Salvar representante'; }
  });

  if (exists('btn-rep-clear')) $('btn-rep-clear').addEventListener('click', () => $('rep-form').reset());

  if (exists('btn-imp-reps')) {
    $('btn-imp-reps').addEventListener('click', async () => {
      const btn = $('btn-imp-reps');
      btn.disabled = true; btn.innerHTML = '<span class="spinner spinner-dark"></span> Importando...';
      try {
        const r = await api('/api/representantes/importar', { method: 'POST' });
        await loadReps(); await loadDashboard();
        if (exists('rep-result-area')) $('rep-result-area').innerHTML = alertHtml('success', 'Importação concluída', `${r.importados} representante(s) importado(s) da base.`);
        toast(`${r.importados} representante(s) importado(s).`, 'success');
      } catch (e) { toast(e.message, 'error'); }
      finally { btn.disabled = false; btn.innerHTML = ' Importar BASE_ENVIO'; }
    });
  }

  if (exists('rep-search')) {
    $('rep-search').addEventListener('input', renderReps);
    $('rep-filter-status').addEventListener('change', renderReps);
  }
}

/* 
   REGRAS
 */
async function loadRegras() {
  if (!exists('regras-list')) return;
  STATE.regras = await api('/api/regras');
  renderRegras();
}

function renderRegras() {
  const container = $('regras-list');
  if (!container) return;
  const search = ($('regra-search')?.value || '').toLowerCase();
  const list = search
    ? STATE.regras.filter(r => `${r.codvend} ${r.rede || ''} ${r.uf || ''} ${r.descricao || ''}`.toLowerCase().includes(search))
    : STATE.regras;

  if ($('regras-count-label')) $('regras-count-label').textContent = `${list.length} regra(s)`;

  if (!list.length) {
    container.innerHTML = emptyState('️', 'Nenhuma regra cadastrada', 'Use o formulário acima para adicionar regras de comissão.');
    return;
  }

  const rows = list.map(r => `
    <tr>
      <td class="mono">${r.id}</td>
      <td class="mono fw-bold">${r.codvend}</td>
      <td class="fw-bold" style="color:var(--blue-dark)">${num(r.percentual)}%</td>
      <td>${r.prioridade}</td>
      <td>${r.codcliente || '<span class="text-muted"></span>'}</td>
      <td>${r.rede || '<span class="text-muted"></span>'}</td>
      <td>${r.uf || '<span class="text-muted"></span>'}</td>
      <td>${r.codprod || '<span class="text-muted"></span>'}</td>
      <td><span class="badge badge-${r.ativo ? 'ativo' : 'inativo'}">${r.ativo ? 'Ativa' : 'Inativa'}</span></td>
      <td class="text-muted text-sm">${r.descricao || ''}</td>
      <td>
        <div class="table-actions">
          <button class="btn btn-ghost btn-sm" data-regra-edit="${r.id}">️ Editar</button>
          <button class="btn btn-danger btn-sm" data-regra-del="${r.id}">️</button>
        </div>
      </td>
    </tr>`).join('');

  container.innerHTML = `<div class="table-wrapper"><table>
    <thead><tr>
      <th>ID</th><th>CODVEND</th><th>%</th><th>Prioridade</th>
      <th>Cliente</th><th>Rede</th><th>UF</th><th>Produto</th>
      <th>Status</th><th>Descrição</th><th>Ações</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table></div>`;

  container.querySelectorAll('button[data-regra-edit]').forEach(b => {
    b.addEventListener('click', () => {
      const id = Number(b.dataset.regraEdit);
      const r = STATE.regras.find(x => x.id === id);
      if (!r || !exists('regra-form')) return;
      const f = $('regra-form');
      ['codvend', 'percentual', 'prioridade', 'codcliente', 'rede', 'uf', 'codprod', 'ativo', 'descricao'].forEach(k => {
        if (f[k]) f[k].value = r[k] ?? '';
      });
      $('regra-id').value = r.id;
      $('regra-form-title').textContent = `Editando Regra #${r.id}`;
      $('regra-submit-btn').textContent = ' Atualizar regra';
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  });

  container.querySelectorAll('button[data-regra-del]').forEach(b => {
    b.addEventListener('click', async () => {
      const ok = await confirmDialog({ title: 'Excluir Regra', message: 'Deseja excluir esta regra de comissao?', confirmText: 'Excluir', danger: true });
      if (!ok) return;
      try {
        await api(`/api/regras/${Number(b.dataset.regraDel)}`, { method: 'DELETE' });
        toast('Regra excluída.', 'success');
        await loadRegras(); await loadDashboard();
      } catch (e) { toast(e.message, 'error'); }
    });
  });
}

function initRegras() {
  if (!exists('regra-form')) return;
  $('regra-form').addEventListener('submit', async e => {
    e.preventDefault();
    const d = formData(e.target);
    const rid = Number($('regra-id').value || 0);
    const btn = e.target.querySelector('[type=submit]');
    btn.disabled = true;
    try {
      await api(rid ? `/api/regras/${rid}` : '/api/regras', {
        method: rid ? 'PUT' : 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ...d, percentual: Number(d.percentual || 0), prioridade: Number(d.prioridade || 100), ativo: Number(d.ativo || 1) }),
      });
      limparRegraForm();
      await loadRegras(); await loadDashboard();
      toast('Regra salva!', 'success');
    } catch (err) { toast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });

  if (exists('btn-regra-limpar')) $('btn-regra-limpar').addEventListener('click', limparRegraForm);
  if (exists('regra-search')) $('regra-search').addEventListener('input', renderRegras);
}

function limparRegraForm() {
  const f = $('regra-form');
  if (!f) return;
  f.reset();
  $('regra-id').value = '';
  f.ativo.value = '1';
  f.prioridade.value = '100';
  if ($('regra-form-title')) $('regra-form-title').textContent = 'Nova Regra';
  if ($('regra-submit-btn')) $('regra-submit-btn').textContent = ' Salvar regra';
}

/* 
   APURAO
 */
async function loadApuracaoMeta(mes, ano) {
  if (!exists('apuracao-id-label')) return;
  try {
    const ap = await api(`/api/apuracoes/${mes}/${ano}`);
    if (ap && ap.id) {
      $('apuracao-id-label').textContent = `Apura\u00e7\u00e3o ID ${ap.id} \u00b7 Per\u00edodo ${String(mes).padStart(2, '0')}/${ano} \u00b7 Arquivo: ${ap.arquivo_nome || 'importado'}`;
    } else {
      $('apuracao-id-label').textContent = `Per\u00edodo ${String(mes).padStart(2, '0')}/${ano} sem apura\u00e7\u00e3o importada.`;
    }
  } catch (e) {
    $('apuracao-id-label').textContent = '';
  }
}

async function initApuracao() {
  if (!exists('btn-calcular')) return;

  try {
    const periodos = await api('/api/lancamentos/periodos');
    const sel = $('sel-periodo');
    if (sel && periodos.length) {
      sel.innerHTML = periodos.map(p => `<option value="${p.mes}|${p.ano}">${String(p.mes).padStart(2,'0')}/${p.ano} (${p.total})</option>`).join('');
      sel.addEventListener('change', () => {
        const [m, a] = sel.value.split('|').map(Number);
        if ($('inp-mes')) $('inp-mes').value = m;
        if ($('inp-ano')) $('inp-ano').value = a;
        if (m && a) loadApuracaoMeta(m, a);
      });
      const first = periodos[0];
      if ($('inp-mes')) $('inp-mes').value = first.mes;
      if ($('inp-ano')) $('inp-ano').value = first.ano;
      await loadApuracaoMeta(first.mes, first.ano);
    }
  } catch (e) {}

  $('btn-calcular').addEventListener('click', async () => {
    const mes = Number($('inp-mes')?.value || 0);
    const ano = Number($('inp-ano')?.value || 0);
    if (!mes || !ano) return toast('Informe m\u00eas e ano.', 'error');
    const btn = $('btn-calcular');
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Calculando...';
    try {
      await api('/api/comissoes/calcular', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ mes, ano }),
      });
      await loadComissoes(mes, ano);
      await loadApuracaoMeta(mes, ano);
      toast('Apura\u00e7\u00e3o conclu\u00edda!', 'success');
    } catch (e) { toast(e.message, 'error'); }
    finally { btn.disabled = false; btn.innerHTML = 'Calcular'; }
  });

  $('btn-carregar').addEventListener('click', async () => {
    const mes = Number($('inp-mes')?.value || 0);
    const ano = Number($('inp-ano')?.value || 0);
    if (!mes || !ano) return toast('Informe m\u00eas e ano.', 'error');
    await loadComissoes(mes, ano);
    await loadApuracaoMeta(mes, ano);
  });

  if (exists('btn-excluir-apuracao')) {
    $('btn-excluir-apuracao').addEventListener('click', async () => {
      const mes = Number($('inp-mes')?.value || 0);
      const ano = Number($('inp-ano')?.value || 0);
      if (!mes || !ano) return toast('Informe m\u00eas e ano.', 'error');
      const ok = await confirmDialog({
        title: 'Excluir Apuracao',
        message: `Deseja excluir toda a apuração ${String(mes).padStart(2, '0')}/${ano}?`,
        confirmText: 'Excluir',
        danger: true,
      });
      if (!ok) return;
      const btn = $('btn-excluir-apuracao');
      btn.disabled = true;
      try {
        await api(`/api/apuracoes/${mes}/${ano}`, { method: 'DELETE' });
        STATE.comissoes = [];
        STATE.lancamentos = [];
        if (exists('comissoes-list')) $('comissoes-list').innerHTML = emptyState('X', 'Apura\u00e7\u00e3o exclu\u00edda', 'Importe a planilha novamente para recriar o per\u00edodo.');
        if (exists('card-lancamentos')) $('card-lancamentos').style.display = 'none';
        await loadApuracaoMeta(mes, ano);
        toast('Apura\u00e7\u00e3o exclu\u00edda.', 'success');
      } catch (e) {
        toast(e.message, 'error');
      } finally {
        btn.disabled = false;
      }
    });
  }

  if (exists('btn-fechar-lanc')) {

    $('btn-fechar-lanc').addEventListener('click', () => {
      $('card-lancamentos').style.display = 'none';
    });
  }

  if (exists('btn-pdf-lote')) {
    $('btn-pdf-lote').addEventListener('click', () => {
      const mes = Number($('inp-mes')?.value || 0);
      const ano = Number($('inp-ano')?.value || 0);
      if (!mes || !ano) return toast('Selecione o período.', 'error');
      window.open(`/api/relatorios/consolidado/${mes}/${ano}/pdf`, '_blank');
    });
    $('btn-csv-lote').addEventListener('click', () => {
      const mes = Number($('inp-mes')?.value || 0);
      const ano = Number($('inp-ano')?.value || 0);
      if (!mes || !ano) return toast('Selecione o período.', 'error');
      window.open(`/api/relatorios/consolidado/${mes}/${ano}/csv`, '_blank');
    });
  }
}

async function loadComissoes(mes, ano) {
  if (!exists('comissoes-list')) return;
  try {
    await loadReps();
    STATE.comissoes = await api(`/api/comissoes/${mes}/${ano}`);
  } catch (e) {
    $('comissoes-list').innerHTML = alertHtml('error', 'Erro ao carregar comissões', e.message);
    return;
  }

  if ($('comissoes-period-label')) $('comissoes-period-label').textContent = `${String(mes).padStart(2,'0')}/${ano}`;

  // KPIs
  const totalLiq = STATE.comissoes.reduce((s, c) => s + Number(c.total_vlrliq || 0), 0);
  const totalComBase = STATE.comissoes.reduce((s, c) => s + Number(c.total_comissao || 0), 0);
  const totalDescontos = STATE.comissoes.reduce((s, c) => s + Number(c.ajuste_desconto || 0), 0);
  const totalPremios = STATE.comissoes.reduce((s, c) => s + Number(c.ajuste_premio || 0), 0);
  const totalComFinal = totalComBase - totalDescontos + totalPremios;
  const aprov = STATE.comissoes.filter(c => {
    const st = String(c.status || '').toLowerCase();
    return st === 'aprovado' || st === 'enviado';
  }).length;
  const percTotal = totalLiq ? (totalComBase / totalLiq) * 100.0 : 0;

  if (exists('kpi-apuracao')) {
    $('kpi-apuracao').classList.remove('hidden');
    $('kpi-apuracao').innerHTML = `
      <div class="kpi-card kpi-blue"><div class="kpi-label">Representantes</div><div class="kpi-value">${STATE.comissoes.length}</div></div>
      <div class="kpi-card"><div class="kpi-label">Fat. Liquido Total</div><div class="kpi-value" style="font-size:16px">${money(totalLiq)}</div></div>
      <div class="kpi-card kpi-green"><div class="kpi-label">Comissao Final</div><div class="kpi-value" style="font-size:16px">${money(totalComFinal)}</div></div>
      <div class="kpi-card kpi-blue"><div class="kpi-label">% Comissao Base / Fat. Liquido</div><div class="kpi-value" style="font-size:16px">${pct2(percTotal)}</div></div>
      <div class="kpi-card kpi-amber"><div class="kpi-label">Aprovadas + Enviadas</div><div class="kpi-value">${aprov} / ${STATE.comissoes.length}</div></div>
    `;
    // show export buttons
    if (exists('btn-pdf-lote')) { $('btn-pdf-lote').classList.remove('hidden'); $('btn-csv-lote').classList.remove('hidden'); }
  }

  const container = $('comissoes-list');
  if (!STATE.comissoes.length) {
    container.innerHTML = emptyState('', 'Sem comissões para o período', 'Verifique se as regras estão configuradas e tente calcular novamente.');
    return;
  }

  const rows = STATE.comissoes.map(c => {
    const rep = STATE.repByCodvend[String(c.codvend || '')];
    const rid = rep ? rep.id : '';
    const status = String(c.status || '').toLowerCase();
    const podeAprovar = status === 'rascunho';
    const podeCancelar = status === 'aprovado';
    const podeEnviar = status === 'aprovado';
    const desconto = Number(c.ajuste_desconto || 0);
    const premio = Number(c.ajuste_premio || 0);
    const comFinal = Number(c.total_comissao || 0) - desconto + premio;
    const percLinha = Number(c.total_vlrliq || 0) ? (Number(c.total_comissao || 0) / Number(c.total_vlrliq || 0)) * 100.0 : 0;
    return `<tr>
      <td><input type="checkbox" class="rep-chk" data-rep-id="${rid}" ${!rid ? 'disabled' : ''} ${STATE.selectedRepIds.has(Number(rid)) ? 'checked' : ''}></td>
      <td class="mono">${c.id}</td>
      <td class="mono fw-bold">${c.codvend}</td>
      <td>${rep?.nome || '<span class="text-muted"></span>'}</td>
      <td class="text-sm text-muted">${rep?.email || ''}</td>
      <td class="text-right">${money(c.total_vlrliq)}</td>
      <td class="text-right">${money(c.total_comissao)}</td>
      <td class="text-right" style="color:#b45309">${money(desconto)}</td>
      <td class="text-right" style="color:#0f766e">${money(premio)}</td>
      <td class="text-right fw-bold" style="color:var(--green)">${money(comFinal)}</td>
      <td class="text-right fw-bold">${pct2(percLinha)}</td>
      <td>${badgeHtml(c.status)}</td>
      <td>
        <div class="table-actions">
          <button class="btn btn-ghost btn-sm" data-lanc="${c.id}" data-rep-nome="${rep?.nome || c.codvend}">Lan\u00e7amentos</button>
          ${podeAprovar ? `<button class="btn btn-success btn-sm" data-aprovar="${c.id}">Aprovar</button>` : ''}
          ${podeCancelar ? `<button class="btn btn-danger btn-sm" data-cancelar="${c.id}">Cancelar</button>` : ''}
          ${rid ? `<button class="btn btn-ghost btn-sm" data-pdf="${rid}" title="Gerar PDF">PDF</button>` : ''}
          ${rid ? `<button class="btn btn-warn btn-sm" data-email-comissao="${c.id}" title="Enviar e-mail" ${podeEnviar ? '' : 'disabled'}>E-mail</button>` : ''}
        </div>
      </td>
    </tr>`;
  }).join('');

  container.innerHTML = `<div class="table-wrapper"><table class="table-apuracao">
    <thead><tr>
      <th style="width:30px"><input type="checkbox" id="chk-all" title="Selecionar todos"></th>
      <th>ID</th><th>CODVEND</th><th>Representante</th><th>E-mail</th>
      <th class="text-right">Vlr Liq.</th><th class="text-right">Com. Base</th><th class="text-right">Desc.</th><th class="text-right">Premio</th><th class="text-right">Com. Final</th><th class="text-right">% Base</th>
      <th>Status</th><th>Ações</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table></div>`;

  // Select all
  $('chk-all')?.addEventListener('change', e => {
    container.querySelectorAll('.rep-chk:not(:disabled)').forEach(cb => {
      cb.checked = e.target.checked;
      const id = Number(cb.dataset.repId || 0);
      if (id) { e.target.checked ? STATE.selectedRepIds.add(id) : STATE.selectedRepIds.delete(id); }
    });
    persistSelectedRepIds();
    syncEnvioIdsField();
  });

  container.querySelectorAll('.rep-chk').forEach(cb => {
    cb.addEventListener('change', () => {
      const id = Number(cb.dataset.repId || 0);
      if (!id) return;
      cb.checked ? STATE.selectedRepIds.add(id) : STATE.selectedRepIds.delete(id);
      persistSelectedRepIds();
      syncEnvioIdsField();
    });
  });

  // Launch details
  container.querySelectorAll('button[data-lanc]').forEach(b => b.addEventListener('click', async () => {
    try {
      const id = Number(b.dataset.lanc);
      STATE.lancamentos = await api(`/api/comissoes/${id}/lancamentos`);
      const com = STATE.comissoes.find(x => Number(x.id) === id);
      STATE.currentComissaoId = id;
      STATE.currentComissaoStatus = String(com?.status || '').toLowerCase();
      $('lanc-rep-nome').textContent = b.dataset.repNome || id;
      renderLancamentosApuracao();
      $('card-lancamentos').style.display = 'block';
      $('card-lancamentos').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) { toast(e.message, 'error'); }
  }));

  container.querySelectorAll('button[data-aprovar]').forEach(b => b.addEventListener('click', async () => {
    b.disabled = true; b.innerHTML = '<span class="spinner"></span>';
    try {
      await api(`/api/comissoes/${Number(b.dataset.aprovar)}/aprovar`, { method: 'POST' });
      toast('Comissão aprovada!', 'success');
      await loadComissoes(mes, ano);
    } catch (e) { toast(e.message, 'error'); b.disabled = false; b.innerHTML = 'Aprovar'; }
  }));

  container.querySelectorAll('button[data-cancelar]').forEach(b => b.addEventListener('click', async () => {
    const ok = await confirmDialog({
      title: 'Cancelar Aprovacao',
      message: 'Deseja cancelar a aprovacao desta comissao?',
      confirmText: 'Cancelar aprovacao',
      danger: true,
    });
    if (!ok) return;
    b.disabled = true; b.innerHTML = '<span class="spinner"></span>';
    try {
      await api(`/api/comissoes/${Number(b.dataset.cancelar)}/cancelar-aprovacao`, { method: 'POST' });
      toast('Aprovacao cancelada.', 'success');
      await loadComissoes(mes, ano);
    } catch (e) {
      toast(e.message, 'error');
      b.disabled = false; b.innerHTML = 'Cancelar';
    }
  }));

  container.querySelectorAll('button[data-pdf]').forEach(b => b.addEventListener('click', () => {
    window.open(`/api/relatorios/${Number(b.dataset.pdf)}/pdf`, '_blank');
  }));


  container.querySelectorAll('button[data-email-comissao]').forEach(b => b.addEventListener('click', async () => {
    const cid = Number(b.dataset.emailComissao || 0);
    b.disabled = true; b.innerHTML = '<span class="spinner spinner-dark"></span>';
    try {
      const r = await api(`/api/comissoes/${cid}/enviar-email`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ assunto: 'Relatorio de Comissao', anexos: [] }),
      });
      toast(r.status === 'ok' ? 'E-mail enviado!' : 'Falha no envio', r.status === 'ok' ? 'success' : 'error');
      await loadComissoes(mes, ano);
    } catch (e) { toast(e.message, 'error'); }
    finally { b.disabled = false; b.innerHTML = 'E-mail'; }
  }));
}

function renderLancamentosApuracao() {
  const container = $('lanc-list');
  if (!container) return;
  const statusComissao = String(STATE.currentComissaoStatus || '').toLowerCase();
  const bloqueado = statusComissao === 'aprovado' || statusComissao === 'enviado';
  const bloqueioFinanceiro = statusComissao === 'aprovado' || statusComissao === 'enviado';
  const comissaoSel = STATE.comissoes.find(x => Number(x.id) === Number(STATE.currentComissaoId)) || {};

  const totalLiq = STATE.lancamentos.reduce((s, l) => s + Number(l.vlrliq || 0), 0);
  const totalComBase = STATE.lancamentos.reduce((s, l) => s + Number(l.tcomisprod || 0), 0);
  const descontoAtual = Number(comissaoSel.ajuste_desconto || 0);
  const premioAtual = Number(comissaoSel.ajuste_premio || 0);
  const totalFinal = totalComBase - descontoAtual + premioAtual;
  const percComissao = totalLiq ? (totalComBase / totalLiq) * 100.0 : 0;

  if ($('lanc-kpis')) {
    $('lanc-kpis').innerHTML = `
      <div class="kpi-card"><div class="kpi-label">Lancamentos</div><div class="kpi-value">${STATE.lancamentos.length}</div></div>
      <div class="kpi-card"><div class="kpi-label">Vlr. Liquido</div><div class="kpi-value" style="font-size:16px">${money(totalLiq)}</div></div>
      <div class="kpi-card kpi-green"><div class="kpi-label">Comissao Final</div><div class="kpi-value" style="font-size:16px">${money(totalFinal)}</div></div>
      <div class="kpi-card kpi-blue"><div class="kpi-label">% Comissao Base / Vlr. Liquido</div><div class="kpi-value" style="font-size:16px">${pct2(percComissao)}</div></div>
    `;
  }

  if (exists('lanc-ajustes-fin')) {
    $('lanc-ajustes-fin').innerHTML = `
      <div class="card" style="margin:0">
        <div class="card-body">
          <div class="form-grid cols-4">
            <div class="form-group">
              <label class="form-label">Desconto na comissao (R$)</label>
              <input id="ajf-desconto" class="form-control" type="number" step="0.01" value="${numInput(descontoAtual)}" ${bloqueioFinanceiro ? 'disabled' : ''}>
            </div>
            <div class="form-group">
              <label class="form-label">Premiacao (R$)</label>
              <input id="ajf-premio" class="form-control" type="number" step="0.01" value="${numInput(premioAtual)}" ${bloqueioFinanceiro ? 'disabled' : ''}>
            </div>
            <div class="form-group form-col-full">
              <label class="form-label">Observacao dos ajustes</label>
              <input id="ajf-obs" class="form-control" type="text" placeholder="Ex: desconto por avaria / premio por meta" value="${esc(comissaoSel.ajuste_obs || '')}" ${bloqueioFinanceiro ? 'disabled' : ''}>
            </div>
            <div class="form-group">
              <label class="form-label">&nbsp;</label>
              <button id="btn-ajf-salvar" class="btn btn-primary" ${bloqueioFinanceiro ? 'disabled' : ''}>Salvar ajustes financeiros</button>
            </div>
          </div>
        </div>
      </div>`;

    $('btn-ajf-salvar')?.addEventListener('click', async () => {
      const cid = Number(STATE.currentComissaoId || 0);
      if (!cid) return;
      const desconto = toFloatOrNull($('ajf-desconto')?.value) || 0;
      const premio = toFloatOrNull($('ajf-premio')?.value) || 0;
      const observacao = $('ajf-obs')?.value || '';
      const btn = $('btn-ajf-salvar');
      btn.disabled = true;
      btn.innerHTML = '<span class="spinner"></span> Salvando...';
      try {
        await api(`/api/comissoes/${cid}/ajustes-financeiros`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ desconto, premio, observacao }),
        });
        const mes = Number($('inp-mes')?.value || 0);
        const ano = Number($('inp-ano')?.value || 0);
        if (mes && ano) await loadComissoes(mes, ano);
        renderLancamentosApuracao();
        toast('Ajustes financeiros salvos.', 'success');
      } catch (e) {
        toast(e.message, 'error');
      } finally {
        btn.disabled = false;
        btn.innerHTML = 'Salvar ajustes financeiros';
      }
    });
  }

  if (!STATE.lancamentos.length) {
    container.innerHTML = emptyState('X', 'Nenhum lancamento encontrado');
    return;
  }

  const rows = STATE.lancamentos.map(l => {
    const vend = l.comis_vend != null ? Number(l.comis_vend) : 0;
    const prod = l.comis_prod != null ? Number(l.comis_prod) : 0;
    const percBase = Math.abs(prod) > 1e-12 ? prod : vend;
    const comCalc = Number(l.vlrliq || 0) * (percBase / 100);
    return `
    <tr data-lid="${l.id}" data-vlrliq="${Number(l.vlrliq || 0)}">
      <td class="mono">${l.nf || ''}</td>
      <td class="mono">${l.pedido || ''}</td>
      <td class="mono text-sm">${l.codprod || ''}</td>
      <td class="truncate" title="${l.produto || ''}">${l.produto || ''}</td>
      <td class="text-right">${money(l.vlrliq)}</td>
      <td class="text-right"><input class="inline-input w120 f-vend" value="${num(vend)}" ${bloqueado ? 'disabled' : ''}></td>
      <td class="text-right"><input class="inline-input w120 f-prod" value="${num(prod)}" ${bloqueado ? 'disabled' : ''}></td>
      <td class="text-right fw-bold f-comissao" style="color:var(--green)">${money(comCalc)}</td>
      <td>
        <input class="inline-input w180 f-motivo" placeholder="Motivo do ajuste" value="Ajuste manual" ${bloqueado ? 'disabled' : ''}>
        <button class="btn btn-ghost btn-sm mt-12" data-lsave="${l.id}" style="margin-top:4px" ${bloqueado ? 'disabled' : ''}>Salvar</button>
      </td>
    </tr>`;
  }).join('');

  const avisoBloqueio = bloqueado ? `<div class="alert alert-warn"><span class="alert-icon">!</span><div class="alert-body"><div class="alert-title">Edicao bloqueada</div><div>Comissao aprovada/enviada nao permite alteracao de lancamentos.</div></div></div>` : '';
  container.innerHTML = `${avisoBloqueio}<div class="table-wrapper"><table>
    <thead><tr>
      <th>NF</th><th>Pedido</th><th>Cd. Produto</th><th>Produto</th>
      <th class="text-right">Vlr. Liquido</th><th class="text-right">% Vend</th>
      <th class="text-right">% Prod</th><th class="text-right">Comissao</th><th>Ajuste</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table></div>`;

  const recalcRow = row => {
    const vlrliq = Number(row.dataset.vlrliq || 0);
    const vend = toFloatOrNull(row.querySelector('.f-vend')?.value);
    const prod = toFloatOrNull(row.querySelector('.f-prod')?.value);
    const percBase = (prod !== null && Math.abs(prod) > 1e-12) ? prod : (vend || 0);
    const comissao = vlrliq * (percBase / 100.0);
    row.querySelector('.f-comissao').textContent = money(comissao);
  };

  container.querySelectorAll('tr[data-lid]').forEach(row => {
    row.querySelector('.f-vend')?.addEventListener('input', () => recalcRow(row));
    row.querySelector('.f-prod')?.addEventListener('input', () => recalcRow(row));
  });

  container.querySelectorAll('button[data-lsave]').forEach(b => b.addEventListener('click', async () => {
    if (bloqueado) return;
    const lid = Number(b.dataset.lsave);
    const row = container.querySelector(`tr[data-lid="${lid}"]`);
    const comisVend = toFloatOrNull(row.querySelector('.f-vend')?.value);
    const comisProd = toFloatOrNull(row.querySelector('.f-prod')?.value);
    const motivo = row.querySelector('.f-motivo')?.value || 'Ajuste manual';
    if (comisVend === null && comisProd === null) {
      toast('Informe ao menos um percentual valido.', 'error');
      return;
    }
    b.disabled = true;
    b.innerHTML = '<span class="spinner spinner-dark"></span>';
    try {
      const resp = await api(`/api/lancamentos/${lid}/percentuais`, {
        method: 'PUT', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ comis_vend: comisVend, comis_prod: comisProd, motivo }),
      });
      const atual = resp?.lancamento || {};
      const idx = STATE.lancamentos.findIndex(x => Number(x.id) === lid);
      if (idx >= 0) STATE.lancamentos[idx] = { ...STATE.lancamentos[idx], ...atual };
      recalcRow(row);
      const mes = Number($('inp-mes')?.value || 0);
      const ano = Number($('inp-ano')?.value || 0);
      if (mes && ano) await loadComissoes(mes, ano);
      renderLancamentosApuracao();
      toast(`Lancamento ${lid} atualizado.`, 'success');
    } catch (e) {
      toast(e.message, 'error');
    } finally {
      b.disabled = false;
      b.innerHTML = 'Salvar';
    }
  }));
}
/* 
   ENVIO
 */
async function loadEmailHist() {
  if (!exists('email-hist')) return;
  try {
    const hist = await api('/api/email/historico');
    if ($('hist-count-label')) $('hist-count-label').textContent = `${hist.length} registro(s)`;
    if (!hist.length) {
      $('email-hist').innerHTML = emptyState('', 'Nenhum envio registrado');
      return;
    }
    const rows = hist.slice(0, 200).map(h => `<tr>
      <td class="mono">${h.id}</td>
      <td>${h.representante_id}</td>
      <td>${h.destinatario || ''}</td>
      <td>${badgeHtml(h.status)}</td>
      <td><span class="badge badge-inativo">${h.tipo || ''}</span></td>
      <td class="text-sm text-muted">${h.data || ''}</td>
    </tr>`).join('');
    $('email-hist').innerHTML = `<div class="table-wrapper"><table>
      <thead><tr><th>ID</th><th>Rep. ID</th><th>Destinatário</th><th>Status</th><th>Tipo</th><th>Data</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div>`;
  } catch (e) { toast(e.message, 'error'); }
}

function initEnvio() {
  if (!exists('btn-enviar-lote')) return;
  syncEnvioIdsField();
  $('btn-enviar-lote').addEventListener('click', async () => {
    const idsStr = $('email-ids')?.value || '';
    const idsFromField = idsStr.split(',').map(x => Number(x.trim())).filter(Boolean);
    const ids = idsFromField.length ? idsFromField : Array.from(STATE.selectedRepIds);
    if (!ids.length) return toast('Informe os IDs ou selecione representantes na apuração.', 'error');
    const assunto = $('email-assunto')?.value || 'Relatório de Comissão';

    const btn = $('btn-enviar-lote');
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Enviando...';
    $('envio-progress-area').classList.remove('hidden');

    try {
      const r = await api('/api/email/lote', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids, assunto, anexos: [] }),
      });
      const resultados = r.resultados || [];
      const ok = resultados.filter(x => x.status === 'ok').length;
      const err = resultados.filter(x => x.status !== 'ok').length;
      $('envio-result-area').innerHTML = alertHtml(err ? 'warn' : 'success',
        `Lote concluído: ${ok} enviados${err ? `, ${err} falha(s)` : ''}`,
        resultados.map(x => `<span class="badge badge-${x.status === 'ok' ? 'ok' : 'erro'}" style="margin:2px">ID ${x.id}: ${x.status}</span>`).join(''));
      await loadEmailHist();
      toast(`${ok} e-mail(s) enviado(s).`, ok === ids.length ? 'success' : 'default');
    } catch (e) {
      $('envio-result-area').innerHTML = alertHtml('error', 'Erro no envio em lote', e.message);
      toast(e.message, 'error');
    } finally {
      btn.disabled = false; btn.innerHTML = ' Enviar em lote';
      $('envio-progress-area').classList.add('hidden');
    }
  });

  if (exists('btn-limpar-sel')) $('btn-limpar-sel').addEventListener('click', () => {
    STATE.selectedRepIds.clear();
    persistSelectedRepIds();
    if ($('email-ids')) $('email-ids').value = '';
    toast('Seleção limpa.');
  });

  if (exists('btn-refresh-hist')) $('btn-refresh-hist').addEventListener('click', loadEmailHist);
}

/* 
   CONFIGURAES
 */
async function loadCfg() {
  if (!exists('cfg-form')) return;
  try {
    const cfg = await api('/api/configuracoes');
    const f = $('cfg-form');
    f.smtp_host.value = cfg.smtp_host || '';
    f.smtp_port.value = cfg.smtp_port || 587;
    f.smtp_user.value = cfg.smtp_user || '';
    f.smtp_pass.value = cfg.smtp_pass || '';
    f.smtp_from.value = cfg.smtp_from || '';
  } catch (e) {}
}

function initConfiguracoes() {
  if (!exists('cfg-form')) return;
  $('cfg-form').addEventListener('submit', async e => {
    e.preventDefault();
    const btn = e.target.querySelector('[type=submit]');
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Salvando...';
    try {
      await api('/api/configuracoes', {
        method: 'PUT', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(formData(e.target)),
      });
      if (exists('cfg-result-area')) $('cfg-result-area').innerHTML = alertHtml('success', 'Salvo!', 'Configurações SMTP salvas com sucesso.');
      toast('SMTP salvo!', 'success');
    } catch (err) { toast(err.message, 'error'); }
    finally { btn.disabled = false; btn.innerHTML = ' Salvar configurações'; }
  });

  if (exists('btn-test-smtp')) {
    $('btn-test-smtp').addEventListener('click', async () => {
      const btn = $('btn-test-smtp');
      btn.disabled = true; btn.innerHTML = '<span class="spinner spinner-dark"></span> Testando...';
      try {
        const r = await api('/api/configuracoes/testar-smtp', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(formData($('cfg-form'))),
        });
        const ok = r.status === 'ok';
        if (exists('cfg-result-area'))
          $('cfg-result-area').innerHTML = alertHtml(ok ? 'success' : 'error', ok ? 'Conexão OK!' : 'Falha na conexão', ok ? 'SMTP configurado corretamente.' : (r.motivo || 'Verifique as configurações.'));
        toast(ok ? 'Conexão OK!' : 'Falha no teste.', ok ? 'success' : 'error');
      } catch (e) { toast(e.message, 'error'); }
      finally { btn.disabled = false; btn.innerHTML = ' Testar conexão'; }
    });
  }
}

/*  utils  */
function esc(s) { return String(s || '').replace(/"/g, '&quot;').replace(/</g, '&lt;'); }

/* 
   BOOTSTRAP
 */
async function bootstrap() {
  hydrateSelectedRepIds();
  // Init section-specific modules
  initImportacao();
  initRepresentantes();
  initRegras();
  await initApuracao();
  initEnvio();
  initConfiguracoes();

  // Load data
  try {
    await Promise.all([
      loadDashboard(),
      loadCfg(),
      loadReps(),
      loadRegras(),
      loadEmailHist(),
    ]);
  } catch (e) {
    console.warn('Bootstrap partial error:', e);
  }

  // Dashboard refresh button
  if (exists('btn-refresh')) $('btn-refresh').addEventListener('click', async () => {
    try { await loadDashboard(); toast('Atualizado.', 'success'); } catch (e) { toast(e.message, 'error'); }
  });
}

document.addEventListener('DOMContentLoaded', bootstrap);



