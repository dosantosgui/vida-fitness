/* Vida Fitness — app de registro de treinos (fase 1: offline, dados no celular).
   O app NÃO cria nem altera treinos: só importa o plano e registra a execução. */
'use strict';

const VERSAO_APP = '1.0.1';
const DIAS = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
const DIAS_LONGOS = ['domingo', 'segunda', 'terça', 'quarta', 'quinta', 'sexta', 'sábado'];

const CONFIG_PADRAO = {
  incrementos: { 'Barra': 5, 'Barra T': 2.5, 'Halteres <10 kg': 1, 'Halteres ≥10 kg': 2, 'Máquina': 5, 'Polia': 5, 'Peso corporal': 2.5 },
  descanso: { ativacao: 45, trabalhoBloco: 75, trabalhoSimples: 150 }
};

/* ---------- utilidades ---------- */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad2 = (n) => String(n).padStart(2, '0');
const isoDe = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const dataDe = (iso) => { const [y, m, d] = iso.split('-').map(Number); return new Date(y, m - 1, d); };
const hojeISO = () => isoDe(new Date());
const somaDias = (iso, n) => { const d = dataDe(iso); d.setDate(d.getDate() + n); return isoDe(d); };
const segundaDe = (iso) => { const d = dataDe(iso); d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); return isoDe(d); };
const ddmm = (iso) => { const [, m, d] = iso.split('-'); return `${d}/${m}`; };
const diaSem = (iso) => DIAS[dataDe(iso).getDay()];
const fmtNum = (n) => (n === null || n === undefined || n === '') ? '' : String(Math.round(n * 100) / 100).replace('.', ',');
const lerNum = (v) => { if (v === null || v === undefined) return null; const s = String(v).trim().replace(',', '.'); if (s === '') return null; const n = Number(s); return Number.isFinite(n) ? n : null; };
const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
const fmtSeg = (s) => { s = Math.max(0, Math.round(s)); return `${Math.floor(s / 60)}:${pad2(s % 60)}`; };
const fmtTempo = (s) => { if (s == null) return ''; s = Math.round(s); const h = Math.floor(s / 3600); const m = Math.floor((s % 3600) / 60); const x = s % 60; return h ? `${h}:${pad2(m)}:${pad2(x)}` : `${m}:${pad2(x)}`; };
const lerTempo = (t) => { if (!t) return null; const p = String(t).trim().split(':').map((x) => Number(x)); if (p.some((x) => !Number.isFinite(x))) return null; if (p.length === 3) return p[0] * 3600 + p[1] * 60 + p[2]; if (p.length === 2) return p[0] * 60 + p[1]; if (p.length === 1) return p[0] * 60; return null; };
const zonaDe = (z) => { const n = String(z ?? '').match(/\d+/g); if (!n) return null; return [Number(n[0]), Number(n[n.length - 1])]; };

function toast(msg, ms = 2500) {
  const t = $('#toast'); t.textContent = msg; t.classList.remove('oculto');
  clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.add('oculto'), ms);
}

/* ---------- banco local (IndexedDB) ---------- */
const DB = {
  _db: null,
  abrir() {
    return new Promise((ok, erro) => {
      const r = indexedDB.open('vida-fitness', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('kv');
      r.onsuccess = () => { this._db = r.result; ok(); };
      r.onerror = () => erro(r.error);
    });
  },
  ler(chave) {
    return new Promise((ok, erro) => {
      const r = this._db.transaction('kv').objectStore('kv').get(chave);
      r.onsuccess = () => ok(r.result); r.onerror = () => erro(r.error);
    });
  },
  gravar(chave, valor) {
    return new Promise((ok, erro) => {
      const tx = this._db.transaction('kv', 'readwrite');
      tx.objectStore('kv').put(valor, chave);
      tx.oncomplete = () => ok(); tx.onerror = () => erro(tx.error);
    });
  }
};

const E = { plano: null, config: null, regMusc: {}, sessoes: {}, regCorrida: {}, ultimoBackup: null };

async function carregar() {
  await DB.abrir();
  E.plano = (await DB.ler('plano')) || null;
  E.config = Object.assign({}, structuredClone(CONFIG_PADRAO), (await DB.ler('config')) || {});
  E.config.incrementos = Object.assign({}, CONFIG_PADRAO.incrementos, E.config.incrementos);
  E.config.descanso = Object.assign({}, CONFIG_PADRAO.descanso, E.config.descanso);
  E.regMusc = (await DB.ler('regMusc')) || {};
  E.sessoes = (await DB.ler('sessoes')) || {};
  E.regCorrida = (await DB.ler('regCorrida')) || {};
  E.ultimoBackup = (await DB.ler('ultimoBackup')) || null;
}
const salvar = (chave) => DB.gravar(chave, E[chave]).catch((e) => { alert('Erro ao salvar: ' + e.message); });

/* ---------- importação do plano (.xlsx) ---------- */
function lerData(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') { const p = XLSX.SSF.parse_date_code(v); return `${p.y}-${pad2(p.m)}-${pad2(p.d)}`; }
  if (v instanceof Date) return isoDe(v);
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;
  return null;
}
function lerBool(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  const s = norm(v); return ['true', 'verdadeiro', 'sim', 'x', '1'].includes(s);
}
function lerAba(wb, nome, colunas) {
  const aba = wb.Sheets[nome];
  if (!aba) throw new Error(`A aba "${nome}" não foi encontrada no arquivo.`);
  const linhas = XLSX.utils.sheet_to_json(aba, { header: 1, raw: true, defval: null, blankrows: false });
  if (!linhas.length) throw new Error(`A aba "${nome}" está vazia.`);
  const cab = linhas[0].map(norm);
  const idx = {};
  for (const [chave, titulo] of Object.entries(colunas)) {
    const i = cab.indexOf(norm(titulo));
    if (i < 0) throw new Error(`Coluna "${titulo}" não encontrada na aba "${nome}".`);
    idx[chave] = i;
  }
  return linhas.slice(1)
    .filter((l) => l.some((c) => c !== null && c !== ''))
    .map((l) => { const o = {}; for (const k in idx) o[k] = l[idx[k]]; return o; });
}
function interpretarPlanilha(wb) {
  const ex = lerAba(wb, 'Exercicios', { id: 'ID', nome: 'Exercício', equip: 'Equipamento', tipo: 'Tipo', zona: 'Zona', grupo: 'Grupo', m1: 'Músculo 1', p1: 'Peso 1', m2: 'Músculo 2', p2: 'Peso 2', uso: 'Uso' });
  const mu = lerAba(wb, 'Plano_Musculacao', { semana: 'Semana', data: 'Data', dia: 'Dia', sessao: 'Sessão', ordem: 'Ordem', bloco: 'Bloco', id: 'ID', nome: 'Exercício', ativ: 'Séries ativação', repsAtiv: 'Reps ativação', trab: 'Séries trabalho', zona: 'Zona', conta: 'Conta no volume', obs: 'Observação' });
  const co = lerAba(wb, 'Plano_Corrida', { semana: 'Semana', data: 'Data', dia: 'Dia', tipo: 'Tipo', dist: 'Distância (km)', pace: 'Pace alvo', strides: 'Strides', deload: 'Deload', obs: 'Observação' });

  const exercicios = {};
  for (const e of ex) {
    if (!e.id) continue;
    exercicios[String(e.id).trim()] = {
      id: String(e.id).trim(), nome: String(e.nome || '').trim(), equip: String(e.equip || '').trim(), tipo: String(e.tipo || '').trim(),
      zona: String(e.zona || '').trim(), grupo: lerNum(e.grupo), m1: e.m1 || null, p1: lerNum(e.p1), m2: e.m2 || null, p2: lerNum(e.p2), uso: e.uso || ''
    };
  }
  const erros = [];
  const musc = mu.map((r, i) => {
    const o = {
      semana: lerNum(r.semana), data: lerData(r.data), dia: r.dia, sessao: String(r.sessao || '').trim(), ordem: lerNum(r.ordem),
      bloco: String(r.bloco || 'Simples').trim(), id: String(r.id || '').trim(), nome: String(r.nome || '').trim(),
      ativ: lerNum(r.ativ) || 0, repsAtiv: lerNum(r.repsAtiv), trab: lerNum(r.trab) || 0, zona: String(r.zona || '').trim(),
      conta: lerBool(r.conta), obs: r.obs ? String(r.obs) : ''
    };
    if (!o.data) erros.push(`Plano_Musculacao, linha ${i + 2}: data inválida.`);
    if (!exercicios[o.id]) erros.push(`Plano_Musculacao, linha ${i + 2}: ID ${o.id} não existe em Exercicios.`);
    return o;
  });
  const corrida = co.map((r, i) => {
    const o = {
      semana: lerNum(r.semana), data: lerData(r.data), dia: r.dia, tipo: String(r.tipo || '').trim(), dist: lerNum(r.dist),
      pace: r.pace ? String(r.pace) : '', strides: lerNum(r.strides) || 0, deload: lerBool(r.deload), obs: r.obs ? String(r.obs) : ''
    };
    if (!o.data) erros.push(`Plano_Corrida, linha ${i + 2}: data inválida.`);
    return o;
  });
  if (erros.length) throw new Error(erros.slice(0, 5).join('\n') + (erros.length > 5 ? `\n…e mais ${erros.length - 5}.` : ''));
  return { importadoEm: new Date().toISOString(), exercicios, musc, corrida };
}

/* ---------- visão do plano ---------- */
function sessoesDoPlano() {
  if (!E.plano) return [];
  const mapa = new Map();
  for (const r of E.plano.musc) {
    const chave = `${r.data}_${r.sessao}`;
    if (!mapa.has(chave)) mapa.set(chave, { chave, data: r.data, semana: r.semana, sessao: r.sessao, itens: [] });
    mapa.get(chave).itens.push(r);
  }
  const lista = [...mapa.values()];
  for (const s of lista) {
    s.itens.sort((a, b) => a.ordem - b.ordem);
    s.deload = s.itens.some((i) => /deload/i.test(i.obs)) || E.plano.corrida.some((c) => c.semana === s.semana && c.deload && segundaDe(c.data) === segundaDe(s.data));
  }
  return lista.sort((a, b) => a.data.localeCompare(b.data));
}
const sessaoPorChave = (k) => sessoesDoPlano().find((s) => s.chave === k);

function blocosDe(sessao) {
  const blocos = [];
  for (const it of sessao.itens) {
    const ultimo = blocos[blocos.length - 1];
    if (it.bloco.toLowerCase() !== 'simples' && ultimo && ultimo.nome === it.bloco) ultimo.itens.push(it);
    else blocos.push({ nome: it.bloco, itens: [it] });
  }
  return blocos;
}
function seriesDe(sessao, it) {
  const s = [];
  for (let i = 1; i <= it.ativ; i++) s.push({ id: `${sessao.chave}|${it.ordem}|${s.length + 1}`, n: s.length + 1, tipo: 'A', idx: i, alvo: it.repsAtiv });
  for (let i = 1; i <= it.trab; i++) s.push({ id: `${sessao.chave}|${it.ordem}|${s.length + 1}`, n: s.length + 1, tipo: 'T', idx: i, alvo: it.zona });
  return s;
}
// Ordem de execução: dentro de um bi/tri-set, alterna os exercícios série a série.
function ordemExecucao(sessao) {
  const ordem = [];
  for (const b of blocosDe(sessao)) {
    const listas = b.itens.map((it) => seriesDe(sessao, it).map((s) => ({ ...s, it, bloco: b })));
    const max = Math.max(...listas.map((l) => l.length));
    for (let i = 0; i < max; i++) for (const l of listas) if (l[i]) ordem.push(l[i]);
  }
  return ordem;
}

function semanasDoPlano() {
  if (!E.plano) return [];
  const mapa = new Map();
  for (const r of [...E.plano.musc, ...E.plano.corrida]) {
    const seg = segundaDe(r.data);
    if (!mapa.has(seg)) mapa.set(seg, { segunda: seg, semana: r.semana });
  }
  return [...mapa.values()].sort((a, b) => a.segunda.localeCompare(b.segunda));
}
const corridaDoDia = (iso) => E.plano ? E.plano.corrida.filter((c) => c.data === iso) : [];
const sessoesDoDia = (iso) => sessoesDoPlano().filter((s) => s.data === iso);

/* ---------- registros e sugestão de carga ---------- */
const metaSessao = (k) => E.sessoes[k] || null;
const exFeito = (sessaoChave, it) => (E.sessoes[sessaoChave]?.ex?.[it.ordem]?.exFeito) || it.id;
const registrosDaSessao = (k) => Object.values(E.regMusc).filter((r) => r.sessao === k);
const statusSessao = (k) => {
  const m = metaSessao(k);
  if (!m) return 'nao';
  return m.status === 'concluida' ? 'ok' : 'andamento';
};

function incrementoDe(equip, carga) {
  const e = norm(equip); const inc = E.config.incrementos;
  if (e.startsWith('halter')) return (carga ?? 0) < 10 ? inc['Halteres <10 kg'] : inc['Halteres ≥10 kg'];
  if (e.startsWith('barrat')) return inc['Barra T'];
  if (e.startsWith('barra')) return inc['Barra'];
  if (e.startsWith('maquina')) return inc['Máquina'];
  if (e.startsWith('polia')) return inc['Polia'];
  if (e.startsWith('pesocorporal')) return inc['Peso corporal'];
  return 2.5;
}
function rotuloCarga(equip) {
  const e = norm(equip);
  if (e.startsWith('pesocorporal')) return 'Lastro';
  if (e.startsWith('halter')) return 'Kg/halter';
  return 'Kg';
}

// Última sessão (fora a atual e, se a atual não é deload, fora as de deload) com séries de trabalho do exercício.
function ultimaSessaoDoEx(exId, chaveAtual, atualEhDeload) {
  const porSessao = {};
  for (const r of Object.values(E.regMusc)) {
    if (r.exFeito !== exId || r.sessao === chaveAtual) continue;
    if (!atualEhDeload && E.sessoes[r.sessao]?.deload) continue;
    (porSessao[r.sessao] ||= []).push(r);
  }
  const chaves = Object.keys(porSessao).filter((k) => porSessao[k].some((r) => r.tipo === 'T' && r.carga !== null));
  if (!chaves.length) return null;
  chaves.sort((a, b) => {
    const da = E.sessoes[a]?.dataReal || a, db = E.sessoes[b]?.dataReal || b;
    return da.localeCompare(db) || (E.sessoes[a]?.inicio || '').localeCompare(E.sessoes[b]?.inicio || '');
  });
  const regs = porSessao[chaves[chaves.length - 1]].sort((a, b) => a.n - b.n);
  return { trabalho: regs.filter((r) => r.tipo === 'T'), ativacao: regs.filter((r) => r.tipo === 'A') };
}

function sugestao(sessao, it) {
  const id = exFeito(sessao.chave, it);
  const ex = E.plano.exercicios[id] || { equip: '', zona: it.zona };
  const zona = zonaDe(id === it.id ? it.zona : (ex.zona || it.zona));
  const ult = ultimaSessaoDoEx(id, sessao.chave, sessao.deload);
  if (!ult) return { carga: null, ativ: {}, texto: 'Sem histórico · calibrar a carga', sobe: false };
  const cargas = ult.trabalho.map((r) => r.carga).filter((c) => c !== null);
  const ref = Math.max(...cargas);
  const ativ = {}; ult.ativacao.forEach((r, i) => { ativ[i + 1] = r.carga; });
  const resumo = ult.trabalho.map((r) => r.reps).join(', ');
  const u = rotuloCarga(ex.equip) === 'Lastro' ? 'kg de lastro' : 'kg';
  if (sessao.deload) return { carga: ref, ativ, texto: `Deload: mesma carga (${fmtNum(ref)} ${u}), parar com 2–3 reps sobrando.`, sobe: false };
  const bateuTopo = zona && ult.trabalho.length && ult.trabalho.every((r) => r.reps !== null && r.reps >= zona[1]);
  if (bateuTopo) {
    const inc = incrementoDe(ex.equip, ref);
    return { carga: ref + inc, ativ, texto: `Última: ${fmtNum(ref)} ${u} × ${resumo}. Topo da zona → subir para ${fmtNum(ref + inc)} ${u} (+${fmtNum(inc)}).`, sobe: true };
  }
  return { carga: ref, ativ, texto: `Última: ${fmtNum(ref)} ${u} × ${resumo}. Mesma carga, buscar mais reps.`, sobe: false };
}

/* ---------- timer de descanso ---------- */
const Timer = {
  estado: null, _int: null, _audio: null,
  iniciar(seg, rotulo) {
    this.estado = { fim: Date.now() + seg * 1000, total: seg, rotulo, avisado: false };
    this._persistir(); this._ligar();
  },
  ajustar(d) { if (!this.estado) return; this.estado.fim += d * 1000; this.estado.total = Math.max(1, this.estado.total + d); this.estado.avisado = false; this._persistir(); this.tick(); },
  parar() { this.estado = null; this._persistir(); clearInterval(this._int); this._int = null; $('#timer').classList.add('oculto'); $('#timer').classList.remove('fim'); document.body.classList.remove('com-timer'); },
  restaurar() { try { const s = JSON.parse(localStorage.getItem('vf-timer') || 'null'); if (s && s.fim > Date.now() - 30000) { this.estado = s; this._ligar(); } } catch (e) { /* sem storage */ } },
  _persistir() { try { localStorage.setItem('vf-timer', JSON.stringify(this.estado)); } catch (e) { /* sem storage */ } },
  _ligar() { $('#timer').classList.remove('oculto'); document.body.classList.add('com-timer'); clearInterval(this._int); this._int = setInterval(() => this.tick(), 250); this.tick(); },
  tick() {
    if (!this.estado) return;
    const resta = (this.estado.fim - Date.now()) / 1000;
    $('#timer-rotulo').textContent = resta > 0 ? this.estado.rotulo : 'Hora da próxima série';
    $('#timer-tempo').textContent = fmtSeg(Math.ceil(Math.max(0, resta)));
    $('#timer-progresso').style.width = `${Math.min(100, 100 * (1 - Math.max(0, resta) / this.estado.total))}%`;
    $('#timer').classList.toggle('fim', resta <= 0);
    if (resta <= 0 && !this.estado.avisado) { this.estado.avisado = true; this._persistir(); this.alertar(); }
    if (resta < -30) this.parar();
  },
  desbloquearSom() { try { if (!this._audio) this._audio = new (window.AudioContext || window.webkitAudioContext)(); if (this._audio.state === 'suspended') this._audio.resume(); } catch (e) { /* sem áudio */ } },
  alertar() {
    try { navigator.vibrate && navigator.vibrate([400, 150, 400, 150, 400]); } catch (e) { /* */ }
    try {
      const ctx = this._audio; if (!ctx) return;
      [0, 0.35, 0.7].forEach((t) => {
        const o = ctx.createOscillator(); const g = ctx.createGain();
        o.frequency.value = 880; o.connect(g); g.connect(ctx.destination);
        g.gain.setValueAtTime(0.0001, ctx.currentTime + t);
        g.gain.exponentialRampToValueAtTime(0.4, ctx.currentTime + t + 0.02);
        g.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + t + 0.25);
        o.start(ctx.currentTime + t); o.stop(ctx.currentTime + t + 0.3);
      });
    } catch (e) { /* */ }
  }
};

/* ---------- manter a tela acesa durante o treino ---------- */
const Tela = {
  _lock: null, ativo: false,
  async ligar() { this.ativo = true; try { if ('wakeLock' in navigator && !this._lock) { this._lock = await navigator.wakeLock.request('screen'); this._lock.addEventListener('release', () => { this._lock = null; }); } } catch (e) { /* sem suporte */ } },
  desligar() { this.ativo = false; try { this._lock && this._lock.release(); } catch (e) { /* */ } this._lock = null; }
};
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') { if (Tela.ativo) Tela.ligar(); Timer.tick(); } });

/* ---------- telas ---------- */
const app = () => $('#app');
const rascunho = {}; // valores digitados ainda não salvos (id da série → campos)

function rota() {
  const h = decodeURIComponent(location.hash.slice(1)) || 'hoje';
  const [tela, ...resto] = h.split('/');
  const arg = resto.join('/');
  $$('#nav button').forEach((b) => b.classList.toggle('ativo', b.dataset.tela === (tela === 'treino' || tela === 'corrida' ? (arg && arg >= hojeISO().slice(0, 10) ? 'hoje' : 'semana') : tela)));
  if (tela !== 'treino') Tela.desligar();
  window.scrollTo(0, 0);
  if (tela === 'semana') return telaSemana(arg);
  if (tela === 'treino') return telaTreino(arg);
  if (tela === 'corrida') return telaCorrida(arg);
  if (tela === 'mais') return telaMais();
  return telaHoje();
}
const ir = (h) => { if (location.hash === '#' + h) rota(); else location.hash = h; };

function iconeStatus(st) { return st === 'ok' ? '<span class="status" style="color:var(--verde)">✓</span>' : st === 'andamento' ? '<span class="status" style="color:var(--laranja)">●</span>' : '<span class="status" style="color:var(--suave)">○</span>'; }

function htmlItemSessao(s) {
  const st = statusSessao(s.chave);
  const txt = st === 'ok' ? 'Concluído' : st === 'andamento' ? 'Em andamento' : `${s.itens.length} exercícios`;
  return `<div class="item-dia" data-ir="treino/${esc(s.chave)}"><div><b>💪 ${esc(s.sessao)}</b>${s.deload ? ' <span class="tag laranja">deload</span>' : ''}<div class="mini">${txt}</div></div>${iconeStatus(st)}</div>`;
}
function htmlItemCorrida(c) {
  const st = E.regCorrida[c.data] ? 'ok' : 'nao';
  const det = `${fmtNum(c.dist)} km · ${esc(c.pace)}${c.strides ? ` · ${c.strides} strides` : ''}`;
  return `<div class="item-dia" data-ir="corrida/${c.data}"><div><b>🏃 ${esc(c.tipo)}</b>${c.deload ? ' <span class="tag laranja">deload</span>' : ''}<div class="mini">${det}</div></div>${iconeStatus(st)}</div>`;
}

function avisoBackup() {
  const temDados = Object.keys(E.regMusc).length || Object.keys(E.regCorrida).length;
  if (!temDados) return '';
  const dias = E.ultimoBackup ? Math.floor((Date.now() - new Date(E.ultimoBackup).getTime()) / 864e5) : null;
  if (dias !== null && dias < 7) return '';
  return `<div class="card alerta clicavel" data-ir="mais"><b class="alerta-txt">Faça um backup</b><p class="suave">${dias === null ? 'Nenhum backup feito ainda.' : `Último backup há ${dias} dias.`} Seus registros só existem neste celular.</p></div>`;
}

function telaHoje() {
  const hoje = hojeISO();
  if (!E.plano) {
    app().innerHTML = `<h1>Vida Fitness</h1><div class="card"><h3>Nenhum plano importado</h3><p class="suave">Baixe a planilha em .xlsx e importe aqui.</p><button class="btn-primario btn-largo" data-ir="mais">Importar plano</button></div>`;
    return ligar();
  }
  const sem = semanasDoPlano(); const seg = segundaDe(hoje);
  const semAtual = sem.find((s) => s.segunda === seg);
  const ms = sessoesDoDia(hoje); const cs = corridaDoDia(hoje);
  let h = `<h1>Hoje</h1><p class="suave">${DIAS_LONGOS[dataDe(hoje).getDay()]}, ${ddmm(hoje)}${semAtual ? ` · <span class="tag verde">Semana ${semAtual.semana}</span>` : ''}</p>`;
  h += avisoBackup();
  if (ms.length || cs.length) {
    h += ms.map(htmlItemSessao).join('') + cs.map(htmlItemCorrida).join('');
  } else {
    h += `<div class="card"><b>${dataDe(hoje).getDay() === 0 ? 'Descanso total' : 'Sem treino hoje'}</b></div>`;
  }
  const primeiro = sem[0]?.segunda;
  if (primeiro && hoje < primeiro) h += `<p class="suave">O plano começa em ${ddmm(primeiro)}.</p>`;
  const ultimo = sem[sem.length - 1]?.segunda;
  if (ultimo && hoje > somaDias(ultimo, 6)) h += `<div class="card"><b>Plano encerrado</b><p class="suave">Importe o próximo ciclo em Mais.</p></div>`;
  // próximos dias com treino
  const prox = [];
  for (let i = 1; i <= 14 && prox.length < 3; i++) {
    const d = somaDias(hoje, i); const m = sessoesDoDia(d); const c = corridaDoDia(d);
    if (m.length || c.length) prox.push({ d, m, c });
  }
  if (prox.length) {
    h += `<h2>Próximos</h2>`;
    for (const p of prox) h += `<div class="dia"><div class="mini">${diaSem(p.d)} ${ddmm(p.d)}</div>${p.m.map(htmlItemSessao).join('')}${p.c.map(htmlItemCorrida).join('')}</div>`;
  }
  app().innerHTML = h; ligar();
}

function telaSemana(segunda) {
  if (!E.plano) return telaHoje();
  const sem = semanasDoPlano(); const hoje = hojeISO();
  if (!segunda || !sem.find((s) => s.segunda === segunda)) segunda = (sem.find((s) => s.segunda === segundaDe(hoje)) || sem.find((s) => s.segunda > hoje) || sem[sem.length - 1]).segunda;
  let h = `<h1>Semana</h1><div class="chips">${sem.map((s) => `<button class="chip ${s.segunda === segunda ? 'ativo' : ''}" data-ir="semana/${s.segunda}">S${s.semana} · ${ddmm(s.segunda)}</button>`).join('')}</div>`;
  for (let i = 0; i < 7; i++) {
    const d = somaDias(segunda, i); const m = sessoesDoDia(d); const c = corridaDoDia(d);
    h += `<div class="dia ${d === hoje ? 'hoje' : ''}"><div class="mini">${DIAS_LONGOS[dataDe(d).getDay()]} ${ddmm(d)}${d === hoje ? ' · hoje' : ''}</div>`;
    h += m.map(htmlItemSessao).join('') + c.map(htmlItemCorrida).join('');
    if (!m.length && !c.length) h += `<div class="mini" style="padding:6px 0">${i === 6 ? 'Descanso total' : '—'}</div>`;
    h += `</div>`;
  }
  app().innerHTML = h; ligar();
  const ativo = $('.chip.ativo'); ativo && ativo.scrollIntoView({ inline: 'center', block: 'nearest' });
}

/* ----- treino de musculação ----- */
function prescricao(it) {
  const total = it.ativ + it.trab;
  if (!it.ativ) return `${it.trab} × ${it.zona}`;
  return `${total} séries (${it.ativ} ativação × ${fmtNum(it.repsAtiv)} + ${it.trab} trabalho × ${it.zona})`;
}

function telaTreino(chave, manterScroll) {
  const s = sessaoPorChave(chave);
  if (!s) { app().innerHTML = `<p>Treino não encontrado no plano.</p><button data-ir="semana">Voltar</button>`; return ligar(); }
  const meta = metaSessao(chave);
  const ordem = ordemExecucao(s);
  const feitas = ordem.filter((x) => E.regMusc[x.id]);
  const atual = ordem.find((x) => !E.regMusc[x.id]);
  const trabTot = ordem.filter((x) => x.tipo === 'T').length;
  const trabFeitas = feitas.filter((x) => x.tipo === 'T').length;
  const sugs = {}; s.itens.forEach((it) => { sugs[it.ordem] = sugestao(s, it); });
  const concluida = meta?.status === 'concluida';

  let h = `<div class="cab-treino"><div class="linha entre"><button class="btn-peq" data-ir="semana/${segundaDe(s.data)}">‹ Voltar</button><span class="mini">S${s.semana} · ${diaSem(s.data)} ${ddmm(s.data)}</span></div>
    <h1>${esc(s.sessao)} ${s.deload ? '<span class="tag laranja">deload</span>' : ''}</h1>
    <div class="linha entre mini"><span>Séries feitas: ${feitas.length}/${ordem.length}</span><span>Trabalho: ${trabFeitas}/${trabTot}</span></div>
    <div class="progresso"><div style="width:${ordem.length ? (100 * feitas.length / ordem.length) : 0}%"></div></div></div>`;
  if (s.data !== hojeISO() && !meta) h += `<p class="mini">Planejado para ${diaSem(s.data)} ${ddmm(s.data)}. O registro fica com a data em que você fizer.</p>`;
  if (concluida) h += `<div class="card"><b style="color:var(--verde)">Treino concluído</b><p class="mini">Você ainda pode corrigir valores.</p></div>`;

  for (const b of blocosDe(s)) {
    const multi = b.itens.length > 1;
    h += `<div class="bloco">${multi ? `<div class="bloco-titulo">${esc(b.nome)} · alternar ${b.itens.length === 2 ? 'os dois' : 'os três'}</div>` : ''}`;
    for (const it of b.itens) {
      const idFeito = exFeito(chave, it);
      const ex = E.plano.exercicios[idFeito] || {};
      const trocado = idFeito !== it.id;
      const sg = sugs[it.ordem];
      const mex = meta?.ex?.[it.ordem] || {};
      const rc = rotuloCarga(ex.equip);
      h += `<div class="ex" id="ex-${it.ordem}">
        <div class="ex-nome">${esc(ex.nome || it.nome)}</div>
        ${trocado ? `<div class="mini">Trocado (plano: ${esc(it.nome)})</div>` : ''}
        <div class="ex-presc">${esc(ex.equip || '')} · ${esc(prescricao(it))}</div>
        ${it.obs ? `<div class="obs-plano">${esc(it.obs)}</div>` : ''}
        ${!it.conta ? '<div class="mini">Não conta no volume nesta semana.</div>' : ''}
        <div class="sugestao ${sg.sobe ? 'sobe' : ''}">${esc(sg.texto)}</div>
        <div class="series-cab"><span>Série</span><span>${rc}</span><span>Reps</span><span>Sobra</span><span></span></div>`;
      for (const se of seriesDe(s, it)) {
        const r = E.regMusc[se.id]; const rs = rascunho[se.id] || {};
        const carga = r ? r.carga : (rs.carga ?? (se.tipo === 'T' ? sg.carga : sg.ativ[se.idx]) ?? '');
        const reps = r ? r.reps : (rs.reps ?? '');
        const sobra = r ? r.sobrando : (rs.sobrando ?? '');
        const cls = ['serie', se.tipo === 'A' ? 'ativacao' : '', r ? 'feita' : '', atual && atual.id === se.id ? 'atual' : ''].join(' ');
        h += `<div class="${cls}" id="s-${esc(se.id)}" data-serie="${esc(se.id)}">
          <div class="rot"><b>${se.tipo === 'A' ? 'Ativ' : 'Trab'} ${se.idx}</b>${se.tipo === 'A' ? `${fmtNum(se.alvo)} reps` : esc(se.alvo)}</div>
          <input type="text" inputmode="decimal" data-campo="carga" value="${esc(fmtNum(carga))}" placeholder="kg">
          <input type="text" inputmode="numeric" data-campo="reps" value="${esc(reps ?? '')}" placeholder="${se.tipo === 'A' ? esc(fmtNum(se.alvo)) : esc((zonaDe(se.alvo) || [''])[1])}">
          ${se.tipo === 'T' ? `<select data-campo="sobrando"><option value=""></option>${[0, 1, 2, 3, 4].map((v) => `<option value="${v}" ${String(sobra) === String(v) ? 'selected' : ''}>${v === 4 ? '4+' : v}</option>`).join('')}</select>` : '<span class="mini" style="text-align:center">4–5</span>'}
          <button class="ok" data-acao="ok">${r ? '✓' : '○'}</button>
        </div>`;
      }
      const inc = mex.incomodo ?? null;
      h += `<details class="ex-extra" ${(mex.obs || inc || trocado) ? 'open' : ''}><summary class="mini">Incômodo, observação, trocar exercício</summary>
          <label>Incômodo (0–10)</label>
          <div class="notas" data-ordem="${it.ordem}">${[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map((v) => `<button data-acao="incomodo" data-v="${v}" class="${inc === v ? 'sel' : ''} ${v >= 4 ? 'alto' : ''}">${v}</button>`).join('')}</div>
          <label>Observação</label>
          <textarea rows="2" data-acao="obs" data-ordem="${it.ordem}" placeholder="ex.: aparelho ocupado, dor, ajuste de banco">${esc(mex.obs || '')}</textarea>
          <label>Trocar por (aparelho ocupado etc.)</label>
          <select data-acao="trocar" data-ordem="${it.ordem}">${Object.values(E.plano.exercicios).map((e) => `<option value="${esc(e.id)}" ${e.id === idFeito ? 'selected' : ''}>${esc(e.id)} · ${esc(e.nome)}${e.id === it.id ? ' (plano)' : ''}</option>`).join('')}</select>
        </details></div>`;
    }
    h += `</div>`;
  }
  h += concluida
    ? `<button class="btn-largo" data-acao="reabrir">Reabrir treino</button>`
    : `<button class="btn-primario btn-largo" data-acao="concluir">Concluir treino</button>`;
  h += `<p class="mini" style="margin-top:12px">Ativação não conta no volume. Trabalho: 0–1 rep sobrando.</p>`;

  const y = window.scrollY;
  app().innerHTML = h; ligar();
  if (manterScroll) window.scrollTo(0, y);
  Tela.ligar();
  ligarTreino(s);
}

function lerSerie(el) {
  const g = (c) => { const x = el.querySelector(`[data-campo="${c}"]`); return x ? x.value : ''; };
  return { carga: lerNum(g('carga')), reps: lerNum(g('reps')), sobrando: g('sobrando') === '' ? null : Number(g('sobrando')) };
}

function garantirMeta(s) {
  if (!E.sessoes[s.chave]) E.sessoes[s.chave] = { status: 'andamento', inicio: new Date().toISOString(), dataReal: hojeISO(), deload: s.deload, ex: {} };
  E.sessoes[s.chave].ex ||= {};
  return E.sessoes[s.chave];
}

function ligarTreino(s) {
  const ordem = ordemExecucao(s);
  const porId = Object.fromEntries(ordem.map((x) => [x.id, x]));

  $$('[data-serie]').forEach((el) => {
    const id = el.dataset.serie; const se = porId[id];
    el.addEventListener('input', () => { rascunho[id] = lerSerie(el); });
    el.addEventListener('change', async () => {
      if (!E.regMusc[id]) return; // só salva sozinho se a série já foi marcada
      Object.assign(E.regMusc[id], lerSerie(el), { editadoEm: new Date().toISOString() });
      await salvar('regMusc'); toast('Correção salva');
    });
    el.querySelector('[data-acao="ok"]').addEventListener('click', async () => {
      Timer.desbloquearSom();
      if (E.regMusc[id]) {
        if (!confirm('Desmarcar esta série? O registro dela será apagado.')) return;
        delete E.regMusc[id]; await salvar('regMusc'); return telaTreino(s.chave, true);
      }
      const v = lerSerie(el);
      if (v.reps === null) { toast('Preencha as reps'); el.querySelector('[data-campo="reps"]').focus(); return; }
      const meta = garantirMeta(s);
      if (meta.status === 'concluida') meta.status = 'andamento';
      E.regMusc[id] = {
        id, sessao: s.chave, dataPlano: s.data, semana: s.semana, sessaoNome: s.sessao, ordem: se.it.ordem,
        exPlano: se.it.id, exFeito: exFeito(s.chave, se.it), n: se.n, tipo: se.tipo, ...v, ts: new Date().toISOString()
      };
      delete rascunho[id];
      await Promise.all([salvar('regMusc'), salvar('sessoes')]);
      const prox = ordem.find((x) => !E.regMusc[x.id]);
      if (prox) {
        const d = E.config.descanso;
        const seg = se.tipo === 'A' ? d.ativacao : (se.bloco.itens.length > 1 ? d.trabalhoBloco : d.trabalhoSimples);
        const nomeProx = (E.plano.exercicios[exFeito(s.chave, prox.it)] || prox.it).nome;
        Timer.iniciar(seg, `Próxima: ${nomeProx} · ${prox.tipo === 'A' ? 'Ativ' : 'Trab'} ${prox.idx}`);
      } else { Timer.parar(); toast('Todas as séries feitas. Toque em Concluir treino.', 4000); }
      telaTreino(s.chave, true);
      const alvo = prox && document.getElementById('s-' + prox.id);
      if (alvo) alvo.scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
  });

  $$('.notas').forEach((box) => box.addEventListener('click', async (ev) => {
    const b = ev.target.closest('[data-acao="incomodo"]'); if (!b) return;
    const meta = garantirMeta(s); const o = Number(box.dataset.ordem);
    meta.ex[o] = { ...(meta.ex[o] || {}), incomodo: Number(b.dataset.v) };
    await salvar('sessoes');
    $$('button', box).forEach((x) => x.classList.toggle('sel', x === b));
    if (Number(b.dataset.v) >= 4) toast('Incômodo ≥4: anote o que sentiu e avise no projeto.', 4000);
  }));
  $$('textarea[data-acao="obs"]').forEach((t) => t.addEventListener('change', async () => {
    const meta = garantirMeta(s); const o = Number(t.dataset.ordem);
    meta.ex[o] = { ...(meta.ex[o] || {}), obs: t.value.trim() };
    await salvar('sessoes'); toast('Observação salva');
  }));
  $$('select[data-acao="trocar"]').forEach((sel) => sel.addEventListener('change', async () => {
    const meta = garantirMeta(s); const o = Number(sel.dataset.ordem);
    meta.ex[o] = { ...(meta.ex[o] || {}), exFeito: sel.value };
    for (const r of Object.values(E.regMusc)) if (r.sessao === s.chave && r.ordem === o) r.exFeito = sel.value;
    await Promise.all([salvar('sessoes'), salvar('regMusc')]);
    toast('Exercício trocado'); telaTreino(s.chave, true);
  }));
  const conc = $('[data-acao="concluir"]');
  conc && conc.addEventListener('click', async () => {
    const faltam = ordem.filter((x) => !E.regMusc[x.id]).length;
    if (!E.sessoes[s.chave] && !confirm('Nenhuma série registrada. Concluir mesmo assim?')) return;
    if (faltam && !confirm(`Faltam ${faltam} séries. Concluir assim mesmo?`)) return;
    const meta = garantirMeta(s); meta.status = 'concluida'; meta.fim = new Date().toISOString();
    await salvar('sessoes'); Timer.parar(); Tela.desligar(); toast('Treino concluído 💪'); ir('hoje');
  });
  const reab = $('[data-acao="reabrir"]');
  reab && reab.addEventListener('click', async () => { E.sessoes[s.chave].status = 'andamento'; await salvar('sessoes'); telaTreino(s.chave); });
}

/* ----- corrida ----- */
function telaCorrida(data) {
  const c = corridaDoDia(data)[0];
  if (!c) { app().innerHTML = `<p>Corrida não encontrada no plano.</p><button data-ir="semana">Voltar</button>`; return ligar(); }
  const r = E.regCorrida[data] || {};
  const h = `<div class="linha entre"><button class="btn-peq" data-ir="semana/${segundaDe(data)}">‹ Voltar</button><span class="mini">S${c.semana} · ${diaSem(data)} ${ddmm(data)}</span></div>
    <h1>🏃 ${esc(c.tipo)} ${c.deload ? '<span class="tag laranja">deload</span>' : ''}</h1>
    <div class="card"><div class="grade2"><div><div class="mini">Distância</div><b>${fmtNum(c.dist)} km</b></div><div><div class="mini">Pace alvo</div><b>${esc(c.pace)}</b></div></div>
      ${c.strides ? `<p><b>${c.strides} strides</b> no fim</p>` : ''}${c.obs ? `<p class="obs-plano">${esc(c.obs)}</p>` : ''}</div>
    <h2>Registro ${E.regCorrida[data] ? '<span class="tag verde">salvo</span>' : ''}</h2>
    <div class="grade2">
      <div><label>Distância (km)</label><input id="c-dist" inputmode="decimal" value="${esc(fmtNum(r.dist ?? c.dist))}"></div>
      <div><label>Tempo (mm:ss ou h:mm:ss)</label><input id="c-tempo" inputmode="numeric" placeholder="35:20" value="${esc(fmtTempo(r.tempo))}"></div>
      <div><label>FC média</label><input id="c-fcm" inputmode="numeric" value="${esc(r.fcMed ?? '')}"></div>
      <div><label>FC máx</label><input id="c-fcx" inputmode="numeric" value="${esc(r.fcMax ?? '')}"></div>
    </div>
    <p class="suave">Pace médio: <b id="c-pace">${r.pace ? r.pace + '/km' : '—'}</b></p>
    <label>Teste da fala (conseguiu conversar?)</label>
    <div class="toggle" id="c-fala"><button data-v="sim" class="${r.fala === 'Sim' ? 'sel' : ''}">Sim</button><button data-v="nao" class="${r.fala === 'Não' ? 'sel' : ''}">Não</button></div>
    <label>Dor lateral no joelho (0–10)</label>
    <div class="notas" id="c-dor">${[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map((v) => `<button data-v="${v}" class="${r.dor === v ? 'sel' : ''} ${v >= 4 ? 'alto' : ''}">${v}</button>`).join('')}</div>
    <label>Observação</label><textarea id="c-obs" rows="2">${esc(r.obs || '')}</textarea>
    <div style="height:12px"></div>
    <button class="btn-primario btn-largo" id="c-salvar">Salvar corrida</button>
    ${E.regCorrida[data] ? '<div style="height:8px"></div><button class="btn-perigo btn-largo" id="c-apagar">Apagar registro</button>' : ''}`;
  app().innerHTML = h; ligar();
  let fala = r.fala || null; let dor = r.dor ?? null;
  const calcPace = () => {
    const d = lerNum($('#c-dist').value); const t = lerTempo($('#c-tempo').value);
    const p = d && t ? fmtSeg(t / d) : null; $('#c-pace').textContent = p ? p + '/km' : '—'; return p;
  };
  $('#c-dist').addEventListener('input', calcPace); $('#c-tempo').addEventListener('input', calcPace);
  $('#c-fala').addEventListener('click', (e) => { const b = e.target.closest('button'); if (!b) return; fala = b.dataset.v === 'sim' ? 'Sim' : 'Não'; $$('#c-fala button').forEach((x) => x.classList.toggle('sel', x === b)); });
  $('#c-dor').addEventListener('click', (e) => { const b = e.target.closest('button'); if (!b) return; dor = Number(b.dataset.v); $$('#c-dor button').forEach((x) => x.classList.toggle('sel', x === b)); });
  $('#c-salvar').addEventListener('click', async () => {
    const dist = lerNum($('#c-dist').value); const tempo = lerTempo($('#c-tempo').value);
    if (!dist || !tempo) { toast('Preencha distância e tempo'); return; }
    E.regCorrida[data] = {
      dataPlano: data, data: r.data || hojeISO(), semana: c.semana, tipo: c.tipo, dist, tempo, pace: calcPace(),
      fcMed: lerNum($('#c-fcm').value), fcMax: lerNum($('#c-fcx').value), fala, dor, obs: $('#c-obs').value.trim(), ts: new Date().toISOString()
    };
    await salvar('regCorrida');
    if (dor !== null && dor >= 4) alert('Dor lateral no joelho ≥4: reduza o volume e avise no projeto.');
    else toast('Corrida salva');
    ir('semana/' + segundaDe(data));
  });
  const ap = $('#c-apagar');
  ap && ap.addEventListener('click', async () => { if (!confirm('Apagar o registro desta corrida?')) return; delete E.regCorrida[data]; await salvar('regCorrida'); telaCorrida(data); });
}

/* ----- mais: importar, backup, exportar, configurações ----- */
function telaMais() {
  const p = E.plano;
  const sem = semanasDoPlano();
  const resumo = p ? `<p class="suave">Importado em ${new Date(p.importadoEm).toLocaleString('pt-BR')}</p>
      <p class="mini">${Object.keys(p.exercicios).length} exercícios · ${sessoesDoPlano().length} treinos de musculação · ${p.corrida.length} corridas · ${sem.length ? `${ddmm(sem[0].segunda)} a ${ddmm(somaDias(sem[sem.length - 1].segunda, 6))}` : ''}</p>` : '<p class="suave">Nenhum plano importado.</p>';
  const nM = Object.keys(E.regMusc).length; const nC = Object.keys(E.regCorrida).length;
  const inc = E.config.incrementos; const d = E.config.descanso;
  app().innerHTML = `<h1>Mais</h1>
    <div class="card"><h3>Plano</h3>${resumo}
      <label for="f-plano">Arquivo da planilha (.xlsx)</label><input type="file" id="f-plano" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">
      <p class="mini">Reimportar troca só o plano. Seus registros ficam.</p></div>
    <div class="card"><h3>Backup</h3>
      <p class="suave">${nM} séries e ${nC} corridas registradas. ${E.ultimoBackup ? `Último backup: ${new Date(E.ultimoBackup).toLocaleDateString('pt-BR')}.` : 'Nenhum backup ainda.'}</p>
      <button class="btn-primario btn-largo" id="b-backup">Fazer backup</button>
      <label for="f-restaurar">Restaurar backup (.json)</label><input type="file" id="f-restaurar" accept=".json,application/json">
      <p class="mini">Guarde o backup fora do celular (Drive, email). Restaurar substitui tudo o que está no app.</p>
      <div style="height:8px"></div><button class="btn-perigo btn-largo" id="b-zerar">Apagar todos os registros</button>
      <p class="mini">Use para limpar treinos de teste. O plano fica.</p></div>
    <div class="card"><h3>Exportar registros (CSV)</h3>
      <div class="grade2"><button id="b-csv-m">Musculação</button><button id="b-csv-c">Corrida</button></div>
      <p class="mini">Mesmas colunas das abas Registro_* da planilha.</p></div>
    <div class="card"><h3>Configurações</h3>
      <p class="mini">Quanto a carga sobe quando você bate o topo da zona.</p>
      <div class="grade2">${Object.keys(CONFIG_PADRAO.incrementos).map((k) => `<div><label>${k} (+kg)</label><input inputmode="decimal" data-inc="${k}" value="${fmtNum(inc[k])}"></div>`).join('')}</div>
      <p class="mini" style="margin-top:12px">Descanso automático (segundos).</p>
      <div class="grade2">
        <div><label>Após ativação</label><input inputmode="numeric" data-desc="ativacao" value="${d.ativacao}"></div>
        <div><label>Trabalho em bi/tri-set</label><input inputmode="numeric" data-desc="trabalhoBloco" value="${d.trabalhoBloco}"></div>
        <div><label>Trabalho em exercício simples</label><input inputmode="numeric" data-desc="trabalhoSimples" value="${d.trabalhoSimples}"></div>
      </div>
      <div style="height:10px"></div><button class="btn-largo" id="b-config">Salvar configurações</button></div>
    <p class="mini">Vida Fitness ${VERSAO_APP} · dados guardados só neste celular.</p>`;
  ligar();

  $('#f-plano').addEventListener('change', async (e) => {
    const f = e.target.files[0]; if (!f) return;
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
      const novo = interpretarPlanilha(wb); novo.arquivo = f.name;
      const nS = new Set(novo.musc.map((r) => `${r.data}_${r.sessao}`)).size;
      if (E.plano && !confirm(`Substituir o plano atual?\n\nNovo: ${Object.keys(novo.exercicios).length} exercícios, ${nS} treinos de musculação, ${novo.corrida.length} corridas.\n\nSeus registros continuam.`)) { e.target.value = ''; return; }
      E.plano = novo; await salvar('plano');
      try { navigator.storage && navigator.storage.persist && await navigator.storage.persist(); } catch (x) { /* */ }
      toast(`Plano importado: ${nS} treinos, ${novo.corrida.length} corridas`, 3500);
      telaMais();
    } catch (err) { alert('Não consegui importar:\n\n' + err.message); e.target.value = ''; }
  });
  $('#b-backup').addEventListener('click', fazerBackup);
  $('#b-zerar').addEventListener('click', async () => {
    if (!confirm('Apagar TODAS as séries e corridas registradas? O plano continua. Não dá para desfazer.')) return;
    if (!confirm('Confirma mesmo? Faça um backup antes se tiver dados reais.')) return;
    E.regMusc = {}; E.sessoes = {}; E.regCorrida = {};
    await Promise.all(['regMusc', 'sessoes', 'regCorrida'].map(salvar));
    Timer.parar(); toast('Registros apagados'); telaMais();
  });
  $('#f-restaurar').addEventListener('change', async (e) => {
    const f = e.target.files[0]; if (!f) return;
    try {
      const b = JSON.parse(await f.text());
      if (b.app !== 'vida-fitness') throw new Error('Este arquivo não é um backup do Vida Fitness.');
      const nm = Object.keys(b.regMusc || {}).length; const nc = Object.keys(b.regCorrida || {}).length;
      if (!confirm(`Restaurar backup de ${new Date(b.exportadoEm).toLocaleString('pt-BR')}?\n\n${nm} séries e ${nc} corridas.\n\nIsso substitui tudo o que está no app agora.`)) { e.target.value = ''; return; }
      E.plano = b.plano || null; E.config = Object.assign(structuredClone(CONFIG_PADRAO), b.config || {});
      E.regMusc = b.regMusc || {}; E.sessoes = b.sessoes || {}; E.regCorrida = b.regCorrida || {};
      await Promise.all(['plano', 'config', 'regMusc', 'sessoes', 'regCorrida'].map(salvar));
      toast('Backup restaurado'); telaMais();
    } catch (err) { alert('Não consegui restaurar:\n\n' + err.message); e.target.value = ''; }
  });
  $('#b-csv-m').addEventListener('click', () => baixar(`registro-musculacao-${hojeISO()}.csv`, csvMusculacao(), 'text/csv'));
  $('#b-csv-c').addEventListener('click', () => baixar(`registro-corrida-${hojeISO()}.csv`, csvCorrida(), 'text/csv'));
  $('#b-config').addEventListener('click', async () => {
    $$('[data-inc]').forEach((i) => { const v = lerNum(i.value); if (v !== null && v >= 0) E.config.incrementos[i.dataset.inc] = v; });
    $$('[data-desc]').forEach((i) => { const v = lerNum(i.value); if (v !== null && v > 0) E.config.descanso[i.dataset.desc] = Math.round(v); });
    await salvar('config'); toast('Configurações salvas'); telaMais();
  });
}

async function baixar(nome, conteudo, tipo) {
  const blob = new Blob([conteudo], { type: tipo });
  const arquivo = new File([blob], nome, { type: tipo });
  try {
    if (navigator.canShare && navigator.canShare({ files: [arquivo] })) { await navigator.share({ files: [arquivo], title: nome }); return true; }
  } catch (e) { if (e.name === 'AbortError') return false; }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nome;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  return true;
}
async function fazerBackup() {
  const b = { app: 'vida-fitness', versao: 1, versaoApp: VERSAO_APP, exportadoEm: new Date().toISOString(), plano: E.plano, config: E.config, regMusc: E.regMusc, sessoes: E.sessoes, regCorrida: E.regCorrida };
  const ok = await baixar(`vida-fitness-backup-${hojeISO()}.json`, JSON.stringify(b), 'application/json');
  if (ok) { E.ultimoBackup = new Date().toISOString(); await salvar('ultimoBackup'); toast('Backup gerado'); if (location.hash === '#mais') telaMais(); }
}

const csvCel = (v) => { if (v === null || v === undefined) return ''; const s = String(v); return /[",\n;]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
const csv = (linhas) => '﻿' + linhas.map((l) => l.map(csvCel).join(',')).join('\r\n');
function csvMusculacao() {
  const cab = ['Data', 'Semana', 'Sessão', 'ID', 'Exercício', 'Série nº', 'Tipo de série', 'Carga (kg)', 'Reps', 'Reps sobrando', 'Observação', 'Incômodo (0–10)'];
  const regs = Object.values(E.regMusc).sort((a, b) => (a.dataPlano + a.sessao).localeCompare(b.dataPlano + b.sessao) || a.ordem - b.ordem || a.n - b.n);
  const linhas = regs.map((r) => {
    const m = E.sessoes[r.sessao] || {}; const mx = m.ex?.[r.ordem] || {};
    const ex = E.plano?.exercicios[r.exFeito];
    const obs = [r.exFeito !== r.exPlano ? `Trocado de ${r.exPlano}` : '', mx.obs || ''].filter(Boolean).join(' · ');
    return [m.dataReal || r.dataPlano, r.semana, r.sessaoNome, r.exFeito, ex ? ex.nome : '', r.n, r.tipo === 'A' ? 'Ativação' : 'Trabalho', r.carga, r.reps, r.sobrando, obs, mx.incomodo ?? ''];
  });
  return csv([cab, ...linhas]);
}
function csvCorrida() {
  const cab = ['Data', 'Semana', 'Tipo', 'Distância (km)', 'Tempo', 'Pace médio (/km)', 'FC média', 'FC máx', 'Teste da fala ok', 'Dor lateral joelho (0–10)', 'Observação'];
  const regs = Object.values(E.regCorrida).sort((a, b) => a.dataPlano.localeCompare(b.dataPlano));
  return csv([cab, ...regs.map((r) => [r.data, r.semana, r.tipo, r.dist, fmtTempo(r.tempo), r.pace, r.fcMed, r.fcMax, r.fala, r.dor, r.obs])]);
}

/* ---------- ligação geral ---------- */
function ligar() {
  $$('[data-ir]', app()).forEach((el) => el.addEventListener('click', () => ir(el.dataset.ir)));
}

async function iniciar() {
  try { await carregar(); } catch (e) { app().innerHTML = `<p>Erro ao abrir o banco de dados do app: ${esc(e.message)}</p>`; return; }
  $$('#nav button').forEach((b) => b.addEventListener('click', () => ir(b.dataset.tela)));
  $$('[data-timer]').forEach((b) => b.addEventListener('click', () => { const v = b.dataset.timer; if (v === 'pular') Timer.parar(); else Timer.ajustar(Number(v)); }));
  document.addEventListener('click', () => Timer.desbloquearSom(), { once: true });
  window.addEventListener('hashchange', rota);
  Timer.restaurar();
  rota();
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    try {
      const reg = await navigator.serviceWorker.register('sw.js');
      const avisar = (w) => { $('#atualizacao').classList.remove('oculto'); $('#btn-atualizar').onclick = () => w.postMessage('atualizar'); };
      if (reg.waiting && navigator.serviceWorker.controller) avisar(reg.waiting);
      reg.addEventListener('updatefound', () => {
        const w = reg.installing;
        w && w.addEventListener('statechange', () => { if (w.state === 'installed' && navigator.serviceWorker.controller) avisar(w); });
      });
      let recarregou = false;
      navigator.serviceWorker.addEventListener('controllerchange', () => { if (!recarregou) { recarregou = true; location.reload(); } });
    } catch (e) { /* sem service worker */ }
  }
}
iniciar();
