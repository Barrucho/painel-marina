
// ============================================================ FASE 4.2 — INTELIGÊNCIA (aprendizado + previsão + categorias)
// APRENDIZADO: cada correção da Marina vira uma "assinatura" (CNPJ/CPF ou nome limpo do estabelecimento) que é
// aplicada na hora a todos os lançamentos semelhantes — passados e futuros. Lançamentos sem regra recebem uma
// categoria PREVISTA a partir do que já foi aprendido (votação por palavras, pesada pela raridade da palavra).
// CATEGORIAS_MAPA: renomear/juntar categorias vale também para o que vem classificado da API.

Object.assign(ABAS, {
  APRENDIZADO: ['id', 'ativa', 'assinatura', 'direcao', 'natureza', 'categoria', 'subcategoria', 'tipo_receita',
    'exemplo', 'criado_em', 'criado_por'],
  CATEGORIAS_MAPA: ['de_categoria', 'de_subcategoria', 'para_categoria', 'para_subcategoria', 'criado_em']
});
['assinatura', 'criado_em'].forEach(function (c) { if (COLUNAS_TEXTO.indexOf(c) < 0) COLUNAS_TEXTO.push(c); });

const NAT_PROTEGIDAS = ['IGNORADA', 'PAGAMENTO_FATURA', 'PAGAMENTO_FATURA_CARTAO', 'AJUSTE_CARTAO'];
const PALAVRAS_GENERICAS = ('PIX ENVIADO RECEBIDO TRANSF TRANSFERENCIA TED DOC PAG PAGTO PAGAMENTO PAGTO COMPRA COMPRAS ' +
  'DEBITO CREDITO BOLETO CARTAO CART PARC PARCELA LTDA EIRELI EPP SAO PAULO BRA BRASIL COM WWW HTTP SERVICOS SERV ' +
  'DEB AUT AUTOMATICO VISA MASTER ELO INT NACIONAL APLICACAO QRS QRCODE CONTA CORRENTE').split(' ');

function garantirAba_(ss, nome) {
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, ABAS[nome].length).setValues([ABAS[nome]]).setFontWeight('bold')
      .setBackground('#1f2937').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function limparTexto_(s) {
  return norm_(s).replace(/[^A-Z ]/g, ' ').split(' ').filter(function (t) {
    return t.length >= 2 && PALAVRAS_GENERICAS.indexOf(t) < 0;
  });
}

// mesma regra existe no painel (JS) — manter as duas iguais
function assinatura_(doc, contraparte, descricao) {
  const d = String(doc || '').replace(/\D/g, '');
  const cp = norm_(contraparte);
  const gateway = GATEWAYS.some(function (g) { return cp.indexOf(g.trim()) >= 0; });
  if (d.length >= 11 && !gateway) return 'DOC:' + d;
  let t = limparTexto_(contraparte);
  if (gateway || !t.length) t = limparTexto_(descricao).concat(gateway ? t : []);
  if (!t.length) return 'TXT:' + norm_(descricao).replace(/\d/g, '').replace(/\s+/g, ' ').trim().substring(0, 30);
  return 'TXT:' + t.slice(0, 4).join(' ');
}

function tokens_(contraparte, descricao) {
  const vistos = {};
  return limparTexto_(String(contraparte || '') + ' ' + String(descricao || '')).filter(function (t) {
    if (t.length < 3 || vistos[t]) return false;
    vistos[t] = 1; return true;
  });
}

function lentes_(row, H) {
  const banco = row[H.conta_tipo] !== 'CARTAO';
  const nat = row[H.natureza];
  row[H.entra_caixa] = banco && row[H.status] === 'POSTED' && NAT_CAIXA.indexOf(nat) >= 0 ? 'SIM' : 'NAO';
  row[H.entra_consumo] = NAT_CONSUMO.indexOf(nat) >= 0 && (banco ? row[H.status] === 'POSTED' : true) ? 'SIM' : 'NAO';
}

// rows: linhas no formato de MOVIMENTOS (alteradas no lugar). Retorna estatísticas.
function aplicarInteligencia_(ss, rows) {
  garantirAba_(ss, 'APRENDIZADO'); garantirAba_(ss, 'CATEGORIAS_MAPA');
  const H = indice_(ABAS.MOVIMENTOS);
  const st = { renomeados: 0, aprendidos: 0, previstos: 0 };

  // 0. renomear/juntar categorias
  const mapa = lerTabela_(ss, 'CATEGORIAS_MAPA').filter(function (r) { return r[0]; });
  if (mapa.length) rows.forEach(function (row) {
    for (let p = 0; p < 5; p++) {
      let mudou = false;
      mapa.forEach(function (m) {
        if (row[H.categoria] !== m[0]) return;
        if (m[1] && row[H.subcategoria] !== m[1]) return;
        row[H.categoria] = m[2] || row[H.categoria];
        if (m[1] && m[3]) row[H.subcategoria] = m[3];
        mudou = true; st.renomeados++;
      });
      if (!mudou) break;
    }
  });

  // 1. aprendizado por assinatura
  const apr = {};
  lerTabela_(ss, 'APRENDIZADO').forEach(function (r) {
    const a = String(r[1]).trim().toUpperCase();
    if (!r[2] || a === 'FALSE' || a === 'NAO' || a === 'FALSO') return;
    apr[r[2] + '|' + r[3]] = r;
  });
  rows.forEach(function (row) {
    if (row[H.origem] === 'MANUAL') return;
    const nat = row[H.natureza];
    if (NAT_PROTEGIDAS.indexOf(nat) >= 0 || (nat === 'TRANSFERENCIA_INTERNA' && row[H.par_id])) return;
    const dir = Number(row[H.valor]) > 0 ? 'ENTRADA' : 'SAIDA';
    const a = apr[assinatura_(row[H.contraparte_doc], row[H.contraparte], row[H.descricao]) + '|' + dir];
    if (!a) return;
    const n = natureza_(a[4]) || nat;
    row[H.natureza] = n;
    if (n === 'RECEITA') {
      row[H.categoria] = 'Receitas'; row[H.tipo_receita] = a[7] || row[H.tipo_receita]; row[H.subcategoria] = row[H.tipo_receita];
    } else {
      row[H.categoria] = a[5] || row[H.categoria]; row[H.subcategoria] = a[6] || ''; row[H.tipo_receita] = '';
    }
    row[H.origem] = 'APRENDIDO'; row[H.revisar] = ''; row[H.motivo_revisao] = '';
    lentes_(row, H);
    st.aprendidos++;
  });

  // 2. previsão (só despesas sem regra confiável)
  const confiaveis = { MANUAL: 1, APRENDIDO: 1, REGRA: 1 };
  const voto = {}, df = {}, exemplos = {};
  let nTreino = 0;
  rows.forEach(function (row) {
    if (row[H.natureza] !== 'DESPESA' || !confiaveis[row[H.origem]] || !row[H.categoria]) return;
    const rot = row[H.categoria] + '\u0001' + (row[H.subcategoria] || '');
    const sig = assinatura_(row[H.contraparte_doc], row[H.contraparte], row[H.descricao]);
    nTreino++;
    tokens_(row[H.contraparte], row[H.descricao]).forEach(function (t) {
      df[t] = (df[t] || 0) + 1;
      const v = voto[t] || (voto[t] = {});
      v[rot] = (v[rot] || 0) + 1;
      const e = exemplos[t + '\u0002' + rot] || (exemplos[t + '\u0002' + rot] = {});
      e[sig] = 1;
    });
  });
  if (nTreino >= 5) rows.forEach(function (row) {
    if (row[H.natureza] !== 'DESPESA') return;
    const orig = row[H.origem];
    if (orig !== 'PADRAO' && orig !== 'API' && orig !== 'PREVISTO') return;
    const placar = {}, apoio = {};
    let total = 0;
    tokens_(row[H.contraparte], row[H.descricao]).forEach(function (t) {
      const v = voto[t];
      if (!v) return;
      const idf = Math.log(1 + nTreino / df[t]);
      const soma = Object.keys(v).reduce(function (s, k) { return s + v[k]; }, 0);
      Object.keys(v).forEach(function (rot) {
        const w = idf * v[rot] / soma;
        placar[rot] = (placar[rot] || 0) + w; total += w;
        const e = exemplos[t + '\u0002' + rot] || {};
        apoio[rot] = Math.max(apoio[rot] || 0, Object.keys(e).length);
      });
    });
    if (!total) return;
    const melhor = Object.keys(placar).sort(function (a, b) { return placar[b] - placar[a]; })[0];
    const conf = placar[melhor] / total;
    const p = melhor.split('\u0001');
    if (p[0] === row[H.categoria] && (p[1] || '') === (row[H.subcategoria] || '')) return;
    const limite = orig === 'API' ? 0.85 : 0.7;
    if (conf < limite || placar[melhor] < 0.5 || (orig === 'API' && apoio[melhor] < 2)) return;
    row[H.categoria] = p[0]; row[H.subcategoria] = p[1] || '';
    row[H.origem] = 'PREVISTO';
    const pc = Math.round(conf * 100);
    row[H.motivo_revisao] = 'Categoria prevista pela plataforma (' + pc + '% de confiança)';
    row[H.revisar] = conf < 0.85 ? 'SIM' : '';
    st.previstos++;
  });
  return st;
}

// ---------- dados para o painel (inclui origem e o que foi aprendido)
function dadosPainel2_() {
  const ss = planilhaPainel_();
  garantirAba_(ss, 'APRENDIZADO'); garantirAba_(ss, 'CATEGORIAS_MAPA');
  const H = indice_(ABAS.MOVIMENTOS);
  const mov = lerTabela_(ss, 'MOVIMENTOS').filter(function (r) { return r[H.id]; }).map(function (r) {
    return [r[H.id], r[H.data], r[H.mes_caixa], r[H.mes_consumo], r[H.banco], r[H.conta_tipo], r[H.conta_nome],
      r[H.descricao], r[H.contraparte], Number(r[H.valor]) || 0, r[H.natureza], r[H.entra_caixa] === 'SIM' ? 1 : 0,
      r[H.entra_consumo] === 'SIM' ? 1 : 0, r[H.categoria], r[H.subcategoria], r[H.tipo_receita],
      r[H.revisar] === 'SIM' ? 1 : 0, r[H.motivo_revisao], r[H.parcela], r[H.status], r[H.contraparte_doc],
      r[H.origem], r[H.par_id]];
  });
  const log = lerTabela_(ss, 'SYNC_LOG');
  const ult = log.length ? log[log.length - 1] : [];
  return JSON.stringify({
    v: 2,
    mov: mov,
    contas: lerTabela_(ss, 'CONTAS').map(function (r) { return r.slice(0, 11); }),
    invest: lerTabela_(ss, 'INVESTIMENTOS').map(function (r) { return r.slice(1, 13); }),
    pag: lerTabela_(ss, 'PAGADORES'),
    cats: lerTabela_(ss, 'CATEGORIAS').map(function (r) { return [r[0], r[1]]; }),
    apr: lerTabela_(ss, 'APRENDIZADO').filter(function (r) { return r[0]; }),
    patr: lerTabela_(ss, 'PATRIMONIO_HIST'),
    sync: { fim: ult[1] || '', status: ult[3] || '' },
    hojeMes: fmt_(new Date(), 'yyyy-MM'), hoje: fmt_(new Date(), 'yyyy-MM-dd'),
    tiposReceita: ['Serviços médicos (PJ/clínicas)', 'Plantões', 'Consultas particulares', 'Outros serviços médicos',
      'Repasse conta salário', 'Rendimentos financeiros', 'Outras receitas']
  });
}

function comLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function garantirCategoria_(ss, cat, sub) {
  if (!cat || cat === 'Receitas') return;
  const atuais = lerTabela_(ss, 'CATEGORIAS');
  const tem = atuais.some(function (r) { return r[0] === cat && String(r[1] || '') === String(sub || ''); });
  if (!tem) anexar_(ss, 'CATEGORIAS', [[cat, sub || '', '', '', 'criada no painel']]);
}

function lerMov_(ss) {
  const sh = ss.getSheetByName('MOVIMENTOS');
  const n = sh.getLastRow() - 1;
  return { sh: sh, n: n, rows: n > 0 ? sh.getRange(2, 1, n, ABAS.MOVIMENTOS.length).getValues() : [] };
}
function gravarMov_(m) {
  if (m.n <= 0) return;
  formatarTexto_(m.sh, ABAS.MOVIMENTOS, m.n + 1);
  m.sh.getRange(2, 1, m.n, ABAS.MOVIMENTOS.length).setValues(m.rows);
}

// p = {id, natureza, categoria, sub, tipoReceita, todos}
function salvarCorrecao2_(json) {
  const p = JSON.parse(json);
  return comLock_(function () {
    const ss = planilhaPainel_();
    garantirAba_(ss, 'APRENDIZADO');
    const agora = fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss');
    const usuario = Session.getEffectiveUser().getEmail();
    const nat = natureza_(p.natureza);
    const rec = nat === 'RECEITA';
    const cat = rec ? 'Receitas' : (p.categoria || '');
    const sub = rec ? (p.tipoReceita || '') : (p.sub || '');
    const tr = rec ? (p.tipoReceita || '') : '';
    garantirCategoria_(ss, cat, sub);

    // CORRECOES do lançamento clicado
    const shC = ss.getSheetByName('CORRECOES');
    const corr = lerTabela_(ss, 'CORRECOES');
    const linhaC = [p.id, cat, rec ? '' : sub, tr, nat, p.todos ? 'aprendizado' : '', agora, usuario];
    let idx = -1;
    corr.forEach(function (r, i) { if (String(r[0]) === String(p.id)) idx = i; });
    if (idx >= 0) { formatarTexto_(shC, ABAS.CORRECOES, idx + 2); shC.getRange(idx + 2, 1, 1, linhaC.length).setValues([linhaC]); }
    else anexar_(ss, 'CORRECOES', [linhaC]);

    const H = indice_(ABAS.MOVIMENTOS);
    const m = lerMov_(ss);
    let alvo = null;
    m.rows.forEach(function (r) { if (String(r[H.id]) === String(p.id)) alvo = r; });
    if (!alvo) throw new Error('Lançamento não encontrado.');
    alvo[H.natureza] = nat; alvo[H.categoria] = cat; alvo[H.subcategoria] = sub; alvo[H.tipo_receita] = tr;
    alvo[H.origem] = 'MANUAL'; alvo[H.revisar] = ''; alvo[H.motivo_revisao] = '';
    lentes_(alvo, H);

    let aprId = '', sig = '', n = 0;
    if (p.todos) {
      sig = assinatura_(alvo[H.contraparte_doc], alvo[H.contraparte], alvo[H.descricao]);
      const dir = Number(alvo[H.valor]) > 0 ? 'ENTRADA' : 'SAIDA';
      const shA = ss.getSheetByName('APRENDIZADO');
      const ap = lerTabela_(ss, 'APRENDIZADO');
      aprId = 'A' + new Date().getTime();
      const linhaA = [aprId, 'TRUE', sig, dir, nat, cat, rec ? '' : sub, tr,
        String(alvo[H.contraparte] || alvo[H.descricao]).substring(0, 80), agora, usuario];
      let ia = -1;
      ap.forEach(function (r, i) { if (r[2] === sig && r[3] === dir) ia = i; });
      if (ia >= 0) { aprId = ap[ia][0] || aprId; linhaA[0] = aprId; formatarTexto_(shA, ABAS.APRENDIZADO, ia + 2); shA.getRange(ia + 2, 1, 1, linhaA.length).setValues([linhaA]); }
      else anexar_(ss, 'APRENDIZADO', [linhaA]);
    }
    const antes = {};
    m.rows.forEach(function (r) { antes[r[H.id]] = [r[H.natureza], r[H.categoria], r[H.subcategoria], r[H.tipo_receita]].join('|'); });
    const st = aplicarInteligencia_(ss, m.rows);
    m.rows.forEach(function (r) {
      if (String(r[H.id]) === String(p.id)) return;
      if (antes[r[H.id]] !== [r[H.natureza], r[H.categoria], r[H.subcategoria], r[H.tipo_receita]].join('|')) n++;
    });
    gravarMov_(m);
    return JSON.stringify({ ok: true, aprId: aprId, n: n, previstos: st.previstos, dados: dadosPainel2_() });
  });
}

// desfaz um aprendizado (e a correção do lançamento de origem, se informado) e reprocessa o histórico
function desfazerAprendizado_(json) {
  const p = JSON.parse(json);
  return comLock_(function () {
    const ss = planilhaPainel_();
    garantirAba_(ss, 'APRENDIZADO');
    if (p.aprId) {
      const shA = ss.getSheetByName('APRENDIZADO');
      lerTabela_(ss, 'APRENDIZADO').forEach(function (r, i) { if (r[0] === p.aprId) shA.getRange(i + 2, 2).setValue('FALSE'); });
    }
    if (p.id) {
      const shC = ss.getSheetByName('CORRECOES');
      const corr = lerTabela_(ss, 'CORRECOES');
      for (let i = corr.length - 1; i >= 0; i--) if (String(corr[i][0]) === String(p.id)) shC.deleteRow(i + 2);
    }
    processar_(obterPlanilha_());
    return dadosPainel2_();
  });
}

// p = {acao: 'criar'|'renomear'|'juntar', de, deSub, para, paraSub}
function gerenciarCategoria_(json) {
  const p = JSON.parse(json);
  return comLock_(function () {
    const ss = planilhaPainel_();
    garantirAba_(ss, 'CATEGORIAS_MAPA');
    const agora = fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss');
    const para = String(p.para || '').trim(), paraSub = String(p.paraSub || '').trim();
    if (!para) throw new Error('Informe o nome da categoria.');
    if (p.acao === 'criar') {
      garantirCategoria_(ss, para, paraSub);
      return dadosPainel2_();
    }
    const de = String(p.de || ''), deSub = String(p.deSub || '');
    if (!de) throw new Error('Informe a categoria de origem.');
    // mapa permanente (vale para tudo que vier da API no futuro)
    anexar_(ss, 'CATEGORIAS_MAPA', [[de, deSub, para, deSub ? (paraSub || deSub) : '', agora]]);
    // CATEGORIAS
    const cats = lerTabela_(ss, 'CATEGORIAS').map(function (r) {
      if (r[0] === de && (!deSub || r[1] === deSub)) { r[0] = para; if (deSub) r[1] = paraSub || deSub; }
      return r;
    });
    const vistos = {};
    escreverTabela_(ss, 'CATEGORIAS', cats.filter(function (r) { const k = r[0] + '|' + r[1]; if (vistos[k]) return false; vistos[k] = 1; return true; }));
    // CORRECOES, APRENDIZADO e REGRAS
    const troca = function (nome, ic, is) {
      const rows = lerTabela_(ss, nome);
      let mudou = false;
      rows.forEach(function (r) {
        if (r[ic] === de && (!deSub || r[is] === deSub)) { r[ic] = para; if (deSub) r[is] = paraSub || deSub; mudou = true; }
      });
      if (mudou) escreverTabela_(ss, nome, rows);
    };
    troca('CORRECOES', 1, 2); troca('APRENDIZADO', 5, 6); troca('REGRAS', 4, 5);
    // MOVIMENTOS (aplica o mapa + reaprende na hora)
    const m = lerMov_(ss);
    aplicarInteligencia_(ss, m.rows);
    gravarMov_(m);
    return dadosPainel2_();
  });
}
