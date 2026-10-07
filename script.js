/* =====================================================================
   WebGeo - portal de monitoramento ambiental
   JavaScript puro, sem build. Bibliotecas por CDN (index.html):
   Supabase JS, Leaflet (mapa) e SheetJS (leitura do Excel).
   ===================================================================== */

// ---------------------------------------------------------------------
// Conexão com o Supabase: o MESMO projeto do Perfil de Sondagem (mesmo
// login, mesma empresa). A chave publishable é pública por natureza — quem
// protege os dados de cada empresa são as regras de segurança (RLS) do banco.
// ---------------------------------------------------------------------
const SUPABASE_URL = 'https://wbocmpekamjfwkwmxtjd.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_rackRbyFgs42jUbWBu0Krw_tn7VTJOw';


/* =====================================================================
   1. IMPORTAÇÃO DA PLANILHA MODELO (funções puras, sem tela)
   Mesmas regras do importador: colunas achadas pelo NOME do cabeçalho,
   "< 1,00" = abaixo do LQ (não vira zero), erros com aba e linha.
   ===================================================================== */
const WGImport = (() => {

  /** "Bário total" -> "BARIOTOTAL": compara nomes ignorando acento, espaço e símbolo. */
  function chave(t) {
    if (t === null || t === undefined) return '';
    return String(t).normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toUpperCase().replace(/[^A-Z0-9]/g, '');
  }

  function texto(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    const s = String(v).trim();
    return s === '' ? null : s;
  }

  /** Número BR ou internacional: "12,5" | "1.234,5" | "12.5" | 12.5 */
  function numero(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    let t = String(v).trim().replace(/\s/g, '');
    if (t === '') return null;
    if (t.includes(',')) t = t.replace(/\./g, '').replace(',', '.');
    else if ((t.match(/\./g) || []).length > 1) t = t.replace(/\./g, '');
    if (!/^[-+]?\d*\.?\d+(e[-+]?\d+)?$/i.test(t)) return null;
    return Number(t);
  }

  /** Data: célula de data do Excel, dd/mm/aaaa, aaaa-mm-dd ou aaaa-mm. Devolve 'aaaa-mm-dd' ou null. */
  function data(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date && !isNaN(v)) {
      // SheetJS cria a data no fuso local; somar 12h evita "voltar um dia"
      const d = new Date(v.getTime() + 12 * 3600 * 1000);
      return d.toISOString().slice(0, 10);
    }
    const t = String(v).trim();
    let m;
    if ((m = t.match(/^(\d{4})-(\d{2})-(\d{2})/))) return valida(+m[1], +m[2], +m[3]);
    if ((m = t.match(/^(\d{4})-(\d{2})$/))) return valida(+m[1], +m[2], 1);
    if ((m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) return valida(+m[3], +m[2], +m[1]);
    return null;
  }
  function valida(a, m, d) {
    const dt = new Date(Date.UTC(a, m - 1, d));
    if (dt.getUTCFullYear() !== a || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return dt.toISOString().slice(0, 10);
  }

  /**
   * Valor do laudo:
   *  12,5 -> quantificado | < 1,00 -> abaixo do LQ (LQ 1,00) | ND, <LQ -> abaixo do LQ sem LQ
   *  vazio, "-", "NA" -> não analisado (null)
   */
  function valorLab(v) {
    const original = texto(v);
    if (original === null) return null;
    const u = original.toUpperCase().replace(/\s/g, '');
    if (['-', 'NA', 'N.A.'].includes(u)) return null;
    if (['ND', 'N.D.', '<LQ', '<LD'].includes(u)) return { valor: null, menor_que_lq: true, lq: null, valor_original: original };
    if (u.startsWith('<')) {
      const lq = numero(u.slice(1));
      if (lq === null) throw new Error(`valor inválido: '${original}'`);
      return { valor: null, menor_que_lq: true, lq, valor_original: original };
    }
    const n = numero(typeof v === 'number' ? v : u);
    if (n === null) throw new Error(`valor inválido: '${original}'`);
    if (n < 0) throw new Error(`concentração negativa: '${original}'`);
    return { valor: n, menor_que_lq: false, lq: null, valor_original: original };
  }

  function codigoPoco(v) {
    const t = texto(v);
    return t ? t.toUpperCase().replace(/\s+/g, '') : null;
  }
  function redeDoCodigo(cod) {
    return cod.includes('-') ? cod.slice(0, cod.indexOf('-')) : cod.replace(/[0-9].*$/, '');
  }

  // ---------------------------------------------------------------- leitura das abas
  function acharAba(wb, nome) {
    const n = wb.SheetNames.find(s => chave(s) === chave(nome));
    return n ? wb.Sheets[n] : null;
  }
  function linhasDaAba(XLSX, ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }

  /** Aba em tabela: { cabecalhos:[{nome,chave,col}], linhas:[{n, get(...nomes), bruto(chave)}] } */
  function tabela(XLSX, wb, nome) {
    const ws = acharAba(wb, nome);
    if (!ws) return null;
    const rows = linhasDaAba(XLSX, ws);
    if (!rows.length) return { nome, cabecalhos: [], linhas: [] };
    const cab = [];
    (rows[0] || []).forEach((h, col) => { const t = texto(h); if (t) cab.push({ nome: t, chave: chave(t), col }); });
    const linhas = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const vals = {};
      let vazia = true;
      cab.forEach(c => { const v = r[c.col]; vals[c.chave] = v; if (texto(v) !== null) vazia = false; });
      if (vazia) continue;
      linhas.push({
        n: i + 1,
        bruto: k => vals[k],
        get: (...nomes) => {
          for (const nm of nomes) { const v = vals[chave(nm)]; if (texto(v) !== null) return v; }
          return null;
        }
      });
    }
    return { nome, cabecalhos: cab, linhas };
  }

  /** Aba Projeto: rótulo na coluna A, valor na B. Busca pelo começo do rótulo. */
  function lerProjeto(XLSX, wb) {
    const ws = acharAba(wb, 'Projeto');
    const kv = {};
    if (ws) for (const r of linhasDaAba(XLSX, ws)) {
      if (r && texto(r[0]) && r[1] !== null && r[1] !== undefined && texto(r[1]) !== null) kv[chave(r[0])] = r[1];
    }
    const campo = prefixo => {
      const k = chave(prefixo);
      const achada = Object.keys(kv).find(x => x.startsWith(k));
      return achada ? kv[achada] : null;
    };
    return {
      // "Projeto" (igual ao Perfil de Sondagem); planilhas antigas usavam "Sítio / área"
      projeto: texto(campo('Projeto')) || texto(campo('Sítio')),
      municipio: texto(campo('Município')), uf: texto(campo('UF')),
      campanha: texto(campo('Campanha')), descricao: texto(campo('Descrição')),
      dataInicioBruta: campo('Data início'), dataFimBruta: campo('Data fim'),
      laboratorio: texto(campo('Laboratório'))
    };
  }

  const COL_POCO = ['Poço', 'Poco', 'ID_POCO', 'Ponto'];
  const COL_LAT = ['Latitude', 'Latitude_Corrigida', 'Lat'];
  const COL_LON = ['Longitude', 'Longitude_Corrigida', 'Lon'];
  const COL_COTA = ['Cota topo (m)', 'Cota (m)', 'Cota', 'Cota topo'];
  const NAO_PARAMETRO = new Set(['Poço', 'Poco', 'ID_POCO', 'Ponto', 'Data coleta', 'Data', 'Laudo',
    'Latitude', 'Longitude', 'Latitude_Corrigida', 'Longitude_Corrigida', 'Lat', 'Lon',
    'Cota (m)', 'Cota', 'Cota topo (m)', 'X', 'Y', 'Observação', 'Obs'].map(chave));
  const ESTRUTURA = { SOLO: 'SOLO', SELO: 'SELO', BRITA: 'BRITA', PREFILTRO: 'BRITA', CEGO: 'CEGO',
    REVESTIMENTO: 'CEGO', TUBO: 'CEGO', TUBOCEGO: 'CEGO', FILTRO: 'FILTRO' };
  const COR_PADRAO = { SOLO: '#d4c4a8', SELO: '#556b2f', BRITA: '#9e9e9e', CEGO: '#d3d3d3', FILTRO: '#87cefa' };
  const SITUACOES = ['ATIVO', 'SECO', 'DANIFICADO', 'DESATIVADO'];

  /**
   * Lê e valida a planilha inteira.
   * ctx.parametros: [{nome, sinonimos}] já cadastrados no banco
   * ctx.pocosExistentes: Set com os códigos já cadastrados naquele sítio
   * Devolve { payload, rel } - payload vai para a função wg_importar do banco.
   */
  function ler(XLSX, wb, arquivo, ctx = {}) {
    const rel = { contadores: {}, avisos: [], erros: [] };
    const contar = (k, n = 1) => { rel.contadores[k] = (rel.contadores[k] || 0) + n; };
    const aviso = (aba, n, m) => rel.avisos.push({ aba, linha: n, msg: m });
    const erro = (aba, n, m) => rel.erros.push({ aba, linha: n, msg: m });
    const pocosExistentes = ctx.pocosExistentes || new Set();

    const payload = { arquivo, projeto: null, campanha: null, parametros: [], pocos: [], resultados: [], campo: [], litologia: [] };

    // ---- Projeto
    const pj = lerProjeto(XLSX, wb);
    // Projeto de destino: o da planilha (pelo nome) ou, se vazio, o escolhido na tela
    const destino = ctx.projetoDestino || null; // {id, nome} já resolvido pela tela
    if (!pj.projeto && !destino) erro('Projeto', 0, 'preencha o nome do Projeto (o mesmo do Perfil de Sondagem)');
    payload.projeto = { id: destino?.id || null, nome: destino?.nome || pj.projeto };

    // ---- Parâmetros: mapa nome/sinônimo -> nome oficial
    const mapaParam = new Map();
    const registrar = (nome, sinonimos) => {
      mapaParam.set(chave(nome), nome);
      (sinonimos || '').split(';').map(s => s.trim()).filter(Boolean).forEach(s => { if (!mapaParam.has(chave(s))) mapaParam.set(chave(s), nome); });
    };
    (ctx.parametros || []).forEach(p => registrar(p.nome, p.sinonimos));

    const tParam = tabela(XLSX, wb, 'Parametros');
    if (tParam) for (const l of tParam.linhas) {
      const nome = texto(l.get('Nome', 'Parâmetro'));
      if (!nome) { erro(tParam.nome, l.n, 'Nome do parâmetro vazio'); continue; }
      const vo = l.get('Valor orientador', 'VI');
      if (vo !== null && numero(vo) === null) { erro(tParam.nome, l.n, `valor orientador inválido '${texto(vo)}'`); continue; }
      const sin = texto(l.get('Sinônimos', 'Sinonimos'));
      payload.parametros.push({ nome, sinonimos: sin, grupo: texto(l.get('Grupo')), cas: texto(l.get('CAS')),
        unidade: texto(l.get('Unidade')), valor_orientador: numero(vo), referencia: texto(l.get('Referência')) });
      registrar(nome, sin);
      contar('Parâmetros cadastrados/atualizados');
    }

    // ---- Poços
    const pocosPlanilha = new Set();
    const tPocos = tabela(XLSX, wb, 'Pocos');
    if (tPocos) for (const l of tPocos.linhas) {
      const cod = codigoPoco(l.get(...COL_POCO));
      if (!cod) { erro(tPocos.nome, l.n, 'código do poço vazio'); continue; }
      if (pocosPlanilha.has(cod)) { erro(tPocos.nome, l.n, `${cod} aparece mais de uma vez`); continue; }
      pocosPlanilha.add(cod);

      const lat = coordenada(l.get(...COL_LAT), -90, 90, 'latitude', tPocos.nome, l.n);
      const lon = coordenada(l.get(...COL_LON), -180, 180, 'longitude', tPocos.nome, l.n);
      if (lat === false || lon === false) continue;
      if ((lat === null) !== (lon === null)) { erro(tPocos.nome, l.n, `${cod}: informe latitude E longitude (ou deixe as duas vazias)`); continue; }
      if (lat === null) aviso(tPocos.nome, l.n, `${cod} sem coordenadas: não aparece no mapa até ser preenchido`);

      let situacao = texto(l.get('Situação', 'Situacao'));
      if (situacao) {
        situacao = chave(situacao);
        if (!SITUACOES.includes(situacao)) { erro(tPocos.nome, l.n, 'situação inválida (use ATIVO, SECO, DANIFICADO ou DESATIVADO)'); continue; }
      }
      payload.pocos.push({
        codigo: cod, rede: (texto(l.get('Rede')) || redeDoCodigo(cod)).toUpperCase(),
        latitude: lat, longitude: lon, cota_topo: numero(l.get(...COL_COTA)),
        altitude: numero(l.get('Altitude (m)', 'Altitude')), profundidade: numero(l.get('Profundidade (m)', 'Profundidade')),
        situacao, observacao: texto(l.get('Observação', 'Observacao', 'Obs'))
      });
      contar(pocosExistentes.has(cod) ? 'Poços atualizados' : 'Poços novos');
    }

    function coordenada(v, min, max, nome, aba, n) {
      if (texto(v) === null) return null;
      if (typeof v === 'string' && (v.match(/\./g) || []).length > 1) {
        erro(aba, n, `${nome} mal formatada '${v}' (vários pontos). Use -23.471383 ou -23,471383`); return false;
      }
      const x = numero(v);
      if (x === null || x < min || x > max) { erro(aba, n, `${nome} inválida '${texto(v)}'. Use graus decimais WGS84, ex.: -23.471383`); return false; }
      return x;
    }

    const avisadoSemCadastro = new Set();
    function pocoReferenciado(cod, aba, n) {
      if (pocosPlanilha.has(cod) || pocosExistentes.has(cod) || avisadoSemCadastro.has(cod)) return;
      avisadoSemCadastro.add(cod);
      aviso(aba, n, `${cod} não está cadastrado: será criado SEM coordenadas (preencha na aba Pocos)`);
      contar('Poços criados sem coordenadas');
    }

    // ---- Campanha (obrigatória se houver Resultados ou Campo)
    const tRes = tabela(XLSX, wb, 'Resultados');
    const tCampo = tabela(XLSX, wb, 'Campo');
    const precisaCampanha = (tRes && tRes.linhas.length) || (tCampo && tCampo.linhas.length);
    if (precisaCampanha) {
      const ini = data(pj.dataInicioBruta);
      const fim = data(pj.dataFimBruta);
      if (!pj.campanha) erro('Projeto', 0, 'preencha a Campanha (ex.: C1) - obrigatória quando há Resultados ou Campo');
      if (!ini) erro('Projeto', 0, 'preencha a Data início da campanha (dd/mm/aaaa)');
      if (pj.dataFimBruta && !fim) erro('Projeto', 0, 'Data fim inválida (use dd/mm/aaaa)');
      if (pj.campanha && ini) {
        payload.campanha = { codigo: pj.campanha.toUpperCase(), descricao: pj.descricao, data_inicio: ini, data_fim: fim, laboratorio: pj.laboratorio };
        contar('Campanha ' + payload.campanha.codigo);
      }
    }
    const dataOuErro = (v, aba, n) => {
      const d = data(v);
      if (texto(v) !== null && !d) erro(aba, n, `data inválida '${texto(v)}' (use dd/mm/aaaa)`);
      return d;
    };

    // ---- Resultados: uma linha por poço, uma coluna por parâmetro
    if (tRes && tRes.linhas.length) {
      const colunas = tRes.cabecalhos.filter(c => !NAO_PARAMETRO.has(c.chave));
      if (!colunas.length) erro(tRes.nome, 1, 'nenhuma coluna de parâmetro encontrada');
      const oficial = new Map();
      const usados = new Map();
      for (const c of colunas) {
        let nome = mapaParam.get(c.chave);
        if (!nome) {
          nome = c.nome;
          mapaParam.set(c.chave, nome);
          payload.parametros.push({ nome });
          aviso(tRes.nome, 1, `parâmetro '${nome}' não existia e será criado SEM valor orientador. Se for outro nome de um parâmetro existente, cadastre como sinônimo.`);
        }
        if (usados.has(nome)) { erro(tRes.nome, 1, `as colunas '${usados.get(nome)}' e '${c.nome}' são o mesmo parâmetro (${nome})`); continue; }
        usados.set(nome, c.nome);
        oficial.set(c.chave, nome);
      }
      const vistos = new Set();
      for (const l of tRes.linhas) {
        const cod = codigoPoco(l.get(...COL_POCO));
        if (!cod) { erro(tRes.nome, l.n, 'código do poço vazio'); continue; }
        if (vistos.has(cod)) { erro(tRes.nome, l.n, `${cod} aparece mais de uma vez nesta campanha`); continue; }
        vistos.add(cod);
        pocoReferenciado(cod, tRes.nome, l.n);
        const dataColeta = dataOuErro(l.get('Data coleta', 'Data'), tRes.nome, l.n);
        const laudo = texto(l.get('Laudo'));
        for (const [k, nome] of oficial) {
          let v;
          try { v = valorLab(l.bruto(k)); }
          catch (e) { erro(tRes.nome, l.n, `${cod} / ${usados.get(nome)}: ${e.message}`); continue; }
          if (!v) continue;
          payload.resultados.push({ poco: cod, parametro: nome, data_coleta: dataColeta, laudo, ...v });
          contar(v.menor_que_lq ? 'Resultados abaixo do LQ' : 'Resultados quantificados');
        }
      }
    }

    // ---- Campo
    if (tCampo) {
      const vistos = new Set();
      for (const l of tCampo.linhas) {
        const cod = codigoPoco(l.get(...COL_POCO));
        if (!cod) { erro(tCampo.nome, l.n, 'código do poço vazio'); continue; }
        if (vistos.has(cod)) { erro(tCampo.nome, l.n, `${cod} aparece mais de uma vez`); continue; }
        vistos.add(cod);
        pocoReferenciado(cod, tCampo.nome, l.n);
        const naBruto = l.get('N.A. (m)', 'NA (m)', 'NA', 'N.A.', "Nível d'água (m)");
        const na = numero(naBruto);
        if (naBruto !== null && na === null) { erro(tCampo.nome, l.n, `N.A. inválido '${texto(naBruto)}'`); continue; }
        const esp = numero(l.get('Espessura FL (m)', 'Espessura FL'));
        const fl = chave(l.get('Fase livre (S/N)', 'Fase livre'));
        const faseLivre = ['S', 'SIM', 'TRUE', '1', 'X'].includes(fl) || (esp !== null && esp > 0);
        payload.campo.push({ poco: cod, data: dataOuErro(l.get('Data'), tCampo.nome, l.n), nivel_agua: na, fase_livre: faseLivre, espessura_fl: esp });
        contar('Medições de campo');
        if (faseLivre) contar('Poços com fase livre');
      }
    }

    // ---- Litologia
    const tLit = tabela(XLSX, wb, 'Litologia');
    if (tLit) for (const l of tLit.linhas) {
      const cod = codigoPoco(l.get(...COL_POCO));
      if (!cod) { erro(tLit.nome, l.n, 'código do poço vazio'); continue; }
      let de = numero(l.get('De (m)', 'De')), ate = numero(l.get('Até (m)', 'Ate (m)', 'Até', 'Ate'));
      if (de === null || ate === null) { erro(tLit.nome, l.n, 'preencha De e Até'); continue; }
      de = Math.abs(de); ate = Math.abs(ate);
      if (ate <= de) { erro(tLit.nome, l.n, "'Até' deve ser maior que 'De'"); continue; }
      const estBruta = texto(l.get('Estrutura'));
      const est = ESTRUTURA[chave(estBruta)];
      if (!est) { erro(tLit.nome, l.n, `estrutura inválida '${estBruta || ''}' (use Solo, Selo, Brita, Cego ou Filtro)`); continue; }
      let cor = texto(l.get('Cor'));
      if (!cor || !/^#[0-9a-f]{6}$/i.test(cor)) cor = COR_PADRAO[est];
      pocoReferenciado(cod, tLit.nome, l.n);
      payload.litologia.push({ poco: cod, de, ate, estrutura: est, descricao: texto(l.get('Descrição', 'Descricao')), cor });
      contar('Intervalos de perfil');
    }

    if (!tPocos && !tRes && !tCampo && !tLit && !tParam) erro(null, 0, 'nenhuma aba reconhecida. Use a planilha modelo do WebGeo.');
    else if (!payload.pocos.length && !payload.resultados.length && !payload.campo.length && !payload.litologia.length && !payload.parametros.length)
      erro(null, 0, 'a planilha não tem dados para importar (abas Pocos, Resultados, Campo e Litologia vazias).');
    payload.resumo = { contadores: rel.contadores, avisos: rel.avisos.length };
    return { payload, rel };
  }

  return { ler, lerProjeto, valorLab, numero, data, chave };
})();

/* =====================================================================
   2. PLUMA (funções puras, sem tela)
   Interpolação IDW em escala log numa grade fina em metros, área acima
   do limiar, centróide e linha do limiar (marching squares).
   ===================================================================== */
const WGPluma = (() => {

  const M_POR_GRAU_LAT = 110574;

  /** Projeção local em metros (equiretangular), exata o bastante para uma área de alguns km. */
  function projecao(lat0, lon0) {
    const kx = 111320 * Math.cos(lat0 * Math.PI / 180);
    return {
      paraXY: (lat, lon) => ({ x: (lon - lon0) * kx, y: (lat - lat0) * M_POR_GRAU_LAT }),
      paraLatLon: (x, y) => ({ lat: lat0 + y / M_POR_GRAU_LAT, lon: lon0 + x / kx })
    };
  }

  /**
   * Monta a grade comum (mesma para todas as campanhas, para dar para comparar).
   * pocos: [{lat, lon}] — todos os poços com coordenadas que têm resultado do parâmetro.
   * A grade cobre a rede de poços + uma margem do tamanho do maior espaçamento entre poços.
   * Se a área for muito grande para a célula pedida, a célula é aumentada (limite de células).
   */
  function grade(pocos, cel, maxCelulas = 80000) {
    if (!pocos.length) return null;
    const lat0 = pocos.reduce((s, p) => s + p.lat, 0) / pocos.length;
    const lon0 = pocos.reduce((s, p) => s + p.lon, 0) / pocos.length;
    const proj = projecao(lat0, lon0);
    const pts = pocos.map(p => proj.paraXY(p.lat, p.lon));
    const margem = Math.max(...espacamentos(pts)) + cel;
    const x0 = Math.min(...pts.map(p => p.x)) - margem, x1 = Math.max(...pts.map(p => p.x)) + margem;
    const y0 = Math.min(...pts.map(p => p.y)) - margem, y1 = Math.max(...pts.map(p => p.y)) + margem;
    while ((x1 - x0) * (y1 - y0) / (cel * cel) > maxCelulas) cel = Math.ceil(cel * 1.5);
    const xmin = Math.floor(x0 / cel) * cel, ymin = Math.floor(y0 / cel) * cel;
    const nx = Math.max(1, Math.ceil((x1 - xmin) / cel)), ny = Math.max(1, Math.ceil((y1 - ymin) / cel));
    return { proj, cel, xmin, ymin, nx, ny };
  }

  /**
   * Junta poços que ficam praticamente no mesmo lugar (até "raio" metros, padrão 5 m) num
   * único ponto, valendo o MAIOR valor. É o caso dos poços multinível (PM-26, PMN-26A,
   * PMN-26B): em planta estão no mesmo ponto, e um poço limpo colado num contaminado
   * anularia a pluma. Usar o maior valor é o critério conservador.
   * amostras: [{lat, lon, v}] -> devolve a lista reduzida (com n = quantos poços no ponto).
   */
  function agrupar(amostras, raio = 5) {
    if (amostras.length < 2) return amostras.map(a => ({ ...a, n: 1 }));
    const proj = projecao(amostras[0].lat, amostras[0].lon);
    const xy = amostras.map(a => proj.paraXY(a.lat, a.lon));
    const pai = amostras.map((_, i) => i);
    const raiz = i => { while (pai[i] !== i) { pai[i] = pai[pai[i]]; i = pai[i]; } return i; };
    for (let i = 0; i < xy.length; i++) for (let j = i + 1; j < xy.length; j++) {
      if (Math.hypot(xy[i].x - xy[j].x, xy[i].y - xy[j].y) <= raio) pai[raiz(i)] = raiz(j);
    }
    const grupos = new Map();
    amostras.forEach((a, i) => {
      const k = raiz(i), g = grupos.get(k);
      if (!g) grupos.set(k, { ...a, n: 1 });
      else { g.n++; if (a.v > g.v) Object.assign(g, a, { n: g.n }); }
    });
    return [...grupos.values()];
  }

  /** Distância de cada poço ao vizinho mais próximo (m), limitada entre 5 e 150 m. */
  function espacamentos(pts) {
    return pts.map((a, i) => {
      let m = Infinity;
      pts.forEach((b, j) => { if (i !== j) m = Math.min(m, Math.hypot(a.x - b.x, a.y - b.y)); });
      return Math.min(150, Math.max(5, Number.isFinite(m) ? m : 20));
    });
  }

  /**
   * IDW em ESCALA LOGARÍTMICA: interpola log10(concentração) com peso 1/distância^p,
   * usando todos os poços (o resultado é uma superfície suave).
   *
   * Por que em log: concentrações variam por ordens de grandeza. Na média direta, um poço
   * com 1000 "contamina" os vizinhos limpos e a pluma engole poços abaixo do limiar. Em log,
   * a pluma fica em volta dos poços contaminados e termina antes de chegar aos poços limpos.
   * Nos pontos dos poços o valor é exatamente o do poço.
   *
   * Fora da rede de poços (além da envoltória) não há dado: o valor decai suavemente até o
   * piso numa distância igual ao espaçamento local entre poços, fechando a pluma em curva.
   *
   * amostras: [{lat, lon, v}] com v > 0 (abaixo do LQ já convertido pelo chamador).
   * piso: valor "limpo" para onde a pluma decai fora da rede.
   */
  function idw(g, amostras, p = 2, piso = null) {
    const v = new Float64Array(g.nx * g.ny);
    if (!amostras.length) return v.fill(NaN);
    const menor = Math.min(...amostras.map(a => a.v));
    const lPiso = Math.log10(Math.max(1e-9, Math.min(menor, piso ?? menor)));
    const pts = amostras.map(a => ({ ...g.proj.paraXY(a.lat, a.lon), l: Math.log10(Math.max(1e-9, a.v)) }));
    const esp = espacamentos(pts);
    const env = envoltoria(pts);
    for (let j = 0; j < g.ny; j++) for (let i = 0; i < g.nx; i++) {
      const x = g.xmin + (i + .5) * g.cel, y = g.ymin + (j + .5) * g.cel;
      let num = 0, den = 0, numR = 0, exato = null;
      for (let q = 0; q < pts.length; q++) {
        const a = pts[q], d2 = (a.x - x) ** 2 + (a.y - y) ** 2;
        if (d2 < 1e-6) { exato = a.l; break; }
        const w = p === 2 ? 1 / d2 : 1 / Math.pow(d2, p / 2);
        num += w * a.l; den += w; numR += w * esp[q];
      }
      let l = exato !== null ? exato : num / den;
      if (exato === null) {
        const fora = distanciaFora(x, y, env);
        if (fora > 0) {
          const t = Math.min(1, fora / (numR / den));       // 0 na borda da rede, 1 a um espaçamento de distância
          l = lPiso + (l - lPiso) * (1 - t * t * (3 - 2 * t)); // decaimento suave (smoothstep)
        }
      }
      v[j * g.nx + i] = Math.pow(10, l);
    }
    return v;
  }

  /** Distância (m) de um ponto até a envoltória da rede; 0 se estiver dentro. */
  function distanciaFora(x, y, env) {
    if (env.length >= 3 && dentroPoligono(x, y, env)) return 0;
    if (env.length === 1) return Math.hypot(x - env[0].x, y - env[0].y);
    let m = Infinity;
    for (let i = 0; i < env.length; i++) {
      const a = env[i], b = env[(i + 1) % env.length];
      const dx = b.x - a.x, dy = b.y - a.y, c = dx * dx + dy * dy;
      const t = c ? Math.max(0, Math.min(1, ((x - a.x) * dx + (y - a.y) * dy) / c)) : 0;
      m = Math.min(m, Math.hypot(x - a.x - t * dx, y - a.y - t * dy));
    }
    return m;
  }

  /** Área (m²) e centróide (lat/lon) das células com valor >= limiar. */
  function areaECentroide(g, v, limiar) {
    let n = 0, sx = 0, sy = 0;
    for (let j = 0; j < g.ny; j++) for (let i = 0; i < g.nx; i++) {
      if (v[j * g.nx + i] >= limiar) { n++; sx += g.xmin + (i + .5) * g.cel; sy += g.ymin + (j + .5) * g.cel; }
    }
    const area = n * g.cel * g.cel;
    if (!n) return { area: 0, centroide: null };
    const c = { x: sx / n, y: sy / n };
    return { area, centroide: { ...c, ...g.proj.paraLatLon(c.x, c.y) } };
  }

  /** Distância em metros entre dois centróides. */
  function distancia(a, b) {
    if (!a || !b) return null;
    return Math.hypot(a.x - b.x, a.y - b.y);
  }

  /**
   * Linha do limiar (marching squares sobre os centros das células).
   * Os pedaços são emendados em curvas contínuas e suavizados (Chaikin), para o
   * traçado sair liso. Devolve uma lista de curvas, cada uma [[lat,lon], ...].
   */
  function isolinha(g, v, limiar) {
    const segs = [];
    const val = (i, j) => {
      if (i < 0 || j < 0 || i >= g.nx || j >= g.ny) return -Infinity;
      const x = v[j * g.nx + i];
      return Number.isNaN(x) ? -Infinity : x;
    };
    const ponto = (i, j) => ({ x: g.xmin + (i + .5) * g.cel, y: g.ymin + (j + .5) * g.cel });
    const interp = (p1, v1, p2, v2) => {
      // interpola em log (a superfície é suave em log); fora da grade, meio do caminho
      const t = (!Number.isFinite(v1) || !Number.isFinite(v2)) ? .5
        : (v1 > 0 && v2 > 0 && limiar > 0) ? Math.log(limiar / v1) / Math.log(v2 / v1) : (limiar - v1) / (v2 - v1);
      return { x: p1.x + t * (p2.x - p1.x), y: p1.y + t * (p2.y - p1.y) };
    };
    for (let j = -1; j < g.ny; j++) for (let i = -1; i < g.nx; i++) {
      const c = [[i, j], [i + 1, j], [i + 1, j + 1], [i, j + 1]];
      const vv = c.map(([a, b]) => val(a, b));
      const pp = c.map(([a, b]) => ponto(a, b));
      const caso = vv.reduce((s, x, k) => s | ((x >= limiar ? 1 : 0) << k), 0);
      if (caso === 0 || caso === 15) continue;
      const e = k => interp(pp[k], vv[k], pp[(k + 1) % 4], vv[(k + 1) % 4]); // aresta k: canto k -> k+1
      const cruzam = [0, 1, 2, 3].filter(k => ((caso >> k) & 1) !== ((caso >> ((k + 1) % 4)) & 1));
      if (cruzam.length === 2) segs.push([e(cruzam[0]), e(cruzam[1])]);
      else if (cruzam.length === 4) { segs.push([e(0), e(1)]); segs.push([e(2), e(3)]); }
    }
    return emendar(segs).map(c => suavizar(c, 2).map(p => { const ll = g.proj.paraLatLon(p.x, p.y); return [ll.lat, ll.lon]; }));
  }

  /** Emenda segmentos soltos [[a,b],...] em curvas contínuas (pontas iguais se juntam). */
  function emendar(segs) {
    const chave = p => Math.round(p.x * 1000) + ',' + Math.round(p.y * 1000);
    const pontas = new Map(); // chave do ponto -> [{s, ponta}]
    segs.forEach((sg, s) => sg.forEach((p, ponta) => {
      const k = chave(p); if (!pontas.has(k)) pontas.set(k, []); pontas.get(k).push({ s, ponta });
    }));
    const usado = new Uint8Array(segs.length), curvas = [];
    const seguir = (s, ponta) => { // anda a partir da ponta de um segmento, devolve os pontos seguintes
      const pts = [];
      for (;;) {
        const p = segs[s][ponta];
        const prox = (pontas.get(chave(p)) || []).find(o => o.s !== s && !usado[o.s]);
        if (!prox) return pts;
        usado[prox.s] = 1; s = prox.s; ponta = 1 - prox.ponta;
        pts.push(segs[s][ponta]);
      }
    };
    for (let s = 0; s < segs.length; s++) {
      if (usado[s]) continue;
      usado[s] = 1;
      const frente = seguir(s, 1), tras = seguir(s, 0);
      curvas.push([...tras.reverse(), segs[s][0], segs[s][1], ...frente]);
    }
    return curvas;
  }

  /** Suavização de Chaikin (corta os cantos); mantém fechada a curva que já era fechada. */
  function suavizar(c, vezes) {
    const fechada = c.length > 3 && Math.hypot(c[0].x - c[c.length - 1].x, c[0].y - c[c.length - 1].y) < 1e-3;
    let pts = fechada ? c.slice(0, -1) : c;
    for (let k = 0; k < vezes && pts.length > 2; k++) {
      const n = pts.length, novo = [];
      if (!fechada) novo.push(pts[0]);
      for (let i = 0; i < (fechada ? n : n - 1); i++) {
        const a = pts[i], b = pts[(i + 1) % n];
        novo.push({ x: .75 * a.x + .25 * b.x, y: .75 * a.y + .25 * b.y }, { x: .25 * a.x + .75 * b.x, y: .25 * a.y + .75 * b.y });
      }
      if (!fechada) novo.push(pts[n - 1]);
      pts = novo;
    }
    return fechada ? [...pts, pts[0]] : pts;
  }

  // ------------------------------------------------------------ geometria auxiliar
  function envoltoria(pts) { // convex hull (monotone chain)
    const p = [...pts].sort((a, b) => a.x - b.x || a.y - b.y);
    if (p.length < 3) return p;
    const cross = (o, a, b) => (a.x - o.x) * (b.y - o.y) - (a.y - o.y) * (b.x - o.x);
    const lo = [], hi = [];
    for (const q of p) { while (lo.length >= 2 && cross(lo[lo.length - 2], lo[lo.length - 1], q) <= 0) lo.pop(); lo.push(q); }
    for (const q of [...p].reverse()) { while (hi.length >= 2 && cross(hi[hi.length - 2], hi[hi.length - 1], q) <= 0) hi.pop(); hi.push(q); }
    return lo.slice(0, -1).concat(hi.slice(0, -1));
  }
  function dentroPoligono(x, y, pol) {
    let dentro = false;
    for (let i = 0, j = pol.length - 1; i < pol.length; j = i++) {
      const a = pol[i], b = pol[j];
      if ((a.y > y) !== (b.y > y) && x < (b.x - a.x) * (y - a.y) / (b.y - a.y) + a.x) dentro = !dentro;
    }
    return dentro;
  }

  // ------------------------------------------------------------ cor (escala log, um tom: laranja -> vermelho escuro)
  const RAMPA = ['#fee6ce', '#fdd0a2', '#fdae6b', '#fd8d3c', '#f16913', '#d94801', '#a63603', '#7f2704'];
  function hexRgb(h) { return [1, 3, 5].map(k => parseInt(h.slice(k, k + 2), 16)); }
  /** Cor de um valor numa escala log10 entre vmin e vmax (vmin > 0). */
  function cor(valor, vmin, vmax) {
    const t = Math.max(0, Math.min(1, (Math.log10(valor) - Math.log10(vmin)) / (Math.log10(vmax) - Math.log10(vmin) || 1)));
    const f = (0.3 + 0.7 * t) * (RAMPA.length - 1), k = Math.min(RAMPA.length - 2, Math.floor(f)), r = f - k;
    const a = hexRgb(RAMPA[k]), b = hexRgb(RAMPA[k + 1]);
    return a.map((x, n) => Math.round(x + (b[n] - x) * r));
  }

  return { grade, idw, agrupar, areaECentroide, distancia, isolinha, cor, RAMPA, projecao,
    emendar, suavizar, envoltoria, espacamentos, distanciaFora };
})();

/* =====================================================================
   3. PLANTA EM DXF (funções puras, sem tela)
   Lê um arquivo DXF (texto) e devolve as linhas do desenho por camada,
   já com os blocos "explodidos". Também converte UTM <-> latitude/longitude.
   ===================================================================== */
const WGDxf = (() => {

  // ------------------------------------------------------------ UTM (SIRGAS 2000 / WGS84), série de Krüger
  const A_ELIP = 6378137, F_ELIP = 1 / 298.257222101, K0 = 0.9996;
  const N3 = F_ELIP / (2 - F_ELIP), N3_2 = N3 * N3, N3_3 = N3_2 * N3;
  const A_RET = A_ELIP / (1 + N3) * (1 + N3_2 / 4 + N3_2 * N3_2 / 64);
  const ALFA = [N3 / 2 - 2 * N3_2 / 3 + 5 * N3_3 / 16, 13 * N3_2 / 48 - 3 * N3_3 / 5, 61 * N3_3 / 240];
  const BETA = [N3 / 2 - 2 * N3_2 / 3 + 37 * N3_3 / 96, N3_2 / 48 + N3_3 / 15, 17 * N3_3 / 480];
  const DELTA = [2 * N3 - 2 * N3_2 / 3 - 2 * N3_3, 7 * N3_2 / 3 - 8 * N3_3 / 5, 56 * N3_3 / 15];
  const meridiano = zona => (zona * 6 - 183) * Math.PI / 180;

  /** UTM -> {lat, lon} em graus. sul = hemisfério sul (padrão no Brasil). */
  function utmParaLatLon(E, N, zona, sul = true) {
    const xi = (N - (sul ? 1e7 : 0)) / (K0 * A_RET), eta = (E - 5e5) / (K0 * A_RET);
    let xi1 = xi, eta1 = eta;
    for (let j = 1; j <= 3; j++) {
      xi1 -= BETA[j - 1] * Math.sin(2 * j * xi) * Math.cosh(2 * j * eta);
      eta1 -= BETA[j - 1] * Math.cos(2 * j * xi) * Math.sinh(2 * j * eta);
    }
    const chi = Math.asin(Math.sin(xi1) / Math.cosh(eta1));
    let phi = chi;
    for (let j = 1; j <= 3; j++) phi += DELTA[j - 1] * Math.sin(2 * j * chi);
    const lam = meridiano(zona) + Math.atan2(Math.sinh(eta1), Math.cos(xi1));
    return { lat: phi * 180 / Math.PI, lon: lam * 180 / Math.PI };
  }

  /** {lat, lon} em graus -> UTM {E, N} na zona dada. */
  function latLonParaUtm(lat, lon, zona, sul = true) {
    const phi = lat * Math.PI / 180, dl = lon * Math.PI / 180 - meridiano(zona);
    const c = 2 * Math.sqrt(N3) / (1 + N3);
    const t = Math.sinh(Math.atanh(Math.sin(phi)) - c * Math.atanh(c * Math.sin(phi)));
    const xi1 = Math.atan2(t, Math.cos(dl)), eta1 = Math.atanh(Math.sin(dl) / Math.sqrt(1 + t * t));
    let xi = xi1, eta = eta1;
    for (let j = 1; j <= 3; j++) {
      xi += ALFA[j - 1] * Math.sin(2 * j * xi1) * Math.cosh(2 * j * eta1);
      eta += ALFA[j - 1] * Math.cos(2 * j * xi1) * Math.sinh(2 * j * eta1);
    }
    return { E: 5e5 + K0 * A_RET * eta, N: (sul ? 1e7 : 0) + K0 * A_RET * xi };
  }
  const zonaDaLongitude = lon => Math.floor((lon + 180) / 6) + 1;

  // ------------------------------------------------------------ leitura do DXF
  /** Cores básicas do AutoCAD (índice ACI) para as camadas. */
  const ACI = { 1: '#ff3b30', 2: '#ffd60a', 3: '#34c759', 4: '#32d7e6', 5: '#3a82f7', 6: '#d85bd8', 7: '#ffffff', 8: '#9a9a9a', 9: '#c8c8c8',
    30: '#ff9f0a', 40: '#ffb340', 50: '#e6d200', 140: '#3aa0ff', 250: '#555555', 251: '#6b6b6b', 252: '#858585', 253: '#a0a0a0', 254: '#c0c0c0' };
  const corAci = n => ACI[Math.abs(n)] || '#ffffff';

  const mult = (m, n) => [ // compõe duas transformações afins [a,b,c,d,e,f]: x' = a x + b y + e ; y' = c x + d y + f
    m[0] * n[0] + m[1] * n[2], m[0] * n[1] + m[1] * n[3], m[2] * n[0] + m[3] * n[2], m[2] * n[1] + m[3] * n[3],
    m[0] * n[4] + m[1] * n[5] + m[4], m[2] * n[4] + m[3] * n[5] + m[5]];
  const aplicar = (m, x, y) => [m[0] * x + m[1] * y + m[4], m[2] * x + m[3] * y + m[5]];
  const IDENT = [1, 0, 0, 1, 0, 0];

  /**
   * Lê o texto de um DXF. Devolve:
   *   camadas: [{nome, cor, linhas: [[x,y,x,y,...], ...]}]  (coordenadas do desenho)
   *   bbox: {xmin, ymin, xmax, ymax}, entidades: {TIPO: n}, ignoradas: {TIPO: n}
   * Entidades lidas: LINE, LWPOLYLINE, POLYLINE, CIRCLE, ARC, ELLIPSE, SPLINE e INSERT (blocos, inclusive aninhados).
   * Textos, hachuras e cotas são ignorados.
   */
  function ler(texto) {
    const linhas = texto.split(/\r\n|\r|\n/);
    const pares = [];
    for (let i = 0; i + 1 < linhas.length; i += 2) pares.push([parseInt(linhas[i], 10), linhas[i + 1]]);
    if (!pares.length || pares.every(p => Number.isNaN(p[0]))) throw new Error('Este arquivo não parece ser um DXF de texto (se for DXF binário ou DWG, salve como "DXF ASCII" no CAD).');

    // separa em seções
    const secoes = {};
    for (let i = 0; i < pares.length; i++) {
      if (pares[i][0] === 0 && pares[i][1].trim() === 'SECTION' && pares[i + 1] && pares[i + 1][0] === 2) {
        const nome = pares[i + 1][1].trim(); let j = i + 2;
        while (j < pares.length && !(pares[j][0] === 0 && pares[j][1].trim() === 'ENDSEC')) j++;
        secoes[nome] = pares.slice(i + 2, j); i = j;
      }
    }
    if (!secoes.ENTITIES) throw new Error('DXF sem a seção ENTITIES (arquivo vazio ou incompleto).');

    // camadas e suas cores
    const corCamada = new Map();
    (function () {
      const t = secoes.TABLES || []; let nome = null, emLayer = false;
      for (const [c, v] of t) {
        if (c === 0) { emLayer = v.trim() === 'LAYER'; nome = null; }
        else if (emLayer && c === 2) nome = v.trim();
        else if (emLayer && c === 62 && nome !== null) corCamada.set(nome, parseInt(v, 10));
      }
    })();

    // quebra uma seção em entidades cruas: {tipo, g: [[codigo, valor], ...]}
    const entidades = lista => {
      const out = []; let atual = null;
      for (const [c, v] of lista) {
        if (c === 0) { atual = { tipo: v.trim(), g: [] }; out.push(atual); }
        else if (atual) atual.g.push([c, v]);
      }
      return out;
    };
    const num = (e, cod, pad = 0) => { const p = e.g.find(x => x[0] === cod); return p ? parseFloat(p[1]) : pad; };
    const str = (e, cod, pad = '') => { const p = e.g.find(x => x[0] === cod); return p ? p[1].trim() : pad; };

    // blocos
    const blocos = new Map();
    (function () {
      let atual = null;
      for (const e of entidades(secoes.BLOCKS || [])) {
        if (e.tipo === 'BLOCK') { atual = { nome: str(e, 2), bx: num(e, 10), by: num(e, 20), ents: [] }; }
        else if (e.tipo === 'ENDBLK') { if (atual) blocos.set(atual.nome, atual); atual = null; }
        else if (atual) atual.ents.push(e);
      }
    })();

    const porCamada = new Map(), contagem = {}, ignoradas = {};
    const bbox = { xmin: Infinity, ymin: Infinity, xmax: -Infinity, ymax: -Infinity };
    let totalPontos = 0;
    const emitir = (camada, pts, m) => {
      if (pts.length < 2) return;
      const l = new Array(pts.length * 2);
      for (let i = 0; i < pts.length; i++) {
        const [x, y] = aplicar(m, pts[i][0], pts[i][1]);
        l[2 * i] = x; l[2 * i + 1] = y;
        if (x < bbox.xmin) bbox.xmin = x; if (x > bbox.xmax) bbox.xmax = x;
        if (y < bbox.ymin) bbox.ymin = y; if (y > bbox.ymax) bbox.ymax = y;
      }
      totalPontos += pts.length;
      if (totalPontos > 600000) throw new Error('Desenho grande demais (mais de 600 mil pontos). No CAD, limpe o que não precisa (PURGE) ou exporte só as camadas principais.');
      if (!porCamada.has(camada)) porCamada.set(camada, []);
      porCamada.get(camada).push(l);
    };
    const arco = (cx, cy, r, a0, a1) => { // ângulos em radianos, sentido anti-horário de a0 a a1
      while (a1 <= a0) a1 += 2 * Math.PI;
      const n = Math.max(2, Math.ceil((a1 - a0) / (Math.PI / 18)));
      const pts = [];
      for (let i = 0; i <= n; i++) { const a = a0 + (a1 - a0) * i / n; pts.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]); }
      return pts;
    };
    const comBulge = (v, fechada) => { // v: [{x,y,b}], b = "bulge" (arco entre este vértice e o próximo)
      const pts = [], n = v.length, fim = fechada ? n : n - 1;
      for (let i = 0; i < fim; i++) {
        const p = v[i], q = v[(i + 1) % n];
        pts.push([p.x, p.y]);
        if (p.b && Math.abs(p.b) > 1e-9) {
          const dx = q.x - p.x, dy = q.y - p.y, corda = Math.hypot(dx, dy);
          if (corda > 1e-9) {
            const ang = 4 * Math.atan(p.b), r = corda / (2 * Math.sin(ang / 2));
            const mx = (p.x + q.x) / 2, my = (p.y + q.y) / 2, h = r * Math.cos(ang / 2);
            const cx = mx - h * dy / corda, cy = my + h * dx / corda;
            const a0 = Math.atan2(p.y - cy, p.x - cx), passos = Math.max(2, Math.ceil(Math.abs(ang) / (Math.PI / 18)));
            for (let k = 1; k < passos; k++) { const a = a0 + ang * k / passos; pts.push([cx + Math.abs(r) * Math.cos(a), cy + Math.abs(r) * Math.sin(a)]); }
          }
        }
      }
      pts.push(fechada ? [v[0].x, v[0].y] : [v[n - 1].x, v[n - 1].y]);
      return pts;
    };
    const spline = e => { // B-spline por de Boor; sem nós válidos, usa os pontos de ajuste ou de controle
      const xs = e.g.filter(x => x[0] === 10).map(x => parseFloat(x[1])), ys = e.g.filter(x => x[0] === 20).map(x => parseFloat(x[1]));
      const ctrl = xs.map((x, i) => [x, ys[i]]);
      const fx = e.g.filter(x => x[0] === 11).map(x => parseFloat(x[1])), fy = e.g.filter(x => x[0] === 21).map(x => parseFloat(x[1]));
      const nos = e.g.filter(x => x[0] === 40).map(x => parseFloat(x[1])), grau = num(e, 71, 3);
      if (ctrl.length < 2) return fx.map((x, i) => [x, fy[i]]);
      if (nos.length !== ctrl.length + grau + 1) return fx.length > 1 ? fx.map((x, i) => [x, fy[i]]) : ctrl;
      const t0 = nos[grau], t1 = nos[ctrl.length], passos = Math.min(400, ctrl.length * 8), pts = [];
      for (let s = 0; s <= passos; s++) {
        const t = Math.min(t1 - 1e-9, t0 + (t1 - t0) * s / passos);
        let k = grau; while (k < ctrl.length - 1 && nos[k + 1] <= t) k++;
        const d = []; for (let j = 0; j <= grau; j++) d.push(ctrl[j + k - grau].slice());
        for (let r = 1; r <= grau; r++) for (let j = grau; j >= r; j--) {
          const den = nos[j + 1 + k - r] - nos[j + k - grau], a = den ? (t - nos[j + k - grau]) / den : 0;
          d[j] = [(1 - a) * d[j - 1][0] + a * d[j][0], (1 - a) * d[j - 1][1] + a * d[j][1]];
        }
        pts.push(d[grau]);
      }
      return pts;
    };

    function desenhar(lista, m, camadaPai, nivel) {
      for (let i = 0; i < lista.length; i++) {
        const e = lista[i];
        if (num(e, 67, 0) === 1) continue; // espaço de papel (layout de impressão)
        let camada = str(e, 8, '0'); if (camada === '0' && camadaPai) camada = camadaPai;
        const esp = num(e, 230, 1) < 0; // entidade espelhada (extrusão negativa)
        const mm = esp ? mult(m, [-1, 0, 0, 1, 0, 0]) : m;
        let ok = true;
        switch (e.tipo) {
          case 'LINE': emitir(camada, [[num(e, 10), num(e, 20)], [num(e, 11), num(e, 21)]], m); break;
          case 'LWPOLYLINE': {
            const v = []; let at = null;
            for (const [c, val] of e.g) {
              if (c === 10) { at = { x: parseFloat(val), y: 0, b: 0 }; v.push(at); }
              else if (c === 20 && at) at.y = parseFloat(val);
              else if (c === 42 && at) at.b = parseFloat(val);
            }
            if (v.length > 1) emitir(camada, comBulge(v, (num(e, 70, 0) & 1) === 1), mm);
            break;
          }
          case 'POLYLINE': {
            const fechada = (num(e, 70, 0) & 1) === 1, v = [];
            let j = i + 1;
            for (; j < lista.length && lista[j].tipo === 'VERTEX'; j++) {
              if (num(lista[j], 70, 0) & 128 && !(num(lista[j], 70, 0) & 64)) continue; // face de malha
              v.push({ x: num(lista[j], 10), y: num(lista[j], 20), b: num(lista[j], 42, 0) });
            }
            i = j - 1;
            if (v.length > 1) emitir(camada, comBulge(v, fechada), mm);
            break;
          }
          case 'VERTEX': case 'SEQEND': case 'ATTRIB': continue;
          case 'CIRCLE': emitir(camada, arco(num(e, 10), num(e, 20), num(e, 40), 0, 2 * Math.PI), mm); break;
          case 'ARC': emitir(camada, arco(num(e, 10), num(e, 20), num(e, 40), num(e, 50) * Math.PI / 180, num(e, 51) * Math.PI / 180), mm); break;
          case 'ELLIPSE': {
            const cx = num(e, 10), cy = num(e, 20), ax = num(e, 11), ay = num(e, 21), razao = num(e, 40, 1);
            let p0 = num(e, 41, 0), p1 = num(e, 42, 2 * Math.PI); while (p1 <= p0) p1 += 2 * Math.PI;
            const n = Math.max(8, Math.ceil((p1 - p0) / (Math.PI / 24))), pts = [];
            for (let k = 0; k <= n; k++) {
              const t = p0 + (p1 - p0) * k / n, c = Math.cos(t), s = Math.sin(t);
              pts.push([cx + ax * c - ay * razao * s, cy + ay * c + ax * razao * s]);
            }
            emitir(camada, pts, m); break;
          }
          case 'SPLINE': { const pts = spline(e); if (pts.length > 1) { if (num(e, 70, 0) & 1) pts.push(pts[0]); emitir(camada, pts, m); } break; }
          case 'INSERT': {
            const b = blocos.get(str(e, 2));
            if (!b || nivel >= 12) { ok = false; break; }
            const rot = num(e, 50, 0) * Math.PI / 180, sx = num(e, 41, 1), sy = num(e, 42, 1);
            const cr = Math.cos(rot), sr = Math.sin(rot);
            // ponto do bloco -> tira a base, escala, gira, leva ao ponto de inserção
            const t = mult([cr * sx, -sr * sy, sr * sx, cr * sy, num(e, 10), num(e, 20)], [1, 0, 0, 1, -b.bx, -b.by]);
            desenhar(b.ents, mult(mm, t), camada, nivel + 1);
            break;
          }
          default: ok = false;
        }
        const alvo = ok ? contagem : ignoradas;
        alvo[e.tipo] = (alvo[e.tipo] || 0) + 1;
      }
    }
    desenhar(entidades(secoes.ENTITIES), IDENT, null, 0);
    if (!porCamada.size) throw new Error('Não encontrei linhas neste DXF (só textos, hachuras ou imagens?).');

    const camadas = [...porCamada.entries()]
      .map(([nome, ls]) => ({ nome, cor: corAci(corCamada.get(nome) ?? 7), linhas: ls }))
      .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
    return { camadas, bbox, entidades: contagem, ignoradas, pontos: totalPontos };
  }

  /**
   * Descobre como as coordenadas do desenho se relacionam com o UTM.
   * Devolve { tipo, afim, aviso }: afim = [a,b,c,d,e,f] leva (x,y) do desenho a (E,N).
   *   'utm'      X = Leste, Y = Norte: usa direto.
   *   'trocado'  X tem valor de Norte e Y de Leste: mantém a orientação do desenho e leva o
   *              centro para o ponto que os números indicam (precisa conferir/ajustar).
   *   'local'    coordenadas sem relação com UTM: centra no ponto dado (precisa ajustar).
   */
  function georreferenciar(bbox, centroUtm) {
    const eLeste = v => v > 1e5 && v < 9.5e5, eNorte = v => v > 1e6 && v < 1.0001e7;
    const cx = (bbox.xmin + bbox.xmax) / 2, cy = (bbox.ymin + bbox.ymax) / 2;
    if (eLeste(bbox.xmin) && eLeste(bbox.xmax) && eNorte(bbox.ymin) && eNorte(bbox.ymax))
      return { tipo: 'utm', afim: [1, 0, 0, 1, 0, 0], aviso: null };
    if (eNorte(bbox.xmin) && eNorte(bbox.xmax) && eLeste(bbox.ymin) && eLeste(bbox.ymax))
      return { tipo: 'trocado', afim: [1, 0, 0, 1, cy - cx, cx - cy],
        aviso: 'As coordenadas do desenho parecem estar com os eixos trocados (X com valor de Norte e Y de Leste). A planta foi colocada perto do lugar, mas confira com o satélite e use "Ajustar posição".' };
    if (!centroUtm) return { tipo: 'local', afim: [1, 0, 0, 1, 0, 0], aviso: 'O desenho não está em coordenadas UTM e o projeto ainda não tem poços para servir de referência.' };
    return { tipo: 'local', afim: [1, 0, 0, 1, centroUtm.E - cx, centroUtm.N - cy],
      aviso: 'O desenho não está em coordenadas UTM. A planta foi colocada no centro dos poços: use "Ajustar posição" para encaixar no satélite.' };
  }

  /**
   * Compacta o desenho para gravar no banco: coordenadas relativas à origem, em centímetros.
   * dados = { v, nome, zona, sul, origem:[x0,y0], afim, tipo, camadas:[{nome, cor, l:[[x,y,...]]}], ocultas:[] }
   */
  function empacotar(desenho, geo, nome, zona, sul = true) {
    const x0 = Math.floor(desenho.bbox.xmin), y0 = Math.floor(desenho.bbox.ymin);
    return {
      v: 1, nome, zona, sul, tipo: geo.tipo, origem: [x0, y0], afim: geo.afim, ocultas: [],
      camadas: desenho.camadas.map(c => ({ nome: c.nome, cor: c.cor,
        l: c.linhas.map(l => l.map((v, i) => Math.round((v - (i % 2 ? y0 : x0)) * 100) / 100)) }))
    };
  }

  /** Ajuste por pontos de controle: pares = [{de:{E,N}, para:{E,N}}] (1 = só desloca; 2 = desloca e gira). */
  function ajustar(afim, pares, mudarEscala = false) {
    if (!pares.length) return afim;
    let k = [1, 0], t; // k = fator complexo (rotação/escala)
    if (pares.length >= 2) {
      const p = pares[0], q = pares[1];
      const dz = [q.de.E - p.de.E, q.de.N - p.de.N], dw = [q.para.E - p.para.E, q.para.N - p.para.N];
      const d2 = dz[0] * dz[0] + dz[1] * dz[1];
      if (d2 > 1e-6) {
        k = [(dw[0] * dz[0] + dw[1] * dz[1]) / d2, (dw[1] * dz[0] - dw[0] * dz[1]) / d2];
        if (!mudarEscala) { const m = Math.hypot(k[0], k[1]) || 1; k = [k[0] / m, k[1] / m]; }
      }
      // translação pelo ponto médio (divide o erro entre os dois pontos)
      const mz = [(p.de.E + q.de.E) / 2, (p.de.N + q.de.N) / 2], mw = [(p.para.E + q.para.E) / 2, (p.para.N + q.para.N) / 2];
      t = [mw[0] - (k[0] * mz[0] - k[1] * mz[1]), mw[1] - (k[1] * mz[0] + k[0] * mz[1])];
    } else {
      t = [pares[0].para.E - pares[0].de.E, pares[0].para.N - pares[0].de.N];
    }
    return mult([k[0], -k[1], k[1], k[0], t[0], t[1]], afim);
  }

  /** Ponto do desenho (coordenadas gravadas, relativas à origem) -> UTM. */
  function paraUtm(dados, x, y) {
    const [E, N] = aplicar(dados.afim, x + dados.origem[0], y + dados.origem[1]);
    return { E, N };
  }

  return { ler, georreferenciar, empacotar, ajustar, paraUtm, utmParaLatLon, latLonParaUtm, zonaDaLongitude };
})();




/* =====================================================================
   3b. MAPA POTENCIOMÉTRICO (funções puras, sem tela)
   Superfície da carga hidráulica (cota do topo - N.A.), curvas
   equipotenciais e sentido do fluxo (do maior para o menor potencial).
   ===================================================================== */
const WGPot = (() => {

  /** Resolve A.x = b (eliminação de Gauss com pivô). Devolve null se o sistema não tem solução única. */
  function resolver(A, b) {
    const n = b.length, M = A.map((l, i) => [...l, b[i]]);
    for (let c = 0; c < n; c++) {
      let piv = c;
      for (let l = c + 1; l < n; l++) if (Math.abs(M[l][c]) > Math.abs(M[piv][c])) piv = l;
      if (Math.abs(M[piv][c]) < 1e-12) return null;
      [M[c], M[piv]] = [M[piv], M[c]];
      for (let l = c + 1; l < n; l++) {
        const f = M[l][c] / M[c][c];
        if (f) for (let k = c; k <= n; k++) M[l][k] -= f * M[c][k];
      }
    }
    const x = new Array(n);
    for (let l = n - 1; l >= 0; l--) {
      let s = M[l][n];
      for (let k = l + 1; k < n; k++) s -= M[l][k] * x[k];
      x[l] = s / M[l][l];
    }
    return x;
  }

  /** Poços no mesmo ponto (até `raio` m) viram um ponto só, com a MÉDIA das cargas. */
  function agrupar(pts, raio = 5) {
    const grupos = [];
    pts.forEach(p => {
      const g = grupos.find(q => Math.hypot(q.x - p.x, q.y - p.y) <= raio);
      if (g) { g.soma += p.h; g.n++; g.h = g.soma / g.n; } else grupos.push({ x: p.x, y: p.y, h: p.h, soma: p.h, n: 1 });
    });
    return grupos;
  }

  /**
   * Ajusta a superfície potenciométrica: plano regional + spline de placa fina (thin plate spline).
   * É a superfície mais "lisa" que passa pelos poços. Com suav > 0 ela deixa de passar exatamente
   * em cada poço e amortece os valores destoantes (erro de cota, poço em nível diferente).
   * pts: [{x, y, h}] em metros. Devolve { f(x,y), grad(x,y), residuos[] } ou null.
   */
  function ajustar(pts, suav = 0) {
    const n = pts.length;
    if (n < 3) return null;
    const cx = pts.reduce((s, p) => s + p.x, 0) / n, cy = pts.reduce((s, p) => s + p.y, 0) / n;
    const L = Math.max(1, ...pts.map(p => Math.hypot(p.x - cx, p.y - cy)));  // escala: melhora a conta
    const q = pts.map(p => ({ x: (p.x - cx) / L, y: (p.y - cy) / L, h: p.h }));
    const hm = q.reduce((s, p) => s + p.h, 0) / n;
    const phi = r2 => r2 > 0 ? .5 * r2 * Math.log(r2) : 0;                   // r² ln r
    const N = n + 3, A = Array.from({ length: N }, () => new Array(N).fill(0)), b = new Array(N).fill(0);
    for (let i = 0; i < n; i++) {
      for (let j = 0; j < n; j++) A[i][j] = phi((q[i].x - q[j].x) ** 2 + (q[i].y - q[j].y) ** 2);
      A[i][i] += suav;
      A[i][n] = A[n][i] = 1; A[i][n + 1] = A[n + 1][i] = q[i].x; A[i][n + 2] = A[n + 2][i] = q[i].y;
      b[i] = q[i].h - hm;
    }
    let s = resolver(A, b), plano = false;
    if (!s) {                                                               // poços alinhados: só o plano médio
      plano = true;
      s = new Array(N).fill(0);
    }
    const w = s.slice(0, n), a0 = s[n], a1 = s[n + 1], a2 = s[n + 2];
    const f = (x, y) => {
      const u = (x - cx) / L, v = (y - cy) / L;
      let z = hm + a0 + a1 * u + a2 * v;
      for (let i = 0; i < n; i++) z += w[i] * phi((u - q[i].x) ** 2 + (v - q[i].y) ** 2);
      return z;
    };
    /** Gradiente (dh/dx, dh/dy) em m/m. */
    const grad = (x, y) => {
      const u = (x - cx) / L, v = (y - cy) / L;
      let gx = a1, gy = a2;
      for (let i = 0; i < n; i++) {
        const dx = u - q[i].x, dy = v - q[i].y, r2 = dx * dx + dy * dy;
        if (r2 > 0) { const k = w[i] * (Math.log(r2) + 1); gx += k * dx; gy += k * dy; }
      }
      return { x: gx / L, y: gy / L };
    };
    return { f, grad, plano, residuos: pts.map(p => f(p.x, p.y) - p.h) };
  }

  /**
   * Superfície na grade. Só dentro da rede de poços (envoltória + uma folga): fora dela não há dado.
   * amostras: [{lat, lon, h}]. Devolve { g, v, modelo, pts, hmin, hmax } ou null (menos de 3 pontos).
   */
  function superficie(amostras, suav = 0, cel = 2) {
    if (amostras.length < 3) return null;
    const g = WGPluma.grade(amostras, cel, 40000);
    const pts = agrupar(amostras.map(a => ({ ...g.proj.paraXY(a.lat, a.lon), h: a.h })), 5);
    if (pts.length < 3) return null;
    const modelo = ajustar(pts, suav);
    const env = WGPluma.envoltoria(pts);
    const esp = WGPluma.espacamentos(pts).sort((a, b) => a - b);
    const folga = Math.min(15, esp[Math.floor(esp.length / 2)] / 2);
    const hmin = Math.min(...pts.map(p => p.h)), hmax = Math.max(...pts.map(p => p.h));
    const v = new Float64Array(g.nx * g.ny).fill(NaN);
    for (let j = 0; j < g.ny; j++) for (let i = 0; i < g.nx; i++) {
      const x = g.xmin + (i + .5) * g.cel, y = g.ymin + (j + .5) * g.cel;
      if (WGPluma.distanciaFora(x, y, env) > folga) continue;
      v[j * g.nx + i] = Math.min(hmax, Math.max(hmin, modelo.f(x, y)));    // nunca passa do medido
    }
    return { g, v, modelo, pts, env, folga, hmin, hmax };
  }

  /** Intervalo "redondo" entre curvas, para dar de 6 a 12 curvas. */
  function intervaloAuto(hmin, hmax) {
    const faixa = hmax - hmin;
    if (!(faixa > 0)) return 0.1;
    return [0.01, 0.02, 0.05, 0.1, 0.2, 0.25, 0.5, 1, 2, 5, 10, 20, 50].find(p => faixa / p <= 12) || 100;
  }

  /** Níveis das curvas: múltiplos do intervalo dentro da faixa medida. */
  function niveis(hmin, hmax, passo) {
    const l = [];
    for (let k = Math.ceil(hmin / passo - 1e-9); k * passo <= hmax + 1e-9 && l.length < 200; k++) l.push(+(k * passo).toFixed(4));
    return l;
  }

  /** Curva equipotencial de um nível: lista de curvas [[lat,lon], ...], emendadas e suavizadas. */
  function curva(sup, nivel) {
    const { g, v } = sup, segs = [];
    const ponto = (i, j) => ({ x: g.xmin + (i + .5) * g.cel, y: g.ymin + (j + .5) * g.cel });
    for (let j = 0; j < g.ny - 1; j++) for (let i = 0; i < g.nx - 1; i++) {
      const c = [[i, j], [i + 1, j], [i + 1, j + 1], [i, j + 1]];
      const vv = c.map(([a, b]) => v[b * g.nx + a]);
      if (vv.some(Number.isNaN)) continue;
      const caso = vv.reduce((s, x, k) => s | ((x >= nivel ? 1 : 0) << k), 0);
      if (caso === 0 || caso === 15) continue;
      const pp = c.map(([a, b]) => ponto(a, b));
      const e = k => {
        const k2 = (k + 1) % 4, t = (nivel - vv[k]) / (vv[k2] - vv[k]);
        return { x: pp[k].x + t * (pp[k2].x - pp[k].x), y: pp[k].y + t * (pp[k2].y - pp[k].y) };
      };
      const cruzam = [0, 1, 2, 3].filter(k => ((caso >> k) & 1) !== ((caso >> ((k + 1) % 4)) & 1));
      if (cruzam.length === 2) segs.push([e(cruzam[0]), e(cruzam[1])]);
      else if (cruzam.length === 4) { segs.push([e(0), e(1)]); segs.push([e(2), e(3)]); }
    }
    return WGPluma.emendar(segs).filter(c => c.length > 2)
      .map(c => WGPluma.suavizar(c, 2).map(p => { const ll = g.proj.paraLatLon(p.x, p.y); return [ll.lat, ll.lon]; }));
  }

  /**
   * Sentido do fluxo: setas numa malha regular, apontando para onde a carga diminui.
   * Devolve { setas: [{de:[lat,lon], ate:[lat,lon], asa1, asa2, i}], gradMedio, azimute, rumo }.
   * azimute: graus a partir do Norte, sentido horário (direção média do fluxo).
   */
  function fluxo(sup, passo = null) {
    const { g, v, modelo } = sup;
    const larg = g.nx * g.cel, alt = g.ny * g.cel;
    passo = passo || Math.max(8, Math.round(Math.sqrt(larg * alt / 120)));
    const comp = passo * .5, setas = [];
    let sx = 0, sy = 0, si = 0, n = 0;
    const ll = (x, y) => { const p = g.proj.paraLatLon(x, y); return [p.lat, p.lon]; };
    for (let y = g.ymin + passo / 2; y < g.ymin + alt; y += passo) for (let x = g.xmin + passo / 2; x < g.xmin + larg; x += passo) {
      const i = Math.floor((x - g.xmin) / g.cel), j = Math.floor((y - g.ymin) / g.cel);
      if (i < 0 || j < 0 || i >= g.nx || j >= g.ny || Number.isNaN(v[j * g.nx + i])) continue;
      const gr = modelo.grad(x, y), m = Math.hypot(gr.x, gr.y);
      if (!(m > 1e-7)) continue;
      const ux = -gr.x / m, uy = -gr.y / m;                 // fluxo = contra o gradiente
      sx += ux * m; sy += uy * m; si += m; n++;
      const x0 = x - ux * comp / 2, y0 = y - uy * comp / 2, x1 = x + ux * comp / 2, y1 = y + uy * comp / 2;
      const a = comp * .32, cs = Math.cos(2.6), sn = Math.sin(2.6); // asas a ~150° do sentido
      setas.push({
        de: ll(x0, y0), ate: ll(x1, y1), i: m,
        asa1: ll(x1 + a * (ux * cs - uy * sn), y1 + a * (ux * sn + uy * cs)),
        asa2: ll(x1 + a * (ux * cs + uy * sn), y1 + a * (-ux * sn + uy * cs))
      });
    }
    if (!n) return { setas, gradMedio: null, azimute: null, rumo: null };
    const azimute = (Math.atan2(sx, sy) * 180 / Math.PI + 360) % 360;
    const rumo = ['N', 'NE', 'L', 'SE', 'S', 'SO', 'O', 'NO'][Math.round(azimute / 45) % 8];
    return { setas, gradMedio: si / n, azimute, rumo, constancia: Math.hypot(sx, sy) / si };
  }

  return { ajustar, agrupar, superficie, intervaloAuto, niveis, curva, fluxo };
})();

/* =====================================================================
   3c. SEÇÃO GEOLÓGICA (funções puras, sem tela)
   Corte vertical ao longo de uma linha A–A' traçada no mapa: posição de
   cada sondagem na linha, ligação das camadas iguais e escalas.
   ===================================================================== */
const WGSecao = (() => {

  /** Comprimento (m) de uma linha quebrada [{x,y}, ...]. */
  function comprimento(linha) {
    let s = 0;
    for (let i = 1; i < linha.length; i++) s += Math.hypot(linha[i].x - linha[i - 1].x, linha[i].y - linha[i - 1].y);
    return s;
  }

  /** Distância de cada vértice da linha até o início (m). */
  function vertices(linha) {
    const d = [0];
    for (let i = 1; i < linha.length; i++) d.push(d[i - 1] + Math.hypot(linha[i].x - linha[i - 1].x, linha[i].y - linha[i - 1].y));
    return d;
  }

  /**
   * Projeta um ponto na linha: devolve a distância ao longo da linha (dist) e o
   * afastamento perpendicular (afast), no trecho mais próximo.
   */
  function projetar(linha, p) {
    let melhor = null, antes = 0;
    for (let i = 1; i < linha.length; i++) {
      const a = linha[i - 1], b = linha[i], dx = b.x - a.x, dy = b.y - a.y, c = dx * dx + dy * dy;
      const t = c ? Math.max(0, Math.min(1, ((p.x - a.x) * dx + (p.y - a.y) * dy) / c)) : 0;
      const afast = Math.hypot(p.x - a.x - t * dx, p.y - a.y - t * dy);
      if (!melhor || afast < melhor.afast) melhor = { dist: antes + t * Math.sqrt(c), afast };
      antes += Math.sqrt(c);
    }
    return melhor;
  }

  /** "Argila Siltosa " e "argila siltosa" são a mesma camada. */
  function chaveLit(t) {
    return String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  }

  /**
   * Liga as camadas iguais de duas sondagens vizinhas, de cima para baixo e sem cruzar
   * (maior sequência comum de nomes de camada). Devolve pares [índice em A, índice em B].
   */
  function correlacionar(a, b) {
    const ka = a.map(c => chaveLit(c.lit)), kb = b.map(c => chaveLit(c.lit)), n = ka.length, m = kb.length;
    const t = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--)
      t[i][j] = (ka[i] && ka[i] === kb[j]) ? t[i + 1][j + 1] + 1 : Math.max(t[i + 1][j], t[i][j + 1]);
    const pares = [];
    for (let i = 0, j = 0; i < n && j < m;) {
      if (ka[i] && ka[i] === kb[j]) { pares.push([i, j]); i++; j++; }
      else if (t[i + 1][j] >= t[i][j + 1]) i++; else j++;
    }
    return pares;
  }

  /**
   * Tipo principal do material de uma camada ("Aterro de argila siltosa" -> aterro,
   * "Passagem de cascalho" -> cascalho). Serve para ligar camadas parecidas, mesmo com nomes diferentes.
   */
  function classeLit(t) {
    const k = chaveLit(t);
    const grupos = [
      ['piso', /^(piso|concreto|asfalto|pavimento|contrapiso|laje|calcada)/],
      ['aterro', /^(aterro|entulho)/],
      ['argila', /^argil/], ['silte', /^silt/], ['areia', /^(areia|arenito)/],
      ['cascalho', /^(cascalho|pedregulho|brita|seixo)/],
      ['rocha', /^(rocha|saprolito|alteracao|granito|gnaisse|basalto|xisto|filito)/],
      ['organico', /^(turfa|materia|solo organico)/]
    ];
    const palavras = k.split(' ').filter(p => !/^(de|da|do|com|e|passagem|camada|lente|nivel|solo)$/.test(p));
    for (const p of palavras) for (const [nome, re] of grupos) if (re.test(p)) return nome;
    return palavras[0] || '';
  }

  /** Quanto duas camadas se parecem: 0 = materiais diferentes; 1 a 2 = mesmo material, mais alto quanto mais palavras em comum. */
  function semelhanca(a, b) {
    const ca = classeLit(a.lit), cb = classeLit(b.lit);
    if (!ca || ca !== cb) return 0;
    const pa = new Set(chaveLit((a.lit || '') + ' ' + (a.cor || '')).split(' ')), pb = new Set(chaveLit((b.lit || '') + ' ' + (b.cor || '')).split(' '));
    let comuns = 0; pa.forEach(p => { if (pb.has(p)) comuns++; });
    return 1 + comuns / Math.max(1, pa.size + pb.size - comuns);
  }

  /** Camadas de uma sondagem sem buracos: trecho sem descrição vira uma camada "vazia" (mantém as profundidades certas). */
  function completar(camadas) {
    const r = []; let z = 0;
    [...camadas].sort((x, y) => x.de - y.de).forEach((c, i) => {
      if (c.de > z + 0.005) r.push({ de: z, ate: c.de, lit: '', cor: '', hex: '#E4E7E5', vazio: true, orig: -1 });
      if (c.ate > Math.max(z, c.de)) { r.push({ ...c, de: Math.max(z, c.de), orig: i }); z = c.ate; }
    });
    return r;
  }

  /**
   * Coloca as camadas de duas sondagens vizinhas na mesma ordem, de cima para baixo, casando as parecidas
   * que estão em alturas próximas (sem cruzar). Devolve a lista de unidades [{ia, ib}]: ia ou ib = null
   * quando a camada só existe de um lado.
   */
  function alinhar(a, b, cotaA = 0, cotaB = 0) {
    const n = a.length, m = b.length;
    const alcance = Math.max(1, (a[n - 1]?.ate || 0), (b[m - 1]?.ate || 0));
    const nota = (i, j) => {
      if (a[i].vazio || b[j].vazio) return (a[i].vazio && b[j].vazio) ? .5 : 0;
      const s = semelhanca(a[i], b[j]); if (!s) return 0;
      const dz = Math.abs((cotaA - (a[i].de + a[i].ate) / 2) - (cotaB - (b[j].de + b[j].ate) / 2));
      return s * Math.max(.15, 1 - dz / alcance);       // mesma altura vale mais
    };
    const t = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) {
      const s = nota(i, j);
      t[i][j] = Math.max(t[i + 1][j], t[i][j + 1], s > 0 ? t[i + 1][j + 1] + s : 0);
    }
    const un = []; let i = 0, j = 0;
    while (i < n || j < m) {
      const s = (i < n && j < m) ? nota(i, j) : 0;
      if (s > 0 && Math.abs(t[i][j] - (t[i + 1][j + 1] + s)) < 1e-9) { un.push({ ia: i, ib: j }); i++; j++; }
      else if (j >= m || (i < n && t[i + 1][j] >= t[i][j + 1] - 1e-9 && (cotaA - a[i].de) >= (cotaB - (b[j]?.de ?? 0)) - 1e-9)) { un.push({ ia: i, ib: null }); i++; }
      else if (i >= n || t[i][j + 1] >= t[i + 1][j] - 1e-9) { un.push({ ia: null, ib: j }); j++; }
      else { un.push({ ia: i, ib: null }); i++; }
    }
    return un;
  }

  /**
   * Preenche todo o espaço entre duas sondagens vizinhas, sem deixar vazio e sem cruzar camadas:
   *  - camada que existe nas duas: a espessura muda suavemente de uma para a outra;
   *  - camada que só existe numa: afina até acabar no caminho (cunha / lente).
   * Devolve { s: [0..1], faixas: [{ ia, ib, topo: [...], base: [...] }] }, com topo e base em metros
   * abaixo do terreno em cada posição s (0 = sondagem A, 1 = sondagem B). ia/ib são índices nas listas completas.
   */
  function preencher(camA, camB, cotaA = 0, cotaB = 0, passos = 28) {
    const a = completar(camA), b = completar(camB);
    if (!a.length || !b.length) return { s: [], faixas: [], a, b };
    const un = alinhar(a, b, cotaA, cotaB);
    const suave = x => { x = Math.max(0, Math.min(1, x)); return x * x * (3 - 2 * x); };
    let DA = a[a.length - 1].ate, DB = b[b.length - 1].ate;
    const ALC = .62;                                                        // quem não tem par acaba a ~60% do caminho
    // Sondagem mais rasa que a vizinha: as camadas que ficam abaixo do fim dela não "acabam" ali,
    // só não foram alcançadas. Elas passam por baixo, com a mesma espessura (marcadas como "abaixo").
    const fundoA = cotaA - DA, fundoB = cotaB - DB, TOL = .3;
    for (let q = un.length - 1; q >= 0; q--) {
      const u = un[q];
      if (u.ia == null && cotaB - b[u.ib].de <= fundoA + TOL) u.abaixo = 'A';
      else if (u.ib == null && cotaA - a[u.ia].de <= fundoB + TOL) u.abaixo = 'B';
      else break;
    }
    un.forEach(u => { if (u.abaixo === 'A') DA += b[u.ib].ate - b[u.ib].de; else if (u.abaixo === 'B') DB += a[u.ia].ate - a[u.ia].de; });
    const ss = Array.from({ length: passos + 1 }, (_, q) => q / passos);
    const faixas = un.map(u => ({ ...u, topo: [], base: [] }));
    ss.forEach(s => {
      const e = suave(s);
      const esp = un.map(u => {
        const tA = u.ia != null ? a[u.ia].ate - a[u.ia].de : 0, tB = u.ib != null ? b[u.ib].ate - b[u.ib].de : 0;
        if (u.ia != null && u.ib != null) return tA + (tB - tA) * e;
        if (u.abaixo) return tA + tB;                       // passa por baixo da sondagem mais rasa
        return u.ia != null ? tA * (1 - suave(s / ALC)) : tB * suave((s - (1 - ALC)) / ALC);
      });
      const soma = esp.reduce((x, y) => x + y, 0), D = DA + (DB - DA) * e, k = soma > 1e-9 ? D / soma : 0;
      let z = 0;
      esp.forEach((t, q) => { faixas[q].topo.push(z); z += t * k; faixas[q].base.push(z); });
    });
    return { s: ss, faixas, a, b };
  }

  /**
   * Afasta posições que ficariam uma em cima da outra no desenho (poços colados).
   * xs em ordem crescente; devolve novas posições com pelo menos `minimo` entre elas, dentro de [x0, x1].
   */
  function espalhar(xs, minimo, x0, x1) {
    const n = xs.length; if (!n) return [];
    if ((n - 1) * minimo > x1 - x0) minimo = (x1 - x0) / Math.max(1, n - 1);
    const r = xs.slice();
    // blocos de vizinhos apertados são centrados em torno da média das posições reais
    let i = 0;
    const blocos = [];
    while (i < n) { blocos.push({ ini: i, fim: i, soma: xs[i] }); i++; }
    const pos = b => { const k = b.fim - b.ini + 1; return b.soma / k - (k - 1) * minimo / 2; };
    for (let k = 1; k < blocos.length; k++) {
      const a = blocos[k - 1], b = blocos[k];
      if (pos(b) < pos(a) + (a.fim - a.ini + 1) * minimo - 1e-9) {
        a.fim = b.fim; a.soma += b.soma; blocos.splice(k, 1); k = Math.max(0, k - 2);
      }
    }
    blocos.forEach(b => {
      const k = b.fim - b.ini + 1;
      const ini = Math.max(x0, Math.min(x1 - (k - 1) * minimo, pos(b)));
      for (let q = 0; q < k; q++) r[b.ini + q] = ini + q * minimo;
    });
    for (let q = 1; q < n; q++) if (r[q] < r[q - 1] + minimo - 1e-6) r[q] = r[q - 1] + minimo;   // blocos empurrados pela borda
    for (let q = n - 1; q >= 0; q--) { const lim = q === n - 1 ? x1 : r[q + 1] - minimo; if (r[q] > lim) r[q] = lim; }
    return r;
  }

  /** Passo "redondo" para marcar um eixo de tamanho `faixa` com cerca de `alvo` marcas. */
  function passoBonito(faixa, alvo = 6) {
    if (!(faixa > 0)) return 1;
    const bruto = faixa / alvo, pot = Math.pow(10, Math.floor(Math.log10(bruto)));
    return [1, 2, 2.5, 5, 10].map(m => m * pot).find(p => p >= bruto - 1e-12);
  }

  /** Exagero vertical "redondo" para o desenho ficar com a altura desejada. */
  function exageroAuto(larguraPx, comprimentoM, alturaPx, desnivelM) {
    if (!(comprimentoM > 0) || !(desnivelM > 0)) return 1;
    const ideal = (alturaPx / desnivelM) / (larguraPx / comprimentoM);
    return [1, 2, 3, 5, 10, 15, 20, 30, 50, 100].reduce((a, b) => Math.abs(Math.log(b / ideal)) < Math.abs(Math.log(a / ideal)) ? b : a);
  }

  return { comprimento, vertices, projetar, chaveLit, correlacionar, classeLit, semelhanca, completar, alinhar, preencher, espalhar, passoBonito, exageroAuto };
})();

if (typeof module !== 'undefined') module.exports = { WGImport, WGPluma, WGDxf, WGPot, WGSecao };

/* =====================================================================
   4. APLICAÇÃO (tela)
   Login e empresa seguem exatamente o Perfil de Sondagem:
   auth -> profiles.organization_id -> organizations.
   ===================================================================== */
if (typeof window !== 'undefined' && window.document) {
  document.addEventListener('DOMContentLoaded', () => WebGeo.iniciar());
}

const WebGeo = (() => {
  let sb;                       // cliente Supabase
  let mapa, camadaSecao, camadaPocos, camadaRotulos, camadaPluma, camadaPot, camadaPlanta, camadaAjuste, rendPlanta;
  let rotulosDoMapa = [];        // nomes a escrever no mapa: {ll, html, texto, status}
  let graficos = {};
  let usuarioCarregado = null;  // evita recarregar a tela quando o Supabase só renova o token
  const st = {
    user: null, profile: null, org: null,
    projetos: [], projetoId: null, soltas: [], parametros: [], campanhas: [], pocos: [], resultados: [], medicoes: [],
    sondagens: new Map(),       // código do poço (PM-30) -> fichas do Perfil
    parametroId: null, campanhaId: null, redesOcultas: new Set(), importacao: null,
    pontos: [], enquadrarPendente: false,
    perfis: new Map(),            // poco_id -> intervalos da aba Litologia (wg_perfil_poco)
    pluma: { ativa: true, limiar: null, p: 2, cel: 2 }, // limiar null = usa o VI do parâmetro
    calc: null,                   // resultado da última interpolação
    pot: { ativa: false, rede: null, intervalo: 0, suav: 0, setas: true }, // rede null = escolhe sozinho; intervalo 0 = automático
    potCalc: null,                // superfície potenciométrica da campanha selecionada
    secao: { tracando: null, linha: null, faixa: 10, exag: 0, ligar: true }, // seção geológica A–A' (exag 0 = automático)
    secaoCalc: null,              // sondagens que entram no corte
    planta: null, plantaLL: null, plantaErro: null, dxfNovo: null, ajuste: null // planta em DXF do projeto
  };

  const $ = s => document.querySelector(s);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = (n, casas = 2) => n === null || n === undefined ? '-' :
    Number(n).toLocaleString('pt-BR', { maximumFractionDigits: casas });
  const fmtData = d => d ? String(d).slice(0, 10).split('-').reverse().join('/') : '-';
  const codigo = c => String(c || '').trim().toUpperCase().replace(/\s+/g, '');

  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg; t.classList.add('show');
    clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('show'), 3200);
  }

  // ------------------------------------------------------------- telas e login
  function showScreen(qual) {
    $('#auth-screen').hidden = qual !== 'auth';
    $('#pending-org-screen').hidden = qual !== 'pending';
    $('#app').hidden = qual !== 'app';
  }

  function traduzErroAuth(msg) {
    msg = String(msg || '');
    if (/Invalid login credentials/i.test(msg)) return 'E-mail ou senha incorretos.';
    if (/Email not confirmed/i.test(msg)) return 'Confirme seu e-mail antes de entrar (veja sua caixa de entrada).';
    if (/Unable to validate email address/i.test(msg)) return 'Digite um e-mail válido.';
    return 'Não foi possível entrar agora: ' + msg;
  }

  function iniciar() {
    if (!window.supabase) {
      document.body.innerHTML = '<p style="padding:24px">Não consegui carregar o Supabase. Copie o arquivo supabase.min.js do Perfil para esta pasta.</p>';
      return;
    }
    sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
    ligarEventos();
    // Mesmo esquema do Perfil: o listener já entrega o estado atual ao ser registrado
    sb.auth.onAuthStateChange((_evento, sessao) => {
      if (!sessao || !sessao.user) {
        usuarioCarregado = null; st.user = st.profile = st.org = null;
        showScreen('auth');
        return;
      }
      st.user = sessao.user;
      if (usuarioCarregado === sessao.user.id) return; // só renovou o token
      usuarioCarregado = sessao.user.id;
      carregarPerfilEEntrar();
    });
  }

  async function carregarPerfilEEntrar() {
    try {
      const { data: profile, error } = await sb.from('profiles').select('*').eq('id', st.user.id).single();
      if (error) throw error;
      st.profile = profile;
      if (!profile.organization_id) { showScreen('pending'); return; }
      const { data: org } = await sb.from('organizations').select('*').eq('id', profile.organization_id).single();
      st.org = org || { id: profile.organization_id, name: 'Empresa' };
      $('#rail-user-org').textContent = st.org.name || 'Empresa';
      $('#rail-user-email').textContent = st.user.email || '';
      $('#imp-empresa').textContent = st.org.name || 'sua empresa';
      showScreen('app');
      if (!mapa) criarMapa();
      await carregarProjetos();
    } catch (e) {
      console.error(e);
      toast('Não foi possível carregar sua conta agora. Tente novamente em instantes.');
    }
  }

  function ligarEventos() {
    $('#auth-form').addEventListener('submit', async e => {
      e.preventDefault();
      $('#auth-error').hidden = true;
      $('#auth-submit').disabled = true;
      const { error } = await sb.auth.signInWithPassword({ email: $('#auth-email').value.trim(), password: $('#auth-password').value });
      $('#auth-submit').disabled = false;
      if (error) { $('#auth-error').textContent = traduzErroAuth(error.message); $('#auth-error').hidden = false; }
    });
    $('#btn-logout').addEventListener('click', () => sb.auth.signOut());
    $('#pending-logout-btn').addEventListener('click', () => sb.auth.signOut());
    $('#pending-refresh-btn').addEventListener('click', async () => {
      $('#pending-refresh-btn').disabled = true;
      await carregarPerfilEEntrar();
      $('#pending-refresh-btn').disabled = false;
    });

    document.querySelectorAll('[data-view]').forEach(b => b.addEventListener('click', () => trocarView(b.dataset.view)));
    $('#f-projeto').addEventListener('change', e => {
      if (e.target.value === '__novo__') { criarProjeto(); return; }
      st.projetoId = e.target.value || null; lembrarProjeto(st.projetoId);
      st.campanhaId = null; st.parametroId = null;
      carregarProjeto();
    });
    $('#f-parametro').addEventListener('change', e => { st.parametroId = +e.target.value; st.pluma.limiar = null; render(); });
    $('#f-campanha').addEventListener('change', e => { st.campanhaId = +e.target.value; render(); });
    $('#btn-csv').addEventListener('click', exportarCSV);

    // Planta em DXF
    $('#btn-planta-inserir').addEventListener('click', () => $('#arquivo-dxf').click());
    $('#btn-planta-trocar').addEventListener('click', () => $('#arquivo-dxf').click());
    $('#arquivo-dxf').addEventListener('change', e => e.target.files[0] && lerDxf(e.target.files[0]));
    const fecharPlanta = () => { $('#modal-planta').hidden = true; st.dxfNovo = null; };
    $('#planta-cancelar').addEventListener('click', fecharPlanta);
    $('#planta-fechar').addEventListener('click', fecharPlanta);
    $('#planta-confirmar').addEventListener('click', inserirPlanta);
    $('#planta-zona').addEventListener('input', () => st.dxfNovo && mostrarResumoDxf());
    $('#planta-sul').addEventListener('change', () => st.dxfNovo && mostrarResumoDxf());
    $('#f-planta').addEventListener('change', e => { cfgPlanta.setMostrar(e.target.checked); renderPlanta(); });
    $('#f-planta-cor').addEventListener('change', e => { cfgPlanta.setCor(e.target.value); renderPlanta(); });
    const marcarCamadas = todas => {
      cfgPlanta.setOcultas(st.projetoId, todas ? new Set() : new Set(st.planta.dados.camadas.map(c => c.nome)));
      montarPainelPlanta(); renderPlanta();
    };
    $('#btn-camadas-todas').addEventListener('click', () => marcarCamadas(true));
    $('#btn-camadas-nenhuma').addEventListener('click', () => marcarCamadas(false));
    $('#btn-planta-remover').addEventListener('click', removerPlanta);
    $('#btn-planta-ajustar').addEventListener('click', iniciarAjuste);
    $('#ajuste-salvar').addEventListener('click', salvarAjuste);
    $('#ajuste-cancelar').addEventListener('click', () => cancelarAjuste(false));
    $('#ajuste-recomecar').addEventListener('click', () => { cancelarAjuste(false); iniciarAjuste(); });

    // Excluir projeto: só libera o botão quando o nome digitado é igual ao do projeto
    $('#btn-excluir-projeto').addEventListener('click', abrirExcluirProjeto);
    $('#excluir-cancelar').addEventListener('click', fecharExcluirProjeto);
    $('#excluir-fechar').addEventListener('click', fecharExcluirProjeto);
    $('#modal-excluir').addEventListener('click', e => { if (e.target.id === 'modal-excluir') fecharExcluirProjeto(); });
    document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('#modal-excluir').hidden) fecharExcluirProjeto(); });
    $('#excluir-confirma').addEventListener('input', e => {
      const proj = st.projetos.find(p => p.id === st.projetoId);
      $('#excluir-confirmar').disabled = !proj || e.target.value.trim() !== proj.nome.trim();
    });
    $('#excluir-confirmar').addEventListener('click', excluirProjeto);
    $('#f-pluma').addEventListener('change', e => { st.pluma.ativa = e.target.checked; render(); });
    $('#f-idw').addEventListener('change', e => { st.pluma.p = +e.target.value; render(); });
    $('#f-celula').addEventListener('change', e => { st.pluma.cel = +e.target.value; render(); });
    $('#f-pot').addEventListener('change', e => { st.pot.ativa = e.target.checked; render(); });
    $('#f-pot-rede').addEventListener('change', e => { st.pot.rede = e.target.value; st.pot.redeEscolhida = true; render(); });
    $('#f-pot-int').addEventListener('change', e => { st.pot.intervalo = +e.target.value; render(); });
    $('#f-pot-suav').addEventListener('change', e => { st.pot.suav = +e.target.value; render(); });
    $('#f-pot-setas').addEventListener('change', e => { st.pot.setas = e.target.checked; render(); });
    $('#btn-secao-tracar').addEventListener('click', iniciarSecao);
    $('#btn-secao-apagar').addEventListener('click', apagarSecao);
    $('#btn-secao-baixar').addEventListener('click', baixarSecao);
    $('#secao-concluir').addEventListener('click', concluirSecao);
    $('#secao-desfazer').addEventListener('click', desfazerSecao);
    $('#secao-cancelar').addEventListener('click', cancelarSecao);
    $('#f-secao-faixa').addEventListener('change', e => { st.secao.faixa = +e.target.value; render(); });
    $('#f-secao-exag').addEventListener('change', e => { st.secao.exag = +e.target.value; render(); });
    $('#f-secao-ligar').addEventListener('change', e => { st.secao.ligar = e.target.checked; render(); });
    $('#f-limiar').addEventListener('change', e => {
      const v = WGImport.numero(e.target.value);
      st.pluma.limiar = v !== null && v > 0 ? v : null; // vazio = volta para o VI
      render();
    });
    // modo claro/escuro mudou: redesenha os gráficos com as cores certas
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => render());

    // Upload
    const drop = $('#drop');
    $('#arquivo').addEventListener('change', e => e.target.files[0] && lerArquivo(e.target.files[0]));
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('sobre'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('sobre'));
    drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('sobre'); e.dataTransfer.files[0] && lerArquivo(e.dataTransfer.files[0]); });
    $('#btn-gravar').addEventListener('click', gravarImportacao);
    $('#btn-cancelar').addEventListener('click', limparImportacao);
  }

  function trocarView(v) {
    document.querySelectorAll('.rail-nav [data-view]').forEach(b => b.classList.toggle('is-active', b.dataset.view === v));
    document.querySelectorAll('.view').forEach(s => { s.hidden = s.id !== 'view-' + v; });
    if (v === 'visao' && mapa) setTimeout(() => {
      mapa.invalidateSize();
      if (st.enquadrarPendente) enquadrarMapa();
      atualizarRotulos();
      if (st.secao.linha) renderSecao();   // o desenho usa a largura da tela
    }, 0);
  }

  // ------------------------------------------------------------- dados
  /** Busca todas as linhas (o Supabase devolve no máximo 1000 por vez). */
  async function buscarTudo(montarConsulta) {
    const tam = 1000; let ini = 0; const tudo = [];
    for (;;) {
      const { data, error } = await montarConsulta().range(ini, ini + tam - 1);
      if (error) throw error;
      tudo.push(...data);
      if (data.length < tam) return tudo;
      ini += tam;
    }
  }

  // ------------------------------------------------------------- projetos (os mesmos do Perfil)
  const CHAVE_PROJETO = 'webgeo.projeto';
  function lembrarProjeto(id) { try { localStorage.setItem(CHAVE_PROJETO, id || ''); } catch (e) { /* sem armazenamento: tudo bem */ } }
  function projetoLembrado() { try { return localStorage.getItem(CHAVE_PROJETO); } catch (e) { return null; } }

  async function carregarProjetos() {
    try {
      const [projetos, params] = await Promise.all([
        buscarTudo(() => sb.from('projetos').select('id, nome').eq('organization_id', st.org.id).order('nome')),
        buscarTudo(() => sb.from('wg_parametro').select('*').order('nome'))
      ]);
      st.projetos = projetos.sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR', { sensitivity: 'base', numeric: true }));
      st.parametros = params;
      $('#aviso-geral').hidden = true;
      $('#stat-projetos').textContent = projetos.length;
      if (!projetos.some(p => p.id === st.projetoId)) {
        const lembrado = projetoLembrado();
        st.projetoId = projetos.some(p => p.id === lembrado) ? lembrado : (projetos[0]?.id || null);
      }
      montarSeletorProjeto();
      if (!st.projetoId) {
        st.campanhas = []; st.pocos = []; st.resultados = []; st.medicoes = []; st.soltas = [];
        await carregarPlanta();
        carregarSecao();
        render(); trocarView('importar'); return;
      }
      await carregarProjeto();
    } catch (e) { erroGeral(e); }
  }

  function montarSeletorProjeto() {
    $('#f-projeto').innerHTML = (st.projetos.length
      ? st.projetos.map(p => `<option value="${esc(p.id)}">${esc(p.nome)}</option>`).join('')
      : '<option value="">(nenhum projeto)</option>') + '<option value="__novo__">+ Novo projeto…</option>';
    $('#f-projeto').value = st.projetoId || '';
    $('#btn-excluir-projeto').hidden = !st.projetoId;
  }

  /** Cria um projeto (na mesma tabela do Perfil — ele aparece lá também). */
  async function criarProjeto() {
    const nome = (window.prompt('Nome do novo projeto:') || '').trim();
    if (!nome) { montarSeletorProjeto(); return; }
    const { data, error } = await sb.from('projetos')
      .insert({ organization_id: st.org.id, created_by: st.profile?.id || null, nome }).select().single();
    if (error || !data) { console.error(error); toast('Não foi possível criar o projeto agora.'); montarSeletorProjeto(); return; }
    toast('Projeto criado. Ele também aparece no Perfil de Sondagem.');
    st.projetoId = data.id; lembrarProjeto(data.id);
    await carregarProjetos();
  }

  // ------------------------------------------------------------- planta em DXF
  const CORES_PLANTA = { branco: '#ffffff', amarelo: '#ffd60a', preto: '#111111' };
  const cfgPlanta = (() => { // preferências de exibição, guardadas neste navegador
    const ler = (k, pad) => { try { const v = localStorage.getItem(k); return v === null ? pad : v; } catch (e) { return pad; } };
    const gravar = (k, v) => { try { localStorage.setItem(k, v); } catch (e) { /* sem armazenamento */ } };
    return {
      mostrar: () => ler('webgeo.planta.mostrar', '1') === '1', setMostrar: v => gravar('webgeo.planta.mostrar', v ? '1' : '0'),
      cor: () => ler('webgeo.planta.cor', 'branco'), setCor: v => gravar('webgeo.planta.cor', v),
      ocultas: proj => { try { return new Set(JSON.parse(ler('webgeo.planta.ocultas.' + proj, '[]'))); } catch (e) { return new Set(); } },
      setOcultas: (proj, s) => gravar('webgeo.planta.ocultas.' + proj, JSON.stringify([...s]))
    };
  })();

  /** Carrega a planta do projeto (se a tabela ainda não existir no banco, só avisa ao tentar inserir). */
  async function carregarPlanta() {
    st.planta = null; st.plantaLL = null; st.plantaErro = null;
    cancelarAjuste(true);
    if (!st.projetoId) { montarPainelPlanta(); return; }
    const { data, error } = await sb.from('wg_planta').select('*').eq('projeto_id', st.projetoId);
    if (error) st.plantaErro = error.message || String(error);
    else if (data && data.length) st.planta = { nome: data[0].nome, dados: data[0].dados };
    montarPainelPlanta();
  }

  function montarPainelPlanta() {
    const p = st.planta;
    $('#planta-painel').hidden = !p;
    $('#btn-planta-inserir').hidden = !!p || !st.projetoId;
    $('#leg-planta').hidden = !p;
    if (!p) return;
    $('#planta-nome').textContent = p.nome;
    $('#f-planta').checked = cfgPlanta.mostrar();
    $('#f-planta-cor').value = cfgPlanta.cor();
    const ocultas = cfgPlanta.ocultas(st.projetoId);
    $('#planta-camadas-tit').textContent = `Camadas (${p.dados.camadas.length - [...ocultas].filter(n => p.dados.camadas.some(c => c.nome === n)).length} de ${p.dados.camadas.length})`;
    $('#planta-camadas').innerHTML = p.dados.camadas.map(c => `
      <label class="chk"><input type="checkbox" value="${esc(c.nome)}" ${ocultas.has(c.nome) ? '' : 'checked'}><span class="box"></span>
        <i class="amostra" style="background:${esc(c.cor)}"></i>${esc(c.nome)}<small>${c.l.length}</small></label>`).join('');
    $('#planta-camadas').querySelectorAll('input').forEach(i => i.addEventListener('change', () => {
      const s = cfgPlanta.ocultas(st.projetoId);
      i.checked ? s.delete(i.value) : s.add(i.value);
      cfgPlanta.setOcultas(st.projetoId, s); montarPainelPlanta(); renderPlanta();
    }));
  }

  /** Converte as linhas da planta para latitude/longitude (guarda em cache até a posição mudar). */
  function linhasDaPlanta() {
    const d = st.planta.dados, chave = d.afim.join(',') + '|' + d.zona + '|' + d.sul;
    if (st.plantaLL && st.plantaLL.chave === chave) return st.plantaLL;
    const camadas = d.camadas.map(c => ({
      nome: c.nome, cor: c.cor,
      linhas: c.l.map(l => {
        const out = new Array(l.length / 2);
        for (let i = 0; i < l.length; i += 2) {
          const u = WGDxf.paraUtm(d, l[i], l[i + 1]), ll = WGDxf.utmParaLatLon(u.E, u.N, d.zona, d.sul);
          out[i / 2] = [ll.lat, ll.lon];
        }
        return out;
      })
    }));
    st.plantaLL = { chave, camadas };
    return st.plantaLL;
  }

  function renderPlanta() {
    if (!camadaPlanta) return;
    camadaPlanta.clearLayers();
    if (!st.planta || !cfgPlanta.mostrar()) return;
    const ocultas = cfgPlanta.ocultas(st.projetoId), modo = cfgPlanta.cor();
    linhasDaPlanta().camadas.forEach(c => {
      if (ocultas.has(c.nome) || !c.linhas.length) return;
      L.polyline(c.linhas, { renderer: rendPlanta, pane: 'planta', interactive: false, weight: 1.2, opacity: .92,
        color: modo === 'cad' ? c.cor : CORES_PLANTA[modo] || '#ffffff' }).addTo(camadaPlanta);
    });
  }

  function limitesDaPlanta() {
    const pts = [];
    linhasDaPlanta().camadas.forEach(c => c.linhas.forEach(l => { pts.push(l[0], l[l.length - 1]); }));
    return pts.length ? L.latLngBounds(pts) : null;
  }

  // ---- inserir
  async function lerDxf(arquivo) {
    $('#arquivo-dxf').value = '';
    if (st.plantaErro) {
      toast('Falta criar a tabela da planta no banco: rode banco/04_planta_dxf.sql no Supabase.');
      console.error('wg_planta:', st.plantaErro); return;
    }
    $('#modal-planta').hidden = false;
    $('#planta-resumo').hidden = true; $('#planta-confirmar').disabled = true;
    $('#planta-status').hidden = false; $('#planta-status').textContent = 'Lendo ' + arquivo.name + '...';
    try {
      const buf = await arquivo.arrayBuffer();
      let texto;
      try { texto = new TextDecoder('utf-8', { fatal: true }).decode(buf); }
      catch (e) { texto = new TextDecoder('windows-1252').decode(buf); } // DXF antigo (acentos em ANSI)
      const desenho = WGDxf.ler(texto);
      // zona UTM e centro de referência: pelos poços do projeto
      const comCoord = st.pocos.filter(p => p.latitude !== null);
      const lat = comCoord.length ? comCoord.reduce((s, p) => s + +p.latitude, 0) / comCoord.length : null;
      const lon = comCoord.length ? comCoord.reduce((s, p) => s + +p.longitude, 0) / comCoord.length : null;
      const zona = lon !== null ? WGDxf.zonaDaLongitude(lon) : 23;
      st.dxfNovo = { desenho, nome: arquivo.name, lat, lon };
      $('#planta-zona').value = zona; $('#planta-sul').checked = lat === null ? true : lat < 0;
      mostrarResumoDxf();
    } catch (e) {
      console.error(e);
      $('#planta-status').innerHTML = '<span class="txt-acima"><b>Não consegui ler o arquivo:</b></span> ' + esc(e.message);
    }
  }

  function geoDoDxfNovo() {
    const n = st.dxfNovo, zona = parseInt($('#planta-zona').value, 10), sul = $('#planta-sul').checked;
    if (!(zona >= 1 && zona <= 60)) return null;
    const centro = n.lat !== null ? WGDxf.latLonParaUtm(n.lat, n.lon, zona, sul) : null;
    return { zona, sul, geo: WGDxf.georreferenciar(n.desenho.bbox, centro) };
  }

  function mostrarResumoDxf() {
    const n = st.dxfNovo, d = n.desenho, g = geoDoDxfNovo();
    $('#planta-status').hidden = true; $('#planta-resumo').hidden = false;
    const nEnt = Object.values(d.entidades).reduce((a, b) => a + b, 0);
    const ign = Object.entries(d.ignoradas).map(([k, v]) => `${v} ${k}`).join(', ');
    const tipo = { utm: 'UTM (X = Leste, Y = Norte)', trocado: 'UTM com os eixos trocados', local: 'coordenadas locais (não é UTM)' };
    $('#planta-info').innerHTML = [
      ['Arquivo', esc(n.nome)],
      ['Desenho', `${fmt(nEnt, 0)} elementos em ${d.camadas.length} camadas`],
      ['Tamanho', `${fmt(d.bbox.xmax - d.bbox.xmin, 0)} m × ${fmt(d.bbox.ymax - d.bbox.ymin, 0)} m`],
      ['Coordenadas', g ? tipo[g.geo.tipo] : 'informe a zona UTM'],
      ...(ign ? [['Não importado', esc(ign)]] : [])
    ].map(([a, b]) => `<tr><td>${a}</td><td><b>${b}</b></td></tr>`).join('');
    $('#planta-aviso').hidden = !(g && g.geo.aviso);
    $('#planta-aviso').textContent = g?.geo.aviso || '';
    $('#planta-confirmar').disabled = !g || (g.geo.tipo === 'local' && n.lat === null);
  }

  async function inserirPlanta() {
    const n = st.dxfNovo, g = geoDoDxfNovo();
    if (!n || !g) return;
    const dados = WGDxf.empacotar(n.desenho, g.geo, n.nome, g.zona, g.sul);
    if (JSON.stringify(dados).length > 6e6) { $('#planta-aviso').hidden = false; $('#planta-aviso').textContent = 'Desenho grande demais para gravar (mais de 6 MB). No CAD, apague o que não precisa e exporte de novo.'; return; }
    const btn = $('#planta-confirmar'); btn.disabled = true; btn.textContent = 'Gravando...';
    const { error } = await sb.from('wg_planta').upsert({ projeto_id: st.projetoId, nome: n.nome, dados, atualizado_em: new Date().toISOString() });
    btn.textContent = 'Inserir no mapa';
    if (error) { btn.disabled = false; $('#planta-aviso').hidden = false; $('#planta-aviso').textContent = 'O banco recusou: ' + error.message; return; }
    $('#modal-planta').hidden = true;
    st.planta = { nome: n.nome, dados }; st.plantaLL = null; st.dxfNovo = null;
    cfgPlanta.setMostrar(true); cfgPlanta.setOcultas(st.projetoId, new Set());
    montarPainelPlanta(); trocarView('visao'); renderPlanta();
    setTimeout(() => { const b = limitesDaPlanta(); if (b && mapaVisivel()) mapa.fitBounds(b.extend(st.pontos.length ? L.latLngBounds(st.pontos) : b), { padding: [30, 30], maxZoom: 20 }); }, 50);
    toast(g.geo.tipo === 'utm' ? 'Planta inserida. Confira a posição com o satélite.' : 'Planta inserida fora de posição: use "Ajustar posição".');
  }

  async function removerPlanta() {
    if (!st.planta || !window.confirm(`Remover a planta "${st.planta.nome}" deste projeto?\n\nOs poços e resultados não são afetados.`)) return;
    const { error } = await sb.from('wg_planta').delete().eq('projeto_id', st.projetoId);
    if (error) { toast('Não foi possível remover a planta.'); console.error(error); return; }
    st.planta = null; st.plantaLL = null; cancelarAjuste(true);
    montarPainelPlanta(); renderPlanta(); toast('Planta removida.');
  }

  // ---- ajustar posição por pontos de controle
  // 1 ponto desloca a planta; 2 pontos deslocam e giram (a escala do desenho é mantida).
  function iniciarAjuste() {
    if (!st.planta) return;
    cancelarSecao();
    cfgPlanta.setMostrar(true); $('#f-planta').checked = true;
    trocarView('visao');
    st.ajuste = { original: st.planta.dados.afim.slice(), pares: [], de: null };
    camadaAjuste.clearLayers();
    mapa.closePopup();
    mapa.getContainer().classList.add('mapa-ajustando');
    $('#ajuste-barra').hidden = false;
    renderPlanta(); textoAjuste();
    $('#ajuste-barra').scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }

  function textoAjuste() {
    const a = st.ajuste; if (!a) return;
    const n = a.pares.length + 1;
    if (a.pares.length >= 2) { $('#ajuste-passo').textContent = 'Confira o encaixe'; $('#ajuste-dica').textContent = 'Se ficou bom, clique em "Salvar posição". Senão, "Recomeçar".'; }
    else if (!a.de) {
      $('#ajuste-passo').textContent = `Ponto ${n}: clique num canto da PLANTA`;
      $('#ajuste-dica').textContent = n === 1 ? 'Escolha um ponto fácil de achar no satélite (canto de prédio, muro). O clique gruda no vértice mais próximo.'
        : 'Opcional: um 2º ponto, longe do primeiro, corrige também a rotação. Ou clique em "Salvar posição".';
    } else {
      $('#ajuste-passo').textContent = `Ponto ${n}: agora clique onde ele fica no MAPA`;
      $('#ajuste-dica').textContent = 'Clique no lugar correspondente no satélite (ou num poço, se o ponto da planta for um poço).';
    }
    $('#ajuste-salvar').disabled = a.pares.length === 0;
  }

  function cliqueAjuste(ev) {
    const a = st.ajuste; if (!a || a.pares.length >= 2) return;
    const d = st.planta.dados, alvo = WGDxf.latLonParaUtm(ev.latlng.lat, ev.latlng.lng, d.zona, d.sul);
    const pino = (ll, txt, cls) => L.marker(ll, { interactive: false, keyboard: false,
      icon: L.divIcon({ className: '', html: `<div class="pino-ajuste ${cls}">${txt}</div>`, iconSize: [20, 20] }) }).addTo(camadaAjuste);
    if (!a.de) {
      // gruda no vértice da planta mais próximo (até 15 pixels); senão usa o ponto clicado
      const mPorPx = 40075016 * Math.cos(ev.latlng.lat * Math.PI / 180) / (256 * Math.pow(2, mapa.getZoom()));
      let melhor = null, dist = 15 * mPorPx;
      const ocultas = cfgPlanta.ocultas(st.projetoId);
      d.camadas.forEach(c => { if (ocultas.has(c.nome)) return; c.l.forEach(l => {
        for (let i = 0; i < l.length; i += 2) {
          const u = WGDxf.paraUtm(d, l[i], l[i + 1]), dd = Math.hypot(u.E - alvo.E, u.N - alvo.N);
          if (dd < dist) { dist = dd; melhor = u; }
        }
      }); });
      const atual = melhor || alvo; // posição do ponto com a planta onde ela está AGORA
      // o mesmo ponto na posição ORIGINAL (antes deste ajuste): desfaz o ajuste parcial já aplicado
      a.de = { original: desfazer(a, atual), atual };
      const ll = WGDxf.utmParaLatLon(atual.E, atual.N, d.zona, d.sul);
      a.pinoDe = pino([ll.lat, ll.lon], a.pares.length + 1, 'de');
    } else {
      a.pares.push({ de: a.de.original, para: alvo });
      a.de = null;
      if (a.pinoDe) { camadaAjuste.removeLayer(a.pinoDe); a.pinoDe = null; }
      pino(ev.latlng, a.pares.length, '');
      d.afim = WGDxf.ajustar(a.original, a.pares);
      st.plantaLL = null; renderPlanta();
    }
    textoAjuste();
  }

  /** Leva um ponto UTM da posição atual da planta de volta para a posição original (antes do ajuste em curso). */
  function desfazer(a, p) {
    const atual = st.planta.dados.afim, o = a.original;
    // inverte "atual" para achar o ponto no desenho e reaplica "original"
    const det = atual[0] * atual[3] - atual[1] * atual[2];
    const x = (atual[3] * (p.E - atual[4]) - atual[1] * (p.N - atual[5])) / det, y = (-atual[2] * (p.E - atual[4]) + atual[0] * (p.N - atual[5])) / det;
    return { E: o[0] * x + o[1] * y + o[4], N: o[2] * x + o[3] * y + o[5] };
  }

  function cancelarAjuste(silencioso) {
    if (!st.ajuste) return;
    if (st.planta) { st.planta.dados.afim = st.ajuste.original; st.plantaLL = null; }
    st.ajuste = null;
    if (camadaAjuste) camadaAjuste.clearLayers();
    if (mapa) mapa.getContainer().classList.remove('mapa-ajustando');
    $('#ajuste-barra').hidden = true;
    if (!silencioso) renderPlanta();
  }

  async function salvarAjuste() {
    const a = st.ajuste; if (!a || !a.pares.length) return;
    const btn = $('#ajuste-salvar'); btn.disabled = true; btn.textContent = 'Salvando...';
    const dados = st.planta.dados;
    dados.tipo = 'ajustado';
    const { error } = await sb.from('wg_planta').upsert({ projeto_id: st.projetoId, nome: st.planta.nome, dados, atualizado_em: new Date().toISOString() });
    btn.textContent = 'Salvar posição';
    if (error) { btn.disabled = false; toast('Não foi possível salvar a posição: ' + error.message); return; }
    st.ajuste = null; camadaAjuste.clearLayers();
    mapa.getContainer().classList.remove('mapa-ajustando');
    $('#ajuste-barra').hidden = true;
    toast('Posição da planta salva.');
  }

  // ------------------------------------------------------------- excluir projeto
  /** Abre a confirmação: mostra o que será apagado e exige digitar o nome do projeto. */
  function abrirExcluirProjeto() {
    const proj = st.projetos.find(p => p.id === st.projetoId);
    if (!proj) return;
    const pocos = st.pocos.filter(p => !p.virtual).length, n = (x, um, varios) => `${x} ${x === 1 ? um : varios}`;
    const itens = [];
    if (pocos) itens.push(`${n(pocos, 'poço', 'poços')} cadastrados no WebGeo`);
    if (st.campanhas.length) itens.push(`${n(st.campanhas.length, 'campanha', 'campanhas')} e ${n(st.resultados.length, 'resultado', 'resultados')} de laboratório`);
    if (st.medicoes.length) itens.push(`${n(st.medicoes.length, 'medição', 'medições')} de nível d'água`);
    if (st.planta) itens.push(`A planta "${st.planta.nome}"`);
    if (!itens.length) itens.push('Este projeto não tem dados no WebGeo.');
    if (st.totalFichas) itens.push(`${n(st.totalFichas, 'ficha do Perfil fica', 'fichas do Perfil ficam')} sem projeto (não ${st.totalFichas === 1 ? 'é apagada' : 'são apagadas'})`);
    $('#excluir-nome').textContent = proj.nome;
    $('#excluir-lista').innerHTML = itens.map(i => `<li>${esc(i)}</li>`).join('');
    $('#excluir-confirma').value = '';
    $('#excluir-confirmar').disabled = true;
    $('#excluir-erro').hidden = true;
    $('#modal-excluir').hidden = false;
    $('#excluir-confirma').focus();
  }
  function fecharExcluirProjeto() { $('#modal-excluir').hidden = true; }

  async function excluirProjeto() {
    const proj = st.projetos.find(p => p.id === st.projetoId);
    if (!proj || $('#excluir-confirma').value.trim() !== proj.nome.trim()) return;
    const btn = $('#excluir-confirmar'); btn.disabled = true; btn.textContent = 'Excluindo...';
    // Apagar o projeto apaga junto (no banco) os poços, campanhas e resultados dele.
    // .select() devolve a linha apagada: se voltar vazio, o banco não deixou apagar.
    const { data, error } = await sb.from('projetos').delete().eq('id', proj.id).select();
    btn.textContent = 'Excluir projeto';
    if (error || !data || !data.length) {
      console.error(error);
      $('#excluir-erro').textContent = 'Não foi possível excluir: ' + (error?.message || 'o banco não permitiu (projeto de outra empresa ou já excluído).');
      $('#excluir-erro').hidden = false; btn.disabled = false;
      return;
    }
    fecharExcluirProjeto();
    toast(`Projeto "${proj.nome}" excluído.`);
    st.projetoId = null; st.campanhaId = null; st.parametroId = null; st.popupAberto = null; lembrarProjeto('');
    await carregarProjetos();
  }

  /**
   * Carrega tudo de um projeto: poços, campanhas, resultados e as FICHAS do
   * Perfil de Sondagem do mesmo projeto. Toda ficha com coordenada aparece no
   * mapa: se tem "Poço de monitoramento nº", vira/completa o poço com esse código;
   * se não tem, aparece como ponto de sondagem.
   */
  async function carregarProjeto() {
    const id = st.projetoId;
    try {
      const [campanhas, pocos, resultados, fichas] = await Promise.all([
        buscarTudo(() => sb.from('wg_campanha').select('*').eq('projeto_id', id).order('data_inicio')),
        buscarTudo(() => sb.from('wg_poco').select('*').eq('projeto_id', id).order('codigo')),
        buscarTudo(() => sb.from('wg_vw_resultado').select('*').eq('projeto_id', id).order('id')),
        buscarTudo(() => sb.from('sondagens').select('id, sondagem_no, poco_no, obra, updated_at, data')
          .eq('projeto_id', id).order('updated_at')).catch(e => { console.warn('Sem acesso às fichas do Perfil:', e.message || e); return []; })
      ]);
      st.campanhas = campanhas; st.resultados = resultados;
      st.medicoes = campanhas.length
        ? await buscarTudo(() => sb.from('wg_medicao_campo').select('*').in('campanha_id', campanhas.map(c => c.id)).order('id'))
        : [];
      st.perfis = new Map();
      for (let k = 0; k < pocos.length; k += 200) {
        const ids = pocos.slice(k, k + 200).map(p => p.id);
        (await buscarTudo(() => sb.from('wg_perfil_poco').select('*').in('poco_id', ids).order('de_m')))
          .forEach(r => { if (!st.perfis.has(r.poco_id)) st.perfis.set(r.poco_id, []); st.perfis.get(r.poco_id).push(r); });
      }
      juntarFichas(pocos, fichas);
      await carregarPlanta();
        carregarSecao();

      // Parâmetros com resultado neste projeto; começa pelo que tem mais poços acima do VI
      const comDados = new Set(resultados.map(r => r.parametro_id));
      const params = st.parametros.filter(p => comDados.has(p.id));
      $('#f-parametro').innerHTML = params.length
        ? params.map(p => `<option value="${p.id}">${esc(p.nome)}</option>`).join('')
        : '<option value="">(sem resultados)</option>';
      if (!params.some(p => p.id === st.parametroId)) {
        const acima = {}; resultados.forEach(r => { if (r.acima_vi) acima[r.parametro_id] = (acima[r.parametro_id] || 0) + 1; });
        st.parametroId = params.length ? [...params].sort((a, b) => (acima[b.id] || 0) - (acima[a.id] || 0))[0].id : null;
      }
      $('#f-parametro').value = st.parametroId ?? '';
      $('#f-campanha').innerHTML = campanhas.length
        ? [...campanhas].reverse().map(c => `<option value="${c.id}">${esc(c.codigo)} · ${fmtData(c.data_inicio)}</option>`).join('')
        : '<option value="">(sem campanhas)</option>';
      if (!campanhas.some(c => c.id === st.campanhaId)) st.campanhaId = campanhas.length ? campanhas[campanhas.length - 1].id : null;
      $('#f-campanha').value = st.campanhaId ?? '';

      montarFiltroRede();
      montarFiltroPot();
      render(true);
    } catch (e) { erroGeral(e); }
  }

  /** Coordenada gravada na ficha do Perfil (data.local.latitude/longitude, já convertida do UTM). */
  function coordDaFicha(f) {
    const lat = WGImport.numero(f.data?.local?.latitude), lon = WGImport.numero(f.data?.local?.longitude);
    if (lat === null || lon === null || Math.abs(lat) > 90 || Math.abs(lon) > 180 || (lat === 0 && lon === 0)) return null;
    return { lat, lon };
  }

  /**
   * Junta os poços do WebGeo com as fichas do Perfil do mesmo projeto:
   *  - poço sem coordenada + ficha com coordenada  -> usa a coordenada da ficha
   *  - ficha com "Poço nº" que não existe no WebGeo -> vira um poço (ainda sem resultados)
   *  - ficha sem "Poço nº"                          -> ponto de sondagem no mapa
   */
  function juntarFichas(pocos, fichas) {
    st.sondagens = new Map(); // código do poço -> fichas
    st.soltas = [];           // sondagens sem poço, com coordenada
    st.totalFichas = fichas.length;
    const porCodigo = new Map(pocos.map(p => [codigo(p.codigo), p]));
    fichas.forEach(f => {
      const k = codigo(f.poco_no || f.data?.meta?.pocoNo);
      const c = coordDaFicha(f);
      if (!k) { if (c) st.soltas.push({ ...f, latitude: c.lat, longitude: c.lon }); return; }
      if (!st.sondagens.has(k)) st.sondagens.set(k, []);
      st.sondagens.get(k).push(f);
      const p = porCodigo.get(k);
      if (p) {
        if (p.latitude === null && c) { p.latitude = c.lat; p.longitude = c.lon; p.coordFonte = `ficha ${f.sondagem_no || ''} do Perfil`; }
      } else {
        const novo = { id: 'ficha-' + f.id, virtual: true, codigo: f.poco_no || f.data?.meta?.pocoNo, rede: redeDe(k),
          latitude: c ? c.lat : null, longitude: c ? c.lon : null, situacao: 'ATIVO', cota_topo: null,
          coordFonte: c ? `ficha ${f.sondagem_no || ''} do Perfil` : null };
        porCodigo.set(k, novo);
        pocos.push(novo);
      }
    });
    pocos.forEach(p => { if (p.latitude !== null && !p.coordFonte) p.coordFonte = 'planilha'; });
    st.pocos = pocos.sort((a, b) => String(a.codigo).localeCompare(String(b.codigo), 'pt-BR', { numeric: true }));
  }
  function redeDe(cod) { return cod.includes('-') ? cod.slice(0, cod.indexOf('-')) : cod.replace(/[0-9].*$/, ''); }

  function montarFiltroRede() {
    const redes = [...new Set(st.pocos.map(p => p.rede || '-'))].sort();
    if (st.soltas.length) redes.push('Sondagens');
    $('#f-redes').innerHTML = redes.map(r => `
      <label class="chk"><input type="checkbox" value="${esc(r)}" ${st.redesOcultas.has(r) ? '' : 'checked'}><span class="box"></span>${esc(r)}</label>`).join('');
    $('#f-redes').querySelectorAll('input').forEach(i => i.addEventListener('change', () => {
      i.checked ? st.redesOcultas.delete(i.value) : st.redesOcultas.add(i.value); render();
    }));
  }

  function erroGeral(e) {
    console.error(e);
    $('#aviso-geral').textContent = 'Erro ao falar com o banco: ' + (e.message || e) + '. Confira se os scripts SQL do WebGeo foram rodados no Supabase.';
    $('#aviso-geral').hidden = false;
  }

  // ------------------------------------------------------------- render
  function selecao() {
    const visivel = p => !st.redesOcultas.has(p.rede || '-');
    const res = st.resultados.filter(r => r.parametro_id === st.parametroId && r.campanha_id === st.campanhaId);
    const porPoco = new Map(res.map(r => [r.poco_id, r]));
    return { porPoco, pocos: st.pocos.filter(visivel) };
  }

  function status(r) {
    if (!r) return 'sem';
    if (r.acima_vi) return 'acima';
    if (r.menor_que_lq) return 'lq';
    return 'abaixo';
  }
  const COR_STATUS = { acima: '#D7301F', abaixo: '#F29E2E', lq: '#2B7BBA', sem: '#9AA5A8' };
  const ROTULO_STATUS = { acima: 'Acima do VI', abaixo: 'Abaixo do VI', lq: 'Abaixo do LQ', sem: 'Sem resultado' };

  function textoValor(r) {
    if (!r) return '-';
    if (r.menor_que_lq) return r.lq !== null ? `< ${fmt(r.lq, 4)}` : '< LQ';
    return fmt(r.valor, 4);
  }

  function render(enquadrar = false) {
    if (!sb) return;
    const projeto = st.projetos.find(p => p.id === st.projetoId);
    const param = st.parametros.find(p => p.id === st.parametroId);
    const camp = st.campanhas.find(c => c.id === st.campanhaId);
    $('#tb-sitio').textContent = projeto ? projeto.nome : 'Nenhum projeto — crie um no menu ou importe uma planilha';
    $('#tb-campanha').textContent = camp ? camp.codigo : '—';
    $('#tb-parametro').hidden = !param;
    $('#tb-parametro').textContent = param ? param.nome : '';
    $('#stat-pocos').textContent = st.pocos.length;
    $('#stat-campanhas').textContent = st.campanhas.length;

    const limiar = limiarAtual(param);
    $('#f-limiar').value = limiar != null ? fmt(limiar, 4) : '';
    $('#f-limiar-un').textContent = param?.unidade || 'µg/L';
    st.calc = calcularPlumas(param, limiar);
    st.potCalc = calcularPot();

    renderKPIs(param);
    renderMapa(param, enquadrar);
    renderPluma(param);
    renderPot();
    renderSecao();
    renderPlanta();
    renderEvolucao(param);
    renderTabela(param);
  }

  function renderKPIs(param) {
    const { pocos, porPoco } = selecao();
    const resultados = pocos.map(p => porPoco.get(p.id)).filter(Boolean);
    const acima = resultados.filter(r => r.acima_vi);
    const lq = resultados.filter(r => r.menor_que_lq);
    const max = resultados.filter(r => !r.menor_que_lq).sort((a, b) => b.valor - a.valor)[0];
    const un = param?.unidade || 'µg/L';

    $('#kpi-analisados').textContent = resultados.length;
    $('#kpi-analisados-sub').textContent = `de ${pocos.length} poços · ${pocos.filter(p => p.latitude === null).length} sem coordenadas`;
    $('#kpi-acima').textContent = acima.length;
    $('#kpi-acima-sub').textContent = param?.valor_orientador != null ? `VI ${fmt(param.valor_orientador)} ${un}` : 'parâmetro sem valor orientador';
    $('#kpi-max').textContent = max ? fmt(max.valor) : '—';
    $('#kpi-max-un').textContent = max ? un : '';
    $('#kpi-max-sub').textContent = max ? `no ${max.poco}` : 'nenhum valor quantificado';
    $('#kpi-lq').textContent = lq.length;
    $('#kpi-lq-sub').textContent = resultados.length ? `${Math.round(100 * lq.length / resultados.length)}% dos analisados` : '';
  }

  // ------------------------------------------------------------- mapa
  function criarMapa() {
    mapa = L.map('mapa').setView([-23.471, -46.576], 17);
    L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
      maxZoom: 21, maxNativeZoom: 19, attribution: 'Imagem © Esri'
    }).addTo(mapa);
    L.control.scale({ imperial: false, position: 'bottomright' }).addTo(mapa);
    // a pluma fica por baixo dos poços: imagem (350) < linha do limiar (380) < poços (400)
    mapa.createPane('pluma').style.zIndex = 350;
    mapa.createPane('planta').style.zIndex = 360;      // planta DXF: acima da pluma, abaixo da linha do limiar
    mapa.createPane('plumaLinha').style.zIndex = 380;
    mapa.createPane('pot').style.zIndex = 385;         // curvas equipotenciais e setas de fluxo
    mapa.createPane('secao').style.zIndex = 390;       // linha da seção geológica
    rendPlanta = L.canvas({ pane: 'planta', padding: .5 }); // canvas: aguenta milhares de linhas sem pesar
    camadaPlanta = L.layerGroup().addTo(mapa);
    camadaAjuste = L.layerGroup().addTo(mapa);
    mapa.on('click', cliqueAjuste);
    mapa.on('popupopen', () => { if (st.ajuste || st.secao.tracando) setTimeout(() => mapa.closePopup(), 0); }); // ajustando a planta ou traçando a seção: sem popups
    mapa.on('click', cliqueSecao);
    mapa.on('dblclick', () => { if (st.secao.tracando) concluirSecao(); });
    camadaPluma = L.layerGroup().addTo(mapa);
    camadaPot = L.layerGroup().addTo(mapa);
    camadaSecao = L.layerGroup().addTo(mapa);
    camadaPocos = L.layerGroup().addTo(mapa);
    camadaRotulos = L.layerGroup().addTo(mapa);
    mapa.on('zoomend', atualizarRotulos);
    atualizarRotulos();
  }

  /**
   * Escreve os nomes dos poços sem deixar um em cima do outro:
   *  - poços colados na tela (PM-26, PMN-26A, PMN-26B) viram um bloco só, um nome por linha;
   *  - cada bloco procura um lado livre do ponto (direita, esquerda, cima, baixo e diagonais).
   * Abaixo do zoom 19 só os poços acima do VI mostram o nome. Refaz a cada mudança de zoom.
   */
  function atualizarRotulos() {
    if (!mapa || !camadaRotulos) return;
    camadaRotulos.clearLayers();
    if (!mapaVisivel()) return;
    const perto = mapa.getZoom() >= 19;
    const todos = rotulosDoMapa.map(r => ({ ...r, pt: mapa.latLngToLayerPoint(r.ll) }));
    const itens = todos.filter(r => perto || r.status === 'acima');
    // 1) junta os que estão colados na tela (até 14 px, em cadeia)
    const grupos = [];
    itens.forEach(it => {
      const g = grupos.find(g => g.itens.some(o => Math.hypot(o.pt.x - it.pt.x, o.pt.y - it.pt.y) <= 14));
      if (g) g.itens.push(it); else grupos.push({ itens: [it] });
    });
    const LARG_LETRA = 6.7, ALT_LINHA = 14, FOLGA = 9;
    grupos.forEach(g => {
      // de cima para baixo, como os pontos estão na tela (as linhas de chamada não se cruzam); empate: ordem do nome
      g.itens.sort((a, b) => (Math.round(a.pt.y / 4) - Math.round(b.pt.y / 4)) || a.texto.localeCompare(b.texto, 'pt-BR', { numeric: true }));
      g.x = g.itens.reduce((s, o) => s + o.pt.x, 0) / g.itens.length;
      g.y = g.itens.reduce((s, o) => s + o.pt.y, 0) / g.itens.length;
      g.raio = Math.max(...g.itens.map(o => Math.hypot(o.pt.x - g.x, o.pt.y - g.y))) + FOLGA + (g.itens.length > 1 ? 7 : 0); // bloco fica um pouco afastado: cabe a linha de chamada
      g.w = Math.max(...g.itens.map(o => o.largura * LARG_LETRA)) + 4;
      g.h = g.itens.length * ALT_LINHA;
      g.acima = g.itens.some(o => o.status === 'acima');
    });
    // 2) coloca primeiro os acima do VI e os blocos maiores; cada um escolhe o lado com menos sobreposição
    const ocupado = todos.map(o => ({ x0: o.pt.x - 7, y0: o.pt.y - 7, x1: o.pt.x + 7, y1: o.pt.y + 7 })); // os pontos dos poços
    const sobra = (a, b) => Math.max(0, Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0)) * Math.max(0, Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0));
    grupos.sort((a, b) => (b.acima - a.acima) || (b.itens.length - a.itens.length) || (a.y - b.y) || (a.x - b.x));
    grupos.forEach(g => {
      const r = g.raio, w = g.w, h = g.h;
      const lados = [                       // [x do canto esquerdo, y do topo, alinhamento do texto]
        [r, -h / 2, 'left'], [-r - w, -h / 2, 'right'], [-w / 2, -r - h, 'center'], [-w / 2, r, 'center'],
        [r * .7, -r * .7 - h, 'left'], [r * .7, r * .7, 'left'], [-r * .7 - w, -r * .7 - h, 'right'], [-r * .7 - w, r * .7, 'right']
      ];
      let melhor = null;
      for (const [dx, dy, alinha] of lados) {
        const cx = { x0: g.x + dx, y0: g.y + dy, x1: g.x + dx + w, y1: g.y + dy + h };
        const custo = ocupado.reduce((s, o) => s + sobra(cx, o), 0);
        if (!melhor || custo < melhor.custo - .5) melhor = { cx, custo, dx, dy, alinha };
        if (custo === 0) break;
      }
      ocupado.push(melhor.cx);
      // linha de chamada: liga cada nome ao seu poço (da borda do ponto até a ponta do texto mais próxima)
      let linhas = '';
      g.itens.forEach((o, i) => {
        const px = o.pt.x - g.x, py = o.pt.y - g.y, larg = o.largura * LARG_LETRA;
        const ini = melhor.alinha === 'left' ? melhor.dx : melhor.alinha === 'right' ? melhor.dx + w - larg : melhor.dx + (w - larg) / 2;
        const ty = melhor.dy + i * ALT_LINHA + ALT_LINHA / 2;
        const tx = Math.abs(px - (ini - 2)) <= Math.abs(px - (ini + larg + 2)) ? ini - 2 : ini + larg + 2;
        const d = Math.hypot(tx - px, ty - py), borda = o.status === 'acima' ? 9 : 7;
        if (d < borda + 4) return;                         // nome já encostado no ponto: não precisa de linha
        const x0 = px + (tx - px) * borda / d, y0 = py + (ty - py) * borda / d;
        const c = `x1="${x0.toFixed(1)}" y1="${y0.toFixed(1)}" x2="${tx.toFixed(1)}" y2="${ty.toFixed(1)}"`;
        linhas += `<line class="chamada-fundo" ${c}/><line class="chamada" ${c}/>`;
      });
      const svg = linhas ? `<svg class="rotulos-linhas" width="1" height="1">${linhas}</svg>` : '';
      const html = svg + `<div class="rotulos-bloco" style="left:${melhor.dx.toFixed(1)}px;top:${melhor.dy.toFixed(1)}px;width:${w.toFixed(1)}px;text-align:${melhor.alinha}">`
        + g.itens.map(o => `<div class="rotulo-pm rotulo-${o.status}">${o.html}</div>`).join('') + '</div>';
      L.marker(mapa.layerPointToLatLng([g.x, g.y]), { icon: L.divIcon({ className: 'rotulos-ancora', html, iconSize: [0, 0] }), interactive: false, keyboard: false })
        .addTo(camadaRotulos);
    });
  }

  function renderMapa(param, enquadrar) {
    if (!mapa) return;
    const aberto = st.popupAberto;   // se havia um popup aberto, reabre depois de redesenhar
    st.redesenhando = true;
    camadaPocos.clearLayers();
    st.redesenhando = false;
    const marcadores = new Map();
    const { pocos, porPoco } = selecao();
    const pts = [];
    rotulosDoMapa = [];
    const ordem = { sem: 0, lq: 1, abaixo: 2, acima: 3 }; // os acima do VI ficam por cima
    const carga = st.potCalc?.porPoco || new Map();       // potenciométrico ligado: carga ao lado do nome
    pocos.filter(p => p.latitude !== null)
      .map(p => ({ p, r: porPoco.get(p.id) }))
      .sort((a, b) => ordem[status(a.r)] - ordem[status(b.r)])
      .forEach(({ p, r }) => {
        const s = status(r);
        const ll = [+p.latitude, +p.longitude];
        pts.push(ll);
        if (st.sondagens.has(codigo(p.codigo))) {
          // anel tracejado = tem ficha de sondagem no Perfil
          L.circleMarker(ll, { radius: 12, color: '#5FBE9C', weight: 2, dashArray: '3 3', fill: false, interactive: false }).addTo(camadaPocos);
        }
        const m = L.circleMarker(ll, { radius: s === 'acima' ? 8 : 6, weight: 2, color: '#fff', fillColor: COR_STATUS[s], fillOpacity: 1, className: 'pm pm-' + s })
          .bindPopup(() => popupPoco(p, param), { maxWidth: 320 })
          .on('click', cliqueSecao)   // o clique num poço não chega ao mapa (o popup segura): avisa a seção daqui
          .on('popupopen', () => { st.popupAberto = p.id; })
          .on('popupclose', () => { if (st.popupAberto === p.id && !st.redesenhando) st.popupAberto = null; })
          .addTo(camadaPocos);
        marcadores.set(p.id, m);
        const h = carga.has(p.id) ? carga.get(p.id).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '';
        rotulosDoMapa.push({ ll, status: s, texto: p.codigo, largura: p.codigo.length + (h ? h.length + 2 : 0),
          html: esc(p.codigo) + (h ? ` <span class="rotulo-carga">${h}</span>` : '') });
      });
    // fichas do Perfil sem poço: quadrado verde com o nº da sondagem
    if (!st.redesOcultas.has('Sondagens')) st.soltas.forEach(f => {
      const ll = [f.latitude, f.longitude];
      pts.push(ll);
      const m = L.marker(ll, { icon: L.divIcon({ className: '', html: '<div class="ponto-sondagem"></div>', iconSize: [12, 12] }) })
        .bindPopup(() => popupSondagem(f), { maxWidth: 320 })
        .on('click', cliqueSecao)
        .on('popupopen', () => { st.popupAberto = 'ficha:' + f.id; })
        .on('popupclose', () => { if (st.popupAberto === 'ficha:' + f.id && !st.redesenhando) st.popupAberto = null; })
        .addTo(camadaPocos);
      marcadores.set('ficha:' + f.id, m);
      const nome = f.sondagem_no || 'Sondagem';
      rotulosDoMapa.push({ ll, status: 'sond', texto: nome, largura: nome.length, html: esc(nome) });
    });
    st.pontos = pts;
    // reabre o popup que estava aberto — só com o mapa visível (escondido, o Leaflet erra a largura)
    if (aberto && marcadores.has(aberto) && mapaVisivel()) marcadores.get(aberto).openPopup();
    else if (aberto) st.popupAberto = null;
    if (enquadrar) enquadrarMapa();
    atualizarRotulos();
    const semCoord = pocos.filter(p => p.latitude === null).length;
    $('#mapa-nota').textContent = semCoord ? `${semCoord} poço(s) sem coordenadas não aparecem no mapa.` : '';
  }

  /** O mapa está na tela? (getSize() do Leaflet guarda o tamanho antigo quando a aba é escondida) */
  function mapaVisivel() { return !!(mapa && mapa.getContainer().clientWidth); }

  /** Ajusta o zoom para mostrar todos os poços; se o mapa estiver escondido, faz quando ele aparecer. */
  function enquadrarMapa() {
    if (!st.pontos.length) return;
    if (!mapaVisivel()) { st.enquadrarPendente = true; return; }
    st.enquadrarPendente = false;
    mapa.fitBounds(st.pontos, { padding: [30, 30], maxZoom: 19 });
  }

  function popupPoco(p, param) {
    const hist = st.resultados
      .filter(r => r.poco_id === p.id && r.parametro_id === st.parametroId)
      .sort((a, b) => String(a.data_inicio).localeCompare(String(b.data_inicio)));
    const atual = hist.find(r => r.campanha_id === st.campanhaId);
    const med = st.medicoes.find(m => m.poco_id === p.id && m.campanha_id === st.campanhaId);
    const fichas = st.sondagens.get(codigo(p.codigo)) || [];
    const s = status(atual);
    return `
      <div class="pop">
        <div class="pop-tit">${esc(p.codigo)} <span class="st-tag st-${s}">${ROTULO_STATUS[s]}</span></div>
        ${param ? `<div class="pop-valor">${esc(param.nome)}: <b>${textoValor(atual)}</b> ${atual ? esc(atual.unidade) : ''}</div>` : ''}
        ${param?.valor_orientador != null ? `<div class="pop-mini">VI CETESB: ${fmt(param.valor_orientador)} ${esc(param.unidade)}</div>` : ''}
        ${med ? `<div class="pop-mini">N.A.: ${fmt(med.nivel_agua)} m${med.carga_hidraulica != null ? ` · carga ${fmt(med.carga_hidraulica, 3)} m` : ''}${med.fase_livre ? ' · <b class="txt-acima">FASE LIVRE</b>' : ''}</div>` : ''}
        ${hist.length > 1 ? `<table class="pop-hist"><tr><th>Campanha</th><th>Valor</th></tr>
          ${hist.map(r => `<tr class="${r.acima_vi ? 'txt-acima' : ''}"><td>${esc(r.campanha)}</td><td>${textoValor(r)}</td></tr>`).join('')}</table>` : ''}
        <div class="pop-mini">${p.cota_topo != null ? `Cota topo ${fmt(p.cota_topo, 3)} m · ` : ''}${esc(p.situacao)}${p.coordFonte ? ` · coordenada: ${esc(p.coordFonte)}` : ''}</div>
        ${p.virtual ? '<div class="pop-mini">Poço vindo da ficha do Perfil (ainda sem resultados de laboratório).</div>' : ''}
        ${fichas.length ? `<div class="pop-sond">Ficha no Perfil: ${fichas.map(f => `<b>${esc(f.sondagem_no || 'sondagem')}</b>${f.obra ? ' · ' + esc(f.obra) : ''}`).join('<br>')}</div>` : ''}
        ${perfilSVG(p, med)}
      </div>`;
  }

  // ------------------------------------------------------------- pluma (IDW)
  function limiarAtual(param) {
    if (!param) return null;
    return st.pluma.limiar ?? (param.valor_orientador != null ? +param.valor_orientador : null);
  }

  /**
   * Interpola o parâmetro selecionado em TODAS as campanhas, na mesma grade,
   * para dar para comparar área e centróide entre elas. Respeita o filtro de rede.
   * A interpolação é em escala log, então todo poço precisa de um valor > 0:
   *   - quantificado: o próprio valor
   *   - abaixo do LQ: metade do LQ (ou o piso, se o laudo não trouxe o LQ)
   * piso = 1/10 do valor orientador (sem VI: 1/10 do limiar ou do menor valor medido).
   */
  function calcularPlumas(param, limiar) {
    if (!param) return null;
    const visivel = p => !st.redesOcultas.has(p.rede || '-');
    const pocosCoord = new Map(st.pocos.filter(p => p.latitude !== null && visivel(p)).map(p => [p.id, p]));
    const res = st.resultados.filter(r => r.parametro_id === param.id && pocosCoord.has(r.poco_id));
    if (!res.length) return null;
    const ids = [...new Set(res.map(r => r.poco_id))];
    const g = WGPluma.grade(ids.map(id => ({ lat: +pocosCoord.get(id).latitude, lon: +pocosCoord.get(id).longitude })), st.pluma.cel);
    const quantTodos = res.filter(r => !r.menor_que_lq && +r.valor > 0).map(r => +r.valor);
    const base = param.valor_orientador != null ? +param.valor_orientador : (limiar ?? (quantTodos.length ? Math.min(...quantTodos) : 1));
    const piso = base / 10;
    const valorInterp = r => r.menor_que_lq ? (r.lq != null && +r.lq > 0 ? +r.lq / 2 : piso) : (+r.valor > 0 ? +r.valor : piso);
    let maxGlobal = 0;
    const porCamp = st.campanhas.map(c => {
      const rc = res.filter(r => r.campanha_id === c.id);
      const quant = rc.filter(r => !r.menor_que_lq).map(r => +r.valor);
      const max = quant.length ? Math.max(...quant) : null;
      if (max) maxGlobal = Math.max(maxGlobal, max);
      const item = { campanha: c, n: rc.length, max, nAcima: rc.filter(r => r.acima_vi).length, v: null, area: null, centroide: null };
      const p = pocosCoord;
      // poços no mesmo ponto (até 5 m: PM-26, PMN-26A, PMN-26B...) contam como um, com o maior valor
      const pontos = WGPluma.agrupar(rc.map(r => ({ lat: +p.get(r.poco_id).latitude, lon: +p.get(r.poco_id).longitude, v: valorInterp(r) })), 5);
      item.pontos = pontos.length;
      if (pontos.length >= 3) {
        item.v = WGPluma.idw(g, pontos, st.pluma.p, piso);
        if (limiar != null) Object.assign(item, WGPluma.areaECentroide(g, item.v, limiar));
      }
      return item;
    });
    // a pluma é pintada só onde o valor interpolado passa do limiar (sem limiar: 3 ordens abaixo do máximo)
    let vmax = Math.pow(10, Math.max(1, Math.ceil(Math.log10(maxGlobal || 10))));
    const vmin = limiar != null ? limiar : vmax / 1e3;
    if (vmax <= vmin * 1.5) vmax = vmin * 10;
    return { g, limiar, porCamp, vmax, vmin };
  }

  function plumaDaCampanha(id) { return st.calc?.porCamp.find(x => x.campanha.id === id) || null; }
  function primeiraComArea() { return st.calc?.porCamp.find(x => x.v) || null; }

  function renderPluma(param) {
    if (!camadaPluma) return;
    camadaPluma.clearLayers();
    const calc = st.calc;
    const sel = plumaDaCampanha(st.campanhaId);
    const mostrar = st.pluma.ativa && calc && sel && sel.v;
    $('#legenda-pluma').hidden = !mostrar;
    if (!mostrar) return;
    const { g, vmin, vmax } = calc;

    // 1) imagem da pluma: cada célula vira k x k pixels, interpolando (em log) entre as células
    //    vizinhas, para a borda sair lisa e coincidir com a linha do limiar
    const k = g.nx * g.ny * 16 <= 1500000 ? 4 : (g.nx * g.ny * 4 <= 1500000 ? 2 : 1);
    const W = g.nx * k, H = g.ny * k;
    const lg = new Float64Array(g.nx * g.ny);
    for (let q = 0; q < lg.length; q++) lg[q] = Math.log10(Math.max(1e-12, sel.v[q]));
    const lmin = Math.log10(vmin);
    const cv = document.createElement('canvas');
    cv.width = W; cv.height = H;
    const ctx = cv.getContext('2d');
    const img = ctx.createImageData(W, H);
    for (let py = 0; py < H; py++) {
      const gy = Math.min(g.ny - 1, Math.max(0, (py + .5) / k - .5)), j0 = Math.min(g.ny - 2, Math.floor(gy)), fy = g.ny > 1 ? gy - j0 : 0;
      for (let px = 0; px < W; px++) {
        const gx = Math.min(g.nx - 1, Math.max(0, (px + .5) / k - .5)), i0 = Math.min(g.nx - 2, Math.floor(gx)), fx = g.nx > 1 ? gx - i0 : 0;
        const a = lg[Math.max(0, j0) * g.nx + Math.max(0, i0)], b = lg[Math.max(0, j0) * g.nx + Math.min(g.nx - 1, i0 + 1)];
        const c = lg[Math.min(g.ny - 1, j0 + 1) * g.nx + Math.max(0, i0)], d = lg[Math.min(g.ny - 1, j0 + 1) * g.nx + Math.min(g.nx - 1, i0 + 1)];
        const l = (a * (1 - fx) + b * fx) * (1 - fy) + (c * (1 - fx) + d * fx) * fy;
        if (!(l >= lmin)) continue; // transparente
        const o = ((H - 1 - py) * W + px) * 4; // linha 0 da imagem = norte
        const [r, gg, bb] = WGPluma.cor(Math.pow(10, l), vmin, vmax);
        img.data[o] = r; img.data[o + 1] = gg; img.data[o + 2] = bb; img.data[o + 3] = 190;
      }
    }
    ctx.putImageData(img, 0, 0);
    const sw = g.proj.paraLatLon(g.xmin, g.ymin), ne = g.proj.paraLatLon(g.xmin + g.nx * g.cel, g.ymin + g.ny * g.cel);
    L.imageOverlay(cv.toDataURL(), [[sw.lat, sw.lon], [ne.lat, ne.lon]], { pane: 'pluma', interactive: false, className: 'img-pluma' }).addTo(camadaPluma);

    // 2) linha do limiar
    if (calc.limiar != null) {
      const segs = WGPluma.isolinha(g, sel.v, calc.limiar);
      if (segs.length) {
        L.polyline(segs, { pane: 'plumaLinha', color: '#172621', weight: 4, opacity: .55, interactive: false }).addTo(camadaPluma);
        L.polyline(segs, { pane: 'plumaLinha', color: '#ffffff', weight: 2, dashArray: '6 5', interactive: false }).addTo(camadaPluma);
      }
    }

    // 3) centróides: o da campanha e o da 1ª campanha, ligados por uma linha
    const ini = primeiraComArea();
    if (ini && ini !== sel && ini.centroide && sel.centroide) {
      L.polyline([[ini.centroide.lat, ini.centroide.lon], [sel.centroide.lat, sel.centroide.lon]],
        { pane: 'plumaLinha', color: '#ffffff', weight: 2, dashArray: '2 4' }).addTo(camadaPluma);
      L.marker([ini.centroide.lat, ini.centroide.lon], { icon: L.divIcon({ className: '', html: '<div class="centroide-mapa ini"></div>', iconSize: [14, 14] }), keyboard: false, interactive: false })
        .addTo(camadaPluma);
    }
    if (sel.centroide) {
      L.marker([sel.centroide.lat, sel.centroide.lon], { icon: L.divIcon({ className: '', html: '<div class="centroide-mapa"></div>', iconSize: [14, 14] }), keyboard: false, interactive: false })
        .addTo(camadaPluma);
    }

    // legenda
    const un = param?.unidade || 'µg/L';
    $('#leg-tit').textContent = `Pluma (${un})`;
    $('#leg-gradiente').style.background = `linear-gradient(to right, ${WGPluma.RAMPA.slice(2).join(', ')})`;
    const pos = v => 100 * Math.log10(v / vmin) / Math.log10(vmax / vmin);
    const ticks = [vmin];
    for (let e = Math.ceil(Math.log10(vmin * 2)); Math.pow(10, e) <= vmax * 1.0001; e++) ticks.push(Math.pow(10, e));
    $('#leg-ticks').innerHTML = ticks.map(v => `<span style="left:${pos(v)}%">${fmt(v, 3)}</span>`).join('');
    $('#leg-limiar').textContent = calc.limiar != null ? `Limiar (${fmt(calc.limiar, 4)} ${un})` : 'Limiar: sem VI, defina no menu';
  }

  // ------------------------------------------------------------- mapa potenciométrico
  /** Carga hidráulica (m) de cada poço na campanha: a gravada, ou cota do topo - N.A. */
  function cargasDaCampanha(campId) {
    const pocos = new Map(st.pocos.map(p => [p.id, p]));
    return st.medicoes.filter(m => m.campanha_id === campId).map(m => {
      const p = pocos.get(m.poco_id);
      if (!p || p.latitude === null) return null;
      const h = m.carga_hidraulica != null ? +m.carga_hidraulica
        : (p.cota_topo != null && m.nivel_agua != null ? +p.cota_topo - +m.nivel_agua : null);
      return h === null || !Number.isFinite(h) ? null : { p, h };
    }).filter(Boolean);
  }

  /** Lista de redes com carga medida (em qualquer campanha). O padrão é a rede rasa "PM". */
  function montarFiltroPot() {
    const pocos = new Map(st.pocos.map(p => [p.id, p]));
    const cont = {};
    st.medicoes.forEach(m => { const p = pocos.get(m.poco_id); if (p && m.nivel_agua != null) cont[p.rede || '-'] = (cont[p.rede || '-'] || 0) + 1; });
    const redes = Object.keys(cont).sort();
    $('#f-pot-rede').innerHTML = redes.map(r => `<option value="${esc(r)}">${esc(r)}</option>`).join('') + '<option value="*">Todas</option>';
    if (!st.pot.redeEscolhida || (st.pot.rede !== '*' && !redes.includes(st.pot.rede))) {
      st.pot.rede = redes.includes('PM') ? 'PM' : (redes.sort((a, b) => cont[b] - cont[a])[0] || '*');
    }
    $('#f-pot-rede').value = st.pot.rede;
  }

  function calcularPot() {
    if (!st.pot.ativa || !st.campanhaId) return null;
    const usados = cargasDaCampanha(st.campanhaId).filter(c => st.pot.rede === '*' || (c.p.rede || '-') === st.pot.rede);
    const semCota = st.medicoes.filter(m => m.campanha_id === st.campanhaId && m.nivel_agua != null).length - cargasDaCampanha(st.campanhaId).length;
    const r = { n: usados.length, semCota, porPoco: new Map(usados.map(c => [c.p.id, c.h])), sup: null };
    r.sup = WGPot.superficie(usados.map(c => ({ lat: +c.p.latitude, lon: +c.p.longitude, h: c.h })), st.pot.suav);
    if (!r.sup) return r;
    r.passo = st.pot.intervalo || WGPot.intervaloAuto(r.sup.hmin, r.sup.hmax);
    if ((r.sup.hmax - r.sup.hmin) / r.passo > 60) r.passo = WGPot.intervaloAuto(r.sup.hmin, r.sup.hmax); // intervalo pequeno demais
    r.niveis = WGPot.niveis(r.sup.hmin, r.sup.hmax, r.passo);
    r.fluxo = WGPot.fluxo(r.sup);
    r.residuo = Math.max(...r.sup.modelo.residuos.map(Math.abs));
    return r;
  }

  function renderPot() {
    if (!camadaPot) return;
    camadaPot.clearLayers();
    $('#pot-painel').hidden = !st.pot.ativa;
    const c = st.potCalc, ok = !!(c && c.sup);
    $('#legenda-pot').hidden = !ok;
    if (!st.pot.ativa) return;
    if (!ok) {
      $('#pot-resumo').textContent = !st.campanhaId ? 'Sem campanha.'
        : `Só ${c ? c.n : 0} poço(s) com N.A. e cota do topo nesta campanha e rede. São precisos pelo menos 3.`;
      return;
    }
    // curvas equipotenciais (uma faixa branca por baixo, para destacar do satélite)
    const linhas = [], rotulos = [];
    c.niveis.forEach(nv => {
      const cs = WGPot.curva(c.sup, nv);
      if (!cs.length) return;
      linhas.push(...cs);
      const maior = cs.reduce((a, b) => b.length > a.length ? b : a);
      rotulos.push({ ll: maior[Math.floor(maior.length / 2)], nv });
    });
    L.polyline(linhas, { pane: 'pot', color: '#ffffff', weight: 4, opacity: .85, interactive: false }).addTo(camadaPot);
    L.polyline(linhas, { pane: 'pot', color: '#1565A8', weight: 2, interactive: false }).addTo(camadaPot);
    const casas = c.passo < 0.1 ? 2 : (c.passo < 1 ? (Number.isInteger(c.passo * 10) ? 1 : 2) : 1);
    const txt = v => Number(v).toLocaleString('pt-BR', { minimumFractionDigits: casas, maximumFractionDigits: casas });
    rotulos.forEach(r => L.marker(r.ll, { icon: L.divIcon({ className: 'rotulo-pot', html: `<span>${txt(r.nv)}</span>`, iconSize: [0, 0] }), interactive: false, keyboard: false }).addTo(camadaPot));
    // setas de fluxo
    if (st.pot.setas && c.fluxo.setas.length) {
      const tr = c.fluxo.setas.flatMap(s => [[s.de, s.ate], [s.asa1, s.ate, s.asa2]]);
      L.polyline(tr, { pane: 'pot', color: '#ffffff', weight: 5, opacity: .9, interactive: false, lineCap: 'round' }).addTo(camadaPot);
      L.polyline(tr, { pane: 'pot', color: '#0B3558', weight: 2.5, interactive: false, lineCap: 'round' }).addTo(camadaPot);
    }
    $('#leg-pot-setas').hidden = !st.pot.setas;
    // resumo
    const f = c.fluxo, partes = [
      `${c.n} poço(s)`,
      `carga de ${fmt(c.sup.hmin, 2)} a ${fmt(c.sup.hmax, 2)} m`,
      `curvas a cada ${fmt(c.passo, 2)} m`
    ];
    if (f.gradMedio != null) {
      partes.push(`gradiente médio ${fmt(f.gradMedio, 3)} m/m`);
      partes.push(f.constancia >= .5 ? `fluxo para ${f.rumo} (${Math.round(f.azimute)}°)` : `fluxo sem sentido único (médio para ${f.rumo})`);
    }
    if (st.pot.suav > 0) partes.push(`suavizada: até ${fmt(c.residuo, 2)} m do medido`);
    if (c.semCota > 0) partes.push(`${c.semCota} poço(s) com N.A. ficaram de fora por falta de cota ou coordenada`);
    $('#pot-resumo').textContent = partes.join(' · ');
    $('#leg-pot-resumo').textContent = partes.slice(0, 5).join(' · ');
  }

  // ------------------------------------------------------------- seção geológica (corte A–A')
  const chaveSecao = () => 'webgeo.secao.' + (st.projetoId || '');
  const NOME_FIM_SECAO = "A'";

  /** A linha da seção fica guardada neste navegador, por projeto. */
  function carregarSecao() {
    st.secao.tracando = null; st.secao.linha = null;
    try {
      const v = JSON.parse(localStorage.getItem(chaveSecao()) || 'null');
      if (Array.isArray(v) && v.length >= 2 && v.every(p => Array.isArray(p) && Number.isFinite(p[0]) && Number.isFinite(p[1]))) st.secao.linha = v;
    } catch (e) { /* sem armazenamento: começa sem seção */ }
    $('#secao-barra').hidden = true;
    if (mapa) { mapa.getContainer().classList.remove('mapa-ajustando'); mapa.doubleClickZoom.enable(); }
  }
  function guardarSecao() {
    try {
      if (st.secao.linha) localStorage.setItem(chaveSecao(), JSON.stringify(st.secao.linha)); else localStorage.removeItem(chaveSecao());
    } catch (e) { /* sem armazenamento */ }
  }

  function iniciarSecao() {
    if (!mapa || !st.projetoId) { toast('Escolha um projeto primeiro.'); return; }
    if (st.ajuste) cancelarAjuste(true);
    trocarView('visao');
    st.secao.tracando = { pontos: [] };
    mapa.closePopup();
    mapa.doubleClickZoom.disable();
    mapa.getContainer().classList.add('mapa-ajustando');
    $('#secao-barra').hidden = false;
    textoSecao(); renderSecaoMapa();
    $('#secao-barra').scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }

  function textoSecao() {
    const t = st.secao.tracando; if (!t) return;
    const n = t.pontos.length;
    $('#secao-passo').textContent = n === 0 ? 'Clique no mapa onde a seção começa (ponto A)'
      : n === 1 ? `Agora clique onde ela termina (ponto ${NOME_FIM_SECAO})` : `${n} pontos marcados. Clique em "Concluir" ou marque mais um ponto`;
    $('#secao-dica').textContent = n === 0 ? 'O clique gruda no poço mais próximo. Passe a linha pelos poços que devem aparecer no corte.'
      : 'Pode marcar pontos intermediários para a linha fazer curva. Dois cliques rápidos também concluem.';
    $('#secao-concluir').disabled = n < 2;
    $('#secao-desfazer').disabled = n === 0;
  }

  function cliqueSecao(ev) {
    const t = st.secao.tracando; if (!t) return;
    let ll = [ev.latlng.lat, ev.latlng.lng];
    const c = mapa.latLngToContainerPoint(ev.latlng);
    let melhor = 15;                                    // gruda no poço mais próximo, até 15 px
    st.pontos.forEach(p => { const d = mapa.latLngToContainerPoint(p).distanceTo(c); if (d <= melhor) { melhor = d; ll = [p[0], p[1]]; } });
    const u = t.pontos[t.pontos.length - 1];
    if (u && mapa.latLngToContainerPoint(u).distanceTo(mapa.latLngToContainerPoint(ll)) < 5) return;   // 2º clique de um clique duplo
    t.pontos.push(ll);
    textoSecao(); renderSecaoMapa();
  }

  function concluirSecao() {
    const t = st.secao.tracando; if (!t || t.pontos.length < 2) return;
    st.secao.linha = t.pontos;
    fecharTracadoSecao();
    guardarSecao();
    render();
    $('#sec-secao').scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }
  function fecharTracadoSecao() {
    st.secao.tracando = null;
    $('#secao-barra').hidden = true;
    mapa.getContainer().classList.remove('mapa-ajustando');
    setTimeout(() => mapa.doubleClickZoom.enable(), 300);   // depois do clique duplo que concluiu
  }
  function cancelarSecao() { if (!st.secao.tracando) return; fecharTracadoSecao(); renderSecaoMapa(); }
  function desfazerSecao() { const t = st.secao.tracando; if (!t) return; t.pontos.pop(); textoSecao(); renderSecaoMapa(); }
  function apagarSecao() { st.secao.linha = null; guardarSecao(); render(); toast('Seção apagada.'); }

  /** Linha da seção no mapa (a definitiva ou a que está sendo traçada), com A e A' nas pontas. */
  function renderSecaoMapa() {
    if (!camadaSecao) return;
    camadaSecao.clearLayers();
    const t = st.secao.tracando, pts = t ? t.pontos : st.secao.linha;
    if (!pts || !pts.length) return;
    if (pts.length >= 2) {
      L.polyline(pts, { pane: 'secao', color: '#172621', weight: 6, opacity: .6, interactive: false }).addTo(camadaSecao);
      L.polyline(pts, { pane: 'secao', color: '#FFD60A', weight: 2.5, dashArray: t ? '6 6' : null, interactive: false }).addTo(camadaSecao);
    }
    const pino = (ll, txt) => L.marker(ll, { icon: L.divIcon({ className: '', html: `<div class="pino-secao">${txt}</div>`, iconSize: [22, 22] }), interactive: false, keyboard: false, zIndexOffset: 500 }).addTo(camadaSecao);
    pino(pts[0], 'A');
    if (pts.length >= 2) pino(pts[pts.length - 1], NOME_FIM_SECAO);
    pts.slice(1, -1).forEach(ll => L.circleMarker(ll, { pane: 'secao', radius: 4, color: '#172621', weight: 1.5, fillColor: '#FFD60A', fillOpacity: 1, interactive: false }).addTo(camadaSecao));
    // poços que entram no corte: anel amarelo
    if (!t && st.secaoCalc) st.secaoCalc.cols.forEach(c =>
      L.circleMarker(c.ll, { pane: 'secao', radius: 11, color: '#FFD60A', weight: 2, fill: false, interactive: false }).addTo(camadaSecao));
  }

  /** Dados de uma sondagem para o corte: camadas, tubo, N.A., cota e profundidade. p = poço; ficha = sondagem sem poço. */
  function colunaSecao(p, ficha) {
    const num = WGImport.numero;
    const fichas = ficha ? [ficha] : (st.sondagens.get(codigo(p.codigo)) || []);
    const f = fichas.length ? fichas[fichas.length - 1] : null, dados = f?.data || {};
    const wg = p ? (st.perfis.get(p.id) || []) : [];
    let camadas = (dados.litologia || [])
      .map(r => ({ de: num(r.inicio), ate: num(r.termino), lit: String(r.litologia || '').trim(), cor: String(r.cor || '').trim(), hex: corLitologia(r.cor, r.litologia) }))
      .filter(r => r.de != null && r.ate != null && r.ate > r.de);
    let fonte = camadas.length ? `ficha ${f.sondagem_no || ''} do Perfil`.replace('  ', ' ') : '';
    if (!camadas.length) {
      camadas = wg.filter(r => r.estrutura === 'SOLO').map(r => ({ de: +r.de_m, ate: +r.ate_m, lit: String(r.descricao || '').trim(), cor: '', hex: r.cor_hex || '#9C8F78' }));
      if (camadas.length) fonte = 'aba Litologia da planilha';
    }
    camadas.sort((a, b) => a.de - b.de);
    let tubo = [];
    const liso = num(dados.perfil?.tuboLiso), filtro = num(dados.perfil?.tuboFiltro);
    if (liso != null && filtro != null && liso + filtro > 0) tubo = [{ tipo: 'CEGO', de: 0, ate: liso }, { tipo: 'FILTRO', de: liso, ate: liso + filtro }];
    else tubo = wg.filter(r => r.estrutura === 'CEGO' || r.estrutura === 'FILTRO').map(r => ({ tipo: r.estrutura, de: +r.de_m, ate: +r.ate_m }));
    const med = p ? st.medicoes.find(m => m.poco_id === p.id && m.campanha_id === st.campanhaId) : null;
    let na = null, naFonte = '';
    if (med?.nivel_agua != null) { na = +med.nivel_agua; naFonte = 'campanha'; }
    else if (num(dados.perfil?.naEstabilizado) != null) { na = num(dados.perfil.naEstabilizado); naFonte = 'ficha'; }
    const prof = Math.max(0, ...camadas.map(r => r.ate), ...tubo.map(r => r.ate), num(p?.profundidade) || 0,
      num(dados.perfil?.profTotalSondagem) || 0, num(dados.perfil?.profTotalPoco) || 0);
    return {
      nome: p ? p.codigo : (ficha.sondagem_no || 'Sondagem'), rede: p ? (p.rede || '-') : 'Sondagem', pocoId: p ? p.id : null,
      cota: p && p.cota_topo != null ? +p.cota_topo : null, prof, camadas, tubo, na, naFonte, fonte
    };
  }

  /** Quem entra no corte: poços e sondagens a até "faixa" metros da linha, na ordem em que aparecem de A para A'. */
  function montarSecao() {
    const ll = st.secao.linha; if (!ll) return null;
    const proj = WGPluma.projecao(ll[0][0], ll[0][1]);
    const linha = ll.map(p => proj.paraXY(p[0], p[1]));
    const comp = WGSecao.comprimento(linha);
    if (!(comp > 0.5)) return null;
    const { pocos, porPoco } = selecao();
    const cand = pocos.filter(p => p.latitude !== null).map(p => ({ p, lat: +p.latitude, lon: +p.longitude }));
    if (!st.redesOcultas.has('Sondagens')) st.soltas.forEach(f => cand.push({ f, lat: f.latitude, lon: f.longitude }));
    const comCota = st.pocos.filter(p => p.latitude !== null && p.cota_topo != null)
      .map(p => ({ ...proj.paraXY(+p.latitude, +p.longitude), z: +p.cota_topo }));
    const cols = [];
    cand.forEach(c => {
      const xy = proj.paraXY(c.lat, c.lon), pr = WGSecao.projetar(linha, xy);
      if (pr.afast > st.secao.faixa) return;
      const col = colunaSecao(c.p || null, c.f || null);
      Object.assign(col, { dist: pr.dist, afast: pr.afast, ll: [c.lat, c.lon], res: c.p ? (porPoco.get(c.p.id) || null) : null });
      if (col.cota == null) {                            // sem cota: usa a do poço mais próximo (marcado com *)
        col.cotaEstimada = true;
        if (comCota.length) col.cota = comCota.reduce((a, b) => Math.hypot(b.x - xy.x, b.y - xy.y) < Math.hypot(a.x - xy.x, a.y - xy.y) ? b : a).z;
        else col.cota = 0;
      }
      cols.push(col);
    });
    cols.sort((a, b) => (a.dist - b.dist) || a.nome.localeCompare(b.nome, 'pt-BR', { numeric: true }));
    return { comp, cols, vertices: WGSecao.vertices(linha), semDatum: !comCota.length };
  }

  function renderSecao() {
    const tem = !!st.secao.linha;
    $('#sec-secao').hidden = !tem;
    $('#secao-painel').hidden = !tem;
    $('#btn-secao-tracar').textContent = tem ? 'Traçar outra seção' : `Traçar seção A–${NOME_FIM_SECAO}`;
    st.secaoCalc = tem ? montarSecao() : null;
    renderSecaoMapa();
    if (!tem) return;
    const d = st.secaoCalc, host = $('#secao-desenho');
    $('#btn-secao-baixar').disabled = !(d && d.cols.length);
    if (!d || !d.cols.length) {
      host.innerHTML = `<p class="secao-vazio">Nenhum poço ou sondagem a até ${fmt(st.secao.faixa, 0)} m da linha. Aumente a faixa no menu ou trace a linha passando pelos poços.</p>`;
      $('#secao-legenda').innerHTML = ''; $('#secao-nota').textContent = ''; $('#secao-hint').textContent = '';
      return;
    }
    const cols = d.cols, COL = 18;
    const W = Math.max(720, (host.clientWidth || 1020) - 40), mL = 56, mR = 62, mT = 70, mB = 46, pw = W - mL - mR;
    const zTopo = Math.max(...cols.map(c => c.cota));
    const zBase = Math.min(...cols.map(c => c.cota - Math.max(c.prof, c.na || 0, 1)));
    const folga = Math.max(.3, (zTopo - zBase) * .06), zmax = zTopo + folga, zmin = zBase - folga;
    const pxH = pw / d.comp;                                         // pixels por metro na horizontal
    let ev = st.secao.exag || WGSecao.exageroAuto(pw, d.comp, 330, zmax - zmin), limitado = false;
    let ph = (zmax - zmin) * pxH * ev;
    if (ph > 900) { ph = 900; limitado = true; } else if (ph < 120) { ph = 120; limitado = true; }
    if (limitado) ev = ph / ((zmax - zmin) * pxH);
    const H = mT + ph + mB, pxV = ph / (zmax - zmin);
    const X = s => mL + s * pxH, Y = z => mT + (zmax - z) * pxV;
    const xr = cols.map(c => X(c.dist));
    const xd = WGSecao.espalhar(xr, COL + 8, mL + COL / 2, W - mR - COL / 2);  // poços colados ficam lado a lado
    const n1 = v => (+v).toFixed(1);
    let g = '';

    // eixos e grade
    const pz = WGSecao.passoBonito(zmax - zmin, Math.max(3, Math.round(ph / 55)));
    for (let z = Math.ceil(zmin / pz) * pz; z <= zmax + 1e-9; z += pz) {
      g += `<line x1="${mL}" x2="${W - mR}" y1="${n1(Y(z))}" y2="${n1(Y(z))}" stroke="currentColor" stroke-opacity=".13"/>`
        + `<text x="${mL - 7}" y="${n1(Y(z) + 3.5)}" font-size="10.5" text-anchor="end" fill="currentColor" fill-opacity=".8">${fmt(d.semDatum ? Math.abs(z) : z, 2)}</text>`;
    }
    const pd = WGSecao.passoBonito(d.comp, Math.max(4, Math.round(pw / 110)));
    for (let s = 0; s <= d.comp + 1e-9; s += pd) {
      g += `<line x1="${n1(X(s))}" x2="${n1(X(s))}" y1="${mT + ph}" y2="${mT + ph + 5}" stroke="currentColor" stroke-opacity=".6"/>`
        + `<text x="${n1(X(s))}" y="${mT + ph + 17}" font-size="10.5" text-anchor="middle" fill="currentColor" fill-opacity=".8">${fmt(s, 1)}</text>`;
    }
    g += `<rect x="${mL}" y="${mT}" width="${pw}" height="${n1(ph)}" fill="none" stroke="currentColor" stroke-opacity=".45"/>`
      + `<text x="${mL + pw / 2}" y="${n1(H - 8)}" font-size="11" text-anchor="middle" fill="currentColor" fill-opacity=".8">Distância ao longo da seção (m)</text>`
      + `<text transform="rotate(-90 14 ${n1(mT + ph / 2)})" x="14" y="${n1(mT + ph / 2)}" font-size="11" text-anchor="middle" fill="currentColor" fill-opacity=".8">${d.semDatum ? 'Profundidade (m)' : 'Cota (m)'}</text>`
      + `<text x="${mL}" y="20" font-size="17" font-weight="700" fill="currentColor">A</text>`
      + `<text x="${W - mR}" y="20" font-size="17" font-weight="700" text-anchor="end" fill="currentColor">${NOME_FIM_SECAO}</text>`;
    d.vertices.slice(1, -1).forEach(s => {
      g += `<line x1="${n1(X(s))}" x2="${n1(X(s))}" y1="${mT}" y2="${mT + ph}" stroke="currentColor" stroke-opacity=".4" stroke-dasharray="2 4"/>`
        + `<text x="${n1(X(s))}" y="${mT + ph - 5}" font-size="9.5" text-anchor="middle" fill="currentColor" fill-opacity=".6">muda de direção</text>`;
    });

    // preenchimento entre sondagens vizinhas: liga as camadas parecidas e afina as que não têm par (interpretação automática)
    const comLito = cols.map((c, i) => i).filter(i => cols[i].camadas.length);   // pula as sondagens sem litologia
    let defs = '', temVazio = false;
    if (st.secao.ligar) for (let q = 0; q + 1 < comLito.length; q++) {
      const i = comLito[q], j = comLito[q + 1], a = cols[i], b = cols[j], x1 = xd[i], x2 = xd[j];   // até o eixo: a coluna é desenhada por cima
      const P = WGSecao.preencher(a.camadas, b.camadas, a.cota, b.cota);
      const ponto = (k, prof) => `${n1(x1 + (x2 - x1) * P.s[k])},${n1(Y(a.cota + (b.cota - a.cota) * P.s[k] - prof))}`;
      const nome = c => c.vazio ? 'trecho sem descrição' : [c.lit, c.cor].filter(Boolean).join(', ');
      P.faixas.forEach((f, k) => {
        const ca = f.ia != null ? P.a[f.ia] : null, cb = f.ib != null ? P.b[f.ib] : null;
        if (!f.base.some((z, w) => z - f.topo[w] > 1e-4)) return;
        if ((ca || cb).vazio) temVazio = true;
        let cor = esc((ca || cb).hex);
        if (ca && cb && ca.hex !== cb.hex) {                      // mesma camada com cores diferentes: passa de uma para a outra
          const id = `sg${i}_${k}`;
          defs += `<linearGradient id="${id}" gradientUnits="userSpaceOnUse" x1="${n1(x1)}" x2="${n1(x2)}" y1="0" y2="0"><stop offset="0" stop-color="${esc(ca.hex)}"/><stop offset="1" stop-color="${esc(cb.hex)}"/></linearGradient>`;
          cor = `url(#${id})`;
        }
        const d = 'M' + P.s.map((_, w) => ponto(w, f.topo[w])).join(' L') + ' L' + P.s.map((_, w) => ponto(P.s.length - 1 - w, f.base[P.s.length - 1 - w])).join(' L') + ' Z';
        const dica = ca && cb ? `${nome(ca)} (${a.nome})  ↔  ${nome(cb)} (${b.nome})`
          : f.abaixo ? `${nome(ca || cb)} (${(ca ? a : b).nome}): abaixo do fim de ${(ca ? b : a).nome}, que é mais rasa`
          : ca ? `${nome(ca)} (${a.nome}): não aparece em ${b.nome}, acaba no caminho` : `${nome(cb)} (${b.nome}): não aparece em ${a.nome}, acaba no caminho`;
        g += `<path class="secao-faixa" d="${d}" fill="${cor}" fill-opacity=".8" stroke="#172621" stroke-opacity=".45" stroke-width=".7" stroke-linejoin="round"><title>${esc(dica)}\nInterpretação automática entre as sondagens</title></path>`;
      });
    }
    if (defs) g = `<defs>${defs}</defs>` + g;

    // terreno (cota do topo de cada ponto)
    const terreno = [[mL, Y(cols[0].cota)], ...cols.map((c, i) => [xd[i], Y(c.cota)]), [W - mR, Y(cols[cols.length - 1].cota)]];
    g += `<polyline points="${terreno.map(p => n1(p[0]) + ',' + n1(p[1])).join(' ')}" fill="none" stroke="#7A5A2E" stroke-width="2"/>`;

    // sondagens
    const litos = new Map();
    cols.forEach((c, i) => {
      const x = xd[i], x0 = x - COL / 2, yt = Y(c.cota);
      const dica = [c.nome + (c.rede && c.rede !== '-' ? ` (${c.rede})` : ''),
        d.semDatum ? '' : `Cota do topo: ${fmt(c.cota, 3)} m${c.cotaEstimada ? ' (estimada pelo poço mais próximo)' : ''}`,
        c.prof ? `Profundidade: ${fmt(c.prof)} m` : '', c.na != null ? `N.A.: ${fmt(c.na)} m (${c.naFonte === 'campanha' ? 'medido na campanha' : 'da ficha do Perfil'})` : '',
        c.res ? `Resultado: ${textoValor(c.res)} (${ROTULO_STATUS[status(c.res)]})` : '', c.fonte ? `Litologia: ${c.fonte}` : 'Sem litologia cadastrada',
        `A ${fmt(c.dist, 1)} m do ponto A · ${fmt(c.afast, 1)} m fora da linha`].filter(Boolean).join('\n');
      g += `<g class="secao-coluna"><title>${esc(dica)}</title>`;
      g += `<path d="M${n1(xr[i] - 3)} ${mT + ph} h6 l-3 -5 z" fill="currentColor" fill-opacity=".55"/>`;   // posição real na linha
      if (c.camadas.length) {
        c.camadas.forEach(r => {
          g += `<rect x="${n1(x0)}" y="${n1(Y(c.cota - r.de))}" width="${COL}" height="${n1(Math.max(.6, (r.ate - r.de) * pxV))}" fill="${esc(r.hex)}" stroke="#ffffff" stroke-width=".6"/>`;
          const k = WGSecao.chaveLit(r.lit) + '|' + r.hex;
          if (!litos.has(k)) litos.set(k, { hex: r.hex, texto: [r.lit, r.cor].filter(Boolean).join(' · ') || 'sem descrição' });
        });
        g += `<rect x="${n1(x0)}" y="${n1(yt)}" width="${COL}" height="${n1(Math.max(1, c.prof * pxV))}" fill="none" stroke="currentColor" stroke-opacity=".75"/>`;
      } else if (c.prof > 0) {
        g += `<rect x="${n1(x0)}" y="${n1(yt)}" width="${COL}" height="${n1(c.prof * pxV)}" fill="currentColor" fill-opacity=".05" stroke="currentColor" stroke-opacity=".6" stroke-dasharray="3 3"/>`;
      } else {
        g += `<line x1="${n1(x)}" x2="${n1(x)}" y1="${n1(yt)}" y2="${n1(yt + 14)}" stroke="currentColor" stroke-opacity=".6" stroke-dasharray="3 3"/>`;
      }
      c.tubo.forEach(r => {
        const y0 = Y(c.cota - r.de), h = (r.ate - r.de) * pxV;
        g += r.tipo === 'FILTRO'
          ? `<rect x="${n1(x - 3.5)}" y="${n1(y0)}" width="7" height="${n1(h)}" fill="#DDEAF6" stroke="#2B7BBA"/>`
            + Array.from({ length: Math.floor(h / 4) }, (_, q) => `<line x1="${n1(x - 2.5)}" x2="${n1(x + 2.5)}" y1="${n1(y0 + 2 + q * 4)}" y2="${n1(y0 + 2 + q * 4)}" stroke="#2B7BBA" stroke-width=".8"/>`).join('')
          : `<rect x="${n1(x - 3.5)}" y="${n1(y0)}" width="7" height="${n1(h)}" fill="#F3F5F4" stroke="#56655D"/>`;
      });
      if (c.na != null) {
        const yn = Y(c.cota - c.na);
        g += `<path d="M${n1(x0 - 9)} ${n1(yn - 7)} h8 l-4 6.5 z" fill="#1565A8" stroke="#ffffff" stroke-width=".6"/>`
          + `<line x1="${n1(x0 - 2)}" x2="${n1(x0 + COL + 2)}" y1="${n1(yn)}" y2="${n1(yn)}" stroke="#1565A8" stroke-width="1.6"/>`;
      }
      if (c.res) g += `<circle cx="${n1(x)}" cy="${n1(yt - 7)}" r="4" fill="${COR_STATUS[status(c.res)]}" stroke="#ffffff" stroke-width="1"/>`;
      g += `<text transform="rotate(-55 ${n1(x)} ${n1(yt - 15)})" x="${n1(x)}" y="${n1(yt - 15)}" font-size="10.5" font-weight="600" fill="currentColor" font-family="IBM Plex Mono, monospace">${esc(c.nome)}${c.cotaEstimada && !d.semDatum ? '*' : ''}</text></g>`;
    });

    // nível d'água: uma linha por rede (PM e PMN medem níveis diferentes do aquífero)
    const TRACOS = ['7 3', '2 3', '9 3 2 3', '4 4'];
    const redes = [...new Set(cols.filter(c => c.na != null).map(c => c.rede))];
    redes.forEach((r, k) => {
      const pts = cols.map((c, i) => c.rede === r && c.na != null ? n1(xd[i]) + ',' + n1(Y(c.cota - c.na)) : null).filter(Boolean);
      if (pts.length >= 2) g += `<polyline points="${pts.join(' ')}" fill="none" stroke="#1565A8" stroke-width="1.5" stroke-dasharray="${TRACOS[k % TRACOS.length]}"/>`;
    });

    host.innerHTML = `<svg id="secao-svg" xmlns="http://www.w3.org/2000/svg" width="${W}" height="${n1(H)}" viewBox="0 0 ${W} ${n1(H)}" font-family="IBM Plex Sans, sans-serif" role="img" aria-label="Seção geológica A–${NOME_FIM_SECAO}">${g}</svg>`;

    // legenda e notas
    const item = (desenho, texto) => `<span>${desenho}${esc(texto)}</span>`;
    let leg = [...litos.values()].map(l => item(`<i class="leg-lito" style="background:${esc(l.hex)}"></i>`, l.texto)).join('');
    leg += item('<i class="leg-terreno"></i>', 'Terreno (cota do topo)');
    if (cols.some(c => c.tubo.length)) leg += item('<i class="leg-tubo"></i>', 'Tubo liso') + item('<i class="leg-tubo leg-filtro"></i>', 'Filtro');
    redes.forEach((r, k) => { leg += `<span><svg width="26" height="8"><line x1="0" x2="26" y1="4" y2="4" stroke="#1565A8" stroke-width="1.6" stroke-dasharray="${TRACOS[k % TRACOS.length]}"/></svg>${esc("Nível d'água" + (r !== '-' ? ' · ' + r : ''))}</span>`; });
    if (temVazio) leg += item('<i class="leg-lito" style="background:#E4E7E5"></i>', 'Trecho sem descrição');
    if (cols.some(c => !c.camadas.length)) leg += item('<i class="leg-semlito"></i>', 'Sem litologia cadastrada');
    $('#secao-legenda').innerHTML = leg;
    const camp = st.campanhas.find(c => c.id === st.campanhaId);
    const notas = [
      `Exagero vertical ${fmt(ev, 1)}x${limitado && st.secao.exag ? ' (ajustado para caber na tela)' : ''}`,
      `comprimento ${fmt(d.comp, 1)} m`, `faixa de ${fmt(st.secao.faixa, 0)} m para cada lado`, `${cols.length} ponto(s)`
    ];
    if (cols.some(c => c.naFonte === 'campanha') && camp) notas.push(`N.A. da campanha ${camp.codigo}`);
    if (d.semDatum) notas.push('nenhum poço tem cota: o desenho está em profundidade, com todos os topos no mesmo nível');
    else if (cols.some(c => c.cotaEstimada)) notas.push('* sem cota cadastrada: usada a cota do poço mais próximo');
    if (st.secao.ligar && comLito.length > 1) notas.push('o preenchimento entre as sondagens é uma interpretação automática (liga camadas parecidas e afina as que não têm par) e precisa de conferência');
    $('#secao-nota').textContent = notas.join(' · ') + '. Passe o mouse sobre uma sondagem para ver os detalhes.';
    $('#secao-hint').textContent = `${cols.length} ponto(s) · ${fmt(d.comp, 0)} m`;
  }

  /** Baixa o desenho da seção em SVG (abre no navegador, no Inkscape, no CorelDRAW e no AutoCAD mais novo). */
  function baixarSecao() {
    const svg = document.querySelector('#secao-svg'); if (!svg) return;
    const c = svg.cloneNode(true);
    c.setAttribute('style', 'color:#172621;background:#ffffff');
    const fundo = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
    fundo.setAttribute('width', '100%'); fundo.setAttribute('height', '100%'); fundo.setAttribute('fill', '#ffffff');
    c.insertBefore(fundo, c.firstChild);
    const proj = st.projetos.find(p => p.id === st.projetoId);
    const url = URL.createObjectURL(new Blob(['<?xml version="1.0" encoding="UTF-8"?>\n' + new XMLSerializer().serializeToString(c)], { type: 'image/svg+xml' }));
    const a = document.createElement('a');
    a.href = url; a.download = `secao_A-A_${(proj?.nome || 'projeto').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\-]+/g, '_')}.svg`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  function renderEvolucao(param) {
    const calc = st.calc, un = param?.unidade || 'µg/L';
    const sel = plumaDaCampanha(st.campanhaId), ini = primeiraComArea();
    const fa = a => a == null ? '—' : fmt(Math.round(a), 0);
    const semLimiar = !calc || calc.limiar == null;
    $('#pluma-hint').textContent = calc
      ? `IDW em escala log, p = ${st.pluma.p} · célula ${calc.g.cel} m · limiar ${semLimiar ? '—' : fmt(calc.limiar, 4) + ' ' + un} · poços a até 5 m contam como um ponto (maior valor)`
      : 'sem resultados deste parâmetro';

    $('#kpi-a1-lab').textContent = ini ? `Área na ${ini.campanha.codigo}` : 'Área na 1ª campanha';
    $('#kpi-a2-lab').textContent = sel ? `Área na ${sel.campanha.codigo}` : 'Área na campanha';
    $('#kpi-a1').textContent = fa(ini?.area);
    $('#kpi-a2').textContent = fa(sel?.area);
    const sub = semLimiar ? 'defina o limiar' : `acima de ${fmt(calc.limiar, 4)} ${un}`;
    $('#kpi-a1-sub').textContent = ini ? sub : 'precisa de 3+ pontos com resultado';
    $('#kpi-a2-sub').textContent = sel?.v ? sub : 'precisa de 3+ pontos com resultado';

    const comparar = ini && sel && ini !== sel && ini.area != null && sel.area != null;
    if (comparar && ini.area > 0) {
      const red = (1 - sel.area / ini.area) * 100;
      $('#kpi-red').textContent = (red >= 0 ? '' : '+') + fmt(Math.abs(red), 1) + '%';
      $('#kpi-red-sub').textContent = red >= 0 ? `redução entre ${ini.campanha.codigo} e ${sel.campanha.codigo}` : `AUMENTO entre ${ini.campanha.codigo} e ${sel.campanha.codigo}`;
    } else {
      $('#kpi-red').textContent = '—';
      $('#kpi-red-sub').textContent = comparar ? `sem área acima do limiar na ${ini.campanha.codigo}` : 'escolha outra campanha para comparar com a 1ª';
    }
    const d = comparar ? WGPluma.distancia(ini.centroide, sel.centroide) : null;
    $('#kpi-desl').textContent = d != null ? fmt(d, 1) : '—';
    $('#kpi-desl-un').textContent = d != null ? 'm' : '';
    $('#kpi-desl-sub').textContent = d != null ? `entre ${ini.campanha.codigo} e ${sel.campanha.codigo}` : (comparar ? 'sem pluma em uma das campanhas' : 'escolha outra campanha para comparar com a 1ª');

    // tabela por campanha (também é a versão em texto dos gráficos)
    $('#tabela-evolucao tbody').innerHTML = (calc?.porCamp || []).map(x => {
      const dd = ini && x !== ini ? WGPluma.distancia(ini.centroide, x.centroide) : null;
      return `<tr class="${x.campanha.id === st.campanhaId ? 'linha-sel' : ''}">
        <td class="cod">${esc(x.campanha.codigo)}</td><td>${fmtData(x.campanha.data_inicio)}</td>
        <td class="num">${x.n}</td><td class="num">${x.nAcima}</td>
        <td class="num">${x.max != null ? fmt(x.max, 4) : '< LQ'}</td>
        <td class="num">${x.v ? fa(x.area) : 'poucos poços'}</td>
        <td class="num">${dd != null ? fmt(dd, 1) + ' m' : (x === ini ? 'referência' : '—')}</td></tr>`;
    }).join('') || '<tr><td colspan="7" class="muted">Sem resultados deste parâmetro.</td></tr>';

    desenharGraficos(param);
  }

  // ------------------------------------------------------------- gráficos (Chart.js)
  function desenharGraficos(param) {
    if (!window.Chart) return;
    const cs = getComputedStyle(document.documentElement);
    const c = n => cs.getPropertyValue(n).trim();
    const cor = { linha: c('--accent'), texto: c('--ink-soft'), grade: c('--line'), fundo: c('--surface'), perigo: c('--danger') };
    const itens = st.calc?.porCamp || [];
    const labels = itens.map(x => x.campanha.codigo);
    const un = param?.unidade || 'µg/L';
    const vo = param?.valor_orientador != null ? +param.valor_orientador : null;

    // linha de referência (VI / limiar) desenhada à mão, sem virar uma "série"
    const linhaRef = (valor, rotulo) => ({
      id: 'linhaRef',
      afterDatasetsDraw(ch) {
        if (valor == null) return;
        const y = ch.scales.y.getPixelForValue(valor);
        if (!(y >= ch.chartArea.top && y <= ch.chartArea.bottom)) return;
        const k = ch.ctx; k.save();
        k.strokeStyle = cor.perigo; k.lineWidth = 1.5; k.setLineDash([5, 4]);
        k.beginPath(); k.moveTo(ch.chartArea.left, y); k.lineTo(ch.chartArea.right, y); k.stroke();
        k.setLineDash([]); k.fillStyle = cor.perigo; k.font = '600 11px "IBM Plex Sans", sans-serif';
        k.textAlign = 'right'; k.fillText(rotulo, ch.chartArea.right - 4, y - 5); k.restore();
      }
    });
    const base = (yExtra, fmtY) => ({
      responsive: true, maintainAspectRatio: false, animation: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ctx.parsed.y == null ? 'sem dado' : fmtY(ctx.parsed.y) } }
      },
      scales: {
        x: { grid: { display: false }, ticks: { color: cor.texto, font: { family: 'IBM Plex Mono', size: 11 } }, border: { color: cor.grade } },
        y: Object.assign({ grid: { color: cor.grade + '80' }, border: { display: false },
          ticks: { color: cor.texto, font: { family: 'IBM Plex Sans', size: 11 }, callback: v => fmtY(v) } }, yExtra)
      }
    });
    const serie = dados => ({
      data: dados, borderColor: cor.linha, backgroundColor: cor.linha, borderWidth: 2,
      pointRadius: 4, pointHoverRadius: 6, pointBorderColor: cor.fundo, pointBorderWidth: 2, spanGaps: true, tension: 0
    });

    Object.values(graficos).forEach(g => g.destroy());
    graficos = {};
    graficos.area = new Chart($('#g-area'), {
      type: 'line',
      data: { labels, datasets: [serie(itens.map(x => x.v && x.area != null ? Math.round(x.area) : null))] },
      options: base({ beginAtZero: true }, v => fmt(v, 0) + ' m²')
    });
    const maxs = itens.map(x => x.max && x.max > 0 ? x.max : null);
    const vals = [...maxs.filter(Boolean), ...(vo ? [vo] : [])];
    const ymax = Math.pow(10, Math.ceil(Math.log10(Math.max(...vals, 1) * 1.2)));
    const ymin = Math.pow(10, Math.floor(Math.log10(Math.min(...vals, ymax / 10) / 1.2)));
    const optConc = base({ type: 'logarithmic', min: ymin, max: ymax }, v => fmt(v, 3) + ' ' + un);
    optConc.scales.y.ticks.callback = v => Number.isInteger(Math.round(Math.log10(v) * 1e6) / 1e6) ? fmt(v, 3) + ' ' + un : '';
    graficos.conc = new Chart($('#g-conc'), {
      type: 'line',
      data: { labels, datasets: [serie(maxs)] },
      options: optConc,
      plugins: [linhaRef(vo, vo != null ? `VI ${fmt(vo)} ${un}` : '')]
    });
    $('#g2-sub').textContent = `${un} · escala log · linha = VI`;
  }

  // ------------------------------------------------------------- perfil do poço (popup)
  // Mesma regra de cor do Perfil de Sondagem (colorFromText), para as camadas ficarem iguais.
  function corLitologia(cor, lit) {
    const t = String((cor || '') + ' ' + (lit || '')).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
    if (/asfalto/.test(t)) return '#3B3A38';
    if (/marrom escur/.test(t)) return '#5A3A22';
    if (/cinza escur/.test(t)) return '#5C645D';
    if (/preta|preto/.test(t)) return '#2A2521';
    if (/marrom/.test(t)) return '#8A5A34';
    if (/amarel/.test(t)) return '#C9A227';
    if (/cinza/.test(t)) return '#8D958D';
    if (/vermelh/.test(t)) return '#A34632';
    if (/roxo|roxa|lilas/.test(t)) return '#7A5C88';
    if (/verde/.test(t)) return '#5E7A4E';
    if (/laranja/.test(t)) return '#C07A32';
    if (/bege/.test(t)) return '#C9B48A';
    if (/branc/.test(t)) return '#E7E3D8';
    return '#9C8F78';
  }

  /** Desenha o perfil: litologia e tubo da ficha do Perfil (se houver) ou da aba Litologia. */
  function perfilSVG(p, med, fichaDireta) {
    const num = WGImport.numero;
    const fichas = fichaDireta ? [fichaDireta] : (st.sondagens.get(codigo(p.codigo)) || []);
    const ficha = fichas.length ? fichas[fichas.length - 1] : null; // a mais recente
    const dados = ficha?.data || {};
    const wg = st.perfis.get(p.id) || [];

    let camadas = (dados.litologia || []).map(r => ({ de: num(r.inicio), ate: num(r.termino), texto: [r.litologia, r.cor].filter(Boolean).join(' · '), hex: corLitologia(r.cor, r.litologia) }))
      .filter(r => r.de != null && r.ate != null && r.ate > r.de);
    let fonte = ficha ? `ficha ${ficha.sondagem_no || ''} do Perfil` : '';
    if (!camadas.length) {
      camadas = wg.filter(r => r.estrutura === 'SOLO').map(r => ({ de: +r.de_m, ate: +r.ate_m, texto: r.descricao || '', hex: r.cor_hex || '#9C8F78' }));
      fonte = camadas.length ? 'aba Litologia da planilha' : fonte;
    }
    let tubo = [];
    const liso = num(dados.perfil?.tuboLiso), filtro = num(dados.perfil?.tuboFiltro);
    if (liso != null && filtro != null && liso + filtro > 0) tubo = [{ tipo: 'CEGO', de: 0, ate: liso }, { tipo: 'FILTRO', de: liso, ate: liso + filtro }];
    else tubo = wg.filter(r => r.estrutura === 'CEGO' || r.estrutura === 'FILTRO').map(r => ({ tipo: r.estrutura, de: +r.de_m, ate: +r.ate_m }));
    const anular = wg.filter(r => r.estrutura === 'SELO' || r.estrutura === 'BRITA').map(r => ({ tipo: r.estrutura, de: +r.de_m, ate: +r.ate_m, hex: r.cor_hex }));
    const na = med?.nivel_agua != null ? +med.nivel_agua : num(dados.perfil?.naEstabilizado);
    if (!camadas.length && !tubo.length) return '';

    const prof = Math.max(...camadas.map(r => r.ate), ...tubo.map(r => r.ate), ...anular.map(r => r.ate), na || 0, 1);
    const H = 170, top = 8, W = 270, esc_ = H / prof;
    const y = d => top + d * esc_;
    const passo = prof <= 3 ? .5 : prof <= 8 ? 1 : prof <= 15 ? 2 : 5;
    let svg = `<svg viewBox="0 0 ${W} ${H + top + 10}" role="img" aria-label="Perfil do poço ${esc(p.codigo)}" font-family="IBM Plex Sans, sans-serif">`;
    for (let d = 0; d <= prof + 1e-9; d += passo) {
      svg += `<line x1="26" x2="30" y1="${y(d)}" y2="${y(d)}" stroke="#8B978F"/><text x="23" y="${y(d) + 3}" font-size="8.5" text-anchor="end" fill="#56655D">${fmt(d, 1)}</text>`;
    }
    svg += `<line x1="30" x2="30" y1="${top}" y2="${y(prof)}" stroke="#8B978F"/>`;
    camadas.forEach(r => {
      svg += `<rect x="32" y="${y(r.de)}" width="52" height="${Math.max(.5, (r.ate - r.de) * esc_)}" fill="${esc(r.hex)}" stroke="#ffffff" stroke-width=".6"/>`;
      if ((r.ate - r.de) * esc_ >= 10 && r.texto) {
        const t = r.texto.length > 30 ? r.texto.slice(0, 29) + '…' : r.texto;
        svg += `<text x="90" y="${y((r.de + r.ate) / 2) + 3}" font-size="9" fill="#172621">${esc(t)}</text>`;
      }
    });
    anular.forEach(r => {
      svg += `<rect x="47" y="${y(r.de)}" width="22" height="${(r.ate - r.de) * esc_}" fill="${r.tipo === 'SELO' ? (r.hex || '#556b2f') : (r.hex || '#9e9e9e')}" opacity=".85"/>`;
    });
    tubo.forEach(r => {
      const h = (r.ate - r.de) * esc_;
      svg += r.tipo === 'FILTRO'
        ? `<rect x="53" y="${y(r.de)}" width="10" height="${h}" fill="#DDEAF6" stroke="#2B7BBA" stroke-width="1"/>` +
          Array.from({ length: Math.floor(h / 4) }, (_, k) => `<line x1="54" x2="62" y1="${y(r.de) + 2 + k * 4}" y2="${y(r.de) + 2 + k * 4}" stroke="#2B7BBA" stroke-width=".8"/>`).join('')
        : `<rect x="53" y="${y(r.de)}" width="10" height="${h}" fill="#F3F5F4" stroke="#56655D" stroke-width="1"/>`;
    });
    if (na != null && na <= prof) {
      svg += `<line x1="32" x2="84" y1="${y(na)}" y2="${y(na)}" stroke="#2B7BBA" stroke-width="1.5" stroke-dasharray="3 2"/>` +
        `<path d="M68 ${y(na) - 7} h8 l-4 6 z" fill="#2B7BBA"/>`;
    }
    svg += `</svg>`;
    const legenda = [tubo.length ? 'tubo liso / <span style="color:#2B7BBA">filtro</span>' : '', na != null ? `<span style="color:#2B7BBA">▼ N.A. ${fmt(na)} m</span>` : ''].filter(Boolean).join(' · ');
    return `<div class="pop-perfil"><div class="pop-perfil-tit">Perfil · ${esc(fonte)}</div>${svg}${legenda ? `<div class="pop-mini">${legenda}</div>` : ''}</div>`;
  }

  /** Popup de uma ficha do Perfil que não tem poço de monitoramento. */
  function popupSondagem(f) {
    const m = f.data?.meta || {};
    return `
      <div class="pop">
        <div class="pop-tit">${esc(f.sondagem_no || 'Sondagem')} <span class="st-tag st-sem">Sondagem</span></div>
        <div class="pop-sond">Ficha do Perfil de Sondagem${f.obra ? ' · ' + esc(f.obra) : ''}</div>
        <div class="pop-mini">${m.dataInicio ? 'Início ' + fmtData(m.dataInicio) + ' · ' : ''}${esc(f.data?.local?.endereco || '')}</div>
        <div class="pop-mini">Lat ${fmt(f.latitude, 6)} · Lon ${fmt(f.longitude, 6)}</div>
        ${perfilSVG({ id: null, codigo: f.sondagem_no || '' }, null, f)}
      </div>`;
  }

  // ------------------------------------------------------------- tabela
  function linhasTabela() {
    const { pocos, porPoco } = selecao();
    return pocos.map(p => ({ p, r: porPoco.get(p.id) }));
  }

  function renderTabela(param) {
    $('#tab-param').textContent = param ? `${param.nome} (${param.unidade})` : 'Resultado';
    $('#tabela-dados tbody').innerHTML = linhasTabela().map(({ p, r }) => {
      const s = status(r);
      const fichas = st.sondagens.get(codigo(p.codigo)) || [];
      return `<tr>
        <td class="cod">${esc(p.codigo)}</td><td>${esc(p.rede || '')}</td>
        <td class="num">${p.latitude !== null ? fmt(p.latitude, 6) : '<span class="muted">sem coord.</span>'}</td>
        <td class="num">${p.longitude !== null ? fmt(p.longitude, 6) : ''}</td>
        <td class="num">${textoValor(r)}</td>
        <td class="num">${param?.valor_orientador != null ? fmt(param.valor_orientador) : '-'}</td>
        <td><span class="st-tag st-${s}">${ROTULO_STATUS[s]}</span></td>
        <td>${esc(r?.laudo || '')}</td>
        <td>${fichas.map(f => esc(f.sondagem_no || 'sim')).join(', ')}</td></tr>`;
    }).join('') || '<tr><td colspan="9" class="muted">Sem poços para mostrar.</td></tr>';
  }

  function exportarCSV() {
    const param = st.parametros.find(p => p.id === st.parametroId);
    const camp = st.campanhas.find(c => c.id === st.campanhaId);
    if (!param || !camp) { toast('Nada para exportar ainda.'); return; }
    const cab = ['Poço', 'Rede', 'Latitude', 'Longitude', 'Campanha', 'Parâmetro', 'Resultado', 'Unidade', 'Abaixo do LQ', 'VI', 'Situação'];
    const lin = linhasTabela().map(({ p, r }) => [p.codigo, p.rede, p.latitude ?? '', p.longitude ?? '', camp.codigo, param.nome,
      r ? (r.menor_que_lq ? (r.lq ?? 'LQ') : r.valor) : '', param.unidade, r?.menor_que_lq ? 'sim' : '', param.valor_orientador ?? '', ROTULO_STATUS[status(r)]]);
    const csv = [cab, ...lin].map(l => l.map(v => `"${String(v ?? '').replace(/"/g, '""')}"`).join(';')).join('\r\n');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' }));
    a.download = `webgeo_${param.nome}_${camp.codigo}.csv`.replace(/\s+/g, '_');
    a.click();
    URL.revokeObjectURL(a.href);
  }

  // ------------------------------------------------------------- upload da planilha
  async function lerArquivo(arquivo) {
    limparImportacao();
    $('#imp-status').textContent = 'Lendo ' + arquivo.name + '...';
    try {
      const wb = XLSX.read(await arquivo.arrayBuffer(), { type: 'array', cellDates: true });
      const pj = WGImport.lerProjeto(XLSX, wb);
      // Projeto de destino: o nome escrito na planilha (se já existe na empresa, usa ele;
      // se não, será criado) ou, com o campo vazio, o projeto escolhido no menu.
      const nomeK = s => String(s || '').trim().toLowerCase();
      let destino = null, novo = false;
      if (pj.projeto) {
        const achado = st.projetos.find(p => nomeK(p.nome) === nomeK(pj.projeto));
        if (achado) destino = { id: achado.id, nome: achado.nome }; else novo = true;
      } else if (st.projetoId) {
        const atual = st.projetos.find(p => p.id === st.projetoId);
        destino = { id: atual.id, nome: atual.nome };
      }
      const existentes = new Set();
      if (destino) (await buscarTudo(() => sb.from('wg_poco').select('codigo').eq('projeto_id', destino.id).order('codigo')))
        .forEach(p => existentes.add(p.codigo));
      const { payload, rel } = WGImport.ler(XLSX, wb, arquivo.name, { parametros: st.parametros, pocosExistentes: existentes, projetoDestino: destino });
      st.importacao = { payload, rel, novo };
      mostrarRelatorio(rel, payload, novo);
    } catch (e) {
      console.error(e);
      $('#imp-status').innerHTML = '<span class="erro">Não consegui ler o arquivo:</span> ' + esc(e.message);
    }
  }

  function mostrarRelatorio(rel, payload, novo) {
    const p = payload.projeto || {};
    $('#imp-status').textContent = '';
    $('#imp-resumo').innerHTML = `
      <div class="imp-cab">
        <div><span>Arquivo</span><b>${esc(payload.arquivo)}</b></div>
        <div><span>Projeto</span><b>${esc(p.nome || '?')}</b>${p.nome ? `<em class="imp-proj">${novo ? 'projeto novo: será criado (aparece também no Perfil)' : 'projeto existente'}</em>` : ''}</div>
        <div><span>Campanha</span><b>${payload.campanha ? esc(payload.campanha.codigo) + ' · ' + fmtData(payload.campanha.data_inicio) : '-'}</b></div>
      </div>
      <table class="imp-cont">${Object.entries(rel.contadores).map(([k, v]) => `<tr><td>${esc(k)}</td><td class="num">${v}</td></tr>`).join('')}</table>`;
    const item = x => `<li><span class="loc">${x.aba ? esc(x.aba) + (x.linha ? ', linha ' + x.linha : '') : ''}</span> ${esc(x.msg)}</li>`;
    $('#imp-erros').innerHTML = rel.erros.length ? `<h4>Erros (${rel.erros.length}): corrija na planilha e envie de novo</h4><ul>${rel.erros.map(item).join('')}</ul>` : '';
    $('#imp-avisos').innerHTML = rel.avisos.length ? `<h4>Avisos (${rel.avisos.length})</h4><ul>${rel.avisos.map(item).join('')}</ul>` : '';
    $('#imp-acoes').hidden = false;
    $('#btn-gravar').disabled = rel.erros.length > 0;
    $('#btn-gravar').textContent = rel.erros.length ? 'Corrija os erros para gravar' : 'Gravar no banco';
    $('#imp-relatorio').hidden = false;
  }

  async function gravarImportacao() {
    if (!st.importacao || st.importacao.rel.erros.length) return;
    const btn = $('#btn-gravar'); btn.disabled = true; btn.textContent = 'Gravando...';
    const { data, error } = await sb.rpc('wg_importar', { p: st.importacao.payload });
    if (error) {
      btn.disabled = false; btn.textContent = 'Gravar no banco';
      $('#imp-status').innerHTML = '<span class="erro">O banco recusou a importação (nada foi gravado):</span> ' + esc(error.message);
      return;
    }
    $('#imp-status').innerHTML = '<span class="ok">Importação gravada.</span> Os dados já estão no mapa.';
    $('#imp-acoes').hidden = true;
    toast('Importação gravada.');
    st.projetoId = data.projeto_id; lembrarProjeto(data.projeto_id);
    if (data.campanha_id) st.campanhaId = data.campanha_id;
    st.importacao = null;
    await carregarProjetos();
  }

  function limparImportacao() {
    st.importacao = null;
    $('#arquivo').value = '';
    $('#imp-relatorio').hidden = true;
    $('#imp-status').textContent = '';
  }

  return { iniciar, _teste: { mapa: () => mapa, estado: st } }; // _teste: usado só pelos testes automáticos
})();
