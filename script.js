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
   Interpolação IDW numa grade em metros, área acima do limiar,
   centróide e linha do limiar (marching squares).
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
   * A grade cobre a rede de poços + uma margem.
   */
  function grade(pocos, cel) {
    if (!pocos.length) return null;
    const lat0 = pocos.reduce((s, p) => s + p.lat, 0) / pocos.length;
    const lon0 = pocos.reduce((s, p) => s + p.lon, 0) / pocos.length;
    const proj = projecao(lat0, lon0);
    const pts = pocos.map(p => proj.paraXY(p.lat, p.lon));
    const margem = Math.max(2 * cel, 10);
    const xmin = Math.floor((Math.min(...pts.map(p => p.x)) - margem) / cel) * cel;
    const ymin = Math.floor((Math.min(...pts.map(p => p.y)) - margem) / cel) * cel;
    const xmax = Math.ceil((Math.max(...pts.map(p => p.x)) + margem) / cel) * cel;
    const ymax = Math.ceil((Math.max(...pts.map(p => p.y)) + margem) / cel) * cel;
    const nx = Math.max(1, Math.round((xmax - xmin) / cel));
    const ny = Math.max(1, Math.round((ymax - ymin) / cel));

    // Máscara: só calcula onde há poço por perto (dentro da envoltória da rede
    // ou a até "margem" metros de um poço). Fora disso o IDW só repetiria a média.
    const env = envoltoria(pts);
    const dentro = new Uint8Array(nx * ny);
    for (let j = 0; j < ny; j++) for (let i = 0; i < nx; i++) {
      const x = xmin + (i + .5) * cel, y = ymin + (j + .5) * cel;
      let ok = env.length >= 3 && dentroPoligono(x, y, env);
      if (!ok) for (const p of pts) { if ((p.x - x) ** 2 + (p.y - y) ** 2 <= margem * margem) { ok = true; break; } }
      dentro[j * nx + i] = ok ? 1 : 0;
    }
    return { proj, cel, xmin, ymin, nx, ny, dentro, margem };
  }

  /**
   * IDW: cada célula recebe a média dos N poços mais próximos (padrão 12),
   * ponderada por 1/distância^p. Usar só os vizinhos evita que um poço muito
   * contaminado "pinte" a área inteira com valores pequenos.
   */
  function idw(g, amostras, p, vizinhos = 12) {
    const v = new Float64Array(g.nx * g.ny).fill(NaN);
    const pts = amostras.map(a => ({ ...g.proj.paraXY(a.lat, a.lon), v: a.v }));
    if (!pts.length) return v;
    const N = Math.min(vizinhos, pts.length);
    const dist = new Float64Array(N), val = new Float64Array(N);
    for (let j = 0; j < g.ny; j++) for (let i = 0; i < g.nx; i++) {
      const k = j * g.nx + i;
      if (!g.dentro[k]) continue;
      const x = g.xmin + (i + .5) * g.cel, y = g.ymin + (j + .5) * g.cel;
      let n = 0, exato = null;
      for (const a of pts) {                       // guarda os N mais próximos (inserção ordenada)
        const d2 = (a.x - x) ** 2 + (a.y - y) ** 2;
        if (d2 < 1e-6) { exato = a.v; break; }
        if (n < N || d2 < dist[n - 1]) {
          let q = n < N ? n++ : n - 1;
          while (q > 0 && dist[q - 1] > d2) { dist[q] = dist[q - 1]; val[q] = val[q - 1]; q--; }
          dist[q] = d2; val[q] = a.v;
        }
      }
      if (exato !== null) { v[k] = exato; continue; }
      let num = 0, den = 0;
      for (let q = 0; q < n; q++) { const w = 1 / Math.pow(dist[q], p / 2); num += w * val[q]; den += w; }
      v[k] = num / den;
    }
    return v;
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
   * Devolve segmentos [[lat,lon],[lat,lon]] para desenhar no mapa.
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
      const t = (!Number.isFinite(v1) || !Number.isFinite(v2)) ? .5 : (limiar - v1) / (v2 - v1);
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
    return segs.map(s => s.map(p => { const ll = g.proj.paraLatLon(p.x, p.y); return [ll.lat, ll.lon]; }));
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
    const f = t * (RAMPA.length - 1), k = Math.min(RAMPA.length - 2, Math.floor(f)), r = f - k;
    const a = hexRgb(RAMPA[k]), b = hexRgb(RAMPA[k + 1]);
    return a.map((x, n) => Math.round(x + (b[n] - x) * r));
  }

  return { grade, idw, areaECentroide, distancia, isolinha, cor, RAMPA, projecao };
})();

if (typeof module !== 'undefined') module.exports = { WGImport, WGPluma };



/* =====================================================================
   3. APLICAÇÃO (tela)
   Login e empresa seguem exatamente o Perfil de Sondagem:
   auth -> profiles.organization_id -> organizations.
   ===================================================================== */
if (typeof window !== 'undefined' && window.document) {
  document.addEventListener('DOMContentLoaded', () => WebGeo.iniciar());
}

const WebGeo = (() => {
  let sb;                       // cliente Supabase
  let mapa, camadaPocos, camadaPluma;
  let graficos = {};
  let usuarioCarregado = null;  // evita recarregar a tela quando o Supabase só renova o token
  const st = {
    user: null, profile: null, org: null,
    projetos: [], projetoId: null, soltas: [], parametros: [], campanhas: [], pocos: [], resultados: [], medicoes: [],
    sondagens: new Map(),       // código do poço (PM-30) -> fichas do Perfil
    parametroId: null, campanhaId: null, redesOcultas: new Set(), importacao: null,
    pontos: [], enquadrarPendente: false,
    perfis: new Map(),            // poco_id -> intervalos da aba Litologia (wg_perfil_poco)
    pluma: { ativa: true, limiar: null, p: 2, cel: 10 }, // limiar null = usa o VI do parâmetro
    calc: null                    // resultado da última interpolação
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
    $('#f-pluma').addEventListener('change', e => { st.pluma.ativa = e.target.checked; render(); });
    $('#f-idw').addEventListener('change', e => { st.pluma.p = +e.target.value; render(); });
    $('#f-celula').addEventListener('change', e => { st.pluma.cel = +e.target.value; render(); });
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

    renderKPIs(param);
    renderMapa(param, enquadrar);
    renderPluma(param);
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
    mapa.createPane('plumaLinha').style.zIndex = 380;
    camadaPluma = L.layerGroup().addTo(mapa);
    camadaPocos = L.layerGroup().addTo(mapa);
    mapa.on('zoomend', atualizarRotulos);
    atualizarRotulos();
  }

  /** Abaixo do zoom 19 os nomes se sobrepõem: mostra só os poços acima do VI. */
  function atualizarRotulos() {
    mapa.getContainer().classList.toggle('rotulos-ocultos', mapa.getZoom() < 19);
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
    const ordem = { sem: 0, lq: 1, abaixo: 2, acima: 3 }; // os acima do VI ficam por cima
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
          .bindTooltip(esc(p.codigo), { permanent: true, direction: 'right', offset: [8, 0], className: 'rotulo-pm rotulo-' + s })
          .bindPopup(() => popupPoco(p, param), { maxWidth: 320 })
          .on('popupopen', () => { st.popupAberto = p.id; })
          .on('popupclose', () => { if (st.popupAberto === p.id && !st.redesenhando) st.popupAberto = null; })
          .addTo(camadaPocos);
        marcadores.set(p.id, m);
      });
    // fichas do Perfil sem poço: quadrado verde com o nº da sondagem
    if (!st.redesOcultas.has('Sondagens')) st.soltas.forEach(f => {
      const ll = [f.latitude, f.longitude];
      pts.push(ll);
      const m = L.marker(ll, { icon: L.divIcon({ className: '', html: '<div class="ponto-sondagem"></div>', iconSize: [12, 12] }) })
        .bindTooltip(esc(f.sondagem_no || 'Sondagem'), { permanent: true, direction: 'right', offset: [8, 0], className: 'rotulo-pm rotulo-sond' })
        .bindPopup(() => popupSondagem(f), { maxWidth: 320 })
        .on('popupopen', () => { st.popupAberto = 'ficha:' + f.id; })
        .on('popupclose', () => { if (st.popupAberto === 'ficha:' + f.id && !st.redesenhando) st.popupAberto = null; })
        .addTo(camadaPocos);
      marcadores.set('ficha:' + f.id, m);
    });
    st.pontos = pts;
    // reabre o popup que estava aberto — só com o mapa visível (escondido, o Leaflet erra a largura)
    if (aberto && marcadores.has(aberto) && mapaVisivel()) marcadores.get(aberto).openPopup();
    else if (aberto) st.popupAberto = null;
    if (enquadrar) enquadrarMapa();
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
   * para dar para comparar área e centróide entre elas.
   * Resultado abaixo do LQ entra como 0. Respeita o filtro de rede.
   */
  function calcularPlumas(param, limiar) {
    if (!param) return null;
    const visivel = p => !st.redesOcultas.has(p.rede || '-');
    const pocosCoord = new Map(st.pocos.filter(p => p.latitude !== null && visivel(p)).map(p => [p.id, p]));
    const res = st.resultados.filter(r => r.parametro_id === param.id && pocosCoord.has(r.poco_id));
    if (!res.length) return null;
    const ids = [...new Set(res.map(r => r.poco_id))];
    const g = WGPluma.grade(ids.map(id => ({ lat: +pocosCoord.get(id).latitude, lon: +pocosCoord.get(id).longitude })), st.pluma.cel);
    let maxGlobal = 0;
    const porCamp = st.campanhas.map(c => {
      const rc = res.filter(r => r.campanha_id === c.id);
      const quant = rc.filter(r => !r.menor_que_lq).map(r => +r.valor);
      const max = quant.length ? Math.max(...quant) : null;
      if (max) maxGlobal = Math.max(maxGlobal, max);
      const item = { campanha: c, n: rc.length, max, nAcima: rc.filter(r => r.acima_vi).length, v: null, area: null, centroide: null };
      if (rc.length >= 3) {
        const p = pocosCoord;
        item.v = WGPluma.idw(g, rc.map(r => ({ lat: +p.get(r.poco_id).latitude, lon: +p.get(r.poco_id).longitude, v: +r.valor_calculo || 0 })), st.pluma.p);
        if (limiar != null) Object.assign(item, WGPluma.areaECentroide(g, item.v, limiar));
      }
      return item;
    });
    const vmax = Math.pow(10, Math.max(1, Math.ceil(Math.log10(maxGlobal || 10))));
    // cor só a partir de 1/10 do limiar (ou 3 ordens abaixo do máximo, sem limiar)
    const vmin = Math.min(vmax / 10, limiar != null ? limiar / 10 : vmax / 1e3);
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

    // 1) imagem da pluma: um pixel por célula, o navegador suaviza ao ampliar
    const cv = document.createElement('canvas');
    cv.width = g.nx; cv.height = g.ny;
    const ctx = cv.getContext('2d');
    const img = ctx.createImageData(g.nx, g.ny);
    for (let j = 0; j < g.ny; j++) for (let i = 0; i < g.nx; i++) {
      const val = sel.v[j * g.nx + i];
      const o = ((g.ny - 1 - j) * g.nx + i) * 4; // linha 0 da imagem = norte
      if (!(val >= vmin)) { img.data[o + 3] = 0; continue; }
      const [r, gg, b] = WGPluma.cor(val, vmin, vmax);
      img.data[o] = r; img.data[o + 1] = gg; img.data[o + 2] = b; img.data[o + 3] = 185;
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
    $('#leg-gradiente').style.background = `linear-gradient(to right, ${WGPluma.RAMPA.join(', ')})`;
    const pos = v => 100 * Math.log10(v / vmin) / Math.log10(vmax / vmin);
    const ticks = [];
    for (let e = Math.ceil(Math.log10(vmin) - 1e-9); Math.pow(10, e) <= vmax * 1.0001; e++) ticks.push(Math.pow(10, e));
    $('#leg-ticks').innerHTML = ticks.map(v => `<span style="left:${pos(v)}%">${fmt(v, 3)}</span>`).join('');
    $('#leg-limiar').textContent = calc.limiar != null ? `Limiar (${fmt(calc.limiar, 4)} ${un})` : 'Limiar: sem VI, defina no menu';
  }

  function renderEvolucao(param) {
    const calc = st.calc, un = param?.unidade || 'µg/L';
    const sel = plumaDaCampanha(st.campanhaId), ini = primeiraComArea();
    const fa = a => a == null ? '—' : fmt(Math.round(a), 0);
    const semLimiar = !calc || calc.limiar == null;
    $('#pluma-hint').textContent = calc
      ? `IDW p = ${st.pluma.p} (12 poços mais próximos) · célula ${st.pluma.cel} m · limiar ${semLimiar ? '—' : fmt(calc.limiar, 4) + ' ' + un} · calculada dentro da rede de poços`
      : 'sem resultados deste parâmetro';

    $('#kpi-a1-lab').textContent = ini ? `Área na ${ini.campanha.codigo}` : 'Área na 1ª campanha';
    $('#kpi-a2-lab').textContent = sel ? `Área na ${sel.campanha.codigo}` : 'Área na campanha';
    $('#kpi-a1').textContent = fa(ini?.area);
    $('#kpi-a2').textContent = fa(sel?.area);
    const sub = semLimiar ? 'defina o limiar' : `acima de ${fmt(calc.limiar, 4)} ${un}`;
    $('#kpi-a1-sub').textContent = ini ? sub : 'precisa de 3+ poços com resultado';
    $('#kpi-a2-sub').textContent = sel?.v ? sub : 'precisa de 3+ poços com resultado';

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

  return { iniciar };
})();
