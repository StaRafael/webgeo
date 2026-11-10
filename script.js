/* =====================================================================
   WebGeo - portal de monitoramento ambiental
   JavaScript puro, sem build. Bibliotecas por CDN (index.html):
   Supabase JS, Leaflet (mapa) e SheetJS (leitura do Excel).
   ===================================================================== */

// ---------------------------------------------------------------------
// Conexão com o Supabase: o MESMO projeto do Perfil de Sondagem (mesmo
// login, mesma empresa). Endereço e chave vêm do config.js.
// A chave publishable é pública por natureza — quem protege os
// dados de cada empresa são as regras de segurança (RLS) do banco.
// ---------------------------------------------------------------------
const SUPABASE_URL = window.APP_CONFIG.supabaseUrl;
const SUPABASE_ANON_KEY = window.APP_CONFIG.supabaseAnonKey;


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
      dataInicioBruta: campo('Data início'), dataFimBruta: campo('Data fim'), fuso: texto(campo('Fuso')), somarE: campo('Somar ao E'),
      laboratorio: texto(campo('Laboratório'))
    };
  }

  const COL_POCO = ['Poço', 'Poco', 'ID_POCO', 'Ponto'];
  const COL_LAT = ['Latitude', 'Latitude_Corrigida', 'Lat'];
  const COL_LON = ['Longitude', 'Longitude_Corrigida', 'Lon'];
  const COL_COTA = ['Cota topo (m)', 'Cota (m)', 'Cota', 'Cota topo'];
  const COL_E = ['E (m)', 'E', 'Este', 'Leste', 'UTM E', 'UTM_E', 'Coord E', 'Coordenada E', 'X (m)', 'X'];
  const COL_N = ['N (m)', 'N', 'Norte', 'UTM N', 'UTM_N', 'Coord N', 'Coordenada N', 'Y (m)', 'Y'];
  const COL_FUSO = ['Fuso', 'Fuso UTM', 'Zona', 'Zona UTM'];
  // fuso UTM mais comum de cada UF (quando a planilha não diz o fuso e o projeto ainda não tem poço com coordenada)
  const FUSO_UF = { AC: 19, AM: 20, RO: 20, RR: 20, MT: 21, MS: 21, PA: 22, GO: 22, TO: 22, PR: 22, SC: 22, RS: 22,
    AP: 22, MA: 23, PI: 23, DF: 23, MG: 23, SP: 23, RJ: 23, CE: 24, RN: 24, PB: 25, PE: 25, AL: 24, SE: 24, BA: 24, ES: 24 };

  // centro aproximado de cada UF [lat, lon]: última referência para completar um E com 5 dígitos
  const CENTRO_UF = { AC: [-9, -70.5], AL: [-9.6, -36.6], AP: [1.4, -51.8], AM: [-4.2, -64.6], BA: [-12.5, -41.7], CE: [-5.2, -39.5],
    DF: [-15.8, -47.9], ES: [-19.6, -40.7], GO: [-16, -49.6], MA: [-5.1, -45.3], MT: [-12.9, -55.9], MS: [-20.5, -54.8], MG: [-18.5, -44.5],
    PA: [-3.8, -52.5], PB: [-7.1, -36.8], PR: [-24.6, -51.6], PE: [-8.4, -37.9], PI: [-7.7, -42.7], RJ: [-22.3, -42.7], RN: [-5.8, -36.6],
    RS: [-29.8, -53.3], RO: [-10.9, -63], RR: [2.1, -61.4], SC: [-27.2, -50.4], SP: [-23.0, -47.5], SE: [-10.6, -37.4], TO: [-10.2, -48.3] };
  const WGMalha_texto = v => Number(v).toLocaleString('pt-BR', { maximumFractionDigits: 0 });

  /** "23", "23S", "23 K", "fuso 22" -> {zona, sul} (sul por padrão; letra N ou de N a X = norte, exceto S). null se não deu. */
  function lerFuso(v) {
    const m = /(\d{1,2})\s*([A-Za-z])?/.exec(String(v ?? ''));
    if (!m || +m[1] < 1 || +m[1] > 60) return null;
    const L = (m[2] || 'S').toUpperCase();
    return { zona: +m[1], sul: L === 'S' || (L >= 'C' && L <= 'M') };
  }

  /**
   * Coordenada de um poço em graus decimais OU em UTM (metros). Aceita lat/lon nas colunas Latitude/Longitude,
   * E/N nas colunas E/N (ou X/Y), e até E e N digitados nas colunas Latitude/Longitude (em qualquer ordem).
   * Devolve {lat, lon, utm: {E, N}} ou {erro} ou null (vazio).
   */
  function coordPoco(vLat, vLon, vE, vN) {
    const a = numero(vLat), b = numero(vLon), e = numero(vE), n = numero(vN);
    const vazio = x => x === null;
    if (vazio(a) && vazio(b) && vazio(e) && vazio(n)) return null;
    const pareceGrau = (la, lo) => la !== null && lo !== null && Math.abs(la) <= 90 && Math.abs(lo) <= 180;
    const comoUtm = (x, y) => {                          // x, y em qualquer ordem: N tem 7 dígitos, E tem 6
      if (x === null || y === null) return null;
      const [E, N] = Math.abs(x) >= 1e6 && Math.abs(y) < 1e6 ? [y, x] : [x, y];
      return E >= 100000 && E <= 900000 && N >= 0 && N <= 10000000 ? { E, N } : null;
    };
    if (pareceGrau(a, b)) return { lat: a, lon: b };
    const u = comoUtm(e, n) || comoUtm(a, b);
    if (u) return { utm: u };
    // E com 5 dígitos (ex.: 37859,50 em vez de 337859,50): planta de CAD que tirou a centena de milhar do E.
    // O dígito que falta é descoberto depois, pela posição de referência (ver "E incompleto" em ler)
    const comoCurto = (x, y) => {
      if (x === null || y === null) return null;
      const [E, N] = Math.abs(x) >= 1e6 && Math.abs(y) < 1e6 ? [y, x] : [x, y];
      return E >= 0 && E < 100000 && N >= 1e6 && N <= 10000000 ? { E, N } : null;
    };
    const c = comoCurto(e, n) || comoCurto(a, b);
    if (c) return { utm: c, curto: true };
    return { erro: true };
  }
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
    const rel = { contadores: {}, avisos: [], erros: [], convertidos: [] };
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
    const utmPendentes = [];                     // poços em UTM: convertidos depois, quando já se sabe o fuso
    const tPocos = tabela(XLSX, wb, 'Pocos');
    if (tPocos) for (const l of tPocos.linhas) {
      const cod = codigoPoco(l.get(...COL_POCO));
      if (!cod) { erro(tPocos.nome, l.n, 'código do poço vazio'); continue; }
      if (pocosPlanilha.has(cod)) { erro(tPocos.nome, l.n, `${cod} aparece mais de uma vez`); continue; }
      pocosPlanilha.add(cod);

      const vLat = l.get(...COL_LAT), vLon = l.get(...COL_LON);
      if ([vLat, vLon].some(v => typeof v === 'string' && (v.match(/\./g) || []).length > 1 && Math.abs(numero(v) ?? 0) < 1000)) {
        erro(tPocos.nome, l.n, `${cod}: coordenada mal formatada '${texto(vLat)}' / '${texto(vLon)}' (vários pontos). Use -23.471383 ou -23,471383`); continue;
      }
      const cp = coordPoco(vLat, vLon, l.get(...COL_E), l.get(...COL_N));
      if (cp && cp.erro) {
        erro(tPocos.nome, l.n, `${cod}: coordenada inválida '${texto(vLat) ?? texto(l.get(...COL_E))}' / '${texto(vLon) ?? texto(l.get(...COL_N))}'. Use graus decimais (ex.: -23.471383 / -46.575588) ou UTM em metros (ex.: E 338979 / N 7403370)`);
        continue;
      }
      if (!cp) aviso(tPocos.nome, l.n, `${cod} sem coordenadas: não aparece no mapa até ser preenchido`);
      else if (cp.lat == null && (cp.utm.E == null || cp.utm.N == null)) { erro(tPocos.nome, l.n, `${cod}: informe E e N (ou deixe os dois vazios)`); continue; }
      const lat = cp ? cp.lat ?? null : null, lon = cp ? cp.lon ?? null : null;
      if (cp && cp.utm) utmPendentes.push({ poco: null, utm: cp.utm, E0: cp.utm.E, N0: cp.utm.N, curto: !!cp.curto, fuso: lerFuso(texto(l.get(...COL_FUSO))), aba: tPocos.nome, n: l.n, cod });

      let situacao = texto(l.get('Situação', 'Situacao'));
      if (situacao) {
        situacao = chave(situacao);
        if (situacao === 'DESTRUIDO') situacao = 'DANIFICADO';   // o banco guarda destruído como DANIFICADO
        if (!SITUACOES.includes(situacao)) { erro(tPocos.nome, l.n, 'situação inválida (use ATIVO, SECO, DANIFICADO/DESTRUÍDO ou DESATIVADO)'); continue; }
      }
      payload.pocos.push({
        codigo: cod, rede: (texto(l.get('Rede')) || redeDoCodigo(cod)).toUpperCase(),
        latitude: lat, longitude: lon, cota_topo: numero(l.get(...COL_COTA)),
        altitude: numero(l.get('Altitude (m)', 'Altitude')), profundidade: numero(l.get('Profundidade (m)', 'Profundidade')),
        situacao, observacao: texto(l.get('Observação', 'Observacao', 'Obs'))
      });
      if (cp && cp.utm) utmPendentes[utmPendentes.length - 1].poco = payload.pocos[payload.pocos.length - 1];
      contar(pocosExistentes.has(cod) ? 'Poços atualizados' : 'Poços novos');
    }

    // ---- UTM -> graus. Fuso: o da linha; senão o da aba Projeto; senão o dos poços em graus (planilha ou projeto); senão o da UF
    if (utmPendentes.length) {
      const graus = payload.pocos.filter(p => p.latitude !== null && p.longitude !== null);
      const media = arr => arr.reduce((x, y) => x + y, 0) / arr.length;
      let padrao = lerFuso(pj.fuso), origem = 'aba Projeto';
      if (!padrao && graus.length) { padrao = { zona: WGDxf.zonaDaLongitude(media(graus.map(p => p.longitude))), sul: media(graus.map(p => p.latitude)) < 0 }; origem = 'poços da planilha em graus'; }
      if (!padrao && ctx.fusoProjeto) { padrao = ctx.fusoProjeto; origem = 'poços já cadastrados no projeto'; }
      if (!padrao && ctx.referencia) { padrao = { zona: WGDxf.zonaDaLongitude(ctx.referencia.lon), sul: ctx.referencia.lat < 0 }; origem = 'outros projetos da empresa'; }
      if (!padrao && pj.uf && FUSO_UF[String(pj.uf).trim().toUpperCase()]) { padrao = { zona: FUSO_UF[String(pj.uf).trim().toUpperCase()], sul: true }; origem = `UF ${String(pj.uf).trim().toUpperCase()}`; }
      if (!padrao) { padrao = { zona: 23, sul: true }; origem = 'padrão'; }
      // E incompleto (5 dígitos): soma a centena de milhar que põe os poços mais perto da referência
      const curtos = utmPendentes.filter(q => q.curto);
      if (curtos.length) {
        const med = arr => { const o = [...arr].sort((x, y) => x - y); return o[Math.floor(o.length / 2)]; };
        const Em = med(curtos.map(q => q.utm.E)), Nm = med(curtos.map(q => q.utm.N)), f0 = curtos[0].fuso || padrao;
        let soma = numero(pj.somarE), origemE = 'campo "Somar ao E" da aba Projeto', ref = null;
        if (soma == null) {
          const completos = utmPendentes.filter(q => !q.curto).map(q => WGDxf.utmParaLatLon(q.utm.E, q.utm.N, (q.fuso || padrao).zona, (q.fuso || padrao).sul));
          if (graus.length) { ref = { lat: media(graus.map(p => p.latitude)), lon: media(graus.map(p => p.longitude)) }; origemE = 'poços da planilha em graus'; }
          else if (completos.length) { ref = { lat: media(completos.map(r => r.lat)), lon: media(completos.map(r => r.lon)) }; origemE = 'poços da planilha com o E completo'; }
          else if (ctx.centroProjeto) { ref = ctx.centroProjeto; origemE = 'poços já cadastrados no projeto'; }
          else if (ctx.referencia) { ref = ctx.referencia; origemE = 'poços dos projetos da empresa'; }
          else if (pj.uf && CENTRO_UF[String(pj.uf).trim().toUpperCase()]) { const c = CENTRO_UF[String(pj.uf).trim().toUpperCase()]; ref = { lat: c[0], lon: c[1] }; origemE = `centro da UF ${String(pj.uf).trim().toUpperCase()}`; }
          if (ref) {
            let melhor = null;
            for (let d = 1; d <= 8; d++) {
              const E = Em + d * 100000; if (E < 160000 || E > 840000) continue;
              const r = WGDxf.utmParaLatLon(E, Nm, f0.zona, f0.sul);
              const dist = Math.hypot((r.lon - ref.lon) * Math.cos(ref.lat * Math.PI / 180), r.lat - ref.lat) * 111;
              if (!melhor || dist < melhor.dist) melhor = { soma: d * 100000, dist };
            }
            if (melhor) soma = melhor.soma;
          }
        }
        if (soma == null) {
          curtos.forEach(q => erro(q.aba, q.n, `${q.cod}: E com 5 dígitos ('${q.utm.E}'): falta a centena de milhar. Preencha "Somar ao E" na aba Projeto (ex.: 300000) ou escreva o E completo (ex.: 337859)`));
          curtos.forEach(q => { q.invalido = true; });
        } else {
          curtos.forEach(q => { q.utm.E += soma; });
          aviso(tPocos.nome, 0, `${curtos.length} poço(s) com o E incompleto (5 dígitos, ex.: ${WGMalha_texto(Em)}): somado ${WGMalha_texto(soma)} ao E (ficou ${WGMalha_texto(Em + soma)}), escolhido pelos ${origemE}. Confira no mapa; se não for isso, preencha "Somar ao E" na aba Projeto.`);
          contar('Coordenadas com o E completado', curtos.length);
        }
      }
      const semFuso = utmPendentes.filter(q => !q.fuso).length;
      if (semFuso) aviso(tPocos.nome, 0, `${semFuso} poço(s) em UTM sem coluna Fuso: usado o fuso ${padrao.zona}${padrao.sul ? 'S' : 'N'} (${origem}), SIRGAS 2000/WGS 84. Se estiver errado, preencha a coluna Fuso (ex.: 23S) ou o campo Fuso na aba Projeto.`);
      utmPendentes.forEach(q => {
        if (q.invalido) return;
        const f = q.fuso || padrao, r = WGDxf.utmParaLatLon(q.utm.E, q.utm.N, f.zona, f.sul);
        if (!Number.isFinite(r.lat) || !Number.isFinite(r.lon) || Math.abs(r.lat) > 90) { erro(q.aba, q.n, `${q.cod}: UTM E ${q.utm.E} / N ${q.utm.N} não deu uma posição válida no fuso ${f.zona}`); return; }
        if (q.poco) { q.poco.latitude = Math.round(r.lat * 1e8) / 1e8; q.poco.longitude = Math.round(r.lon * 1e8) / 1e8; }
        rel.convertidos.push({ cod: q.cod, E0: q.E0, N0: q.N0, E: q.utm.E, N: q.utm.N, zona: f.zona, sul: f.sul, lat: Math.round(r.lat * 1e8) / 1e8, lon: Math.round(r.lon * 1e8) / 1e8 });
      });
      contar('Coordenadas UTM convertidas para graus', utmPendentes.filter(q => !q.invalido).length);
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

  /**
   * Botão "Converter UTM em latitude/longitude": troca, na aba Pocos, as coordenadas UTM pelas latitudes e longitudes
   * (mesmas regras da importação: fuso, E de 5 dígitos...) e guarda o E/N original em colunas novas no fim.
   * Mexe no próprio wb. Devolve { n, rel, erro? } (rel = relatório da leitura, só para os avisos de coordenada).
   */
  function converterPlanilha(XLSX, wb, ctx = {}) {
    const { rel } = ler(XLSX, wb, 'conversao', ctx);
    const ws = acharAba(wb, 'Pocos'), t = tabela(XLSX, wb, 'Pocos');
    if (!ws || !t) return { n: 0, rel, erro: 'a planilha não tem a aba Pocos' };
    const conv = new Map(rel.convertidos.map(c => [c.cod, c]));
    if (!conv.size) return { n: 0, rel };
    const col = nomes => { for (const nm of nomes) { const c = t.cabecalhos.find(h => h.chave === chave(nm)); if (c) return c.col; } return null; };
    let fim = Math.max(...t.cabecalhos.map(h => h.col));
    const nova = titulo => { fim += 1; ws[XLSX.utils.encode_cell({ r: 0, c: fim })] = { t: 's', v: titulo }; return fim; };
    const cPoco = col(COL_POCO);
    const cLat = col(COL_LAT) ?? nova('Latitude'), cLon = col(COL_LON) ?? nova('Longitude');
    const cE = nova('E UTM original (m)'), cN = nova('N UTM original (m)'), cF = nova('Fuso UTM usado');
    const cel = (r, c, v) => { ws[XLSX.utils.encode_cell({ r, c })] = typeof v === 'number' ? { t: 'n', v } : { t: 's', v: String(v) }; };
    let n = 0;
    t.linhas.forEach(l => {
      const cod = codigoPoco(l.get(...COL_POCO)), c = cod && conv.get(cod);
      if (!c || cPoco == null) return;
      const r = l.n - 1;
      cel(r, cLat, c.lat); cel(r, cLon, c.lon);
      cel(r, cE, c.E0); cel(r, cN, c.N0);
      cel(r, cF, `${c.zona}${c.sul ? 'S' : 'N'}${c.E !== c.E0 ? ` (E + ${Math.round(c.E - c.E0)})` : ''}`);
      n++;
    });
    const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    ref.e.c = Math.max(ref.e.c, fim); ws['!ref'] = XLSX.utils.encode_range(ref);
    return { n, rel };
  }

  return { ler, lerProjeto, valorLab, numero, data, chave, coordPoco, lerFuso, converterPlanilha };
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

  /** Triangulação de Delaunay (Bowyer-Watson) de pontos {x, y}: devolve [[i, j, k]...]. */
  function delaunay(pts) {
    const n = pts.length; if (n < 3) return [];
    const xs = pts.map(p => p.x), ys = pts.map(p => p.y);
    const x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys), d = Math.max(x1 - x0, y1 - y0) * 20 + 1;
    const cx = (x0 + x1) / 2, cy = (y0 + y1) / 2;
    const P = pts.map(p => [p.x, p.y]).concat([[cx - d, cy - d], [cx + d, cy - d], [cx, cy + d]]);
    const circ = (a, b, c) => {
      const [ax, ay] = P[a], [bx, by] = P[b], [qx, qy] = P[c];
      const D = 2 * (ax * (by - qy) + bx * (qy - ay) + qx * (ay - by)) || 1e-12;
      const ux = ((ax * ax + ay * ay) * (by - qy) + (bx * bx + by * by) * (qy - ay) + (qx * qx + qy * qy) * (ay - by)) / D;
      const uy = ((ax * ax + ay * ay) * (qx - bx) + (bx * bx + by * by) * (ax - qx) + (qx * qx + qy * qy) * (bx - ax)) / D;
      return { a, b, c, ux, uy, r2: (ax - ux) ** 2 + (ay - uy) ** 2 };
    };
    let tri = [circ(n, n + 1, n + 2)];
    for (let i = 0; i < n; i++) {
      const [px, py] = P[i], ruins = [], bons = [];
      tri.forEach(t => ((px - t.ux) ** 2 + (py - t.uy) ** 2 < t.r2 ? ruins : bons).push(t));
      const arestas = new Map();
      ruins.forEach(t => [[t.a, t.b], [t.b, t.c], [t.c, t.a]].forEach(([u, v]) => { const k = u < v ? u + ',' + v : v + ',' + u; arestas.set(k, arestas.has(k) ? null : [u, v]); }));
      tri = bons;
      arestas.forEach(e => { if (e) tri.push(circ(e[0], e[1], i)); });
    }
    return tri.filter(t => t.a < n && t.b < n && t.c < n).map(t => [t.a, t.b, t.c]);
  }

  /**
   * Divide a linha do limite em trechos DEFINIDOS (contínuos) e INFERIDOS (tracejados).
   * Definido = o trecho passa dentro de um triângulo (Delaunay dos poços) que tem pelo menos um poço acima e um
   * abaixo do limite: ali a linha está presa entre um ponto com valor alto e um com valor baixo. Fora disso
   * (borda da rede, ou região sem poço baixo por perto) a linha é uma extrapolação: inferida.
   * curvas: [[[lat, lon]...]] (saída de isolinha); amostras: [{lat, lon, v}]. Devolve [{pts, definido}].
   */
  function trechosLimite(g, curvas, amostras, limiar) {
    if (!amostras || amostras.length < 3) return curvas.map(c => ({ pts: c, definido: false }));
    const pts = amostras.map(a => ({ ...g.proj.paraXY(a.lat, a.lon), alto: a.v >= limiar }));
    const esp = espacamentos(pts), med = [...esp].sort((a, b) => a - b)[Math.floor(esp.length / 2)];
    const maxAresta = Math.max(4 * med, 30);
    const L_ = (i, j) => Math.hypot(pts[i].x - pts[j].x, pts[i].y - pts[j].y);
    const tris = delaunay(pts).filter(([a, b, c]) => Math.max(L_(a, b), L_(b, c), L_(c, a)) <= maxAresta)
      .filter(([a, b, c]) => { const n = [a, b, c].filter(k => pts[k].alto).length; return n > 0 && n < 3; })   // misto: alto e baixo
      .map(([a, b, c]) => {
        const A = pts[a], B = pts[b], C = pts[c], det = (B.y - C.y) * (A.x - C.x) + (C.x - B.x) * (A.y - C.y);
        return { A, B, C, det, x0: Math.min(A.x, B.x, C.x) - .5, x1: Math.max(A.x, B.x, C.x) + .5, y0: Math.min(A.y, B.y, C.y) - .5, y1: Math.max(A.y, B.y, C.y) + .5 };
      }).filter(t => Math.abs(t.det) > 1e-9);
    const dentro = (x, y) => tris.some(t => {
      if (x < t.x0 || x > t.x1 || y < t.y0 || y > t.y1) return false;
      const w1 = ((t.B.y - t.C.y) * (x - t.C.x) + (t.C.x - t.B.x) * (y - t.C.y)) / t.det;
      const w2 = ((t.C.y - t.A.y) * (x - t.C.x) + (t.A.x - t.C.x) * (y - t.C.y)) / t.det;
      return w1 >= -.02 && w2 >= -.02 && 1 - w1 - w2 >= -.02;
    });
    const out = [];
    curvas.forEach(c => {
      if (c.length < 2) return;
      // classifica cada pedaço pelo ponto do meio; junta pedaços seguidos da mesma classe
      let atual = null;
      for (let k = 0; k + 1 < c.length; k++) {
        const ma = g.proj.paraXY((c[k][0] + c[k + 1][0]) / 2, (c[k][1] + c[k + 1][1]) / 2);
        const def = dentro(ma.x, ma.y);
        if (!atual || atual.definido !== def) { atual = { pts: [c[k]], definido: def }; out.push(atual); }
        atual.pts.push(c[k + 1]);
      }
    });
    // pedacinhos soltos (menos de 3 vértices) de uma classe no meio da outra: incorpora no vizinho, para não piscar
    return out;
  }

  /** Para cada pedaço (k -> k+1) da linha [[lat, lon]...]: true = definido (entre poço acima e abaixo do limite). */
  function classificarSegmentos(g, pts, amostras, limiar) {
    const t = trechosLimite(g, [pts], amostras, limiar), out = [];
    t.forEach(r => { for (let k = 0; k + 1 < r.pts.length; k++) out.push(r.definido); });
    return out;
  }

  /** Simplifica uma linha [[lat, lon]...] (Douglas-Peucker, tolerância em metros na projeção da grade). */
  function simplificar(g, pts, tol = 1) {
    if (pts.length < 4) return pts.slice();
    const xy = pts.map(([la, lo]) => g.proj.paraXY(la, lo)), manter = new Uint8Array(pts.length);
    manter[0] = manter[pts.length - 1] = 1;
    const pilha = [[0, pts.length - 1]];
    while (pilha.length) {
      const [a, b] = pilha.pop(); let m = -1, dm = 0;
      const A = xy[a], B = xy[b], dx = B.x - A.x, dy = B.y - A.y, L2 = dx * dx + dy * dy;
      for (let i = a + 1; i < b; i++) {
        const t = L2 ? Math.max(0, Math.min(1, ((xy[i].x - A.x) * dx + (xy[i].y - A.y) * dy) / L2)) : 0;
        const d = Math.hypot(xy[i].x - A.x - t * dx, xy[i].y - A.y - t * dy);
        if (d > dm) { dm = d; m = i; }
      }
      if (dm > tol && m > 0) { manter[m] = 1; pilha.push([a, m], [m, b]); }
    }
    return pts.filter((_, i) => manter[i]);
  }

  /** Área (m²) e centróide de polígonos [[lat, lon]...] (fechados ou não), na projeção da grade. */
  function areaPoligonos(g, polis) {
    let A = 0, cx = 0, cy = 0;
    polis.forEach(pl => {
      const q = pl.map(([la, lo]) => g.proj.paraXY(la, lo));
      for (let i = 0, j = q.length - 1; i < q.length; j = i++) {
        const c = q[j].x * q[i].y - q[i].x * q[j].y; A += c; cx += (q[j].x + q[i].x) * c; cy += (q[j].y + q[i].y) * c;
      }
    });
    if (Math.abs(A) < 1e-9) return { area: 0, centroide: null };
    const c = { x: cx / (3 * A), y: cy / (3 * A) };
    return { area: Math.abs(A / 2), centroide: { ...c, ...g.proj.paraLatLon(c.x, c.y) } };
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

  // mapa de calor: azul (baixo) -> verde -> amarelo -> vermelho (alto), em escala log
  const RAMPA_CALOR = ['#2c7bb6', '#00a6ca', '#00ccbc', '#90eb9d', '#ffff8c', '#f9d057', '#f29e2e', '#e76818', '#d7191c'];
  function corCalor(valor, vmin, vmax) {
    const t = Math.max(0, Math.min(1, (Math.log10(valor) - Math.log10(vmin)) / (Math.log10(vmax) - Math.log10(vmin) || 1)));
    const f = t * (RAMPA_CALOR.length - 1), k = Math.min(RAMPA_CALOR.length - 2, Math.floor(f)), r = f - k;
    const a = hexRgb(RAMPA_CALOR[k]), b = hexRgb(RAMPA_CALOR[k + 1]);
    return a.map((x, n) => Math.round(x + (b[n] - x) * r));
  }

  return { corCalor, RAMPA_CALOR, grade, idw, trechosLimite, classificarSegmentos, simplificar, areaPoligonos, delaunay, agrupar, areaECentroide, distancia, isolinha, cor, RAMPA, projecao,
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
    const comp = Math.min(passo * .35, 9), setas = [];   // seta curta: um terço do espaçamento, no máximo 9 m
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
      const a = comp * .3, cs = Math.cos(2.6), sn = Math.sin(2.6); // asas a ~150° do sentido
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

/* =====================================================================
   3d. REVISÃO DA IMPORTAÇÃO (funções puras, sem tela)
   Antes de gravar, compara a planilha com o histórico do projeto e aponta
   o que merece conferência: valores muito fora do histórico, primeira vez
   acima do VI, possível erro de unidade, N.A. e cota que mudaram muito.
   ===================================================================== */
const WGRevisao = (() => {
  const K = s => String(s ?? '').trim().toUpperCase().replace(/\s+/g, '');
  const num = v => (v === null || v === undefined || v === '' || !Number.isFinite(+v)) ? null : +v;
  const f = (v, c = 4) => v === null || v === undefined ? '-' : Number(v).toLocaleString('pt-BR', { maximumFractionDigits: c });
  const vezes = x => x >= 100 ? f(Math.round(x), 0) : f(x, 1);

  /**
   * payload: o que a planilha vai gravar (resultados, campo, pocos, campanha).
   * hist: o que o projeto já tem, por código:
   *   { resultados: [{poco, parametro, campanha, data, valor, menor_que_lq, lq}],
   *     medicoes:   [{poco, campanha, data, nivel_agua}],
   *     pocos:      [{codigo, cota_topo, latitude, longitude, profundidade}],
   *     campanhas:  [{codigo, data_inicio}] }
   * parametros: [{nome, valor_orientador, unidade}]
   * Devolve { alertas: [{nivel: 'alto'|'medio', poco, item, antes, agora, motivo}], substitui: {campanha, resultados} | null }
   */
  function conferir(payload, hist, parametros) {
    hist = hist || { resultados: [], medicoes: [], pocos: [], campanhas: [] };
    const alertas = [];
    const camp = payload.campanha ? K(payload.campanha.codigo) : null;
    const VI = new Map((parametros || []).map(p => [K(p.nome), { vi: num(p.valor_orientador), un: p.unidade || 'µg/L' }]));
    (payload.parametros || []).forEach(p => { if (!VI.has(K(p.nome))) VI.set(K(p.nome), { vi: num(p.valor_orientador), un: p.unidade || 'µg/L' }); });
    const dataCamp = new Map((hist.campanhas || []).map(c => [K(c.codigo), String(c.data_inicio || '')]));
    // compara só com campanhas ANTERIORES à da planilha (reimportar uma campanha antiga não compara com as mais novas)
    const dataNova = payload.campanha?.data_inicio ? String(payload.campanha.data_inicio).slice(0, 10) : '';
    const dataDe = r => String(dataCamp.get(K(r.campanha)) || r.data || '').slice(0, 10);
    const outras = r => K(r.campanha) !== camp && (!dataNova || !dataDe(r) || dataDe(r) < dataNova);

    // resultados do laboratório
    const porChave = new Map();
    (hist.resultados || []).filter(outras).forEach(r => {
      const k = K(r.poco) + '|' + K(r.parametro);
      if (!porChave.has(k)) porChave.set(k, []);
      porChave.get(k).push(r);
    });
    (payload.resultados || []).forEach(r => {
      const p = VI.get(K(r.parametro)) || { vi: null, un: '' }, vi = p.vi, un = p.un;
      const ant = (porChave.get(K(r.poco) + '|' + K(r.parametro)) || [])
        .sort((a, b) => String(a.data || dataCamp.get(K(a.campanha)) || '').localeCompare(String(b.data || dataCamp.get(K(b.campanha)) || '')));
      const quant = ant.filter(a => !a.menor_que_lq && num(a.valor) !== null);
      const maior = quant.reduce((m, a) => (!m || +a.valor > +m.valor) ? a : m, null);
      const item = `${r.parametro}${un ? ` (${un})` : ''}`;
      const agora = r.menor_que_lq ? (num(r.lq) !== null ? `< ${f(r.lq)}` : '< LQ') : f(r.valor);
      const v = r.menor_que_lq ? null : num(r.valor);
      if (v !== null && vi && v >= 100 * vi) {
        alertas.push({ nivel: 'alto', poco: r.poco, item, antes: maior ? `${f(maior.valor)} (${maior.campanha})` : '-', agora,
          motivo: `${vezes(v / vi)} vezes o valor orientador (${f(vi)}). Confira o laudo e a unidade (µg/L × mg/L).` });
        return;
      }
      if (v !== null && maior && +maior.valor > 0 && v >= 10 * +maior.valor && (!vi || v >= vi)) {
        alertas.push({ nivel: 'alto', poco: r.poco, item, antes: `${f(maior.valor)} (${maior.campanha})`, agora,
          motivo: `${vezes(v / +maior.valor)} vezes o maior valor que este poço já teve.` });
        return;
      }
      if (v !== null && vi && v >= vi && ant.length && !quant.some(a => +a.valor >= vi)) {
        alertas.push({ nivel: 'medio', poco: r.poco, item, antes: maior ? `${f(maior.valor)} (${maior.campanha})` : 'abaixo do LQ', agora,
          motivo: `Primeira vez acima do valor orientador (${f(vi)}) em ${ant.length} campanha(s).` });
        return;
      }
      const ultimo = ant[ant.length - 1];
      if (r.menor_que_lq && vi && ultimo && !ultimo.menor_que_lq && num(ultimo.valor) >= 10 * vi) {
        alertas.push({ nivel: 'medio', poco: r.poco, item, antes: `${f(ultimo.valor)} (${ultimo.campanha})`, agora,
          motivo: 'Caiu de muito acima do valor orientador para abaixo do LQ de uma campanha para a outra.' });
      }
    });

    // nível d'água
    const pocoHist = new Map((hist.pocos || []).map(p => [K(p.codigo), p]));
    const pocoNovo = new Map((payload.pocos || []).map(p => [K(p.codigo), p]));
    const medAnt = new Map();
    (hist.medicoes || []).filter(outras).forEach(m => {
      const k = K(m.poco), d = String(m.data || dataCamp.get(K(m.campanha)) || '');
      if (num(m.nivel_agua) === null) return;
      if (!medAnt.has(k) || d > medAnt.get(k).d) medAnt.set(k, { ...m, d });
    });
    (payload.campo || []).forEach(c => {
      const na = num(c.nivel_agua); if (na === null) return;
      const prof = num(pocoNovo.get(K(c.poco))?.profundidade) ?? num(pocoHist.get(K(c.poco))?.profundidade);
      if (na < 0 || (prof && na > prof)) {
        alertas.push({ nivel: 'alto', poco: c.poco, item: "N.A. (m)", antes: prof ? `profundidade ${f(prof, 2)}` : '-', agora: f(na, 2),
          motivo: na < 0 ? 'N.A. negativo.' : 'N.A. mais fundo que o próprio poço.' });
        return;
      }
      const a = medAnt.get(K(c.poco));
      if (a && Math.abs(na - +a.nivel_agua) > 2) {
        alertas.push({ nivel: 'medio', poco: c.poco, item: "N.A. (m)", antes: `${f(a.nivel_agua, 2)} (${a.campanha})`, agora: f(na, 2),
          motivo: `Mudou ${f(Math.abs(na - +a.nivel_agua), 2)} m desde a última medição.` });
      }
    });

    // cota e coordenada do poço
    (payload.pocos || []).forEach(p => {
      const h = pocoHist.get(K(p.codigo)); if (!h) return;
      const c1 = num(h.cota_topo), c2 = num(p.cota_topo);
      if (c1 !== null && c2 !== null && Math.abs(c1 - c2) > .5) {
        alertas.push({ nivel: 'medio', poco: p.codigo, item: 'Cota do topo (m)', antes: f(c1, 3), agora: f(c2, 3),
          motivo: `Mudou ${f(Math.abs(c1 - c2), 2)} m: muda a carga hidráulica e o potenciométrico.` });
      }
      const la1 = num(h.latitude), lo1 = num(h.longitude), la2 = num(p.latitude), lo2 = num(p.longitude);
      if ([la1, lo1, la2, lo2].every(x => x !== null)) {
        const d = Math.hypot((la2 - la1) * 110574, (lo2 - lo1) * 111320 * Math.cos(la1 * Math.PI / 180));
        if (d > 10) alertas.push({ nivel: 'medio', poco: p.codigo, item: 'Coordenada', antes: `${f(la1, 6)}, ${f(lo1, 6)}`, agora: `${f(la2, 6)}, ${f(lo2, 6)}`,
          motivo: `O poço mudou ${f(Math.round(d), 0)} m de lugar.` });
      }
    });

    const ordem = { alto: 0, medio: 1 };
    alertas.sort((a, b) => (ordem[a.nivel] - ordem[b.nivel]) || String(a.poco).localeCompare(String(b.poco), 'pt-BR', { numeric: true }));
    const jaTinha = camp ? (hist.resultados || []).filter(r => K(r.campanha) === camp).length : 0;
    const existia = camp && (hist.campanhas || []).some(c => K(c.codigo) === camp);
    return { alertas, substitui: existia ? { campanha: payload.campanha.codigo, resultados: jaTinha } : null };
  }

  return { conferir };
})();

/* =====================================================================
   3e. FOTOS 360° (funções puras, sem tela)
   Lê o GPS gravado pela câmera (EXIF do JPEG) e confere se a imagem é
   uma panorâmica equiretangular (largura = 2 x altura).
   ===================================================================== */
const WG360 = (() => {
  /** Latitude/longitude do EXIF de um JPEG (ArrayBuffer). null se não tiver GPS. */
  function gpsDoJpeg(buf) {
    try {
      const v = new DataView(buf);
      if (v.byteLength < 12 || v.getUint16(0) !== 0xFFD8) return null;
      let o = 2;
      while (o + 4 <= v.byteLength) {
        const m = v.getUint16(o);
        if ((m & 0xFF00) !== 0xFF00) return null;
        if (m === 0xFFDA || m === 0xFFD9) return null;              // começou a imagem: não há EXIF
        const len = v.getUint16(o + 2);
        if (m === 0xFFE1 && o + 10 <= v.byteLength && v.getUint32(o + 4) === 0x45786966) return lerTiff(v, o + 10);
        o += 2 + len;
      }
    } catch (e) { /* arquivo estranho: sem GPS */ }
    return null;
  }
  function lerTiff(v, t) {
    const le = v.getUint16(t) === 0x4949;
    const u16 = a => v.getUint16(a, le), u32 = a => v.getUint32(a, le);
    const tag = (ifd, id) => { const n = u16(ifd); for (let i = 0; i < n; i++) { const e = ifd + 2 + i * 12; if (u16(e) === id) return e; } return null; };
    const g = tag(t + u32(t + 4), 0x8825);
    if (!g) return null;
    const gps = t + u32(g + 8);
    const graus = e => { const a = t + u32(e + 8); const r = k => u32(a + k * 8) / u32(a + k * 8 + 4); return r(0) + r(1) / 60 + r(2) / 3600; };
    const ref = e => e ? String.fromCharCode(v.getUint8(e + 8)) : '';
    const eLat = tag(gps, 2), eLon = tag(gps, 4);
    if (!eLat || !eLon) return null;
    let lat = graus(eLat), lon = graus(eLon);
    if (ref(tag(gps, 1)) === 'S') lat = -lat;
    if (ref(tag(gps, 3)) === 'W') lon = -lon;
    if (!Number.isFinite(lat) || !Number.isFinite(lon) || Math.abs(lat) > 90 || Math.abs(lon) > 180 || (lat === 0 && lon === 0)) return null;
    return { lat, lon };
  }

  /** Panorâmica 360° completa tem proporção 2:1 (aceita 5% de folga). */
  function eh360(largura, altura) { return largura > 0 && altura > 0 && Math.abs(largura / altura - 2) <= 0.1; }

  /** Tamanho final da foto gravada: no máximo 4096 x 2048 (leve e aceito por qualquer celular). */
  function tamanhoFinal(largura, altura, max = 4096) {
    const L = Math.min(max, largura);
    return { largura: L, altura: Math.round(L / 2) };
  }

  return { gpsDoJpeg, eh360, tamanhoFinal };
})();

/* =====================================================================
   Visor 360°: WebGL puro, sem biblioteca. Um retângulo na tela e um
   shader que, para cada pixel, calcula a direção do olhar e busca a cor
   na imagem equiretangular. Arrastar = girar; roda/pinça = aproximar.
   ===================================================================== */
const Visor360 = (() => {
  const VS = 'attribute vec2 p; void main(){ gl_Position = vec4(p, 0.0, 1.0); }';
  const FS = `
    #ifdef GL_FRAGMENT_PRECISION_HIGH
    precision highp float;
    #else
    precision mediump float;
    #endif
    uniform sampler2D t; uniform vec2 r; uniform float yaw; uniform float pitch; uniform float f;
    void main(){
      vec2 q = (gl_FragCoord.xy * 2.0 - r) / r.y;
      vec3 d = normalize(vec3(q.x * f, q.y * f, 1.0));
      float cp = cos(pitch), sp = sin(pitch);
      d = vec3(d.x, d.y * cp + d.z * sp, -d.y * sp + d.z * cp);
      float cy = cos(yaw), sy = sin(yaw);
      d = vec3(d.x * cy + d.z * sy, d.y, -d.x * sy + d.z * cy);
      float lon = atan(d.x, d.z);
      float lat = asin(clamp(d.y, -1.0, 1.0));
      gl_FragColor = texture2D(t, vec2(0.5 + lon / 6.28318530718, 0.5 - lat / 3.14159265359));
    }`;
  const FOV_MIN = 30, FOV_MAX = 110, FOV_INI = 80;

  /** Abre a imagem (já carregada) dentro de "area". Devolve {zoom, centro, destruir}. */
  function abrir(area, img) {
    const cv = document.createElement('canvas');
    cv.className = 'visor360-canvas';
    area.appendChild(cv);
    const gl = cv.getContext('webgl', { antialias: false, preserveDrawingBuffer: false }) || cv.getContext('experimental-webgl');
    if (!gl) { cv.remove(); return planoSemWebgl(area, img); }

    const sh = (tipo, src) => { const s = gl.createShader(tipo); gl.shaderSource(s, src); gl.compileShader(s); return s; };
    const prog = gl.createProgram();
    gl.attachShader(prog, sh(gl.VERTEX_SHADER, VS)); gl.attachShader(prog, sh(gl.FRAGMENT_SHADER, FS));
    gl.linkProgram(prog);
    if (!gl.getProgramParameter(prog, gl.LINK_STATUS)) { cv.remove(); return planoSemWebgl(area, img); }
    gl.useProgram(prog);
    const buf = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, buf);
    gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([-1, -1, 1, -1, -1, 1, 1, 1]), gl.STATIC_DRAW);
    const loc = gl.getAttribLocation(prog, 'p');
    gl.enableVertexAttribArray(loc); gl.vertexAttribPointer(loc, 2, gl.FLOAT, false, 0, 0);

    // textura: se a placa de vídeo não aceitar o tamanho, reduz antes
    const max = gl.getParameter(gl.MAX_TEXTURE_SIZE) || 4096;
    let fonte = img;
    const w0 = img.naturalWidth || img.width, h0 = img.naturalHeight || img.height;
    if (w0 > max || h0 > max) {
      const k = Math.min(max / w0, max / h0), c = document.createElement('canvas');
      c.width = Math.floor(w0 * k); c.height = Math.floor(h0 * k);
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
      fonte = c;
    }
    const tex = gl.createTexture();
    gl.bindTexture(gl.TEXTURE_2D, tex);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGB, gl.RGB, gl.UNSIGNED_BYTE, fonte);
    const u = n => gl.getUniformLocation(prog, n);
    const U = { r: u('r'), yaw: u('yaw'), pitch: u('pitch'), f: u('f') };

    const vista = { yaw: 0, pitch: 0, fov: FOV_INI };
    let pedido = 0, vivo = true;
    function desenhar() {
      pedido = 0; if (!vivo) return;
      const dpr = Math.min(2, window.devicePixelRatio || 1);
      const W = Math.max(1, Math.round(cv.clientWidth * dpr)), H = Math.max(1, Math.round(cv.clientHeight * dpr));
      if (cv.width !== W || cv.height !== H) { cv.width = W; cv.height = H; }
      gl.viewport(0, 0, W, H);
      gl.uniform2f(U.r, W, H);
      gl.uniform1f(U.yaw, vista.yaw * Math.PI / 180);
      gl.uniform1f(U.pitch, vista.pitch * Math.PI / 180);
      gl.uniform1f(U.f, Math.tan(vista.fov * Math.PI / 360));
      gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }
    const pedir = () => { if (!pedido) pedido = requestAnimationFrame(desenhar); };
    const limitar = () => {
      vista.fov = Math.max(FOV_MIN, Math.min(FOV_MAX, vista.fov));
      vista.pitch = Math.max(-89, Math.min(89, vista.pitch));
      vista.yaw = ((vista.yaw + 540) % 360) - 180;
    };

    // arrastar (mouse, caneta ou dedo) e pinça com dois dedos
    const toques = new Map();
    let pinca = null;
    cv.addEventListener('pointerdown', e => { cv.setPointerCapture(e.pointerId); toques.set(e.pointerId, { x: e.clientX, y: e.clientY }); cv.classList.add('arrastando'); });
    cv.addEventListener('pointermove', e => {
      const a = toques.get(e.pointerId); if (!a) return;
      if (toques.size === 2) {
        a.x = e.clientX; a.y = e.clientY;
        const [p1, p2] = [...toques.values()], d = Math.hypot(p1.x - p2.x, p1.y - p2.y);
        if (pinca) { vista.fov = pinca.fov * pinca.d / Math.max(10, d); limitar(); pedir(); } else pinca = { d, fov: vista.fov };
        return;
      }
      const grausPorPx = vista.fov / Math.max(1, cv.clientHeight);
      vista.yaw -= (e.clientX - a.x) * grausPorPx;
      vista.pitch += (e.clientY - a.y) * grausPorPx;
      a.x = e.clientX; a.y = e.clientY;
      limitar(); pedir();
    });
    const soltar = e => { toques.delete(e.pointerId); pinca = null; if (!toques.size) cv.classList.remove('arrastando'); };
    cv.addEventListener('pointerup', soltar); cv.addEventListener('pointercancel', soltar);
    cv.addEventListener('wheel', e => { e.preventDefault(); vista.fov *= Math.exp(e.deltaY * 0.0012); limitar(); pedir(); }, { passive: false });
    const ro = window.ResizeObserver ? new ResizeObserver(pedir) : null;
    if (ro) ro.observe(cv); else window.addEventListener('resize', pedir);
    pedir();

    return {
      zoom(k) { vista.fov *= k; limitar(); pedir(); },
      girar(dy, dp) { vista.yaw += dy; vista.pitch += dp; limitar(); pedir(); },
      centro() { vista.yaw = 0; vista.pitch = 0; vista.fov = FOV_INI; pedir(); },
      destruir() {
        vivo = false; if (pedido) cancelAnimationFrame(pedido);
        if (ro) ro.disconnect(); else window.removeEventListener('resize', pedir);
        gl.deleteTexture(tex); gl.deleteBuffer(buf); gl.deleteProgram(prog);
        const ext = gl.getExtension('WEBGL_lose_context'); if (ext) ext.loseContext();
        cv.remove();
      }
    };
  }

  /** Sem WebGL (raro): mostra a panorâmica plana, com rolagem para o lado. */
  function planoSemWebgl(area, img) {
    const d = document.createElement('div');
    d.className = 'visor360-plano';
    const i = img.cloneNode(); d.appendChild(i); area.appendChild(d);
    return { zoom() {}, girar(dy) { d.scrollLeft += dy * 8; }, centro() { d.scrollLeft = (d.scrollWidth - d.clientWidth) / 2; }, destruir() { d.remove(); } };
  }

  return { abrir };
})();

/* =====================================================================
   3f. DIAGRAMA DE CERCAS 3D (funções puras, sem tela)
   Monta os triângulos das colunas de sondagem e dos painéis entre elas.
   Coordenadas em metros: x = leste, y = norte, z = cota (já com exagero).
   Os painéis usam o mesmo preenchimento da seção 2D (WGSecao.preencher).
   ===================================================================== */
const WGCerca = (() => {
  const rgb = h => {
    const m = /^#?([0-9a-f]{6})$/i.exec(String(h || '').trim());
    if (!m) return [0.6, 0.6, 0.6];
    const n = parseInt(m[1], 16);
    return [(n >> 16 & 255) / 255, (n >> 8 & 255) / 255, (n & 255) / 255];
  };

  /** Árvore geradora mínima (Prim): liga todos os pontos sem cruzar, pelo caminho mais curto. Devolve [[i, j]]. */
  function arvoreMinima(pts) {
    const n = pts.length; if (n < 2) return [];
    const dentro = new Array(n).fill(false), dist = new Array(n).fill(Infinity), pai = new Array(n).fill(-1), lig = [];
    dist[0] = 0;
    for (let k = 0; k < n; k++) {
      let u = -1;
      for (let i = 0; i < n; i++) if (!dentro[i] && (u < 0 || dist[i] < dist[u])) u = i;
      dentro[u] = true;
      if (pai[u] >= 0) lig.push([pai[u], u]);
      for (let v = 0; v < n; v++) {
        if (dentro[v]) continue;
        const d = Math.hypot(pts[u].x - pts[v].x, pts[u].y - pts[v].y);
        if (d < dist[v]) { dist[v] = d; pai[v] = u; }
      }
    }
    return lig;
  }

  /** Poços no mesmo lugar (PM-26, PMN-26A...) ficam em roda, para as colunas não se sobreporem. */
  function espalharColunas(cols, w) {
    const usados = new Array(cols.length).fill(false);
    cols.forEach((c, i) => {
      if (usados[i]) return;
      const g = [i];
      for (let j = i + 1; j < cols.length; j++) if (!usados[j] && Math.hypot(cols[j].x - c.x, cols[j].y - c.y) < w * 1.6) g.push(j);
      g.forEach(j => { usados[j] = true; });
      if (g.length < 2) return;
      const cx = g.reduce((s, j) => s + cols[j].x, 0) / g.length, cy = g.reduce((s, j) => s + cols[j].y, 0) / g.length;
      const r = w * 0.9 / Math.sin(Math.PI / g.length);
      g.forEach((j, k) => { const a = 2 * Math.PI * k / g.length; cols[j].x = cx + r * Math.cos(a); cols[j].y = cy + r * Math.sin(a); });
    });
    return cols;
  }

  /**
   * cols: [{nome, x, y, cota, prof, camadas:[{de, ate, hex, lit}], na}]
   * ligacoes: [[i, j]] (painéis); ex = exagero vertical; w = largura da coluna (m).
   * Devolve { tri: {pos, cor, nor}, lin: {pos, cor}, caixa, rotulos, colunas }.
   */
  function geometria(cols, ligacoes, { ex = 1, w = 1, paineis = true, agua = true } = {}) {
    const P = [], C = [], N = [], LP = [], LC = [];
    const tri = (a, b, c, cor, n) => { P.push(...a, ...b, ...c); for (let k = 0; k < 3; k++) { C.push(...(Array.isArray(cor[0]) ? cor[k] : cor)); N.push(...n); } };
    const quad = (a, b, c, d, cor, n) => { const cs = Array.isArray(cor[0]) ? cor : [cor, cor, cor, cor]; tri(a, b, c, [cs[0], cs[1], cs[2]], n); tri(a, c, d, [cs[0], cs[2], cs[3]], n); };
    const linha = (a, b, cor) => { LP.push(...a, ...b); LC.push(...cor, ...cor); };
    const Z = z => z * ex;
    const AZUL = [0.08, 0.40, 0.66], TERRA = [0.48, 0.35, 0.18], CINZA = [0.72, 0.74, 0.73];

    // painéis entre sondagens ligadas
    let paineisFeitos = 0;
    if (paineis) ligacoes.forEach(([i, j]) => {
      const a = cols[i], b = cols[j];
      if (!a.camadas.length || !b.camadas.length) return;
      const L = Math.hypot(b.x - a.x, b.y - a.y); if (L < 1e-6) return;
      const n = [-(b.y - a.y) / L, (b.x - a.x) / L, 0];
      const F = WGSecao.preencher(a.camadas, b.camadas, a.cota, b.cota);
      const pt = (s, prof) => [a.x + (b.x - a.x) * s, a.y + (b.y - a.y) * s, Z(a.cota + (b.cota - a.cota) * s - prof)];
      F.faixas.forEach(f => {
        const ca = f.ia != null ? F.a[f.ia] : null, cb = f.ib != null ? F.b[f.ib] : null;
        const c0 = rgb((ca || cb).hex), c1 = rgb((cb || ca).hex);
        for (let k = 0; k + 1 < F.s.length; k++) {
          if (f.base[k] - f.topo[k] < 1e-4 && f.base[k + 1] - f.topo[k + 1] < 1e-4) continue;
          const s0 = F.s[k], s1 = F.s[k + 1];
          const mix = s => c0.map((v, q) => v + (c1[q] - v) * s);
          quad(pt(s0, f.topo[k]), pt(s1, f.topo[k + 1]), pt(s1, f.base[k + 1]), pt(s0, f.base[k]), [mix(s0), mix(s1), mix(s1), mix(s0)], n);
        }
      });
      // terreno (topo do painel) e nível d'água ao longo do painel
      const off = [n[0] * w * 0.04, n[1] * w * 0.04, 0];
      const lado = (p, sgn) => [p[0] + off[0] * sgn, p[1] + off[1] * sgn, p[2]];
      linha(pt(0, 0), pt(1, 0), TERRA);
      if (agua && a.na != null && b.na != null && (a.rede || '-') === (b.rede || '-')) {
        const pa = [a.x, a.y, Z(a.cota - a.na)], pb = [b.x, b.y, Z(b.cota - b.na)];
        [1, -1].forEach(sg => linha(lado(pa, sg), lado(pb, sg), AZUL));
      }
      paineisFeitos++;
    });

    // colunas: uma caixa por camada
    const h = w / 2;
    cols.forEach(c => {
      const cam = c.camadas.length ? WGSecao.completar(c.camadas) : [{ de: 0, ate: Math.max(c.prof || 0, 0.5), hex: null, vazio: true }];
      cam.forEach((r, q) => {
        const cor = r.hex ? rgb(r.hex) : CINZA;
        const zt = Z(c.cota - r.de), zb = Z(c.cota - r.ate);
        const v = [[c.x - h, c.y - h], [c.x + h, c.y - h], [c.x + h, c.y + h], [c.x - h, c.y + h]];
        const nn = [[0, -1, 0], [1, 0, 0], [0, 1, 0], [-1, 0, 0]];
        for (let k = 0; k < 4; k++) {
          const p0 = v[k], p1 = v[(k + 1) % 4];
          quad([p0[0], p0[1], zt], [p1[0], p1[1], zt], [p1[0], p1[1], zb], [p0[0], p0[1], zb], cor, nn[k]);
        }
        if (q === 0) quad([v[0][0], v[0][1], zt], [v[1][0], v[1][1], zt], [v[2][0], v[2][1], zt], [v[3][0], v[3][1], zt], cor, [0, 0, 1]);
        if (q === cam.length - 1) quad([v[0][0], v[0][1], zb], [v[3][0], v[3][1], zb], [v[2][0], v[2][1], zb], [v[1][0], v[1][1], zb], cor, [0, 0, -1]);
      });
      // contorno da coluna (arestas verticais) para separar do painel
      const zt = Z(c.cota), zb = Z(c.cota - Math.max(c.prof || 0, cam[cam.length - 1].ate));
      [[-h, -h], [h, -h], [h, h], [-h, h]].forEach(([dx, dy]) => linha([c.x + dx, c.y + dy, zt], [c.x + dx, c.y + dy, zb], [0.09, 0.15, 0.13]));
      // anel azul no nível d'água
      if (agua && c.na != null) {
        const zn = Z(c.cota - c.na), r = h * 1.6;
        const q4 = [[-r, -r], [r, -r], [r, r], [-r, r]];
        for (let k = 0; k < 4; k++) linha([c.x + q4[k][0], c.y + q4[k][1], zn], [c.x + q4[(k + 1) % 4][0], c.y + q4[(k + 1) % 4][1], zn], AZUL);
      }
    });

    // caixa e grade no plano da base
    const xs = cols.map(c => c.x), ys = cols.map(c => c.y);
    const zTop = Math.max(...cols.map(c => c.cota)), zBot = Math.min(...cols.map(c => c.cota - Math.max(c.prof || 0, c.na || 0, 0.5)));
    const m = Math.max(2 * w, 0.06 * Math.max(Math.max(...xs) - Math.min(...xs), Math.max(...ys) - Math.min(...ys), 10));
    const caixa = { x0: Math.min(...xs) - m, x1: Math.max(...xs) + m, y0: Math.min(...ys) - m, y1: Math.max(...ys) + m, z0: Z(zBot), z1: Z(zTop), zTop, zBot };
    const passo = passoGrade(Math.max(caixa.x1 - caixa.x0, caixa.y1 - caixa.y0));
    const GRADE = [0.55, 0.6, 0.58];
    for (let x = Math.ceil(caixa.x0 / passo) * passo; x <= caixa.x1; x += passo) linha([x, caixa.y0, caixa.z0], [x, caixa.y1, caixa.z0], GRADE);
    for (let y = Math.ceil(caixa.y0 / passo) * passo; y <= caixa.y1; y += passo) linha([caixa.x0, y, caixa.z0], [caixa.x1, y, caixa.z0], GRADE);
    caixa.passo = passo;

    return {
      tri: { pos: new Float32Array(P), cor: new Float32Array(C), nor: new Float32Array(N) },
      lin: { pos: new Float32Array(LP), cor: new Float32Array(LC) },
      caixa, paineis: paineisFeitos,
      rotulos: cols.map(c => ({ nome: c.nome, x: c.x, y: c.y, z: Z(c.cota) }))
    };
  }
  function passoGrade(ext) {
    const alvo = ext / 8, p = Math.pow(10, Math.floor(Math.log10(Math.max(alvo, 1e-6))));
    return [1, 2, 5, 10].map(k => k * p).find(v => v >= alvo) || 10 * p;
  }

  /** Exagero automático: o bloco fica com altura ~35% da largura. */
  function exageroAuto(largura, altura) {
    if (!(altura > 0) || !(largura > 0)) return 1;
    const e = 0.35 * largura / altura;
    const opcoes = [1, 2, 3, 5, 10, 20, 50];
    return opcoes.reduce((m, v) => Math.abs(Math.log(v / e)) < Math.abs(Math.log(m / e)) ? v : m, 1);
  }

  return { rgb, arvoreMinima, espalharColunas, geometria, passoGrade, exageroAuto };
})();

/* =====================================================================
   Visor 3D do diagrama de cercas: WebGL puro (sem biblioteca).
   Girar = arrastar; mover = botão direito, Shift+arrastar ou dois dedos;
   aproximar = roda ou pinça.
   ===================================================================== */
const Visor3D = (() => {
  const VS = `attribute vec3 p; attribute vec3 c; attribute vec3 n;
    uniform mat4 m; uniform vec3 luz; uniform float lit;
    varying vec3 vc;
    void main(){
      gl_Position = m * vec4(p, 1.0);
      float d = abs(dot(normalize(n), luz));
      vc = lit > 0.5 ? c * (0.62 + 0.38 * d) : c;
    }`;
  const FS = `precision mediump float; varying vec3 vc; void main(){ gl_FragColor = vec4(vc, 1.0); }`;

  // --- matrizes 4x4 (coluna principal, como o WebGL espera)
  const mul = (a, b) => { const r = new Float32Array(16); for (let i = 0; i < 4; i++) for (let j = 0; j < 4; j++) { let s = 0; for (let k = 0; k < 4; k++) s += a[k * 4 + j] * b[i * 4 + k]; r[i * 4 + j] = s; } return r; };
  const persp = (fov, asp, n, f) => { const t = 1 / Math.tan(fov / 2); return new Float32Array([t / asp, 0, 0, 0, 0, t, 0, 0, 0, 0, (f + n) / (n - f), -1, 0, 0, 2 * f * n / (n - f), 0]); };
  function olhar(o, alvo, cima) {
    const sub = (a, b) => [a[0] - b[0], a[1] - b[1], a[2] - b[2]];
    const cruz = (a, b) => [a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0]];
    const nor = a => { const l = Math.hypot(...a) || 1; return a.map(v => v / l); };
    const z = nor(sub(o, alvo)), x = nor(cruz(cima, z)), y = cruz(z, x);
    const d = v => -(v[0] * o[0] + v[1] * o[1] + v[2] * o[2]);
    return new Float32Array([x[0], y[0], z[0], 0, x[1], y[1], z[1], 0, x[2], y[2], z[2], 0, d(x), d(y), d(z), 1]);
  }

  function abrir(area, opcoes = {}) {
    const cv = document.createElement('canvas');
    cv.className = 'cerca3d-canvas';
    area.appendChild(cv);
    const camRot = document.createElement('div'); camRot.className = 'cerca3d-rotulos'; area.appendChild(camRot);
    const gl = cv.getContext('webgl', { antialias: true, preserveDrawingBuffer: true });
    if (!gl) { cv.remove(); camRot.remove(); return null; }
    const sh = (t, s) => { const o = gl.createShader(t); gl.shaderSource(o, s); gl.compileShader(o); return o; };
    const prog = gl.createProgram();
    gl.attachShader(prog, sh(gl.VERTEX_SHADER, VS)); gl.attachShader(prog, sh(gl.FRAGMENT_SHADER, FS)); gl.linkProgram(prog);
    if (!gl.getProgramParameter(prog, gl.LINK_STATUS)) { cv.remove(); camRot.remove(); return null; }
    gl.useProgram(prog);
    const A = { p: gl.getAttribLocation(prog, 'p'), c: gl.getAttribLocation(prog, 'c'), n: gl.getAttribLocation(prog, 'n') };
    const U = { m: gl.getUniformLocation(prog, 'm'), luz: gl.getUniformLocation(prog, 'luz'), lit: gl.getUniformLocation(prog, 'lit') };
    gl.enable(gl.DEPTH_TEST);
    gl.enable(gl.POLYGON_OFFSET_FILL); gl.polygonOffset(1, 1);   // as linhas ficam por cima das faces
    const bufs = [];
    const novoBuf = dados => { const b = gl.createBuffer(); gl.bindBuffer(gl.ARRAY_BUFFER, b); gl.bufferData(gl.ARRAY_BUFFER, dados, gl.STATIC_DRAW); bufs.push(b); return b; };
    let G = null, B = null, rotulos = [];
    const cam = { yaw: -35, pitch: 32, dist: 100, alvo: [0, 0, 0] };
    let fundo = [1, 1, 1], vistaIni = { yaw: -35, pitch: 32 };

    // o mundo do WebGL tem y para cima: (x, y, z) dos dados -> (x, z, -y)
    const mundo = (x, y, z) => [x, z, -y];
    function matriz() {
      const W = cv.width, H = cv.height;
      const yr = cam.yaw * Math.PI / 180, pr = cam.pitch * Math.PI / 180;
      const o = [cam.alvo[0] + cam.dist * Math.cos(pr) * Math.sin(yr), cam.alvo[1] + cam.dist * Math.sin(pr), cam.alvo[2] + cam.dist * Math.cos(pr) * Math.cos(yr)];
      return { M: mul(persp(45 * Math.PI / 180, W / H, cam.dist / 200, cam.dist * 20), olhar(o, cam.alvo, [0, 1, 0])), o };
    }
    function projetar(M, x, y, z) {
      const [a, b, c] = mundo(x, y, z);
      const X = M[0] * a + M[4] * b + M[8] * c + M[12], Y = M[1] * a + M[5] * b + M[9] * c + M[13], Wc = M[3] * a + M[7] * b + M[11] * c + M[15];
      if (Wc <= 0) return null;
      return [(X / Wc * 0.5 + 0.5) * cv.clientWidth, (1 - (Y / Wc * 0.5 + 0.5)) * cv.clientHeight];
    }

    let pedido = 0, vivo = true;
    function desenhar() {
      pedido = 0; if (!vivo) return;
      const dpr = Math.min(2, window.devicePixelRatio || 1);
      const W = Math.max(1, Math.round(cv.clientWidth * dpr)), H = Math.max(1, Math.round(cv.clientHeight * dpr));
      if (cv.width !== W || cv.height !== H) { cv.width = W; cv.height = H; }
      gl.viewport(0, 0, W, H);
      gl.clearColor(fundo[0], fundo[1], fundo[2], 1);
      gl.clear(gl.COLOR_BUFFER_BIT | gl.DEPTH_BUFFER_BIT);
      if (!G) return;
      const { M } = matriz();
      gl.uniformMatrix4fv(U.m, false, M);
      const yr = cam.yaw * Math.PI / 180;
      const l = [Math.sin(yr + 0.6) * 0.75, 0.55, Math.cos(yr + 0.6) * 0.75], ll = Math.hypot(...l);
      gl.uniform3f(U.luz, l[0] / ll, l[1] / ll, l[2] / ll);
      const ligar = (buf, loc, n) => { if (loc < 0) return; gl.bindBuffer(gl.ARRAY_BUFFER, buf); gl.enableVertexAttribArray(loc); gl.vertexAttribPointer(loc, n, gl.FLOAT, false, 0, 0); };
      if (B.nt) {
        gl.uniform1f(U.lit, 1);
        ligar(B.tp, A.p, 3); ligar(B.tc, A.c, 3); ligar(B.tn, A.n, 3);
        gl.drawArrays(gl.TRIANGLES, 0, B.nt);
      }
      if (B.nl) {
        gl.uniform1f(U.lit, 0);
        ligar(B.lp, A.p, 3); ligar(B.lc, A.c, 3);
        if (A.n >= 0) { gl.disableVertexAttribArray(A.n); gl.vertexAttrib3f(A.n, 0, 1, 0); }
        gl.drawArrays(gl.LINES, 0, B.nl);
      }
      // nomes das sondagens
      rotulos.forEach(r => {
        const s = projetar(M, r.x, r.y, r.z);
        if (!s) { r.el.hidden = true; return; }
        r.el.hidden = false; r.el.style.transform = `translate(${s[0].toFixed(1)}px, ${s[1].toFixed(1)}px) translate(-50%, -150%)`;
      });
      if (opcoes.aoDesenhar) opcoes.aoDesenhar();
    }
    const pedir = () => { if (!pedido) pedido = requestAnimationFrame(desenhar); };

    /** Distância da câmera para a caixa toda caber na tela (com margem), na direção atual. */
    function enquadrar(k) {
      if (!cv.clientWidth) { cv.width = 800; cv.height = 500; }
      const cantos = [];
      [k.x0, k.x1].forEach(x => [k.y0, k.y1].forEach(y => [k.z0, k.z1].forEach(z => cantos.push([x, y, z]))));
      const diag = Math.hypot(k.x1 - k.x0, k.y1 - k.y0, k.z1 - k.z0) || 10;
      const cabe = () => {
        const { M } = matriz();
        return cantos.every(([x, y, z]) => {
          const [a, b, c] = mundo(x, y, z);
          const X = M[0] * a + M[4] * b + M[8] * c + M[12], Y = M[1] * a + M[5] * b + M[9] * c + M[13], Wc = M[3] * a + M[7] * b + M[11] * c + M[15];
          return Wc > 0 && Math.abs(X / Wc) < 0.94 && Math.abs(Y / Wc) < 0.9;
        });
      };
      let lo = diag * 0.1, hi = diag * 4;
      for (let q = 0; q < 30; q++) { cam.dist = (lo + hi) / 2; if (cabe()) hi = cam.dist; else lo = cam.dist; }
      cam.dist = hi;
    }

    // os dados chegam em (x leste, y norte, z cota); passam para o mundo do WebGL
    function trocarGeometria(g, { reenquadrar = false } = {}) {
      bufs.forEach(b => gl.deleteBuffer(b)); bufs.length = 0;
      const conv = arr => { const o = new Float32Array(arr.length); for (let i = 0; i < arr.length; i += 3) { o[i] = arr[i]; o[i + 1] = arr[i + 2]; o[i + 2] = -arr[i + 1]; } return o; };
      B = {
        tp: novoBuf(conv(g.tri.pos)), tc: novoBuf(g.tri.cor), tn: novoBuf(conv(g.tri.nor)), nt: g.tri.pos.length / 3,
        lp: novoBuf(conv(g.lin.pos)), lc: novoBuf(g.lin.cor), nl: g.lin.pos.length / 3
      };
      const k = g.caixa;
      const centro = mundo((k.x0 + k.x1) / 2, (k.y0 + k.y1) / 2, (k.z0 + k.z1) / 2);
      if (reenquadrar || !G) { cam.alvo = centro; enquadrar(k); }
      else { cam.alvo[1] = centro[1]; }                 // trocou o exagero: mantém a vista, ajusta a altura
      G = g;
      camRot.innerHTML = '';
      rotulos = g.rotulos.map(r => { const el = document.createElement('span'); el.textContent = r.nome; camRot.appendChild(el); return { ...r, el }; });
      pedir();
    }

    // mouse e toque
    const toques = new Map(); let pinca = null, modo = null;
    cv.addEventListener('contextmenu', e => e.preventDefault());
    cv.addEventListener('pointerdown', e => {
      cv.setPointerCapture(e.pointerId); toques.set(e.pointerId, { x: e.clientX, y: e.clientY });
      modo = (e.button === 2 || e.shiftKey) ? 'mover' : 'girar'; cv.classList.add('arrastando');
    });
    function mover(dx, dy) {
      const yr = cam.yaw * Math.PI / 180, k = cam.dist * 0.0016;
      const dir = [Math.cos(yr), 0, -Math.sin(yr)];
      cam.alvo[0] -= dir[0] * dx * k; cam.alvo[2] -= dir[2] * dx * k; cam.alvo[1] += dy * k;
    }
    cv.addEventListener('pointermove', e => {
      const a = toques.get(e.pointerId);
      if (!a) { if (opcoes.aoPassar) opcoes.aoPassar(e); return; }
      const dx = e.clientX - a.x, dy = e.clientY - a.y;
      a.x = e.clientX; a.y = e.clientY;
      if (toques.size === 2) {
        const [p1, p2] = [...toques.values()], d = Math.hypot(p1.x - p2.x, p1.y - p2.y);
        if (pinca) { cam.dist = Math.max(1, pinca.dist * pinca.d / Math.max(10, d)); mover(dx / 2, dy / 2); } else pinca = { d, dist: cam.dist };
      } else if (modo === 'mover') mover(dx, dy);
      else { cam.yaw -= dx * 0.4; cam.pitch = Math.max(-5, Math.min(89, cam.pitch + dy * 0.3)); }
      pedir();
    });
    const soltar = e => { toques.delete(e.pointerId); pinca = null; if (!toques.size) { modo = null; cv.classList.remove('arrastando'); } };
    cv.addEventListener('pointerup', soltar); cv.addEventListener('pointercancel', soltar);
    cv.addEventListener('pointerleave', () => { if (opcoes.aoSair) opcoes.aoSair(); });
    cv.addEventListener('wheel', e => { e.preventDefault(); cam.dist = Math.max(1, cam.dist * Math.exp(e.deltaY * 0.0012)); pedir(); }, { passive: false });
    const ro = window.ResizeObserver ? new ResizeObserver(pedir) : null;
    if (ro) ro.observe(cv); else window.addEventListener('resize', pedir);

    return {
      trocarGeometria,
      fundo(c) { fundo = c; pedir(); },
      /** Direção inicial da câmera (graus): usada agora e no botão "voltar à vista inicial". */
      vistaInicial(yaw, pitch) { vistaIni = { yaw, pitch }; cam.yaw = yaw; cam.pitch = pitch; },
      vista(qual) {
        if (!G) return;
        if (qual === 'cima') { cam.pitch = 89; cam.yaw = 0; }
        else if (qual === 'inicio') { cam.yaw = vistaIni.yaw; cam.pitch = vistaIni.pitch; trocarGeometria(G, { reenquadrar: true }); return; }
        pedir();
      },
      zoom(k) { cam.dist = Math.max(1, cam.dist * k); pedir(); },
      girar(d) { cam.yaw += d; pedir(); },
      /** Ponto da tela -> coluna mais próxima (para a dica ao passar o mouse). */
      colunaEm(px, py, cols, ex, w) {
        if (!G) return null;
        const { M, o } = matriz(); let melhor = null;
        cols.forEach((c, i) => {
          const fundo = Math.max(c.prof || 0, 0.5);
          const t = projetar(M, c.x, c.y, c.cota * ex), b = projetar(M, c.x, c.y, (c.cota - fundo) * ex);
          if (!t || !b) return;
          const vx = b[0] - t[0], vy = b[1] - t[1], L2 = vx * vx + vy * vy || 1;
          const s = Math.max(0, Math.min(1, ((px - t[0]) * vx + (py - t[1]) * vy) / L2));
          const d = Math.hypot(px - t[0] - vx * s, py - t[1] - vy * s);
          if (d > 12) return;
          const p = mundo(c.x, c.y, (c.cota - s * fundo) * ex), longe = Math.hypot(p[0] - o[0], p[1] - o[1], p[2] - o[2]);
          if (!melhor || longe < melhor.longe) melhor = { i, prof: s * fundo, longe };
        });
        return melhor;
      },
      /** Imagem PNG com os nomes das sondagens. */
      imagem() {
        desenhar();
        const out = document.createElement('canvas'); out.width = cv.width; out.height = cv.height;
        const ctx = out.getContext('2d'); ctx.drawImage(cv, 0, 0);
        const k = cv.width / cv.clientWidth, { M } = matriz();
        ctx.font = `600 ${Math.round(11 * k)}px "IBM Plex Mono", monospace`; ctx.textAlign = 'center';
        rotulos.forEach(r => {
          const s = projetar(M, r.x, r.y, r.z); if (!s) return;
          const x = s[0] * k, y = s[1] * k - 8 * k;
          ctx.lineWidth = 3 * k; ctx.strokeStyle = 'rgba(255,255,255,.9)'; ctx.strokeText(r.nome, x, y);
          ctx.fillStyle = '#172621'; ctx.fillText(r.nome, x, y);
        });
        return out.toDataURL('image/png');
      },
      destruir() {
        vivo = false; if (pedido) cancelAnimationFrame(pedido);
        if (ro) ro.disconnect(); else window.removeEventListener('resize', pedir);
        bufs.forEach(b => gl.deleteBuffer(b)); gl.deleteProgram(prog);
        const ext = gl.getExtension('WEBGL_lose_context'); if (ext) ext.loseContext();
        cv.remove(); camRot.remove();
      }
    };
  }
  return { abrir };
})();

/* =====================================================================
   3g. MALHA UTM DO MAPA (funções puras, sem tela)
   Linhas de E e N constantes que cobrem a área visível, num passo "redondo"
   (10, 20, 25, 50, 100 m...). As linhas são calculadas em UTM e convertidas
   para latitude/longitude com as mesmas contas da planta DXF.
   ===================================================================== */
const WGMalha = (() => {
  /** Passo redondo para ter perto de `alvo` linhas na extensão dada (m). */
  function passo(ext, alvo = 6) {
    const bruto = Math.max(ext, 1e-6) / Math.max(1, alvo), p = Math.pow(10, Math.floor(Math.log10(bruto)));
    return [1, 2, 2.5, 5, 10].map(k => k * p).find(v => v >= bruto * 0.999);
  }

  /**
   * b = {s, w, n, e} em graus (limites do mapa). Devolve
   * { zona, sul, passo, verticais: [{v: E, pts: [[lat, lon]...]}], horizontais: [{v: N, pts}] }.
   */
  function linhas(b, alvo = 6) {
    const lat0 = (b.s + b.n) / 2, lon0 = (b.w + b.e) / 2;
    const zona = WGDxf.zonaDaLongitude(lon0), sul = lat0 < 0;
    const cantos = [[b.s, b.w], [b.s, b.e], [b.n, b.w], [b.n, b.e]].map(([la, lo]) => WGDxf.latLonParaUtm(la, lo, zona, sul));
    const E0 = Math.min(...cantos.map(c => c.E)), E1 = Math.max(...cantos.map(c => c.E));
    const N0 = Math.min(...cantos.map(c => c.N)), N1 = Math.max(...cantos.map(c => c.N));
    const p = passo(E1 - E0, alvo);
    const fora = (E1 - E0) / p > 60 || (N1 - N0) / p > 60;   // segurança: nunca desenha centenas de linhas
    const AMOSTRAS = 8;
    const ll = (E, N) => { const r = WGDxf.utmParaLatLon(E, N, zona, sul); return [r.lat, r.lon]; };
    const verticais = [], horizontais = [];
    if (!fora) {
      for (let E = Math.ceil(E0 / p) * p; E <= E1 + 1e-6; E += p)
        verticais.push({ v: E, pts: Array.from({ length: AMOSTRAS + 1 }, (_, k) => ll(E, N0 + (N1 - N0) * k / AMOSTRAS)) });
      for (let N = Math.ceil(N0 / p) * p; N <= N1 + 1e-6; N += p)
        horizontais.push({ v: N, pts: Array.from({ length: AMOSTRAS + 1 }, (_, k) => ll(E0 + (E1 - E0) * k / AMOSTRAS, N)) });
    }
    return { zona, sul, passo: p, verticais, horizontais };
  }

  /** Onde uma linha (em pixels da tela) cruza a reta x = valor ou y = valor. null se não cruza. */
  function cruzamento(pxs, eixo, valor) {
    for (let k = 0; k + 1 < pxs.length; k++) {
      const a = pxs[k], b = pxs[k + 1], da = a[eixo] - valor, db = b[eixo] - valor;
      if (da === 0) return a;
      if (da * db < 0) { const t = da / (da - db); return { x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t }; }
    }
    return null;
  }

  /** 345600 -> "345.600"; com passo menor que 1 m mostra decimais. */
  function texto(v, p = 1) {
    return Number(v).toLocaleString('pt-BR', { maximumFractionDigits: p < 1 ? 1 : 0, minimumFractionDigits: p < 1 ? 1 : 0 });
  }

  return { passo, linhas, cruzamento, texto };
})();

/* =====================================================================
   3h. PLUMA VERTICAL NA SEÇÃO (funções puras, sem tela)
   Interpola a concentração no plano do corte (distância ao longo da seção ×
   cota). Cada amostra fica no meio do filtro do poço. IDW em escala log, com
   anisotropia (a distância vertical pesa mais que a horizontal, como Kh > Kv),
   só na zona saturada (entre o N.A. e o fundo das sondagens). Fora das amostras
   o valor cai até o piso, como na pluma do mapa.
   ===================================================================== */
const WGPlumaV = (() => {
  /** Interpolação linear por partes de [[s, z]...] (ordenado por s); fora das pontas, o valor da ponta. */
  function perfil(pts) {
    const p = pts.filter(q => Number.isFinite(q[0]) && Number.isFinite(q[1])).sort((a, b) => a[0] - b[0]);
    return s => {
      if (!p.length) return null;
      if (s <= p[0][0]) return p[0][1];
      for (let k = 0; k + 1 < p.length; k++) {
        const [s0, z0] = p[k], [s1, z1] = p[k + 1];
        if (s <= s1) return s1 - s0 < 1e-9 ? Math.max(z0, z1) : z0 + (z1 - z0) * (s - s0) / (s1 - s0);
      }
      return p[p.length - 1][1];
    };
  }

  /**
   * amostras: [{s, z, v}] (v > 0). o = {s0, s1, z0, z1, nx, nz, p, anis, piso, topo(s), base(s)}.
   * Devolve { nx, nz, v: Float64Array(nx*nz) } com NaN fora da zona saturada. Célula (i, j): i ao longo de s, j de cima (z1) para baixo.
   */
  function grade(amostras, o) {
    const am = amostras.filter(a => a.v > 0 && Number.isFinite(a.s) && Number.isFinite(a.z));
    const v = new Float64Array(o.nx * o.nz).fill(NaN);
    if (!am.length) return { nx: o.nx, nz: o.nz, v };
    const p = o.p || 2, k = o.anis || 10, lPiso = Math.log10(o.piso);
    const L = am.map(a => Math.log10(a.v));
    const ss = am.map(a => a.s), zs = am.map(a => a.z);
    const sMin = Math.min(...ss), sMax = Math.max(...ss), zMin = Math.min(...zs);
    // distância de decaimento para os lados: metade do espaçamento típico entre os poços (mínimo 5 m)
    const sord = [...new Set(ss.map(x => Math.round(x * 10) / 10))].sort((a, b) => a - b);
    const gaps = sord.slice(1).map((x, q) => x - sord[q]).filter(g => g > 0.5).sort((a, b) => a - b);
    const espH = Math.max(5, 0.5 * (gaps.length ? gaps[Math.floor(gaps.length / 2)] : (o.s1 - o.s0) * 0.2));
    const espV = Math.max(2, (Math.max(...zs) - zMin) * 0.5 + 1);
    for (let j = 0; j < o.nz; j++) {
      const z = o.z1 - (j + 0.5) * (o.z1 - o.z0) / o.nz;
      for (let i = 0; i < o.nx; i++) {
        const s = o.s0 + (i + 0.5) * (o.s1 - o.s0) / o.nx;
        const tp = o.topo ? o.topo(s) : null, bs = o.base ? o.base(s) : null;
        if ((tp != null && z > tp) || (bs != null && z < bs)) continue;
        let sw = 0, sl = 0, exato = null;
        for (let q = 0; q < am.length; q++) {
          const d = Math.hypot(s - ss[q], (z - zs[q]) * k);
          if (d < 1e-6) { exato = L[q]; break; }
          const w = 1 / Math.pow(d, p); sw += w; sl += w * L[q];
        }
        let l = exato != null ? exato : sl / sw;
        // fora da faixa de poços (para os lados e para baixo) o valor vai caindo até o piso
        const fh = s < sMin ? (sMin - s) / espH : s > sMax ? (s - sMax) / espH : 0;
        const fv = z < zMin ? (zMin - z) / espV : 0;
        const f = Math.min(1, Math.max(fh, fv));
        if (f > 0 && l > lPiso) l = l * (1 - f) + lPiso * f;
        v[j * o.nx + i] = Math.pow(10, l);
      }
    }
    return { nx: o.nx, nz: o.nz, v };
  }

  /** Linha do limite (marching squares) em coordenadas de célula: [[x0, y0, x1, y1]...] (x = i + 0,5, y = j + 0,5). */
  function isolinha(g, limiar) {
    const segs = [], V = (i, j) => { const x = g.v[j * g.nx + i]; return Number.isNaN(x) ? -Infinity : x; };
    const t = (a, b) => (a === b || !Number.isFinite(a) || !Number.isFinite(b)) ? 0.5 : Math.min(1, Math.max(0, (limiar - a) / (b - a)));
    for (let j = 0; j + 1 < g.nz; j++) for (let i = 0; i + 1 < g.nx; i++) {
      const a = V(i, j), b = V(i + 1, j), c = V(i + 1, j + 1), d = V(i, j + 1);
      const caso = (a >= limiar ? 8 : 0) | (b >= limiar ? 4 : 0) | (c >= limiar ? 2 : 0) | (d >= limiar ? 1 : 0);
      if (caso === 0 || caso === 15) continue;
      const x = i + 0.5, y = j + 0.5;
      const topo = [x + t(a, b), y], dir = [x + 1, y + t(b, c)], baixo = [x + t(d, c), y + 1], esq = [x, y + t(a, d)];
      const lig = { 1: [[esq, baixo]], 2: [[baixo, dir]], 3: [[esq, dir]], 4: [[topo, dir]], 5: [[esq, topo], [baixo, dir]], 6: [[topo, baixo]], 7: [[esq, topo]],
        8: [[esq, topo]], 9: [[topo, baixo]], 10: [[esq, baixo], [topo, dir]], 11: [[topo, dir]], 12: [[esq, dir]], 13: [[baixo, dir]], 14: [[esq, baixo]] }[caso];
      lig.forEach(([p, q]) => segs.push([p[0], p[1], q[0], q[1]]));
    }
    return segs;
  }

  /**
   * Pluma em "perfis": cada ponto da linha (um poço, ou um grupo multinível) tem um perfil do log da
   * concentração com a profundidade abaixo do N.A.: um poço só vai do valor medido no N.A. até o limite no fundo
   * do filtro; um grupo multinível passa por cada amostra e cai até o limite no fundo do filtro mais fundo.
   * Na horizontal, IDW em log entre os perfis, na mesma profundidade abaixo do N.A.
   * grupos: [{s, pts: [{d, dBase, l}]}] (d = prof. do meio do filtro abaixo do N.A., dBase = do fundo do filtro).
   * o = {s0, s1, z0, z1, nx, nz, p, lLim, lPiso, espH, topo(s)}.
   */
  function gradePerfis(grupos, o) {
    const v = new Float64Array(o.nx * o.nz).fill(NaN);
    if (!grupos.length) return { nx: o.nx, nz: o.nz, v };
    const perfis = grupos.map(g => {
      const pts = [...g.pts].sort((a, b) => a.d - b.d), u = pts[pts.length - 1];
      // um poço só: cai em linha reta do N.A. até o limite no fundo do filtro (a pluma afina com a distância);
      // multinível: passa por cada amostra e cai abaixo da mais funda
      const nos = [[0, pts[0].l], ...(pts.length > 1 ? pts.map(q => [q.d, q.l]) : [])];
      const dB = Math.max(u.dBase, u.d + 0.3);
      if (u.l > o.lLim) nos.push([dB, o.lLim]);
      const lUlt = Math.min(u.l, o.lLim);
      nos.push([dB + Math.max(1, dB - u.d), Math.min(lUlt, o.lPiso)]);
      return { s: g.s, f: d => {
        if (d <= nos[0][0]) return nos[0][1];
        for (let k = 0; k + 1 < nos.length; k++) if (d <= nos[k + 1][0]) {
          const [d0, l0] = nos[k], [d1, l1] = nos[k + 1];
          return d1 - d0 < 1e-9 ? l1 : l0 + (l1 - l0) * (d - d0) / (d1 - d0);
        }
        return nos[nos.length - 1][1];
      } };
    });
    const ss = perfis.map(q => q.s), sMin = Math.min(...ss), sMax = Math.max(...ss), p = o.p || 2;
    for (let i = 0; i < o.nx; i++) {
      const s = o.s0 + (i + 0.5) * (o.s1 - o.s0) / o.nx, na = o.topo(s);
      const w = perfis.map(q => { const dd = Math.abs(s - q.s); return dd < 1e-6 ? Infinity : 1 / Math.pow(dd, p); });
      const exato = w.findIndex(x => x === Infinity);
      const sw = w.reduce((a, b) => a + b, 0);
      const fh = Math.min(1, s < sMin ? (sMin - s) / o.espH : s > sMax ? (s - sMax) / o.espH : 0);
      for (let j = 0; j < o.nz; j++) {
        const z = o.z1 - (j + 0.5) * (o.z1 - o.z0) / o.nz, d = na - z;
        let l;
        if (exato >= 0) l = perfis[exato].f(d);
        else { l = 0; perfis.forEach((q, k) => { l += w[k] * q.f(d); }); l /= sw; }
        if (fh > 0 && l > o.lPiso) l = l * (1 - fh) + o.lPiso * fh;
        v[j * o.nx + i] = Math.pow(10, l);
      }
    }
    return { nx: o.nx, nz: o.nz, v };
  }

  return { perfil, grade, gradePerfis, isolinha };
})();

/* =====================================================================
   3i. GRAVAÇÃO DE DXF (funções puras, sem tela)
   DXF R12 (AC1009) em texto, que abre no AutoCAD, BricsCAD, QGIS, LibreCAD...
   Entidades: POLYLINE, SOLID (preenchimento), TEXT e CIRCLE, com camadas
   e cores ACI (o R12 não guarda cor RGB: usa a cor ACI mais próxima).
   ===================================================================== */
const WGDxfSaida = (() => {
  // paleta ACI aproximada: 1-9 fixas, 10-249 por matiz (15°) e brilho, 250-255 cinzas
  const PALETA = (() => {
    const p = [null, [255, 0, 0], [255, 255, 0], [0, 255, 0], [0, 255, 255], [0, 0, 255], [255, 0, 255], [0, 0, 0], [128, 128, 128], [192, 192, 192]];
    const brilho = [1, .8, .6, .5, .3];
    for (let i = 10; i <= 249; i++) {
      const h = Math.floor((i - 10) / 10) * 15, sub = i % 10, v = brilho[Math.floor(sub / 2)], sat = sub % 2 ? .5 : 1;
      const f = n => { const k = (n + h / 60) % 6; return v * (1 - sat * Math.max(0, Math.min(k, 4 - k, 1))); };
      p[i] = [f(5), f(3), f(1)].map(x => Math.round(x * 255));
    }
    [51, 91, 132, 173, 214, 255].forEach((g, k) => { p[250 + k] = [g, g, g]; });
    return p;
  })();

  /** "#a63603" -> índice ACI mais próximo (distância simples em RGB, com peso de percepção). 7 = preto/branco. */
  function aci(hex) {
    const m = /^#?([0-9a-f]{6})$/i.exec(String(hex || '').trim());
    if (!m) return 7;
    const c = [0, 2, 4].map(k => parseInt(m[1].slice(k, k + 2), 16));
    let melhor = 7, dm = Infinity;
    for (let i = 1; i <= 255; i++) {
      if (i === 7) continue;
      const q = PALETA[i], dr = c[0] - q[0], dg = c[1] - q[1], db = c[2] - q[2];
      const d = 2 * dr * dr + 4 * dg * dg + 3 * db * db;
      if (d < dm) { dm = d; melhor = i; }
    }
    return c[0] + c[1] + c[2] < 135 ? 7 : melhor;   // quase preto: usa a 7 (preto no papel, branco na tela do CAD)
  }

  /** Nome de camada aceito no R12: sem acento, maiúsculas, letras/números/_ e -, até 31 caracteres. */
  function nomeCamada(t) {
    return (String(t || 'CAMADA').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z0-9_\-]+/g, '_').replace(/^_+|_+$/g, '') || 'CAMADA').slice(0, 31);
  }

  /** Texto para o DXF em Windows-1252: tira os caracteres que não existem nele. */
  function texto1252(t) {
    return String(t).replace(/[–—]/g, '-').replace(/[‘’′]/g, "'").replace(/[“”]/g, '"').replace(/↔/g, '<->').replace(/·/g, '-')
      .replace(/[^\x20-\x7e\xa0-\xff]/g, '?');
  }

  const num = v => { const r = Math.round(v * 1e4) / 1e4; return Object.is(r, -0) ? '0' : String(r); };

  /**
   * ents: [{t:'poli', pts:[[x,y]...], fechado, camada, cor, tipo:'TRACO'?} | {t:'solido', pts:[3 ou 4 pontos em ordem de contorno], camada, cor}
   *        | {t:'texto', x, y, alt, txt, rot, al:0|1|2, camada, cor} | {t:'circ', x, y, r, camada, cor}]
   * o = {traco: [comp. do traço, comp. do espaço]} (em unidades do desenho). Devolve o texto do DXF.
   */
  function escrever(ents, o = {}) {
    const L = [], gc = (c, v) => { L.push(String(c), String(v)); };
    const camadas = new Map();
    ents.forEach(e => { if (!camadas.has(e.camada)) camadas.set(e.camada, e.cor || 7); });
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    const ext = (x, y) => { x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y); };
    ents.forEach(e => (e.pts || [[e.x, e.y]]).forEach(p => ext(p[0], p[1])));
    if (!Number.isFinite(x0)) { x0 = y0 = 0; x1 = y1 = 1; }
    const [tr, es] = o.traco || [1, .6];

    gc(0, 'SECTION'); gc(2, 'HEADER');
    gc(9, '$ACADVER'); gc(1, 'AC1009');
    gc(9, '$DWGCODEPAGE'); gc(3, 'ANSI_1252');
    gc(9, '$INSUNITS'); gc(70, 6);
    gc(9, '$EXTMIN'); gc(10, num(x0)); gc(20, num(y0)); gc(30, 0);
    gc(9, '$EXTMAX'); gc(10, num(x1)); gc(20, num(y1)); gc(30, 0);
    gc(9, '$FILLMODE'); gc(70, 1);
    gc(0, 'ENDSEC');

    gc(0, 'SECTION'); gc(2, 'TABLES');
    gc(0, 'TABLE'); gc(2, 'LTYPE'); gc(70, 2);
    gc(0, 'LTYPE'); gc(2, 'CONTINUOUS'); gc(70, 0); gc(3, 'Solid line'); gc(72, 65); gc(73, 0); gc(40, 0);
    gc(0, 'LTYPE'); gc(2, 'TRACO'); gc(70, 0); gc(3, '__ __ __'); gc(72, 65); gc(73, 2); gc(40, num(tr + es)); gc(49, num(tr)); gc(49, num(-es));
    gc(0, 'ENDTAB');
    gc(0, 'TABLE'); gc(2, 'LAYER'); gc(70, camadas.size + 1);
    gc(0, 'LAYER'); gc(2, '0'); gc(70, 0); gc(62, 7); gc(6, 'CONTINUOUS');
    camadas.forEach((cor, nome) => { gc(0, 'LAYER'); gc(2, nome); gc(70, 0); gc(62, cor); gc(6, 'CONTINUOUS'); });
    gc(0, 'ENDTAB');
    gc(0, 'ENDSEC');

    gc(0, 'SECTION'); gc(2, 'ENTITIES');
    ents.forEach(e => {
      const cab = tipo => { gc(0, tipo); gc(8, e.camada); gc(62, e.cor || 7); if (e.tipo) gc(6, e.tipo); };
      if (e.t === 'poli' && e.pts.length >= 2) {
        cab('POLYLINE'); gc(66, 1); gc(10, 0); gc(20, 0); gc(30, 0); gc(70, e.fechado ? 1 : 0);
        e.pts.forEach(p => { gc(0, 'VERTEX'); gc(8, e.camada); gc(10, num(p[0])); gc(20, num(p[1])); gc(30, 0); });
        gc(0, 'SEQEND'); gc(8, e.camada);
      } else if (e.t === 'solido' && e.pts.length >= 3) {
        const [a, b, c, d = c] = e.pts;                 // ordem do SOLID: 1, 2, 4, 3
        cab('SOLID');
        gc(10, num(a[0])); gc(20, num(a[1])); gc(30, 0); gc(11, num(b[0])); gc(21, num(b[1])); gc(31, 0);
        gc(12, num(d[0])); gc(22, num(d[1])); gc(32, 0); gc(13, num(c[0])); gc(23, num(c[1])); gc(33, 0);
      } else if (e.t === 'texto' && String(e.txt).trim()) {
        cab('TEXT'); gc(10, num(e.x)); gc(20, num(e.y)); gc(30, 0); gc(40, num(e.alt)); gc(1, texto1252(e.txt));
        if (e.rot) gc(50, num(e.rot));
        if (e.al) { gc(72, e.al); gc(11, num(e.x)); gc(21, num(e.y)); gc(31, 0); }
      } else if (e.t === 'circ') {
        cab('CIRCLE'); gc(10, num(e.x)); gc(20, num(e.y)); gc(30, 0); gc(40, num(e.r));
      }
    });
    gc(0, 'ENDSEC'); gc(0, 'EOF');
    return L.join('\r\n') + '\r\n';
  }

  /** Texto -> bytes Windows-1252 (os caracteres já foram limpos por texto1252). */
  function bytes(txt) {
    const b = new Uint8Array(txt.length);
    for (let i = 0; i < txt.length; i++) { const c = txt.charCodeAt(i); b[i] = c < 256 ? c : 63; }
    return b;
  }

  /** "M1 2 L3 4 h5 v-2 l1 1 Z M..." -> [{pts:[[x,y]...], fechado}] (só comandos de reta). */
  function caminho(d) {
    const tk = String(d).match(/[MLHVZmlhvz]|-?\d*\.?\d+(?:e[-+]?\d+)?/g) || [];
    const out = []; let cmd = 'M', cur = null, x = 0, y = 0, k = 0;
    const novo = () => { cur = { pts: [[x, y]], fechado: false }; out.push(cur); };
    while (k < tk.length) {
      if (/[A-Za-z]/.test(tk[k])) { cmd = tk[k++]; if (cmd === 'Z' || cmd === 'z') { if (cur) { cur.fechado = true; [x, y] = cur.pts[0]; } cur = null; continue; } }
      const n = () => +tk[k++];
      if (cmd === 'M' || cmd === 'm') { const a = n(), b = n(); if (cmd === 'm') { x += a; y += b; } else { x = a; y = b; } novo(); cmd = cmd === 'm' ? 'l' : 'L'; continue; }
      if (!cur) novo();
      if (cmd === 'L') { x = n(); y = n(); } else if (cmd === 'l') { x += n(); y += n(); }
      else if (cmd === 'H') x = n(); else if (cmd === 'h') x += n();
      else if (cmd === 'V') y = n(); else if (cmd === 'v') y += n();
      else { k++; continue; }
      cur.pts.push([x, y]);
    }
    return out.filter(c => c.pts.length > 1);
  }

  return { aci, nomeCamada, texto1252, escrever, bytes, caminho, PALETA };
})();

if (typeof module !== 'undefined') module.exports = { WGImport, WGPluma, WGDxf, WGPot, WGSecao, WGRevisao, WG360, WGCerca, WGMalha };

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
  let mapa, camadaSecao, camadaF360, camadaPocos, camadaRotulos, camadaPluma, camadaPot, camadaPlanta, camadaAjuste, rendPlanta;
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
    pluma: { ativa: true, limiar: null, p: 2, cel: 2, calor: (() => { try { return localStorage.getItem('webgeo.pluma.calor') === '1'; } catch (e) { return false; } })() }, // limiar null = usa o VI do parâmetro
    calc: null,                   // resultado da última interpolação
    excluindo: null,              // {tipo: 'projeto'|'campanha', id, nome} enquanto o aviso de exclusão está aberto
    pot: { ativa: false, rede: null, intervalo: 0, suav: 0, setas: true }, // rede null = escolhe sozinho; intervalo 0 = automático
    potCalc: null,                // superfície potenciométrica da campanha selecionada
    secao: { tracando: null, lista: [], ativa: 0, linha: null, letra: 'A', faixa: 10, exag: 0, ligar: true, pluma: true }, // seções geológicas A–A', B–B'... (linha/letra = a seção aberta; exag 0 = automático)
    secaoCalc: null,
    anot: { projeto: null, texto: '', base: '', versao: null, autor: null, erro: null, sujo: false, salvando: false, timer: null, conflito: null }, // anotações do projeto
    f360: { lista: [], erro: null, mostrar: true, novo: null, marcar: null, urls: new Map() }, // fotos 360° do projeto              // sondagens que entram no corte
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
      aplicarPapel();
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
      cancelarEditPoco(); sairRotulosPrancha(false); if (st.editLim) { mapa.removeLayer(st.editLim.camada); st.editLim = null; $('#lim-barra').hidden = true; mapa.doubleClickZoom.enable(); }
      if (e.target.value === '__novo__') { criarProjeto(); return; }
      st.projetoId = e.target.value || null; lembrarProjeto(st.projetoId);
      st.campanhaId = null; st.parametroId = null;
      carregarProjeto();
    });
    $('#f-parametro').addEventListener('change', e => { st.parametroId = +e.target.value; st.pluma.limiar = null; render(); });
    $('#f-campanha').addEventListener('change', e => { st.campanhaId = +e.target.value; aplicarPapel(); render(); });
    $('#btn-csv').addEventListener('click', exportarCSV);
    ligarEventosF360();
    ligarEventosCampo();
    ligarEventosDesenho();
    ligarEventosAnotacao();
    ligarEventosCerca();

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
    $('#btn-excluir-projeto').addEventListener('click', () => abrirExcluir('projeto'));
    $('#btn-excluir-campanha').addEventListener('click', () => abrirExcluir('campanha'));
    $('#excluir-fechar').addEventListener('click', fecharExcluir);
    $('#excluir-cancelar').addEventListener('click', fecharExcluir);
    $('#modal-excluir').addEventListener('click', e => { if (e.target.id === 'modal-excluir') fecharExcluir(); });
    $('#excluir-confirma').addEventListener('input', () => { $('#excluir-confirmar').disabled = !confirmouExclusao(); });
    $('#excluir-confirma').addEventListener('keydown', e => { if (e.key === 'Enter' && confirmouExclusao()) excluirConfirmado(); });
    $('#excluir-confirmar').addEventListener('click', excluirConfirmado);
    document.addEventListener('keydown', e => {
      if (e.key !== 'Escape') return;
      if (!$('#modal-excluir').hidden) fecharExcluir();
      if (!$('#modal-relatorio').hidden) fecharRelatorio();
      if (!$('#modal-senha').hidden) fecharSenha();
    });
    // equipe
    $('#equipe-form').addEventListener('submit', adicionarPessoa);
    $('#equipe-gerar').addEventListener('click', () => { $('#equipe-senha').value = gerarSenha(); });
    $('#equipe-form').addEventListener('input', () => { $('#equipe-erro').hidden = true; });
    $('#btn-senha').addEventListener('click', abrirSenha);
    $('#senha-fechar').addEventListener('click', fecharSenha);
    $('#senha-cancelar').addEventListener('click', fecharSenha);
    $('#senha-form').addEventListener('submit', salvarSenha);
    $('#equipe-tabela').addEventListener('change', e => { if (e.target.classList.contains('equipe-papel')) mudarPapel(e.target.closest('tr').dataset.id, e.target.value); });
    $('#equipe-tabela').addEventListener('click', e => {
      const tr = e.target.closest('tr'), c = e.target.classList;
      if (c.contains('equipe-remover')) removerPessoa(tr.dataset.id);
      if (c.contains('equipe-editar')) editarLinhaEquipe(tr);
      if (c.contains('equipe-salvar')) salvarLinhaEquipe(tr);
      if (c.contains('equipe-cancelar')) { $('#equipe-erro').hidden = true; carregarEquipe(); }
    });
    $('#equipe-tabela').addEventListener('keydown', e => {
      const tr = e.target.closest('tr.editando'); if (!tr || e.target.tagName !== 'INPUT') return;
      if (e.key === 'Enter') { e.preventDefault(); salvarLinhaEquipe(tr); }
      if (e.key === 'Escape') { $('#equipe-erro').hidden = true; carregarEquipe(); }
    });
    // relatório
    $('#btn-relatorio').addEventListener('click', abrirRelatorio);
    $('#rel-fechar').addEventListener('click', fecharRelatorio);
    $('#rel-cancelar').addEventListener('click', fecharRelatorio);
    $('#rel-gerar').addEventListener('click', gerarRelatorio);
    window.addEventListener('afterprint', () => setTimeout(encerrarRelatorio, 100));
    $('#f-pluma').addEventListener('change', e => { st.pluma.ativa = e.target.checked; render(); });
    $('#f-calor').checked = st.pluma.calor;
    $('#f-calor').addEventListener('change', e => {
      st.pluma.calor = e.target.checked;
      try { localStorage.setItem('webgeo.pluma.calor', st.pluma.calor ? '1' : '0'); } catch (er) { /* sem armazenamento */ }
      if (st.pluma.calor && !st.pluma.ativa) { st.pluma.ativa = true; $('#f-pluma').checked = true; }
      render();
    });
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
    $('#btn-secao-dxf').addEventListener('click', baixarSecaoDxf);
    $('#secao-concluir').addEventListener('click', concluirSecao);
    $('#secao-desfazer').addEventListener('click', desfazerSecao);
    $('#secao-cancelar').addEventListener('click', cancelarSecao);
    $('#f-secao-faixa').addEventListener('change', e => { st.secao.faixa = +e.target.value; render(); });
    $('#f-secao-exag').addEventListener('change', e => { st.secao.exag = +e.target.value; render(); });
    $('#f-secao-ligar').addEventListener('change', e => { st.secao.ligar = e.target.checked; render(); });
    $('#f-secao-pluma').addEventListener('change', e => { st.secao.pluma = e.target.checked; render(); });
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
    $('#poco-salvar').addEventListener('click', salvarEditPoco);
    $('#btn-mapa-dxf').addEventListener('click', baixarMapaDxf);
    $('#btn-prancha').addEventListener('click', abrirPrancha);
    ['#pr-fechar', '#pr-cancelar'].forEach(id => $(id).addEventListener('click', () => { $('#modal-prancha').hidden = true; }));
    $('#pr-gerar').addEventListener('click', () => gerarPrancha(false));
    $('#pr-visualizar').addEventListener('click', () => gerarPrancha(true));
    $('#previa-voltar').addEventListener('click', () => fecharPrevia(true));
    $('#previa-fechar').addEventListener('click', () => fecharPrevia(false));
    $('#previa-baixar').addEventListener('click', baixarPrevia);
    $('#previa-imprimir').addEventListener('click', () => {
      if (!previaPrancha) return;
      // imprime o próprio PDF (escala real, A3): abre num quadro escondido e chama a impressão do navegador
      const f = document.createElement('iframe'); f.style.cssText = 'position:fixed;right:0;bottom:0;width:1px;height:1px;border:0;opacity:0';
      f.src = previaPrancha.url; document.body.appendChild(f);
      f.onload = () => { try { f.contentWindow.focus(); f.contentWindow.print(); } catch (e) { window.open(previaPrancha.url, '_blank'); } setTimeout(() => f.remove(), 60000); };
    });
    $('#previa-mais').addEventListener('click', () => zoomPrevia(1.4));
    $('#previa-menos').addEventListener('click', () => zoomPrevia(1 / 1.4));
    $('#previa-ajustar').addEventListener('click', () => zoomPrevia(0));
    window.addEventListener('resize', () => { if (!$('#modal-previa').hidden && previaPrancha && !previaPrancha.zoom) desenharPrevia(); });
    $('#pr-rotulos').addEventListener('click', () => ajustarRotulosPrancha(true));
    $('#btn-rotulos').addEventListener('click', () => ajustarRotulosPrancha(false));
    $('#btn-limite').addEventListener('click', editarLimite);
    $('#lim-salvar').addEventListener('click', () => sairEdicaoLimite(true));
    $('#lim-cancelar').addEventListener('click', () => sairEdicaoLimite(false));
    $('#lim-calculado').addEventListener('click', voltarLimiteCalculado);
    $('#lim-estado-voltar').addEventListener('click', voltarLimiteCalculado);
    $('#lim-novo').addEventListener('click', iniciarDesenhoLimite);
    $('#lim-fim').addEventListener('click', concluirDesenhoLimite);
    $('#lim-limpar').addEventListener('click', () => { const ed = st.editLim; if (!ed) return; ed.polis = []; desenharEdicaoLimite(); iniciarDesenhoLimite(); });
    $('#lim-traco').addEventListener('change', e => { const ed = st.editLim; if (!ed) return; ed.reta = e.target.value === 'reta'; desenharEdicaoLimite(); previaDesenhoLimite(); });
    document.addEventListener('keydown', e => { if (e.key === 'Escape' && st.editLim?.desenho) sairDesenhoLimite(); });
    $('#f-valores').checked = prefsMapa.valores;
    $('#f-valores').addEventListener('change', e => { prefsMapa.valores = e.target.checked; guardarPrefsMapa(); atualizarRotulos(); });
    $('#rot-concluir').addEventListener('click', () => sairRotulosPrancha(true, st.rotAjuste?.daPrancha));
    $('#rot-cancelar').addEventListener('click', () => sairRotulosPrancha(false, st.rotAjuste?.daPrancha));
    $('#rot-padrao').addEventListener('click', rotulosPadraoPrancha);
    $('#pr-conteudo').addEventListener('change', e => { $('#pr-titulo').value = tituloPadraoPrancha(e.target.value); $('#pr-opc-mapa').hidden = e.target.value.startsWith('secao'); });
    $('#pr-escala').addEventListener('change', e => { $('#pr-escala-n').hidden = e.target.value !== 'outra'; if (e.target.value === 'outra') $('#pr-escala-n').focus(); });
    $('#pr-logo').addEventListener('change', e => { const f = e.target.files[0]; if (f) escolherLogoPrancha(f); e.target.value = ''; });
    $('#pr-logo-tirar').addEventListener('click', () => { try { localStorage.removeItem(chaveLogo()); } catch (e) { /* nada */ } mostrarLogoPrancha(); });
    $('#poco-situacao').addEventListener('change', habilitarSalvarPoco);
    $('#poco-cancelar').addEventListener('click', cancelarEditPoco);
    ['#poco-lat', '#poco-lon'].forEach(id => $(id).addEventListener('change', () => digitouEditPoco('graus')));
    ['#poco-e', '#poco-n'].forEach(id => $(id).addEventListener('change', () => digitouEditPoco('utm')));
    ['#poco-lat', '#poco-lon', '#poco-e', '#poco-n'].forEach(id => $(id).addEventListener('keydown', e => { if (e.key === 'Enter') e.target.dispatchEvent(new Event('change')); if (e.key === 'Escape') cancelarEditPoco(); }));
    $('#arquivo-utm').addEventListener('change', e => { const f = e.target.files[0]; if (f) converterUtmArquivo(f); e.target.value = ''; });
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('sobre'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('sobre'));
    drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('sobre'); e.dataTransfer.files[0] && lerArquivo(e.dataTransfer.files[0]); });
    $('#btn-gravar').addEventListener('click', gravarImportacao);
    $('#btn-cancelar').addEventListener('click', limparImportacao);
  }

  function trocarView(v) {
    document.querySelectorAll('.rail-nav [data-view]').forEach(b => b.classList.toggle('is-active', b.dataset.view === v));
    document.querySelectorAll('.view').forEach(s => { s.hidden = s.id !== 'view-' + v; });
    if (v === 'equipe') carregarEquipe();
    if (v === 'anotacoes') { renderAnotacao(); setTimeout(() => $('#anot-texto').focus(), 0); }
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
        await carregarFotos360();
        await carregarAnotacao();
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
    aplicarPapel();
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
      await carregarFotos360();
      await carregarOrto();
      await carregarDesenhos();
      await carregarAnotacao();
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
  const ROTULO_SITUACAO = { ATIVO: 'Ativo', SECO: 'Seco', DANIFICADO: 'Danificado/destruído', DESATIVADO: 'Desativado' };
  const ROTULO_STATUS = { acima: 'Acima do VI', abaixo: 'Abaixo do VI', lq: 'Abaixo do LQ', sem: 'Sem resultado' };

  function textoValor(r) {
    if (!r) return '-';
    if (r.menor_que_lq) return r.lq !== null ? `< ${fmt(r.lq, 4)}` : '< LQ';
    return fmt(r.valor, 4);
  }

  function render(enquadrar = false) {
    if (!sb) return;
    if (st.profile) aplicarPapel();
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
  // ------------------------------------------------------------- mapa: satélite, malha UTM e rosa dos ventos
  const CHAVE_MAPA = 'webgeo.mapa';
  const prefsMapa = (() => {
    const p = { satelite: true, malha: false, valores: true, orto: true };
    try { Object.assign(p, JSON.parse(localStorage.getItem(CHAVE_MAPA) || '{}')); } catch (e) { /* sem armazenamento */ }
    return p;
  })();
  const guardarPrefsMapa = () => { try { localStorage.setItem(CHAVE_MAPA, JSON.stringify(prefsMapa)); } catch (e) { /* sem armazenamento */ } };
  let camadaSatelite, camadaMalha, malhaCalc = null, rotulosMalha = null, controleRosa = null;

  /** Rosa dos ventos (SVG): 4 pontas principais e 4 intermediárias; o norte em vermelho. */
  function svgRosa() {
    const ponta = (ang, comp, larg, c1, c2) => {
      const r = a => [Math.sin(a), -Math.cos(a)], [ux, uy] = r(ang), [vx, vy] = r(ang + Math.PI / 2);
      const P = (t, s) => `${(ux * t + vx * s).toFixed(2)},${(uy * t + vy * s).toFixed(2)}`;
      return `<polygon points="${P(comp, 0)} ${P(0, -larg)} 0,0" fill="${c1}"/><polygon points="${P(comp, 0)} 0,0 ${P(0, larg)}" fill="${c2}"/>`;
    };
    let g = '';
    [45, 135, 225, 315].forEach(a => { g += ponta(a * Math.PI / 180, 22, 5, '#56655D', '#C9D3CE'); });
    [90, 180, 270].forEach(a => { g += ponta(a * Math.PI / 180, 33, 7, '#172621', '#ffffff'); });
    g += ponta(0, 33, 7, '#B3402A', '#E2694F');
    const letra = (t, x, y, cor) => `<text x="${x}" y="${y}" text-anchor="middle" dominant-baseline="central" fill="${cor}">${t}</text>`;
    return `<svg viewBox="-56 -56 112 112" width="84" height="84" role="img" aria-label="Rosa dos ventos: norte para cima">
      <circle r="38" fill="rgba(255,255,255,.88)" stroke="#172621" stroke-width="1.2"/>
      <circle r="29" fill="none" stroke="#172621" stroke-opacity=".3"/>
      ${g}<circle r="2.6" fill="#172621"/>
      <g font-family="IBM Plex Sans, sans-serif" font-weight="700" font-size="12" stroke="#ffffff" stroke-width="3" paint-order="stroke">
        ${letra('N', 0, -47, '#B3402A')}${letra('S', 0, 47, '#172621')}${letra('L', 47, 0, '#172621')}${letra('O', -47, 0, '#172621')}
      </g></svg>`;
  }

  function criarExtrasMapa(urlSatelite, opcoesSatelite) {
    camadaSatelite = L.tileLayer(urlSatelite, opcoesSatelite);
    if (prefsMapa.satelite) camadaSatelite.addTo(mapa);
    mapa.getContainer().classList.toggle('mapa-sem-satelite', !prefsMapa.satelite);
    mapa.createPane('malha').style.zIndex = 340;      // acima da imagem, abaixo da pluma e da planta
    camadaMalha = L.layerGroup().addTo(mapa);
    rotulosMalha = document.createElement('div');
    rotulosMalha.className = 'malha-rotulos';
    // posição e camada definidas aqui mesmo (não dependem do style.css estar atualizado)
    rotulosMalha.style.cssText = 'position:absolute;left:0;top:0;right:0;bottom:0;z-index:450;pointer-events:none;overflow:hidden;';
    mapa.getContainer().appendChild(rotulosMalha);

    const Rosa = L.Control.extend({ options: { position: 'topright' }, onAdd() { const d = L.DomUtil.create('div', 'rosa-ventos'); d.innerHTML = svgRosa(); return d; } });
    controleRosa = new Rosa();
    controleRosa.addTo(mapa);                          // a rosa dos ventos fica sempre no mapa

    const Botoes = L.Control.extend({
      options: { position: 'topleft' },
      onAdd() {
        const d = L.DomUtil.create('div', 'leaflet-bar mapa-botoes');
        d.style.cssText = 'display:flex;flex-direction:column;';   // um embaixo do outro
        d.innerHTML = [['satelite', 'Satélite', 'Ligar ou desligar a imagem de satélite'], ['malha', 'Malha UTM', 'Mostrar a malha de coordenadas UTM'], ['orto', 'Ortofoto', 'Ligar ou desligar a ortofoto do drone']]
          .map(([k, t, dica]) => `<button type="button" data-k="${k}" title="${dica}" aria-pressed="${!!prefsMapa[k]}"${k === 'orto' && !st.orto ? ' hidden' : ''}>${t}</button>`).join('');
        L.DomEvent.disableClickPropagation(d); L.DomEvent.disableScrollPropagation(d);
        d.addEventListener('click', e => { const b = e.target.closest('button[data-k]'); if (b) alternarMapa(b.dataset.k); });
        return d;
      }
    });
    new Botoes().addTo(mapa);
    mapa.on('moveend zoomend resize', renderMalha);
    mapa.on('move zoom', posicionarRotulosMalha);
    renderMalha();
  }

  function alternarMapa(k) {
    prefsMapa[k] = !prefsMapa[k];
    guardarPrefsMapa();
    document.querySelectorAll(`.mapa-botoes button[data-k="${k}"]`).forEach(b => b.setAttribute('aria-pressed', String(prefsMapa[k])));
    if (k === 'orto') { renderOrto(); return; }
    if (k === 'satelite') {
      if (prefsMapa.satelite) camadaSatelite.addTo(mapa); else mapa.removeLayer(camadaSatelite);
      mapa.getContainer().classList.toggle('mapa-sem-satelite', !prefsMapa.satelite);
    }
    renderMalha();                                     // a cor da malha muda com o satélite
  }

  function renderMalha() {
    if (!camadaMalha) return;
    camadaMalha.clearLayers(); malhaCalc = null;
    if (!prefsMapa.malha || !mapaVisivel()) { rotulosMalha.innerHTML = ''; return; }
    const b = mapa.getBounds();
    malhaCalc = WGMalha.linhas({ s: b.getSouth(), w: b.getWest(), n: b.getNorth(), e: b.getEast() }, Math.max(3, Math.round(mapa.getSize().x / 170)));
    // pontilhada: branca sobre o satélite, preta sem o satélite
    const cor = prefsMapa.satelite ? '#ffffff' : '#000000';
    [...malhaCalc.verticais, ...malhaCalc.horizontais].forEach(l =>
      L.polyline(l.pts, { pane: 'malha', className: 'malha-linha', color: cor, weight: 1.3, opacity: prefsMapa.satelite ? 0.9 : 0.75,
        dashArray: '2 5', lineCap: 'round', interactive: false }).addTo(camadaMalha));
    posicionarRotulosMalha();
  }

  /** Valores de E no alto e de N na esquerda do mapa, onde cada linha encosta na borda (como na carta impressa). */
  function posicionarRotulosMalha() {
    if (!rotulosMalha) return;
    if (!malhaCalc) { rotulosMalha.innerHTML = ''; return; }
    const { x: W, y: H } = mapa.getSize(), M = 4;
    const px = pts => pts.map(p => mapa.latLngToContainerPoint(p));
    // não escreve por baixo dos botões, do zoom e da rosa dos ventos
    const caixa = mapa.getContainer().getBoundingClientRect();
    const tampas = [...mapa.getContainer().querySelectorAll('.leaflet-top .leaflet-control')].map(el => {
      const r = el.getBoundingClientRect(); return { x0: r.left - caixa.left - 6, x1: r.right - caixa.left + 6, y0: r.top - caixa.top - 6, y1: r.bottom - caixa.top + 6 };
    });
    const livre = (x0, x1, y0, y1) => x0 > 2 && x1 < W - 2 && y0 > 2 && y1 < H - 26 && !tampas.some(t => x0 < t.x1 && x1 > t.x0 && y0 < t.y1 && y1 > t.y0);
    let html = '';
    const est = 'position:absolute;white-space:nowrap;font:600 10.5px/1 \'IBM Plex Mono\',ui-monospace,monospace;' + (prefsMapa.satelite
      ? 'color:#fff;text-shadow:0 0 3px #000,0 0 2px #000,0 0 1px #000;' : 'color:#000;text-shadow:0 0 3px #fff,0 0 2px #fff,0 0 1px #fff;');
    malhaCalc.verticais.forEach(l => {
      // E na vertical, encostado à esquerda da linha, lendo de baixo para cima
      const c = WGMalha.cruzamento(px(l.pts), 'y', M), t = 'E ' + WGMalha.texto(l.v, malhaCalc.passo), w = t.length * 6.4;
      if (c && livre(c.x - 15, c.x, 4, 8 + w)) html += `<span class="malha-e" style="${est}left:${(c.x - 14).toFixed(1)}px;top:6px;writing-mode:vertical-rl;transform:rotate(180deg);">${t}</span>`;
    });
    malhaCalc.horizontais.forEach(l => {
      // N em cima da linha, apoiado nela
      const c = WGMalha.cruzamento(px(l.pts), 'x', M), t = 'N ' + WGMalha.texto(l.v, malhaCalc.passo), w = t.length * 6.4;
      if (!c) return;
      // se os botões do mapa cobrem a borda nessa altura, escreve logo depois deles
      const xs = [6, ...tampas.filter(tp => c.y - 15 < tp.y1 && c.y > tp.y0 && tp.x0 < 60).map(tp => tp.x1 + 2)];
      const x = xs.find(x0 => livre(x0 - 2, x0 + w + 2, c.y - 15, c.y));
      if (x != null) html += `<span class="malha-n" style="${est}top:${(c.y - 13).toFixed(1)}px;left:${x.toFixed(1)}px;">${t}</span>`;
    });
    html += `<span class="malha-info" style="${est}left:8px;bottom:8px;font-family:'IBM Plex Sans',sans-serif;font-weight:500;">UTM fuso ${malhaCalc.zona}${malhaCalc.sul ? 'S' : 'N'} · WGS 84 / SIRGAS 2000 · malha de ${WGMalha.texto(malhaCalc.passo, malhaCalc.passo)} m</span>`;
    rotulosMalha.innerHTML = html;
  }

  function criarMapa() {
    mapa = L.map('mapa').setView([-23.471, -46.576], 17);
    L.control.scale({ imperial: false, position: 'bottomright' }).addTo(mapa);
    criarExtrasMapa('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
      maxZoom: 21, maxNativeZoom: 19, attribution: 'Imagem © Esri'
    });
    // a pluma fica por baixo dos poços: imagem (350) < linha do limiar (380) < poços (400)
    mapa.createPane('orto').style.zIndex = 250;          // ortofoto do drone: acima do satélite, abaixo de tudo o mais
    mapa.createPane('pluma').style.zIndex = 350;
    mapa.createPane('planta').style.zIndex = 360;      // planta DXF: acima da pluma, abaixo da linha do limiar
    mapa.createPane('plumaLinha').style.zIndex = 380;
    mapa.createPane('pot').style.zIndex = 385;         // curvas equipotenciais e setas de fluxo
    mapa.createPane('secao').style.zIndex = 390;       // linha da seção geológica
    rendPlanta = L.canvas({ pane: 'planta', padding: .5 }); // canvas: aguenta milhares de linhas sem pesar
    camadaPlanta = L.layerGroup().addTo(mapa);
    camadaAjuste = L.layerGroup().addTo(mapa);
    mapa.on('click', cliqueAjuste);
    mapa.on('click', cliqueF360);
    mapa.on('popupopen', () => { if (st.ajuste || st.secao.tracando || st.f360.marcar || st.editPoco || st.rotAjuste || st.editLim || st.desAtivo) setTimeout(() => mapa.closePopup(), 0); });
    mapa.on('click', cliqueEditPoco);
    ligarMapaDesenho();
    mapa.on('click', cliqueLimite);
    mapa.on('dblclick', () => { if (st.editLim?.desenho) concluirDesenhoLimite(); });
    mapa.getContainer().addEventListener('click', e => { const b = e.target.closest('[data-editar-poco]'); if (b) { e.preventDefault(); editarPosicaoPoco(+b.dataset.editarPoco || b.dataset.editarPoco); } }); // ajustando a planta ou traçando a seção: sem popups
    mapa.on('click', cliqueSecao);
    mapa.on('dblclick', () => { if (st.secao.tracando) concluirSecao(); });
    camadaPluma = L.layerGroup().addTo(mapa);
    camadaPot = L.layerGroup().addTo(mapa);
    camadaSecao = L.layerGroup().addTo(mapa);
    camadaPocos = L.layerGroup().addTo(mapa);
    camadaF360 = L.layerGroup().addTo(mapa);
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
    if (!mapaVisivel() || st.rotAjuste) return;          // ajustando os rótulos: só as etiquetas arrastáveis aparecem
    const perto = mapa.getZoom() >= 19;
    const todos = rotulosDoMapa.map(r => ({ ...r, pt: mapa.latLngToLayerPoint(r.ll) }));
    const pos = lerLS(chaveRotulos(), {}), manual = r => r.pid != null && pos[r.pid];
    const mostra = r => perto || r.status === 'acima';
    // nome arrastado pela pessoa: fica onde ela pôs, com linha de chamada; o resto entra na arrumação automática
    const itens = todos.filter(r => mostra(r) && !manual(r)?.nome);
    const chamadaMapa = (de, para) => L.polyline([de, para], { color: '#ffffff', weight: 1, opacity: .85, interactive: false }).addTo(camadaRotulos);
    todos.filter(r => mostra(r) && manual(r)?.nome).forEach(r => {
      chamadaMapa(r.ll, manual(r).nome);
      L.marker(manual(r).nome, { icon: L.divIcon({ className: 'rotulos-ancora', html: `<div class="rotulo-pm rotulo-${r.status} rotulo-solto">${r.html}</div>`, iconSize: [0, 0] }), interactive: false, keyboard: false }).addTo(camadaRotulos);
    });
    // valores dos resultados: embaixo de cada poço (ou onde a pessoa arrastou)
    const caixasValor = [];
    // carga hidráulica (mapa potenciométrico): em cima de cada poço, numa etiqueta azul (ou onde a pessoa arrastou)
    todos.filter(r => r.carga).forEach(r => {
      const m = manual(r)?.carga, w = r.carga.length * 6.2 + 8;
      if (m) chamadaMapa(r.ll, m);
      else caixasValor.push({ x0: r.pt.x - w / 2, y0: r.pt.y - 23, x1: r.pt.x + w / 2, y1: r.pt.y - 8 });
      L.marker(m || r.ll, { icon: L.divIcon({ className: 'rotulos-ancora', iconSize: [0, 0],
        html: `<div class="rotulo-carga-mapa${m ? ' solto' : ''}">${esc(r.carga)}</div>` }), interactive: false, keyboard: false }).addTo(camadaRotulos);
    });
    if (prefsMapa.valores) todos.filter(r => r.valor && mostra(r)).forEach(r => {
      const m = manual(r)?.valor, w = r.valor.length * 6.2 + 8;
      if (m) chamadaMapa(r.ll, m);
      else caixasValor.push({ x0: r.pt.x - w / 2, y0: r.pt.y + 7, x1: r.pt.x + w / 2, y1: r.pt.y + 21 });
      L.marker(m || r.ll, { icon: L.divIcon({ className: 'rotulos-ancora', iconSize: [0, 0],
        html: `<div class="rotulo-valor-mapa${m ? ' solto' : ''} rotulo-${r.status}">${esc(r.valor)}</div>` }), interactive: false, keyboard: false }).addTo(camadaRotulos);
    });
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
    const ocupado = todos.map(o => ({ x0: o.pt.x - 7, y0: o.pt.y - 7, x1: o.pt.x + 7, y1: o.pt.y + 7 })).concat(caixasValor); // os pontos dos poços e os valores
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
        const inativo = p.situacao && p.situacao !== 'ATIVO';   // seco, danificado, destruído, desativado: borda preta tracejada e cor apagada
        const m = L.circleMarker(ll, { radius: s === 'acima' ? 8 : 6, weight: 2, color: inativo ? '#172621' : '#fff', dashArray: inativo ? '2 2' : null,
          fillColor: COR_STATUS[s], fillOpacity: inativo ? .45 : 1, className: 'pm pm-' + s + (inativo ? ' pm-inativo' : '') })
          .bindPopup(() => popupPoco(p, param), { maxWidth: 320 })
          .on('click', cliqueSecao)   // o clique num poço não chega ao mapa (o popup segura): avisa a seção daqui
          .on('popupopen', () => { st.popupAberto = p.id; })
          .on('popupclose', () => { if (st.popupAberto === p.id && !st.redesenhando) st.popupAberto = null; })
          .addTo(camadaPocos);
        marcadores.set(p.id, m);
        const h = carga.has(p.id) ? carga.get(p.id).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '';
        const sit = inativo ? (ROTULO_SITUACAO[p.situacao] || p.situacao).toLowerCase() : '';
        rotulosDoMapa.push({ pid: p.id, valor: r ? valorRotulo(r) : null, carga: h || null, ll, status: s, texto: p.codigo, largura: p.codigo.length + (sit ? sit.length + 1 : 0),
          html: esc(p.codigo) + (sit ? ` <span class="rotulo-sit">${esc(sit)}</span>` : '') });
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

  // ------------------------------------------------------------- editar um poço: posição (arrastar, clicar ou digitar) e situação
  const SITUACOES_POCO = ['ATIVO', 'SECO', 'DANIFICADO', 'DESATIVADO'];   // os mesmos que o banco já aceita
  /** Abre o modo de edição: pino arrastável no mapa + campos de latitude/longitude e UTM E/N. Grava em wg_poco ao salvar. */
  function editarPosicaoPoco(id) {
    const p = st.pocos.find(x => String(x.id) === String(id));
    if (!p || p.virtual) return;
    if (st.ajuste || st.secao.tracando || st.f360.marcar) { toast('Termine o que está fazendo no mapa antes de editar a posição.'); return; }
    cancelarEditPoco();
    mapa.closePopup();
    const tem = p.latitude != null && p.longitude != null;
    const c = mapa.getCenter(), ll0 = tem ? [+p.latitude, +p.longitude] : [c.lat, c.lng];
    const zona = WGDxf.zonaDaLongitude(ll0[1]), sul = ll0[0] < 0;
    const mk = L.marker(ll0, { draggable: true, zIndexOffset: 2000, keyboard: false,
      icon: L.divIcon({ className: '', html: `<div class="pino-poco-edit"><span>${esc(p.codigo)}</span></div>`, iconSize: [18, 18], iconAnchor: [9, 9] }) }).addTo(mapa);
    const linha = L.polyline([ll0, ll0], { color: '#FFD60A', weight: 2, dashArray: '4 4', interactive: false }).addTo(mapa);
    st.editPoco = { p, mk, linha, ll0, tem, zona, sul, moveu: false };
    mk.on('drag', () => { st.editPoco.moveu = true; posicaoEditPoco(mk.getLatLng(), 'arrastar'); });
    $('#poco-situacao').value = SITUACOES_POCO.includes(p.situacao) ? p.situacao : 'ATIVO';
    mapa.getContainer().classList.add('mapa-ajustando');
    $('#poco-passo').textContent = `Editar poço ${p.codigo}`;
    $('#poco-fuso').textContent = `UTM fuso ${zona}${sul ? 'S' : 'N'} (SIRGAS 2000 / WGS 84)`;
    $('#poco-barra').hidden = false;
    $('#poco-barra').scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    posicaoEditPoco({ lat: ll0[0], lng: ll0[1] }, 'inicio');
    if (tem) mapa.panTo(ll0);
  }

  /** Atualiza pino, linha até a posição antiga e os 4 campos (menos o que a pessoa está digitando). */
  function posicaoEditPoco(ll, origem) {
    const ed = st.editPoco; if (!ed) return;
    ed.atual = [ll.lat, ll.lng];
    if (origem !== 'arrastar') ed.mk.setLatLng(ll);
    ed.linha.setLatLngs([ed.ll0, ed.atual]);
    const u = WGDxf.latLonParaUtm(ll.lat, ll.lng, ed.zona, ed.sul);
    const f = (v, d) => (+v).toFixed(d).replace('.', ',');
    if (origem !== 'graus') { $('#poco-lat').value = f(ll.lat, 8); $('#poco-lon').value = f(ll.lng, 8); }
    if (origem !== 'utm') { $('#poco-e').value = f(u.E, 3); $('#poco-n').value = f(u.N, 3); }
    const d = ed.tem ? mapa.distance(ed.ll0, ed.atual) : null;
    $('#poco-info').textContent = !ed.tem ? (ed.moveu ? 'poço sem coordenada: posição nova' : 'poço sem coordenada') : d < 0.05 ? 'posição sem mudança' : `movido ${fmt(d, d < 10 ? 2 : 1)} m`;
    habilitarSalvarPoco();
  }

  /** Salvar só fica ativo se mudou a posição ou a situação. */
  function habilitarSalvarPoco() {
    const ed = st.editPoco; if (!ed) return;
    const moveu = ed.tem ? mapa.distance(ed.ll0, ed.atual) >= 0.05 : ed.moveu;
    const sit = $('#poco-situacao').value !== (ed.p.situacao || 'ATIVO');
    $('#poco-salvar').disabled = !(moveu || sit);
  }

  function digitouEditPoco(tipo) {
    const ed = st.editPoco; if (!ed) return;
    const n = id => WGImport.numero($(id).value);
    if (tipo === 'graus') {
      const la = n('#poco-lat'), lo = n('#poco-lon');
      if (la === null || lo === null || Math.abs(la) > 90 || Math.abs(lo) > 180) { $('#poco-info').textContent = 'latitude/longitude inválida (graus decimais, ex.: -23,471383)'; $('#poco-salvar').disabled = true; return; }
      ed.moveu = true; posicaoEditPoco({ lat: la, lng: lo }, 'graus');
    } else {
      const E = n('#poco-e'), N = n('#poco-n');
      if (E === null || N === null || E < 100000 || E > 900000 || N < 0 || N > 10000000) { $('#poco-info').textContent = 'E/N inválido (UTM em metros, ex.: E 338979 / N 7403370)'; $('#poco-salvar').disabled = true; return; }
      const r = WGDxf.utmParaLatLon(E, N, ed.zona, ed.sul);
      ed.moveu = true; posicaoEditPoco({ lat: r.lat, lng: r.lon }, 'utm');
    }
    mapa.panTo(ed.atual);
  }

  function cliqueEditPoco(e) { if (st.editPoco) { st.editPoco.moveu = true; posicaoEditPoco(e.latlng, 'clique'); } }

  function cancelarEditPoco() {
    const ed = st.editPoco; if (!ed) return;
    mapa.removeLayer(ed.mk); mapa.removeLayer(ed.linha);
    st.editPoco = null;
    mapa.getContainer().classList.remove('mapa-ajustando');
    $('#poco-barra').hidden = true;
  }

  async function salvarEditPoco() {
    const ed = st.editPoco; if (!ed || !ed.atual) return;
    const [lat, lon] = ed.atual.map(v => Math.round(v * 1e8) / 1e8);
    const moveu = ed.tem ? mapa.distance(ed.ll0, ed.atual) >= 0.05 : ed.moveu;
    const situacao = $('#poco-situacao').value, mudouSit = situacao !== (ed.p.situacao || 'ATIVO');
    const dados = {};
    if (moveu) Object.assign(dados, { latitude: lat, longitude: lon });
    if (mudouSit) dados.situacao = situacao;
    if (!Object.keys(dados).length) return;
    const btn = $('#poco-salvar'); btn.disabled = true; btn.textContent = 'Salvando...';
    const { error } = await sb.from('wg_poco').update(dados).eq('id', ed.p.id);
    btn.textContent = 'Salvar';
    if (error) {
      btn.disabled = false;
      toast('Não consegui salvar: ' + error.message);
      return;
    }
    Object.assign(ed.p, dados);
    const cod = ed.p.codigo;
    cancelarEditPoco();
    render();
    toast(`${cod} salvo${mudouSit ? ` como ${ROTULO_SITUACAO[situacao].toLowerCase()}` : ''}${moveu ? (mudouSit ? ', com a posição nova' : ': posição nova') : ''}.`);
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
        <div class="pop-mini">${p.cota_topo != null ? `Cota topo ${fmt(p.cota_topo, 3)} m · ` : ''}${esc(ROTULO_SITUACAO[p.situacao] || p.situacao)}${p.coordFonte ? ` · coordenada: ${esc(p.coordFonte)}` : ''}</div>
        ${p.virtual ? '<div class="pop-mini">Poço vindo da ficha do Perfil (ainda sem resultados de laboratório).</div>' : ''}
        ${fichas.length ? `<div class="pop-sond">Ficha no Perfil: ${fichas.map(f => `<b>${esc(f.sondagem_no || 'sondagem')}</b>${f.obra ? ' · ' + esc(f.obra) : ''}`).join('<br>')}</div>` : ''}
        ${perfilSVG(p, med)}
        ${p.virtual ? '' : `<div class="pop-acoes"><button type="button" class="btn-mini" data-ficha-campo="${esc(String(p.id))}">Ficha de campo</button><button type="button" class="btn-mini" data-editar-poco="${esc(String(p.id))}">Editar poço</button></div>`}
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
        item.amostras = pontos;
        if (limiar != null) Object.assign(item, WGPluma.areaECentroide(g, item.v, limiar));
        const man = limiteManual(c.id, param.id);
        if (man) { Object.assign(item, WGPluma.areaPoligonos(g, polisSuaves(man.polis).map(q => q.pts))); item.manual = true; }
      }
      return item;
    });
    // a pluma é pintada só onde o valor interpolado passa do limiar (sem limiar: 3 ordens abaixo do máximo)
    let vmax = Math.pow(10, Math.max(1, Math.ceil(Math.log10(maxGlobal || 10))));
    const vmin = limiar != null ? limiar : vmax / 1e3;
    if (vmax <= vmin * 1.5) vmax = vmin * 10;
    return { g, limiar, porCamp, vmax, vmin, piso };
  }

  function plumaDaCampanha(id) { return st.calc?.porCamp.find(x => x.campanha.id === id) || null; }
  function primeiraComArea() { return st.calc?.porCamp.find(x => x.v) || null; }

  // ------------------------------------------------------------- linha do limite ajustada à mão
  const chaveLimite = (campId = st.campanhaId, paramId = st.parametroId) => `webgeo.limite.${st.projetoId}.${paramId}.${campId}`;
  /** Linha ajustada à mão para o projeto/parâmetro/campanha: { polis: [{ pts: [[lat, lon]...], def: [bool por trecho] }] } ou null. */
  function limiteManual(campId = st.campanhaId, paramId = st.parametroId) {
    try { const v = localStorage.getItem(chaveLimite(campId, paramId)); const o = v ? JSON.parse(v) : null; return o && o.polis && o.polis.length ? o : null; } catch (e) { return null; }
  }
  /** Curva suave (Catmull-Rom, passa pelos pontos) de um polígono fechado: pedaço k = do ponto k ao k+1. */
  function curvaTrecho(pts, k, n = 8) {
    const N = pts.length, P = i => pts[(i + N) % N], p0 = P(k - 1), p1 = P(k), p2 = P(k + 1), p3 = P(k + 2), out = [];
    for (let t = 0; t <= n; t++) {
      const u = t / n, u2 = u * u, u3 = u2 * u;
      out.push([0, 1].map(c => .5 * (2 * p1[c] + (-p0[c] + p2[c]) * u + (2 * p0[c] - 5 * p1[c] + 4 * p2[c] - p3[c]) * u2 + (-p0[c] + 3 * p1[c] - 3 * p2[c] + p3[c]) * u3)));
    }
    return out;
  }
  /** Pedaço k da linha: curva (spline) ou reta, conforme o traço escolhido. */
  const trechoDe = (pts, k, reta) => reta ? [pts[k], pts[(k + 1) % pts.length]] : curvaTrecho(pts, k);
  /** Polígonos da linha ajustada, já com a curva (usados no recorte da pluma, na área e no DXF). */
  const polisSuaves = (polis, reta = limiteManual()?.reta) => polis.map(({ pts }) => ({ pts: pts.flatMap((_, k) => trechoDe(pts, k, reta).slice(0, -1)) }));

  function dentroLimiteManual(lat, lon, polis) {
    let dentro = false;
    polis.forEach(({ pts }) => { for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
      const [ya, xa] = pts[i], [yb, xb] = pts[j];
      if ((ya > lat) !== (yb > lat) && lon < (xb - xa) * (lat - ya) / (yb - ya) + xa) dentro = !dentro;
    } });
    return dentro;
  }
  /** Trechos da linha do limite: os da linha ajustada à mão (se houver) ou os calculados. [{pts, definido}] */
  function linhasDoLimite(calc, sel) {
    const m = limiteManual();
    if (!m) return WGPluma.trechosLimite(calc.g, WGPluma.isolinha(calc.g, sel.v, calc.limiar), sel.amostras, calc.limiar);
    const out = [];
    m.polis.forEach(({ pts, def }) => {
      const n = pts.length; let atual = null;
      for (let k = 0; k < n; k++) {
        const d = !!def[k], cv = trechoDe(pts, k, m.reta);
        if (!atual || atual.definido !== d) { atual = { pts: [cv[0]], definido: d }; out.push(atual); }
        atual.pts.push(...cv.slice(1));
      }
    });
    return out;
  }
  /** Recorta a imagem da pluma pela linha ajustada à mão (k = pixels por célula). */
  function mascararPluma(cv, g, k, polis) {
    const m = document.createElement('canvas'); m.width = cv.width; m.height = cv.height;
    const c = m.getContext('2d'), alto = g.ny * g.cel;
    c.beginPath();
    polis.forEach(({ pts }) => pts.forEach(([la, lo], i) => {
      const q = g.proj.paraXY(la, lo), x = (q.x - g.xmin) / g.cel * k, y = (alto - (q.y - g.ymin)) / g.cel * k;
      if (i) c.lineTo(x, y); else c.moveTo(x, y);
    }));
    c.closePath(); c.fill('evenodd');
    c.globalCompositeOperation = 'source-in'; c.drawImage(cv, 0, 0);
    return m;
  }

  /** Modo de edição: pontos arrastáveis na linha; clique no trecho alterna contínua/tracejada. */
  function editarLimite() {
    const calc = st.calc, sel = plumaDaCampanha(st.campanhaId);
    if (!st.pluma.ativa || !calc || !sel || !sel.v || calc.limiar == null) { toast('Ligue a pluma (com limite definido) para editar a linha.'); return; }
    if (st.ajuste || st.secao.tracando || st.f360.marcar || st.editPoco || st.rotAjuste || st.desAtivo) { toast('Termine o que está fazendo no mapa antes.'); return; }
    sairEdicaoLimite(false);
    const m = limiteManual();
    let polis;
    if (m) polis = m.polis.map(q => ({ pts: q.pts.map(p => p.slice()), def: q.def.slice() }));
    else polis = WGPluma.isolinha(calc.g, sel.v, calc.limiar).filter(c => c.length > 3).map(c => {
      let pts = WGPluma.simplificar(calc.g, c, Math.max(.6, calc.g.cel * .6));
      if (pts.length > 3 && Math.hypot(pts[0][0] - pts[pts.length - 1][0], pts[0][1] - pts[pts.length - 1][1]) < 1e-7) pts = pts.slice(0, -1);
      const def = WGPluma.classificarSegmentos(calc.g, pts.concat([pts[0]]), sel.amostras, calc.limiar);
      return { pts, def };
    }).filter(q => q.pts.length >= 3);
    st.editLim = { polis, camada: L.layerGroup().addTo(mapa), chave: chaveLimite(), reta: !!m?.reta, desenho: null };
    $('#lim-traco').value = st.editLim.reta ? 'reta' : 'spline';
    mapa.doubleClickZoom.disable();
    mapa.getContainer().classList.add('mapa-editando-limite');
    $('#lim-barra').hidden = false;
    $('#lim-barra').scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    desenharEdicaoLimite();
    if (!polis.length) iniciarDesenhoLimite();      // sem linha calculada: já começa desenhando
  }

  // ---- desenhar uma pluma nova (clique a clique)
  function iniciarDesenhoLimite() {
    const ed = st.editLim; if (!ed) return;
    ed.desenho = { pts: [], camada: L.layerGroup().addTo(mapa) };
    mapa.getContainer().classList.add('mapa-desenhando-limite');
    $('#lim-dica').textContent = 'Clique no mapa para marcar os pontos da pluma nova, em volta. Duplo clique (ou "Concluir desenho") fecha a pluma; Esc desiste.';
    $('#lim-fim').hidden = false; $('#lim-novo').hidden = true;
  }
  function previaDesenhoLimite() {
    const ed = st.editLim, dz = ed?.desenho; if (!dz) return;
    dz.camada.clearLayers();
    const pts = dz.pts;
    if (pts.length >= 2) {
      const linha = pts.length >= 3 ? pts.flatMap((_, k) => trechoDe(pts, k, ed.reta).slice(0, -1)).concat([pts[0]]) : pts;
      L.polyline(linha, { color: '#ffffff', weight: 6, opacity: .7, interactive: false }).addTo(dz.camada);
      L.polyline(linha, { color: '#B3402A', weight: 2.5, dashArray: pts.length >= 3 ? null : '4 4', interactive: false }).addTo(dz.camada);
    }
    pts.forEach(p => L.marker(p, { interactive: false, keyboard: false, icon: L.divIcon({ className: '', html: '<div class="lim-vertice novo"></div>', iconSize: [12, 12], iconAnchor: [6, 6] }) }).addTo(dz.camada));
  }
  function concluirDesenhoLimite() {
    const ed = st.editLim, dz = ed?.desenho; if (!dz) return;
    // duplo clique gera dois cliques antes: tira pontos repetidos no fim
    const pts = dz.pts.filter((p, i, a) => i === 0 || Math.hypot(p[0] - a[i - 1][0], p[1] - a[i - 1][1]) > 1e-7);
    if (pts.length < 3) { toast('Marque pelo menos 3 pontos para fechar a pluma.'); return; }
    ed.polis.push({ pts, def: pts.map(() => true) });
    sairDesenhoLimite();
    desenharEdicaoLimite();
  }
  function sairDesenhoLimite() {
    const ed = st.editLim; if (!ed || !ed.desenho) return;
    mapa.removeLayer(ed.desenho.camada); ed.desenho = null;
    mapa.getContainer().classList.remove('mapa-desenhando-limite');
    $('#lim-dica').textContent = TEXTO_DICA_LIMITE;
    $('#lim-fim').hidden = true; $('#lim-novo').hidden = false;
  }
  const TEXTO_DICA_LIMITE = 'Arraste os pontos brancos. Para criar um ponto, clique ou arraste a bolinha do meio de um trecho; para apagar, duplo clique (ou botão direito) no ponto. Clique num trecho para trocar entre contínuo e tracejado.';
  function cliqueLimite(e) {
    const dz = st.editLim?.desenho; if (!dz) return;
    dz.pts.push([e.latlng.lat, e.latlng.lng]); previaDesenhoLimite();
  }

  function desenharEdicaoLimite() {
    const ed = st.editLim; if (!ed) return;
    ed.camada.clearLayers();
    ed.polis.forEach((pl, pi) => {
      const n = pl.pts.length;
      for (let k = 0; k < n; k++) {
        const cv = trechoDe(pl.pts, k, ed.reta), mm = ed.reta ? [(cv[0][0] + cv[1][0]) / 2, (cv[0][1] + cv[1][1]) / 2] : cv[4];
        L.polyline(cv, { color: '#ffffff', weight: 7, opacity: .7 }).addTo(ed.camada);
        const seg = L.polyline(cv, { color: '#172621', weight: 3, dashArray: pl.def[k] ? null : '7 6' }).addTo(ed.camada);
        seg.on('click', e => { L.DomEvent.stop(e); pl.def[k] = !pl.def[k]; desenharEdicaoLimite(); });
        // ponto do meio: arrastar ou clicar cria um ponto novo
        const meio = L.marker(mm, { draggable: true, keyboard: false, zIndexOffset: 1500,
          icon: L.divIcon({ className: '', html: '<div class="lim-meio"></div>', iconSize: [12, 12], iconAnchor: [6, 6] }) }).addTo(ed.camada);
        const inserir = ll => { pl.pts.splice(k + 1, 0, [ll.lat, ll.lng]); pl.def.splice(k + 1, 0, pl.def[k]); desenharEdicaoLimite(); };
        meio.on('click', e => { L.DomEvent.stop(e); inserir(meio.getLatLng()); });
        meio.on('dragend', () => inserir(meio.getLatLng()));
      }
      pl.pts.forEach((pt, k) => {
        const v = L.marker(pt, { draggable: true, keyboard: false, zIndexOffset: 2000,
          icon: L.divIcon({ className: '', html: '<div class="lim-vertice"></div>', iconSize: [14, 14], iconAnchor: [7, 7] }) }).addTo(ed.camada);
        v.on('drag', () => { const q = v.getLatLng(); pl.pts[k] = [q.lat, q.lng]; });
        v.on('dragend', desenharEdicaoLimite);
        v.on('dblclick contextmenu', e => {
          L.DomEvent.stop(e);
          if (pl.pts.length <= 3) { if (ed.polis.length > 1) { ed.polis.splice(pi, 1); desenharEdicaoLimite(); } else toast('A linha precisa de pelo menos 3 pontos.'); return; }
          pl.pts.splice(k, 1); pl.def.splice(k, 1); desenharEdicaoLimite();
        });
      });
    });
  }

  function sairEdicaoLimite(salvar) {
    const ed = st.editLim; if (!ed) return;
    if (salvar && ed.desenho && ed.desenho.pts.length >= 3) concluirDesenhoLimite();
    if (salvar && !ed.polis.length) { toast('Não há nenhuma pluma desenhada para salvar. Desenhe uma ou use "Voltar à calculada".'); return; }
    if (salvar) { try { localStorage.setItem(chaveLimite(), JSON.stringify({ polis: ed.polis, reta: ed.reta, em: new Date().toISOString() })); } catch (e) { toast('O navegador não deixou guardar a linha.'); } }
    sairDesenhoLimite();
    mapa.removeLayer(ed.camada);
    st.editLim = null;
    mapa.doubleClickZoom.enable();
    mapa.getContainer().classList.remove('mapa-editando-limite');
    $('#lim-barra').hidden = true;
    render();
    if (salvar) toast('Linha da pluma salva (ajustada à mão).');
  }

  function voltarLimiteCalculado() {
    try { localStorage.removeItem(chaveLimite()); } catch (e) { /* nada */ }
    const ed = st.editLim; if (ed) { sairDesenhoLimite(); mapa.removeLayer(ed.camada); st.editLim = null; mapa.doubleClickZoom.enable(); mapa.getContainer().classList.remove('mapa-editando-limite'); $('#lim-barra').hidden = true; }
    render(); toast('Linha da pluma voltou a ser a calculada.');
  }

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
    const manual = limiteManual();
    // mapa de calor: pinta todos os valores (do piso ao máximo) com a rampa azul-vermelho; a transparência cai nos valores baixos
    const calor = st.pluma.calor && !manual, cmin = Math.max(1e-9, calc.piso || vmin / 100), lcmin = Math.log10(cmin), lcmax = Math.log10(vmax);
    $('#lim-estado').hidden = !manual;
    if (st.editLim && st.editLim.chave !== chaveLimite()) { mapa.removeLayer(st.editLim.camada); st.editLim = null; mapa.doubleClickZoom.enable(); mapa.getContainer().classList.remove('mapa-editando-limite'); $('#lim-barra').hidden = true; }
    let cv = document.createElement('canvas');
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
        const o = ((H - 1 - py) * W + px) * 4; // linha 0 da imagem = norte
        if (calor) {
          const t = (l - lcmin) / ((lcmax - lcmin) || 1);
          if (!(t > .02)) continue;
          const [r, gg, bb] = WGPluma.corCalor(Math.pow(10, l), cmin, vmax);
          img.data[o] = r; img.data[o + 1] = gg; img.data[o + 2] = bb; img.data[o + 3] = Math.round(60 + 140 * Math.min(1, t * 1.6));
          continue;
        }
        if (!(l >= lmin) && !manual) continue; // transparente (com a linha ajustada à mão, o recorte é feito pela linha)
        const [r, gg, bb] = WGPluma.cor(Math.pow(10, Math.max(l, lmin)), vmin, vmax);
        img.data[o] = r; img.data[o + 1] = gg; img.data[o + 2] = bb; img.data[o + 3] = 190;
      }
    }
    ctx.putImageData(img, 0, 0);
    if (manual) cv = mascararPluma(cv, g, k, polisSuaves(manual.polis));
    const sw = g.proj.paraLatLon(g.xmin, g.ymin), ne = g.proj.paraLatLon(g.xmin + g.nx * g.cel, g.ymin + g.ny * g.cel);
    L.imageOverlay(cv.toDataURL(), [[sw.lat, sw.lon], [ne.lat, ne.lon]], { pane: 'pluma', interactive: false, className: 'img-pluma' }).addTo(camadaPluma);

    // 2) linha do limiar
    if (calc.limiar != null) {
      // contínua onde a linha passa entre um poço acima e um abaixo do limite; tracejada onde é inferida
      linhasDoLimite(calc, sel).forEach(t => {
        L.polyline(t.pts, { pane: 'plumaLinha', color: '#ffffff', weight: 5, opacity: .75, interactive: false }).addTo(camadaPluma);
        L.polyline(t.pts, { pane: 'plumaLinha', color: '#172621', weight: 2.4, dashArray: t.definido ? null : '7 6', interactive: false }).addTo(camadaPluma);
      });
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
    $('#leg-tit').textContent = calor ? `Mapa de calor (${un})` : `Pluma (${un})`;
    $('#leg-gradiente').style.background = `linear-gradient(to right, ${(calor ? WGPluma.RAMPA_CALOR : WGPluma.RAMPA.slice(2)).join(', ')})`;
    const v0 = calor ? cmin : vmin;
    const pos = v => 100 * Math.log10(v / v0) / Math.log10(vmax / v0);
    const ticks = [v0];
    for (let e = Math.ceil(Math.log10(v0 * 2)); Math.pow(10, e) <= vmax * 1.0001; e++) ticks.push(Math.pow(10, e));
    $('#leg-ticks').innerHTML = ticks.map(v => `<span style="left:${pos(v)}%">${fmt(v, 3)}</span>`).join('');
    $('#leg-limiar').textContent = calc.limiar != null ? `Limite da pluma delimitada (${fmt(calc.limiar, 4)} ${un})` : 'Limite: sem VI, defina no menu';
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
      L.polyline(tr, { pane: 'pot', color: '#ffffff', weight: 3.5, opacity: .85, interactive: false, lineCap: 'round', lineJoin: 'round' }).addTo(camadaPot);
      L.polyline(tr, { pane: 'pot', color: '#0B3558', weight: 1.6, interactive: false, lineCap: 'round', lineJoin: 'round' }).addTo(camadaPot);
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

  // ------------------------------------------------------------- fotos 360°
  const BUCKET_FOTOS = 'fotos-sondagem';   // o mesmo bucket das fotos do Perfil (pasta por empresa)
  let visor = null;

  async function carregarFotos360() {
    st.f360.lista = []; st.f360.erro = null;
    if (st.projetoId) {
      const { data, error } = await sb.from('wg_foto360').select('*').eq('projeto_id', st.projetoId).order('criado_em');
      if (error) { st.f360.erro = error.message || String(error); console.warn('Fotos 360°:', st.f360.erro); }
      else st.f360.lista = data || [];
    }
    montarPainelF360();
    renderF360Mapa();
  }

  function montarPainelF360() {
    const l = st.f360.lista;
    $('#btn-f360-inserir').hidden = !st.projetoId;
    $('#f360-painel').hidden = !l.length && !st.f360.erro;
    $('#f-f360').checked = st.f360.mostrar;
    $('#f360-conta').textContent = l.length ? `(${l.length})` : '';
    $('#f360-lista').innerHTML = st.f360.erro
      ? `<li class="f360-erro">Fotos 360° ainda não liberadas no banco. Rode o arquivo <b>07_fotos_360.sql</b>.</li>`
      : l.map(f => `<li data-id="${esc(f.id)}">
          <button type="button" class="f360-nome" data-acao="ver" title="Ver em 360°">${esc(f.nome)}</button>
          ${f.latitude == null ? '<span class="f360-sem">sem local</span>' : ''}
          <button type="button" class="btn-mini" data-acao="local" title="Marcar no mapa onde a foto foi tirada">Local</button>
        </li>`).join('');
  }

  const URL_VALE_MS = 50 * 60 * 1000;
  async function urlFoto(caminho) {
    const c = st.f360.urls.get(caminho);
    if (c && c.ate > Date.now()) return c.url;
    const { data, error } = await sb.storage.from(BUCKET_FOTOS).createSignedUrl(caminho, 3600);
    if (error || !data) throw new Error(error?.message || 'Foto não encontrada.');
    st.f360.urls.set(caminho, { url: data.signedUrl, ate: Date.now() + URL_VALE_MS });
    return data.signedUrl;
  }

  function renderF360Mapa() {
    if (!camadaF360) return;
    camadaF360.clearLayers();
    if (!st.f360.mostrar) return;
    st.f360.lista.filter(f => f.latitude != null && f.longitude != null).forEach(f => {
      const m = L.marker([+f.latitude, +f.longitude], {
        icon: L.divIcon({ className: '', html: '<div class="pino-360" aria-hidden="true">360°</div>', iconSize: [34, 22], iconAnchor: [17, 11] }),
        title: f.nome, keyboard: true, zIndexOffset: 300
      });
      m.bindPopup(() => popupF360(f), { maxWidth: 300, minWidth: 260, className: 'popup-360' });
      m.on('click', cliqueSecao);
      m.addTo(camadaF360);
    });
  }

  function popupF360(f) {
    const d = document.createElement('div');
    d.className = 'pop-360';
    d.innerHTML = `
      <button type="button" class="pop-360-mini" data-acao="ver" title="Ver em 360°"><span class="pop-360-carr">carregando…</span><img alt="" hidden><b>Ver em 360°</b></button>
      <div class="pop-360-nome">${esc(f.nome)}</div>
      <div class="pop-360-info">${fmtData(f.criado_em)} · ${fmt(+f.latitude, 6)}, ${fmt(+f.longitude, 6)}</div>
      <div class="pop-360-acoes">
        <button type="button" class="btn btn-primary btn-sm2" data-acao="ver">Ver foto 360°</button>
        <button type="button" class="btn-mini" data-acao="local">Mudar local</button>
        <button type="button" class="btn-mini btn-mini-perigo" data-acao="excluir">Excluir</button>
      </div>`;
    const img = d.querySelector('img');
    if (f.miniatura) urlFoto(f.miniatura).then(u => { img.onload = () => { img.hidden = false; d.querySelector('.pop-360-carr').remove(); }; img.src = u; })
      .catch(() => { d.querySelector('.pop-360-carr').textContent = 'miniatura indisponível'; });
    else d.querySelector('.pop-360-carr').textContent = '';
    d.addEventListener('click', e => {
      const b = e.target.closest('[data-acao]'); if (!b) return;
      if (b.dataset.acao === 'ver') { mapa.closePopup(); verFoto360(f); }
      if (b.dataset.acao === 'local') { mapa.closePopup(); marcarLocalF360({ foto: f }); }
      if (b.dataset.acao === 'excluir') {
        if (b.dataset.confirmar) excluirFoto360(f, b);
        else { b.dataset.confirmar = '1'; b.textContent = 'Confirmar exclusão'; b.classList.add('confirmando'); }
      }
    });
    return d;
  }

  // ---- inserir
  function carregarImagem(src) {
    return new Promise((ok, falha) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => falha(new Error('Não consegui abrir a imagem.')); i.src = src; });
  }
  const paraBlob = (cv, q) => new Promise((ok, falha) => cv.toBlob(b => b ? ok(b) : falha(new Error('Não consegui preparar a imagem.')), 'image/jpeg', q));
  const tamanhoTxt = b => b >= 1048576 ? fmt(b / 1048576, 1) + ' MB' : fmt(Math.round(b / 1024), 0) + ' KB';

  /** Reduz a foto (até 4096 x 2048, JPEG) e faz a miniatura com a vista da frente. */
  async function prepararFoto360(arquivo) {
    const gps = /jpe?g$/i.test(arquivo.type) || /\.jpe?g$/i.test(arquivo.name) ? WG360.gpsDoJpeg(await arquivo.arrayBuffer()) : null;
    const url = URL.createObjectURL(arquivo);
    try {
      const img = await carregarImagem(url);
      const w = img.naturalWidth, h = img.naturalHeight;
      // tenta a melhor qualidade que fique abaixo de 1,8 MB (leve para abrir no celular e
      // abaixo do limite de envio mais comum em hospedagem, 2 MB)
      const LIMITE = 1.8 * 1048576;
      const cv = document.createElement('canvas');
      const ctx = cv.getContext('2d');
      let foto = null, fim = null;
      for (const [max, q] of [[4096, 0.85], [4096, 0.78], [4096, 0.7], [3072, 0.8], [3072, 0.7], [2048, 0.75]]) {
        const t = WG360.tamanhoFinal(w, h, max);
        if (!fim || t.largura !== fim.largura) {
          fim = t; cv.width = t.largura; cv.height = t.altura;
          ctx.imageSmoothingQuality = 'high';
          ctx.drawImage(img, 0, 0, cv.width, cv.height);
        }
        foto = await paraBlob(cv, q);
        if (foto.size <= LIMITE) break;
      }
      const mc = document.createElement('canvas');
      mc.width = 600; mc.height = 300;
      const mctx = mc.getContext('2d'); mctx.imageSmoothingQuality = 'high';
      mctx.drawImage(cv, cv.width * 0.3, cv.height * 0.3, cv.width * 0.4, cv.height * 0.4, 0, 0, 600, 300); // frente da câmera
      const mini = await paraBlob(mc, 0.75);
      cv.width = cv.height = mc.width = mc.height = 1;   // libera a memória (celular)
      return { foto, mini, largura: fim.largura, altura: fim.altura, original: { w, h, bytes: arquivo.size }, eh360: WG360.eh360(w, h), gps };
    } finally { URL.revokeObjectURL(url); }
  }

  async function escolheuArquivoF360(arquivo) {
    if (!arquivo) return;
    if (!/^image\/(jpeg|png|webp)$/i.test(arquivo.type) && !/\.(jpe?g|png|webp)$/i.test(arquivo.name)) { toast('Escolha uma foto JPG, PNG ou WebP.'); return; }
    toast('Preparando a foto…');
    let p;
    try { p = await prepararFoto360(arquivo); } catch (e) { console.error(e); toast('Não consegui abrir essa foto: ' + (e.message || e)); return; }
    if (st.f360.novo?.previa) URL.revokeObjectURL(st.f360.novo.previa);
    st.f360.novo = { ...p, nome: arquivo.name.replace(/\.[^.]+$/, '').replace(/[_]+/g, ' ').trim().slice(0, 120) || 'Foto 360°',
      lat: p.gps?.lat ?? null, lon: p.gps?.lon ?? null, doGps: !!p.gps, previa: URL.createObjectURL(p.mini) };
    abrirModalF360();
  }

  function abrirModalF360() {
    const n = st.f360.novo; if (!n) return;
    $('#f360-previa-img').src = n.previa;
    $('#f360-previa-info').textContent = `${n.original.w} x ${n.original.h} px · ${tamanhoTxt(n.original.bytes)} → grava ${n.largura} x ${n.altura} px · ${tamanhoTxt(n.foto.size + n.mini.size)}`;
    $('#f360-nome').value = n.nome;
    const aviso = $('#f360-aviso');
    aviso.hidden = n.eh360;
    aviso.textContent = n.eh360 ? '' : `Esta imagem tem proporção ${fmt(n.original.w / n.original.h, 2)}:1. Uma foto 360° completa tem 2:1; ela vai aparecer esticada no visor.`;
    textoLocalF360();
    $('#f360-erro').hidden = true;
    $('#f360-salvar').disabled = false; $('#f360-salvar').textContent = 'Salvar foto';
    $('#modal-f360').hidden = false;
  }
  function textoLocalF360() {
    const n = st.f360.novo; if (!n) return;
    $('#f360-local-txt').textContent = n.lat != null
      ? `${fmt(n.lat, 6)}, ${fmt(n.lon, 6)}${n.doGps ? ' (lida do GPS da câmera)' : ''}`
      : 'ainda não marcada (sem ela a foto não aparece no mapa)';
    $('#f360-marcar').textContent = n.lat != null ? 'Mudar no mapa' : 'Marcar no mapa';
  }
  function fecharModalF360(descartar = true) {
    $('#modal-f360').hidden = true;
    if (descartar && st.f360.novo) { URL.revokeObjectURL(st.f360.novo.previa); st.f360.novo = null; }
  }

  async function salvarFoto360() {
    const n = st.f360.novo; if (!n || !st.projetoId || !st.org?.id) return;
    n.nome = $('#f360-nome').value.trim() || 'Foto 360°';
    const btn = $('#f360-salvar'); btn.disabled = true; btn.textContent = 'Enviando…';
    const erro = msg => { $('#f360-erro').textContent = msg; $('#f360-erro').hidden = false; btn.disabled = false; btn.textContent = 'Salvar foto'; };
    const id = (crypto.randomUUID ? crypto.randomUUID() : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => { const r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 3 | 8)).toString(16); }));
    const pasta = `${st.org.id}/f360-${id}`;
    const caminho = pasta + '/foto.jpg', miniatura = pasta + '/miniatura.jpg';
    try {
      const up1 = await sb.storage.from(BUCKET_FOTOS).upload(caminho, n.foto, { upsert: true, contentType: 'image/jpeg' });
      if (up1.error) throw up1.error;
      const up2 = await sb.storage.from(BUCKET_FOTOS).upload(miniatura, n.mini, { upsert: true, contentType: 'image/jpeg' });
      if (up2.error) throw up2.error;
      const { error } = await sb.from('wg_foto360').upsert({ id, projeto_id: st.projetoId, nome: n.nome.slice(0, 200), latitude: n.lat, longitude: n.lon,
        caminho, miniatura, largura: n.largura, altura: n.altura });
      if (error) { sb.storage.from(BUCKET_FOTOS).remove([caminho, miniatura]); throw error; }
    } catch (e) {
      console.error(e);
      const m = e?.message || String(e);
      erro(/wg_foto360|relation|schema cache|doesn't exist/i.test(m) ? 'O banco ainda não tem a tabela das fotos 360°. Rode o arquivo 07_fotos_360.sql e tente de novo.' : 'Não consegui salvar: ' + m);
      return;
    }
    fecharModalF360();
    toast('Foto 360° salva.');
    await carregarFotos360();
  }

  // ---- marcar localização no mapa (para a foto nova ou para uma já salva)
  function marcarLocalF360(alvo) {
    if (!mapa) return;
    if (st.ajuste) cancelarAjuste(true);
    if (st.secao.tracando) cancelarSecao();
    st.f360.marcar = alvo;
    if (alvo.novo) $('#modal-f360').hidden = true;
    trocarView('visao');
    mapa.closePopup();
    mapa.getContainer().classList.add('mapa-ajustando');
    $('#f360-passo').textContent = `Clique no mapa onde a foto "${alvo.novo ? $('#f360-nome').value.trim() || 'nova' : alvo.foto.nome}" foi tirada`;
    $('#f360-barra').hidden = false;
    $('#f360-barra').scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }
  function fecharMarcarF360() {
    const alvo = st.f360.marcar; st.f360.marcar = null;
    $('#f360-barra').hidden = true;
    mapa.getContainer().classList.remove('mapa-ajustando');
    if (alvo?.novo) { textoLocalF360(); $('#modal-f360').hidden = false; }
  }
  async function cliqueF360(ev) {
    const alvo = st.f360.marcar; if (!alvo) return;
    const { lat, lng } = ev.latlng;
    if (alvo.novo) { const n = st.f360.novo; if (n) { n.nome = $('#f360-nome').value; n.lat = lat; n.lon = lng; n.doGps = false; } fecharMarcarF360(); return; }
    fecharMarcarF360();
    const f = alvo.foto;
    const { error } = await sb.from('wg_foto360').upsert({ id: f.id, projeto_id: f.projeto_id, nome: f.nome, latitude: lat, longitude: lng,
      caminho: f.caminho, miniatura: f.miniatura, largura: f.largura, altura: f.altura });
    if (error) { toast('Não consegui salvar a localização: ' + (error.message || error)); return; }
    toast('Localização salva.');
    await carregarFotos360();
  }

  async function excluirFoto360(f, btn) {
    btn.disabled = true; btn.textContent = 'Excluindo…';
    const { data, error } = await sb.from('wg_foto360').delete().eq('id', f.id).select();
    if (error || !data || !data.length) { btn.disabled = false; btn.textContent = 'Excluir'; delete btn.dataset.confirmar; toast('Não consegui excluir: ' + (error?.message || 'o banco não permitiu.')); return; }
    sb.storage.from(BUCKET_FOTOS).remove([f.caminho, f.miniatura].filter(Boolean)).then(() => {}, () => {});
    mapa.closePopup();
    toast(`Foto "${f.nome}" excluída.`);
    await carregarFotos360();
  }

  // ---- visor
  async function verFoto360(f, blob = null) {
    $('#visor360-nome').textContent = f.nome;
    $('#visor360-info').textContent = f.latitude != null ? `${fmt(+f.latitude, 6)}, ${fmt(+f.longitude, 6)}` : 'sem localização';
    const msg = $('#visor360-msg');
    msg.textContent = 'Carregando foto…'; msg.hidden = false;
    $('#visor360').hidden = false;
    document.body.classList.add('visor-aberto');
    $('#visor360-fechar').focus();
    let src = null, temp = null;
    try {
      if (blob) src = temp = URL.createObjectURL(blob);
      else src = await urlFoto(f.caminho);
      const img = new Image();
      if (!temp) img.crossOrigin = 'anonymous';
      await new Promise((ok, falha) => { img.onload = ok; img.onerror = () => falha(new Error('Não consegui baixar a foto.')); img.src = src; });
      if ($('#visor360').hidden) return;      // fechou enquanto carregava
      if (visor) visor.destruir();
      visor = Visor360.abrir($('#visor360-area'), img);
      msg.hidden = true;
    } catch (e) {
      console.error(e);
      msg.textContent = e.message || String(e);
    } finally { if (temp) setTimeout(() => URL.revokeObjectURL(temp), 1000); }
  }
  function fecharVisor() {
    if (visor) { visor.destruir(); visor = null; }
    $('#visor360').hidden = true;
    document.body.classList.remove('visor-aberto');
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
  }

  function ligarEventosF360() {
    $('#btn-f360-inserir').addEventListener('click', () => { if (!st.projetoId) { toast('Escolha um projeto primeiro.'); return; } $('#arquivo-f360').click(); });
    $('#arquivo-f360').addEventListener('change', e => { const a = e.target.files[0]; e.target.value = ''; escolheuArquivoF360(a); });
    $('#f-f360').addEventListener('change', e => { st.f360.mostrar = e.target.checked; renderF360Mapa(); });
    $('#f360-lista').addEventListener('click', e => {
      const b = e.target.closest('[data-acao]'), li = e.target.closest('li[data-id]'); if (!b || !li) return;
      const f = st.f360.lista.find(x => String(x.id) === li.dataset.id); if (!f) return;
      if (b.dataset.acao === 'ver') verFoto360(f);
      if (b.dataset.acao === 'local') marcarLocalF360({ foto: f });
    });
    $('#f360-fechar').addEventListener('click', () => fecharModalF360());
    $('#f360-cancelar').addEventListener('click', () => fecharModalF360());
    $('#modal-f360').addEventListener('click', e => { if (e.target.id === 'modal-f360') fecharModalF360(); });
    $('#f360-marcar').addEventListener('click', () => { if (st.f360.novo) st.f360.novo.nome = $('#f360-nome').value; marcarLocalF360({ novo: true }); });
    $('#f360-previa-ver').addEventListener('click', () => { const n = st.f360.novo; if (n) verFoto360({ nome: $('#f360-nome').value || n.nome, latitude: n.lat, longitude: n.lon }, n.foto); });
    $('#f360-salvar').addEventListener('click', salvarFoto360);
    $('#f360-cancelar-marcar').addEventListener('click', fecharMarcarF360);
    $('#visor360-fechar').addEventListener('click', fecharVisor);
    $('#visor360-mais').addEventListener('click', () => visor && visor.zoom(0.8));
    $('#visor360-menos').addEventListener('click', () => visor && visor.zoom(1.25));
    $('#visor360-centro').addEventListener('click', () => visor && visor.centro());
    $('#visor360-tela').addEventListener('click', () => {
      const el = $('#visor360');
      if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
      else if (el.requestFullscreen) el.requestFullscreen().catch(() => {});
    });
    document.addEventListener('keydown', e => {
      if ($('#visor360').hidden) {
        if (e.key === 'Escape' && st.f360.marcar) fecharMarcarF360();
        return;
      }
      if (e.key === 'Escape' && !document.fullscreenElement) fecharVisor();
      if (!visor) return;
      const passos = { ArrowLeft: [-10, 0], ArrowRight: [10, 0], ArrowUp: [0, 8], ArrowDown: [0, -8] };
      if (passos[e.key]) { e.preventDefault(); visor.girar(...passos[e.key]); }
      if (e.key === '+' || e.key === '=') visor.zoom(0.8);
      if (e.key === '-') visor.zoom(1.25);
    });
  }

  // ------------------------------------------------------------- anotações do projeto
  const ESPERA_SALVAR = 1500;   // salva sozinho 1,5 s depois que a pessoa para de digitar

  async function carregarAnotacao() {
    const a = st.anot;
    if (a.sujo && a.projeto && a.projeto !== st.projetoId) await salvarAnotacao({ forcar: true });   // trocou de projeto com texto por salvar
    clearTimeout(a.timer);
    Object.assign(a, { projeto: st.projetoId, texto: '', base: '', versao: null, autor: null, erro: null, sujo: false, conflito: null });
    if (st.projetoId) {
      const { data, error } = await sb.from('wg_anotacao').select('*').eq('projeto_id', st.projetoId);
      if (st.projetoId !== a.projeto) return;                 // trocou de projeto enquanto carregava
      if (error) { a.erro = error.message || String(error); console.warn('Anotações:', a.erro); }
      else if (data && data.length) Object.assign(a, { texto: data[0].texto || '', base: data[0].texto || '', versao: data[0].atualizado_em, autor: data[0].atualizado_por_nome });
    }
    $('#anot-texto').value = a.texto;
    renderAnotacao();
  }

  function renderAnotacao(estado) {
    const a = st.anot, proj = st.projetos.find(p => p.id === st.projetoId);
    $('#anot-projeto').textContent = proj ? proj.nome : '';
    $('#anot-texto').disabled = !st.projetoId || !!a.erro;
    $('#anot-texto').placeholder = st.projetoId
      ? 'Escreva aqui o que a equipe precisa saber deste projeto: histórico da área, combinados com o cliente, pendências de campo, observações das campanhas...'
      : 'Escolha um projeto no menu.';
    const n = $('#anot-texto').value.length;
    $('#anot-conta').textContent = n ? `${n.toLocaleString('pt-BR')} caractere${n === 1 ? '' : 's'}` : '';
    const s = $('#anot-status');
    s.className = 'anot-status';
    if (a.erro) { s.classList.add('erro'); s.innerHTML = 'As anotações ainda não estão liberadas no banco. Rode o arquivo <b>banco/08_anotacoes.sql</b> no SQL Editor do Supabase.'; }
    else if (estado === 'salvando') s.textContent = 'Salvando…';
    else if (estado === 'falhou') { s.classList.add('erro'); s.textContent = 'Não consegui salvar. Confira a internet; vou tentar de novo quando você voltar a digitar.'; }
    else if (a.sujo) s.textContent = 'Alterações ainda não salvas…';
    else if (a.versao) { s.classList.add('ok'); s.textContent = `Salvo · última alteração em ${new Date(a.versao).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' })}${a.autor ? ' por ' + a.autor : ''}`; }
    else s.textContent = st.projetoId ? 'Nenhuma anotação ainda. O texto é salvo sozinho enquanto você escreve.' : '';
    $('#anot-conflito').hidden = !a.conflito;
    if (a.conflito) $('#anot-conflito-txt').textContent = `${a.conflito.autor || 'Outra pessoa'} salvou outra versão em ${new Date(a.conflito.versao).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' })}, enquanto você editava.`;
  }

  function digitouAnotacao() {
    const a = st.anot; if (!a.projeto || a.erro) return;
    a.sujo = $('#anot-texto').value !== a.base;
    clearTimeout(a.timer);
    if (a.sujo && !a.conflito) a.timer = setTimeout(() => salvarAnotacao(), ESPERA_SALVAR);
    renderAnotacao();
  }

  /** Grava o texto. Antes confere se ninguém salvou outra versão no meio do caminho. */
  async function salvarAnotacao({ forcar = false } = {}) {
    const a = st.anot, proj = a.projeto;
    if (!proj || !a.sujo || a.salvando || a.erro) return;
    clearTimeout(a.timer);
    a.salvando = true;
    const texto = $('#anot-texto').value;            // o campo ainda mostra o texto deste projeto (é trocado só depois de salvar)
    a.texto = texto;
    if (proj === st.projetoId) renderAnotacao('salvando');
    try {
      if (!forcar) {
        const { data, error } = await sb.from('wg_anotacao').select('texto, atualizado_em, atualizado_por_nome').eq('projeto_id', proj);
        if (error) throw error;
        const r = data && data[0];
        if (r && r.atualizado_em !== a.versao && (r.texto || '') !== a.base && (r.texto || '') !== texto) {
          a.conflito = { texto: r.texto || '', versao: r.atualizado_em, autor: r.atualizado_por_nome };
          a.salvando = false;
          if (proj === st.projetoId) renderAnotacao();
          return;
        }
      }
      const { error } = await sb.from('wg_anotacao').upsert({ projeto_id: proj, texto, atualizado_por_nome: (st.profile?.full_name || st.user?.email || '').slice(0, 200) });
      if (error) throw error;
      const { data } = await sb.from('wg_anotacao').select('atualizado_em, atualizado_por_nome').eq('projeto_id', proj);
      if (proj !== st.projetoId) return;
      a.base = texto; a.conflito = null;
      if (data && data[0]) { a.versao = data[0].atualizado_em; a.autor = data[0].atualizado_por_nome; }
      a.sujo = $('#anot-texto').value !== a.base;               // continuou digitando enquanto salvava
      if (a.sujo) a.timer = setTimeout(() => salvarAnotacao(), ESPERA_SALVAR);
    } catch (e) {
      console.error(e);
      if (proj === st.projetoId) { a.salvando = false; renderAnotacao('falhou'); return; }
    } finally { a.salvando = false; }
    if (proj === st.projetoId) renderAnotacao();
  }

  function resolverConflito(qual) {
    const a = st.anot, c = a.conflito; if (!c) return;
    if (qual === 'deles') {
      Object.assign(a, { texto: c.texto, base: c.texto, versao: c.versao, autor: c.autor, sujo: false, conflito: null });
      $('#anot-texto').value = c.texto;
      renderAnotacao();
      toast('Carregada a versão salva por ' + (c.autor || 'outra pessoa') + '.');
    } else {
      a.conflito = null;
      salvarAnotacao({ forcar: true }).then(() => toast('A sua versão foi salva.'));
    }
  }

  function ligarEventosAnotacao() {
    const t = $('#anot-texto');
    t.addEventListener('input', digitouAnotacao);
    t.addEventListener('blur', () => { if (st.anot.sujo && !st.anot.conflito) salvarAnotacao(); });
    t.addEventListener('keydown', e => { if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); if (!st.anot.conflito) salvarAnotacao(); } });
    $('#anot-manter-minha').addEventListener('click', () => resolverConflito('minha'));
    $('#anot-usar-deles').addEventListener('click', () => resolverConflito('deles'));
    window.addEventListener('beforeunload', e => { if (st.anot.sujo) { salvarAnotacao({ forcar: true }); e.preventDefault(); e.returnValue = ''; } });
  }

  /** Bloco das anotações no relatório em PDF (mantém as quebras de linha). */
  function anotacaoRelatorio() {
    const a = st.anot, texto = ($('#anot-texto').value || '').trim();
    if (!texto) return '';
    const quando = a.versao ? `Última alteração em ${new Date(a.versao).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' })}${a.autor ? ' por ' + esc(a.autor) : ''}.` : '';
    return `<h2>Anotações do projeto</h2><div class="rel-anot-texto">${esc(texto)}</div>${quando ? `<p class="rel-anot-quando">${quando}</p>` : ''}`;
  }

  // ------------------------------------------------------------- seções geológicas (cortes A–A', B–B', C–C'...)
  const chaveSecao = () => 'webgeo.secao.' + (st.projetoId || '');
  const LETRAS_SECAO = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const fimSecao = l => l + "'";
  const nomeSecao = l => `${l}–${fimSecao(l)}`;
  const linhaValida = v => Array.isArray(v) && v.length >= 2 && v.every(p => Array.isArray(p) && Number.isFinite(p[0]) && Number.isFinite(p[1]));
  /** Primeira letra que ainda não tem seção (se apagou a B, a próxima nova volta a ser B). */
  const proximaLetra = () => [...LETRAS_SECAO].find(l => !st.secao.lista.some(x => x.letra === l)) || 'Z';
  /** st.secao.linha e st.secao.letra = a seção aberta (o resto do código desenha e calcula só ela). */
  function sincronizarSecao() {
    const l = st.secao.lista;
    if (st.secao.ativa >= l.length) st.secao.ativa = Math.max(0, l.length - 1);
    const a = l[st.secao.ativa];
    st.secao.linha = a ? a.linha : null;
    st.secao.letra = a ? a.letra : proximaLetra();
  }

  /** As seções ficam guardadas neste navegador, por projeto. Aceita também o formato antigo (uma linha só = A–A'). */
  function carregarSecao() {
    st.secao.tracando = null; st.secao.lista = []; st.secao.ativa = 0;
    try {
      const v = JSON.parse(localStorage.getItem(chaveSecao()) || 'null');
      if (linhaValida(v)) st.secao.lista = [{ letra: 'A', linha: v }];
      else if (v && Array.isArray(v.lista)) {
        st.secao.lista = v.lista.filter(x => x && LETRAS_SECAO.includes(x.letra) && linhaValida(x.linha)).map(x => ({ letra: x.letra, linha: x.linha }));
        st.secao.ativa = Number.isInteger(v.ativa) ? v.ativa : 0;
      }
    } catch (e) { /* sem armazenamento: começa sem seção */ }
    st.secao.lista.sort((a, b) => a.letra.localeCompare(b.letra));
    sincronizarSecao();
    $('#secao-barra').hidden = true;
    if (mapa) { mapa.getContainer().classList.remove('mapa-ajustando'); mapa.doubleClickZoom.enable(); }
  }
  function guardarSecao() {
    try {
      if (st.secao.lista.length) localStorage.setItem(chaveSecao(), JSON.stringify({ v: 2, lista: st.secao.lista, ativa: st.secao.ativa }));
      else localStorage.removeItem(chaveSecao());
    } catch (e) { /* sem armazenamento */ }
  }
  function abrirSecao(i) {
    if (i < 0 || i >= st.secao.lista.length || st.secao.tracando) return;
    st.secao.ativa = i; sincronizarSecao(); guardarSecao(); renderSecao();
  }

  function iniciarSecao() {
    if (!mapa || !st.projetoId) { toast('Escolha um projeto primeiro.'); return; }
    if (st.ajuste) cancelarAjuste(true);
    trocarView('visao');
    st.secao.tracando = { pontos: [], letra: proximaLetra() };
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
    $('#secao-passo').textContent = n === 0 ? `Seção ${nomeSecao(t.letra)}: clique no mapa onde ela começa (ponto ${t.letra})`
      : n === 1 ? `Agora clique onde ela termina (ponto ${fimSecao(t.letra)})` : `${n} pontos marcados. Clique em "Concluir" ou marque mais um ponto`;
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
    st.secao.lista.push({ letra: t.letra, linha: t.pontos });
    st.secao.lista.sort((a, b) => a.letra.localeCompare(b.letra));
    st.secao.ativa = st.secao.lista.findIndex(x => x.letra === t.letra);
    sincronizarSecao();
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
  function apagarSecao() {
    const a = st.secao.lista[st.secao.ativa]; if (!a) return;
    st.secao.lista.splice(st.secao.ativa, 1);
    sincronizarSecao(); guardarSecao(); render();
    toast(`Seção ${nomeSecao(a.letra)} apagada.`);
  }

  /** Linhas das seções no mapa, com a letra nas pontas (A e A', B e B'...). A aberta em amarelo; as outras em branco.
   *  Clicar numa ponta de outra seção abre essa seção. */
  function renderSecaoMapa() {
    if (!camadaSecao) return;
    camadaSecao.clearLayers();
    const t = st.secao.tracando;
    const pino = (ll, txt, outra, i) => {
      const m = L.marker(ll, { icon: L.divIcon({ className: '', html: `<div class="pino-secao${outra ? ' outra' : ''}" style="${outra ? 'background:#ffffff;' : ''}">${txt}</div>`, iconSize: [24, 22] }),
        interactive: outra && !t, keyboard: false, zIndexOffset: outra ? 400 : 500, title: outra ? `Abrir a seção ${nomeSecao(txt.replace("'", ''))}` : '' }).addTo(camadaSecao);
      if (outra && !t) m.on('click', () => abrirSecao(i));
    };
    const desenhar = (pts, letra, outra, i, tracando) => {
      if (!pts || !pts.length) return;
      if (pts.length >= 2) {
        L.polyline(pts, { pane: 'secao', color: '#172621', weight: outra ? 5 : 6, opacity: outra ? .45 : .6, interactive: false }).addTo(camadaSecao);
        L.polyline(pts, { pane: 'secao', color: outra ? '#ffffff' : '#FFD60A', weight: outra ? 2 : 2.5, dashArray: tracando ? '6 6' : (outra ? '8 5' : null), interactive: false }).addTo(camadaSecao);
      }
      pino(pts[0], letra, outra, i);
      if (pts.length >= 2) pino(pts[pts.length - 1], fimSecao(letra), outra, i);
      pts.slice(1, -1).forEach(ll => L.circleMarker(ll, { pane: 'secao', radius: outra ? 3 : 4, color: '#172621', weight: 1.5, fillColor: outra ? '#ffffff' : '#FFD60A', fillOpacity: 1, interactive: false }).addTo(camadaSecao));
    };
    st.secao.lista.forEach((x, i) => { if (t || i !== st.secao.ativa) desenhar(x.linha, x.letra, true, i, false); });
    if (t) desenhar(t.pontos, t.letra, false, -1, true);
    else if (st.secao.linha) desenhar(st.secao.linha, st.secao.letra, false, st.secao.ativa, false);
    if (t || !st.secao.linha) return;
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

  /**
   * Pluma vertical da seção: cada poço com resultado entra no meio do filtro (sem filtro: no meio entre o N.A. e o fundo).
   * Mesma escala de cores, limite e regra do LQ da pluma do mapa. Devolve null se não há nenhum resultado.
   */
  function plumaVerticalSecao(cols, comp, zmin, zmax, nx, nz) {
    const param = st.parametros.find(p => p.id === st.parametroId);
    if (!param) return null;
    const comRes = cols.filter(c => c.res);
    if (!comRes.length) return null;
    const calc = st.calc, limiar = limiarAtual(param);
    const quant = comRes.filter(c => !c.res.menor_que_lq && +c.res.valor > 0).map(c => +c.res.valor);
    const base = param.valor_orientador != null ? +param.valor_orientador : (limiar ?? (quant.length ? Math.min(...quant) : 1));
    const piso = base / 10;
    const valor = r => r.menor_que_lq ? (r.lq != null && +r.lq > 0 ? +r.lq / 2 : piso) : (+r.valor > 0 ? +r.valor : piso);
    let semFiltro = 0;
    const amostras = cols.map((c, i) => {
      if (!c.res) return null;
      const f = c.tubo.filter(t => t.tipo === 'FILTRO');
      let prof, fundo, fonte = 'filtro';
      if (f.length) { fundo = Math.max(...f.map(t => t.ate)); prof = (Math.min(...f.map(t => t.de)) + fundo) / 2; }
      else if (c.prof > 0) { fundo = c.prof; prof = ((c.na ?? 0) + c.prof) / 2; fonte = 'trecho saturado'; semFiltro++; }
      else return null;
      return { i, s: c.dist, z: c.cota - prof, zBase: c.cota - fundo, v: valor(c.res), nome: c.nome, texto: `${textoValor(c.res)} ${c.res.unidade || param.unidade || ''}`.trim(), fonte };
    }).filter(Boolean);
    const out = { amostras, param: param.nome, un: param.unidade || 'µg/L', limiar, semFiltro, img: null, segs: [] };
    if (!amostras.length) { out.motivo = 'pluma vertical: os poços com resultado não têm filtro nem profundidade cadastrados'; return out; }
    // topo da pluma = nível d'água (o mais raso de cada ponto); sem N.A., o terreno. Base = fundo das sondagens.
    const agrupa = (pts, maisAlto) => {
      const m = new Map();
      pts.forEach(([s, z]) => { const k = Math.round(s); if (!m.has(k) || (maisAlto ? z > m.get(k)[1] : z < m.get(k)[1])) m.set(k, [s, z]); });
      return [...m.values()];
    };
    const comNA = cols.filter(c => c.na != null);
    const topo = WGPlumaV.perfil(agrupa((comNA.length ? comNA.map(c => [c.dist, c.cota - c.na]) : cols.map(c => [c.dist, c.cota])), true));
    // base = fundo das sondagens ligado de um poço ao outro; um poço mais fundo (PMN) abre um "cone" que estreita para os lados
    const fundos = cols.filter(c => c.prof > 0).map(c => [c.dist, c.cota - c.prof]);
    amostras.forEach(a => fundos.push([a.s, a.zBase]));
    const zAm = Math.min(...amostras.map(a => a.z)), INCL = 0.15;
    const linhaFundo = WGPlumaV.perfil(agrupa(fundos, false));
    const baseF = s => Math.min(linhaFundo(s), ...fundos.map(([sf, zf]) => zf + Math.abs(s - sf) * INCL));
    let vmax = calc?.vmax, vmin = calc?.vmin;
    if (!(vmax > 0 && vmin > 0)) {
      const mx = Math.max(...amostras.map(a => a.v));
      vmax = Math.pow(10, Math.max(1, Math.ceil(Math.log10(mx || 10))));
      vmin = limiar != null ? limiar : vmax / 1e3;
      if (vmax <= vmin * 1.5) vmax = vmin * 10;
    }
    // poços a até 5 m um do outro (PM + PMN) formam um perfil só, com uma amostra por profundidade
    const grupos = [];
    [...amostras].sort((a, b) => a.s - b.s).forEach(a => {
      const g = grupos.find(x => Math.abs(x.s - a.s) <= 5);
      const na = topo(a.s), pt = { d: na - a.z, dBase: na - a.zBase, l: Math.log10(a.v) };
      if (g) { g.pts.push(pt); g.s = (g.s * (g.pts.length - 1) + a.s) / g.pts.length; } else grupos.push({ s: a.s, pts: [pt] });
    });
    const sg = grupos.map(g => g.s), gaps = sg.slice(1).map((x, q) => x - sg[q]).filter(x => x > 0.5).sort((a, b) => a - b);
    const espH = Math.max(5, 0.5 * (gaps.length ? gaps[Math.floor(gaps.length / 2)] : comp * 0.2));
    nx = Math.max(60, Math.min(420, nx)); nz = Math.max(40, Math.min(300, nz));
    // a grade cobre o corte todo (sem recorte), para a linha do limite atravessar o N.A.; o recorte é feito no desenho
    const G = WGPlumaV.gradePerfis(grupos, { s0: 0, s1: comp, z0: zmin, z1: zmax, nx, nz, p: st.pluma.p || 2,
      lLim: Math.log10(limiar ?? vmin), lPiso: Math.log10(piso), espH, topo });
    // só pinta onde passa do mínimo da escala (o limite, como no mapa)
    const cv = document.createElement('canvas'); cv.width = nx; cv.height = nz;
    const ctx = cv.getContext('2d'), img = ctx.createImageData(nx, nz);
    let pinta = 0;
    for (let q = 0; q < nx * nz; q++) {
      const v = G.v[q];
      if (!(v >= vmin)) continue;
      const sq = (q % nx + 0.5) * comp / nx, zq = zmax - (Math.floor(q / nx) + 0.5) * (zmax - zmin) / nz;
      if (zq <= topo(sq) + (zmax - zmin) / nz && zq >= baseF(sq) - (zmax - zmin) / nz) pinta++;
      const [r, gg, bb] = WGPluma.cor(v, vmin, vmax);
      img.data[q * 4] = r; img.data[q * 4 + 1] = gg; img.data[q * 4 + 2] = bb; img.data[q * 4 + 3] = 255;
    }
    Object.assign(out, { topo, base: baseF, vmin, vmax, s0: 0, ds: comp / nx, z1: zmax, dz: (zmax - zmin) / nz,
      mesmaProf: amostras.length > 1 && Math.max(...amostras.map(a => a.z)) - zAm < 1 });
    if (!pinta) { out.motivo = `pluma vertical: nenhum ponto da seção passa do ${limiar != null ? 'limite' : 'mínimo da escala'} (${fmt(vmin, 4)} ${out.un})`; return out; }
    ctx.putImageData(img, 0, 0);
    out.img = cv.toDataURL();
    // linhas (em metros: [s, cota]) do limite e de 10x, 100x... o limite; só o trecho dentro da zona saturada
    const dzc = (zmax - zmin) / nz, dsc = comp / nx;
    const curvasDe = val => {
      const pt = (x, y) => ({ x: x * dsc, y: zmax - y * dzc });
      const dentro = q => q.y <= topo(q.x) + dzc && q.y >= baseF(q.x) - dzc;
      const segs = WGPlumaV.isolinha(G, val).map(q => [pt(q[0], q[1]), pt(q[2], q[3])]).filter(([a, b]) => dentro(a) || dentro(b));
      return WGPluma.emendar(segs).map(c => WGPluma.suavizar(c, 3)).filter(c => c.length > 1).map(c => c.map(q => [q.x, q.y]));
    };
    if (limiar != null) {
      out.segs = [1];
      out.curvas = [{ v: limiar, limite: true, linhas: curvasDe(limiar) }];
      for (let v = limiar * 10; v < vmax * 0.999; v *= 10) out.curvas.push({ v, limite: false, linhas: curvasDe(v) });
    }
    out.dica = `Pluma vertical de ${param.nome} (estimativa entre os poços, só abaixo do nível d'água)`;
    return out;
  }

  function renderSecao() {
    const tem = !!st.secao.linha, L_ = st.secao.letra;
    $('#sec-secao').hidden = !tem;
    $('#secao-painel').hidden = !tem;
    $('#btn-secao-tracar').textContent = tem ? `Traçar seção ${nomeSecao(proximaLetra())}` : `Traçar seção ${nomeSecao(proximaLetra())}`;
    $('#btn-secao-apagar').textContent = tem ? `Apagar seção ${nomeSecao(L_)}` : 'Apagar seção';
    const tit = $('#sec-secao .section-head h2'); if (tit) tit.textContent = `Seção geológica ${nomeSecao(L_)}`;
    const num = $('#sec-secao .section-head .section-num'); if (num) num.textContent = L_;
    abasSecao();
    const opc = $('#cerca3d-modo option[value="secao"]'); if (opc) opc.textContent = tem ? `Da seção ${nomeSecao(L_)}` : 'Da seção';
    st.secaoCalc = tem ? montarSecao() : null;
    renderSecaoMapa();
    if (!tem) return;
    const d = st.secaoCalc, host = $('#secao-desenho');
    $('#btn-secao-baixar').disabled = $('#btn-secao-dxf').disabled = !(d && d.cols.length);
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
      + `<text x="${mL}" y="20" font-size="17" font-weight="700" fill="currentColor">${L_}</text>`
      + `<text x="${W - mR}" y="20" font-size="17" font-weight="700" text-anchor="end" fill="currentColor">${fimSecao(L_)}</text>`;
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

    // pluma vertical: concentração interpolada no plano do corte, entre os filtros dos poços (zona saturada)
    const PV = st.secao.pluma ? plumaVerticalSecao(cols, d.comp, zmin, zmax, Math.round(pw / 3), Math.round(ph / 3)) : null;
    if (PV && PV.img) {
      // recorte da zona saturada (do N.A. ao fundo): a pluma e a linha do limite terminam nele, com borda lisa
      const idc = `pv-recorte-${L_}`, NS = 160;
      const ss = Array.from({ length: NS + 1 }, (_, k) => d.comp * k / NS);
      const contorno = [...ss.map(s => `${n1(X(s))},${n1(Y(PV.topo(s)))}`), ...ss.reverse().map(s => `${n1(X(s))},${n1(Y(PV.base(s)))}`)].join(' ');
      g += `<clipPath id="${idc}"><polygon points="${contorno}"/></clipPath><g clip-path="url(#${idc})">`;
      g += `<image href="${PV.img}" x="${n1(X(0))}" y="${n1(Y(zmax))}" width="${n1(pw)}" height="${n1(ph)}" preserveAspectRatio="none" opacity=".6"><title>${esc(PV.dica)}</title></image>`;
      if (PV.curvas && PV.curvas[0].linhas.length) {
        const dd = PV.curvas[0].linhas.map(c => 'M' + c.map(p => `${n1(X(p[0]))} ${n1(Y(p[1]))}`).join('L')).join('');
        g += `<path d="${dd}" fill="none" stroke="#ffffff" stroke-opacity=".7" stroke-width="4.4" stroke-linecap="round" stroke-linejoin="round"/><path d="${dd}" fill="none" stroke="#D7191C" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>`;
      }
      g += '</g>';
    }

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

    // pontos amostrados da pluma vertical (no meio do filtro), por cima das colunas
    if (PV && PV.img) PV.amostras.forEach(a => {
      g += `<circle cx="${n1(xd[a.i])}" cy="${n1(Y(a.z))}" r="3.4" fill="#ffffff" stroke="#172621" stroke-width="1.3"><title>${esc(`${a.nome}: ${a.texto} no meio do ${a.fonte} (cota ${fmt(a.z, 2)} m)`)}</title></circle>`;
    });

    // nível d'água: uma linha por rede (PM e PMN medem níveis diferentes do aquífero)
    const TRACOS = ['7 3', '2 3', '9 3 2 3', '4 4'];
    const redes = [...new Set(cols.filter(c => c.na != null).map(c => c.rede))];
    redes.forEach((r, k) => {
      const pts = cols.map((c, i) => c.rede === r && c.na != null ? n1(xd[i]) + ',' + n1(Y(c.cota - c.na)) : null).filter(Boolean);
      if (pts.length >= 2) g += `<polyline points="${pts.join(' ')}" fill="none" stroke="#1565A8" stroke-width="1.5" stroke-dasharray="${TRACOS[k % TRACOS.length]}"/>`;
    });

    const camp0 = st.campanhas.find(c => c.id === st.campanhaId), par0 = st.parametros.find(p => p.id === st.parametroId);
    st.secaoDes = { mL, mT, pxH, ev, zmax, ph, W, semDatum: d.semDatum, PV: PV && PV.img ? PV : null,
      titulo: [`Seção ${nomeSecao(L_)}`, st.projetos.find(p => p.id === st.projetoId)?.nome, PV && PV.img && par0 ? `pluma de ${par0.nome}` : '', camp0 ? `campanha ${camp0.codigo}` : ''].filter(Boolean).join(' - ') };
    host.innerHTML = `<svg id="secao-svg" xmlns="http://www.w3.org/2000/svg" width="${W}" height="${n1(H)}" viewBox="0 0 ${W} ${n1(H)}" font-family="IBM Plex Sans, sans-serif" role="img" aria-label="Seção geológica ${nomeSecao(L_)}">${g}</svg>`;

    // legenda e notas
    const item = (desenho, texto) => `<span>${desenho}${esc(texto)}</span>`;
    let leg = [...litos.values()].map(l => item(`<i class="leg-lito" style="background:${esc(l.hex)}"></i>`, l.texto)).join('');
    leg += item('<i class="leg-terreno"></i>', 'Terreno (cota do topo)');
    if (cols.some(c => c.tubo.length)) leg += item('<i class="leg-tubo"></i>', 'Tubo liso') + item('<i class="leg-tubo leg-filtro"></i>', 'Filtro');
    redes.forEach((r, k) => { leg += `<span><svg width="26" height="8"><line x1="0" x2="26" y1="4" y2="4" stroke="#1565A8" stroke-width="1.6" stroke-dasharray="${TRACOS[k % TRACOS.length]}"/></svg>${esc("Nível d'água" + (r !== '-' ? ' · ' + r : ''))}</span>`; });
    if (PV && PV.img) {
      leg += `<span><i class="leg-pluma-v" style="background:linear-gradient(to right, ${WGPluma.RAMPA.slice(2).join(', ')})"></i>${esc(`Pluma vertical · ${PV.param} (${fmt(PV.vmin, 3)} a ${fmt(PV.vmax, 3)} ${PV.un})`)}</span>`;
      if (PV.limiar != null) leg += `<span><svg width="26" height="8"><line x1="0" x2="26" y1="4" y2="4" stroke="#D7191C" stroke-width="2.4"/></svg>${esc(`Limite (${fmt(PV.limiar, 4)} ${PV.un})`)}</span>`;
      leg += `<span><svg width="10" height="10"><circle cx="5" cy="5" r="3.2" fill="#fff" stroke="#172621" stroke-width="1.2"/></svg>Ponto amostrado (meio do filtro)</span>`;
    }
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
    if (st.secao.pluma) {
      if (!PV) notas.push('pluma vertical: sem resultados do parâmetro nesta campanha para os poços da seção');
      else if (!PV.img) notas.push(PV.motivo);
      else {
        notas.push(`pluma vertical de ${PV.param}${camp ? ` na campanha ${camp.codigo}` : ''}: interpolada (IDW em escala log) entre ${PV.amostras.length} ponto(s) amostrado(s), só abaixo do nível d'água; em cada poço a concentração parte do valor medido no N.A. e cai até o limite no fundo do filtro, por isso a pluma afina com a distância; abaixo do LQ entra como metade do LQ; é uma estimativa e precisa de conferência`);
        if (PV.mesmaProf) notas.push('todos os poços amostram quase a mesma profundidade, então a pluma mostra só a variação lateral (para ver a variação com a profundidade são precisos poços multinível, como PM e PMN)');
        if (PV.semFiltro) notas.push(`${PV.semFiltro} poço(s) sem filtro cadastrado: amostra no meio da zona saturada`);
      }
    }
    $('#secao-nota').textContent = notas.join(' · ') + '. Passe o mouse sobre uma sondagem para ver os detalhes.';
    $('#secao-hint').textContent = `${cols.length} ponto(s) · ${fmt(d.comp, 0)} m`;
  }

  // ------------------------------------------------------------- diagrama de cercas 3D
  let visor3d = null;
  const cerca = { modo: 'secao', exag: 0, paineis: true, agua: true, cols: [], ex: 1, w: 1 };

  /** Sondagens do diagrama, com posição em metros (x leste, y norte) e cota relativa ao ponto mais alto. */
  function colunasCerca(modo) {
    let base = [];
    if (modo === 'secao') {
      const d = st.secao.linha ? montarSecao() : null;
      base = d ? d.cols.map(c => ({ ...c })) : [];
    } else {
      const { pocos, porPoco } = selecao();
      const cand = pocos.filter(p => p.latitude !== null).map(p => ({ p, lat: +p.latitude, lon: +p.longitude }));
      if (!st.redesOcultas.has('Sondagens')) st.soltas.forEach(f => cand.push({ f, lat: f.latitude, lon: f.longitude }));
      base = cand.map(c => Object.assign(colunaSecao(c.p || null, c.f || null), { ll: [c.lat, c.lon], res: c.p ? (porPoco.get(c.p.id) || null) : null }))
        .filter(c => c.camadas.length);
      const comCota = st.pocos.filter(p => p.latitude !== null && p.cota_topo != null);
      base.forEach(c => {
        if (c.cota != null) return;
        c.cotaEstimada = true;
        if (!comCota.length) { c.cota = 0; return; }
        const k = Math.cos(c.ll[0] * Math.PI / 180), d = p => Math.hypot(+p.latitude - c.ll[0], (+p.longitude - c.ll[1]) * k);
        const perto = comCota.reduce((a, b) => d(b) < d(a) ? b : a);
        c.cota = +perto.cota_topo;
      });
      base.sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR', { numeric: true }));
    }
    if (!base.length) return { cols: [], semDatum: false };
    const lat0 = base.reduce((s, c) => s + c.ll[0], 0) / base.length, lon0 = base.reduce((s, c) => s + c.ll[1], 0) / base.length;
    const proj = WGPluma.projecao(lat0, lon0);
    const zRef = Math.max(...base.map(c => c.cota));
    const semDatum = !st.pocos.some(p => p.cota_topo != null);
    base.forEach(c => { const xy = proj.paraXY(c.ll[0], c.ll[1]); c.x = xy.x; c.y = xy.y; c.cotaReal = c.cota; c.cota = c.cota - zRef; });
    return { cols: base, semDatum };
  }

  function montarCerca3D(reenquadrar = false) {
    const { cols, semDatum } = colunasCerca(cerca.modo);
    cerca.cols = cols; cerca.semDatum = semDatum;
    const msg = $('#cerca3d-msg');
    if (!cols.length) {
      msg.hidden = false;
      msg.textContent = cerca.modo === 'secao'
        ? (st.secao.linha ? 'Nenhuma sondagem perto da linha da seção. Aumente a faixa no menu ou use "Todas as sondagens".' : 'Ainda não há seção traçada. Use "Todas as sondagens" ou trace a seção A–A\' no mapa.')
        : 'Nenhuma sondagem com litologia e coordenada neste projeto (e nas redes visíveis).';
      if (visor3d) visor3d.trocarGeometria(WGCerca.geometria([{ nome: '', x: 0, y: 0, cota: 0, prof: 0, camadas: [] }], [], {}), { reenquadrar: true });
      $('#cerca3d-legenda').innerHTML = ''; $('#cerca3d-nota').textContent = '';
      return;
    }
    msg.hidden = true;
    const ext = Math.max(10, Math.max(...cols.map(c => c.x)) - Math.min(...cols.map(c => c.x)), Math.max(...cols.map(c => c.y)) - Math.min(...cols.map(c => c.y)));
    cerca.w = Math.max(0.4, ext / 70);
    WGCerca.espalharColunas(cols, cerca.w);
    const alt = Math.max(1, Math.max(...cols.map(c => c.cota)) - Math.min(...cols.map(c => c.cota - Math.max(c.prof, c.na || 0, 0.5))));
    cerca.ex = cerca.exag || WGCerca.exageroAuto(ext, alt);
    let lig;
    if (cerca.modo === 'secao') {
      const comLito = cols.map((c, i) => i).filter(i => cols[i].camadas.length);
      lig = comLito.slice(1).map((j, q) => [comLito[q], j]);
    } else lig = WGCerca.arvoreMinima(cols);
    const g = WGCerca.geometria(cols, lig, { ex: cerca.ex, w: cerca.w, paineis: cerca.paineis, agua: cerca.agua });
    if (reenquadrar) {
      // seção: de frente para a linha A–A' (A à esquerda), um pouco de lado e de cima; todas: vista oblíqua
      if (cerca.modo === 'secao' && cols.length > 1) {
        const a = cols[0], b = cols[cols.length - 1];
        visor3d.vistaInicial(Math.atan2(b.y - a.y, b.x - a.x) * 180 / Math.PI - 20, 22);
      } else visor3d.vistaInicial(-35, 32);
    }
    visor3d.trocarGeometria(g, { reenquadrar });

    // legenda e nota
    const litos = new Map();
    cols.forEach(c => c.camadas.forEach(r => { const k = WGSecao.chaveLit(r.lit) + '|' + r.hex; if (!litos.has(k)) litos.set(k, { hex: r.hex, texto: [r.lit, r.cor].filter(Boolean).join(' · ') || 'sem descrição' }); }));
    let leg = [...litos.values()].map(l => `<span><i class="leg-lito" style="background:${esc(l.hex)}"></i>${esc(l.texto)}</span>`).join('');
    leg += `<span><i class="leg-terreno"></i>Terreno</span>`;
    if (cerca.agua && cols.some(c => c.na != null)) leg += `<span><svg width="22" height="8"><line x1="0" x2="22" y1="4" y2="4" stroke="#1565A8" stroke-width="2"/></svg>Nível d'água</span>`;
    $('#cerca3d-legenda').innerHTML = leg;
    const camp = st.campanhas.find(c => c.id === st.campanhaId);
    const notas = [`${cols.length} sondagem(ns)`, `${g.paineis} painel(is)`, `exagero vertical ${cerca.ex}x`, `grade de ${fmt(g.caixa.passo, 0)} m`];
    if (cerca.modo === 'todas') notas.push('ligação automática pelo caminho mais curto entre as sondagens (sem cruzar)');
    if (cols.some(c => c.naFonte === 'campanha') && camp) notas.push(`N.A. da campanha ${camp.codigo}`);
    if (semDatum) notas.push('nenhum poço tem cota: todos os topos no mesmo nível');
    else if (cols.some(c => c.cotaEstimada)) notas.push('algumas sondagens sem cota usam a do poço mais próximo');
    notas.push('os painéis são uma interpretação automática (o mesmo critério da seção 2D) e precisam de conferência');
    $('#cerca3d-nota').textContent = notas.join(' · ') + '.';
  }

  function abrirCerca3D(modo) {
    if (!st.projetoId) { toast('Escolha um projeto primeiro.'); return; }
    const novoModo = modo || (st.secao.linha ? 'secao' : 'todas');
    if (novoModo !== cerca.modo) cerca.exag = 0;
    cerca.modo = novoModo;
    $('#cerca3d-modo').value = cerca.modo;
    $('#cerca3d-exag').value = String(cerca.exag);
    $('#cerca3d-paineis').checked = cerca.paineis;
    $('#cerca3d-agua').checked = cerca.agua;
    $('#cerca3d').hidden = false;
    document.body.classList.add('visor-aberto');
    if (!visor3d) {
      visor3d = Visor3D.abrir($('#cerca3d-area'), { aoPassar: dicaCerca, aoSair: () => { $('#cerca3d-dica').hidden = true; } });
      if (!visor3d) { $('#cerca3d-msg').hidden = false; $('#cerca3d-msg').textContent = 'Este navegador não tem WebGL (3D). Tente o Chrome, o Edge ou o Firefox atualizados.'; return; }
    }
    const s = getComputedStyle(document.documentElement).getPropertyValue('--surface').trim();
    visor3d.fundo(WGCerca.rgb(s.length === 7 ? s : '#ffffff'));
    montarCerca3D(true);
  }
  function fecharCerca3D() {
    if (visor3d) { visor3d.destruir(); visor3d = null; }
    $('#cerca3d').hidden = true; $('#cerca3d-dica').hidden = true;
    document.body.classList.remove('visor-aberto');
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
  }

  function dicaCerca(e) {
    const dica = $('#cerca3d-dica');
    const r = $('#cerca3d-area').getBoundingClientRect();
    const h = visor3d && visor3d.colunaEm(e.clientX - r.left, e.clientY - r.top, cerca.cols, cerca.ex, cerca.w);
    if (!h) { dica.hidden = true; return; }
    const c = cerca.cols[h.i];
    const cam = c.camadas.find(x => h.prof >= x.de && h.prof <= x.ate);
    const linhas = [`<b>${esc(c.nome)}</b>${c.rede && c.rede !== '-' ? ' · ' + esc(c.rede) : ''}`];
    if (cam) linhas.push(`${fmt(cam.de)}–${fmt(cam.ate)} m: ${esc([cam.lit, cam.cor].filter(Boolean).join(', ') || 'sem descrição')}`);
    if (!cerca.semDatum) linhas.push(`Cota do topo: ${fmt(c.cotaReal, 2)} m${c.cotaEstimada ? ' (estimada)' : ''}`);
    if (c.prof) linhas.push(`Profundidade: ${fmt(c.prof)} m`);
    if (c.na != null) linhas.push(`N.A.: ${fmt(c.na)} m`);
    dica.innerHTML = linhas.join('<br>');
    dica.hidden = false;
    dica.style.left = Math.min(e.clientX + 14, window.innerWidth - 270) + 'px'; dica.style.top = (e.clientY + 12) + 'px';
  }

  function ligarEventosCerca() {
    $('#btn-cerca3d').addEventListener('click', () => abrirCerca3D());
    $('#btn-secao-3d').addEventListener('click', () => abrirCerca3D('secao'));
    $('#cerca3d-fechar').addEventListener('click', fecharCerca3D);
    $('#cerca3d-modo').addEventListener('change', e => { cerca.modo = e.target.value; cerca.exag = 0; $('#cerca3d-exag').value = '0'; montarCerca3D(true); });
    $('#cerca3d-exag').addEventListener('change', e => { cerca.exag = +e.target.value; montarCerca3D(false); });
    $('#cerca3d-paineis').addEventListener('change', e => { cerca.paineis = e.target.checked; montarCerca3D(false); });
    $('#cerca3d-agua').addEventListener('change', e => { cerca.agua = e.target.checked; montarCerca3D(false); });
    $('#cerca3d-cima').addEventListener('click', () => visor3d && visor3d.vista('cima'));
    $('#cerca3d-inicio').addEventListener('click', () => visor3d && visor3d.vista('inicio'));
    $('#cerca3d-mais').addEventListener('click', () => visor3d && visor3d.zoom(0.8));
    $('#cerca3d-menos').addEventListener('click', () => visor3d && visor3d.zoom(1.25));
    $('#cerca3d-png').addEventListener('click', () => {
      if (!visor3d || !cerca.cols.length) return;
      const a = document.createElement('a');
      const proj = st.projetos.find(p => p.id === st.projetoId);
      a.href = visor3d.imagem(); a.download = `cercas_3d_${(proj?.nome || 'projeto').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w-]+/g, '_')}.png`; a.click();
    });
    $('#cerca3d-tela').addEventListener('click', () => {
      const el = $('#cerca3d');
      if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); else if (el.requestFullscreen) el.requestFullscreen().catch(() => {});
    });
    document.addEventListener('keydown', e => {
      if ($('#cerca3d').hidden || e.target.closest('select,input')) return;
      if (e.key === 'Escape' && !document.fullscreenElement) fecharCerca3D();
      if (!visor3d) return;
      if (e.key === 'ArrowLeft') visor3d.girar(-8);
      if (e.key === 'ArrowRight') visor3d.girar(8);
      if (e.key === '+' || e.key === '=') visor3d.zoom(0.8);
      if (e.key === '-') visor3d.zoom(1.25);
    });
  }

  /** Abas acima do desenho: uma por seção (A–A', B–B'...). O bloco é criado aqui, sem mexer no index.html. */
  function abasSecao() {
    let d = $('#secao-abas');
    if (!d) {
      d = document.createElement('div');
      d.id = 'secao-abas'; d.className = 'secao-abas'; d.setAttribute('role', 'tablist');
      d.style.cssText = 'display:flex;flex-wrap:wrap;gap:6px;align-items:center;padding:10px 18px 0;';
      const corpo = $('#sec-secao .section-body'); corpo.insertBefore(d, corpo.firstChild);
      d.addEventListener('click', e => {
        const b = e.target.closest('[data-i]'); if (b) abrirSecao(+b.dataset.i);
        if (e.target.closest('[data-nova]')) iniciarSecao();
      });
    }
    d.innerHTML = st.secao.lista.map((x, i) => {
      const on = i === st.secao.ativa;
      return `<button type="button" role="tab" aria-selected="${on}" data-i="${i}" class="secao-aba${on ? ' on' : ''}" style="font:600 12.5px/1 'IBM Plex Mono',monospace;padding:6px 11px;border-radius:999px;cursor:pointer;border:1px solid ${on ? 'var(--accent)' : 'var(--line-strong)'};background:${on ? 'var(--accent)' : 'var(--surface)'};color:${on ? '#fff' : 'var(--ink-soft)'};">${nomeSecao(x.letra)}</button>`;
    }).join('') + `<button type="button" data-nova="1" class="secao-aba-nova" style="font:600 12.5px/1 'IBM Plex Sans',sans-serif;padding:6px 11px;border-radius:999px;cursor:pointer;border:1px dashed var(--line-strong);background:transparent;color:var(--ink-soft);">+ Nova seção</button>`;
  }

  /** No relatório: as outras seções (além da aberta), uma depois da outra, logo abaixo da aberta. */
  function secoesExtrasRelatorio(op) {
    let box = $('#rel-secoes-extra');
    if (!box) { box = document.createElement('div'); box.id = 'rel-secoes-extra'; box.className = 'so-relatorio'; $('#sec-secao').after(box); }
    box.innerHTML = '';
    if (!op.secao || st.secao.lista.length < 2) return;
    // em ordem de letra: a A–A' no lugar de sempre, as outras logo abaixo; a aberta volta ao fechar o relatório
    st.secao.ativaAntesRel = st.secao.ativa;
    const blocos = [];
    for (let i = st.secao.lista.length - 1; i >= 0; i--) {
      st.secao.ativa = i; sincronizarSecao(); renderSecao();
      if (i === 0) break;
      const sec = $('#sec-secao').cloneNode(true);
      sec.removeAttribute('id'); sec.hidden = false;
      sec.querySelectorAll('[id]').forEach(el => el.removeAttribute('id'));
      blocos.unshift(sec);
    }
    blocos.forEach(b => box.appendChild(b));
  }

  /** Baixa o desenho da seção em SVG (abre no navegador, no Inkscape, no CorelDRAW e no AutoCAD mais novo). */
  /** Seção em DXF (R12): o mesmo desenho da tela, em metros, com camadas por tipo (litologia, filtro, N.A., pluma...). */
  // ------------------------------------------------------------- ortofoto do drone (georreferenciada) e modo campo (GPS)
  // Sem tabela nova: a ortofoto e as fotos de campo vão para o bucket das fotos, na pasta da empresa:
  //   <empresa>/campo-<projeto>/ortofoto.jpg + ortofoto.json (limites)
  //   <empresa>/campo-<projeto>/<poço>/<data-hora>-<n>.jpg
  const GEOTIFF_URL = 'https://cdn.jsdelivr.net/npm/geotiff@2.1.3/dist-browser/geotiff.js';
  const pastaCampo = () => `${st.org?.id}/campo-${st.projetoId}`;
  let camadaOrto = null;

  function carregarScript(url, global) {
    if (window[global]) return Promise.resolve(window[global]);
    return new Promise((ok, erro) => {
      const s = document.createElement('script'); s.src = url;
      s.onload = () => window[global] ? ok(window[global]) : erro(new Error('biblioteca não carregou'));
      s.onerror = () => erro(new Error('não consegui baixar a biblioteca (confira a internet)'));
      document.head.appendChild(s);
    });
  }

  /** Fuso do projeto (planta, senão poços): para ler ortofoto em UTM sem fuso no arquivo. */
  function fusoDoProjeto() {
    if (st.planta?.dados?.zona) return { zona: st.planta.dados.zona, sul: st.planta.dados.sul };
    const c = st.pocos.filter(p => p.latitude != null && p.longitude != null);
    if (!c.length) return { zona: 23, sul: true };
    const lo = c.reduce((a, p) => a + +p.longitude, 0) / c.length, la = c.reduce((a, p) => a + +p.latitude, 0) / c.length;
    return { zona: WGDxf.zonaDaLongitude(lo), sul: la < 0 };
  }

  /** Código EPSG -> {zona, sul} (SIRGAS 2000 / WGS 84 / SAD69 UTM) ou {graus:true}. null se não reconhece. */
  function fusoDoEpsg(c) {
    if (!c) return null;
    if (c === 4326 || c === 4674 || c === 4618) return { graus: true };
    if (c >= 31965 && c <= 31976) return { zona: c - 31965 + 11, sul: false };     // SIRGAS 2000 UTM N
    if (c >= 31977 && c <= 31985) return { zona: c - 31977 + 17, sul: true };      // SIRGAS 2000 UTM 17S..25S
    if (c >= 32601 && c <= 32660) return { zona: c - 32600, sul: false };
    if (c >= 32701 && c <= 32760) return { zona: c - 32700, sul: true };
    if (c >= 29187 && c <= 29195) return { zona: c - 29170, sul: true, sad69: true }; // SAD69 17S..25S
    return null;
  }

  /**
   * Põe a imagem (já georreferenciada) no "chão" do mapa: desenha num canvas alinhado ao mapa (Web Mercator),
   * com o canto (0,0), (W,0) e (0,H) da imagem em coordenadas [lat, lon]. Devolve {canvas, bounds}.
   */
  function entortarOrto(fonte, W, H, c00, cW0, c0H, maxLado = 4096) {
    const crs = L.CRS.EPSG3857, pj = ll => crs.project(L.latLng(ll[0], ll[1]));
    const p0 = pj(c00), p1 = pj(cW0), p2 = pj(c0H), p3 = { x: p1.x + p2.x - p0.x, y: p1.y + p2.y - p0.y };
    const xs = [p0.x, p1.x, p2.x, p3.x], ys = [p0.y, p1.y, p2.y, p3.y];
    const x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys);
    const k = maxLado / Math.max(x1 - x0, y1 - y0);
    const cv = document.createElement('canvas'); cv.width = Math.max(1, Math.round((x1 - x0) * k)); cv.height = Math.max(1, Math.round((y1 - y0) * k));
    const T = p => ({ x: (p.x - x0) * k, y: (y1 - p.y) * k });       // y do canvas cresce para baixo
    const a0 = T(p0), a1 = T(p1), a2 = T(p2);
    const ctx = cv.getContext('2d'); ctx.imageSmoothingQuality = 'high';
    ctx.setTransform((a1.x - a0.x) / W, (a1.y - a0.y) / W, (a2.x - a0.x) / H, (a2.y - a0.y) / H, a0.x, a0.y);
    ctx.drawImage(fonte, 0, 0, W, H);
    const sw = crs.unproject(L.point(x0, y0)), ne = crs.unproject(L.point(x1, y1));
    return { canvas: cv, bounds: [[sw.lat, sw.lng], [ne.lat, ne.lng]] };
  }

  /** Lê a ortofoto: GeoTIFF, ou JPG/PNG + arquivo de georreferência (.jgw/.pgw/.tfw/.wld). */
  async function lerOrtofoto(arquivos) {
    const img = arquivos.find(f => /\.(tif|tiff|jpe?g|png)$/i.test(f.name));
    const mundo = arquivos.find(f => /\.(jgw|jpgw|pgw|pngw|tfw|tifw|wld)$/i.test(f.name));
    if (!img) throw new Error('escolha a imagem (GeoTIFF, JPG ou PNG) — junto com o arquivo .jgw/.pgw se for JPG/PNG');
    const msg = t => { $('#orto-estado').hidden = false; $('#orto-estado').textContent = t; };
    let fonte, W, H, cantos, fuso = fusoDoProjeto(), aviso = '';
    const paraLL = (x, y) => fuso.graus ? [y, x] : (r => [r.lat, r.lon])(WGDxf.utmParaLatLon(x, y, fuso.zona, fuso.sul));
    if (/\.tiff?$/i.test(img.name) && !mundo) {
      msg('Lendo o GeoTIFF...');
      const GT = await carregarScript(GEOTIFF_URL, 'GeoTIFF');
      const tiff = await GT.fromArrayBuffer(await img.arrayBuffer()), im = await tiff.getImage();
      const gk = im.getGeoKeys ? im.getGeoKeys() : {};
      const f2 = fusoDoEpsg(gk.ProjectedCSTypeGeoKey) || (gk.ProjectedCSTypeGeoKey ? null : fusoDoEpsg(gk.GeographicTypeGeoKey));
      if (f2) { fuso = f2.graus ? { graus: true } : f2; if (f2.sad69) aviso = ' (atenção: SAD69 — pode ficar deslocada uns 50 m)'; }
      else if (gk.ProjectedCSTypeGeoKey) aviso = ` (sistema EPSG ${gk.ProjectedCSTypeGeoKey} não reconhecido: usado UTM ${fuso.zona}${fuso.sul ? 'S' : 'N'})`;
      const [ox, oy] = im.getOrigin(), [rx, ry] = im.getResolution(), w0 = im.getWidth(), h0 = im.getHeight();
      const f = Math.min(1, 4096 / Math.max(w0, h0));
      W = Math.max(1, Math.round(w0 * f)); H = Math.max(1, Math.round(h0 * f));
      msg(`Lendo o GeoTIFF (${w0} x ${h0} px)...`);
      const rgb = await im.readRGB({ width: W, height: H, interleave: true, enableAlpha: true });
      const n = rgb.length / (W * H), cv = document.createElement('canvas'); cv.width = W; cv.height = H;
      const cx = cv.getContext('2d'), id = cx.createImageData(W, H);
      for (let i = 0; i < W * H; i++) {
        const r = rgb[i * n], g = rgb[i * n + 1], b = rgb[i * n + 2], a = n >= 4 ? rgb[i * n + 3] : 255;
        id.data[i * 4] = r; id.data[i * 4 + 1] = g; id.data[i * 4 + 2] = b; id.data[i * 4 + 3] = (r === 0 && g === 0 && b === 0 && n < 4) ? 0 : a;
      }
      cx.putImageData(id, 0, 0); fonte = cv;
      cantos = [paraLL(ox, oy), paraLL(ox + rx * w0, oy), paraLL(ox, oy + ry * h0)];
    } else {
      if (!mundo) throw new Error('falta o arquivo de georreferência da imagem (.jgw para JPG, .pgw para PNG, .tfw para TIF)');
      const v = (await mundo.text()).split(/\s+/).filter(Boolean).map(t => WGImport.numero(t));
      if (v.length < 6 || v.some(x => x === null)) throw new Error('o arquivo ' + mundo.name + ' não parece um world file (6 números)');
      const [A, D, B, E, C, F] = v;
      if (Math.abs(C) <= 180 && Math.abs(F) <= 90) fuso = { graus: true };
      msg('Lendo a imagem...');
      const url = URL.createObjectURL(img);
      fonte = await new Promise((ok, erro) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => erro(new Error('imagem grande demais ou inválida — reduza para até 8.000 px')); i.src = url; });
      W = fonte.naturalWidth; H = fonte.naturalHeight;
      const X = (c, r) => [A * c + B * r + C - (A + B) / 2, D * c + E * r + F - (D + E) / 2];   // world file: centro do pixel
      cantos = [paraLL(...X(0, 0)), paraLL(...X(W, 0)), paraLL(...X(0, H))];
    }
    msg('Ajustando a imagem ao mapa...');
    const r = entortarOrto(fonte, W, H, ...cantos);
    return { ...r, aviso, fuso };
  }

  async function inserirOrtofoto(arquivos) {
    if (!st.projetoId || !st.org?.id) { toast('Escolha um projeto primeiro.'); return; }
    try {
      const o = await lerOrtofoto([...arquivos]);
      $('#orto-estado').textContent = 'Enviando a ortofoto...';
      const blob = await new Promise((ok, erro) => o.canvas.toBlob(b => b ? ok(b) : erro(new Error('não consegui preparar a imagem')), 'image/jpeg', .85));
      const meta = { bounds: o.bounds, nome: [...arquivos].find(f => /\.(tif|tiff|jpe?g|png)$/i.test(f.name)).name, enviado: new Date().toISOString() };
      const b1 = await sb.storage.from(BUCKET_FOTOS).upload(pastaCampo() + '/ortofoto.jpg', blob, { upsert: true, contentType: 'image/jpeg' });
      if (b1.error) throw b1.error;
      const b2 = await sb.storage.from(BUCKET_FOTOS).upload(pastaCampo() + '/ortofoto.json', new Blob([JSON.stringify(meta)], { type: 'application/json' }), { upsert: true, contentType: 'application/json' });
      if (b2.error) throw b2.error;
      st.orto = { ...meta, url: URL.createObjectURL(blob) };
      prefsMapa.orto = true; guardarPrefsMapa();
      renderOrto(true);
      toast('Ortofoto inserida no mapa' + (o.aviso || '') + '.');
    } catch (e) {
      console.error(e);
      $('#orto-estado').hidden = false;
      $('#orto-estado').textContent = 'Não consegui usar a ortofoto: ' + (e.message || e);
      return;
    }
    mostrarEstadoOrto();
  }

  async function carregarOrto() {
    st.orto = null;
    if (!st.projetoId || !st.org?.id) { renderOrto(); return; }
    try {
      const { data } = await sb.storage.from(BUCKET_FOTOS).createSignedUrl(pastaCampo() + '/ortofoto.json', 3600);
      if (data?.signedUrl) {
        const r = await fetch(data.signedUrl);
        if (r.ok) {
          const meta = await r.json();
          const img = await sb.storage.from(BUCKET_FOTOS).createSignedUrl(pastaCampo() + '/ortofoto.jpg', 3600 * 8);
          if (img.data?.signedUrl && meta?.bounds) st.orto = { ...meta, url: img.data.signedUrl };
        }
      }
    } catch (e) { /* sem ortofoto */ }
    renderOrto(); mostrarEstadoOrto();
  }

  function renderOrto(enquadrar = false) {
    if (!mapa) return;
    if (camadaOrto) { mapa.removeLayer(camadaOrto); camadaOrto = null; }
    document.querySelectorAll('.mapa-botoes button[data-k="orto"]').forEach(b => { b.hidden = !st.orto; b.setAttribute('aria-pressed', String(!!prefsMapa.orto)); });
    if (!st.orto || prefsMapa.orto === false) return;
    camadaOrto = L.imageOverlay(st.orto.url, st.orto.bounds, { pane: 'orto', interactive: false, className: 'img-orto' }).addTo(mapa);
    if (enquadrar) mapa.fitBounds(st.orto.bounds, { padding: [20, 20] });
  }

  function mostrarEstadoOrto() {
    const el = $('#orto-estado'); if (!el) return;
    el.hidden = !st.orto;
    if (st.orto) el.innerHTML = `Ortofoto: <b>${esc(st.orto.nome || 'imagem')}</b> · <button type="button" class="btn-link" id="orto-remover">remover</button>`;
    const rm = $('#orto-remover');
    if (rm) rm.addEventListener('click', async () => {
      if (!confirm('Remover a ortofoto deste projeto?')) return;
      await sb.storage.from(BUCKET_FOTOS).remove([pastaCampo() + '/ortofoto.jpg', pastaCampo() + '/ortofoto.json']);
      st.orto = null; renderOrto(); mostrarEstadoOrto(); toast('Ortofoto removida.');
    });
  }

  // ---- desenho e medida no mapa (reta, mão livre, medir, borracha)
  // Sem tabela nova: os desenhos do projeto ficam num JSON no bucket das fotos, na pasta do projeto:
  //   <empresa>/campo-<projeto>/desenhos.json  = [{ id, tipo: 'reta'|'livre'|'medida', pts: [[lat,lon]...], cor, esp, fechado }]
  let camadaDes = null, desSalvarT = null;
  const des = { lista: [], ativo: false, ferr: 'reta', cor: '#E0261B', esp: 4, atual: null, mouse: null, temp: null, desfazer: [], livre: null };
  const fmtDist = m => m >= 1000 ? (m / 1000).toLocaleString('pt-BR', { maximumFractionDigits: 3 }) + ' km'
    : m.toLocaleString('pt-BR', { minimumFractionDigits: m < 100 ? 2 : 1, maximumFractionDigits: m < 100 ? 2 : 1 }) + ' m';
  const fmtArea = a => a >= 10000
    ? `${(a / 10000).toLocaleString('pt-BR', { maximumFractionDigits: 3 })} ha (${Math.round(a).toLocaleString('pt-BR')} m²)`
    : a.toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + ' m²';

  /** Comprimento de cada trecho (m), em UTM do projeto, e a área (m²) se a figura for fechada. */
  function medirPts(pts, fechado) {
    const { zona, sul } = fusoDoProjeto();
    const u = pts.map(([la, lo]) => { const q = WGDxf.latLonParaUtm(la, lo, zona, sul); return [q.E, q.N]; });
    const n = u.length, trechos = [];
    for (let i = 1; i < n; i++) trechos.push(Math.hypot(u[i][0] - u[i - 1][0], u[i][1] - u[i - 1][1]));
    if (fechado && n > 2) trechos.push(Math.hypot(u[0][0] - u[n - 1][0], u[0][1] - u[n - 1][1]));
    let area = 0;
    if (fechado && n > 2) { for (let i = 0; i < n; i++) { const a = u[i], b = u[(i + 1) % n]; area += a[0] * b[1] - b[0] * a[1]; } area = Math.abs(area) / 2; }
    return { trechos, total: trechos.reduce((a, b) => a + b, 0), area };
  }

  async function carregarDesenhos() {
    des.lista = []; des.desfazer = [];
    if (st.projetoId && st.org?.id) {
      try {
        const { data } = await sb.storage.from(BUCKET_FOTOS).createSignedUrl(pastaCampo() + '/desenhos.json', 600);
        if (data?.signedUrl) { const r = await fetch(data.signedUrl, { cache: 'no-store' }); if (r.ok) { const j = await r.json(); if (Array.isArray(j)) des.lista = j.filter(d => d && Array.isArray(d.pts) && d.pts.length > 1); } }
      } catch (e) { /* sem desenhos */ }
    }
    renderDesenhos();
  }

  function salvarDesenhos() {
    clearTimeout(desSalvarT);
    desSalvarT = setTimeout(async () => {
      if (!st.projetoId || !st.org?.id) return;
      const corpo = new Blob([JSON.stringify(des.lista)], { type: 'application/json' });
      const { error } = await sb.storage.from(BUCKET_FOTOS).upload(pastaCampo() + '/desenhos.json', corpo, { upsert: true, contentType: 'application/json' });
      if (error) toast('Não consegui salvar os desenhos: ' + (error.message || error));
    }, 500);
  }

  function rotuloMedida(ll, txt, classe = '') {
    return L.marker(ll, { pane: 'desenho', interactive: false, keyboard: false,
      icon: L.divIcon({ className: 'rotulo-medida ' + classe, html: `<span>${esc(txt)}</span>`, iconSize: [0, 0] }) });
  }

  /** Desenha um item (salvo ou em andamento) num grupo: linha/polígono e, na medida, os rótulos. */
  function desenharItemDes(d, grupo, previa = false) {
    const fech = d.fechado && d.pts.length > 2;
    const op = { pane: 'desenho', color: d.cor, weight: d.esp, opacity: .95, lineCap: 'round', lineJoin: 'round',
      interactive: !previa && des.ativo && des.ferr === 'apagar', className: 'des-item', bubblingMouseEvents: false };
    const forma = fech ? L.polygon(d.pts, { ...op, fill: true, fillColor: d.cor, fillOpacity: d.tipo === 'medida' ? .12 : .06 })
      : L.polyline(d.pts, { ...op, dashArray: d.tipo === 'medida' ? `${d.esp * 2.5} ${d.esp * 1.6}` : null });
    if (d.tipo === 'medida' || previa) {   // contorno claro por baixo, para ler em cima do satélite
      L.polyline(fech ? [...d.pts, d.pts[0]] : d.pts, { pane: 'desenho', color: '#ffffff', weight: d.esp + 2.5, opacity: .55, interactive: false }).addTo(grupo);
    }
    forma.addTo(grupo);
    if (op.interactive) forma.on('click', () => apagarDesenho(d.id));
    if (d.tipo !== 'medida') return;
    d.pts.forEach(p => L.circleMarker(p, { pane: 'desenho', radius: Math.max(2.5, d.esp * .8), color: '#fff', weight: 1.5, fillColor: d.cor, fillOpacity: 1, interactive: false }).addTo(grupo));
    const m = medirPts(d.pts, fech), n = d.pts.length;
    if (m.trechos.length > 1) m.trechos.forEach((t, i) => {
      const a = d.pts[i], b = d.pts[(i + 1) % n];
      rotuloMedida([(a[0] + b[0]) / 2, (a[1] + b[1]) / 2], fmtDist(t)).addTo(grupo);
    });
    const ult = fech ? d.pts[0] : d.pts[n - 1];
    rotuloMedida(ult, (fech ? 'Perímetro ' : m.trechos.length > 1 ? 'Total ' : '') + fmtDist(m.total), 'total').addTo(grupo);
    if (fech) {
      const c = d.pts.reduce((a, p) => [a[0] + p[0] / n, a[1] + p[1] / n], [0, 0]);
      rotuloMedida(c, 'Área ' + fmtArea(m.area), 'area').addTo(grupo);
    }
  }

  function renderDesenhos() {
    if (!mapa || !camadaDes) return;
    camadaDes.clearLayers();
    const mostrar = prefsMapa.desenhos !== false || des.ativo;
    if (mostrar) des.lista.forEach(d => desenharItemDes(d, camadaDes));
    const linha = $('#f-desenhos-linha');
    if (linha) { linha.hidden = !des.lista.length; $('#des-conta').textContent = des.lista.length; $('#f-desenhos').checked = prefsMapa.desenhos !== false; }
    previaDes();
  }

  /** Traço em andamento + linha elástica até o mouse + medida ao vivo na barra. */
  function previaDes() {
    if (!des.temp) return;
    des.temp.clearLayers();
    const a = des.atual, out = $('#des-medida');
    if (!a || !a.pts.length) { if (out) out.textContent = ''; return; }
    const pts = des.mouse && a.tipo !== 'livre' ? [...a.pts, des.mouse] : a.pts;
    if (pts.length > 1) desenharItemDes({ ...a, pts, fechado: false }, des.temp, true);
    else L.circleMarker(pts[0], { pane: 'desenho', radius: 4, color: '#fff', weight: 1.5, fillColor: a.cor, fillOpacity: 1, interactive: false }).addTo(des.temp);
    if (a.tipo !== 'livre' && a.pts.length > 2) L.circleMarker(a.pts[0], { pane: 'desenho', radius: 7, color: a.cor, weight: 2, fill: false, interactive: false }).addTo(des.temp);
    if (out) {
      const m = medirPts(pts, false), ult = m.trechos[m.trechos.length - 1] || 0;
      out.textContent = pts.length > 1 ? (m.trechos.length > 1 ? `trecho ${fmtDist(ult)} · total ${fmtDist(m.total)}` : fmtDist(m.total)) : '';
    }
  }

  const DICAS_DES = {
    reta: 'Reta: clique ponto a ponto no mapa; duplo clique (ou Concluir traço) termina. Clicar no 1º ponto fecha a figura.',
    livre: 'Mão livre: segure e arraste no mapa para desenhar (o mapa não se move nesse modo).',
    medida: 'Medir: clique ponto a ponto; mostra cada trecho e o total. Clicar no 1º ponto fecha e mostra a área.',
    apagar: 'Borracha: clique num desenho para apagá-lo (Desfazer traz de volta).'
  };

  function escolherFerramenta(f) {
    concluirTraco();
    des.ferr = f;
    document.querySelectorAll('#des-barra [data-ferr]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.ferr === f)));
    $('#des-dica').textContent = DICAS_DES[f];
    const el = mapa.getContainer();
    el.classList.toggle('mapa-desenhando', f !== 'apagar');
    el.classList.toggle('mapa-borracha', f === 'apagar');
    el.classList.toggle('mapa-livre', f === 'livre');
    if (f === 'livre') mapa.dragging.disable(); else mapa.dragging.enable();
    renderDesenhos();
  }

  function abrirDesenho() {
    if (!mapa) return;
    if (!st.projetoId) { toast('Escolha um projeto primeiro.'); return; }
    if (des.ativo) { fecharDesenho(); return; }
    if (st.ajuste || st.secao.tracando || st.f360.marcar || st.editPoco || st.rotAjuste || st.editLim) { toast('Termine o que está fazendo no mapa antes.'); return; }
    des.ativo = true; st.desAtivo = true;
    mapa.closePopup(); mapa.doubleClickZoom.disable();
    des.temp = L.layerGroup().addTo(mapa);
    $('#des-barra').hidden = false;
    $('#btn-desenho').classList.add('ativo');
    escolherFerramenta(des.ferr);
    mapa.invalidateSize();
  }

  function fecharDesenho() {
    concluirTraco();
    des.ativo = false; st.desAtivo = false; des.mouse = null;
    if (des.temp) { mapa.removeLayer(des.temp); des.temp = null; }
    $('#des-barra').hidden = true;
    $('#btn-desenho').classList.remove('ativo');
    const el = mapa.getContainer(); el.classList.remove('mapa-desenhando', 'mapa-borracha', 'mapa-livre');
    mapa.dragging.enable(); mapa.doubleClickZoom.enable();
    renderDesenhos(); mapa.invalidateSize();
  }

  /** Fecha o traço em andamento e guarda (tira pontos repetidos do duplo clique). */
  function concluirTraco(fechar = false) {
    const a = des.atual; des.atual = null; des.mouse = null; $('#des-fim').hidden = true;
    if (a && a.pts.length > 1) {
      const pts = a.pts.filter((p, i) => !i || mapa.latLngToContainerPoint(p).distanceTo(mapa.latLngToContainerPoint(a.pts[i - 1])) > 2);
      if (pts.length > 1) {
        const d = { id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6), tipo: a.tipo, pts: pts.map(p => [+p[0].toFixed(7), +p[1].toFixed(7)]), cor: a.cor, esp: a.esp, fechado: fechar && pts.length > 2 };
        des.lista.push(d); des.desfazer.push({ acao: 'add', id: d.id }); salvarDesenhos();
      }
    }
    $('#des-medida').textContent = '';
    renderDesenhos();
  }

  function apagarDesenho(id) {
    const i = des.lista.findIndex(d => d.id === id); if (i < 0) return;
    const [d] = des.lista.splice(i, 1); des.desfazer.push({ acao: 'del', item: d, i });
    salvarDesenhos(); renderDesenhos();
  }

  function desfazerDes() {
    if (des.atual && des.atual.pts.length) { des.atual.pts.pop(); if (!des.atual.pts.length) concluirTraco(); else previaDes(); return; }
    const u = des.desfazer.pop(); if (!u) { toast('Nada para desfazer.'); return; }
    if (u.acao === 'add') des.lista = des.lista.filter(d => d.id !== u.id);
    else des.lista.splice(Math.min(u.i, des.lista.length), 0, u.item);
    salvarDesenhos(); renderDesenhos();
  }

  function cliqueDes(e) {
    if (!des.ativo || (des.ferr !== 'reta' && des.ferr !== 'medida')) return;
    const p = [e.latlng.lat, e.latlng.lng];
    if (!des.atual) des.atual = { tipo: des.ferr, cor: des.cor, esp: des.esp, pts: [] };
    const a = des.atual;
    if (a.pts.length > 2 && mapa.latLngToContainerPoint(a.pts[0]).distanceTo(e.containerPoint) < 12) { concluirTraco(true); return; }
    a.pts.push(p); $('#des-fim').hidden = false; previaDes();
  }

  function ligarMapaDesenho() {
    mapa.createPane('desenho').style.zIndex = 392;   // acima da seção, abaixo dos poços
    camadaDes = L.layerGroup().addTo(mapa);
    mapa.on('click', cliqueDes);
    mapa.on('dblclick', () => { if (des.ativo && des.atual) concluirTraco(); });
    mapa.on('mousemove', e => { if (des.ativo && des.atual && des.ferr !== 'livre') { des.mouse = [e.latlng.lat, e.latlng.lng]; previaDes(); } });
    // mão livre: ponteiro (mouse, dedo ou caneta) direto no contêiner do mapa
    const el = mapa.getContainer();
    const fora = ev => ev.target.closest('.leaflet-control, .mapa-botoes, .leaflet-popup');
    el.addEventListener('pointerdown', ev => {
      if (!des.ativo || des.ferr !== 'livre' || fora(ev) || (ev.pointerType === 'mouse' && ev.button !== 0)) return;
      ev.preventDefault();
      try { el.setPointerCapture(ev.pointerId); } catch (e) { /* ok */ }
      const ll = mapa.mouseEventToLatLng(ev);
      des.livre = { id: ev.pointerId, ultimo: mapa.mouseEventToContainerPoint(ev) };
      des.atual = { tipo: 'livre', cor: des.cor, esp: des.esp, pts: [[ll.lat, ll.lng]] }; previaDes();
    });
    el.addEventListener('pointermove', ev => {
      if (!des.livre || ev.pointerId !== des.livre.id || !des.atual) return;
      const cp = mapa.mouseEventToContainerPoint(ev); if (cp.distanceTo(des.livre.ultimo) < 3) return;
      des.livre.ultimo = cp; const ll = mapa.containerPointToLatLng(cp);
      des.atual.pts.push([ll.lat, ll.lng]); previaDes();
    });
    const soltar = ev => {
      if (!des.livre || ev.pointerId !== des.livre.id) return;
      des.livre = null;
      if (des.atual) des.atual.pts = simplificarTela(des.atual.pts, 1.2);
      concluirTraco();
    };
    el.addEventListener('pointerup', soltar); el.addEventListener('pointercancel', soltar);
    document.addEventListener('keydown', ev => {
      if (!des.ativo || /input|textarea|select/i.test(ev.target.tagName)) return;
      if (ev.key === 'Enter') concluirTraco();
      else if (ev.key === 'Escape') { if (des.atual) { des.atual = null; des.mouse = null; $('#des-fim').hidden = true; previaDes(); } else fecharDesenho(); }
      else if ((ev.ctrlKey || ev.metaKey) && ev.key.toLowerCase() === 'z') { ev.preventDefault(); desfazerDes(); }
    });
  }

  /** Douglas-Peucker em pixels da tela (traço à mão livre fica leve sem perder a forma). */
  function simplificarTela(pts, tol) {
    if (pts.length < 3) return pts;
    const q = pts.map(p => mapa.latLngToContainerPoint(p)), manter = new Uint8Array(pts.length);
    manter[0] = manter[pts.length - 1] = 1;
    const pilha = [[0, pts.length - 1]];
    while (pilha.length) {
      const [i, j] = pilha.pop(); let dm = 0, k = -1;
      const ax = q[i].x, ay = q[i].y, bx = q[j].x, by = q[j].y, L2 = (bx - ax) ** 2 + (by - ay) ** 2;
      for (let m = i + 1; m < j; m++) {
        let t = L2 ? ((q[m].x - ax) * (bx - ax) + (q[m].y - ay) * (by - ay)) / L2 : 0; t = Math.max(0, Math.min(1, t));
        const d = Math.hypot(q[m].x - (ax + t * (bx - ax)), q[m].y - (ay + t * (by - ay)));
        if (d > dm) { dm = d; k = m; }
      }
      if (dm > tol && k > 0) { manter[k] = 1; pilha.push([i, k], [k, j]); }
    }
    return pts.filter((_, i) => manter[i]);
  }

  function ligarEventosDesenho() {
    $('#btn-desenho').addEventListener('click', abrirDesenho);
    $('#des-fechar').addEventListener('click', fecharDesenho);
    $('#des-fim').addEventListener('click', () => concluirTraco());
    $('#des-desfazer').addEventListener('click', desfazerDes);
    $('#des-limpar').addEventListener('click', () => {
      if (!des.lista.length) return;
      if (!confirm(`Apagar os ${des.lista.length} desenho(s) deste projeto?`)) return;
      des.desfazer.push(...des.lista.map((item, i) => ({ acao: 'del', item, i })).reverse());
      des.lista = []; salvarDesenhos(); renderDesenhos();
    });
    document.querySelectorAll('#des-barra [data-ferr]').forEach(b => b.addEventListener('click', () => escolherFerramenta(b.dataset.ferr)));
    const marcarCor = c => {
      des.cor = c; if (des.atual) { des.atual.cor = c; previaDes(); }
      document.querySelectorAll('#des-barra .des-cor[data-cor]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.cor.toLowerCase() === c.toLowerCase())));
    };
    document.querySelectorAll('#des-barra .des-cor[data-cor]').forEach(b => b.addEventListener('click', () => marcarCor(b.dataset.cor)));
    $('#des-cor-livre').addEventListener('input', e => marcarCor(e.target.value));
    $('#des-esp').addEventListener('change', e => { des.esp = +e.target.value; if (des.atual) { des.atual.esp = des.esp; previaDes(); } });
    $('#f-desenhos').addEventListener('change', e => { prefsMapa.desenhos = e.target.checked; guardarPrefsMapa(); renderDesenhos(); });
  }

  // ---- modo campo: GPS, navegação até o poço e ficha de campo
  const campo = { ativo: false, watch: null, pos: null, prec: null, rumo: null, alvo: null, camada: null, seguir: true, ultimaLista: 0 };

  async function iniciarCampo() {
    if (!st.projetoId) { toast('Escolha um projeto primeiro.'); return; }
    if (!('geolocation' in navigator)) { toast('Este aparelho/navegador não tem GPS disponível.'); return; }
    campo.ativo = true; campo.seguir = true;
    document.body.classList.add('modo-campo');
    $('#campo-painel').hidden = false;
    trocarView('visao');
    setTimeout(() => mapa.invalidateSize(), 60);
    campo.camada = L.layerGroup().addTo(mapa);
    // bússola (no iPhone precisa de permissão, pedida aqui porque veio de um toque)
    try { if (window.DeviceOrientationEvent && typeof DeviceOrientationEvent.requestPermission === 'function') await DeviceOrientationEvent.requestPermission(); } catch (e) { /* sem bússola */ }
    window.addEventListener('deviceorientationabsolute', orientacaoCampo, true);
    window.addEventListener('deviceorientation', orientacaoCampo, true);
    $('#campo-gps').textContent = 'Procurando o GPS...';
    campo.watch = navigator.geolocation.watchPosition(posicaoCampo, e => {
      $('#campo-gps').textContent = e.code === 1 ? 'GPS bloqueado: libere a localização para este site nas configurações do navegador.' : 'Sem sinal de GPS (' + e.message + ')';
    }, { enableHighAccuracy: true, maximumAge: 2000, timeout: 30000 });
    renderListaCampo();
  }

  function sairCampo() {
    campo.ativo = false;
    if (campo.watch != null) navigator.geolocation.clearWatch(campo.watch);
    campo.watch = null;
    window.removeEventListener('deviceorientationabsolute', orientacaoCampo, true);
    window.removeEventListener('deviceorientation', orientacaoCampo, true);
    if (campo.camada) mapa.removeLayer(campo.camada);
    campo.camada = null; campo.alvo = null;
    document.body.classList.remove('modo-campo');
    $('#campo-painel').hidden = true;
    setTimeout(() => mapa.invalidateSize(), 60);
  }

  function orientacaoCampo(e) {
    let h = null;
    if (typeof e.webkitCompassHeading === 'number') h = e.webkitCompassHeading;          // iPhone
    else if ((e.absolute || e.type === 'deviceorientationabsolute') && e.alpha != null) h = (360 - e.alpha) % 360;
    if (h == null) return;
    campo.rumo = h;
    atualizarAlvoCampo();
  }

  function posicaoCampo(p) {
    campo.pos = [p.coords.latitude, p.coords.longitude]; campo.prec = p.coords.accuracy;
    $('#campo-gps').textContent = `GPS ±${fmt(campo.prec, 0)} m`;
    $('#campo-gps').classList.toggle('ruim', campo.prec > 15);
    desenharEuCampo();
    if (campo.seguir) mapa.panTo(campo.pos, { animate: true });
    atualizarAlvoCampo();
    if (Date.now() - campo.ultimaLista > 2500) renderListaCampo();
  }

  function desenharEuCampo() {
    if (!campo.camada) return;
    campo.camada.clearLayers();
    if (!campo.pos) return;
    L.circle(campo.pos, { radius: campo.prec || 5, color: '#1E88E5', weight: 1, fillColor: '#1E88E5', fillOpacity: .12, interactive: false }).addTo(campo.camada);
    if (campo.alvo) {
      const p = st.pocos.find(x => x.id === campo.alvo);
      if (p && p.latitude != null) L.polyline([campo.pos, [+p.latitude, +p.longitude]], { color: '#FFD60A', weight: 4, dashArray: '8 6', interactive: false }).addTo(campo.camada);
    }
    L.circleMarker(campo.pos, { radius: 8, color: '#fff', weight: 3, fillColor: '#1E88E5', fillOpacity: 1, interactive: false }).addTo(campo.camada);
  }

  /** Rumo (0 = norte) e distância (m) de A até B. */
  function rumoDist(a, b) {
    const d = mapa.distance(a, b), la = (a[0] + b[0]) / 2 * Math.PI / 180;
    const dn = (b[0] - a[0]) * 110574, de = (b[1] - a[1]) * 111320 * Math.cos(la);
    return { d, rumo: (Math.atan2(de, dn) * 180 / Math.PI + 360) % 360 };
  }
  const nomeRumo = g => ['N', 'NE', 'L', 'SE', 'S', 'SO', 'O', 'NO'][Math.round(g / 45) % 8];

  function atualizarAlvoCampo() {
    const box = $('#campo-alvo');
    const p = campo.alvo != null ? st.pocos.find(x => x.id === campo.alvo) : null;
    box.hidden = !p;
    if (!p) return;
    $('#campo-alvo-nome').textContent = p.codigo;
    if (!campo.pos || p.latitude == null) { $('#campo-alvo-dist').textContent = p.latitude == null ? 'poço sem coordenada' : 'esperando o GPS...'; return; }
    const { d, rumo } = rumoDist(campo.pos, [+p.latitude, +p.longitude]);
    const chegou = d <= Math.max(5, Math.min(campo.prec || 5, 15));
    $('#campo-alvo-dist').textContent = chegou ? 'Você chegou!' : `${d < 1000 ? fmt(d, d < 20 ? 1 : 0) + ' m' : fmt(d / 1000, 2) + ' km'} · rumo ${nomeRumo(rumo)} (${Math.round(rumo)}°)`;
    box.classList.toggle('chegou', chegou);
    // seta: com bússola, relativa para onde o celular aponta; sem bússola, em relação ao norte do mapa
    const ang = campo.rumo != null ? rumo - campo.rumo : rumo;
    $('#campo-seta').style.transform = `rotate(${ang}deg)`;
    $('#campo-seta-nota').textContent = campo.rumo != null ? 'siga a seta' : 'seta em relação ao norte (sem bússola)';
  }

  function renderListaCampo() {
    campo.ultimaLista = Date.now();
    const feitos = new Set(st.medicoes.filter(m => m.campanha_id === st.campanhaId).map(m => m.poco_id));
    const lista = st.pocos.filter(p => !p.virtual).map(p => ({ p, d: campo.pos && p.latitude != null ? mapa.distance(campo.pos, [+p.latitude, +p.longitude]) : null }))
      .sort((a, b) => (a.d ?? 1e12) - (b.d ?? 1e12) || a.p.codigo.localeCompare(b.p.codigo, 'pt-BR', { numeric: true }));
    $('#campo-conta').textContent = `${feitos.size} de ${lista.length} com N.A. nesta campanha`;
    $('#campo-lista').innerHTML = lista.map(({ p, d }) => `<li class="${feitos.has(p.id) ? 'feito' : ''}${campo.alvo === p.id ? ' alvo' : ''}">
      <button type="button" data-ir="${p.id}"><b>${esc(p.codigo)}</b><span>${d != null ? (d < 1000 ? fmt(d, 0) + ' m' : fmt(d / 1000, 1) + ' km') : (p.latitude == null ? 'sem coordenada' : '')}</span>
      <em>${feitos.has(p.id) ? '✓ medido' : esc((ROTULO_SITUACAO[p.situacao] || '').toLowerCase())}</em></button>
      <button type="button" class="campo-ficha-btn" data-ficha="${p.id}">Ficha</button></li>`).join('');
  }

  function irParaPoco(id) {
    campo.alvo = id;
    const p = st.pocos.find(x => x.id === id);
    if (p && p.latitude != null) {
      if (campo.pos) mapa.fitBounds(L.latLngBounds([campo.pos, [+p.latitude, +p.longitude]]), { padding: [60, 60], maxZoom: 20 });
      else mapa.setView([+p.latitude, +p.longitude], 19);
    }
    campo.seguir = false;
    desenharEuCampo(); atualizarAlvoCampo(); renderListaCampo();
  }

  // ---- ficha de campo
  async function abrirFichaCampo(id) {
    const p = st.pocos.find(x => x.id === id); if (!p) return;
    campo.ficha = { p, fotos: [] };
    const med = st.medicoes.find(m => m.poco_id === p.id && m.campanha_id === st.campanhaId);
    const camp = st.campanhas.find(c => c.id === st.campanhaId);
    $('#cf-titulo').textContent = `Ficha de campo · ${p.codigo}`;
    $('#cf-campanha').textContent = camp ? `Campanha ${camp.codigo}` : 'Sem campanha escolhida: o N.A. não será gravado';
    $('#cf-achado').value = 'sim';
    $('#cf-situacao').value = ['ATIVO', 'SECO', 'DANIFICADO', 'DESATIVADO'].includes(p.situacao) ? p.situacao : 'ATIVO';
    $('#cf-na').value = med?.nivel_agua != null ? String(med.nivel_agua).replace('.', ',') : '';
    $('#cf-fl').checked = !!med?.fase_livre; $('#cf-esp').value = med?.espessura_fl != null ? String(med.espessura_fl).replace('.', ',') : '';
    $('#cf-esp-linha').hidden = !$('#cf-fl').checked;
    $('#cf-desc').value = '';
    $('#cf-fotos').value = ''; $('#cf-fotos-novas').textContent = '';
    const dist = campo.pos && p.latitude != null ? mapa.distance(campo.pos, [+p.latitude, +p.longitude]) : null;
    $('#cf-gps').checked = false;
    $('#cf-gps-linha').hidden = !campo.pos;
    $('#cf-gps-txt').textContent = campo.pos ? `Corrigir a coordenada do poço com o GPS (±${fmt(campo.prec, 0)} m${dist != null ? `; hoje está a ${fmt(dist, 1)} m de você` : ''})` : '';
    $('#cf-obs-atual').textContent = p.observacao ? 'Observação atual: ' + p.observacao : '';
    $('#cf-erro').hidden = true;
    $('#modal-campo').hidden = false;
    // fotos já tiradas deste poço
    const el = $('#cf-galeria'); el.innerHTML = '<span class="muted">carregando fotos...</span>';
    try {
      const { data } = await sb.storage.from(BUCKET_FOTOS).list(`${pastaCampo()}/${p.id}`, { limit: 50, sortBy: { column: 'name', order: 'desc' } });
      const arqs = (data || []).filter(f => /\.jpe?g$/i.test(f.name));
      if (!arqs.length) { el.innerHTML = '<span class="muted">nenhuma foto de campo ainda</span>'; return; }
      const urls = await Promise.all(arqs.slice(0, 12).map(f => sb.storage.from(BUCKET_FOTOS).createSignedUrl(`${pastaCampo()}/${p.id}/${f.name}`, 3600)));
      el.innerHTML = urls.map((u, i) => u.data?.signedUrl ? `<a href="${u.data.signedUrl}" target="_blank" rel="noopener" title="${esc(arqs[i].name)}"><img src="${u.data.signedUrl}" alt="Foto ${esc(arqs[i].name)}"></a>` : '').join('');
    } catch (e) { el.innerHTML = '<span class="muted">não consegui listar as fotos</span>'; }
  }

  async function reduzirFoto(arq, lado = 1600) {
    const url = URL.createObjectURL(arq);
    const img = await new Promise((ok, erro) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => erro(new Error('foto inválida')); i.src = url; });
    const f = Math.min(1, lado / Math.max(img.naturalWidth, img.naturalHeight)), cv = document.createElement('canvas');
    cv.width = Math.round(img.naturalWidth * f); cv.height = Math.round(img.naturalHeight * f);
    cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height);
    URL.revokeObjectURL(url);
    return paraBlob(cv, .82);
  }

  async function salvarFichaCampo() {
    const fc = campo.ficha; if (!fc) return;
    const p = fc.p, btn = $('#cf-salvar'), erro = t => { $('#cf-erro').textContent = t; $('#cf-erro').hidden = false; btn.disabled = false; btn.textContent = 'Salvar'; };
    const achado = $('#cf-achado').value === 'sim', situacao = $('#cf-situacao').value, desc = $('#cf-desc').value.trim();
    const na = WGImport.numero($('#cf-na').value), fl = $('#cf-fl').checked, esp = WGImport.numero($('#cf-esp').value);
    if ($('#cf-na').value.trim() && na === null) return erro('N.A. inválido (use metros, ex.: 3,45).');
    btn.disabled = true; btn.textContent = 'Salvando...';
    try {
      const hoje = new Date(), dataTxt = hoje.toLocaleDateString('pt-BR');
      // 1) poço: situação, observação (com a data) e coordenada do GPS
      const upd = {};
      if (situacao !== (p.situacao || 'ATIVO')) upd.situacao = situacao;
      const notas = [!achado ? 'Não localizado' : '', desc].filter(Boolean).join('. ');
      if (notas) upd.observacao = [`[Campo ${dataTxt}] ${notas}`, p.observacao].filter(Boolean).join(' | ').slice(0, 2000);
      if (achado && $('#cf-gps').checked && campo.pos) { upd.latitude = Math.round(campo.pos[0] * 1e8) / 1e8; upd.longitude = Math.round(campo.pos[1] * 1e8) / 1e8; }
      if (Object.keys(upd).length) {
        const { error } = await sb.from('wg_poco').update(upd).eq('id', p.id);
        if (error) throw error;
        Object.assign(p, upd);
      }
      // 2) N.A. da campanha
      if (st.campanhaId && (na !== null || fl)) {
        const reg = { poco_id: p.id, campanha_id: st.campanhaId, data_medicao: hoje.toISOString().slice(0, 10), nivel_agua: na, fase_livre: fl, espessura_fl: fl ? esp : null };
        const { error } = await sb.from('wg_medicao_campo').upsert(reg, { onConflict: 'poco_id,campanha_id' });
        if (error) throw error;
        const i = st.medicoes.findIndex(m => m.poco_id === p.id && m.campanha_id === st.campanhaId);
        if (i >= 0) Object.assign(st.medicoes[i], reg); else st.medicoes.push(reg);
      }
      // 3) fotos
      const arqs = [...$('#cf-fotos').files];
      const carimbo = hoje.toISOString().replace(/[-:]/g, '').slice(0, 15);
      for (let i = 0; i < arqs.length; i++) {
        btn.textContent = `Enviando foto ${i + 1} de ${arqs.length}...`;
        const blob = await reduzirFoto(arqs[i]);
        const { error } = await sb.storage.from(BUCKET_FOTOS).upload(`${pastaCampo()}/${p.id}/${carimbo}-${i + 1}.jpg`, blob, { upsert: true, contentType: 'image/jpeg' });
        if (error) throw error;
      }
      $('#modal-campo').hidden = true;
      btn.disabled = false; btn.textContent = 'Salvar';
      render();
      if (campo.ativo) { renderListaCampo(); desenharEuCampo(); }
      toast(`${p.codigo} salvo${arqs.length ? ` com ${arqs.length} foto(s)` : ''}.`);
    } catch (e) {
      console.error(e);
      erro('Não consegui salvar: ' + (e.message || e));
    }
  }

  function ligarEventosCampo() {
    $('#btn-campo').addEventListener('click', iniciarCampo);
    $('#campo-sair').addEventListener('click', sairCampo);
    $('#campo-eu').addEventListener('click', () => { campo.seguir = true; if (campo.pos) mapa.setView(campo.pos, Math.max(mapa.getZoom(), 19)); });
    $('#campo-alvo-ficha').addEventListener('click', () => { if (campo.alvo != null) abrirFichaCampo(campo.alvo); });
    $('#campo-alvo-limpar').addEventListener('click', () => { campo.alvo = null; desenharEuCampo(); atualizarAlvoCampo(); renderListaCampo(); });
    $('#campo-lista').addEventListener('click', e => {
      const f = e.target.closest('[data-ficha]'); if (f) { abrirFichaCampo(+f.dataset.ficha || f.dataset.ficha); return; }
      const b = e.target.closest('[data-ir]'); if (b) irParaPoco(+b.dataset.ir || b.dataset.ir);
    });
    $('#campo-recolher').addEventListener('click', () => $('#campo-painel').classList.toggle('recolhido'));
    ['#cf-fechar', '#cf-cancelar'].forEach(id => $(id).addEventListener('click', () => { $('#modal-campo').hidden = true; }));
    $('#cf-salvar').addEventListener('click', salvarFichaCampo);
    $('#cf-fl').addEventListener('change', e => { $('#cf-esp-linha').hidden = !e.target.checked; });
    $('#cf-achado').addEventListener('change', e => { if (e.target.value === 'nao') $('#cf-gps').checked = false; $('#cf-gps-linha').hidden = e.target.value === 'nao' || !campo.pos; });
    $('#cf-fotos').addEventListener('change', e => { const n = e.target.files.length; $('#cf-fotos-novas').textContent = n ? `${n} foto(s) para enviar` : ''; });
    $('#btn-orto').addEventListener('click', () => $('#arquivo-orto').click());
    $('#arquivo-orto').addEventListener('change', e => { const f = [...(e.target.files || [])]; e.target.value = ''; if (f.length) inserirOrtofoto(f); });
    // ficha também pelo popup do poço
    document.addEventListener('click', e => { const b = e.target.closest('[data-ficha-campo]'); if (b) { e.preventDefault(); mapa.closePopup(); abrirFichaCampo(+b.dataset.fichaCampo || b.dataset.fichaCampo); } });
  }

  // ------------------------------------------------------------- prancha A3 em PDF (mapa ou seção + legenda + carimbo com logo)
  const JSPDF_URL = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
  const chavePrancha = () => 'webgeo.prancha.' + (st.projetoId || '-');
  const chaveLogo = () => 'webgeo.logo.' + (st.org?.id || '-');
  const lerLS = (k, padrao) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : padrao; } catch (e) { return padrao; } };
  const gravarLS = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } };

  function carregarJsPdf() {
    if (window.jspdf?.jsPDF) return Promise.resolve(window.jspdf.jsPDF);
    return new Promise((ok, erro) => {
      const s = document.createElement('script');
      s.src = JSPDF_URL; s.onload = () => window.jspdf?.jsPDF ? ok(window.jspdf.jsPDF) : erro(new Error('jsPDF não carregou'));
      s.onerror = () => erro(new Error('não consegui baixar o gerador de PDF (jsPDF). Confira a internet.'));
      document.head.appendChild(s);
    });
  }

  const chaveRotulos = () => 'webgeo.rotulos.' + (st.projetoId || '-');
  /** Valor para o rótulo, sem a unidade (a unidade vai nas notas): "450", "< 1", "0,52". */
  const valorRotulo = r => r.menor_que_lq ? (r.lq != null ? `< ${fmt(r.lq, 4)}` : '< LQ') : fmt(r.valor, 4);

  /** Modo "ajustar rótulos": nome e valor de cada poço viram etiquetas arrastáveis no mapa; a posição vale para a prancha. */
  function ajustarRotulosPrancha(daPrancha = true) {
    const vindoDaPrancha = daPrancha === true;
    $('#modal-prancha').hidden = true;
    sairRotulosPrancha(false);
    const { pocos, porPoco } = selecao();
    const pos = lerLS(chaveRotulos(), {});
    const camada = L.layerGroup().addTo(mapa);
    st.rotAjuste = { camada, pos, mudou: false, daPrancha: vindoDaPrancha };
    atualizarRotulos();
    const etiqueta = (p, tipo, txt, ll0, classe) => {
      const m = L.marker(ll0, { draggable: true, keyboard: false, zIndexOffset: tipo === 'nome' ? 3000 : 2900,
        icon: L.divIcon({ className: 'rot-ajuste-icone', html: `<span class="rot-ajuste ${classe}">${esc(txt)}</span>`, iconSize: null }) }).addTo(camada);
      const linha = L.polyline([[+p.latitude, +p.longitude], ll0], { color: '#FFD60A', weight: 1, interactive: false }).addTo(camada);
      m.on('drag', () => linha.setLatLngs([[+p.latitude, +p.longitude], m.getLatLng()]));
      m.on('dragend', () => { const q = m.getLatLng(); (pos[p.id] = pos[p.id] || {})[tipo] = [q.lat, q.lng]; st.rotAjuste.mudou = true; });
    };
    pocos.filter(p => p.latitude != null && p.longitude != null).forEach(p => {
      const c = mapa.latLngToContainerPoint([+p.latitude, +p.longitude]), sp = pos[p.id] || {};
      const pad = (dx, dy) => mapa.containerPointToLatLng([c.x + dx, c.y + dy]);
      etiqueta(p, 'nome', p.codigo, sp.nome ? L.latLng(sp.nome) : pad(9, 0), 'rot-nome');
      const r = porPoco.get(p.id);
      if (r) etiqueta(p, 'valor', valorRotulo(r), sp.valor ? L.latLng(sp.valor) : pad(0, 13), 'rot-valor' + (status(r) === 'acima' ? ' rot-acima' : ''));
      const h = st.potCalc?.sup && st.potCalc.porPoco.get(p.id);
      if (h != null && h !== false) etiqueta(p, 'carga', Number(h).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }), sp.carga ? L.latLng(sp.carga) : pad(0, -15), 'rot-valor rot-carga');
    });
    mapa.getContainer().classList.add('mapa-rotulos');
    $('#rot-barra').hidden = false;
    $('#rot-barra').scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }

  function sairRotulosPrancha(salvar, voltarPrancha = false) {
    const a = st.rotAjuste; if (!a) return;
    if (salvar) gravarLS(chaveRotulos(), a.pos);
    mapa.removeLayer(a.camada);
    st.rotAjuste = null;
    mapa.getContainer().classList.remove('mapa-rotulos');
    $('#rot-barra').hidden = true;
    atualizarRotulos();
    if (voltarPrancha) { $('#modal-prancha').hidden = false; $('#pr-status').textContent = salvar ? 'Posição dos rótulos salva. Gere o PDF.' : ''; }
  }

  function rotulosPadraoPrancha() {
    if (!st.rotAjuste) return;
    st.rotAjuste.pos = {}; gravarLS(chaveRotulos(), {});
    const dp = st.rotAjuste.daPrancha; sairRotulosPrancha(false); ajustarRotulosPrancha(dp);
  }

  /** Abre a janela da prancha com os campos do carimbo (lembrados por projeto neste navegador). */
  function abrirPrancha() {
    if (!st.projetoId) { toast('Escolha um projeto primeiro.'); return; }
    const proj = st.projetos.find(p => p.id === st.projetoId), param = st.parametros.find(p => p.id === st.parametroId);
    const camp = st.campanhas.find(c => c.id === st.campanhaId);
    const salvo = lerLS(chavePrancha(), {});
    const hoje = new Date(), mes = `${String(hoje.getMonth() + 1).padStart(2, '0')}/${hoje.getFullYear()}`;
    const v = {
      relatorio: salvo.relatorio ?? 'RELATÓRIO DE INVESTIGAÇÃO AMBIENTAL DETALHADA',
      projeto: salvo.projeto ?? (proj?.nome || ''), data: salvo.data ?? mes,
      endereco: salvo.endereco ?? '', contratante: salvo.contratante ?? '', responsavel: salvo.responsavel ?? '',
      folha: salvo.folha ?? '01/01', empresa: salvo.empresa ?? (st.org?.name || ''),
      enquadrar: salvo.enquadrar ?? 'pocos', escala: salvo.escala ?? '0', 'escala-n': salvo['escala-n'] ?? ''
    };
    Object.entries(v).forEach(([k, val]) => { const el = $('#pr-' + k); if (el) el.value = val; });
    // conteúdo: o mapa ou uma das seções
    const temPot = !!(st.pot.ativa && st.potCalc?.sup);
    const opc = [`<option value="mapa">Mapa${param && st.pluma.ativa ? ` - pluma de ${esc(param.nome)}` : ''}${camp ? ` (${esc(camp.codigo)})` : ''}</option>`]
      .concat(temPot ? [`<option value="pot">Mapa potenciométrico${camp ? ` (${esc(camp.codigo)})` : ''}</option>`] : [])
      .concat((st.secao.lista || []).map((x, i) => `<option value="secao:${i}">Seção geológica ${esc(nomeSecao(x.letra))}</option>`));
    $('#pr-conteudo').innerHTML = opc.join('');
    if (temPot && !st.pluma.ativa) $('#pr-conteudo').value = 'pot';
    $('#pr-titulo').value = tituloPadraoPrancha($('#pr-conteudo').value);
    $('#pr-escala-n').hidden = $('#pr-escala').value !== 'outra';
    $('#pr-opc-mapa').hidden = false;
    mostrarLogoPrancha();
    $('#pr-status').textContent = '';
    $('#pr-gerar').disabled = false;
    $('#modal-prancha').hidden = false;
  }

  function tituloPadraoPrancha(conteudo) {
    const param = st.parametros.find(p => p.id === st.parametroId), camp = st.campanhas.find(c => c.id === st.campanhaId);
    if (conteudo.startsWith('secao:')) { const x = st.secao.lista[+conteudo.split(':')[1]]; return `SEÇÃO GEOLÓGICA ${x ? nomeSecao(x.letra) : ''}`.toUpperCase(); }
    if (conteudo === 'pot') return `MAPA POTENCIOMÉTRICO${camp ? ` - CAMPANHA ${camp.codigo}` : ''}`.toUpperCase();
    return ((param && st.pluma.ativa ? `${st.pluma.calor && !limiteManual() ? 'MAPA DE CALOR' : 'MAPA DA PLUMA'} DE ${param.nome}` : 'MAPA DE LOCALIZAÇÃO DOS POÇOS') + (camp ? ` - CAMPANHA ${camp.codigo}` : '')).toUpperCase();
  }

  function mostrarLogoPrancha() {
    const logo = lerLS(chaveLogo(), null);
    $('#pr-logo-ver').innerHTML = logo ? `<img src="${logo}" alt="Logo da empresa">` : '<span>sem logo</span>';
    $('#pr-logo-tirar').hidden = !logo;
  }

  /** Logo: reduz para no máximo 900 px (PNG, mantém transparência) e guarda neste navegador, para a empresa. */
  async function escolherLogoPrancha(arq) {
    try {
      const url = await new Promise((ok, erro) => { const r = new FileReader(); r.onload = () => ok(r.result); r.onerror = erro; r.readAsDataURL(arq); });
      const img = await new Promise((ok, erro) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => erro(new Error('imagem inválida')); i.src = url; });
      const f = Math.min(1, 900 / Math.max(img.width, img.height)), cv = document.createElement('canvas');
      cv.width = Math.round(img.width * f); cv.height = Math.round(img.height * f);
      cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height);
      if (!gravarLS(chaveLogo(), cv.toDataURL('image/png'))) throw new Error('o navegador não deixou guardar a imagem');
      mostrarLogoPrancha();
    } catch (e) { toast('Não consegui usar essa imagem: ' + (e.message || e)); }
  }

  let previaPrancha = null;   // { url, nome, blob, pag, zoom } do último PDF visualizado
  const PDFJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
  const PDFJS_WORKER = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  function carregarPdfJs() {
    if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
    return new Promise((ok, erro) => {
      const s = document.createElement('script');
      s.src = PDFJS_URL;
      s.onload = () => { if (!window.pdfjsLib) return erro(new Error('pdf.js não carregou')); window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER; ok(window.pdfjsLib); };
      s.onerror = () => erro(new Error('não consegui baixar o visualizador de PDF. Confira a internet.'));
      document.head.appendChild(s);
    });
  }

  /** Desenha a folha num canvas (funciona em qualquer navegador, inclusive no celular). zoom 0 = caber na tela. */
  async function desenharPrevia() {
    const pv = previaPrancha, box = $('#previa-quadro'); if (!pv) return;
    try {
      if (!pv.pag) {
        const lib = await carregarPdfJs();
        const pdf = await lib.getDocument({ data: new Uint8Array(await pv.blob.arrayBuffer()) }).promise;
        pv.pag = await pdf.getPage(1);
      }
      const v1 = pv.pag.getViewport({ scale: 1 });
      const caber = Math.min((box.clientWidth - 32) / v1.width, (box.clientHeight - 32) / v1.height);
      const fator = pv.zoom ? caber * pv.zoom : caber;
      const dpr = Math.min(window.devicePixelRatio || 1, 2.5), vp = pv.pag.getViewport({ scale: fator * dpr });
      const cv = document.createElement('canvas'); cv.width = Math.round(vp.width); cv.height = Math.round(vp.height);
      cv.style.width = Math.round(vp.width / dpr) + 'px'; cv.style.height = Math.round(vp.height / dpr) + 'px';
      if (pv.tarefa) { try { pv.tarefa.cancel(); } catch (e) { /* ok */ } }
      pv.tarefa = pv.pag.render({ canvasContext: cv.getContext('2d'), viewport: vp });
      await pv.tarefa.promise; pv.tarefa = null;
      box.innerHTML = ''; box.appendChild(cv);
    } catch (e) {
      if (e?.name === 'RenderingCancelledException') return;
      console.error(e);
      box.innerHTML = `<p class="previa-carregando">Não consegui mostrar a folha aqui (${esc(e.message || String(e))}). Use "Abrir em nova aba" ou "Baixar PDF".</p>`;
    }
  }
  function zoomPrevia(f) {
    const pv = previaPrancha; if (!pv) return;
    pv.zoom = f === 0 ? 0 : Math.max(.5, Math.min(6, (pv.zoom || 1) * f));
    desenharPrevia();
  }

  async function gerarPrancha(previa = false) {
    const campos = ['relatorio', 'projeto', 'data', 'endereco', 'contratante', 'responsavel', 'folha', 'empresa', 'enquadrar', 'escala', 'escala-n'];
    const v = Object.fromEntries(campos.map(k => [k, $('#pr-' + k).value.trim()]));
    gravarLS(chavePrancha(), v);
    v.titulo = $('#pr-titulo').value.trim();
    const conteudo = $('#pr-conteudo').value;
    const btns = [$('#pr-gerar'), $('#pr-visualizar')], msg = $('#pr-status');
    btns.forEach(b => { b.disabled = true; }); msg.textContent = previa ? 'Montando a visualização...' : 'Preparando o PDF...';
    try {
      const JsPDF = await carregarJsPdf();
      const doc = new JsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3', compress: true });
      const G = { MX: 13, MY: 13, MW: 300, MH: 271, PX: 316, PY: 13, PW: 91, PH: 271 };   // área do desenho e painel da direita
      doc.setLineWidth(.6); doc.setDrawColor(0); doc.rect(10, 10, 400, 277);
      let info;
      if (conteudo.startsWith('secao:')) info = await desenharSecaoPrancha(doc, G, +conteudo.split(':')[1], msg);
      else info = await desenharMapaPrancha(doc, G, msg, { modo: conteudo === 'pot' ? 'pot' : 'mapa', valores: $('#pr-valores').checked, enquadrar: $('#pr-enquadrar').value, escala: $('#pr-escala').value === 'outra' ? WGImport.numero($('#pr-escala-n').value) : +$('#pr-escala').value });
      desenharPainelPrancha(doc, G, v, info);
      const proj = st.projetos.find(p => p.id === st.projetoId);
      const nome = `prancha_A3_${proj?.nome || 'projeto'}_${conteudo.startsWith('secao') ? 'secao' : conteudo === 'pot' ? 'potenciometrico' : 'mapa'}`.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\-]+/g, '_') + '.pdf';
      const avisos = info.avisos?.length ? ` (${info.avisos.join('; ')})` : '';
      if (previa) {
        if (previaPrancha) URL.revokeObjectURL(previaPrancha.url);
        const blob = doc.output('blob');
        previaPrancha = { url: URL.createObjectURL(blob), nome, blob, pag: null, zoom: 0 };
        $('#previa-nome').textContent = nome;
        $('#previa-aba').href = previaPrancha.url;
        $('#previa-aviso').hidden = !avisos; $('#previa-aviso').textContent = avisos ? 'Atenção:' + avisos : '';
        $('#previa-quadro').innerHTML = '<p class="previa-carregando">Desenhando a folha...</p>';
        $('#modal-prancha').hidden = true; $('#modal-previa').hidden = false;
        await desenharPrevia();
        msg.textContent = '';
      } else {
        doc.save(nome);
        msg.textContent = `PDF gerado: ${nome}` + avisos;
      }
    } catch (e) {
      console.error(e);
      msg.textContent = 'Não consegui gerar o PDF: ' + (e.message || e);
    }
    btns.forEach(b => { b.disabled = false; });
  }

  function baixarPrevia() {
    if (!previaPrancha) return;
    const a = document.createElement('a'); a.href = previaPrancha.url; a.download = previaPrancha.nome;
    document.body.appendChild(a); a.click(); a.remove();
    toast('PDF baixado: ' + previaPrancha.nome);
  }

  function fecharPrevia(voltar) {
    $('#modal-previa').hidden = true;
    $('#previa-quadro').innerHTML = '';
    if (voltar) $('#modal-prancha').hidden = false;
  }

  /** Texto de uma linha que pode ter "µ" (as fontes padrão do PDF não têm o µ: ele sai com a fonte Symbol, letra m). */
  function textoMu(doc, t, x, y, tam, estilo = 'normal') {
    const partes = String(t).split('µ');
    let cx = x;
    partes.forEach((pt, k) => {
      doc.setFont('helvetica', estilo); doc.setFontSize(tam);
      if (pt) { doc.text(pt, cx, y); cx += doc.getTextWidth(pt); }
      if (k < partes.length - 1) { doc.setFont('symbol', 'normal'); doc.setFontSize(tam); doc.text('m', cx, y); cx += doc.getTextWidth('m'); }
    });
    doc.setFont('helvetica', estilo);
  }
  const hexRgbP = h => { const m = /^#?([0-9a-f]{6})$/i.exec(h || ''); return m ? [0, 2, 4].map(k => parseInt(m[1].slice(k, k + 2), 16)) : [0, 0, 0]; };
  const fmtEsc = n => Math.round(n).toLocaleString('pt-BR');

  /** Texto com contorno branco (legível sobre a imagem de satélite). */
  function textoHalo(doc, t, x, y, o = {}) {
    const op = { align: o.align || 'left', angle: o.angle || 0, baseline: o.baseline || 'alphabetic' };
    doc.setFontSize(o.tam || 6); doc.setFont('helvetica', o.negrito ? 'bold' : 'normal');
    if (o.halo !== false) { doc.setLineWidth(o.larg || .55); doc.setLineJoin('round'); doc.setLineCap('round'); doc.setDrawColor(255, 255, 255); doc.text(t, x, y, { ...op, renderingMode: 'stroke' }); doc.setLineJoin('miter'); doc.setLineCap('butt'); }
    doc.setTextColor(...(o.cor || [23, 38, 33])); doc.text(t, x, y, { ...op, renderingMode: 'fill' });
  }

  function carregarImagem(url) {
    return new Promise(ok => {
      const i = new Image(); i.crossOrigin = 'anonymous';
      const t = setTimeout(() => ok(null), 15000);
      i.onload = () => { clearTimeout(t); ok(i); }; i.onerror = () => { clearTimeout(t); ok(null); };
      i.src = url;
    });
  }

  /** Mapa: a vista atual, ajustada ao retângulo da prancha. Satélite + pluma em imagem; planta, malha, limite e poços em vetor. */
  /** Área que a prancha precisa mostrar: a vista do mapa, os poços, a pluma ou a planta. */
  function limitesPrancha(modo) {
    const pts = [];
    const { pocos } = selecao();
    const pocosLL = pocos.filter(p => p.latitude != null && p.longitude != null).map(p => [+p.latitude, +p.longitude]);
    if (modo === 'pluma') {
      const calc = st.calc, sel = plumaDaCampanha(st.campanhaId);
      if (st.pluma.ativa && calc && sel && sel.v && calc.limiar != null) WGPluma.isolinha(calc.g, sel.v, calc.limiar).forEach(l => pts.push(...l));
      if (!pts.length) pts.push(...pocosLL);
    } else if (modo === 'planta') {
      const b = st.planta && cfgPlanta.mostrar() ? limitesDaPlanta() : null;
      if (b) pts.push([b.getSouth(), b.getWest()], [b.getNorth(), b.getEast()]); else pts.push(...pocosLL);
    } else if (modo === 'pocos') pts.push(...pocosLL);
    if (modo === 'vista' || !pts.length) { const b = mapa.getBounds(); return { b, vista: true }; }
    return { b: L.latLngBounds(pts), vista: false };
  }

  /** Escala "redonda" logo acima de um valor (1:250, 500, 750, 1.000, 1.250, 1.500, 2.000, 2.500, 3.000, 4.000, 5.000...). */
  function escalaRedonda(n) {
    const base = [1, 1.25, 1.5, 2, 2.5, 3, 4, 5, 7.5, 10];
    const p = Math.pow(10, Math.floor(Math.log10(Math.max(n, 1))));
    return base.map(k => k * p).find(v => v >= n * 0.999) || 10 * p;
  }

  async function desenharMapaPrancha(doc, G, msg, opc = {}) {
    const avisos = [];
    // centro e escala: pela área escolhida (com folga de 10%) ou pela escala digitada
    const { b, vista } = limitesPrancha(opc.enquadrar || 'pocos');
    const centro = b.getCenter();
    const larguraM = mapa.distance([centro.lat, b.getWest()], [centro.lat, b.getEast()]);
    const alturaM = mapa.distance([b.getSouth(), centro.lng], [b.getNorth(), centro.lng]);
    const folga = vista ? 1 : 1.2;
    const precisa = Math.max(larguraM * folga / G.MW, alturaM * folga / G.MH, 0.05) * 1000;   // 1:n que faz caber
    const escalaN = opc.escala > 0 ? opc.escala : (vista ? precisa : escalaRedonda(precisa));
    const mPorMmAlvo = escalaN / 1000;
    // zoom da imagem: uns 7 pixels por mm (~180 dpi), sem passar de 4.000 px de largura
    const mpp = z => 156543.03392 * Math.cos(centro.lat * Math.PI / 180) / Math.pow(2, z);
    let zR = 19;
    for (let z = 1; z <= 19; z++) { if (mPorMmAlvo / mpp(z) >= 7) { zR = z; break; } }
    while (zR > 1 && G.MW * mPorMmAlvo / mpp(zR) > 4000) zR--;
    const s = mpp(zR) / mPorMmAlvo;                                              // mm por pixel
    const w = G.MW / s, h = G.MH / s, c = mapa.project(centro, zR);
    const ox = c.x - w / 2, oy = c.y - h / 2;
    const P = (lat, lon) => { const p = mapa.project([lat, lon], zR); return [G.MX + (p.x - ox) * s, G.MY + (p.y - oy) * s]; };
    const ll = (x, y) => mapa.unproject([ox + x, oy + y], zR);
    const nw = ll(0, 0), se = ll(w, h);

    // 1) imagem: satélite e pluma num canvas
    const cv = document.createElement('canvas'); cv.width = Math.round(w); cv.height = Math.round(h);
    const ctx = cv.getContext('2d'); ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, cv.width, cv.height);
    let temImagem = false;
    if (prefsMapa.satelite) {
      msg.textContent = 'Baixando a imagem de satélite...';
      const t0 = Math.floor(ox / 256), t1 = Math.floor((ox + w) / 256), u0 = Math.floor(oy / 256), u1 = Math.floor((oy + h) / 256);
      const jobs = [];
      for (let tx = t0; tx <= t1; tx++) for (let ty = u0; ty <= u1; ty++) {
        const url = `https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/${zR}/${ty}/${tx}`;
        jobs.push(carregarImagem(url).then(img => { if (img) ctx.drawImage(img, tx * 256 - ox, ty * 256 - oy, 256, 256); return !!img; }));
      }
      const r = await Promise.all(jobs);
      if (r.some(Boolean)) temImagem = true; else avisos.push('sem imagem de satélite (não baixou)');
    }
    const calc = st.calc, sel = plumaDaCampanha(st.campanhaId), param = st.parametros.find(p => p.id === st.parametroId);
    const modoPot = opc.modo === 'pot';
    const pc = st.potCalc, comPot = !!(st.pot.ativa && pc && pc.sup);
    const comPluma = !modoPot && !!(st.pluma.ativa && calc && sel && sel.v);
    if (modoPot) opc.valores = false;
    if (comPluma) {
      const { g, vmin, vmax } = calc, manPr = limiteManual();
      const calorPr = st.pluma.calor && !manPr, cminPr = Math.max(1e-9, calc.piso || vmin / 100);
      let pc = document.createElement('canvas'); pc.width = g.nx; pc.height = g.ny;
      const pctx = pc.getContext('2d'), img = pctx.createImageData(g.nx, g.ny);
      for (let j = 0; j < g.ny; j++) for (let i = 0; i < g.nx; i++) {
        const val = sel.v[j * g.nx + i];
        if (calorPr) {
          const t = Math.log10(val / cminPr) / (Math.log10(vmax / cminPr) || 1); if (!(t > .02)) continue;
          const [r, gg, b] = WGPluma.corCalor(val, cminPr, vmax), o = ((g.ny - 1 - j) * g.nx + i) * 4;
          img.data[o] = r; img.data[o + 1] = gg; img.data[o + 2] = b; img.data[o + 3] = Math.round(60 + 140 * Math.min(1, t * 1.6));
          continue;
        }
        if (!(val >= vmin) && !manPr) continue;
        const [r, gg, b] = WGPluma.cor(Math.max(val, vmin), vmin, vmax), o = ((g.ny - 1 - j) * g.nx + i) * 4;
        img.data[o] = r; img.data[o + 1] = gg; img.data[o + 2] = b; img.data[o + 3] = 185;
      }
      pctx.putImageData(img, 0, 0);
      if (manPr) pc = mascararPluma(pc, g, 1, polisSuaves(manPr.polis));
      const swl = g.proj.paraLatLon(g.xmin, g.ymin), nel = g.proj.paraLatLon(g.xmin + g.nx * g.cel, g.ymin + g.ny * g.cel);
      const a = mapa.project([nel.lat, swl.lon], zR), b2 = mapa.project([swl.lat, nel.lon], zR);
      ctx.imageSmoothingEnabled = true; ctx.drawImage(pc, a.x - ox, a.y - oy, b2.x - a.x, b2.y - a.y);
      temImagem = true;
    }
    if (temImagem) {
      let dados;
      try { dados = cv.toDataURL('image/jpeg', .9); }
      catch (e) {   // imagem do satélite sem permissão de cópia: refaz só com a pluma
        avisos.push('imagem de satélite bloqueada pelo servidor');
        ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, cv.width, cv.height); dados = null;
      }
      if (dados) doc.addImage(dados, 'JPEG', G.MX, G.MY, G.MW, G.MH, undefined, 'FAST');
    }
    const fundoEscuro = prefsMapa.satelite && temImagem && !avisos.some(a => a.includes('satélite'));

    msg.textContent = 'Desenhando planta, poços e pluma...';
    doc.saveGraphicsState(); doc.rect(G.MX, G.MY, G.MW, G.MH, null); doc.clip(); doc.discardPath();
    // 2) planta
    let temPlanta = false;
    if (st.planta && cfgPlanta.mostrar()) {
      const ocultas = cfgPlanta.ocultas(st.projetoId), modo = cfgPlanta.cor();
      const corFixa = modo === 'cad' ? null : hexRgbP(modo === 'branco' && !fundoEscuro ? '#111111' : CORES_PLANTA[modo]);
      doc.setLineWidth(.15); doc.setLineDashPattern([], 0);
      linhasDaPlanta().camadas.forEach(cm => {
        if (ocultas.has(cm.nome)) return;
        let cor = corFixa || hexRgbP(cm.cor);
        if (!corFixa && !fundoEscuro && cor[0] + cor[1] + cor[2] > 600) cor = [40, 40, 40];   // branco do CAD no papel branco: escurece
        doc.setDrawColor(...cor);
        cm.linhas.forEach(l => {
          if (l.length < 2) return;
          const pts = l.map(([la, lo]) => P(la, lo));
          doc.lines(pts.slice(1).map((p, k) => [p[0] - pts[k][0], p[1] - pts[k][1]]), pts[0][0], pts[0][1], [1, 1], 'S');
        });
        temPlanta = true;
      });
    }
    // 3) malha UTM
    const malha = WGMalha.linhas({ s: se.lat, w: nw.lng, n: nw.lat, e: se.lng }, 6);
    doc.setLineWidth(.12); doc.setLineDashPattern([.8, 1.2], 0); doc.setDrawColor(...(fundoEscuro ? [255, 255, 255] : [120, 120, 120]));
    const traço = pts => { const q = pts.map(([la, lo]) => P(la, lo)); doc.lines(q.slice(1).map((p, k) => [p[0] - q[k][0], p[1] - q[k][1]]), q[0][0], q[0][1], [1, 1], 'S'); return q; };
    const rotE = [], rotN = [];
    malha.verticais.forEach(l => { const q = traço(l.pts); rotE.push({ v: l.v, x: q[0][0] + (q[q.length - 1][0] - q[0][0]) * ((G.MY + G.MH - q[0][1]) / ((q[q.length - 1][1] - q[0][1]) || 1)) }); });
    malha.horizontais.forEach(l => { const q = traço(l.pts); rotN.push({ v: l.v, y: q[0][1] + (q[q.length - 1][1] - q[0][1]) * ((G.MX - q[0][0]) / ((q[q.length - 1][0] - q[0][0]) || 1)) }); });
    doc.setLineDashPattern([], 0);
    // 4) linha do limite
    if (comPluma && calc.limiar != null) {
      linhasDoLimite(calc, sel).forEach(t => {
        if (t.pts.length < 2) return;
        const q = t.pts.map(([la, lo]) => P(la, lo)), d = q.slice(1).map((p, k) => [p[0] - q[k][0], p[1] - q[k][1]]);
        doc.setLineDashPattern([], 0); doc.setLineWidth(1); doc.setDrawColor(255, 255, 255); doc.lines(d, q[0][0], q[0][1], [1, 1], 'S');
        doc.setLineDashPattern(t.definido ? [] : [1.8, 1.4], 0); doc.setLineWidth(.5); doc.setDrawColor(23, 38, 33); doc.lines(d, q[0][0], q[0][1], [1, 1], 'S');
      });
      doc.setLineDashPattern([], 0);
    }
    // 4b) potenciométrico: curvas equipotenciais (com o valor) e setas de fluxo
    let passoPot = null;
    if (comPot) {
      passoPot = pc.passo;
      const casas = pc.passo < 0.1 ? 2 : (pc.passo < 1 ? (Number.isInteger(pc.passo * 10) ? 1 : 2) : 1);
      const rotPot = [];
      pc.niveis.forEach(nv => {
        const cs = WGPot.curva(pc.sup, nv); if (!cs.length) return;
        cs.forEach(cv => {
          if (cv.length < 2) return;
          const q = cv.map(([la, lo]) => P(la, lo)), d = q.slice(1).map((p, k) => [p[0] - q[k][0], p[1] - q[k][1]]);
          doc.setLineWidth(.75); doc.setDrawColor(255, 255, 255); doc.lines(d, q[0][0], q[0][1], [1, 1], 'S');
          doc.setLineWidth(.32); doc.setDrawColor(21, 101, 168); doc.lines(d, q[0][0], q[0][1], [1, 1], 'S');
        });
        const maior = cs.reduce((a, b) => b.length > a.length ? b : a);
        rotPot.push({ p: P(...maior[Math.floor(maior.length / 2)]), t: Number(nv).toLocaleString('pt-BR', { minimumFractionDigits: casas, maximumFractionDigits: casas }) });
      });
      if (st.pot.setas && pc.fluxo.setas.length) pc.fluxo.setas.forEach(sf => {
        // seta com tamanho fixo no papel (5 mm), centrada no ponto da malha
        const a0 = P(...sf.de), b0 = P(...sf.ate), mx = (a0[0] + b0[0]) / 2, my = (a0[1] + b0[1]) / 2;
        const dl = Math.hypot(b0[0] - a0[0], b0[1] - a0[1]) || 1, ux = (b0[0] - a0[0]) / dl, uy = (b0[1] - a0[1]) / dl, Lm = 2.5, Am = 1.1;
        const a = [mx - ux * Lm, my - uy * Lm], b = [mx + ux * Lm, my + uy * Lm];
        const rot = ang => [b[0] + Am * (ux * Math.cos(ang) - uy * Math.sin(ang)), b[1] + Am * (ux * Math.sin(ang) + uy * Math.cos(ang))];
        const w1 = rot(Math.PI - .5), w2 = rot(Math.PI + .5);
        [[1, [255, 255, 255], .7], [0, [11, 53, 88], .3]].forEach(([, cor, lw]) => {
          doc.setLineWidth(lw); doc.setDrawColor(...cor); doc.setLineCap('round'); doc.setLineJoin('round');
          doc.line(a[0], a[1], b[0], b[1]); doc.line(w1[0], w1[1], b[0], b[1]); doc.line(w2[0], w2[1], b[0], b[1]);
        });
        doc.setLineCap('butt'); doc.setLineJoin('miter');
      });
      rotPot.forEach(r => {
        doc.setFont('helvetica', 'bold'); doc.setFontSize(5.2); const tw = doc.getTextWidth(r.t);
        doc.setFillColor(255, 255, 255); doc.setDrawColor(21, 101, 168); doc.setLineWidth(.15);
        doc.rect(r.p[0] - tw / 2 - .6, r.p[1] - 1.3, tw + 1.2, 2.3, 'FD');
        doc.setTextColor(21, 101, 168); doc.text(r.t, r.p[0], r.p[1] + .55, { align: 'center' });
      });
    }
    // 4b) desenhos e medidas feitos no mapa (Desenhar e medir)
    if (prefsMapa.desenhos !== false) des.lista.forEach(d => {
      const q = d.pts.map(([la, lo]) => P(la, lo)), fech = d.fechado && q.length > 2;
      const seg = q.slice(1).map((p, i) => [p[0] - q[i][0], p[1] - q[i][1]]);
      doc.setLineCap('round'); doc.setLineJoin('round');
      doc.setLineDashPattern(d.tipo === 'medida' ? [1.6, 1] : [], 0);
      doc.setLineWidth(Math.max(.18, d.esp * .16)); doc.setDrawColor(...hexRgbP(d.cor));
      doc.lines(seg, q[0][0], q[0][1], [1, 1], 'S', fech);
      doc.setLineDashPattern([], 0);
      if (d.tipo !== 'medida') return;
      const m = medirPts(d.pts, fech), n = q.length;
      q.forEach(([x, y]) => { doc.setFillColor(...hexRgbP(d.cor)); doc.setDrawColor(255, 255, 255); doc.setLineWidth(.15); doc.circle(x, y, .55, 'FD'); });
      if (m.trechos.length > 1) m.trechos.forEach((t, i) => { const a = q[i], b = q[(i + 1) % n]; textoHalo(doc, fmtDist(t), (a[0] + b[0]) / 2, (a[1] + b[1]) / 2 + .6, { tam: 5.2, align: 'center' }); });
      const u = fech ? q[0] : q[n - 1];
      textoHalo(doc, (fech ? 'Perímetro ' : m.trechos.length > 1 ? 'Total ' : '') + fmtDist(m.total), u[0] + 1.2, u[1] - 1.2, { tam: 5.6, negrito: true });
      if (fech) { const c = q.reduce((a, p) => [a[0] + p[0] / n, a[1] + p[1] / n], [0, 0]); textoHalo(doc, 'Área ' + fmtArea(m.area), c[0], c[1] + .7, { tam: 5.8, negrito: true, align: 'center' }); }
    });
    // 5) poços
    const { pocos, porPoco } = selecao();
    const usados = new Set(), rotulos = []; let inativos = false;
    const ordem = { sem: 0, lq: 1, abaixo: 2, acima: 3 };
    pocos.filter(p => p.latitude != null && p.longitude != null)
      .map(p => ({ p, r: porPoco.get(p.id) })).sort((a, b) => ordem[status(a.r)] - ordem[status(b.r)])
      .forEach(({ p, r }) => {
        const [x, y] = P(+p.latitude, +p.longitude);
        if (x < G.MX || x > G.MX + G.MW || y < G.MY || y > G.MY + G.MH) return;
        const stt = status(r), inat = p.situacao && p.situacao !== 'ATIVO'; if (inat) inativos = true;
        const hPoco = comPot ? pc.porPoco.get(p.id) : undefined;
        const corPoco = modoPot ? (hPoco != null ? '#1565A8' : '#9AA5A8') : COR_STATUS[stt];
        if (modoPot) usados.add(hPoco != null ? 'pot' : 'potsem'); else usados.add(stt);
        doc.setLineWidth(.35); doc.setLineDashPattern(inat ? [.4, .4] : [], 0);
        doc.setDrawColor(...(inat ? [23, 38, 33] : [255, 255, 255])); doc.setFillColor(...hexRgbP(corPoco));
        doc.circle(x, y, stt === 'acima' ? 1.35 : 1.1, 'FD'); doc.setLineDashPattern([], 0);
        const sit = inat ? ` (${(ROTULO_SITUACAO[p.situacao] || p.situacao).toLowerCase()})` : '';
        rotulos.push({ p, r, stt, x, y, nome: p.codigo + sit, h: hPoco });
      });
    // rótulos: nome à direita e valor embaixo do poço, ou onde a pessoa arrastou (com linha de chamada se ficou longe)
    const pos = lerLS(chaveRotulos(), {});
    const chamada = (x, y, tx, ty) => { if (Math.hypot(tx - x, ty - y) > 4.5) { doc.setDrawColor(23, 38, 33); doc.setLineWidth(.15); doc.line(x, y, tx, ty); } };
    rotulos.forEach(({ p, r, stt, x, y, nome, h }) => {
      const sp = pos[p.id] || {};
      if (h != null) {      // carga hidráulica (m), em azul, em cima do poço
        const [hx, hy] = sp.carga ? P(sp.carga[0], sp.carga[1]) : [x, y - 2.4];
        if (sp.carga) chamada(x, y, hx, hy);
        textoHalo(doc, Number(h).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }), hx, hy + .7, { tam: 5.6, negrito: true, align: 'center', cor: [21, 101, 168] });
      }
      const [nx, ny] = sp.nome ? P(sp.nome[0], sp.nome[1]) : [x + 1.9, y];
      const temValor = opc.valores !== false && r;
      const [vx, vy] = sp.valor ? P(sp.valor[0], sp.valor[1]) : [x, y + 3.1];
      if (sp.nome) chamada(x, y, nx, ny);
      if (temValor && sp.valor) chamada(x, y, vx, vy);
      textoHalo(doc, nome, nx, ny + .75, { tam: 6, negrito: true });
      if (temValor) textoHalo(doc, valorRotulo(r), vx, vy + .7, { tam: 5.6, negrito: stt === 'acima', align: 'center', cor: stt === 'acima' ? [179, 30, 20] : [23, 38, 33] });
    });
    // sondagens do Perfil sem poço
    let temSond = false;
    if (!st.redesOcultas.has('Sondagens')) st.soltas.forEach(fc => {
      const [x, y] = P(fc.latitude, fc.longitude); if (x < G.MX || x > G.MX + G.MW || y < G.MY || y > G.MY + G.MH) return;
      doc.setFillColor(95, 190, 156); doc.setDrawColor(255, 255, 255); doc.setLineWidth(.3); doc.rect(x - .9, y - .9, 1.8, 1.8, 'FD'); temSond = true;
      textoHalo(doc, String(fc.sondagem_no || 'SP'), x + 1.6, y + .9, { tam: 5.5 });
    });
    doc.restoreGraphicsState();
    // 6) moldura do desenho, valores da malha e escala
    doc.setLineWidth(.4); doc.setDrawColor(0); doc.rect(G.MX, G.MY, G.MW, G.MH);
    const tx = (vv) => WGMalha.texto(vv, malha.passo);
    const fimEscala = G.MX + 6 + 50 + 40;   // a caixa da escala ocupa o canto de baixo à esquerda
    rotE.forEach(r => { if (r.x > fimEscala && r.x < G.MX + G.MW - 8) textoHalo(doc, `E ${tx(r.v)}`, r.x - 1, G.MY + G.MH - 2, { tam: 5.5, angle: 90 }); });
    rotN.forEach(r => { if (r.y > G.MY + 8 && r.y < G.MY + G.MH - 14) textoHalo(doc, `N ${tx(r.v)}`, G.MX + 1.5, r.y - 1, { tam: 5.5 }); });
    const mPorMm = mPorMmAlvo, passos = [1, 2, 2.5, 5], alvo = mPorMm * 50;          // barra com ~50 mm
    const p10 = Math.pow(10, Math.floor(Math.log10(alvo))), comp = passos.map(k => k * p10).concat([10 * p10]).reduce((a, b) => Math.abs(b - alvo) < Math.abs(a - alvo) ? b : a);
    const bl = comp / mPorMm, bx = G.MX + 6, by = G.MY + G.MH - 8;
    doc.setFillColor(255, 255, 255); doc.setDrawColor(0); doc.setLineWidth(.2); doc.rect(bx - 2, by - 6, bl + 26, 10.5, 'FD');
    for (let k = 0; k < 4; k++) { doc.setFillColor(...(k % 2 ? [255, 255, 255] : [0, 0, 0])); doc.rect(bx + k * bl / 4, by, bl / 4, 1.6, 'FD'); }
    textoHalo(doc, '0', bx, by - 1, { tam: 5.5, halo: false, align: 'center' });
    textoHalo(doc, `${fmt(comp, comp < 10 ? 1 : 0)} m`, bx + bl, by - 1, { tam: 5.5, halo: false, align: 'center' });
    textoHalo(doc, `Escala 1:${fmtEsc(escalaN)} (A3)`, bx + bl + 4, by + 1.6, { tam: 6, halo: false, negrito: true });
    return { tipo: 'mapa', modoPot, comPot, passoPot, avisos, comPluma, calc, param, usados, inativos, temPlanta, temSond, valores: opc.valores !== false, escala: `1:${fmtEsc(escalaN)}`, zona: malha.zona, sul: malha.sul };
  }

  /** Seção: o desenho da tela (SVG) convertido em imagem, centrado no retângulo da prancha. */
  async function desenharSecaoPrancha(doc, G, i, msg) {
    const ativaAntes = st.secao.ativa;
    if (i !== st.secao.ativa) abrirSecao(i);
    const svg = document.querySelector('#secao-svg');
    if (!svg) throw new Error('a seção não tem desenho (nenhum poço perto da linha)');
    msg.textContent = 'Convertendo a seção...';
    const c = svg.cloneNode(true); c.setAttribute('style', 'color:#172621');
    const W = +svg.getAttribute('width'), H = +svg.getAttribute('height');
    const txt = new XMLSerializer().serializeToString(c);
    const url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(txt);
    const img = await new Promise((ok, erro) => { const im = new Image(); im.onload = () => ok(im); im.onerror = () => erro(new Error('não consegui converter a seção')); im.src = url; });
    const k = 3, cv = document.createElement('canvas'); cv.width = W * k; cv.height = H * k;
    const ctx = cv.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cv.width, cv.height); ctx.drawImage(img, 0, 0, cv.width, cv.height);
    const esc_ = Math.min((G.MW - 8) / W, (G.MH - 8) / H), dw = W * esc_, dh = H * esc_;
    doc.addImage(cv.toDataURL('image/png'), 'PNG', G.MX + (G.MW - dw) / 2, G.MY + (G.MH - dh) / 2, dw, dh, undefined, 'FAST');
    doc.setLineWidth(.4); doc.setDrawColor(0); doc.rect(G.MX, G.MY, G.MW, G.MH);
    // legenda da seção: lê a legenda da tela (cor + texto)
    const itens = [...document.querySelectorAll('#secao-legenda > span')].map(sp => {
      const i_ = sp.querySelector('i'), cs = i_ ? getComputedStyle(i_) : null;
      const rgb = t => { const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?/.exec(t || ''); return m && (m[4] == null || +m[4] > .05) ? [+m[1], +m[2], +m[3]] : null; };
      const fundo = cs ? rgb(cs.backgroundColor) : null, borda = cs ? (rgb(cs.borderBottomColor) || rgb(cs.borderTopColor)) : null;
      // sem fundo (terreno, nível d'água...): vira uma linha na cor da borda
      return { txt: sp.textContent.trim(), cor: fundo, corLinha: fundo ? null : borda };
    }).filter(x => x.txt);
    const nota = $('#secao-nota').textContent;
    if (ativaAntes !== st.secao.ativa) abrirSecao(ativaAntes);
    return { tipo: 'secao', avisos: [], itens, nota };
  }

  /** Painel da direita: rosa dos ventos, legenda, notas e carimbo com logo. */
  function desenharPainelPrancha(doc, G, v, info) {
    const { PX, PY, PW, PH } = G, X2 = PX + PW;
    doc.setLineWidth(.4); doc.setDrawColor(0); doc.rect(PX, PY, PW, PH);
    doc.setTextColor(0);
    let y = PY;
    // rosa dos ventos (só no mapa)
    if (info.tipo === 'mapa') {
      const cx = PX + PW / 2, cy = PY + 24, R = 13;
      const ponta = (ang, comp, larg, c1, c2) => {
        const ux = Math.sin(ang), uy = -Math.cos(ang), vx = Math.sin(ang + Math.PI / 2), vy = -Math.cos(ang + Math.PI / 2);
        const pt = (t, sx) => [cx + ux * t + vx * sx, cy + uy * t + vy * sx];
        const a = pt(comp, 0), b = pt(0, -larg), c = pt(0, larg);
        doc.setLineWidth(.15); doc.setDrawColor(23, 38, 33);
        doc.setFillColor(...c1); doc.triangle(a[0], a[1], b[0], b[1], cx, cy, 'FD');
        doc.setFillColor(...c2); doc.triangle(a[0], a[1], cx, cy, c[0], c[1], 'FD');
      };
      doc.setLineWidth(.25); doc.setDrawColor(23, 38, 33); doc.circle(cx, cy, R * .82, 'S'); doc.circle(cx, cy, R * .62, 'S');
      [45, 135, 225, 315].forEach(a => ponta(a * Math.PI / 180, R * .62, R * .13, [86, 101, 93], [201, 211, 206]));
      [90, 180, 270].forEach(a => ponta(a * Math.PI / 180, R * .95, R * .2, [23, 38, 33], [255, 255, 255]));
      ponta(0, R * .95, R * .2, [179, 64, 42], [226, 105, 79]);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(179, 64, 42); doc.text('N', cx, cy - R - 1.2, { align: 'center' });
      doc.setTextColor(23, 38, 33); doc.text('S', cx, cy + R + 4.2, { align: 'center' });
      doc.text('L', cx + R + 2.2, cy + 1.2, { align: 'center' }); doc.text('O', cx - R - 2.2, cy + 1.2, { align: 'center' });
      y = PY + 44;
    } else y = PY + 6;
    // legenda
    doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(0); doc.text('LEGENDA', PX + 4, y); y += 2.5;
    const linhaLeg = (desenho, texto) => {
      const linhas = doc.setFont('helvetica', 'normal').setFontSize(6.5).splitTextToSize(texto, PW - 17);
      const hh = Math.max(6.2, linhas.length * 2.7 + 2.6);
      if (y + hh > PY + PH - 92) return false;
      doc.setFillColor(243, 245, 244); doc.setDrawColor(200, 205, 202); doc.setLineWidth(.15); doc.rect(PX + 3, y, PW - 6, hh, 'FD');
      doc.line(PX + 12, y, PX + 12, y + hh);
      desenho(PX + 7.5, y + hh / 2);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(6.5); doc.setTextColor(23, 38, 33);
      if (linhas.length === 1 && texto.includes('µ')) textoMu(doc, texto, PX + 14, y + hh / 2 + .9, 6.5);
      else doc.text(linhas, PX + 14, y + hh / 2 - (linhas.length - 1) * 1.35 + .9);
      y += hh; return true;
    };
    const bola = cor => (cx, cy) => { doc.setFillColor(...hexRgbP(cor)); doc.setDrawColor(23, 38, 33); doc.setLineWidth(.25); doc.circle(cx, cy, 1.5, 'FD'); };
    if (info.tipo === 'mapa') {
      if (info.temPlanta) linhaLeg((cx, cy) => { doc.setDrawColor(40, 40, 40); doc.setLineWidth(.3); doc.line(cx - 3, cy + 1.2, cx - 1, cy - 1.2); doc.line(cx - 1, cy - 1.2, cx + 1.5, cy - 1.2); doc.line(cx + 1.5, cy - 1.2, cx + 3, cy + 1.2); }, 'Planta / área de estudo');
      if (info.usados.has('pot')) linhaLeg(bola('#1565A8'), 'Poço de monitoramento');
      if (info.usados.has('potsem')) linhaLeg(bola('#9AA5A8'), 'Poço sem medição de nível na campanha');
      ['acima', 'abaixo', 'lq', 'sem'].filter(sx => info.usados.has(sx)).forEach(sx => {
        const txt = { acima: `Poço de monitoramento - acima do VI${info.param ? ` (${info.param.nome})` : ''}`, abaixo: 'Poço de monitoramento - quantificado abaixo do VI', lq: 'Poço de monitoramento - abaixo do limite de quantificação (LQ)', sem: 'Poço de monitoramento - sem resultado na campanha' }[sx];
        linhaLeg(bola(COR_STATUS[sx]), txt);
      });
      if (info.inativos) linhaLeg((cx, cy) => { doc.setFillColor(220, 220, 220); doc.setDrawColor(23, 38, 33); doc.setLineDashPattern([.4, .4], 0); doc.circle(cx, cy, 1.5, 'FD'); doc.setLineDashPattern([], 0); }, 'Poço seco, danificado/destruído ou desativado (borda tracejada)');
      if (info.temSond) linhaLeg((cx, cy) => { doc.setFillColor(95, 190, 156); doc.setDrawColor(23, 38, 33); doc.setLineWidth(.2); doc.rect(cx - 1.3, cy - 1.3, 2.6, 2.6, 'FD'); }, 'Sondagem do Perfil de Sondagem (sem poço)');
      if (info.comPot) {
        linhaLeg((cx, cy) => { doc.setLineWidth(.45); doc.setDrawColor(21, 101, 168); doc.line(cx - 3.5, cy, cx + 3.5, cy); },
          'Linha equipotencial');
        // exemplo do rótulo de carga hidráulica que vai junto de cada poço
        linhaLeg((cx, cy) => { textoHalo(doc, '97,64', cx, cy + .8, { tam: 5.6, negrito: true, align: 'center', cor: [21, 101, 168], halo: false }); },
          'Carga hidráulica (m)');
        if (st.pot.setas) linhaLeg((cx, cy) => { doc.setLineWidth(.35); doc.setDrawColor(11, 53, 88); doc.line(cx - 3, cy, cx + 3, cy); doc.line(cx + 1.6, cy - 1, cx + 3, cy); doc.line(cx + 1.6, cy + 1, cx + 3, cy); },
          'Sentido do fluxo da água subterrânea');
      }
      if (info.comPluma) {
        const { vmin, vmax, limiar } = info.calc, un = info.param?.unidade || 'µg/L';
        linhaLeg((cx, cy) => { doc.setLineDashPattern([], 0); doc.setLineWidth(.5); doc.setDrawColor(23, 38, 33); doc.line(cx - 3.5, cy, cx + 3.5, cy); },
          limiar != null ? `Limite da pluma delimitada (${fmt(limiar, 4)} ${un})` : 'Limite da pluma delimitada');
        linhaLeg((cx, cy) => { doc.setLineDashPattern([1.2, 1], 0); doc.setLineWidth(.5); doc.setDrawColor(23, 38, 33); doc.line(cx - 3.5, cy, cx + 3.5, cy); doc.setLineDashPattern([], 0); },
          'Limite de pluma inferida');
        // barra de cores
        if (y + 14 < PY + PH - 92) {
          doc.setFont('helvetica', 'normal'); doc.setFontSize(6.5); doc.setTextColor(23, 38, 33);
          const calorL = st.pluma.calor && !limiteManual(), v0 = calorL ? Math.max(1e-9, info.calc.piso || vmin / 100) : vmin;
          textoMu(doc, `${calorL ? 'Mapa de calor' : 'Concentração interpolada'} de ${info.param?.nome || ''} (${un})`, PX + 4, y + 4, 6.5);
          const bx = PX + 6, bw = PW - 12, by = y + 6.5, n = 40;
          for (let k = 0; k < n; k++) { const vv = v0 * Math.pow(vmax / v0, (k + .5) / n); doc.setFillColor(...(calorL ? WGPluma.corCalor(vv, v0, vmax) : WGPluma.cor(vv, vmin, vmax))); doc.rect(bx + k * bw / n, by, bw / n + .05, 3, 'F'); }
          doc.setDrawColor(23, 38, 33); doc.setLineWidth(.15); doc.rect(bx, by, bw, 3);
          const ticks = [v0]; for (let e = Math.ceil(Math.log10(v0 * 2)); Math.pow(10, e) <= vmax * 1.0001; e++) ticks.push(Math.pow(10, e));
          doc.setFontSize(5.5);
          ticks.forEach(t => { const tx = bx + bw * Math.log10(t / v0) / Math.log10(vmax / v0); doc.line(tx, by + 3, tx, by + 4);
            doc.text(fmt(t, 3), tx, by + 6.3, { align: tx > bx + bw - 4 ? 'right' : tx < bx + 4 ? 'left' : 'center' }); });
          y = by + 9;
        }
      }
    } else {
      info.itens.forEach(it => linhaLeg((cx, cy) => {
        if (it.cor) { doc.setFillColor(...it.cor); doc.setDrawColor(23, 38, 33); doc.setLineWidth(.15); doc.rect(cx - 2.5, cy - 1.6, 5, 3.2, 'FD'); }
        else { doc.setDrawColor(...(it.corLinha || [21, 101, 168])); doc.setLineWidth(.5); doc.line(cx - 3, cy, cx + 3, cy); }
      }, it.txt));
    }
    // (sem bloco de notas: a legenda e o carimbo já dizem o necessário)
    // carimbo (de baixo para cima)
    const CB = PY + PH, CT = CB - 80;
    doc.setDrawColor(0); doc.setLineWidth(.3); doc.line(PX, CT, X2, CT);
    const rotulo = (t, x, yy) => { doc.setFont('helvetica', 'bold'); doc.setFontSize(4.6); doc.setTextColor(60, 60, 60); doc.text(t, x + 1, yy + 2.2); };
    const valor = (t, x, yy, ww, hh, tamF = 7, negr = false) => {
      doc.setFont('helvetica', negr ? 'bold' : 'normal'); doc.setFontSize(tamF); doc.setTextColor(0);
      const ln = doc.splitTextToSize(t || '-', ww - 3).slice(0, 2);
      doc.text(ln, x + ww / 2, yy + hh / 2 + 1.6 - (ln.length - 1) * tamF * .18, { align: 'center' });
    };
    // logo + empresa
    const logo = lerLS(chaveLogo(), null);
    let yy = CT;
    if (logo) {
      try {
        const p = doc.getImageProperties(logo), mw = PW - 10, mh = 20, k = Math.min(mw / p.width, mh / p.height);
        doc.addImage(logo, 'PNG', PX + (PW - p.width * k) / 2, yy + 2 + (mh - p.height * k) / 2, p.width * k, p.height * k);
      } catch (e) { console.warn('logo', e); }
    } else valor(v.empresa || st.org?.name || '', PX, yy, PW, 24, 10, true);
    yy += 24; doc.line(PX, yy, X2, yy);
    valor(v.relatorio, PX, yy, PW, 8, 6.5, true); yy += 8; doc.line(PX, yy, X2, yy);
    valor(v.titulo, PX, yy, PW, 10, 7.5, true); yy += 10; doc.line(PX, yy, X2, yy);
    const meio = PX + PW * .68;
    rotulo('PROJETO:', PX, yy); valor(v.projeto, PX, yy, meio - PX, 9); rotulo('DATA:', meio, yy); valor(v.data, meio, yy, X2 - meio, 9);
    doc.line(meio, yy, meio, yy + 9); yy += 9; doc.line(PX, yy, X2, yy);
    rotulo('ENDEREÇO:', PX, yy); valor(v.endereco, PX, yy, PW, 9, 6.5); yy += 9; doc.line(PX, yy, X2, yy);
    rotulo('CONTRATANTE:', PX, yy); valor(v.contratante, PX, yy, PW, 9, 6.5); yy += 9; doc.line(PX, yy, X2, yy);
    const c1 = PX + PW * .5, c2 = PX + PW * .78;
    rotulo('RESPONSÁVEL TÉCNICO:', PX, yy); valor(v.responsavel, PX, yy, c1 - PX, CB - yy, 6);
    rotulo('ESCALA:', c1, yy); valor(info.escala || 'indicada', c1, yy, c2 - c1, CB - yy, 6);
    rotulo('FOLHA:', c2, yy); valor(v.folha, c2, yy, X2 - c2, CB - yy, 6.5);
    doc.line(c1, yy, c1, CB); doc.line(c2, yy, c2, CB);
  }

  /**
   * Mapa em DXF (R12), em UTM (metros, SIRGAS 2000): planta (camadas visíveis), poços (círculo + nome + resultado)
   * e a pluma da campanha (faixas preenchidas por classe de concentração, linha do limite e isolinhas de 10x, 100x...).
   */
  function baixarMapaDxf() {
    const param = st.parametros.find(p => p.id === st.parametroId), camp = st.campanhas.find(c => c.id === st.campanhaId);
    const proj = st.projetos.find(p => p.id === st.projetoId);
    const { pocos, porPoco } = selecao();
    const comCoord = pocos.filter(p => p.latitude !== null && p.longitude !== null);
    const temPlanta = !!(st.planta && cfgPlanta.mostrar());
    if (!comCoord.length && !temPlanta) { toast('Nada para exportar: nenhum poço com coordenada e nenhuma planta.'); return; }
    // fuso: o da planta; sem planta, o dos poços
    let zona, sul;
    if (temPlanta) { zona = st.planta.dados.zona; sul = st.planta.dados.sul; }
    else { const lo = comCoord.reduce((a, p) => a + +p.longitude, 0) / comCoord.length, la = comCoord.reduce((a, p) => a + +p.latitude, 0) / comCoord.length; zona = WGDxf.zonaDaLongitude(lo); sul = la < 0; }
    const U = (lat, lon) => { const u = WGDxf.latLonParaUtm(lat, lon, zona, sul); return [u.E, u.N]; };
    const ents = [], fundo = [], C = WGDxfSaida.nomeCamada;   // fundo: preenchimento da pluma, desenhado antes (por baixo) da planta
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    const ext = ([x, y]) => { x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y); };

    // 1) planta
    if (temPlanta) {
      const ocultas = cfgPlanta.ocultas(st.projetoId), modo = cfgPlanta.cor();
      // mesma cor escolhida no menu (Cor das linhas); "Cores do CAD" mantém a cor de cada camada
      const corModo = modo === 'cad' ? null : (modo === 'preto' ? 7 : WGDxfSaida.aci(CORES_PLANTA[modo] || '#ffffff'));
      linhasDaPlanta().camadas.forEach(c => {
        if (ocultas.has(c.nome)) return;
        const cor = corModo ?? WGDxfSaida.aci(c.cor);
        c.linhas.forEach(l => { const pts = l.map(([la, lo]) => U(la, lo)); pts.forEach(ext); ents.push({ t: 'poli', pts, camada: C('PLANTA_' + c.nome), cor }); });
      });
    }
    comCoord.forEach(p => ext(U(+p.latitude, +p.longitude)));
    const tam = Math.max(10, x1 - x0, y1 - y0), hTxt = Math.min(2.5, Math.max(.6, tam / 160)), rPoco = hTxt * .45;

    // 2) pluma: faixas por classe (células da grade emendadas em linha), limite e isolinhas
    const calc = st.calc, sel = plumaDaCampanha(st.campanhaId), un = param?.unidade || 'µg/L';
    const classes = [];
    if (st.pluma.ativa && calc && sel && sel.v) {
      const { g, vmin, vmax } = calc;
      for (let v = vmin; v < vmax * 0.999; v *= 10) classes.push({ de: v, ate: Math.min(v * 10, vmax) });
      if (!classes.length) classes.push({ de: vmin, ate: vmax });
      classes.forEach((k, i) => { k.cor = '#' + WGPluma.cor(Math.sqrt(k.de * k.ate), vmin, vmax).map(x => x.toString(16).padStart(2, '0')).join(''); k.camada = C(`PLUMA_FAIXA_${i + 1}`); });
      const manD = limiteManual(); if (manD) manD.suaves = polisSuaves(manD.polis);
      const classe0 = v => { if (!(v >= vmin)) return -1; for (let i = classes.length - 1; i >= 0; i--) if (v >= classes[i].de) return i; return -1; };
      const cls = new Int16Array(g.nx * g.ny);
      for (let j = 0; j < g.ny; j++) for (let i = 0; i < g.nx; i++) {
        let k = classe0(sel.v[j * g.nx + i]);
        if (manD) { const ll = g.proj.paraLatLon(g.xmin + (i + .5) * g.cel, g.ymin + (j + .5) * g.cel); k = dentroLimiteManual(ll.lat, ll.lon, manD.suaves) ? Math.max(0, k) : -1; }
        cls[j * g.nx + i] = k;
      }
      const classe = (v, q) => cls[q];
      const canto = (i, j) => { const ll = g.proj.paraLatLon(g.xmin + i * g.cel, g.ymin + j * g.cel); return U(ll.lat, ll.lon); };
      for (let j = 0; j < g.ny; j++) {
        let i = 0;
        while (i < g.nx) {
          const k = classe(null, j * g.nx + i);
          if (k < 0) { i++; continue; }
          let f = i + 1; while (f < g.nx && classe(null, j * g.nx + f) === k) f++;
          const q = [canto(i, j), canto(f, j), canto(f, j + 1), canto(i, j + 1)];
          q.forEach(ext);
          fundo.push({ t: 'solido', pts: q, camada: classes[k].camada, cor: WGDxfSaida.aci(classes[k].cor) });
          i = f;
        }
      }
      const linhas = (v, camada, cor, rotulo) => WGPluma.isolinha(g, sel.v, v).forEach(l => {
        const pts = l.map(([la, lo]) => U(la, lo)); if (pts.length < 2) return;
        ents.push({ t: 'poli', pts, camada, cor });
        if (rotulo) { const m = pts[Math.floor(pts.length / 2)]; ents.push({ t: 'texto', x: m[0], y: m[1], alt: hTxt * .8, txt: rotulo, al: 1, camada, cor }); }
      });
      if (calc.limiar != null) {
        let rot = false;
        linhasDoLimite(calc, sel).forEach(t => {
          const pts = t.pts.map(([la, lo]) => U(la, lo)); if (pts.length < 2) return;
          ents.push({ t: 'poli', pts, camada: t.definido ? 'PLUMA_LIMITE' : 'PLUMA_LIMITE_INFERIDO', cor: 1, tipo: t.definido ? undefined : 'TRACO' });
          if (!rot && t.definido && pts.length > 4) { rot = true; const m = pts[Math.floor(pts.length / 2)]; ents.push({ t: 'texto', x: m[0], y: m[1], alt: hTxt * .8, txt: `${fmt(calc.limiar, 4)} ${un}`, al: 1, camada: 'PLUMA_LIMITE', cor: 1 }); }
        });
      }
      for (let v = (calc.limiar ?? vmin) * 10; v < vmax * 0.999; v *= 10) linhas(v, 'PLUMA_ISOLINHAS', 30, `${fmt(v, 3)} ${un}`);
    }

    // 2b) potenciométrico: equipotenciais (com o valor), setas de fluxo e carga hidráulica de cada poço
    const pc = st.potCalc, comPot = !!(st.pot.ativa && pc && pc.sup);
    if (comPot) {
      const casas = pc.passo < 0.1 ? 2 : (pc.passo < 1 ? (Number.isInteger(pc.passo * 10) ? 1 : 2) : 1);
      pc.niveis.forEach(nv => {
        const cs = WGPot.curva(pc.sup, nv); if (!cs.length) return;
        cs.forEach(cv => { const pts = cv.map(([la, lo]) => U(la, lo)); if (pts.length < 2) return; pts.forEach(ext); ents.push({ t: 'poli', pts, camada: 'POT_EQUIPOTENCIAIS', cor: 5 }); });
        const maior = cs.reduce((a, b) => b.length > a.length ? b : a), m = U(...maior[Math.floor(maior.length / 2)]);
        ents.push({ t: 'texto', x: m[0], y: m[1], alt: hTxt * .8, txt: Number(nv).toLocaleString('pt-BR', { minimumFractionDigits: casas, maximumFractionDigits: casas }), al: 1, camada: 'POT_EQUIPOTENCIAIS_VALOR', cor: 5 });
      });
      if (st.pot.setas) pc.fluxo.setas.forEach(sf => {
        const a = U(...sf.de), b = U(...sf.ate), w1 = U(...sf.asa1), w2 = U(...sf.asa2);
        ents.push({ t: 'poli', pts: [a, b], camada: 'POT_FLUXO', cor: 150 });
        ents.push({ t: 'poli', pts: [w1, b, w2], camada: 'POT_FLUXO', cor: 150 });
      });
    }

    // 3) poços: círculo (cor do resultado), nome e valor
    const COR_ACI = { acima: 1, abaixo: 30, lq: 5, sem: 8 };
    comCoord.forEach(p => {
      const [x, y] = U(+p.latitude, +p.longitude), r = porPoco.get(p.id), s = status(r), inativo = p.situacao && p.situacao !== 'ATIVO';
      ents.push({ t: 'circ', x, y, r: rPoco, camada: C('POCOS_' + ROTULO_STATUS[s]), cor: COR_ACI[s] });
      ents.push({ t: 'texto', x: x + rPoco * 1.6, y: y + rPoco * .3, alt: hTxt, txt: p.codigo + (inativo ? ` (${(ROTULO_SITUACAO[p.situacao] || p.situacao).toLowerCase()})` : ''), camada: 'POCOS_NOMES', cor: 7 });
      if (r && !(comPot && !classes.length)) ents.push({ t: 'texto', x: x + rPoco * 1.6, y: y - hTxt * 1.1, alt: hTxt * .75, txt: `${textoValor(r)} ${r.unidade || un}`, camada: 'POCOS_RESULTADOS', cor: COR_ACI[s] });
      const h = comPot ? pc.porPoco.get(p.id) : undefined;
      if (h != null) ents.push({ t: 'texto', x, y: y + rPoco * 1.8, alt: hTxt * .8, al: 1, camada: 'POT_CARGA_HIDRAULICA', cor: 5,
        txt: Number(h).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) });
    });

    // 3b) desenhos e medidas feitos no mapa
    if (prefsMapa.desenhos !== false) des.lista.forEach(d => {
      const pts = d.pts.map(([la, lo]) => U(la, lo)), fech = d.fechado && pts.length > 2, camada = d.tipo === 'medida' ? 'MEDIDAS' : 'DESENHOS';
      const cor = WGDxfSaida.aci(d.cor); pts.forEach(ext);
      ents.push({ t: 'poli', pts: fech ? [...pts, pts[0]] : pts, camada, cor });
      if (d.tipo !== 'medida') return;
      const m = medirPts(d.pts, fech), n = pts.length;
      if (m.trechos.length > 1) m.trechos.forEach((t, i) => { const a = pts[i], b = pts[(i + 1) % n]; ents.push({ t: 'texto', x: (a[0] + b[0]) / 2, y: (a[1] + b[1]) / 2, alt: hTxt * .7, al: 1, txt: fmtDist(t), camada, cor }); });
      const u = fech ? pts[0] : pts[n - 1];
      ents.push({ t: 'texto', x: u[0] + hTxt * .5, y: u[1] + hTxt * .5, alt: hTxt * .8, txt: (fech ? 'Perímetro ' : m.trechos.length > 1 ? 'Total ' : '') + fmtDist(m.total), camada, cor });
      if (fech) { const c = pts.reduce((a, p) => [a[0] + p[0] / n, a[1] + p[1] / n], [0, 0]); ents.push({ t: 'texto', x: c[0], y: c[1], alt: hTxt * .85, al: 1, txt: 'Área ' + fmtArea(m.area), camada, cor }); }
    });

    // 4) legenda e norte, à direita do desenho
    let lx = x1 + tam * .06, ly = y1;
    const linhaLeg = (txt, alt = hTxt, cor = 7) => { ents.push({ t: 'texto', x: lx, y: ly, alt, txt, camada: 'LEGENDA', cor }); ly -= alt * 1.8; };
    linhaLeg(`${proj?.nome || 'Projeto'}${param && !(comPot && !classes.length) ? ' - ' + param.nome : ''}${comPot && !classes.length ? ' - mapa potenciométrico' : ''}${camp ? ' - campanha ' + camp.codigo : ''}`, hTxt * 1.4);
    linhaLeg(`UTM fuso ${zona}${sul ? 'S' : 'N'} - SIRGAS 2000 / WGS 84 - unidades em metros`, hTxt * .8);
    ly -= hTxt;
    if (classes.length) {
      linhaLeg(`Pluma (${un})${calc.limiar != null ? ` - limite ${fmt(calc.limiar, 4)}` : ''}`);
      classes.forEach(k => {
        const q = [[lx, ly - hTxt * .2], [lx + hTxt * 2, ly - hTxt * .2], [lx + hTxt * 2, ly + hTxt], [lx, ly + hTxt]];
        ents.push({ t: 'solido', pts: q, camada: 'LEGENDA', cor: WGDxfSaida.aci(k.cor) });
        ents.push({ t: 'texto', x: lx + hTxt * 2.6, y: ly, alt: hTxt * .85, txt: `${fmt(k.de, 3)} a ${fmt(k.ate, 3)}`, camada: 'LEGENDA', cor: 7 });
        ly -= hTxt * 1.8;
      });
      ly -= hTxt;
    }
    if (comPot) {
      linhaLeg(`Mapa potenciométrico${camp ? ' - campanha ' + camp.codigo : ''}`);
      ents.push({ t: 'poli', pts: [[lx, ly + hTxt * .4], [lx + hTxt * 2, ly + hTxt * .4]], camada: 'LEGENDA', cor: 5 });
      ents.push({ t: 'texto', x: lx + hTxt * 2.6, y: ly, alt: hTxt * .85, txt: `Linha equipotencial (a cada ${fmt(pc.passo, 2)} m)`, camada: 'LEGENDA', cor: 7 }); ly -= hTxt * 1.8;
      ents.push({ t: 'texto', x: lx + hTxt, y: ly, alt: hTxt * .8, al: 1, txt: '97,64', camada: 'LEGENDA', cor: 5 });
      ents.push({ t: 'texto', x: lx + hTxt * 2.6, y: ly, alt: hTxt * .85, txt: 'Carga hidráulica (m)', camada: 'LEGENDA', cor: 7 }); ly -= hTxt * 1.8;
      if (st.pot.setas) {
        const yy = ly + hTxt * .4;
        ents.push({ t: 'poli', pts: [[lx, yy], [lx + hTxt * 2, yy]], camada: 'LEGENDA', cor: 150 });
        ents.push({ t: 'poli', pts: [[lx + hTxt * 1.4, yy + hTxt * .4], [lx + hTxt * 2, yy], [lx + hTxt * 1.4, yy - hTxt * .4]], camada: 'LEGENDA', cor: 150 });
        ents.push({ t: 'texto', x: lx + hTxt * 2.6, y: ly, alt: hTxt * .85, txt: 'Sentido do fluxo da água subterrânea', camada: 'LEGENDA', cor: 7 }); ly -= hTxt * 1.8;
      }
      ly -= hTxt;
    }
    linhaLeg('Poços');
    ['acima', 'abaixo', 'lq', 'sem'].forEach(s => {
      ents.push({ t: 'circ', x: lx + rPoco, y: ly + hTxt * .4, r: rPoco, camada: 'LEGENDA', cor: COR_ACI[s] });
      ents.push({ t: 'texto', x: lx + hTxt * 2.6, y: ly, alt: hTxt * .85, txt: ROTULO_STATUS[s], camada: 'LEGENDA', cor: 7 });
      ly -= hTxt * 1.8;
    });
    const nx = lx + hTxt * 2, ny = ly - hTxt * 5, ha = hTxt * 4;   // seta do norte
    ents.push({ t: 'solido', pts: [[nx, ny + ha], [nx + ha * .3, ny], [nx, ny + ha * .25]], camada: 'LEGENDA', cor: 1 });
    ents.push({ t: 'poli', pts: [[nx, ny + ha], [nx - ha * .3, ny], [nx, ny + ha * .25]], fechado: true, camada: 'LEGENDA', cor: 7 });
    ents.push({ t: 'texto', x: nx, y: ny + ha + hTxt * .5, alt: hTxt * 1.2, txt: 'N', al: 1, camada: 'LEGENDA', cor: 7 });

    const dxf = WGDxfSaida.escrever([...fundo, ...ents], { traco: [hTxt, hTxt * .6] });
    const nome = `${comPot && !classes.length ? 'potenciometrico' : 'mapa'}_${(proj?.nome || 'projeto')}${comPot && !classes.length ? '' : param ? '_' + param.nome : ''}${camp ? '_' + camp.codigo : ''}`.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w\-]+/g, '_') + '.dxf';
    const url = URL.createObjectURL(new Blob([WGDxfSaida.bytes(dxf)], { type: 'application/dxf' }));
    const a = document.createElement('a'); a.href = url; a.download = nome;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    toast(`DXF do mapa baixado (${comCoord.length} poço(s)${temPlanta ? ', planta' : ''}${classes.length ? ', pluma' : ''}${comPot ? ', potenciométrico' : ''}).`);
  }

  function baixarSecaoDxf() {
    const svg = document.querySelector('#secao-svg'), D = st.secaoDes; if (!svg || !D) return;
    // X = distância ao longo da seção (m); Y = cota x exagero vertical (o desenho fica igual ao da tela)
    const X = x => (x - D.mL) / D.pxH, Y = y => D.ev * D.zmax - (y - D.mT) / D.pxH, P = ([x, y]) => [X(x), Y(y)];
    const ents = [], C = WGDxfSaida.nomeCamada;
    const hex = v => { const t = String(v || '').trim(); if (!t || t === 'none') return null; if (t === 'currentColor') return '#172621'; return /^#[0-9a-f]{6}$/i.test(t) ? t : (/^#[0-9a-f]{3}$/i.test(t) ? '#' + [...t.slice(1)].map(c => c + c).join('') : null); };
    const mistura = (a, b, t) => '#' + [1, 3, 5].map(k => Math.round(parseInt(a.slice(k, k + 2), 16) * (1 - t) + parseInt(b.slice(k, k + 2), 16) * t).toString(16).padStart(2, '0')).join('');
    const grads = new Map([...svg.querySelectorAll('linearGradient')].map(gr => [gr.id, { x1: +gr.getAttribute('x1'), x2: +gr.getAttribute('x2'), c: [...gr.querySelectorAll('stop')].map(q => hex(q.getAttribute('stop-color')) || '#999999') }]));
    const corEm = (v, x) => {
      const m = /url\(#([^)]+)\)/.exec(v || '');
      if (!m) return hex(v);
      const gr = grads.get(m[1]); if (!gr) return null;
      return mistura(gr.c[0], gr.c[gr.c.length - 1], Math.max(0, Math.min(1, (x - gr.x1) / ((gr.x2 - gr.x1) || 1))));
    };
    const camadaDe = (el, sc, fc) => {
      const col = el.closest('.secao-coluna'), tag = el.tagName.toLowerCase();
      if (tag === 'text') return col ? 'POCOS_NOMES' : 'TEXTOS';
      if (el.classList.contains('secao-faixa')) return 'LITOLOGIA_ENTRE_SONDAGENS';
      if (tag === 'circle') return col ? 'POCOS_RESULTADO' : 'PLUMA_AMOSTRAS';
      if (sc === '#1565a8' || fc === '#1565a8') return 'NIVEL_AGUA';
      if (sc === '#7a5a2e') return 'TERRENO';
      if (col) {
        if (tag === 'rect' && sc === '#ffffff') return 'LITOLOGIA_SONDAGENS';
        if (fc === '#ddeaf6' || sc === '#2b7bba') return 'FILTRO';
        if (fc === '#f3f5f4') return 'TUBO_LISO';
        return 'SONDAGENS';
      }
      if (el.getAttribute('stroke-opacity') === '.13') return 'GRADE';
      return 'MOLDURA';
    };
    svg.querySelectorAll('line, rect, path, polyline, circle, text').forEach(el => {
      if (el.closest('clipPath, defs, [clip-path]')) return;
      const tag = el.tagName.toLowerCase(), A = k => el.getAttribute(k);
      const scR = A('stroke'), fcR = A('fill');
      const sc = hex(scR)?.toLowerCase() || null, fc = (fcR && fcR.startsWith('url(')) ? 'url' : (hex(fcR)?.toLowerCase() || (tag === 'text' || (tag !== 'line' && tag !== 'polyline' && fcR == null) ? '#172621' : null));
      const camada = camadaDe(el, sc, fc), tipo = A('stroke-dasharray') ? 'TRACO' : undefined;
      const linha = (pts, fechado) => { if (sc && sc !== '#ffffff') ents.push({ t: 'poli', pts: pts.map(P), fechado, camada, cor: WGDxfSaida.aci(sc), tipo }); };
      const solido = (pts, cor) => { if (cor && cor !== '#ffffff') ents.push({ t: 'solido', pts: pts.map(P), camada, cor: WGDxfSaida.aci(cor) }); };
      if (tag === 'line') linha([[+A('x1'), +A('y1')], [+A('x2'), +A('y2')]], false);
      else if (tag === 'polyline') linha(A('points').trim().split(/\s+/).map(q => q.split(',').map(Number)), false);
      else if (tag === 'rect') {
        if (A('width') === '100%') return;
        const x = +A('x'), y = +A('y'), w = +A('width'), h = +A('height'), pts = [[x, y], [x + w, y], [x + w, y + h], [x, y + h]];
        if (fc && fc !== 'url' && (A('fill-opacity') == null || +A('fill-opacity') > .3)) solido(pts, fc);
        linha(pts, true);
      } else if (tag === 'path') {
        WGDxfSaida.caminho(A('d')).forEach(cm => {
          let pts = cm.pts;
          if (pts.length > 2 && Math.hypot(pts[0][0] - pts[pts.length - 1][0], pts[0][1] - pts[pts.length - 1][1]) < 1e-6) pts = pts.slice(0, -1);
          if (el.classList.contains('secao-faixa')) {            // faixa entre sondagens: topo de A para B, base de B para A -> quadriláteros
            const n = pts.length / 2;
            for (let k = 0; k + 1 < n; k++) {
              const q = [pts[k], pts[k + 1], pts[2 * n - 2 - k], pts[2 * n - 1 - k]];
              solido(q, corEm(fcR, (q[0][0] + q[1][0]) / 2));
            }
          } else if (fc && fc !== 'url' && fcR !== 'none' && pts.length <= 4 && cm.fechado) solido(pts, fc);
          if (scR !== 'none') linha(pts, cm.fechado);
        });
      } else if (tag === 'circle') {
        ents.push({ t: 'circ', x: X(+A('cx')), y: Y(+A('cy')), r: +A('r') / D.pxH, camada, cor: WGDxfSaida.aci(fc && fc !== '#ffffff' ? fc : (sc || '#172621')) });
      } else if (tag === 'text') {
        let x = +A('x'), y = +A('y'), rot = 0;
        const m = /rotate\(\s*(-?[\d.]+)/.exec(A('transform') || ''); if (m) rot = -(+m[1]);
        const fs = +(A('font-size') || 11), anc = A('text-anchor');
        ents.push({ t: 'texto', x: X(x), y: Y(y), alt: fs * 0.72 / D.pxH, txt: el.textContent, rot, al: anc === 'middle' ? 1 : anc === 'end' ? 2 : 0, camada, cor: WGDxfSaida.aci(fc || '#172621') });
      }
    });
    // pluma vertical: linha do limite (vermelha) e linhas de 10x, 100x... o limite, com o valor escrito
    const PV = D.PV;
    if (PV && PV.curvas) PV.curvas.forEach(cv => cv.linhas.forEach(l => {
      const pts = l.map(([s0, z]) => [s0, z * D.ev]);
      ents.push({ t: 'poli', pts, fechado: false, camada: cv.limite ? 'PLUMA_LIMITE' : 'PLUMA_ISOLINHAS', cor: cv.limite ? 1 : 30 });
      const m = pts[Math.floor(pts.length / 2)];
      if (!cv.limite && pts.length > 4) ents.push({ t: 'texto', x: m[0], y: m[1], alt: 7 / D.pxH, txt: `${fmt(cv.v, 3)} ${PV.un}`, al: 1, camada: 'PLUMA_ISOLINHAS', cor: 30 });
    }));
    ents.push({ t: 'texto', x: 0, y: Y(-30), alt: 12 / D.pxH, txt: D.titulo, camada: 'TEXTOS', cor: 7 });   // acima de A e A'
    ents.push({ t: 'texto', x: 0, y: Y(-12), alt: 8 / D.pxH, camada: 'TEXTOS', cor: 7,
      txt: `X = distância ao longo da seção (m) · Y = ${D.semDatum ? 'profundidade' : 'cota'} (m) x ${fmt(D.ev, 2)} (exagero vertical ${fmt(D.ev, 1)}x)` });
    const dxf = WGDxfSaida.escrever(ents.map(e => ({ ...e, camada: C(e.camada) })), { traco: [5 / D.pxH, 3 / D.pxH] });
    const proj = st.projetos.find(p => p.id === st.projetoId);
    const url = URL.createObjectURL(new Blob([WGDxfSaida.bytes(dxf)], { type: 'application/dxf' }));
    const a = document.createElement('a');
    a.href = url; a.download = `secao_${st.secao.letra}-${st.secao.letra}_${(proj?.nome || 'projeto').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w\-]+/g, '_')}.dxf`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

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
    a.href = url; a.download = `secao_${st.secao.letra}-${st.secao.letra}_${(proj?.nome || 'projeto').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\-]+/g, '_')}.svg`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  // ------------------------------------------------------------- papéis e equipe
  const ehAdmin = () => st.profile?.role === 'admin';
  const NOME_PAPEL = { admin: 'Administrador', tecnico: 'Técnico' };

  /** Mostra ou esconde o que só o administrador pode fazer. */
  function aplicarPapel() {
    const adm = ehAdmin();
    $('#rail-user-papel').textContent = NOME_PAPEL[st.profile?.role] || '';
    $('#nav-equipe').hidden = !adm;
    $('#btn-excluir-projeto').hidden = !adm || !st.projetoId;
    $('#btn-excluir-campanha').hidden = !adm || !st.campanhaId;
    if (!adm && !$('#view-equipe').hidden) trocarView('visao');
  }

  /** Nome da empresa no topo da tela Equipe (o bloco é criado aqui, sem precisar mexer no index.html). */
  function montarEmpresaEquipe() {
    let d = $('#equipe-empresa');
    if (!d) {
      d = document.createElement('div');
      d.id = 'equipe-empresa'; d.className = 'equipe-empresa';
      d.innerHTML = `<div class="field"><label for="equipe-empresa-nome">Nome da empresa</label>
          <input type="text" id="equipe-empresa-nome" maxlength="200" autocomplete="organization" placeholder="Nome da empresa"></div>
        <button type="button" class="btn btn-primary" id="equipe-empresa-salvar">Salvar</button>
        <span class="equipe-empresa-msg" id="equipe-empresa-msg" aria-live="polite"></span>`;
      $('#view-equipe .equipe-corpo').prepend(d);
      $('#equipe-empresa-salvar').addEventListener('click', salvarEmpresa);
      $('#equipe-empresa-nome').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); salvarEmpresa(); } });
      $('#equipe-empresa-nome').addEventListener('input', () => { $('#equipe-empresa-msg').textContent = ''; });
    }
    $('#equipe-empresa-nome').value = st.org?.name || '';
    $('#equipe-empresa-msg').textContent = '';
  }

  async function salvarEmpresa() {
    const nome = $('#equipe-empresa-nome').value.trim(), msg = $('#equipe-empresa-msg'), btn = $('#equipe-empresa-salvar');
    msg.className = 'equipe-empresa-msg';
    if (!nome) { msg.classList.add('erro'); msg.textContent = 'Digite o nome da empresa.'; return; }
    if (nome === (st.org?.name || '')) { msg.textContent = 'Nada mudou.'; return; }
    btn.disabled = true; btn.textContent = 'Salvando...';
    const { error } = await sb.rpc('wg_editar_empresa', { p_nome: nome });
    btn.disabled = false; btn.textContent = 'Salvar';
    if (error) {
      msg.classList.add('erro');
      msg.textContent = /function .*wg_editar_empresa|could not find|Função desconhecida/i.test(error.message)
        ? 'Ainda não liberado no banco: rode o 09_editar_pessoa.sql de novo (ele foi atualizado).' : error.message;
      return;
    }
    if (st.org) st.org.name = nome;
    $('#rail-user-org').textContent = nome;
    if ($('#imp-empresa')) $('#imp-empresa').textContent = nome;
    msg.classList.add('ok'); msg.textContent = 'Salvo.';
    toast('Nome da empresa salvo. Vale também para o Perfil de Sondagem.');
  }

  async function carregarEquipe() {
    montarEmpresaEquipe();
    const corpo = $('#equipe-tabela tbody');
    corpo.innerHTML = '<tr><td colspan="5" class="muted">Carregando...</td></tr>';
    const { data, error } = await sb.rpc('wg_equipe', {});
    if (error) { corpo.innerHTML = `<tr><td colspan="5" class="erro">${esc(error.message)}</td></tr>`; return; }
    st.equipe = data || [];
    corpo.innerHTML = (data || []).map(p => {
      const eu = p.id === st.user.id;
      return `<tr data-id="${esc(p.id)}">
        <td>${esc(p.nome || '—')}${eu ? ' <span class="muted">(você)</span>' : ''}</td>
        <td>${esc(p.email || '')}</td>
        <td><select class="equipe-papel" aria-label="Papel de ${esc(p.email || p.nome || '')}">
          <option value="tecnico"${p.papel === 'tecnico' ? ' selected' : ''}>Técnico</option>
          <option value="admin"${p.papel === 'admin' ? ' selected' : ''}>Administrador</option></select></td>
        <td>${fmtData(p.desde)}</td>
        <td class="num equipe-acoes"><button type="button" class="btn-mini equipe-editar">Editar</button><button type="button" class="btn-mini btn-mini-perigo equipe-remover">${eu ? 'Sair da empresa' : 'Remover'}</button></td></tr>`;
    }).join('') || '<tr><td colspan="5" class="muted">Ninguém na equipe.</td></tr>';
  }

  /** Linha da equipe vira formulário: nome e e-mail editáveis, com Salvar e Cancelar. */
  function editarLinhaEquipe(tr) {
    const p = (st.equipe || []).find(x => String(x.id) === tr.dataset.id); if (!p) return;
    const aberta = $('#equipe-tabela tr.editando'); if (aberta && aberta !== tr) carregarEquipe();   // só uma por vez
    tr.classList.add('editando');
    tr.children[0].innerHTML = `<input type="text" class="equipe-ed-nome" maxlength="200" value="${esc(p.nome || '')}" placeholder="Nome e sobrenome" aria-label="Nome">`;
    tr.children[1].innerHTML = `<input type="email" class="equipe-ed-email" maxlength="190" value="${esc(p.email || '')}" placeholder="nome@empresa.com.br" aria-label="E-mail">`;
    tr.children[4].innerHTML = `<button type="button" class="btn btn-primary btn-sm2 equipe-salvar">Salvar</button><button type="button" class="btn-mini equipe-cancelar">Cancelar</button>`;
    tr.querySelector('.equipe-ed-nome').focus();
  }

  async function salvarLinhaEquipe(tr) {
    const p = (st.equipe || []).find(x => String(x.id) === tr.dataset.id); if (!p) return;
    const nome = tr.querySelector('.equipe-ed-nome').value.trim(), email = tr.querySelector('.equipe-ed-email').value.trim().toLowerCase();
    const erro = msg => { $('#equipe-erro').textContent = msg; $('#equipe-erro').hidden = false; };
    $('#equipe-erro').hidden = true; $('#equipe-criado').hidden = true;
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { erro('Digite um e-mail válido.'); tr.querySelector('.equipe-ed-email').focus(); return; }
    if (nome === (p.nome || '') && email === (p.email || '').toLowerCase()) { carregarEquipe(); return; }   // nada mudou
    const btn = tr.querySelector('.equipe-salvar'); btn.disabled = true; btn.textContent = 'Salvando...';
    const { error } = await sb.rpc('wg_editar_pessoa', { p_usuario: p.id, p_nome: nome, p_email: email });
    if (error) {
      btn.disabled = false; btn.textContent = 'Salvar';
      erro(/function .*wg_editar_pessoa|could not find|não encontrada a função|Função desconhecida/i.test(error.message)
        ? 'A edição ainda não está liberada no banco. Rode o arquivo 09_editar_pessoa.sql e tente de novo.' : error.message);
      return;
    }
    const mudouEmail = email !== (p.email || '').toLowerCase();
    if (p.id === st.user.id && st.profile) st.profile.full_name = nome;
    if (p.id === st.user.id && mudouEmail) { st.user.email = email; $('#rail-user-email').textContent = email; }
    toast(mudouEmail ? `Salvo. A partir de agora ${nome || 'a pessoa'} entra com ${email}.` : 'Salvo.');
    carregarEquipe();
  }

  /** Senha provisória fácil de ditar: sem letras parecidas (l, 1, O, 0). */
  function gerarSenha() {
    const c = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    const v = new Uint32Array(10); crypto.getRandomValues(v);
    return [...v].map(x => c[x % c.length]).join('');
  }

  /**
   * Cria o usuário (nome, e-mail, senha provisória) e já coloca na empresa com o papel escolhido.
   * O cadastro é feito por um cliente separado, que não mexe na sessão de quem está criando.
   * Sem senha: só adiciona à equipe uma conta que já existe.
   */
  async function adicionarPessoa(ev) {
    ev.preventDefault();
    const nome = $('#equipe-nome').value.trim(), email = $('#equipe-email').value.trim().toLowerCase();
    const senha = $('#equipe-senha').value, papel = $('#equipe-papel-novo').value;
    if (!email) return;
    const erro = msg => { $('#equipe-erro').textContent = msg; $('#equipe-erro').hidden = false; };
    $('#equipe-erro').hidden = true; $('#equipe-criado').hidden = true;
    if (senha && senha.length < 6) { erro('A senha precisa ter pelo menos 6 caracteres.'); return; }
    const btn = $('#equipe-adicionar'); btn.disabled = true; btn.textContent = 'Criando...';
    let criado = false, precisaConfirmar = false;
    try {
      if (senha) {
        const aux = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY,
          { auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false, storageKey: 'webgeo-criar-usuario' } });
        const r = await aux.auth.signUp({ email, password: senha, options: { data: { full_name: nome } } });
        if (r.error && !/already registered/i.test(r.error.message)) {
          erro(/at least/i.test(r.error.message) ? 'A senha precisa ter pelo menos 6 caracteres.'
            : /validate email|invalid/i.test(r.error.message) ? 'Digite um e-mail válido.'
            : /rate limit|too many/i.test(r.error.message) ? 'O Supabase limitou a criação de contas por agora (limite de e-mails por hora). Tente de novo mais tarde.'
            : 'Não foi possível criar a conta: ' + r.error.message);
          return;
        }
        criado = !r.error && !!r.data?.user && (r.data.user.identities?.length ?? 1) > 0;  // identities vazio = e-mail já tinha conta
        precisaConfirmar = criado && !r.data.session;
      }
      const { error } = await sb.rpc('wg_adicionar_pessoa', { p_email: email, p_papel: papel });
      if (error) { erro(error.message); return; }
      const qual = papel === 'admin' ? 'Administrador' : 'Técnico';
      $('#equipe-criado').innerHTML = criado
        ? `<b>Usuário criado como ${qual}.</b> Passe para a pessoa: e-mail <code>${esc(email)}</code> e senha provisória <code>${esc(senha)}</code>.
           ${precisaConfirmar ? '<br>O Supabase está pedindo confirmação de e-mail: a pessoa precisa clicar no link que chegou no e-mail dela antes do primeiro acesso.' : ''}`
        : `<b>${esc(email)}</b> entrou na equipe como ${qual}${senha ? '. Essa pessoa já tinha conta, então a senha dela continua a mesma de antes.' : '.'}`;
      $('#equipe-criado').hidden = false;
      $('#equipe-nome').value = ''; $('#equipe-email').value = ''; $('#equipe-senha').value = ''; $('#equipe-papel-novo').value = 'tecnico';
      carregarEquipe();
    } finally {
      btn.disabled = false; btn.textContent = 'Criar usuário';
    }
  }

  // ------------------------------------------------------------- trocar a própria senha
  function abrirSenha() { $('#senha-form').reset(); $('#senha-erro').hidden = true; $('#modal-senha').hidden = false; $('#senha-nova').focus(); }
  function fecharSenha() { $('#modal-senha').hidden = true; }
  async function salvarSenha(ev) {
    ev.preventDefault();
    const a = $('#senha-nova').value, b = $('#senha-conf').value;
    const erro = m => { $('#senha-erro').textContent = m; $('#senha-erro').hidden = false; };
    if (a.length < 6) { erro('A senha precisa ter pelo menos 6 caracteres.'); return; }
    if (a !== b) { erro('As duas senhas estão diferentes.'); return; }
    const btn = $('#senha-salvar'); btn.disabled = true;
    const { error } = await sb.auth.updateUser({ password: a });
    btn.disabled = false;
    if (error) { erro(/different from the old/i.test(error.message) ? 'A nova senha precisa ser diferente da atual.' : 'Não foi possível trocar a senha: ' + error.message); return; }
    fecharSenha();
    toast('Senha trocada.');
  }

  async function mudarPapel(id, papel) {
    const { error } = await sb.rpc('wg_definir_papel', { p_usuario: id, p_papel: papel });
    if (error) { toast(error.message); carregarEquipe(); return; }
    toast('Papel alterado.');
    if (id === st.user.id) { await recarregarMeuPerfil(); return; }
    carregarEquipe();
  }

  async function removerPessoa(id) {
    const eu = id === st.user.id;
    if (!window.confirm(eu ? 'Sair da empresa? Você perde o acesso aos dados dela.' : 'Remover esta pessoa da empresa? A conta dela continua existindo, mas ela perde o acesso aos dados.')) return;
    const { error } = await sb.rpc('wg_remover_pessoa', { p_usuario: id });
    if (error) { toast(error.message); return; }
    toast(eu ? 'Você saiu da empresa.' : 'Pessoa removida da equipe.');
    if (eu) { await carregarPerfilEEntrar(); return; }
    carregarEquipe();
  }

  /** Relê o próprio papel (depois de mudar o próprio papel na tela Equipe). */
  async function recarregarMeuPerfil() {
    const { data } = await sb.from('profiles').select('*').eq('id', st.user.id).single();
    if (data) st.profile = data;
    aplicarPapel();
    if (ehAdmin()) carregarEquipe();
  }

  // ------------------------------------------------------------- excluir projeto ou campanha (só administrador)
  function abrirExcluir(tipo) {
    const proj = st.projetos.find(p => p.id === st.projetoId);
    const camp = st.campanhas.find(c => c.id === st.campanhaId);
    if (tipo === 'projeto' && !proj) return;
    if (tipo === 'campanha' && !camp) return;
    const n = (x, um, varios) => `${x} ${x === 1 ? um : varios}`;
    const itens = [];
    let nome, nota;
    if (tipo === 'projeto') {
      nome = proj.nome;
      const pocos = st.pocos.filter(p => !p.virtual).length;
      if (pocos) itens.push(`${n(pocos, 'poço', 'poços')} cadastrados no WebGeo`);
      if (st.campanhas.length) itens.push(`${n(st.campanhas.length, 'campanha', 'campanhas')} e ${n(st.resultados.length, 'resultado', 'resultados')} de laboratório`);
      if (st.medicoes.length) itens.push(`${n(st.medicoes.length, 'medição', 'medições')} de nível d'água`);
      if (st.planta) itens.push(`A planta "${st.planta.nome}"`);
      if (!itens.length) itens.push('Este projeto não tem dados no WebGeo.');
      if (st.totalFichas) itens.push(`${n(st.totalFichas, 'ficha do Perfil fica', 'fichas do Perfil ficam')} sem projeto (não ${st.totalFichas === 1 ? 'é apagada' : 'são apagadas'})`);
      nota = 'O projeto é o mesmo do Perfil de Sondagem: ele some de lá também. As fichas de sondagem <b>não são apagadas</b>, ficam "Sem projeto" no Perfil.';
    } else {
      nome = camp.codigo;
      const res = st.resultados.filter(r => r.campanha_id === camp.id).length, med = st.medicoes.filter(m => m.campanha_id === camp.id).length;
      itens.push(`${n(res, 'resultado', 'resultados')} de laboratório da campanha ${camp.codigo} (${fmtData(camp.data_inicio)})`);
      if (med) itens.push(`${n(med, 'medição', 'medições')} de nível d'água dessa campanha`);
      nota = 'Os poços, as outras campanhas e as fichas do Perfil continuam. Para refazer, importe a planilha da campanha de novo.';
    }
    st.excluindo = { tipo, id: tipo === 'projeto' ? proj.id : camp.id, nome };
    $('#modal-excluir-tit').textContent = tipo === 'projeto' ? 'Excluir projeto' : 'Excluir campanha';
    $('#excluir-tipo').textContent = tipo === 'projeto' ? 'o projeto' : 'a campanha';
    $('#excluir-nome').textContent = nome;
    $('#excluir-lista').innerHTML = itens.map(i => `<li>${esc(i)}</li>`).join('');
    $('#excluir-nota').innerHTML = nota;
    $('#excluir-rotulo').textContent = tipo === 'projeto' ? 'Para confirmar, digite o nome do projeto' : 'Para confirmar, digite o código da campanha';
    $('#excluir-confirmar').textContent = tipo === 'projeto' ? 'Excluir projeto' : 'Excluir campanha';
    $('#excluir-confirma').value = '';
    $('#excluir-confirmar').disabled = true;
    $('#excluir-erro').hidden = true;
    $('#modal-excluir').hidden = false;
    $('#excluir-confirma').focus();
  }
  function fecharExcluir() { $('#modal-excluir').hidden = true; st.excluindo = null; }
  const confirmouExclusao = () => !!st.excluindo && $('#excluir-confirma').value.trim().toUpperCase() === String(st.excluindo.nome).trim().toUpperCase();

  async function excluirConfirmado() {
    const x = st.excluindo;
    if (!x || !confirmouExclusao()) return;
    const btn = $('#excluir-confirmar'), rotulo = btn.textContent; btn.disabled = true; btn.textContent = 'Excluindo...';
    // .select() devolve a linha apagada: se voltar vazio, o banco não deixou apagar
    const { data, error } = await sb.from(x.tipo === 'projeto' ? 'projetos' : 'wg_campanha').delete().eq('id', x.id).select();
    btn.textContent = rotulo;
    if (error || !data || !data.length) {
      console.error(error);
      $('#excluir-erro').textContent = 'Não foi possível excluir: ' + (error?.message || 'o banco não permitiu (só administradores excluem, e só da própria empresa).');
      $('#excluir-erro').hidden = false; btn.disabled = false;
      return;
    }
    fecharExcluir();
    if (x.tipo === 'projeto') {
      toast(`Projeto "${x.nome}" excluído.`);
      st.projetoId = null; st.campanhaId = null; st.parametroId = null; st.popupAberto = null; lembrarProjeto('');
      await carregarProjetos();
    } else {
      toast(`Campanha ${x.nome} excluída.`);
      st.campanhaId = null;
      await carregarProjeto();
    }
  }

  // ------------------------------------------------------------- revisão da importação
  /** O que o projeto de destino já tem, por código (para comparar com a planilha). */
  async function historicoDoProjeto(projetoId) {
    const [campanhas, pocos, resultados] = await Promise.all([
      buscarTudo(() => sb.from('wg_campanha').select('id, codigo, data_inicio').eq('projeto_id', projetoId).order('id')),
      buscarTudo(() => sb.from('wg_poco').select('id, codigo, cota_topo, latitude, longitude, profundidade').eq('projeto_id', projetoId).order('id')),
      buscarTudo(() => sb.from('wg_vw_resultado').select('poco, parametro, campanha, data_coleta, data_inicio, valor, menor_que_lq, lq').eq('projeto_id', projetoId).order('id'))
    ]);
    const codPoco = new Map(pocos.map(p => [p.id, p.codigo])), codCamp = new Map(campanhas.map(c => [c.id, c]));
    const medicoes = campanhas.length
      ? await buscarTudo(() => sb.from('wg_medicao_campo').select('poco_id, campanha_id, data_medicao, nivel_agua').in('campanha_id', campanhas.map(c => c.id)).order('id'))
      : [];
    return {
      campanhas,
      pocos,
      resultados: resultados.map(r => ({ ...r, data: r.data_coleta || r.data_inicio })),
      medicoes: medicoes.map(m => ({ poco: codPoco.get(m.poco_id), campanha: codCamp.get(m.campanha_id)?.codigo,
        data: m.data_medicao || codCamp.get(m.campanha_id)?.data_inicio, nivel_agua: m.nivel_agua }))
    };
  }

  function mostrarRevisao(rev) {
    const box = $('#imp-conferencia');
    const a = rev.alertas;
    let html = '';
    if (rev.substitui) html += `<p class="conf-substitui">A campanha <b>${esc(rev.substitui.campanha)}</b> já existe neste projeto: os valores dela serão atualizados com os desta planilha (${rev.substitui.resultados} resultado(s) gravados hoje).</p>`;
    if (!a.length) html += '<p class="conf-ok">Nada fora do comum em relação ao histórico do projeto.</p>';
    else {
      const altos = a.filter(x => x.nivel === 'alto').length;
      html += `<h4>Conferência: ${a.length} ponto(s) para olhar${altos ? `, ${altos} importante(s)` : ''}</h4>
        <div class="tbl-wrap"><table class="dtbl conf-tabela"><thead><tr><th></th><th>Poço</th><th>Item</th><th class="num">Antes</th><th class="num">Agora</th><th>Por quê</th></tr></thead><tbody>
        ${a.map(x => `<tr class="conf-${x.nivel}"><td><span class="conf-marca" title="${x.nivel === 'alto' ? 'Importante' : 'Atenção'}"></span></td>
          <td class="cod">${esc(x.poco)}</td><td>${esc(x.item)}</td><td class="num">${esc(x.antes)}</td><td class="num"><b>${esc(x.agora)}</b></td><td>${esc(x.motivo)}</td></tr>`).join('')}
        </tbody></table></div>
        <label class="chk conf-ciente"><input type="checkbox" id="imp-conferi"><span class="box"></span>Conferi esses pontos e quero gravar assim mesmo</label>`;
    }
    box.innerHTML = html;
    box.hidden = false;
    const conf = $('#imp-conferi');
    if (conf) conf.addEventListener('change', () => { $('#btn-gravar').disabled = !conf.checked; });
    $('#btn-gravar').disabled = !!a.length;
  }

  // ------------------------------------------------------------- relatório em PDF (impressão do navegador)
  function abrirRelatorio() {
    if (!st.projetoId || !st.campanhaId) { toast('Escolha um projeto e uma campanha.'); return; }
    $('#rel-op-secao').disabled = !st.secao.linha;
    if (!st.secao.linha) $('#rel-op-secao').checked = false;
    $('#modal-relatorio').hidden = false;
  }
  function fecharRelatorio() { $('#modal-relatorio').hidden = true; }

  /** Tabela poços x parâmetros da campanha, com o valor orientador no cabeçalho e o que passou em destaque. */
  function tabelaRelatorio() {
    const camp = st.campanhas.find(c => c.id === st.campanhaId);
    const res = st.resultados.filter(r => r.campanha_id === st.campanhaId);
    const params = st.parametros.filter(p => res.some(r => r.parametro_id === p.id));
    const pocos = st.pocos.filter(p => res.some(r => r.poco_id === p.id) || st.medicoes.some(m => m.poco_id === p.id && m.campanha_id === st.campanhaId && m.nivel_agua != null));
    const mapa_ = new Map(res.map(r => [r.poco_id + '|' + r.parametro_id, r]));
    const na = new Map(st.medicoes.filter(m => m.campanha_id === st.campanhaId).map(m => [m.poco_id, m]));
    const POR_TABELA = 7;
    let html = '';
    for (let k = 0; k < Math.max(1, params.length); k += POR_TABELA) {
      const ps = params.slice(k, k + POR_TABELA);
      html += `<table class="rel-tabela"><thead><tr><th>Poço</th>${k === 0 ? '<th class="num">N.A. (m)</th>' : ''}${ps.map(p =>
        `<th class="num">${esc(p.nome)}<small>${esc(p.unidade || '')}${p.valor_orientador != null ? ` · VI ${fmt(p.valor_orientador)}` : ''}</small></th>`).join('')}</tr></thead><tbody>`
        + pocos.map(p => `<tr><td class="cod">${esc(p.codigo)}</td>${k === 0 ? `<td class="num">${na.has(p.id) && na.get(p.id).nivel_agua != null ? fmt(na.get(p.id).nivel_agua, 2) : '-'}</td>` : ''}${ps.map(q => {
          const r = mapa_.get(p.id + '|' + q.id);
          return `<td class="num${r?.acima_vi ? ' rel-acima' : r?.menor_que_lq ? ' rel-lq' : ''}">${textoValor(r)}</td>`;
        }).join('')}</tr>`).join('') + '</tbody></table>';
    }
    const acima = res.filter(r => r.acima_vi).length;
    return `<h2>Resultados da campanha ${esc(camp?.codigo || '')}</h2>
      <p class="rel-texto">${pocos.length} poço(s), ${params.length} parâmetro(s), ${res.length} resultado(s); ${acima} acima do valor orientador.
      Em <b class="rel-acima-txt">vermelho</b>: igual ou acima do valor orientador. "&lt; x": abaixo do limite de quantificação x.</p>${html}`;
  }

  /** Monta o relatório na própria página e chama a impressão (o usuário escolhe "Salvar como PDF"). */
  async function prepararRelatorio(op) {
    const proj = st.projetos.find(p => p.id === st.projetoId);
    const camp = st.campanhas.find(c => c.id === st.campanhaId);
    const param = st.parametros.find(p => p.id === st.parametroId);
    const limiar = limiarAtual(param), un = param?.unidade || 'µg/L';
    const agora = new Date();
    $('#rel-cab').innerHTML = `
      <div class="rel-marca">WebGeo · Relatório de monitoramento</div>
      <h1>${esc(proj?.nome || '')}</h1>
      <table class="rel-ficha">
        <tr><th>Empresa</th><td>${esc(st.org?.name || '')}</td><th>Campanha</th><td>${esc(camp?.codigo || '')} · ${fmtData(camp?.data_inicio)}${camp?.data_fim ? ' a ' + fmtData(camp.data_fim) : ''}</td></tr>
        <tr><th>Laboratório</th><td>${esc(camp?.laboratorio || '-')}</td><th>Parâmetro do mapa</th><td>${esc(param?.nome || '-')}${param?.valor_orientador != null ? ` · VI ${fmt(param.valor_orientador)} ${esc(un)}` : ''}</td></tr>
        <tr><th>Emitido em</th><td>${agora.toLocaleDateString('pt-BR')} ${agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}</td><th>Por</th><td>${esc(st.profile?.full_name || st.user?.email || '')}</td></tr>
      </table>`;
    $('#rel-tabela').innerHTML = op.tabela ? tabelaRelatorio() : '';
    $('#rel-anot').innerHTML = op.anot ? anotacaoRelatorio() : '';
    const notas = [
      `Valores orientadores: ${esc(st.parametros.find(p => p.referencia)?.referencia || 'CETESB DD 125/2021/E')}.`,
      op.pluma ? `Pluma de ${esc(param?.nome || '')}: interpolação IDW em escala logarítmica (p = ${st.pluma.p}), célula de ${st.pluma.cel} m, limite de ${fmt(limiar, 4)} ${esc(un)}. Valores abaixo do LQ entram como metade do LQ; poços a até 5 m um do outro contam como um ponto, com o maior valor.` : '',
      st.pot.ativa && st.potCalc?.sup ? `Potenciométrico: carga hidráulica = cota do topo − N.A.; superfície por spline de placa fina; ${esc($('#pot-resumo').textContent)}.` : '',
      op.secao && st.secao.linha ? `${st.secao.lista.length > 1 ? 'Seções geológicas ' + st.secao.lista.map(x => nomeSecao(x.letra)).join(', ') : 'Seção geológica ' + nomeSecao(st.secao.letra)}: litologia das fichas do Perfil de Sondagem; o preenchimento entre as sondagens é uma interpretação automática.` : '',
      'Mapa: imagem de satélite Esri World Imagery.'
    ].filter(Boolean);
    $('#rel-notas').innerHTML = op.notas ? `<h2>Notas</h2><ul>${notas.map(n => `<li>${n}</li>`).join('')}</ul>` : '';
    // o relatório sai sempre no tema claro (papel branco), mesmo com a tela no modo escuro
    st.temaAntes = document.documentElement.getAttribute('data-theme');
    document.documentElement.setAttribute('data-theme', 'light');
    const corpo = document.body.classList;
    corpo.add('modo-relatorio');
    corpo.toggle('rel-sem-mapa', !op.mapa);
    corpo.toggle('rel-sem-secao', !op.secao);
    corpo.toggle('rel-sem-pluma', !op.pluma);
    trocarView('visao');
    await new Promise(r => setTimeout(r, 50));
    mapa.invalidateSize();
    render();                                   // gráficos e desenhos com as cores do tema claro
    if (op.mapa && st.pontos.length) mapa.fitBounds(st.pontos, { padding: [20, 20], maxZoom: 19, animate: false });
    atualizarRotulos();
    if (st.secao.linha) renderSecao();
    secoesExtrasRelatorio(op);
    Object.values(graficos).forEach(g => g?.resize?.());
    await esperarImagemDoMapa(4000);
  }

  /** Espera os pedaços da imagem de satélite carregarem (no máximo `ms`). */
  function esperarImagemDoMapa(ms) {
    return new Promise(fim => {
      const t = setTimeout(fim, ms);
      const pendentes = () => [...document.querySelectorAll('#mapa .leaflet-tile')].filter(i => !i.complete).length;
      const ver = () => { if (!pendentes()) { clearTimeout(t); setTimeout(fim, 300); } else setTimeout(ver, 150); };
      setTimeout(ver, 200);
    });
  }

  function encerrarRelatorio() {
    if (!document.body.classList.contains('modo-relatorio')) return;
    document.body.classList.remove('modo-relatorio', 'rel-sem-mapa', 'rel-sem-secao', 'rel-sem-pluma');
    if ($('#rel-secoes-extra')) $('#rel-secoes-extra').innerHTML = '';
    if (st.secao.ativaAntesRel != null) { st.secao.ativa = st.secao.ativaAntesRel; st.secao.ativaAntesRel = null; sincronizarSecao(); }
    if (st.temaAntes) document.documentElement.setAttribute('data-theme', st.temaAntes); else document.documentElement.removeAttribute('data-theme');
    setTimeout(() => { mapa.invalidateSize(); render(); }, 50);
  }

  async function gerarRelatorio() {
    const op = { mapa: $('#rel-op-mapa').checked, secao: $('#rel-op-secao').checked, pluma: $('#rel-op-pluma').checked,
      tabela: $('#rel-op-tabela').checked, notas: $('#rel-op-notas').checked, anot: $('#rel-op-anot').checked };
    const btn = $('#rel-gerar'); btn.disabled = true; btn.textContent = 'Preparando...';
    await prepararRelatorio(op);
    btn.disabled = false; btn.textContent = 'Gerar PDF';
    fecharRelatorio();
    window.print();                 // a janela do navegador abre aqui; escolha "Salvar como PDF"
    setTimeout(encerrarRelatorio, 500);
  }

  function renderEvolucao(param) {
    const calc = st.calc, un = param?.unidade || 'µg/L';
    const sel = plumaDaCampanha(st.campanhaId), ini = primeiraComArea();
    const fa = a => a == null ? '—' : fmt(Math.round(a), 0);
    const semLimiar = !calc || calc.limiar == null;
    $('#pluma-hint').textContent = calc
      ? `IDW em escala log, p = ${st.pluma.p} · célula ${calc.g.cel} m · limite ${semLimiar ? '—' : fmt(calc.limiar, 4) + ' ' + un} · poços a até 5 m contam como um ponto (maior valor)`
      : 'sem resultados deste parâmetro';

    $('#kpi-a1-lab').textContent = ini ? `Área na ${ini.campanha.codigo}` : 'Área na 1ª campanha';
    $('#kpi-a2-lab').textContent = sel ? `Área na ${sel.campanha.codigo}` : 'Área na campanha';
    $('#kpi-a1').textContent = fa(ini?.area);
    $('#kpi-a2').textContent = fa(sel?.area);
    const sub = semLimiar ? 'defina o limite' : `acima de ${fmt(calc.limiar, 4)} ${un}`;
    $('#kpi-a1-sub').textContent = ini ? sub : 'precisa de 3+ pontos com resultado';
    $('#kpi-a2-sub').textContent = sel?.v ? sub : 'precisa de 3+ pontos com resultado';

    const comparar = ini && sel && ini !== sel && ini.area != null && sel.area != null;
    if (comparar && ini.area > 0) {
      const red = (1 - sel.area / ini.area) * 100;
      $('#kpi-red').textContent = (red >= 0 ? '' : '+') + fmt(Math.abs(red), 1) + '%';
      $('#kpi-red-sub').textContent = red >= 0 ? `redução entre ${ini.campanha.codigo} e ${sel.campanha.codigo}` : `AUMENTO entre ${ini.campanha.codigo} e ${sel.campanha.codigo}`;
    } else {
      $('#kpi-red').textContent = '—';
      $('#kpi-red-sub').textContent = comparar ? `sem área acima do limite na ${ini.campanha.codigo}` : 'escolha outra campanha para comparar com a 1ª';
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

  // ------------------------------------------------------------- converter UTM -> latitude/longitude (sem gravar nada)
  /** Referências para as coordenadas: fuso e centro dos poços do projeto de destino e dos poços do projeto aberto. */
  async function contextoCoordenadas(destino) {
    const existentes = new Set(), comCoord = [];
    if (destino) (await buscarTudo(() => sb.from('wg_poco').select('codigo,latitude,longitude').eq('projeto_id', destino.id).order('codigo')))
      .forEach(p => { existentes.add(p.codigo); if (p.latitude != null && p.longitude != null) comCoord.push(p); });
    const centro = arr => arr.length ? { lat: arr.reduce((a, p) => a + +p.latitude, 0) / arr.length, lon: arr.reduce((a, p) => a + +p.longitude, 0) / arr.length } : null;
    const centroProjeto = centro(comCoord);
    const fusoProjeto = centroProjeto ? { zona: WGDxf.zonaDaLongitude(centroProjeto.lon), sul: centroProjeto.lat < 0 } : null;
    const referencia = centro(st.pocos.filter(p => p.latitude != null && p.longitude != null && !p.virtual));
    return { existentes, fusoProjeto, centroProjeto, referencia };
  }

  async function converterUtmArquivo(arquivo) {
    const msg = $('#imp-status');
    msg.textContent = 'Convertendo ' + arquivo.name + '...';
    try {
      const wb = XLSX.read(await arquivo.arrayBuffer(), { type: 'array', cellDates: true });
      const pj = WGImport.lerProjeto(XLSX, wb);
      const nomeK = s => String(s || '').trim().toLowerCase();
      const achado = pj.projeto ? st.projetos.find(p => nomeK(p.nome) === nomeK(pj.projeto)) : st.projetos.find(p => p.id === st.projetoId);
      const ctx = await contextoCoordenadas(achado ? { id: achado.id, nome: achado.nome } : null);
      const { n, rel, erro } = WGImport.converterPlanilha(XLSX, wb, { parametros: st.parametros, pocosExistentes: ctx.existentes, ...ctx });
      const avisos = rel.avisos.filter(a => /UTM|fuso|E incompleto/i.test(a.msg)).map(a => a.msg);
      const errosCoord = rel.erros.filter(e => /coordenada|UTM|E com 5|latitude|longitude/i.test(e.msg));
      if (erro || !n) {
        msg.innerHTML = `<span class="erro">Nada para converter:</span> ${esc(erro || 'nenhum poço em UTM na aba Pocos (as coordenadas já estão em graus ou estão vazias)')}`
          + (errosCoord.length ? `<br>${errosCoord.slice(0, 5).map(e => esc(`linha ${e.linha}: ${e.msg}`)).join('<br>')}` : '');
        return;
      }
      const nome = arquivo.name.replace(/\.xlsx?$/i, '') + '_latlong.xlsx';
      XLSX.writeFile(wb, nome);
      msg.innerHTML = `<span class="ok">${n} poço(s) convertido(s) para latitude/longitude.</span> Baixado: <b>${esc(nome)}</b> (o E/N original ficou nas colunas novas no fim da aba Pocos). Confira no mapa depois de importar.`
        + (avisos.length ? `<br><span class="muted">${avisos.map(esc).join('<br>')}</span>` : '')
        + (errosCoord.length ? `<br><span class="erro">${errosCoord.length} linha(s) não convertida(s):</span> ${errosCoord.slice(0, 5).map(e => esc(`linha ${e.linha}: ${e.msg}`)).join('<br>')}` : '');
    } catch (e) {
      console.error(e);
      msg.innerHTML = '<span class="erro">Não consegui converter o arquivo:</span> ' + esc(e.message || e);
    }
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
      const cc = await contextoCoordenadas(destino);
      cc.existentes.forEach(c => existentes.add(c));
      const { fusoProjeto, centroProjeto, referencia } = cc;
      const { payload, rel } = WGImport.ler(XLSX, wb, arquivo.name, { parametros: st.parametros, pocosExistentes: existentes, projetoDestino: destino, fusoProjeto, centroProjeto, referencia });
      st.importacao = { payload, rel, novo };
      mostrarRelatorio(rel, payload, novo);
      if (!rel.erros.length) {                 // confere com o que o projeto já tem antes de deixar gravar
        $('#imp-conferencia').hidden = false;
        $('#imp-conferencia').innerHTML = '<p class="muted">Conferindo com o histórico do projeto...</p>';
        $('#btn-gravar').disabled = true;
        try {
          const hist = destino ? await historicoDoProjeto(destino.id) : null;
          if (st.importacao?.payload !== payload) return;   // outra planilha foi escolhida no meio
          st.importacao.revisao = WGRevisao.conferir(payload, hist, st.parametros);
          mostrarRevisao(st.importacao.revisao);
        } catch (e) {
          console.error(e);
          $('#imp-conferencia').innerHTML = '<p class="muted">Não consegui conferir com o histórico agora (' + esc(e.message || e) + '). Dá para gravar mesmo assim.</p>';
          $('#btn-gravar').disabled = false;
        }
      }
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
    $('#imp-conferencia').hidden = true;
    $('#imp-conferencia').innerHTML = '';
    $('#imp-status').textContent = '';
  }

  return { iniciar, _teste: { mapa: () => mapa, estado: st, prepararRelatorio: op => prepararRelatorio(op), encerrarRelatorio: () => encerrarRelatorio() } }; // _teste: usado só pelos testes automáticos
})();
