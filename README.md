# WebGeo

Portal de monitoramento ambiental de poços (PM/PMN), irmão do **Perfil de Sondagem**.

- HTML + CSS + JavaScript puro, sem framework e sem etapa de build
- Supabase (nuvem): o mesmo projeto do Perfil (login, empresas e **projetos** compartilhados)
- Vercel: hospeda o site

## Publicação e banco de dados

**Site:** publicado no Vercel. Os arquivos listados no `.vercelignore` (pasta `banco/`,
`dados-exemplo/`, `*.md`, `DWG/`) não vão para o site.

**Banco:** o Supabase em `https://wbocmpekamjfwkwmxtjd.supabase.co`. O endereço e a chave
pública (publishable) ficam no `config.js`. A chave é pública por natureza; quem protege os
dados de cada empresa são as regras de segurança (RLS) do banco.

**Mudanças no banco vão como arquivo `.sql` novo em `banco/`** (ex.: `banco/10_minha_mudanca.sql`),
rodado à mão no **SQL Editor** do Supabase:
- rode cada arquivo **uma vez**, em ordem de nome (por isso o número na frente);
- para mudar algo já aplicado, crie outro arquivo em vez de editar o antigo;
- não use comandos do psql (linhas começando com `\`).

## O que faz
- **Mapa do projeto:** os poços e todas as fichas do Perfil com coordenada.
- **Mapa: satélite, malha UTM e rosa dos ventos:** botões para ligar/desligar a imagem de satélite e a malha de coordenadas UTM (pontilhada, E no alto, N na esquerda, passo conforme o zoom); rosa dos ventos sempre visível; saem também no PDF.
- **Planilha do laboratório:** upload da planilha modelo, uma por campanha, com validação por aba e linha. "< 1,00" fica gravado como abaixo do LQ.
- **Resultados:** situação de cada poço frente ao valor orientador CETESB.
- **Pluma (IDW):** mostra a área acima do limite, o centróide e a evolução entre campanhas.
- **Popup do poço:** perfil litológico e construtivo, vindo da ficha do Perfil.
- **Projetos:** os mesmos do Perfil; dá para criar e excluir pelo menu.
- **Mapa potenciométrico:** curvas equipotenciais e setas de fluxo da campanha, a partir da cota do topo e do N.A.
- **Seções geológicas:** cortes A–A', B–B', C–C'... traçados no mapa (abas para trocar; todas no mapa e no PDF), com a litologia das fichas do Perfil, filtros, nível d'água, preenchimento das camadas entre as sondagens e a pluma vertical do parâmetro escolhido; baixa em SVG ou DXF (com camadas, para o AutoCAD).
- **Diagrama de cercas 3D:** as sondagens em 3D, na posição e cota reais, com painéis de litologia entre elas e o nível d'água (WebGL próprio, sem biblioteca).
- **Anotações do projeto:** um texto por projeto para a equipe, salvo sozinho, com aviso quando duas pessoas editam ao mesmo tempo; pode entrar no relatório PDF.
- **Relatório em PDF:** resumo, mapa, seção, evolução da pluma e tabela de resultados comparada com a CETESB.
- **Papéis e equipe:** administrador e técnico; o administrador cuida da equipe pela tela Equipe (cria usuários, muda papel, edita nome e e-mail das pessoas e o nome da empresa).
- **Coordenadas na planilha:** graus decimais ou UTM em metros (E/N, com coluna Fuso opcional); UTM é convertido para graus na importação (E de 5 dígitos, tirado de planta de CAD, é completado); o botão "Converter UTM em lat/long" devolve a própria planilha já em graus.
- **Prancha A3 (PDF):** mapa ou seção em folha A3 com rosa dos ventos, legenda, escala e carimbo com o logo da empresa; visualização na tela (com zoom) antes de baixar ou imprimir.
- **Mapa em DXF:** planta, poços e pluma da campanha em UTM, com camadas, para o AutoCAD.
- **Editar poço:** pelo popup do poço no mapa: situação (ativo, seco, danificado/destruído, desativado) e posição, arrastando o pino, clicando no mapa ou digitando lat/long ou UTM E/N.
- **Conferência da importação:** alerta valores fora do histórico antes de gravar.
- **Desenhar e medir:** no mapa, reta, mão livre, medida de distância (trechos e total) e de área, borracha, cores e espessuras; salvos no projeto e incluídos na prancha PDF e no DXF.
- **Modo campo (GPS):** no celular, leva até cada poço sobre a imagem do drone (GeoTIFF ou imagem + world file, em UTM); ficha de campo com localizado/não localizado, situação (seco, danificado/destruído…), N.A., descrição, correção da coordenada pelo GPS e fotos. Sem SQL novo.
- **Fotos 360°:** fotos panorâmicas marcadas no mapa, com visor próprio (WebGL, sem biblioteca), GPS lido da câmera e redução automática para até 1,8 MB.
- **Planta em DXF:** a planta do projeto por cima do satélite, com camadas, cor e ajuste de posição por pontos.

## Como a pluma é calculada
- **Linha do limite ajustável:** botão "Editar linha da pluma" para arrastar, criar e apagar pontos e escolher trecho contínuo/tracejado; a pluma, a área, a prancha e o DXF seguem a linha ajustada.
- **Linha do limite:** contínua onde passa entre um poço acima e um poço abaixo do limite (triângulos de Delaunay dos poços), tracejada onde é inferida.
- **Interpolação:** IDW em escala logarítmica (interpola o log da concentração), com todos os poços e potência p ajustável (padrão 2). Em log, a pluma fica em volta dos poços acima do limite e termina antes dos poços limpos.
- **Abaixo do LQ:** entra como metade do LQ. Sem LQ no laudo, entra como 1/10 do valor orientador.
- **Poços no mesmo ponto:** poços a até 5 m um do outro (ex.: PM-26, PMN-26A, PMN-26B) contam como um ponto, com o maior valor. Para ver um nível só, use o filtro de rede.
- **Fora da rede de poços:** o valor decai suavemente até o piso numa distância igual ao espaçamento local, fechando a pluma em curva.
- **Desenho:** só é pintado o que passa do limite (padrão = VI). Grade de 2 m; a linha do limite é emendada e suavizada.

## Como a pluma vertical (na seção) é calculada
- **Amostras:** cada poço da seção com resultado do parâmetro na campanha entra no meio do filtro (sem filtro cadastrado: no meio entre o N.A. e o fundo). Poços multinível (PM, PMN...) dão a variação com a profundidade.
- **Perfil em cada poço:** do valor medido no N.A. caindo até o limite no fundo do filtro (grupo multinível: passa por cada amostra). Entre os poços, IDW em escala log na mesma profundidade abaixo do N.A. Abaixo do LQ e piso iguais aos da pluma do mapa.
- **Linha do limite:** vermelha, contínua e suavizada; vai do fundo do poço contaminado até o N.A.
- **Onde é pintada:** só na zona saturada, do N.A. até o fundo das sondagens (um poço mais fundo abre uma faixa que estreita para os lados). Mesma escala de cores e mesma linha do limite do mapa.
- **Limitação:** se todos os poços amostram a mesma profundidade, a pluma mostra só a variação lateral; a nota da seção avisa.

## Como o potenciométrico é calculado
- **Carga hidráulica:** cota do topo do poço menos o N.A. medido na campanha (aba Campo da planilha).
- **Poços usados:** uma rede por vez (padrão PM, a rasa), para não misturar níveis do aquífero. Poços a até 5 m contam como um ponto, com a média.
- **Superfície:** plano regional + spline de placa fina (thin plate spline). Passa exatamente pelos poços; a suavização opcional amortece valores destoantes.
- **Desenho:** só dentro da rede de poços. As setas apontam para onde a carga diminui (perpendiculares às curvas).
- **Resumo:** gradiente médio (m/m) e sentido médio do fluxo.

Para instalar e publicar, veja **PASSO-A-PASSO.md**.
