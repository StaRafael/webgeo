# WebGeo

Portal de monitoramento ambiental de poços (PM/PMN), irmão do **Perfil de Sondagem**.

- HTML + CSS + JavaScript puro, sem framework e sem etapa de build
- Supabase: o mesmo projeto do Perfil (login, empresas e **projetos** compartilhados)
- Vercel: publicação automática a cada push

## O que faz
- **Mapa do projeto:** os poços e todas as fichas do Perfil com coordenada.
- **Planilha do laboratório:** upload da planilha modelo, uma por campanha, com validação por aba e linha. "< 1,00" fica gravado como abaixo do LQ.
- **Resultados:** situação de cada poço frente ao valor orientador CETESB.
- **Pluma (IDW):** mostra a área acima do limiar, o centróide e a evolução entre campanhas.
- **Popup do poço:** perfil litológico e construtivo, vindo da ficha do Perfil.
- **Projetos:** os mesmos do Perfil; dá para criar e excluir pelo menu.
- **Mapa potenciométrico:** curvas equipotenciais e setas de fluxo da campanha, a partir da cota do topo e do N.A.
- **Seção geológica:** corte vertical A–A' traçado no mapa, com a litologia das fichas do Perfil, filtros, nível d'água e preenchimento das camadas entre as sondagens.
- **Relatório em PDF:** resumo, mapa, seção, evolução da pluma e tabela de resultados comparada com a CETESB.
- **Papéis e equipe:** administrador e técnico; o administrador cuida da equipe pela tela Equipe.
- **Conferência da importação:** alerta valores fora do histórico antes de gravar.
- **Análise ambiental:** massa dissolvida e centro de massa da pluma por campanha, sentido e velocidade do fluxo, tendência de Mann-Kendall por poço, meia-vida e tempo até o VI.
- **Planta em DXF:** a planta do projeto por cima do satélite, com camadas, cor e ajuste de posição por pontos.

## Como a pluma é calculada
- **Interpolação:** IDW em escala logarítmica (interpola o log da concentração), com todos os poços e potência p ajustável (padrão 2). Em log, a pluma fica em volta dos poços acima do limiar e termina antes dos poços limpos.
- **Abaixo do LQ:** entra como metade do LQ. Sem LQ no laudo, entra como 1/10 do valor orientador.
- **Poços no mesmo ponto:** poços a até 5 m um do outro (ex.: PM-26, PMN-26A, PMN-26B) contam como um ponto, com o maior valor. Para ver um nível só, use o filtro de rede.
- **Fora da rede de poços:** o valor decai suavemente até o piso numa distância igual ao espaçamento local, fechando a pluma em curva.
- **Desenho:** só é pintado o que passa do limiar (padrão = VI). Grade de 2 m; a linha do limiar é emendada e suavizada.

## Como a análise ambiental é calculada
- **Massa dissolvida:** Σ C × área da célula × b × ne nas células da pluma (mesmo limiar, IDW e rede do mapa). b, ne e K são digitados na página e ficam salvos no navegador, por projeto.
- **Centro de massa:** média das posições ponderada pela concentração; o deslocamento entre campanhas é comparado com o sentido do fluxo.
- **Fluxo por campanha:** a mesma superfície do potenciométrico, campanha a campanha. Velocidade = K × i ÷ ne.
- **Mann-Kendall:** 4+ campanhas por poço, classes do MAROS; resultados < LQ entram como empate abaixo de qualquer valor medido.
- **Atenuação:** reta de ln(C) × tempo (3+ valores quantificados): k, meia-vida e tempo até o VI.

## Como o potenciométrico é calculado
- **Carga hidráulica:** cota do topo do poço menos o N.A. medido na campanha (aba Campo da planilha).
- **Poços usados:** uma rede por vez (padrão PM, a rasa), para não misturar níveis do aquífero. Poços a até 5 m contam como um ponto, com a média.
- **Superfície:** plano regional + spline de placa fina (thin plate spline). Passa exatamente pelos poços; a suavização opcional amortece valores destoantes.
- **Desenho:** só dentro da rede de poços. As setas apontam para onde a carga diminui (perpendiculares às curvas).
- **Resumo:** gradiente médio (m/m) e sentido médio do fluxo.

Para instalar e publicar, veja **PASSO-A-PASSO.md**.
