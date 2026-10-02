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
- **Planta em DXF:** a planta do projeto por cima do satélite, com camadas, cor e ajuste de posição por pontos.

## Como a pluma é calculada
- **Interpolação:** IDW em escala logarítmica (interpola o log da concentração), com todos os poços e potência p ajustável (padrão 2). Em log, a pluma fica em volta dos poços acima do limiar e termina antes dos poços limpos.
- **Abaixo do LQ:** entra como metade do LQ. Sem LQ no laudo, entra como 1/10 do valor orientador.
- **Poços no mesmo ponto:** poços a até 5 m um do outro (ex.: PM-26, PMN-26A, PMN-26B) contam como um ponto, com o maior valor. Para ver um nível só, use o filtro de rede.
- **Fora da rede de poços:** o valor decai suavemente até o piso numa distância igual ao espaçamento local, fechando a pluma em curva.
- **Desenho:** só é pintado o que passa do limiar (padrão = VI). Grade de 2 m; a linha do limiar é emendada e suavizada.

Para instalar e publicar, veja **PASSO-A-PASSO.md**.
