# WebGeo — passo a passo para colocar no ar (teste)

O WebGeo usa o **mesmo Supabase do Perfil de Sondagem**: mesmo login, mesmas empresas e os **mesmos projetos**.
Ele não altera nenhuma tabela do Perfil, só acrescenta tabelas que começam com `wg_`.
O Perfil continua funcionando exatamente como está.

Tempo estimado: 20 a 30 minutos.

---

## O que tem nesta pasta

| Arquivo / pasta | Para que serve |
|---|---|
| `index.html`, `style.css`, `script.js` | O portal (é o que vai para a Vercel) |
| `supabase.min.js` | Biblioteca do Supabase (a mesma do Perfil) |
| `modelo/WebGeo_Modelo_Importacao.xlsx` | Planilha modelo, baixada pelo botão "Baixar planilha modelo" |
| `banco/00_limpar_versao_anterior.sql` | Só se você rodou alguma versão antiga do WebGeo (com `wg_sitio`) |
| `banco/01_webgeo_schema.sql` | Cria as tabelas, as regras de acesso e a função de importação |
| `banco/02_parametros.sql` | Parâmetros e valores orientadores CETESB |
| `banco/03_dados_exemplo_projeto_xyz.sql` | (Opcional) os 88 poços do protótipo Python, via SQL |
| `banco/04_planta_dxf.sql` | Atualização: tabela da planta em DXF, para quem já tinha rodado o 01 antes |
| `dados-exemplo/WebGeo_ProjetoXYZ_C1.xlsx` | Os mesmos dados de exemplo, para importar pela tela |
| `.vercelignore` | Impede que `banco/`, `dados-exemplo/`, `DWG/`, os `.dxf`/`.dwg` e os `.md` sejam publicados no site |

---

## Passo 1 — Banco de dados (Supabase)

1. Entre em supabase.com e abra o **mesmo projeto do Perfil**.
2. Menu da esquerda: **SQL Editor** > **New query**.
3. **Só se você já rodou uma versão antiga do WebGeo**: cole `banco/00_limpar_versao_anterior.sql` e clique em **Run**. Isso apaga só as tabelas `wg_...`; o Perfil não é tocado.
4. Nova query: cole **todo** o `banco/01_webgeo_schema.sql` e clique em **Run**. Deve aparecer "Success. No rows returned".
5. Nova query: cole `banco/02_parametros.sql` e clique em **Run**.
   - **Se você já tinha rodado o 01 antes de existir a planta em DXF:** rode só o `banco/04_planta_dxf.sql` (cria a tabela `wg_planta`). Instalação nova não precisa.
6. Confira em **Table Editor**: devem aparecer as tabelas `wg_poco`, `wg_campanha`, `wg_resultado`, `wg_parametro` (com 12 linhas) etc.

> Se o passo 4 der erro dizendo que `projetos` ou `current_org_id` não existe, falta rodar no
> Supabase os scripts do Perfil `schema.sql` e `adicionar-projetos.sql` (eles já devem estar lá).

---

## Passo 2 — Testar no seu computador (antes de publicar)

1. Abra a pasta `webgeo-portal` no **VS Code**.
2. Instale a extensão **Live Server** (se ainda não tiver).
3. Clique com o botão direito no `index.html` > **Open with Live Server**. O navegador abre em `http://127.0.0.1:5500`.
4. Entre com o **mesmo e-mail e senha do Perfil**.
5. Confira:
   - no menu **Projeto** aparecem os projetos que você já tem no Perfil;
   - as fichas do Perfil **com coordenada** aparecem no mapa do projeto delas;
   - em **Importar planilha**, envie `dados-exemplo/WebGeo_ProjetoXYZ_C1.xlsx`, confira o relatório e clique em **Gravar no banco**.

> Abrir o `index.html` com duplo clique (endereço começando com `file://`) não funciona: precisa do Live Server.

---

## Passo 3 — Subir para o GitHub

Use um **repositório novo**, separado do Perfil (assim um não interfere no outro).

**Pelo site do GitHub (sem comandos):**
1. github.com > **New repository** > nome `webgeo` > marque **Private** > **Create repository**.
2. Na página do repositório: **uploading an existing file**.
3. Arraste **o conteúdo** da pasta `webgeo-portal` (os arquivos e as pastas, não a pasta em si) e clique em **Commit changes**.
   Atenção: o arquivo `.vercelignore` começa com ponto e pode ficar oculto no Windows. No Explorer, ative **Exibir > Itens ocultos** para ele ir junto.

**Ou pelo terminal (se preferir):**
```
cd C:\Users\RafaelSilva\Desktop\Geoportal\WebGeo\webgeo-portal
git init
git add .
git commit -m "WebGeo: primeira versão"
git branch -M main
git remote add origin https://github.com/SEU-USUARIO/webgeo.git
git push -u origin main
```

---

## Passo 4 — Publicar na Vercel

1. vercel.com > **Add New... > Project**.
2. Escolha o repositório `webgeo` > **Import**.
3. Configuração:
   - **Framework Preset:** `Other`
   - **Root Directory:** `./`
   - **Build Command** e **Output Directory:** deixe vazios (não tem build, igual ao Perfil)
4. **Deploy**. Em cerca de 1 minuto sai um endereço como `https://webgeo-xxxx.vercel.app`.
5. A partir daí, cada envio para o GitHub publica sozinho.

**Supabase, ajuste recomendado:** em **Authentication > URL Configuration > Redirect URLs**, adicione o endereço da Vercel
(ex.: `https://webgeo-xxxx.vercel.app/**`). O login por e-mail e senha funciona sem isso; o ajuste vale para os links de
e-mail (confirmação de conta, troca de senha).

---

## Passo 5 — Roteiro de teste no ar

- [ ] Entrar com a conta do Perfil (a mesma empresa aparece no canto do menu)
- [ ] O menu **Projeto** mostra os mesmos projetos do Perfil
- [ ] Uma ficha do Perfil com coordenada aparece no mapa do projeto dela
  - com "Poço de monitoramento nº": círculo com anel tracejado
  - sem poço: losango verde com o nº da sondagem
- [ ] Clicar na ficha mostra o perfil litológico (as mesmas cores do Perfil)
- [ ] Importar a planilha de exemplo: 88 poços, campanha C1, pluma de benzeno no mapa
- [ ] **+ Novo projeto…** no menu: o projeto aparece também no Perfil
- [ ] Entrar com uma conta de **outra empresa**: não vê nada da sua

---

## Papéis, equipe, conferência da importação e relatório (09/10/2026)

**Antes de usar:** rode `banco/05_papeis_e_equipe.sql` no SQL Editor do Supabase. Ele não apaga dados. Em toda empresa sem administrador, a conta mais antiga vira administradora.

- **Papéis:** *Administrador* faz tudo; *Técnico* vê os dados, importa planilhas, insere planta e traça seções, mas não exclui projeto nem campanha e não vê a tela Equipe. O papel aparece embaixo do e-mail, no menu.
- **Cadastro com tipo de conta** (tela de entrada do Perfil, aba *Criar conta*): rode antes `banco/06_cadastro_com_papel.sql` (depois do 05). A pessoa escolhe **Administrador** (informa o nome da empresa; a empresa é criada e ela entra direto como administradora) ou **Técnico** (a conta fica aguardando até um administrador adicionar pela tela Equipe). Ninguém vira administrador de uma empresa que já existe pelo cadastro.
- **Equipe** (menu, só administrador): crie o usuário direto ali, com nome, e-mail, senha provisória (o botão *Gerar* sugere uma) e papel (Técnico ou Administrador). Passe o e-mail e a senha para a pessoa; ela troca a senha em **Trocar senha**, embaixo do e-mail no menu. Se a pessoa já tem conta (criou no Perfil), deixe a senha em branco que ela só é adicionada. Também dá para mudar o papel ou remover alguém. A empresa sempre fica com pelo menos um administrador.
- **Confirmação de e-mail do Supabase:** se em *Authentication > Sign In / Providers > Email* a opção "Confirm email" estiver ligada, a pessoa criada recebe um e-mail e precisa clicar no link antes do primeiro acesso. Desligando, ela entra direto com a senha provisória. O Supabase também limita quantos e-mails de confirmação saem por hora no plano gratuito.
- **Segurança:** o script também fecha uma brecha antiga: antes, uma pessoa logada conseguia mudar a própria empresa e o próprio papel direto no banco.
- **Excluir campanha** (administrador): botão ao lado de "Campanha", no menu. Apaga os resultados e as medições daquela campanha; poços e outras campanhas ficam.
- **Conferência da importação:** depois de ler a planilha, o sistema compara com as campanhas anteriores do projeto e lista o que merece olhar: valor 100 vezes o valor orientador (possível erro de unidade), valor 10 vezes maior que o histórico do poço, primeira vez acima do VI, queda brusca para abaixo do LQ, N.A. que mudou mais de 2 m, N.A. negativo ou mais fundo que o poço, cota que mudou mais de 0,5 m e poço que mudou mais de 10 m de lugar. Com alertas, só grava depois de marcar "Conferi esses pontos".
- **Relatório (PDF):** botão no topo. Escolha as partes (resumo e mapa, seção, pluma e evolução, tabela de resultados com o VI de cada parâmetro, notas) e clique em *Gerar PDF*. Na janela de impressão, em "Destino", escolha *Salvar como PDF*. Desmarque "Cabeçalhos e rodapés" para não sair o endereço nas margens. O relatório sai sempre em fundo branco.

---

## Fotos 360° (14/10/2026)
1. **Banco:** rode `banco/07_fotos_360.sql` no SQL Editor do Supabase (uma vez). As imagens usam o mesmo bucket das fotos do Perfil (`fotos-sondagem`), então não precisa criar bucket nem regra nova.
2. **Publicar:** envie `index.html`, `script.js` e `style.css` ao GitHub.
3. **Usar:** no menu, em **Fotos 360°**, clique em **Inserir foto 360°** e escolha a foto (JPG da câmera 360°, proporção 2:1).
   - Se a câmera gravou o GPS, a localização já vem preenchida. Senão, clique em **Marcar no mapa** e clique no ponto onde a foto foi tirada.
   - **Ver em 360° antes de salvar** abre o visor com a foto ainda no seu computador.
   - O portal reduz a foto para no máximo 4096 x 2048 e abaixo de 1,8 MB antes de enviar. A original fica com você.
4. **Ver:** clique no marcador **360°** no mapa (ou no nome na lista do menu) e em **Ver foto 360°**. Arraste para olhar em volta; roda do mouse ou pinça para aproximar; ⛶ para tela cheia; Esc fecha.
5. **Mudar local / excluir:** pelo balão do marcador, ou **Local** na lista do menu. Excluir pede um segundo clique para confirmar.

## Várias seções: A–A', B–B', C–C'... (28/10/2026)
Não precisa de SQL.
1. No menu, em **Seção geológica**, o botão mostra a próxima letra livre: **Traçar seção B–B'**, **Traçar seção C–C'**... Trace no mapa como antes (clique no começo, no fim e em Concluir).
2. Acima do desenho aparecem as abas **A–A' | B–B' | C–C' | + Nova seção**: clique numa aba para ver aquele corte.
3. No mapa, todas as seções aparecem: a aberta em **amarelo** e as outras em **branco** tracejado, cada uma com a letra nas pontas. Clicar na letra de uma seção branca abre essa seção.
4. **Apagar seção B–B'** (no menu) apaga só a seção aberta. A próxima seção nova reaproveita a letra livre.
5. **Relatório (PDF):** a opção Seção geológica imprime **todas** as seções, em ordem de letra, uma por página.
6. **Diagrama de cercas 3D:** a opção "Da seção…" usa a seção que estiver aberta.
As seções continuam guardadas no navegador, por projeto; a seção única que já existia vira a A–A'.

## Mapa: satélite, malha UTM e rosa dos ventos (23/10/2026)
Não precisa de SQL. No canto de cima, à esquerda do mapa (embaixo do + e −), há dois botões, um embaixo do outro:
- **Satélite:** liga/desliga a imagem de satélite. Desligado, o fundo fica cinza-claro (bom para imprimir ou ver só os poços e a planta; se a planta DXF estiver branca, troque a cor dela para preto no menu).
- **Malha UTM:** linhas pontilhadas de coordenadas UTM, **brancas** sobre o satélite e **pretas** com o satélite desligado, com os valores de **E** no alto e de **N** na esquerda do mapa. O passo se ajusta ao zoom (10, 20, 50, 100 m…), e embaixo aparece o fuso (ex.: UTM 23S · WGS 84 / SIRGAS 2000).
A **rosa dos ventos** fica sempre no canto de cima, à direita, com o norte em vermelho (N, L, S, O).
A escolha dos botões fica guardada no navegador. A malha e a rosa saem também no **Relatório (PDF)**; os botões não.
**Versão dos arquivos:** o `index.html` chama `style.css?v=…` e `script.js?v=…`. Sempre que mudar o `script.js` ou o `style.css`, troque esse número (ex.: `v=20261023-1520` → `v=20261024-0900`) para o navegador de todo mundo baixar a versão nova, em vez de usar a que guardou.

## Prancha A3: visualizar antes de baixar/imprimir (09/11/2026)
- **O que mudou:** `script.js`, `style.css` e `index.html` (versão `?v=20261109-2200`). Nenhum SQL novo.
- **Como usar:** na janela da **Prancha A3 (PDF)**, preencha os campos e clique em **Visualizar**. A folha aparece inteira na tela, do jeito que vai sair.
- **Na visualização:** **−** / **+** aproximam e afastam, **Caber** volta à folha inteira; **Voltar e ajustar** volta à janela com os campos preenchidos (mude o que quiser e visualize de novo); **Imprimir** abre a impressão do navegador (escolha A3, paisagem, "tamanho real"); **Baixar PDF** salva o arquivo; **Abrir em nova aba** abre o PDF no leitor do navegador.
- **Baixar PDF direto** continua lá, para quem não quer ver antes.

## Desenhar e medir no mapa (09/11/2026)
- **O que mudou:** `script.js`, `style.css` e `index.html` (versão `?v=20261109-2130`). **Nenhum SQL novo:** os desenhos ficam num arquivo do projeto, no mesmo armazenamento das fotos (`<empresa>/campo-<projeto>/desenhos.json`), e todo mundo da empresa vê.
- **Onde:** menu da esquerda, botão **Desenhar e medir** (abaixo de "Editar linha da pluma"). Abre uma barra em cima do mapa.
- **Ferramentas:**
  - **Reta:** clique ponto a ponto; duplo clique (ou **Concluir traço**, ou Enter) termina. Clicar de novo no 1º ponto fecha a figura.
  - **Mão livre:** segure e arraste (mouse ou dedo). Nesse modo o mapa não se move.
  - **Medir:** clique ponto a ponto; mostra a medida de cada trecho e o total (ao vivo na barra enquanto desenha). Clicando no 1º ponto, fecha e mostra o **perímetro** e a **área** (m² ou hectares). As medidas são em UTM, em metros.
  - **Borracha:** clique num desenho para apagá-lo.
- **Cor e espessura:** 7 cores prontas e o círculo colorido para qualquer outra; espessura fina, média ou grossa. Vale para o próximo traço.
- **Desfazer** (ou Ctrl+Z) volta o último ponto ou o último desenho; **Apagar tudo** limpa os desenhos do projeto; **Esc** cancela o traço em andamento.
- **Mostrar desenhos:** depois de fechar a barra, a caixa "Mostrar desenhos" (menu da esquerda) liga e desliga.
- **Saem também** na **Prancha A3 (PDF)** e no **Baixar DXF do mapa** (camadas DESENHOS e MEDIDAS, com as cotas e a área).

## Modo campo (GPS) com imagem de drone (09/11/2026)
- **O que mudou:** `script.js`, `style.css` e `index.html` (versão `?v=20261109-2000`). **Nenhum SQL novo:** usa as tabelas que já existem (`wg_poco`, `wg_medicao_campo`) e o mesmo armazenamento das fotos (`fotos-sondagem`).
- **1. Subir a imagem do drone (no escritório):** menu da esquerda, **Ortofoto (drone)** → escolha o **GeoTIFF** (`.tif`), ou a imagem (`.jpg`/`.png`) junto com o arquivo de georreferência dela (`.jgw`/`.pgw`/`.tfw`), selecionando os dois de uma vez. A imagem precisa estar em **UTM** (SIRGAS 2000/WGS 84; o fuso vem do GeoTIFF ou do projeto). Fica gravada no projeto, para todo mundo da empresa, e aparece por baixo dos poços. O botão **Ortofoto** no mapa liga/desliga; "remover" tira do projeto.
- **2. No campo (celular):** abra o WebGeo, escolha o projeto e toque em **Modo campo (GPS)** no alto. O celular pede para usar a localização (e, no iPhone, a bússola): permita. A tela vira só mapa + painel: o ponto azul é você (com o círculo da precisão), e a lista mostra os poços do mais perto para o mais longe, com ✓ nos que já têm N.A. nesta campanha.
- **3. Ir até o poço:** toque no nome do poço. Aparece a linha amarela até ele, a distância e a seta do rumo (com bússola a seta gira com o celular; sem bússola ela é em relação ao norte). Ao chegar perto aparece **"Você chegou!"**.
- **4. Ficha de campo:** toque em **Ficha** (no painel, na lista ou no popup do poço):
  - **Poço:** *Localizado (confirmado)* ou *Não localizado*.
  - **Situação:** ativo, **seco**, **danificado/destruído** ou desativado.
  - **N.A. (m)**, fase livre e espessura (gravados na campanha escolhida).
  - **Descrição:** vai para a observação do poço, com a data (ex.: `[Campo 09/11/2026] Tampa quebrada`).
  - **Corrigir a coordenada com o GPS:** opcional; mostra a precisão e a distância até onde o poço está hoje. Cada poço (PM, PMN…) fica com a sua coordenada.
  - **Fotos:** abre a câmera; pode tirar várias. São reduzidas antes de enviar e ficam na ficha do poço (aparecem da próxima vez).
- **Dicas:** use a céu aberto (precisão de 3–10 m é normal no celular); a tela precisa ficar aberta para o GPS seguir; o site precisa estar em **https** para o celular liberar o GPS. **Sair** volta para a tela normal.

## Mapa de calor (09/11/2026)
- **O que mudou:** só `script.js` e `index.html` (versão `?v=20261109-1650`). Nenhum SQL novo.
- **Onde:** menu da esquerda, logo abaixo de "Mostrar pluma": **Mapa de calor (todos os valores)**. A escolha fica lembrada neste navegador.
- **Como fica:** em vez de pintar só o que passa do limite, pinta a concentração interpolada (mesmo IDW da pluma) em toda a área dos poços, do valor mais baixo (azul) ao mais alto (vermelho), passando por verde e amarelo, em escala log. Os valores baixos ficam mais transparentes. A linha do limite (contínua/tracejada) continua por cima.
- **Legenda e prancha:** a legenda do mapa vira "Mapa de calor" com a barra azul-vermelho; a Prancha A3 sai com o mapa de calor, a barra de cores na legenda e o título "MAPA DE CALOR DE <parâmetro>".
- **Observação:** com a linha da pluma ajustada à mão, vale a pluma ajustada (o mapa de calor fica desligado nessa campanha). O DXF continua com as faixas da pluma.

## Potenciométrico em DXF (09/11/2026)
- **O que mudou:** só `script.js` e `index.html` (versão `?v=20261109-1600`). Nenhum SQL novo.
- **Como usar:** ligue o **Mapa potenciométrico** e clique em **Baixar DXF do mapa** (menu da esquerda, bloco Planta). Com a pluma desligada, o arquivo sai como `potenciometrico_<projeto>_<campanha>.dxf`; com a pluma ligada, o DXF leva as duas coisas.
- **Camadas novas:** POT_EQUIPOTENCIAIS (curvas, azul), POT_EQUIPOTENCIAIS_VALOR (valor de cada curva), POT_FLUXO (setas) e POT_CARGA_HIDRAULICA (carga de cada poço, em cima do poço). A legenda do DXF ganha linha equipotencial (com o intervalo), carga hidráulica e sentido do fluxo. Tudo em UTM (metros, SIRGAS 2000).

## Potenciométrico: PDF, rótulos de carga e setas menores (06/11/2026)
- **O que mudou:** `script.js`, `style.css` e `index.html` (versão `?v=20261106-1400`). Nenhum SQL novo.
- **PDF:** com o mapa potenciométrico ligado, a **Prancha A3 (PDF)** ganha a opção de desenho **Mapa potenciométrico**. Sai com as curvas equipotenciais (com o valor), as setas de fluxo, a carga hidráulica de cada poço em azul e a legenda (linha equipotencial, exemplo do valor da carga hidráulica em azul, sentido do fluxo). Título padrão: "MAPA POTENCIOMÉTRICO - CAMPANHA C1". No desenho "Mapa", se o potenciométrico estiver ligado, ele também aparece.
- **Rótulos da carga:** a carga hidráulica saiu de dentro do nome e virou uma etiqueta azul própria, em cima de cada poço. No modo **Ajustar rótulos (nomes e valores)** ela também é arrastável (etiqueta azul); a posição vale para o mapa e para a prancha.
- **Setas:** mais curtas e finas no mapa (um terço do espaçamento, no máximo 9 m); no PDF têm tamanho fixo de 5 mm.

## Editar a linha da pluma à mão (06/11/2026)
- **O que mudou:** `script.js`, `style.css` e `index.html` (versão `?v=20261106-1130`). Nenhum SQL novo.
- **Onde:** menu da esquerda, botão **Editar linha da pluma** (abaixo de "Ajustar rótulos").
- **Como usar:** a linha calculada vira uma linha com pontos brancos (quadradinhos). Arraste os pontos; para criar um ponto, clique ou arraste a bolinha amarela do meio de um trecho; para apagar, duplo clique (ou botão direito) no ponto. Clique num trecho para trocar entre **contínuo** e **tracejado**. **Salvar linha** grava; **Cancelar** desfaz; **Voltar à calculada** apaga o ajuste.
- **Desenhar uma pluma nova (06/11/2026, versão `?v=20261106-1545`):** na barra de edição, **Desenhar pluma nova** acrescenta mais uma pluma; **Apagar tudo e redesenhar** apaga as linhas e já começa do zero. Clique no mapa ponto a ponto em volta da área; **duplo clique** (ou **Concluir desenho**) fecha a pluma; **Esc** desiste. Depois de fechada, ela se edita como as outras (arrastar, criar e apagar pontos, contínuo/tracejado). O campo **Traço** escolhe **spline (curva)** ou **linha reta** para todas as linhas. **Salvar linha** grava.
- **O que acompanha a linha:** a cor da pluma é recortada pela linha nova (dentro dela a pluma aparece mesmo onde o cálculo dava abaixo do limite), a área e o centróide passam a ser os da linha ajustada, e a prancha A3 e o DXF do mapa usam a linha ajustada (a nota da prancha diz que foi ajustada manualmente). A linha passa suave pelos pontos.
- **Guardada por:** projeto + parâmetro + campanha, **neste navegador** (outro computador vê a calculada). No menu aparece "Linha ajustada à mão nesta campanha · voltar à calculada".

## Linha do limite: contínua e tracejada (06/11/2026)
- **O que mudou:** `script.js`, `style.css` e `index.html` (versão `?v=20261106-1000`). Nenhum SQL novo. O contorno da pluma (IDW) não mudou.
- **Contínua:** o trecho da linha do limite que passa entre um poço com valor acima e um poço com valor abaixo do limite (os dois medidos). Para isso, os poços são ligados em triângulos (Delaunay); o trecho é contínuo quando cai num triângulo que tem poço acima e poço abaixo do limite.
- **Tracejada:** o resto da linha, onde ela é inferida (borda da rede, ou sem poço abaixo do limite por perto).
- **Onde vale:** no mapa (legenda "Limite da pluma delimitada" e "Limite de pluma inferida"), na prancha A3 (legenda: "Limite da pluma delimitada (valor)" e "Limite de pluma inferida"; a explicação saiu das notas) e no DXF do mapa (camadas PLUMA_LIMITE, contínua, e PLUMA_LIMITE_INFERIDO, tracejada).
- Poços sem resultado na campanha não contam.

## Prancha A3 em PDF (05/11/2026)
- **O que mudou:** `script.js`, `style.css` e `index.html` (versão `?v=20261105-1830`). Nenhum SQL novo. O gerador de PDF (jsPDF 2.5.1) é baixado do cdnjs só na hora de gerar.
- **Onde:** botão **Prancha A3 (PDF)** no alto da tela, ao lado de "Relatório (PDF)".
- **Como usar:** deixe o mapa na vista desejada (zoom e posição), clique no botão, escolha o desenho (o mapa ou uma das seções), preencha o carimbo (relatório, título, projeto, data, endereço, contratante, responsável técnico, folha) e o logo, e clique em **Gerar PDF**.
- **A folha:** A3 deitada. À esquerda o desenho: imagem de satélite, pluma, planta, malha UTM com coordenadas, poços com nome e barra de escala com a escala numérica. À direita: rosa dos ventos, legenda (só o que aparece no desenho) e o carimbo com o logo (o bloco de notas foi retirado em 06/11/2026).
- **Centralizar e escala (06/11/2026, versão `?v=20261106-0845`):** na janela, **Centralizar em** (todos os poços, a pluma, a planta ou a vista atual do mapa) e **Escala** (automática ou fixa: 1:250, 1:500, 1:750, 1:1.000, 1:1.500, 1:2.000, 1:2.500, 1:5.000, 1:10.000 ou outra digitada). Na automática, o desenho fica no centro da folha com folga de 10% em volta e a escala é arredondada para um valor de prancha (ex.: 1:500, 1:1.250). Com escala fixa, o centro continua sendo o escolhido.
- **Valores e rótulos (06/11/2026, versão `?v=20261106-0915`):** o valor do resultado sai embaixo de cada poço (vermelho e em negrito quando passa do VI; "<" quando abaixo do LQ; a unidade vai nas notas). A caixa **Valores dos resultados embaixo de cada poço** liga/desliga. O botão **Ajustar posição dos rótulos** fecha a janela e deixa o nome (etiqueta escura) e o valor (etiqueta amarela) de cada poço arrastáveis no mapa; **Concluir** salva e volta para a prancha, **Voltar ao padrão** desfaz tudo. Rótulo arrastado para longe do poço ganha uma linha de chamada no PDF. As posições ficam guardadas neste navegador, por projeto.
- **No próprio mapa (06/11/2026, versão `?v=20261106-0920`):** no menu da esquerda, abaixo de "Mostrar pluma", a caixa **Mostrar valores nos poços** (valor embaixo de cada poço; aparecem nos mesmos poços que mostram o nome: os acima do VI e, bem de perto, todos) e o botão **Ajustar rótulos (nomes e valores)**, que abre o mesmo modo de arrastar. A posição ajustada vale para o mapa e para a prancha.
- **Logo e campos:** o logo fica guardado neste navegador para a empresa; os campos do carimbo, por projeto. Em outro computador, é preciso escolher o logo de novo.
- **Satélite:** se o servidor da imagem não deixar copiar, a prancha sai sem o satélite e a mensagem avisa.

## Mapa em DXF: planta, poços e pluma (05/11/2026)
- **O que mudou:** `script.js`, `style.css` e `index.html` (versão `?v=20261105-1530`). Nenhum SQL novo.
- **Onde:** menu da esquerda, bloco Planta (DXF), botão **Baixar DXF do mapa**.
- **O que sai:** um DXF R12 em **UTM (metros, SIRGAS 2000)**, no mesmo fuso da planta (sem planta, no fuso dos poços). Abre no AutoCAD já na posição real.
- **Cor da planta:** a mesma escolhida em "Cor das linhas" no menu: Preto sai na cor 7 do CAD (preta na impressão, branca no fundo escuro do AutoCAD), Branco e Amarelo na cor mais próxima, Cores do CAD com a cor de cada camada.
- **Camadas:** PLANTA_<camada> (só as camadas visíveis); PLUMA_FAIXA_1, _2... (pluma preenchida por classe: limite a 10x, 10x a 100x...), PLUMA_LIMITE (vermelha), PLUMA_ISOLINHAS (com o valor escrito); POCOS_<situação do resultado> (círculos), POCOS_NOMES, POCOS_RESULTADOS; LEGENDA (título, classes da pluma, símbolos e norte, à direita do desenho).
- **Usa o que está na tela:** parâmetro, campanha, redes visíveis, limite e se a pluma está ligada.
- **Limitação:** a borda do preenchimento sai em degraus de uma célula da grade (2 m); a linha do limite sai lisa.

## Situação do poço: seco, destruído... (05/11/2026)
- **O que mudou:** `script.js`, `style.css` e `index.html` (versão `?v=20261105-1140`). Nenhum SQL novo: usa as situações que o banco já aceita.
- **Como usar:** clique no poço → **Editar poço** → escolha a **Situação** (Ativo, Seco, Danificado/destruído, Desativado) → **Salvar**. Dá para mudar só a situação, só a posição ou as duas juntas.
- **Destruído:** é gravado como DANIFICADO (aparece como "danificado/destruído"). Na planilha, a coluna Situação aceita DESTRUÍDO e grava do mesmo jeito.
- **No mapa:** poço que não está ativo fica com a borda preta tracejada, a cor apagada e a situação escrita ao lado do nome. O popup também mostra a situação.
- Os resultados antigos do poço continuam valendo (histórico, pluma da campanha em que foi amostrado).

## Editar a posição de um poço no mapa (05/11/2026)
- **O que mudou:** `script.js`, `style.css` e `index.html` (versão `?v=20261105-1100`). Nenhum SQL novo (a tabela wg_poco já deixa a equipe da empresa alterar).
- **Como usar:** clique no poço no mapa → **Editar posição**. Aparece um pino amarelo e uma barra em cima do mapa com Latitude, Longitude, E e N.
- **Três jeitos:** arrastar o pino, clicar no mapa onde o poço está, ou digitar a coordenada (em graus ou em UTM E/N) e apertar Enter. A linha tracejada amarela liga à posição antiga e a barra mostra quantos metros moveu.
- **Salvar** grava no banco na hora; **Cancelar** (ou Esc) desfaz.
- **Atenção:** se depois importar de novo uma planilha com a coordenada antiga desse poço, a da planilha volta a valer. Corrija também na planilha.

## Botão "Converter UTM em lat/long" (03/11/2026)
- **O que mudou:** `script.js`, `style.css` e `index.html` (versão `?v=20261103-1030`). Nenhum SQL novo.
- **Onde:** tela Importar planilha, abaixo de "Baixar planilha modelo".
- **Como usar:** clique, escolha a planilha. Ela volta baixada como `<nome>_latlong.xlsx`, com a Latitude e a Longitude da aba Pocos em graus decimais. O E/N original e o fuso usado ficam em três colunas novas no fim da aba. Nada é gravado no banco.
- **Regras:** as mesmas da importação (fuso, E de 5 dígitos completado). A mensagem embaixo do botão diz quantos poços foram convertidos e o que foi somado.
- **Atenção:** a planilha baixada mantém os dados, mas perde a formatação (cores, comentários das células).

## Planilha: E com 5 dígitos (03/11/2026)
- **O que mudou:** só `script.js` e `index.html` (versão `?v=20261103-0945`). Nenhum SQL novo.
- **O caso:** coordenadas tiradas de planta de CAD que apagou a centena de milhar do E (ex.: 37859,50 em vez de 337859,50), com o N completo (7384801,75).
- **O que o sistema faz:** reconhece o E de 5 dígitos e soma a centena de milhar (100.000, 200.000, 300.000...) que põe os poços mais perto da referência: poços em graus da planilha, poços da planilha com E completo, poços já cadastrados no projeto ou, num projeto novo, os poços do projeto aberto no menu. O relatório da importação avisa quanto foi somado.
- **Para forçar:** campo **Somar ao E** na aba Projeto (ex.: 300000). Sem referência nenhuma e sem esse campo, a importação pede para preencher.

## Planilha: cada poço com a sua coordenada (30/10/2026)
- **O que mudou:** `script.js`, `style.css` e `index.html` (versão `?v=20261030-1505`). Nenhum SQL novo.
- Saiu o quadro que copiava a coordenada do PM para o PMN de mesmo número: cada PM e PMN fica com a coordenada da planilha.
- Quem importou com a versão anterior e deixou os PMN marcados: importe a planilha de novo para voltar às coordenadas próprias.

## Planilha: coordenadas em graus ou em UTM (30/10/2026)
- **O que mudou:** só `script.js` e `index.html` (versão `?v=20261030-1010`). Nenhum SQL novo.
- **Graus decimais:** como antes, nas colunas Latitude e Longitude (ex.: -23.471383 / -46.575588, com ponto ou vírgula).
- **UTM em metros:** nas colunas E e N (também aceita Este/Norte, UTM E/UTM N, X/Y) ou até nas próprias colunas Latitude/Longitude. O sistema reconhece pelo tamanho do número (N tem 7 dígitos, E tem 6), mesmo com as colunas trocadas, e converte para graus antes de gravar.
- **Fuso:** coluna **Fuso** na aba Pocos (ex.: 23S) ou campo **Fuso** na aba Projeto. Sem isso, usa o fuso dos poços em graus da planilha, depois o dos poços que o projeto já tem, depois o da UF; o relatório da importação avisa qual fuso foi usado.
- **Datum:** SIRGAS 2000 / WGS 84 (a diferença entre os dois é de centímetros).

## Seção em DXF (28/10/2026)
- **O que mudou:** só `script.js` e `index.html` (versão `?v=20261028-1320`). Nenhum SQL novo.
- **Como usar:** na seção aberta, botão **Baixar DXF** (ao lado de "Baixar imagem (SVG)"). Sai um arquivo `secao_A-A_<projeto>.dxf`.
- **Formato:** DXF R12, que abre no AutoCAD, BricsCAD, QGIS e LibreCAD. Unidade em metros: X = distância ao longo da seção; Y = cota multiplicada pelo exagero vertical da tela (escrito no cabeçalho do desenho), para ficar igual ao que aparece no WebGeo.
- **Camadas:** LITOLOGIA_SONDAGENS, LITOLOGIA_ENTRE_SONDAGENS (preenchidas), SONDAGENS, TUBO_LISO, FILTRO, NIVEL_AGUA, TERRENO, POCOS_NOMES, POCOS_RESULTADO, PLUMA_LIMITE (vermelha), PLUMA_ISOLINHAS (10x, 100x... o limite, com o valor), PLUMA_AMOSTRAS, GRADE, MOLDURA e TEXTOS.
- **Limitação:** o degradê de cor da pluma não vai para o DXF (o formato não tem imagem); vão as linhas do limite e as isolinhas. As cores são as mais próximas da paleta do CAD.

## Pluma vertical na seção (28/10/2026)
- **O que mudou:** só `script.js`, `style.css` e `index.html` (versão `?v=20261028-1205`). Nenhum SQL novo.
- **Como usar:** abra uma seção. A pluma do parâmetro e da campanha escolhidos aparece dentro do corte, abaixo do nível d'água, com a linha tracejada do limite (VI) e um ponto branco no meio de cada filtro amostrado.
- **Linha do limite:** vermelha e contínua, suavizada. Em cada poço a concentração parte do valor medido no N.A. e cai até o limite no fundo do filtro, então a linha sai do fundo do poço contaminado e sobe em curva até o nível d'água, afinando com a distância.
- **Ligar/desligar:** no menu da seção, caixa **Pluma vertical na seção**.
- **Relatório e SVG:** a pluma sai junto no PDF e no "Baixar imagem (SVG)".
- **Para conferir no dev:** depois do push, abra uma seção que passe por um poço PM e um PMN no mesmo ponto; a pluma deve mudar com a profundidade.

## Equipe: editar nome, e-mail e nome da empresa (23/10/2026)
1. **Banco:** rode o `banco/09_editar_pessoa.sql` uma vez no SQL Editor do Supabase. Cria só duas funções, `wg_editar_pessoa` e `wg_editar_empresa`, nenhuma tabela: o nome fica em `profiles`, o e-mail no login do Supabase (`auth.users`) e o nome da empresa em `organizations`.
2. **Usar (só administrador):** tela **Equipe**.
   - **Nome da empresa:** campo no topo + **Salvar** (ou Enter). Muda na hora no menu e vale também para o Perfil de Sondagem.
   - **Pessoa:** **Editar** na linha → mude o nome e/ou o e-mail → **Salvar** (ou Enter; Esc cancela). O e-mail novo passa a valer para entrar no WebGeo e no Perfil; a senha continua a mesma.
3. Avisos: e-mail inválido, e-mail que já é de outra conta, nome da empresa vazio. Técnico não consegue editar (o banco recusa).

## Anotações do projeto (23/10/2026)
1. **Banco:** rode o `banco/08_anotacoes.sql` uma vez no SQL Editor do Supabase. Cria só a tabela `wg_anotacao` com a regra de acesso por empresa; pode rodar de novo.
2. **Usar:** menu **Anotações**. Um texto por projeto, que a equipe toda da empresa vê e edita. Salva sozinho 1,5 s depois que você para de digitar (ou Ctrl+S, ou ao trocar de projeto). Embaixo aparece quando e por quem foi a última alteração.
3. Se duas pessoas editarem ao mesmo tempo, aparece um aviso com **Ver a versão salva** ou **Salvar a minha por cima**: ninguém perde texto sem saber.
4. No **Relatório (PDF)**, a opção **Anotações do projeto** coloca o texto no relatório.

## Diagrama de cercas 3D (15/10/2026)
Não precisa rodar SQL; basta publicar `index.html`, `script.js` e `style.css`.
1. No menu, em **Seção geológica**, clique em **Diagrama de cercas 3D** (ou em **Ver em 3D**, em cima da seção A–A').
2. Em **Sondagens**, escolha:
   - **Da seção A–A'**: as mesmas sondagens da seção traçada, na posição real, ligadas na ordem de A para A'.
   - **Todas as sondagens**: todas as que têm litologia e coordenada (nas redes visíveis), ligadas pelo caminho mais curto entre elas, sem painel cruzando outro.
3. **Arraste** para girar; **botão direito** ou **Shift + arrastar** para mover; **roda** ou pinça para aproximar. Os botões do lado direito aproximam, afastam, mostram de cima, voltam à vista inicial e põem em tela cheia.
4. Passe o mouse numa sondagem para ver a camada naquela profundidade, a cota e o N.A.
5. **Exagero vertical**, **Painéis** e **Nível d'água** ligam e desligam na hora. **Baixar imagem (PNG)** salva a vista atual com os nomes.
Os painéis usam o mesmo preenchimento da seção 2D (liga camadas parecidas e afina as que não têm par): é uma interpretação automática e precisa de conferência.

## Seção geológica

No menu, em **Seção geológica**, clique em **Traçar seção A–A'** e marque no mapa o início e o fim da linha (o clique gruda no poço mais próximo). Pode marcar pontos no meio para a linha fazer curva. Clique em **Concluir** (ou dê dois cliques rápidos no último ponto). Não precisa de nada novo no banco.

- **Quem entra:** poços e sondagens a até a distância da **Faixa** (5 a 40 m) para cada lado da linha. Eles ganham um anel amarelo no mapa.
- **Litologia:** vem da ficha do Perfil de Sondagem com o mesmo "Poço nº" (ou da aba *Litologia* da planilha). Poço sem litologia aparece só com o nome e o nível d'água.
- **Altura:** usa a **Cota topo (m)** do poço. Sem cota, usa a do poço mais próximo e marca o nome com `*`.
- **Nível d'água:** o N.A. da campanha selecionada (ou o da ficha), com uma linha por rede (PM, PMN).
- **Preencher entre as sondagens:** fecha todo o espaço entre sondagens vizinhas. Camadas parecidas (mesmo material, em alturas próximas) são ligadas com curvas suaves; a que não tem par afina até acabar no caminho; a que fica abaixo do fim de uma sondagem mais rasa passa por baixo dela. É uma interpretação automática e precisa de conferência; dá para desligar.
- **Exagero vertical:** automático, ou de 1x a 20x.
- **Baixar imagem (SVG):** salva o desenho em fundo branco.
- A linha fica guardada no navegador de quem traçou, por projeto.

---

## Planta em DXF

No menu, em **Planta (DXF)**, clique em **Inserir planta (DXF)** e escolha o arquivo. A planta fica salva no projeto, para todos da empresa.

- **Formato:** DXF em texto (no AutoCAD: *Salvar como > DXF*). DWG não é aceito.
- **Coordenadas:** o ideal é o desenho em **UTM SIRGAS 2000** (X = Leste, Y = Norte). A zona é sugerida pelos poços do projeto (23 em São Paulo).
- **O que entra:** linhas, polilinhas, círculos, arcos, elipses, splines e blocos. Textos, hachuras e cotas não são importados.
- **Camadas e cor:** dá para ligar/desligar cada camada e escolher a cor das linhas (preferência de cada usuário).
- **Fora de posição?** Use **Ajustar posição**: clique num canto da planta e depois onde ele fica no satélite. Com 1 ponto a planta é deslocada; com 2 pontos, deslocada e girada. A escala do desenho é mantida.
- **Arquivos deste projeto:** `Geoportal\Base.dxf` está em UTM correto (os poços do desenho batem com os reais com 2 a 4 m de diferença). `DWG\base.dxf` está com os eixos trocados e precisa do ajuste por pontos.

---

## Mapa potenciométrico

No menu, marque **Mapa potenciométrico**. Ele usa a campanha selecionada e não precisa de nada novo no banco.

- **De onde vem:** carga = **Cota topo (m)** da aba *Pocos* menos **N.A. (m)** da aba *Campo*. Poço sem cota ou sem coordenada fica de fora.
- **Poços:** escolha uma rede por vez (PM = rasa; PMN = multinível). Misturar níveis distorce o mapa.
- **Curvas a cada:** intervalo entre equipotenciais; no automático o sistema escolhe.
- **Superfície:** "Passa pelos poços" respeita cada medida. As suavizações amortecem poços destoantes e o menu mostra quanto a superfície se afastou do medido.
- **Setas de fluxo:** apontam do maior para o menor potencial.
- **Com o potenciométrico ligado**, a carga de cada poço aparece ao lado do nome no mapa.

---

## Como os dois sistemas se ligam

| No Perfil | No WebGeo |
|---|---|
| Projeto | Projeto (é a mesma tabela `projetos`) |
| Ficha com coordenada (UTM convertida) | Ponto no mapa do projeto |
| "Poço de monitoramento nº" (ex.: PM-30) | Poço com o mesmo código: os resultados de laboratório da planilha se ligam a ele |
| Litologia + tubo liso/filtro + N.A. | Perfil desenhado no popup do poço |

- **Coordenada:** vale a da planilha (aba Pocos). Se o poço não tem coordenada na planilha, vale a da ficha do Perfil.
- **Ficha sem latitude:** fichas salvas antes de preencher as coordenadas UTM podem estar sem latitude/longitude. Abra a ficha no Perfil, confira E/N/zona e salve de novo.

---

## Problemas comuns

| Sintoma | Causa provável / solução |
|---|---|
| "Erro ao falar com o banco: relation wg_... does not exist" | Faltou o Passo 1 (rodar `01_webgeo_schema.sql`) |
| Tela "Conta sem empresa" | O usuário não tem empresa no `profiles` — vincule como no Perfil (Table Editor > profiles > organization_id) |
| Menu Projeto vazio | A empresa ainda não tem projeto: crie em "+ Novo projeto…" (ou no Perfil) |
| Ficha não aparece no mapa | Ficha sem projeto, ou sem coordenada salva (ver "Ficha sem latitude" acima) |
| Fundo do mapa cinza | Sem internet para as imagens de satélite (Esri) — os poços aparecem mesmo assim |
| "O banco recusou a importação" | A mensagem diz o motivo; nada foi gravado, pode corrigir e enviar de novo |
| "Falta criar a tabela da planta no banco" | Rode `banco/04_planta_dxf.sql` no SQL Editor |
| Planta aparece longe ou espelhada | O DXF não está em UTM (ou está com X e Y trocados): use "Ajustar posição" |
