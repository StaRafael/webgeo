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

## Seção geológica

No menu, em **Seção geológica**, clique em **Traçar seção A–A'** e marque no mapa o início e o fim da linha (o clique gruda no poço mais próximo). Pode marcar pontos no meio para a linha fazer curva. Clique em **Concluir** (ou dê dois cliques rápidos no último ponto). Não precisa de nada novo no banco.

- **Quem entra:** poços e sondagens a até a distância da **Faixa** (5 a 40 m) para cada lado da linha. Eles ganham um anel amarelo no mapa.
- **Litologia:** vem da ficha do Perfil de Sondagem com o mesmo "Poço nº" (ou da aba *Litologia* da planilha). Poço sem litologia aparece só com o nome e o nível d'água.
- **Altura:** usa a **Cota topo (m)** do poço. Sem cota, usa a do poço mais próximo e marca o nome com `*`.
- **Nível d'água:** o N.A. da campanha selecionada (ou o da ficha), com uma linha por rede (PM, PMN).
- **Ligar camadas iguais:** liga as camadas de mesmo nome entre sondagens vizinhas. É automático e precisa de conferência; dá para desligar.
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
