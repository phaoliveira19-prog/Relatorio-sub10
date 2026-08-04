# Sistema de Gestão — Sub-10

Dashboard de planejamento e monitoramento da categoria, construído a partir do
"Sistema de gestão.xlsx" (Banco de atletas, Banco de treinos, Avaliação
individual, Pós-jogo, Pós-jogo individual, Caderno de conteúdos, Tipos de
tarefa de treino, Avaliação do treino, Princípios pedagógicos).

## Páginas do site

- `index.html` — capa (logo + botões "Acesso ao sistema" / "Como funciona o
  sistema?"). Estática, sem dados, é a porta de entrada.
- `dashboard.html` — o dashboard em si (era o `index.html` antes da capa
  existir — qualquer link/bookmark antigo pra `index.html` como dashboard
  precisa ser atualizado).
- `metodologia.html` — explicação da metodologia (conteúdos, tipos de
  tarefa, TEA/TAA/AAT), com conteúdo tirado das abas "Caderno de
  conteúdos", "Tipos de tarefa de treino" e "Avaliação do treino" da
  planilha.
- `entrada.html` — preenchimento (ver seção própria abaixo).

## Como funciona hoje

```
Google Sheets --(fetch CSV ao vivo)--> dashboard.html   [fonte principal]
sistema.xlsx  --(build_data.py)-->     data.json         [fallback]
```

`dashboard.html` busca as 5 abas da planilha do Google Sheets direto no
navegador (sem backend) e monta os dados na hora. Se isso falhar por
qualquer motivo (sem internet, aba renomeada, permissão de
compartilhamento), ele cai automaticamente para o `./data.json` local,
gerado a partir de um xlsx antigo — então o site nunca fica fora do ar, mas
os dados podem ficar desatualizados até o problema com o Sheets ser
resolvido.

Pra regenerar o `data.json` de fallback a partir de um xlsx mais novo:

```bash
python3 build_data.py /caminho/para/Sistema_de_gestao.xlsx data.json
```

## Preenchimento da planilha (`entrada.html`)

Página separada (mesma pasta, `dashboard/entrada.html`) com formulários pra
cadastrar atletas e lançar treinos/jogos sem precisar editar a planilha
diretamente — pensada pro preenchimento do dia a dia em campo, principalmente
pelo celular. Escreve na planilha através de um script do Google Apps Script
(`apps-script.gs`, colado dentro da própria planilha — veja o cabeçalho desse
arquivo pras instruções de instalação).

Configurações em `entrada.html` antes de funcionar:
- `SCRIPT_URL`: a URL do Apps Script implantado (veja `apps-script.gs`).
- `LOGIN`: usuário de acesso à página (padrão atual: `pedrohenrique`).
- `PASSWORD_HASH`: hash SHA-256 da senha de acesso à página (senha padrão
  atual: `sub10cruzeiro` — troque assim que possível, veja o comentário no
  arquivo pra gerar um novo hash).

Esse login é uma barreira simples contra acesso casual, não segurança de
verdade (fica visível em quem souber ler o código-fonte da página). Proteção
real fica pendente da migração de hospedagem pra Cloudflare Pages + Access
(ver pendências abaixo).

## Formato da aba "Avaliação individual do treino"

Essa aba é **formato largo**: uma linha por atleta, uma coluna nova a cada
treino (cabeçalho da coluna = data do treino em texto, `DD/MM/AAAA`).
Ausência = célula em branco na grade. `build_data.py`, `dashboard.html` e
`apps-script.gs` leem esse formato — é o jeito mais fácil de preencher no
dia a dia.

## Fotos dos atletas

As fotos hoje apontam para links externos (i.ibb.co) copiados da planilha.
Isso não é ideal a longo prazo (é um ponto único de falha) e **eu não consegui
baixá-las daqui** — a rede deste ambiente bloqueia esse host. Para resolver:

```bash
pip install requests pillow
python3 fetch_photos.py /caminho/para/Sistema_de_gestao.xlsx
python3 build_data.py /caminho/para/Sistema_de_gestao.xlsx data.json
```

Isso baixa, redimensiona e salva cada foto em `dashboard/photos/`, e
`build_data.py` passa a preferir automaticamente o arquivo local em vez do
link externo. Dois atletas hoje não têm link de foto válido na planilha
(Caio Gabriel, Heitor Nery) e um tem um link do Google Drive que não é uma
imagem direta (Henrique Lemes) — esses precisam de foto manual em
`dashboard/photos/<apelido-sem-acento>.png`.

## Pendências conhecidas (decididas na conversa, ainda não implementadas)

- **Autenticação de verdade** — hoje o site (dashboard e entrada.html) é
  público/protegido só por uma senha client-side fraca. Migração planejada
  pra Cloudflare Pages + Cloudflare Access (login por e-mail, gratuito).
- **Versão mobile do dashboard** — hoje o `dashboard.html` quebra em telas
  de celular (foi desenhado pra desktop). `index.html`, `entrada.html` e
  `metodologia.html` já nasceram responsivos.
- **Multi-clube / multi-categoria** — o schema já carrega `categoria` em
  atletas/treinos/jogos, mas o dashboard ainda assume uma única base
  (`data.json`). Quando for expandir para outros clubes, cada um vira sua
  própria pasta/`data.json`, reaproveitando o mesmo `dashboard.html`.

## Lacunas de dados encontradas na extração (não são bugs do dashboard)

- **Conteúdo secundário e concorrente**: as colunas existem em "Banco de
  treinos", mas nenhuma sessão registrada até agora preencheu nenhuma das
  duas — 100% do currículo tem só "conteúdo principal". As páginas/gráficos
  por papel (secundário/concorrente) estão prontos e vão passar a mostrar
  dado assim que essas colunas começarem a ser preenchidas.
- **3 datas com avaliação individual mas sem sessão em "Banco de treinos"**
  (02, 03 e 04/02/2026) — a minutagem de treino dessas datas não entra no
  cálculo de "minutagem do atleta" no card, porque não há uma sessão
  correspondente para atribuir os minutos.
- **2 jogos sem gols registrados** (Bonfim, Florestino) aparecem com
  resultado em branco (não mais como falso "Empate" por conta da fórmula
  original da planilha).
