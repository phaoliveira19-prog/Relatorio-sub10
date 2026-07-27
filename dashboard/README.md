# Sistema de Gestão — Sub-10

Dashboard de planejamento e monitoramento da categoria, construído a partir do
"Sistema de gestão.xlsx" (Banco de atletas, Banco de treinos, Avaliação
individual, Pós-jogo, Pós-jogo individual, Caderno de conteúdos, Tipos de
tarefa de treino, Avaliação do treino, Princípios pedagógicos).

## Como funciona hoje

```
sistema.xlsx  --(build_data.py)-->  data.json  --(fetch)-->  index.html
```

`index.html` não tem nenhum dado embutido: ao abrir, ele busca `./data.json`
(mesma pasta) e renderiza tudo em cima disso. Para atualizar o dashboard com
dados novos:

```bash
python3 build_data.py /caminho/para/Sistema_de_gestao.xlsx data.json
```

e recarregar a página (ou fazer commit/push de `data.json` se estiver hospedado
no GitHub Pages).

**Isso ainda não é "editar a planilha atualiza o site" automaticamente.**
Esse é o próximo passo pendente: migrar a fonte de dados para Google Sheets e
trocar o `fetch('./data.json')` em `index.html` por um fetch direto do Sheets
publicado como CSV (ou da Visualization API). A estrutura já foi desenhada
para essa troca ser pontual — só a função de carregamento muda, nada do resto
do dashboard.

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
`dashboard/photos/<apelido-sem-acento>.jpg`.

## Pendências conhecidas (decididas na conversa, ainda não implementadas)

- **Google Sheets como fonte viva** — decidido, não feito. Precisa da planilha
  criada/compartilhada para eu trocar o loader.
- **Autenticação (login/senha)** — não é urgente, mas o hosting via GitHub
  Pages por si só não tem isso; se for necessário antes de ter tempo para uma
  auth de verdade, dá para colocar o repositório como privado ou um gate
  simples na frente.
- **Multi-clube / multi-categoria** — o schema já carrega `categoria` em
  atletas/treinos/jogos, mas o dashboard ainda assume uma única base
  (`data.json`). Quando for expandir para outros clubes, cada um vira sua
  própria pasta/`data.json`, reaproveitando o mesmo `index.html`.

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
