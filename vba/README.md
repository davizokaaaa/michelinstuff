# Macro de classificação — módulos VBA

Todos os módulos `.bas` deste diretório formam a macro `ClassificarTudo` (em `modMain.bas`).
Importe-os no VBA Editor do Excel (Arquivo > Importar Arquivo, ou arraste pro Project Explorer)
e rode `modMain.ClassificarTudo`.

## O que mudou nesta revisão

### 1. Performance (sem alterar nenhum resultado)

- **Leitura/escrita em bloco** (`modMain.bas`): as colunas de entrada (`DESCRIÇÃODA MERCADORIA`,
  `PROVÁVEL ADQUIRENTE`) agora são lidas uma única vez num array em memória, e todas as colunas
  de saída são escritas de uma vez só no final (`Range.Value = array`), em vez de
  `ws.Cells(i, col)` por linha. Cada acesso `Cells` é uma chamada COM cara; um array em memória
  não é. Essa é a maior causa provável dos travamentos em bases grandes.
- **Cache de resultado por linha** (`modMain.bas`): como toda a extração (MARCA, GAMA, DIMENSÃO,
  SEGMENTO, LP, ARO, RT/OE, ANIP, ESTEPE/COMPETIÇÃO) depende só de
  `(DESCRIÇÃODA MERCADORIA, PROVÁVEL ADQUIRENTE)` — os dicionários carregados no início não mudam
  durante o laço —, linhas com a mesma dupla de valores reaproveitam o resultado já calculado em
  vez de refazer toda a busca. Se a base tiver descrições repetidas (comum em importação), o
  ganho pode ser grande.
- **RegExp cacheado com `Static`** (`modDimensao.bas`, `modDimensaoExtenso.bas`): os objetos
  `VBScript.RegExp` eram recriados a cada chamada; agora são criados uma vez e reaproveitados
  (só o `.Pattern` muda quando necessário). Essas funções rodam várias vezes por linha, em toda
  a base.
- **Bug de cache corrigido** (`modMarca.bas`, `SeguidoPorRotuloCampo`): a lista de rótulos estava
  declarada `Static`, mas era **reatribuída incondicionalmente** a cada chamada — ou seja, o
  cache nunca funcionava de fato. Agora só é montada na primeira chamada.

Nenhuma dessas mudanças altera o valor calculado em nenhuma linha — só a forma de ler/escrever/
evitar recálculo do que já seria calculado de qualquer forma. **Recomendo validar isso na
prática**: rode a macro antiga e a nova na mesma planilha (em cópias/abas separadas) e compare
as colunas de saída linha a linha (ex: com uma fórmula `=EXATO(...)` numa coluna auxiliar, ou um
`COUNTIF` de diferenças). Se algo divergir, me avise antes de usar em produção.

### 2. Nova feature: ESTEPE / COMPETIÇÃO / LCV

Módulo novo `modEstepeCompeticao.bas`, função `ClassificarEstepeCompeticao(texto)`. Portada 1:1
da macro legada `Sub estepe()` (mesmas palavras-chave, mesma ordem de prioridade em cascata
`If/ElseIf`, mesmo `vbTextCompare`):

1. `ESTEPE`
2. `CAMPEONATO`
3. `COMPETICAO`
4. Um conjunto fixo de medidas/textos específicos (`225/65 R16 C`, `205/75R16C` etc.) → `LCV`
5. Nenhum bateu → `""`

Integrada ao laço principal de `modMain.ClassificarTudo`: roda na mesma passada, usando a mesma
`DESCRIÇÃODA MERCADORIA` já lida (não precisa de uma segunda varredura pela planilha), e grava
numa coluna nova, `ESTEPE/COMPETICAO` (nome em `modConfig.COL_ESTEPE_COMPETICAO_NOME`), criada
automaticamente se não existir.

**Premissa a confirmar**: a macro legada lia da coluna `H` (fixa) e escrevia na coluna `T`
(fixa). Assumi que a coluna `H`, na planilha onde a legada rodava, é a mesma coisa que
`DESCRIÇÃODA MERCADORIA` (mesmo texto usado por MARCA/GAMA/DIMENSÃO aqui). Se não for, é só
trocar qual variável é passada para `ClassificarEstepeCompeticao` em `modMain.bas`.

## Plano de teste sugerido

1. Importe todos os `.bas` num workbook de teste (cópia da planilha real).
2. Rode `ClassificarTudo` numa base pequena conhecida primeiro.
3. Compare coluna a coluna com uma execução da versão anterior da macro (mesma base, mesma
   Tabela de Referência, mesma resposta SIM/NÃO no diálogo BR/MIN).
4. Confira a nova coluna `ESTEPE/COMPETICAO` contra o resultado da macro legada `estepe()`
   rodando a parte (mesma base, comparando coluna a coluna).
5. Rode numa base grande e cronometre antes/depois.
