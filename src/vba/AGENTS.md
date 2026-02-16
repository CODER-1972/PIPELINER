# AGENTS.md — PIPELINER (Template Excel + VBA para execução de pipelines e catálogos de prompts)

## Objetivo deste repositório (visão de alto nível)
Este repositório contém o “motor” (VBA) e as regras operacionais de um template Excel (.xlsm) que permite:
- Definir **catálogos de prompts** (em folhas de catálogo) com IDs e metadados.
- Definir **pipelines** (sequências de prompts) na folha **PAINEL** (até 10 pipelines).
- Executar pipelines via VBA, chamando a API (Responses API) e, quando aplicável, a Files API para anexos.
- Registar auditoria e troubleshooting em **Seguimento** e **DEBUG**.
- (Opcional) Gerar e descarregar ficheiros de output para um OUTPUT Folder e registar na folha FILES_MANAGEMENT.

A arquitetura separa duas camadas:
1) **Excel (dados/UI)**: folhas com prompts, pipelines, configurações e logs.
2) **VBA/API (execução)**: macros que lêem o Excel, constroem pedidos à API, encadeiam contexto e aplicam validações/limites.

Regra de ouro:
- Alterações ao VBA **não devem exigir** alterações estruturais no Excel, salvo necessidade explícita.
- Alterações às folhas devem ser feitas como “dados” (conteúdo), preservando layout, nomes e cabeçalhos.

Diretriz de documentação (README):
- O `README.md` é a referência principal de **funcionamento global do projeto** (arquitetura, fluxo, componentes e operação).
- Evitar transformar o README num guia exclusivo de um único teste/caso.
- Guias de teste específicos (ex.: T1/T3/selftests) devem ser adicionados como secções complementares ou documentos dedicados, sem substituir a visão end-to-end do sistema.

---

## “Mapa mental” do funcionamento (pipeline em 60 segundos)
Quando uma pipeline é iniciada:
1) O VBA lê a lista de IDs no PAINEL (pipeline escolhida).
2) Resolve o prompt inicial e encontra o respetivo bloco no catálogo (folha cujo nome coincide com o prefixo do ID).
3) Constrói o request com base em:
   - defaults globais (folha Config),
   - overrides no catálogo (modelo, modos, storage, config extra),
   - anexos definidos na célula INPUTS (via directiva FILES:) e regras de Config para anexos.
4) Faz a chamada à API e escreve:
   - auditoria detalhada no Seguimento,
   - alertas/erros e diagnóstico no DEBUG.
5) Determina o próximo passo via:
   - Next PROMPT determinístico, ou
   - Next PROMPT = AUTO (extrai `NEXT_PROMPT_ID: ...` do output), com fallback para Next PROMPT default.
6) Aplica limites (Max Steps, Max Repetitions, deteção de loops) e termina quando chegar a STOP/limite/erro.

---

## Estrutura do workbook e função de cada folha (INVARIANTES)

### 1) Folha “PAINEL” (obrigatória)
Finalidade:
- UI principal para gerir pipelines, limites, e botões de execução.
- Define INPUT Folder e OUTPUT Folder por pipeline.

Invariantes críticos:
- Existem **10 pipelines**, cada uma num “par de colunas”: (INICIAR à esquerda, REGISTAR à direita).
- A **lista de IDs começa na linha 9** (por pipeline).
- Deve existir um campo para:
  - Nome da pipeline (linha 1),
  - INPUT Folder (linha 2),
  - OUTPUT Folder (linha 3),
  - Max Steps (linha 4),
  - Max Repetitions (linha 5),
  - Botões INICIAR/REGISTAR (linha 7).
- Não reformatar nem “mexer no layout” do PAINEL sem alterar o VBA correspondente.

Comportamentos operacionais esperados (não quebrar):
- Ao clicar em INICIAR, o foco muda para a folha Seguimento (A1).
- Ao iniciar, a folha DEBUG é limpa (mantém cabeçalho) para evitar confusão com logs anteriores.
- Durante execução, a barra de estado do Excel mostra progresso do tipo: `(hh:mm) Step: x of y | Retry: z`.

### 2) Folha “Config” (obrigatória)
Finalidade:
- Defaults globais para execução e controlo de anexos.

Chaves base (invariantes de célula):
- Config!B1 = OPENAI_API_KEY (NUNCA versionar valores reais no repo; apenas placeholder)
- Config!B2 = MODELO (default)
- Config!B3 = TEMPERATURA (default)
- Config!B4 = MAX_OUTPUT_TOKENS (default)

Parâmetros de ficheiros (sintético — ver secção “FILES” abaixo):
- Config!B5 = FILES_TRANSPORT_MODE (FILE_ID recomendado; alternativa: INLINE_BASE64)
- Config!B6 = FILES_ENABLE_IA_FALLBACK (TRUE/FALSE)
- Config!B7 = FILES_INLINE_MAX_MB (limite para INLINE_BASE64)

Outras opções (normalmente em linhas “label -> valor”):
- FILES_UPLOAD_PROFILE (ROBUST_THEN_LEGACY | ROBUST_ONLY | LEGACY_ONLY)
- FILES_UPLOAD_ENGINE_PRIMARY (WINHTTP | MSXML)
- FILES_UPLOAD_ENGINE_FALLBACK (WINHTTP | MSXML | NONE)
- FILES_MULTIPART_FILENAME_MODE (ASCII_SAFE | RAW | RFC5987)
- FILES_UPLOAD_DEBUG_LEVEL (OFF | BASIC | VERBOSE)
- FILES_DOCX_CONTEXT_MODE (AUTO_AS_PDF | AUTO_TEXT_EMBED | ERROR)
- FILES_DOCX_AS_PDF_FALLBACK (TEXT_EMBED | ERROR)
- FILES_TEXT_EMBED_MAX_CHARS (número; default típico ~50000)
- FILES_TEXT_EMBED_OVERFLOW_ACTION (ALERT_ONLY | TRUNCATE | RETRY_AS_PDF | STOP)
- “Reutilização de ficheiros no upload” (TRUE/FALSE)

Invariantes:
- O VBA deve tolerar Config incompleta (compatibilidade retroativa): se uma chave não existir, usar defaults internos.
- Não introduzir dependências frágeis (por ex. exigir novas chaves sem fallback).

### 3) Folha “Seguimento” (obrigatória)
Finalidade:
- Auditoria por passo: prompt, modelo, config, HTTP status, output, pipeline, next decidido, etc.

Invariantes:
- Cabeçalhos devem manter o mesmo texto (a ordem pode mudar, mas os nomes devem ser estáveis).
- Deve ser suficiente para auditar:
  - que prompt correu,
  - com que parâmetros,
  - que request foi feito (pelo menos em forma de resumo),
  - que resposta veio (resumo + ids relevantes),
  - qual foi o Next PROMPT decidido.

### 4) Folha “DEBUG” (obrigatória)
Finalidade:
- Erros/alertas e troubleshooting, incluindo:
  - parsing (Config extra),
  - validações (Next PROMPT, allowed/default),
  - limites (Max Steps / Max Repetitions / loops),
  - anexos (FILES:), uploads e compatibilidades,
  - diagnósticos SELFTEST.

Invariantes:
- Cabeçalhos estáveis.
- Deve conter apenas:
  - ERRO/ALERTA/INFO relevantes,
  - sugestões curtas e acionáveis,
  - sem dados sensíveis (NUNCA API key; evitar dumps extensos).

### 5) Folha “PROMPT PARAMETROS” (recomendada)
Finalidade:
- Guia humano dos parâmetros suportados em “Config extra” (sintaxe amigável).
- NÃO é usada diretamente pela execução; serve para consistência e manutenção.

Regra:
- Quando o parser (VBA) aceitar novas chaves / mudar comportamento, atualizar este guia.

### 6) Folha “PROMPT TEMPLATE” (recomendada)
Finalidade:
- Biblioteca de secções e boas práticas para prompts (agnóstico do domínio).
- Ajuda a manter prompts consistentes e auditáveis.

Estrutura típica (não rígida):
- SEÇÃO | Itens | Propósito | Situações a não usar | Exemplos | IDs em que foi usada

### 7) Folhas de catálogo (obrigatórias)
Finalidade:
- Cada folha de catálogo contém prompts “executáveis”.

Invariantes:
- O **nome da folha tem de coincidir com o prefixo do ID** (antes do primeiro `/`).
- Não usar merges.
- As linhas “Next PROMPT” começam na coluna B (ver layout do bloco abaixo).
- O VBA assume um layout por prompt (blocos fixos).

### 8) Folha “FILES_MANAGEMENT” (se presente/ativa)
Finalidade:
- Auditoria de ficheiros: file_id, hash, reutilização, last_used_at, used_in_prompts, caminhos, e eventos de download/output.

Regra:
- Se o módulo de gestão de ficheiros estiver ativo, esta folha deve existir e ser tratada como “log/auditoria”.

---

## Convenções de IDs (prompt IDs) e nomenclatura
Formato recomendado (compatível com o template):
`ID = <NomeDaFolha>/<número>/<nomeCurto>/<versão>`

Exemplo:
`AvalCap/02/Poema/A`

Regras:
- Prefixo do ID (antes do primeiro `/`) == nome exato da folha onde o prompt está definido.
- Evitar espaços; usar nomes curtos e estáveis (CamelCase quando necessário).
- Número sequencial dentro da folha (01, 02, 03...).
- Versão (A, B, C...) muda quando há alterações relevantes ao comportamento do prompt.
- `STOP` é sentinela (não é ID de prompt).

---

## Layout do catálogo (bloco por prompt) — INVARIANTE
Cada prompt ocupa um bloco de **5 linhas**:
- Linha principal (linha do ID): campos executáveis.
- +3 linhas de configuração/documentação.
- +1 linha em branco.

### Linha principal (linha do ID) — colunas esperadas
A: ID (lookup)
B: Nome curto (documentação)
C: Nome descritivo (documentação)
D: Texto prompt (conteúdo enviado por defeito como input)
E: Modelo (override do default Config!B2)
F: Modos (ativa tools conforme lógica do VBA; ex.: Web search / None)
G: Storage (store:true/false para a API)
H: Config extra (amigável; convertido para JSON)
I: Comentários (documentação)
J: Notas para desenvolvimento (documentação)
K: Histórico de versões (documentação)

### Linhas seguintes (sem merges; texto começa na coluna B)
Linha +1: `Next PROMPT: <ID | AUTO | STOP | vazio>`
Linha +2: `Next PROMPT default: <ID | STOP>`
Linha +3: `Next PROMPT allowed: <ALL; STOP | lista separada por ';'>`

Documentação estrutural em C/D (padrão recomendado):
- C: `Descrição textual:` | D: (descrição)
- C: `INPUTS:` | D: (inputs necessários; aqui entra FILES:)
- C: `OUTPUTS:` | D: (outputs esperados)


### Nota operacional para testes de injeção por placeholder (ContextKV)
Quando estiver a validar mecanismos de **injeção por placeholder** (por exemplo, `{{VAR:RESULTS_JSON}}`):
- O placeholder tem de existir **no Texto da prompt** (coluna D), porque é isso que entra no request enviado à API.
- Uma linha isolada `VARS: ...` é útil para pedir variáveis por “directiva”, mas **não testa** o caminho de *placeholder injection*.
- Para um teste E2E completo, inclua ambos quando necessário:
  - `{{VAR:RESULTS_JSON}}` (valida detecção/substituição de placeholder)
  - `VARS: REGISTO_PESQUISAS` (valida pedido explícito de variável sem placeholder)

Isto evita falsos negativos em testes (por exemplo, ver `(SEM_VAR_...)` no output apesar de a captura existir) e torna o diagnóstico mais imediato.

---

## “Next PROMPT” e execução condicional (pipeline dinâmica)
O template suporta 3 padrões:
1) Determinístico: Next PROMPT = `<ID>`
2) Terminação: Next PROMPT = `STOP` (ou vazio)
3) Dinâmico: Next PROMPT = `AUTO`

### Quando Next PROMPT = AUTO (invariante de parsing)
O VBA tenta extrair do output uma linha no formato exato:
`NEXT_PROMPT_ID: <ID completo ou STOP>`

Boas práticas (para reduzir falhas):
- Incluir sempre `NEXT_PROMPT_ID: ...` numa linha isolada, sem ruído.
- Opcional (para auditoria humana):
  - `DECISION: ...`
  - `RATIONALE: ...` (1–2 linhas)

Fallback obrigatório:
- Se o output não tiver `NEXT_PROMPT_ID`, o VBA usa `Next PROMPT default`.

### Allowed/default (controlo de integridade)
- `Next PROMPT default` deve estar sempre preenchido (ou STOP), salvo intenção explícita de terminar.
- `Next PROMPT allowed` deve incluir:
  - o default,
  - e STOP,
  - e (idealmente) restringir opções para evitar saltos inesperados.

### Controlo de loops e limites (não quebrar)
- Max Steps: limite absoluto de passos por execução (PAINEL).
- Max Repetitions: limite por ID (PAINEL).
- Deteção de alternância A↔B (padrão A-B-A-B): interromper com alerta.
- STOP explícito termina.

---

## Campo “Config extra” (amigável) — sintaxe, validações e precedência
Objetivo:
- Permitir parâmetros extra sem proliferar colunas.

Formato:
- Uma linha = `chave: valor` (linhas separadas por ALT+ENTER).
- Linhas sem `:` são ignoradas com alerta em DEBUG.
- Chaves podem usar notação com pontos (nesting): `text.format.type`
- Listas: `[a, b, c]`
- Objetos: `{k: v, k2: v2}`

Bloco `input:` (override do conteúdo enviado):
- Após `input:`, só aceitar `role:` e `content:` (indentados).

Chaves proibidas em Config extra (devem ser ignoradas com alerta):
- `model`, `temperature`, `max_output_tokens`, `store`, `tools`
(Estas chaves têm colunas dedicadas / lógica própria.)

Regras de coerência:
- Não usar `conversation` e `previous_response_id` em simultâneo; se ambos aparecerem, manter `conversation` e ignorar `previous_response_id` (com alerta).

---

## Directiva FILES: (anexos) — comportamento e invariantes
### Onde e como declarar
Na célula INPUTS (no bloco do prompt; tipicamente linha +2, coluna D):
- `FILES: ficheiro1.pdf; imagem1.png; doc1.docx (as pdf)`
Separador suportado: `;`
Sinónimo suportado: `FICHEIROS:` (com ou sem espaço antes de `:`)

### Flags reconhecidas por ficheiro (exatas)
Entre parênteses no mesmo item:
- Obrigatoriedade:
  - `(required)` ou `(obrigatorio)` ou `(obrigatoria)`
- “mais recente” (wildcards):
  - `(latest)` ou `(mais recente)` ou `(mais_recente)`
- Modo:
  - `(as pdf)` ou `(as_pdf)`
  - `(as is)` ou `(as_is)`
  - `(text)` ou `(text_embed)`

### Regra base: onde o sistema procura
- O VBA resolve ficheiros **apenas** dentro do INPUT Folder do pipeline (PAINEL, linha 2).
- Por segurança, deve ignorar caminhos que tentem sair da pasta (usar apenas filename).
- Wildcard `*` procura dentro da própria pasta (não contar com subpastas).

### Precedência de configuração (quando há conflito)
1) Overrides na directiva FILES (por ficheiro)
2) Config (global)
3) Defaults internos (fallback/compatibilidade)

### Transport mode (como o conteúdo chega ao request)
- `FILES_TRANSPORT_MODE=FILE_ID` (recomendado):
  - Faz upload para `/v1/files` quando necessário e referencia por `file_id`.
  - Pode reutilizar uploads via cache (hash + modo + nome).
- `FILES_TRANSPORT_MODE=INLINE_BASE64`:
  - Embute como data URL no request; sujeito a `FILES_INLINE_MAX_MB`.

### Compatibilidade /v1/responses (ponto crítico)
- Alguns formatos Office (DOC/DOCX/PPT/PPTX/XLSX/...) podem ser aceites no `/v1/files`, mas **não são aceites como `input_file`** no `/v1/responses`.
- Por isso, existe “effective_mode” para evitar falhas a meio:
  - Se alguém pedir `(as is)` para DOCX/PPTX, o sistema pode fazer override automático para:
    - PDF (`AUTO_AS_PDF`), ou
    - text_embed (`AUTO_TEXT_EMBED`),
    - ou erro (`ERROR`), conforme Config.

### Comportamento por tipo (síntese)
A) PDF
- Default: upload como PDF para o modelo (modo “pdf_upload”).
- Recomendação: usar PDF quando precisa de layout/tabelas/imagens.

B) Imagens (png/jpg/jpeg/webp)
- Default: upload como imagem (“image_upload”).

C) Word (DOC/DOCX)
- Default: `text_embed` (extração de texto), se não disser nada.
- `(as pdf)`: converte para PDF (cache `_pdf_cache`) e envia como PDF.
- `(as is)`: pode ser override por effective_mode (Config: FILES_DOCX_CONTEXT_MODE).

D) PowerPoint (PPT/PPTX)
- Similar a DOCX: default `text_embed`; `(as pdf)` converte; `(as is)` pode ser override.

E) Excel (XLS/XLSX/XLSM)
- Em geral, não há extração robusta (text_embed) nem conversão automática fiável no fluxo atual.
- Recomendação prática: exportar previamente para PDF/CSV e anexar esses artefactos.

### Controlo de tamanho (text_embed)
- `FILES_TEXT_EMBED_MAX_CHARS` limita chars do texto extraído.
- Se exceder:
  - emitir alerta `TEXT_EMBED_TOO_LARGE`,
  - aplicar `FILES_TEXT_EMBED_OVERFLOW_ACTION`:
    - ALERT_ONLY | TRUNCATE | RETRY_AS_PDF | STOP.

### Upload robustness (evitar 415/invalid multipart)
- `FILES_UPLOAD_PROFILE` controla estratégia (ROBUST_THEN_LEGACY recomendado).
- `FILES_UPLOAD_ENGINE_PRIMARY` + fallback (WINHTTP/MSXML).
- `FILES_MULTIPART_FILENAME_MODE` (ASCII_SAFE por defeito) para evitar problemas com Unicode/acentos em filename.
- Em erros de upload, registar no DEBUG:
  - HTTP status,
  - inputFolder, resolvedPath, FileExists=SIM/NAO,
  - engine e profile usados.

### Roteiro de verificação (3 checks) — recomendado em troubleshooting
1) JSON final contém `input_file/input_image` e `file_id` quando esperado
2) Upload `/v1/files` devolve HTTP 2xx (quando aplicável)
3) FILES_MANAGEMENT tem registo (file_id, hash, last_used_at, used_in_prompts)

### Fallback IA (ambiguidade) — usar com critério
- `FILES_ENABLE_IA_FALLBACK=TRUE` pode fazer chamada adicional para escolher o ficheiro quando há ambiguidades.
- Deve ser usado apenas quando ambiguidade é frequente e não há regra determinística.

---

## File Output (criação/descarregamento de ficheiros)
Objetivo:
- Permitir que uma prompt produza um artefacto final (.xlsx/.docx/.pptx/.pdf/.csv) guardado no OUTPUT Folder.

Invariantes:
- OUTPUT Folder é definido no PAINEL (linha 3) e a pasta tem de existir/permissões.
- Prompt deve pedir explicitamente um ficheiro e indicar nome + extensão.
- Registar evento em FILES_MANAGEMENT (path completo, tipo, notas, response_id quando aplicável).
- Em caso de colisão de nome (ficheiro já existe):
  - preferir sufixo com timestamp/versão,
  - ou política clara de overwrite (se existir).

---

## Regras de segurança e privacidade (não negociáveis)
- NUNCA commitar API keys reais (Config!B1 deve ser placeholder).
- NUNCA escrever segredos em DEBUG/Seguimento.
- Minimizar exposição de dados sensíveis nos logs (mascarar quando necessário).
- Quando falhar, preferir logs “curtos e acionáveis” a dumps extensos.

---

## Compatibilidade retroativa e disciplina de mudanças
Quando alterar VBA:
- Preservar compatibilidade com templates existentes sempre que possível.
- Se introduzir nova configuração:
  1) adicionar default interno,
  2) tolerar ausência da chave,
  3) atualizar PROMPT PARAMETROS e documentação,
  4) adicionar/atualizar SelfTest correspondente (se aplicável).

Quando alterar prompts/catálogos:
- Se o comportamento mudar, atualizar versão do ID (A->B->C...).
- Manter regras de Next PROMPT (AUTO/default/allowed) coerentes para evitar loops e custos.

---

## Padrões de qualidade VBA (para revisão e alterações)
### Checklist mínimo antes de aceitar um PR com VBA
(1) Host e compatibilidade
- Assumir Excel como host principal.
- Se usar APIs específicas, aplicar guardas e fallbacks.

(2) Auto-contenção
- Código entregue/alterado deve incluir dependências necessárias (evitar chamadas para procedimentos “fantasma”).

(3) Option Explicit e símbolos
- `Option Explicit` obrigatório.
- Sem variáveis não declaradas / nomes inconsistentes.

(4) Objectos e Set
- Atribuições a objectos usam `Set` (incluindo quando o destino é `Variant`).
- Ao ler `Scripting.Dictionary.Item(key)` cujo valor pode ser objecto, testar `IsObject(...)` e usar `Set` para evitar erro 450.
- Em dictionaries/collections, evitar atribuições encadeadas ambíguas; usar helpers consistentes.

(5) Guardas por tipo/capacidade
- Antes de chamar métodos/propriedades não universais, validar TypeName/.Type/.HasX ou usar `On Error Resume Next` apenas em blocos curtos + verificação do resultado.

(6) Late binding e constantes
- Preferir late binding para portabilidade.
- Evitar enums dependentes de versão; usar `Const ... As Long = ...` quando necessário.

(7) Integridade estrutural (copy-paste seguro)
- Sem `With` sem `End With`, sem `Select Case` sem `End Select`, sem linhas truncadas.

(8) Idempotência
- Quando cria/atualiza artefactos (folhas, tabelas, shapes, ficheiros), evitar duplicação em execuções repetidas.

(9) Tratamento e reporte de erros
- Entry points principais com `On Error GoTo EH` e reporte mínimo (rotina, Err.Number, Err.Description; Erl se aplicável).
- Evitar `On Error Resume Next` como “modo global”.

(10) Smoke tests / SelfTest
- Manter rotinas de teste idempotentes (ex.: `SelfTest_RunAll`) e registar PASS/FAIL no DEBUG.

---

## Instruções para o Codex durante code review (como avaliar PRs)
### Prioridades (classificação de severidade)
- P0 (bloqueador): quebra execução, perda de dados, regressão de compatibilidade, erros de anexos/requests, loops infinitos, fuga de segredos.
- P1 (alto): edge cases frequentes, logs insuficientes para diagnosticar, validações fracas (Next PROMPT), parsing frágil.
- P2 (médio): clareza/legibilidade, duplicação, pequenas melhorias de robustez.
- P3 (baixo): estilo, micro-optimizações, “nitpicks”.

### O que o Codex deve sempre verificar
- Não há alterações inadvertidas ao layout esperado (PAINEL, Config, Seguimento, DEBUG, catálogo).
- Next PROMPT e regras de AUTO/default/allowed mantêm-se consistentes.
- FILES: continua compatível (flags, effective_mode, limites, logs).
- Não foi introduzida dependência que exija Office apps ausentes sem fallback.
- Não há segredos/versionamento indevido.
- Todas as alterações que atualizem ou modifiquem o funcionamento devem levar a uma revisão e atualização do ficheiro README, e isso deve ser reportado no chat.
- Todos os erros detectados no VBA, se puderem ser detectados, devem levar a uma atualização no ficheiro src\vba\AGENTS.md, com a revisão e o enunciar de regras ou princípios agnósticos que evitem a repetição do erro.

### Como propor alterações
- Preferir diffs pequenos e testáveis.
- Não fazer refactors massivos “por limpeza” sem necessidade.
- Quando sugerir mudança, indicar:
  - risco,
  - impacto em compatibilidade,
  - como validar (passos de teste / onde ver logs).

---

## Nota sobre tamanho e manutenção das instruções
Este AGENTS.md é propositadamente detalhado, mas deve manter-se “operacional”.
Se começar a ficar demasiado grande:
- mover detalhe para `docs/*.md`,
- criar `src/vba/AGENTS.md` com regras específicas de VBA/módulos,
- manter aqui apenas invariantes e instruções de revisão.

(As instruções mais próximas do ficheiro alterado devem ter prioridade.)


## Checklist anti-erros de sintaxe em strings VBA
- Ao remover aspas duplas em `Replace`, usar literal válido de VBA (`""""`) ou `Chr$(34)`; nunca usar `"""` porque gera erro de compilação.
- Sempre que editar strings com escape (JSON, regex-like, Replace), executar verificação rápida no VBE (Debug > Compile VBAProject) antes de fechar a alteração.
- Em detecção de diretivas via `InStr`, normalize primeiro o texto (espaços/aspas) e compare também por igualdade canónica (`s = "environ(openai_api_key)"`) para evitar `Type mismatch` por string mal escapada.
