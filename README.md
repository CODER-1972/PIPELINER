# PIPELINER

Template Excel + VBA para execução de pipelines de prompts com auditoria operacional, integração API e gestão de contexto.

## Índice hierárquico

- [1. O que é o PIPELINER](#1-o-que-é-o-pipeliner)
- [2. Arquitetura do projeto](#2-arquitetura-do-projeto)
  - [2.1 Camada Excel (dados/UI)](#21-camada-excel-dadosui)
  - [2.2 Camada VBA/API (execução)](#22-camada-vbaapi-execução)
- [3. Estrutura funcional do workbook](#3-estrutura-funcional-do-workbook)
  - [3.1 PAINEL](#31-painel)
  - [3.2 Config](#32-config)
  - [3.3 Seguimento](#33-seguimento)
  - [3.4 DEBUG](#34-debug)
  - [3.5 Folhas de catálogo](#35-folhas-de-catálogo)
  - [3.6 FILES_MANAGEMENT](#36-files_management)
- [4. Modelo de IDs e catálogo de prompts](#4-modelo-de-ids-e-catálogo-de-prompts)
- [5. Fluxo de execução de uma pipeline](#5-fluxo-de-execução-de-uma-pipeline)
  - [5.1 Resolução do prompt e configuração efetiva](#51-resolução-do-prompt-e-configuração-efetiva)
  - [5.2 Chamada API e auditoria](#52-chamada-api-e-auditoria)
  - [5.3 Resolução de Next PROMPT](#53-resolução-de-next-prompt)
  - [5.4 Limites e proteção contra loops](#54-limites-e-proteção-contra-loops)
- [6. Campo Config extra (sintaxe amigável)](#6-campo-config-extra-sintaxe-amigável)
- [7. FILES: anexos, upload e compatibilidade](#7-files-anexos-upload-e-compatibilidade)
- [8. ContextKV (captura e injeção de variáveis)](#8-contextkv-captura-e-injeção-de-variáveis)
- [9. Logs, troubleshooting e validação operacional](#9-logs-troubleshooting-e-validação-operacional)
- [10. Segurança e compatibilidade retroativa](#10-segurança-e-compatibilidade-retroativa)
- [11. Guia rápido de operação](#11-guia-rápido-de-operação)

---

## 1. O que é o PIPELINER

O PIPELINER é um motor de execução em VBA, acoplado a um template Excel (`.xlsm`), para:

- gerir **catálogos de prompts** por ID;
- montar e correr **pipelines** de múltiplos passos no PAINEL;
- chamar a **Responses API** (e, quando aplicável, **Files API**);
- guardar rastreabilidade completa em **Seguimento** e **DEBUG**.

Objetivo principal: dar uma forma auditável e operacional de executar fluxos com IA sem perder controlo de parâmetros, contexto e outputs.

---

## 2. Arquitetura do projeto

### 2.1 Camada Excel (dados/UI)

Inclui folhas de configuração e operação:

- definição de prompts e metadados;
- parametrização global;
- sequência de pipelines;
- auditoria e troubleshooting.

### 2.2 Camada VBA/API (execução)

Inclui módulos que:

- leem dados estruturados do workbook;
- convertem configurações amigáveis para payloads API;
- gerem anexos e uploads;
- executam chamadas API;
- persistem auditoria por passo;
- resolvem encadeamento (`Next PROMPT`) até `STOP`.

---

## 3. Estrutura funcional do workbook

## 3.1 PAINEL

Ponto de operação principal:

- 10 pipelines (pares INICIAR/REGISTAR);
- nome da pipeline;
- `INPUT Folder`, `OUTPUT Folder`;
- limites (`Max Steps`, `Max Repetitions`);
- botões de execução.

Comportamentos esperados:

- foco em `Seguimento!A1` no arranque;
- limpeza de DEBUG da execução anterior;
- status bar com progresso de execução.
- durante cada passo, a status bar inclui fase operacional antes da execução (ex.: `A preparar passo`, `Uploading file`, `A executar prompt`).

## 3.2 Config

Defaults e opções globais, incluindo:

- credenciais/API key (placeholder em repositório);
- modelo/temperatura/tokens;
- estratégia de transporte de ficheiros (`FILE_ID`/`INLINE_BASE64`);
- opções de robustez de upload e fallback.

## 3.3 Seguimento

Auditoria por passo: prompt executado, configuração usada, status HTTP, output, next prompt decidido, ficheiros usados e colunas de contexto (captured/injected).

## 3.4 DEBUG

Registo curto e acionável de erros/alertas/info de parsing, validação de encadeamento, limites e troubleshooting técnico.

## 3.5 Folhas de catálogo

Cada folha contém prompts executáveis. O prefixo do ID deve corresponder ao nome da folha.

## 3.6 FILES_MANAGEMENT

Folha de auditoria de ficheiros (upload/reutilização/download/output), quando o módulo de files está ativo.

---

## 4. Modelo de IDs e catálogo de prompts

Formato recomendado:

`<NomeDaFolha>/<número>/<nomeCurto>/<versão>`

Exemplo:

`AvalCap/02/Poema/A`

Regras importantes:

- prefixo do ID = nome exato da folha;
- manter IDs estáveis;
- usar `STOP` como sentinela de término.

Cada prompt ocupa bloco fixo no catálogo:

- linha principal com campos executáveis (`ID`, prompt, modelo, modos, storage, config extra);
- linhas de `Next PROMPT`, `default`, `allowed`;
- documentação de `INPUTS`/`OUTPUTS`.

---

## 5. Fluxo de execução de uma pipeline

## 5.1 Resolução do prompt e configuração efetiva

Para cada passo, o motor:

1. lê o ID atual no PAINEL;
2. resolve a definição no catálogo;
3. aplica defaults globais + overrides do prompt;
4. converte `Config extra` para fragmentos JSON válidos;
5. prepara input com/sem anexos.

## 5.2 Chamada API e auditoria

Depois de montar o payload:

- executa chamada à API;
- escreve registo no `Seguimento`;
- escreve eventos técnicos no `DEBUG` quando aplicável.

## 5.3 Resolução de Next PROMPT

Suporta:

- próximo prompt determinístico;
- `AUTO` (extração de `NEXT_PROMPT_ID: ...` do output);
- fallback para `Next PROMPT default`;
- validação com `Next PROMPT allowed`.

## 5.4 Limites e proteção contra loops

Execução termina por:

- `STOP` explícito;
- `Max Steps`;
- `Max Repetitions` por ID;
- deteção de alternância A-B-A-B.

---

## 6. Campo Config extra (sintaxe amigável)

Formato por linha: `chave: valor`.

Suporta:

- nesting por pontos (`a.b.c`);
- listas (`[a,b,c]`);
- objetos (`{k:v}`);
- bloco `input:` com `role`/`content`.

Comportamentos de robustez:

- linhas inválidas são ignoradas com alerta no DEBUG;
- chaves proibidas (ex.: `model`, `tools`) são ignoradas com alerta;
- conflitos de parâmetros de encadeamento são resolvidos de forma determinística;
- serialização recursiva de dicionários aninhados usa atribuição segura com `Set` para itens `Object` do `Scripting.Dictionary` (evita erro 450 em estruturas mistas).

---

## 7. FILES: anexos, upload e compatibilidade

Na linha de `INPUTS` do prompt é possível declarar `FILES:`/`FICHEIROS:`.

As linhas de `INPUTS` são anexadas ao prompt final enviado ao modelo num bloco dedicado `INPUTS_DECLARADOS_NO_CATALOGO`, incluindo `URLS_ENTRADA`, `MODO_DE_VERIFICACAO` e também a própria declaração `FILES:`/`FICHEIROS:` como contexto textual. O anexo técnico dos ficheiros continua a ser tratado pelo módulo M09. Esse mesmo texto final montado é o que segue para o `input_text` quando o M09 prepara anexos.

Capacidades principais:

- resolução de ficheiros no `INPUT Folder` da pipeline;
- flags por ficheiro (`required`, `latest`, `as pdf`, `as is`, `text`);
- suporte a wildcard em `FILES:` (ex.: `GUIA_DE_ESTILO*.pdf`), com tentativa inicial por `Dir` e fallback de correspondência flexível para nomes com `_`, `-` e espaço;
- upload para `/v1/files` com reutilização por hash (quando configurado);
- fallback entre engines/perfis de upload.

Nota de compatibilidade importante:

- nem todos os formatos aceites no upload são aceites como `input_file` no `/v1/responses`;
- o sistema pode aplicar `effective_mode` (ex.: converter para PDF ou text embed) conforme configuração.

---

## 8. ContextKV (captura e injeção de variáveis)

O módulo ContextKV permite:

- **capturar** blocos estruturados do output de um passo (`captured_vars`, `captured_vars_meta`);
- **injetar** variáveis em passos seguintes via `{{VAR:...}}`, `VARS:` e `{@OUTPUT: ...}`;
- registar eventos operacionais no DEBUG (`INJECT_*`, `CAPTURE_*`).

É útil para pipelines multi-etapa onde uma resposta precisa ser reutilizada de forma controlada no passo seguinte.

---

## 9. Logs, troubleshooting e validação operacional

Boas práticas de manutenção VBA (preventivas):

- em literais de string com aspas duplas, usar escaping válido do VBA (ex.: `""""`) ou `Chr$(34)`;
- em comparações `If ... = "` e listas `Select Case` para aspas, confirmar literal completo (`""""`) para evitar `Syntax error`;
- em padrões regex com aspas dentro de classe de caracteres (ex.: `[^\"]`), duplicar aspas no literal VBA (ex.: `"""([^""]+)"""`) para evitar erro de compilação;
- em rotinas de escape/unescape JSON, validar o par inverso de `Replace` (escape: `\ -> \\`, `" -> \"`; unescape: `\\ -> \`, `\" -> "`) para não corromper conteúdo silenciosamente;
- após alterações em módulos `.bas`, correr compilação do projeto (`Debug > Compile VBAProject`) para apanhar erros de sintaxe antes de execução.


### Diagnóstico rápido: web_search + anexos + ContextKV

Regra atual do PIPELINER: quando `Modos` contém `Web search`, o payload deve incluir `tools:[{"type":"web_search"}]` por auto-injeção, mesmo que existam anexos (`input_file`/`input_image`).

Checklist objetivo:

1. Confirmar `REQ_INPUT_JSON` com `has_input_file=SIM` e `file_id=file-...` quando o modo de transporte for `FILE_ID`.
2. Confirmar `M05_PAYLOAD_CHECK` com `web_search=ADICIONADO_AUTO` sempre que `Modos=Web search`.
3. Se `web_search` não for auto-adicionado, validar se existe `tools` explícito no fragmento extra (`web_search=NAO_AUTO (tools no extra)`).

Para ContextKV, `CAPTURE_MISS` significa que o output não trouxe rótulos capturáveis esperados (`RESULTS_JSON`, `NEXT_PROMPT_ID`, `MEMORY_SHORT`, etc.). Para aumentar taxa de `CAPTURE_OK`, incluir no prompt instruções explícitas para devolver pelo menos:

- `RESULTS_JSON:` (linha com JSON ou bloco fenced);
- `NEXT_PROMPT_ID: STOP` (ou ID válido, se a pipeline usar AUTO).


Nota: `tools` continua como chave proibida em `Config extra` (é ignorada com alerta), para preservar a coerência com as colunas/lógica dedicadas.

### Diagnóstico rápido: `Erro VBA: The operation timed out`

Quando o `Seguimento` mostra `HTTP Status=0` e `Output=[ERRO] Erro VBA: The operation timed out`, a falha tende a acontecer no cliente HTTP (tempo de espera do host/engine) e não necessariamente num erro de validação do payload.

Checklist recomendado (ordem prática):

1. Confirmar no `DEBUG` se existe `M05_PAYLOAD_CHECK` com `has_input_file=SIM/NAO`, `web_search=...`, `model=...` e `payload_len=...` para validar se o pedido final foi mesmo montado.
2. Confirmar se existe `M05_PAYLOAD_DUMP` e abrir o `payload.json` gravado para inspeção local (estrutura JSON, tamanho e blocos `tools`/`input`).
3. Se `process_mode=code_interpreter`, confirmar se o run devolveu `rawResponseJson`; evento `M10_CI_RAW_MISSING` indica que o fluxo CI não trouxe corpo bruto para pós-processamento e deve ser tratado como pista de diagnóstico, não como causa raiz isolada.
4. Medir tamanho de entrada efetiva (`REQ_INPUT_JSON len=...`): payloads muito grandes (texto + anexos + instruções extensas) aumentam risco de timeout no host VBA.
5. Repetir teste com redução controlada de carga:
   - remover temporariamente `process_mode: code_interpreter`;
   - reduzir anexos a 1 ficheiro essencial;
   - testar com prompt curto (smoke test) no mesmo modelo.
6. Se o timeout persistir com payload pequeno, validar conectividade e engine HTTP ativa (WinHTTP/MSXML), além de quota/latência do endpoint.

Sinais úteis para separar causas:

- `FILES ... Anexacao OK` + `has_input_file=SIM` + `timeout` => anexação concluída, falha provável em execução/resposta.
- `HTTP 4xx/5xx` com body => erro API explícito (não timeout de cliente).
- `timeout` sem `M05_PAYLOAD_CHECK` => falha antes da montagem final (inspecionar parsing/configuração).

SelfTests recomendados para este cenário:

- `SelfTest_WebSearchGating` (com/sem anexos; validar que a mensagem de `M05_PAYLOAD_CHECK` permanece `web_search=ADICIONADO_AUTO`);
- `SelfTest_PayloadHasInputFileId` (valida `REQ_INPUT_JSON` e presença de `file_id`);
- `SelfTest_ContextKV_CaptureOkMiss` (2 outputs sintéticos: um capturável e outro livre);
- `SelfTest_InputsKvExtraction` (linhas `CHAVE: valor` e `CHAVE=valor`, com exclusão de `FILES:`).
- `SELFTEST_FILES_WILDCARD_RESOLUTION` (cria pasta temporária + dummies `GUIA_DE_ESTILO*.pdf` e valida escolha do mais recente com `(latest)`).

Macros utilitárias para troubleshooting rápido de catálogo + Config extra:

- `TOOL_CreateCatalogTemplateSheet` (M15): cria uma nova folha de catálogo com layout compatível (headers A:K, bloco de 5 linhas, `Next PROMPT` e secções `Descrição textual/INPUTS/OUTPUTS`).
- `TOOL_RunConfigExtraSequentialDiagnostics` (M15): executa uma bateria sequencial de casos de `Config extra`, converte via parser oficial (`ConfigExtra_Converter`), injeta fragmento de File Output (`json_schema`) e valida a estrutura JSON final antes do HTTP.
- `Files_Diag_TestarResolucaoWildcard` (M09): testa resolução de anexos `FILES:` com wildcard (ex.: `GUIA_DE_ESTILO*.pdf`) e regista no DEBUG quantos candidatos foram encontrados por `Dir` e por fallback normalizado, além do `status` final (`OK`/`AMBIGUOUS`/`NOT_FOUND`).
- Resultado do diagnóstico fica em `CONFIG_EXTRA_TESTS` + linhas `INFO/ERRO` no `DEBUG` (`M15_CONFIG_EXTRA_DIAG`), com detalhe de causa (ex.: `fecho_sem_abertura`).


### Padrão recomendado para mensagens de aviso/erro

Para tornar mensagens mais informativas e acionáveis, usar sempre 4 partes:

- **PROBLEMA**: o que falhou (facto observável, sem ambiguidade);
- **IMPACTO**: consequência direta na execução;
- **AÇÃO**: próximo passo objetivo para recuperação;
- **DETALHE** (opcional): contexto técnico curto (ex.: `payload_len`, `http_status`, `file_id`).

Formato alvo (1 linha):

`[SCOPE] PROBLEMA=... | IMPACTO=... | ACAO=... | DETALHE=...`

No VBA, o módulo `M16_ErrorMessageFormatter` disponibiliza helpers (`Diag_Format`, `Diag_WithRetryHint`, `Diag_ErrorFingerprint`) para padronizar este formato sem expor segredos.


### Explicação didática (molde para leigos)

Quando uma mensagem de erro aparece no DEBUG, quem não é técnico precisa de uma leitura "traduzida".
Use este **molde de 5 blocos** logo abaixo da mensagem técnica:

1. **O que aconteceu (em linguagem simples)**
2. **Porque isto importa (impacto prático)**
3. **O que fazer agora (passo a passo curto)**
4. **Como confirmar que ficou resolvido**
5. **Quando pedir ajuda e que evidências levar**

Exemplo didático para timeout:

- **O que aconteceu:** o sistema enviou o pedido, mas não recebeu resposta a tempo.
- **Porque importa:** este passo da pipeline ficou incompleto e os seguintes não devem avançar sem validação.
- **O que fazer agora:** (a) repetir com menos anexos, (b) reduzir texto da prompt, (c) testar sem `process_mode=code_interpreter`.
- **Como confirmar resolução:** o `Seguimento` passa a ter HTTP 2xx e o `Output` deixa de mostrar `The operation timed out`.
- **Quando pedir ajuda:** se falhar 3 vezes com payload pequeno; anexar `M05_PAYLOAD_CHECK`, `REQ_INPUT_JSON len` e `M05_PAYLOAD_DUMP`.

Exemplo didático para erro de validação de payload:

- **O que aconteceu:** o pedido foi rejeitado por formato inválido.
- **Porque importa:** o modelo não chegou a processar conteúdo; é necessário corrigir estrutura JSON/config extra.
- **O que fazer agora:** validar chaves/aspas no `Config extra`, remover trailing commas e repetir teste curto.
- **Como confirmar resolução:** aparece HTTP 2xx e desaparece o erro `invalid_json`/`invalid_json_schema`.
- **Quando pedir ajuda:** se persistir após correção local; partilhar fingerprint do erro e trecho mínimo do payload.

Regra prática de escrita para equipas mistas (técnico + negócio):
- 1 linha técnica padronizada (`PROBLEMA|IMPACTO|ACAO|DETALHE`) +
- 3–5 linhas didáticas no molde acima.

### Seguimento

Usar para auditar:

- o que correu;
- com que configuração;
- o que respondeu;
- qual o próximo passo decidido.

### DEBUG

Usar para diagnosticar:

- parsing inválido;
- encadeamento inconsistente (`Next PROMPT`);
- erros de anexos/upload;
- limites de execução;
- eventos de captura/injeção.

### Matriz de troubleshooting (10 checks com mensagens DEBUG curtas)

Objetivo: separar claramente onde a falha ocorre no ciclo completo de execução (montagem de request → API → exportação → consumo local).

| # | Check (o que validar) | Evidência esperada | Mensagem DEBUG curta (sugestão) |
|---|---|---|---|
| 1 | Prompt resolvido no catálogo | ID existe, folha prefixo confere, bloco de 5 linhas lido sem erro | `CHK01_PROMPT_OK` |
| 2 | Config efetiva montada | modelo/temperatura/tokens resolvidos com defaults+overrides | `CHK02_CONFIG_OK` |
| 3 | `Config extra` parseado sem bloqueio | fragmento JSON válido; linhas inválidas apenas com alerta | `CHK03_CFG_EXTRA_OK` |
| 4 | `Next PROMPT` consistente | `AUTO/default/allowed` válidos e compatíveis com fallback | `CHK04_NEXT_OK` |
| 5 | FILES resolvidos no INPUT Folder | paths resolvidos sem fuga de diretório; `required` respeitado | `CHK05_FILES_RESOLVE_OK` |
| 6 | Upload/attach técnico concluído | `file_id` presente quando `FILE_ID`; `input_image`/`input_file` no payload | `CHK06_ATTACH_OK` |
| 7 | Request enviado com HTTP 2xx | status 200/201 e `response_id` registado | `CHK07_HTTP_OK` |
| 8 | Output parseável para decisão | `NEXT_PROMPT_ID` extraído (ou default aplicado sem erro) | `CHK08_OUTPUT_PARSE_OK` |
| 9 | Exportação para OUTPUT Folder concluída | ficheiro final com path auditável e nome sem colisão silenciosa | `CHK09_EXPORT_OK` |
| 10 | Consumo local/COM concluído (se aplicável) | `FileExists=SIM` no destino final e abertura/uso sem timeout | `CHK10_LOCAL_CONSUME_OK` |

Convenção curta recomendada para falhas, mantendo logs acionáveis:

- `CHKxx_*_FAIL` + causa direta (`NOT_FOUND`, `INVALID_JSON`, `HTTP_4XX`, `TIMEOUT_COM`, `PATH_BLOCKED`);
- `CHKxx_*_WARN` para degradações controladas (ex.: fallback aplicado com sucesso);
- evitar dumps longos no DEBUG; detalhar apenas identificadores úteis (ID prompt, status HTTP, nome ficheiro, step).

Notas adicionais para File Output + Structured Outputs (`json_schema`):

- quando `structured_outputs_mode=json_schema` e `strict=true`, o schema do manifest deve manter `required` alinhado com todas as chaves definidas em `properties` (incluindo chaves como `subfolder` quando existirem);
- o motor passa a emitir diagnóstico resumido do schema no DEBUG (`schema_name`, `strict`, contagem de `properties` e `required`), para reduzir tempo de troubleshooting de erros `invalid_json_schema`;
- antes do envio HTTP, o motor executa um preflight de JSON para detetar caracteres de controlo não escapados **e** escapes inválidos com backslash dentro de strings (causas comuns de `invalid_json`), bloqueando o envio e registando posição aproximada + escape sugerido no DEBUG (ex.: `\n`, `\r`, `\t`, `\u00XX`, e escapes após `\`: `\"`, `\\`, `\/`, `\b`, `\f`, `\n`, `\r`, `\t`, `\uXXXX`);
- além disso, valida estrutura mínima de JSON antes do HTTP (aspas/chaves/arrays e vírgula final inválida como `,}`/`,]`), para reduzir tentativas cegas quando há fusão de fragments (`Config extra` + `File Output`).
- ao editar o fragmento `text.format` de File Output, confirmar balanceamento de chaves no schema concatenado (`properties`/`items`/`required`) para evitar `fecho_sem_abertura` no preflight estrutural.
- durante construção do request, o payload final pode ser gravado em `C:\Temp\payload.json` para inspeção local antes de nova execução.

Recomendação operacional:

- limpar `Seguimento`/`DEBUG` antes de testes formais;
- validar sempre evidências mínimas por passo;
- manter logs curtos e acionáveis.

---

## 10. Segurança e compatibilidade retroativa

Regras essenciais:

- nunca commitar API keys reais;
- não expor segredos nos logs;
- preservar layout e cabeçalhos estruturais do workbook;
- mudanças em VBA devem manter fallback/default para templates antigos.

Resolução de `OPENAI_API_KEY` (ordem de precedência atual):

1. variável de ambiente `OPENAI_API_KEY` (recomendado);
2. fallback em `Config!B1` apenas por compatibilidade retroativa.

Notas operacionais:

- `Config!B1` pode manter uma diretiva como `(Environ("OPENAI_API_KEY"))` para documentar a origem da key;
- o parser também aceita variações equivalentes da diretiva (com/sem aspas internas ou espaços), além de `env:OPENAI_API_KEY` e `${OPENAI_API_KEY}`;
- quando o motor usa fallback literal em `Config!B1`, é emitido `ALERTA` no `DEBUG` para incentivar migração;
- quando não há key válida (nem ambiente nem fallback), é emitido `ERRO` no `DEBUG` e a execução é interrompida.

---

## 11. Guia rápido de operação

1. Confirmar `Config` (modelo, limites, opções de files/contexto) e, de preferência, a variável de ambiente `OPENAI_API_KEY`.
2. Confirmar IDs de catálogo e regras de `Next PROMPT`.
3. Preparar pipeline no PAINEL (`INPUT/OUTPUT folders` + limites).
4. Executar via INICIAR.
5. Auditar `Seguimento` e `DEBUG`.
6. Ajustar prompts/configuração e repetir.

---

> Este README é a referência de funcionamento do projeto. Guias de teste específicos (ex.: T3) devem viver como subseções operacionais ou documentação complementar, sem substituir a visão global do sistema.

## 12. EXECUTE Orders (Output Orders)

O PIPELINER suporta execução controlada de ordens pós-output, após resposta HTTP 2xx e sem erro.

### 12.1 Whitelist e sintaxe suportada (v1.3)

- Comando permitido: `LOAD_CSV`.
- Formatos aceites:
  - `EXECUTE: LOAD_CSV([ficheiro.csv])`
  - `<EXECUTE: LOAD_CSV([ficheiro.csv])>`
  - `EXECUTE: LOAD_CSV("ficheiro.csv")`
  - `EXECUTE: LOAD_CSV(ficheiro.csv)`
- Segurança:
  - apenas `basename.csv`;
  - rejeita `..`, `:`, `/`, `\`, `%`, `~`;
  - rejeita comandos fora da whitelist.

### 12.2 Fluxo LOAD_CSV

1. Parser ignora ordens dentro de blocos de código (```...```).
2. Resolve CSV automaticamente a partir de `downloadedFiles` e `OUTPUT Folder` (incluindo subpastas).
3. Faz pré-check técnico:
   - BOM UTF-8 (EF BB BF);
   - presença de CR/LF reais dentro de campos quoted;
   - `colsHint` pela linha de cabeçalho.
4. Cria worksheet nova após `PAINEL` (ou no fim, se `PAINEL` não existir), com nome baseado no prefixo do ID da coluna A do CSV.
5. Importa CSV por `QueryTables` (`;`, UTF-8), com fallback `OpenText`.
6. Verifica importação (linhas/colunas/header) e regista diagnóstico.

### 12.3 Logs

- DEBUG:
  - `OUTPUT_EXECUTE_FOUND`
  - `OUTPUT_EXECUTE_PARSED`
  - `OUTPUT_EXECUTE_UNKNOWN_CMD`
  - `OUTPUT_EXECUTE_INVALID_FILENAME`
  - `OUTPUT_EXECUTE_FILE_NOT_FOUND`
  - `OUTPUT_EXECUTE_CSV_PRECHECK`
  - `OUTPUT_EXECUTE_SHEET_CREATED`
  - `OUTPUT_EXECUTE_CSV_IMPORTED`
  - `OUTPUT_EXECUTE_VERIFIED`
  - `OUTPUT_EXECUTE_IMPORT_FAIL`
- Seguimento (`files_ops_log`): append com separador ` | ` e frase:
  - `CREATED AND LOADED Excel Sheet <sheetName> importing <nome_do_ficheiro.csv>, and verified.`

### 12.4 Patch proposto para prompt CatalogoFromTxt v1.3

Recomendação de alteração de prompt:

- Novo parâmetro: `AUTO_IMPORT_CSV_TO_SHEET: Sim|Não` (default: `Não`).
- Manifesto final curto:
  - `EXPORT_OK: true|false`
  - `FILE_NAME: <nome.csv>`
  - `DELIMITER: ;`
  - `ENCODING: UTF-8-BOM`
  - `COLS: <n>`
  - `ROWS: <n>`
- Se `AUTO_IMPORT_CSV_TO_SHEET=Sim`, emitir linha isolada:
  - `EXECUTE: LOAD_CSV([<nome_exacto_do_csv>])`
- Mitigação operacional:
  - anexar apenas 1 CSV final;
  - não gerar `.txt` auxiliares;
  - remover temporários no CI antes de concluir;
  - exportar CSV com `utf-8-sig` e normalizar quebras em células para literal `\n`.

### 12.5 DOCUMENTO_ORIENTADOR e FORMULARIO_DE_PROMPTS

Se estes artefactos existirem no workbook/documentação da equipa, aplicar:

- DOCUMENTO_ORIENTADOR:
  - secção “EXECUTE Orders” com whitelist, segurança, logs e troubleshooting.
- FORMULARIO_DE_PROMPTS:
  - campo/checkbox “Emitir EXECUTE após exportação?” com valores `OFF | LOAD_CSV_TO_NEW_SHEET`;
  - opção “Nome da folha: AUTO (prefixo do ID) / override”; 
  - nota de segurança para permitir apenas `basename.csv`.
