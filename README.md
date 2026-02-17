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

Quando o DEBUG mostrar `web_search=NAO_AUTO (ha anexos + flag config=FALSE)` em `M05_PAYLOAD_CHECK`, isso indica **gating local do PIPELINER**: o motor não auto-adiciona `tools:[{"type":"web_search"}]` quando já existem `input_file`/`input_image` no `input`.

A partir da revisão atual, existe a chave opcional `TOOLS_WEB_SEARCH_WITH_ATTACHMENTS` (folha `Config`, formato label->valor) para controlar este comportamento de forma retrocompatível:

- `FALSE` (default): mantém a regra histórica de não auto-adicionar web_search com anexos;
- `TRUE`: permite auto-adicionar web_search mesmo com anexos;
- se a chave estiver ausente: assume `FALSE`.

Checklist objetivo:

1. Confirmar `REQ_INPUT_JSON` com `has_input_file=SIM` e `file_id=file-...` quando o modo de transporte for `FILE_ID`.
2. Confirmar `M05_PAYLOAD_CHECK` com `has_input_file=SIM`, `has_file_id=SIM` e `web_search=NAO_AUTO (ha anexos + flag config=FALSE)` para o cenário de anexos (ou `ADICIONADO_AUTO ... flag config=TRUE`, se a flag estiver ativa).
3. Se precisar de pesquisa web + anexos por auto-injeção, definir `TOOLS_WEB_SEARCH_WITH_ATTACHMENTS=TRUE` em `Config`.

Para ContextKV, `CAPTURE_MISS` significa que o output não trouxe rótulos capturáveis esperados (`RESULTS_JSON`, `NEXT_PROMPT_ID`, `MEMORY_SHORT`, etc.). Para aumentar taxa de `CAPTURE_OK`, incluir no prompt instruções explícitas para devolver pelo menos:

- `RESULTS_JSON:` (linha com JSON ou bloco fenced);
- `NEXT_PROMPT_ID: STOP` (ou ID válido, se a pipeline usar AUTO).


Nota: `tools` continua como chave proibida em `Config extra` (é ignorada com alerta), para preservar a coerência com as colunas/lógica dedicadas.

SelfTests recomendados para este cenário:

- `SelfTest_WebSearchGating` (com/sem anexos; valida mensagem em `M05_PAYLOAD_CHECK`);
- `SelfTest_PayloadHasInputFileId` (valida `REQ_INPUT_JSON` e presença de `file_id`);
- `SelfTest_ContextKV_CaptureOkMiss` (2 outputs sintéticos: um capturável e outro livre);
- `SelfTest_InputsKvExtraction` (linhas `CHAVE: valor` e `CHAVE=valor`, com exclusão de `FILES:`).

Macros utilitárias para troubleshooting rápido de catálogo + Config extra:

- `TOOL_CreateCatalogTemplateSheet` (M15): cria uma nova folha de catálogo com layout compatível (headers A:K, bloco de 5 linhas, `Next PROMPT` e secções `Descrição textual/INPUTS/OUTPUTS`).
- `TOOL_RunConfigExtraSequentialDiagnostics` (M15): executa uma bateria sequencial de casos de `Config extra`, converte via parser oficial (`ConfigExtra_Converter`), injeta fragmento de File Output (`json_schema`) e valida a estrutura JSON final antes do HTTP.
- Resultado do diagnóstico fica em `CONFIG_EXTRA_TESTS` + linhas `INFO/ERRO` no `DEBUG` (`M15_CONFIG_EXTRA_DIAG`), com detalhe de causa (ex.: `fecho_sem_abertura`).

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
