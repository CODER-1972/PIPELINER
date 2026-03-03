# Análise de coerência — contrato parseável e gatekeeping CI/CSV

## Contexto
> **Nota de atualização:** este documento começou como análise de lacunas; várias recomendações já foram implementadas no VBA.
> Use a secção **Checklist de execução (estado atual)** como referência principal do que está concluído vs pendente.

Este documento avalia a proposta de robustez para cenários intermitentes de CI/`/mnt/data` (foco no `FLOW_TEMPLATE.csv`) face ao estado atual do PIPELINER.

## Conclusão executiva
Em geral, a proposta **faz sentido** e está alinhada com as dores reais já mitigadas parcialmente no código (guardrails de anexos, fallback CI, diagnósticos em DEBUG/DEBUG_DIAG, prechecks de CSV). O principal risco não é técnico, mas de **governança de contrato**: se for aplicada de forma rígida e global, pode criar falsos bloqueios em prompts que não exigem CI/CSV.

## O que já existe hoje (parcialmente alinhado)
- Bloqueio pré-API quando anexos declarados em `FILES:` não foram efetivamente preparados (`INPUTFILES_MISSING`).
- Registo de breadcrumbs de estágio e trilha de execução por passo (`STEP_STAGE`) + auditoria em `Seguimento`.
- Diagnóstico paralelo em `DEBUG_DIAG` com campos de CI/container, diretivas `EXECUTE`, causa provável e `suggested_fix`.
- Estratégia robusta de resolução CI em File Output com prioridade determinística de evidências (`container_file_citation` > marcadores textuais > fallback por listagem), incluindo mitigação de ambiguidades.
- Execução de `EXECUTE: LOAD_CSV` com whitelist, validação de nome, precheck CSV (BOM/CRLF), import com fallback e verificação estrutural mínima.
- Bundle de diagnóstico opcional suporta pasta e ZIP; continua em aberto padronizar o recorte explícito “PROVA_CI isolado” por passo.

## Lacunas face à proposta (estado após implementação atual)
1. **Estado tri-state persistido em célula dedicada (opcional)**
   - O estado já é decidido por passo (`OK/FAIL/BLOCKED`) e aplicado no gate, mas ainda não existe persistência canónica em célula única de run (ex.: folha de controlo).
2. **Diff “esperado vs encontrado” explícito com base no bloco PROVA_CI**
   - Implementado no contrato (`CONTRACT_PROVA_DIFF` + regra `R5_PROVA_EXPECTED_DIFF`) com decisão inequívoca quando falta ficheiro esperado no `PROVA_CI`.
3. **UX de remediação em UI**
   - Mensagens no DEBUG estão acionáveis, mas não existem ainda botões dedicados no Excel (ex.: “Abrir instruções”, “Reexecutar passo 4”).
4. **Evidência isolada por artefacto de contrato**
   - Implementado no bundle (`step<passo>_PROVA_CI.txt`), incluindo fallback explícito quando o bloco não existe.

## Pontos de potencial conflito/disfunção (se implementar sem guardrails)
- **Risco de regressão por rigidez global**: tornar os marcadores obrigatórios em todos os passos quebra retrocompatibilidade e contradiz o princípio de tolerância/compatibilidade do template.
- **Conflito semântico com fallback CI já existente**: exigir sempre `container_file_citation` pode invalidar cenários em que o motor recupera output legitimamente por marcador textual/listagem do container.
- **Falsos positivos quando CI não deve ser usado**: há lógica intencional de supressão de auto-add de CI com anexos sem intenção explícita; impor contrato de `/mnt/data` nesses passos seria incoerente.
- **Custos operacionais de logging**: verbose sem truncagem/limites pode degradar performance e poluir DEBUG, contrariando a diretriz de logs curtos e acionáveis.
- **Mudança estrutural de workbook**: criar nova folha obrigatória sem fallback pode colidir com a regra de evitar alterações estruturais não estritamente necessárias.

## Recomendação de desenho (para evitar conflitos)
1. **Contrato por perfil/step (opt-in)**
   - Ex.: `diagnostic_contract: ci_csv_v1` no `Config extra` (prompt) ou por pipeline; quando o motor detetar benefício deste modo, registar sugestão no `DEBUG` (coluna `Sugestao`).
2. **Gate tri-state com fallback seguro**
   - `OK/FAIL/BLOCKED`, mas só bloqueante para passos com contrato ativo; estado reportado no `DEBUG` (verdade técnica) e resumo funcional no `Seguimento`.
3. **Compatibilidade progressiva (sem nova folha)**
   - Não criar `PIPELINER_LOG`; consolidar o registo operacional no `DEBUG`, com campos estáveis por run/passo e detalhe técnico estruturado.
4. **Regras de consistência hierárquicas**
   - Distinção `hard fail` / `warn` / `info` explicitamente no `DEBUG` por severidade e parâmetro técnico.
5. **Bundle em três modos configuráveis**
   - `local_only`, `zip_only` ou `local_and_zip`, com precedência do Config extra do prompt sobre Config global; subpasta local também configurável por prompt (prioritário) ou Config.
6. **SelfTests dedicados**
   - cenários: marcador ausente, citation ausente com fallback válido, PROVA_CI sem CSV, CSV inválido e passo sem contrato (com DEBUG detalhado sem bloqueio por contrato).



## Contrato mínimo proposto para eventos DEBUG (tri-state)

### Objetivo
Padronizar **o mínimo obrigatório** que deve ser registado no `DEBUG` em cada passo, para que:
- a decisão (`OK/FAIL/BLOCKED`) seja auditável;
- a leitura seja compreensível por utilizadores não técnicos;
- o suporte consiga correlacionar rapidamente run, passo, regra e evidência.

### Formato legível (com campos obrigatórios entre `[]`)
Recomendação: incluir os campos obrigatórios no texto das células em linhas próprias, por exemplo:
- `[RunID: 20260303_101530_4821]`
- `[Passo: 4]`
- `[PromptID: AvalCap/04/Validacao/A]`
- `[Contrato: ci_csv_v1]`
- `[Estado: BLOCKED]`
- `[Regra: C2_MISSING_MARKER]`
- `[Severidade: ERRO]`

Isto pode aparecer no campo `Problema` e/ou `Sugestao`, mantendo linguagem simples abaixo dos metadados.

### Eventos/códigos mínimos por passo (detalhados)
1. `CONTRACT_EVAL_START` (INFO)
   - Quando: início da avaliação do contrato no passo.
   - Mensagem leiga: “Iniciada validação de consistência deste passo.”
   - Deve incluir: `[RunID] [Passo] [PromptID] [Contrato] [Estado: EM_ANALISE]`.

2. `CONTRACT_MARKERS_PARSED` (INFO ou ALERTA)
   - Quando: após parsing dos marcadores esperados.
   - Mensagem leiga: “Marcadores obrigatórios lidos com sucesso” ou “faltam marcadores esperados”.
   - Deve incluir: `[RunID] [Passo] [Regra: C1_PARSE]` + lista curta de marcadores encontrados/em falta.

3. `CONTRACT_RULE_RESULT` (INFO/ALERTA/ERRO)
   - Quando: resultado de cada regra de consistência aplicada.
   - Mensagem leiga: “Regra de consistência validada” ou “inconsistência detetada”.
   - Deve incluir: `[RunID] [Passo] [Regra: <codigo>] [EstadoParcial: PASS|WARN|FAIL]`.

4. `CONTRACT_STATE_DECISION` (INFO ou ERRO)
   - Quando: decisão final do tri-state do passo.
   - Mensagem leiga: “Passo aprovado para avançar” / “Passo bloqueado para evitar erro em cascata”.
   - Deve incluir: `[RunID] [Passo] [Estado: OK|FAIL|BLOCKED] [RegraFinal]`.

5. `CONTRACT_NEXT_ACTION` (INFO)
   - Quando: após decisão, para orientar ação humana.
   - Mensagem leiga: “Próxima ação recomendada”.
   - Deve incluir: `[RunID] [Passo]` + instrução objetiva (ex.: “Reexecutar passo 4 após confirmar CSV no PROVA_CI”).

### Campos mínimos obrigatórios (sempre presentes)
- `RunID`
- `Passo`
- `PromptID`
- `Contrato` (ou `SEM_CONTRATO`)
- `Estado` (`OK/FAIL/BLOCKED`)
- `Regra` (código curto)
- `Severidade` (`INFO/ALERTA/ERRO`)

### Linguagem recomendada para leigos nos campos do DEBUG
- **Funcionalidade**: “Validação de consistência do passo”, “Verificação de ficheiros de entrada”, etc.
- **Problema**: frase curta em português simples (“CSV esperado não foi comprovado no PROVA_CI”).
- **Sugestao**: ação concreta em 1–2 passos (“Anexar CSV no input_file e reexecutar passo 4”).
- Evitar jargão sem contexto; quando houver siglas, explicar em seguida.

## Explicação simples (para leigo informado) dos 6 pontos

### 1) Contrato por perfil/step (opt-in)
**Em linguagem simples:**
- Nem todos os passos da pipeline precisam do mesmo nível de rigor.
- Então, em vez de obrigar regras novas para tudo, ativa-se o “modo contrato” apenas onde faz sentido (por exemplo, no passo que depende do `FLOW_TEMPLATE.csv`).

**Exemplo prático:**
- Passo 4 tem `diagnostic_contract: ci_csv_v1` → aqui o motor exige marcadores como `PROVA_CI`.
- Passo 7 (resumo textual) não tem esse contrato → segue o fluxo normal, sem bloqueios extra.
- Se o motor detetar padrão recorrente de inconsistência, deve sugerir no `DEBUG` (coluna `Sugestao`) ativar `diagnostic_contract` nesse passo.

**Benefício:**
- Evita quebrar pipelines antigas que nunca foram desenhadas para esse contrato.

### 2) Gate tri-state com fallback seguro
**Em linguagem simples:**
- Cada passo termina com um estado claro:
  - `OK`: está coerente, pode avançar.
  - `FAIL`: erro confirmado.
  - `BLOCKED`: faltam provas/condições para confiar no passo.

**Regra recomendada:**
- Só bloquear automaticamente quando o contrato desse passo está ativo.
- Registar sempre o estado final do passo (`OK/FAIL/BLOCKED`) no `DEBUG` e manter no `Seguimento` apenas o resumo funcional.

**Exemplo prático:**
- Se o passo com contrato diz “CSV encontrado” mas não há evidência suficiente, fica `BLOCKED` e não avança para `LOAD_CSV`.

**Benefício:**
- Evita “falso sucesso” e também evita travar passos que não estavam sob esse contrato.

### 3) Compatibilidade progressiva
**Em linguagem simples:**
- Não criar nova folha de log operacional; a folha `DEBUG` passa a concentrar o registo técnico de verdade.
- A evolução deve ser incremental: manter compatibilidade com o layout atual e enriquecer o `DEBUG` com campos/colunas estáveis por run.

**Exemplo prático:**
- Workbook antigo continua a funcionar sem mudanças estruturais.
- O diagnóstico adicional é escrito no `DEBUG` (mesma folha), sem dependência de nova estrutura.
- Se faltar algum campo novo, a rotina degrada com defaults e não aborta execução.

**Benefício:**
- Introduz melhorias sem “rebentar” ambientes antigos.

### 4) Regras de consistência hierárquicas
**Em linguagem simples:**
- Nem toda anomalia merece “parar tudo”.
- Distinguir no `DEBUG`:
  - inconsistência crítica (hard fail),
  - alerta relevante (warn),
  - informação útil (info).

**Exemplo de hard fail (inequívoco):**
- No mesmo passo/contrato, output declara `EXPORT_OK_CSV=true`, mas o `PROVA_CI` não mostra o CSV esperado.

**Exemplo de warn:**
- Formato da prova está estranho, mas há outras evidências sólidas de que o ficheiro existe.

**Benefício:**
- Menos falsos bloqueios e melhor qualidade de diagnóstico.

### 5) Bundle com três modos
**Em linguagem simples:**
- Modo 1: `local_only` (guarda apenas pasta local estruturada por run).
- Modo 2: `zip_only` (gera apenas ZIP para suporte).
- Modo 3: `local_and_zip` (mantém pasta + ZIP).
- A pasta local pode usar subpasta configurável (`diagnostics_subfolder`) com precedência do prompt sobre Config global.

**Exemplo prático:**
- `diag_bundle_mode: local_only` → só pasta local (mais rápido).
- `diag_bundle_mode: zip_only` → só ZIP para ticket.
- `diag_bundle_mode: local_and_zip` → pasta + ZIP.
- `diagnostics_subfolder: NomeDaSubpasta` pode ser definido no prompt (prioritário) ou na Config.

**Benefício:**
- Mantém rastreabilidade local e facilita partilha quando necessário.

### 6) SelfTests dedicados
**Em linguagem simples:**
- Criar testes automáticos pequenos para os erros mais comuns, antes de chegarem ao utilizador.

**Conjunto mínimo recomendado:**
- marcador obrigatório ausente → deve `BLOCKED`.
- `container_file_citation` ausente, mas fallback válido → não deve falhar injustamente.
- `PROVA_CI` sem CSV quando CSV era obrigatório → deve bloquear com mensagem clara.
- CSV inválido (vazio/encoding/separador/colunas) → deve emitir erro/alerta acionável.

**Benefício:**
- Reduz regressões e acelera troubleshooting quando houver mudanças no VBA.





## Critérios de aceitação por cenário (DoD mínimo)
- **Cenário A — passo com contrato e marcador obrigatório ausente**
  - Resultado esperado: `BLOCKED` + DEBUG detalhado (eventos mínimos do contrato e regra acionada).
- **Cenário B — passo sem contrato**
  - Resultado esperado: **nunca bloquear por regra de contrato**.
  - Mesmo assim, registar DEBUG detalhado de observação (`SEM_CONTRATO`) para auditoria e comparação entre runs.
- **Cenário C — contrato com evidência consistente**
  - Resultado esperado: `OK` + registo de validação e próxima ação (seguir pipeline).
- **Cenário D — inconsistência crítica inequívoca**
  - Resultado esperado: `FAIL` ou `BLOCKED` (conforme regra), com sugestão acionável em linguagem simples.

## Sugestões práticas para reforçar o DEBUG como log operacional único
- **Adicionar colunas estáveis** (sem quebrar cabeçalhos atuais): `RunId`, `ContractMode`, `StepState`, `RuleId`, `EvidenceRef`, `DetailJsonCompact`.
- **Normalizar severidade**: mapear explicitamente `INFO/ALERTA/ERRO` para `info/warn/hard_fail` quando aplicável ao contrato.
- **Separar mensagem humana de detalhe técnico**: manter `Problema/Sugestao` curtos e colocar payload resumido em `DetailJsonCompact` com truncagem previsível.
- **Compact JSON budget configurável**: limitar `DetailJsonCompact` por linha (ex.: 1–2 KB) com parâmetro na folha `Config` (ex.: `DEBUG_DETAIL_JSON_MAX_CHARS`) e fallback interno seguro quando ausente.
- **Criar filtros prontos** no topo da folha para uso em suporte (`RunId`, `Prompt ID`, `Parametro`, `StepState`).
- **Garantir correlação com Seguimento**: escrever sempre `RunId` e `Passo` para cruzamento rápido entre auditoria funcional e diagnóstico técnico.

## Julgamento final
A proposta é **coerente e útil** para o objetivo de reduzir intermitências e melhorar depuração, desde que implementada com:
- ativação por contrato explícito,
- preservação da retrocompatibilidade,
- respeito pelo desenho atual de fallback CI,
- limites de verbosidade e sem exposição de dados sensíveis.

Sem esses cuidados, há risco real de transformar robustez em bloqueio excessivo e de introduzir falsos erros operacionais.


## Checklist de execução (estado atual)
- [x] Contrato por step opt-in (`diagnostic_contract: ci_csv_v1`) com sugestão automática no `DEBUG` quando há indícios de CSV/`LOAD_CSV` sem contrato ativo.
- [x] Gate tri-state com fallback seguro (`OK/FAIL/BLOCKED`) e bloqueio apenas para passos com contrato ativo.
- [x] Estado reportado no `DEBUG` por eventos canónicos (`CONTRACT_*`); `Seguimento` mantém resumo funcional do passo.
- [x] Compatibilidade progressiva sem nova folha (`DEBUG` como log operacional de verdade).
- [x] Regras de consistência hierárquicas no `DEBUG` (`INFO/ALERTA/ERRO`) com hard fail só em inconsistência inequívoca.
- [x] Bundle em três modos (`local_only`, `zip_only`, `local_and_zip`) com precedência prompt > Config e subpasta configurável.
- [x] Budget de detalhe (`DEBUG_DETAIL_JSON_MAX_CHARS`) com truncagem previsível para manter volume controlado.
- [x] SelfTests dedicados para cenários DoD mínimos (sem contrato, marcador ausente, inconsistência EXECUTE/FOUND, fallback válido sem citation, bloqueio sem prova equivalente).
- [x] Diff determinístico expected-vs-PROVA_CI com regra própria no contrato.
- [x] Artefacto dedicado `step<passo>_PROVA_CI.txt` no bundle de diagnóstico.
- [ ] Teste integrado em Excel host real (run completo com workbook de referência) para validar UX final de mensagens e tempos (pendente operacional fora deste ambiente).
