# Testes internos + prompts de teste (ContextKV e passagem de variáveis)

Este guia complementa o README com um plano prático para validar:

1. captura de variáveis no output (`captured_vars`),
2. injeção no passo seguinte (`{{VAR:...}}`, `VARS:`),
3. referência a outputs anteriores (`{@OUTPUT: ...}`),
4. encadeamento dinâmico com `NEXT_PROMPT_ID`.

---

## 1) Testes internos (macros VBA)

Executar no Editor VBA (F5), pela ordem:

1. `SelfTest_RunAll`
2. `SelfTest_RunAll_ContextKV`

Resultado esperado:
- Linhas `SELFTEST_*` no `DEBUG` com `PASS`/`INFO`.
- Sem `ERRO` nos testes de `Parse_RESULTS_JSON`, `Placeholder_Replace` e `OutputRefFormat`.

> Nota: estes testes são idempotentes e não exigem alteração estrutural no workbook.

---

## 2) Pré-requisitos no workbook

- Criar (ou usar) uma folha de catálogo chamada `CTXTEST`.
- No `PAINEL`, configurar uma pipeline com os IDs abaixo (ordem sequencial).
- Em `Config`, garantir:
  - `CONTEXT_KV_ENABLED = TRUE` (se existir a chave),
  - `CONTEXT_KV_STRICT = FALSE` para diagnóstico inicial.

Pipeline sugerida:
- `CTXTEST/01/GerarDados/A`
- `CTXTEST/02/ConsumirVars/A`
- `CTXTEST/03/UsarOutputRef/A`
- `CTXTEST/04/DecisaoAuto/A`
- `CTXTEST/05/Fim/A`

---

## 3) Blocos de prompt para copy-paste no layout do catálogo

> Cada bloco segue o layout padrão: linha principal + `Next PROMPT` + `default` + `allowed`.

### Bloco 1 — gera variáveis estruturadas

**Linha principal (A..K)**

- **A (ID):** `CTXTEST/01/GerarDados/A`
- **B (Nome curto):** `GerarDados`
- **C (Nome descritivo):** `Produz variáveis para injeção`
- **D (Texto prompt):**

```text
Devolve exatamente no formato abaixo (sem texto adicional antes/depois):

A) RESULTS_JSON
```json
{"cliente":"ACME","ano":2026,"indicadores":["receita","margem"]}
```

B) MEMORY_SHORT:
Resumo curto do cliente ACME.

C) NEXT_PROMPT_ID: CTXTEST/02/ConsumirVars/A
```

- **E (Modelo):** *(vazio ou override opcional)*
- **F (Modos):** `None`
- **G (Storage):** `false`
- **H (Config extra):** *(vazio)*
- **I/J/K:** documentação livre

**Linhas seguintes (coluna B):**

```text
Next PROMPT: CTXTEST/02/ConsumirVars/A
Next PROMPT default: CTXTEST/02/ConsumirVars/A
Next PROMPT allowed: CTXTEST/02/ConsumirVars/A; STOP
```

---

### Bloco 2 — consome `{{VAR:...}}` + `VARS:`

**Linha principal (A..K)**

- **A (ID):** `CTXTEST/02/ConsumirVars/A`
- **D (Texto prompt):**

```text
VARS: RESULTS_JSON, MEMORY_SHORT

Usa os dados injetados para escrever:
1) nome do cliente,
2) ano,
3) lista de indicadores.

Mantém no fim a linha:
NEXT_PROMPT_ID: CTXTEST/03/UsarOutputRef/A

RESULTS_JSON injetado:
{{VAR:RESULTS_JSON}}

MEMORY_SHORT injetado:
{{VAR:MEMORY_SHORT}}
```

- **F (Modos):** `None`
- **G (Storage):** `false`

**Linhas seguintes (coluna B):**

```text
Next PROMPT: CTXTEST/03/UsarOutputRef/A
Next PROMPT default: CTXTEST/03/UsarOutputRef/A
Next PROMPT allowed: CTXTEST/03/UsarOutputRef/A; STOP
```

---

### Bloco 3 — usa `{@OUTPUT: "Prompt Anterior"}`

**Linha principal (A..K)**

- **A (ID):** `CTXTEST/03/UsarOutputRef/A`
- **D (Texto prompt):**

```text
Resume em 3 bullets o output anterior:
{@OUTPUT: "Prompt Anterior"}

No final devolve:
NEXT_PROMPT_ID: CTXTEST/04/DecisaoAuto/A
```

- **F (Modos):** `None`
- **G (Storage):** `false`

**Linhas seguintes (coluna B):**

```text
Next PROMPT: CTXTEST/04/DecisaoAuto/A
Next PROMPT default: CTXTEST/04/DecisaoAuto/A
Next PROMPT allowed: CTXTEST/04/DecisaoAuto/A; STOP
```

---

### Bloco 4 — valida AUTO + decisão explícita

**Linha principal (A..K)**

- **A (ID):** `CTXTEST/04/DecisaoAuto/A`
- **D (Texto prompt):**

```text
Responde com:
DECISION: continuar
RATIONALE: validação de AUTO com fallback controlado.
NEXT_PROMPT_ID: CTXTEST/05/Fim/A
```

- **F (Modos):** `None`
- **G (Storage):** `false`

**Linhas seguintes (coluna B):**

```text
Next PROMPT: AUTO
Next PROMPT default: CTXTEST/05/Fim/A
Next PROMPT allowed: CTXTEST/05/Fim/A; STOP
```

---

### Bloco 5 — término limpo

**Linha principal (A..K)**

- **A (ID):** `CTXTEST/05/Fim/A`
- **D (Texto prompt):**

```text
Confirma fim da pipeline e escreve apenas:
NEXT_PROMPT_ID: STOP
```

- **F (Modos):** `None`
- **G (Storage):** `false`

**Linhas seguintes (coluna B):**

```text
Next PROMPT: STOP
Next PROMPT default: STOP
Next PROMPT allowed: STOP
```

---

## 4) Checklist de validação (evidências mínimas)

Após correr a pipeline, validar:

1. **Seguimento**
   - coluna `captured_vars` preenchida no passo 1,
   - coluna `injected_vars` preenchida no passo 2,
   - `next` resolvido corretamente até `STOP`.

2. **DEBUG**
   - eventos `CAPTURE_OK`, `INJECT_OK`, `PLACEHOLDER_REPLACED`, `OUTPUT_REF_FOUND`/`INJECT_OK`,
   - ausência de `INJECT_MISS` (ou, se existir, com detalhe acionável).

3. **Comportamento de fallback AUTO**
   - no passo 4, se o parser não encontrar `NEXT_PROMPT_ID`, deve usar `Next PROMPT default`.

---

## 5) Testes negativos rápidos (opcional)

- Remover `NEXT_PROMPT_ID` do bloco 4 e confirmar fallback para `default`.
- Escrever `{{VAR:NAO_EXISTE}}` no bloco 2 e confirmar:
  - `ALERTA` (`strict=false`) ou
  - `ERRO` (`strict=true`).
- Alterar `allowed` do bloco 4 para não incluir o destino e confirmar alerta de validação.
