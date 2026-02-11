# src/vba/AGENTS.md — Regras técnicas de VBA (PIPELINER)

## 0) Propósito deste ficheiro
Estas regras existem para evitar regressões em VBA e tornar o comportamento:
- compatível (Excel/VBA),
- previsível (determinístico),
- robusto (erros e edge cases),
- auditável (logs úteis),
- seguro (sem segredos em logs/código).

Este ficheiro é “lei” para alterações em `src/vba/*.bas` (e eventuais `.cls/.frm`).

---

## 1) Premissas do host e compatibilidade
### 1.1 Host alvo
- Host principal: **Excel (Windows)**.
- Assumir que o utilizador cola/usa o código num `.xlsm` com macros activas.

### 1.2 32-bit vs 64-bit
- Se houver declarações de API/PtrSafe:
  - usar `#If VBA7 Then` e `LongPtr` quando necessário.
- Evitar tipos “ambíguos” que rebentam em 64-bit.

### 1.3 Referências e binding
Preferência:
- **Late binding** quando possível (portabilidade e menos dependências), sobretudo para:
  - `Scripting.Dictionary` (via `CreateObject("Scripting.Dictionary")`)
  - HTTP clients (WinHTTP/MSXML)
  - FileSystemObject (se usado, ponderar alternativas nativas)

Se usares early binding:
- documentar explicitamente a referência necessária (nome, versão) e justificar.

---

## 2) Regras não negociáveis (falha = P0)
### 2.1 `Option Explicit`
- Todos os módulos `.bas/.cls` têm de começar com:
  - `Option Explicit`
- Deve estar antes de qualquer procedimento.

### 2.2 Compilação e auto-contenção
- O código entregue/alterado tem de compilar sem erros.
- O módulo/feature tem de ser **auto-contido**:
  - Não chamar procedimentos inexistentes no repo.
  - Se criares uma função helper, inclui-a no commit e garante que está acessível.

### 2.3 Sem “Resume Next” global
- Proibido `On Error Resume Next` como modo global.
- Só permitido:
  - em blocos curtos, para “capability checks” (guardas),
  - seguido de validação do resultado,
  - e com limpeza de erro (`Err.Clear`) e reposição (`On Error GoTo 0`).

### 2.4 Objectos: `Set` obrigatório
- Sempre que a atribuição for um objecto, usar `Set`.
- Em colecções/dicionários que armazenem objectos:
  - usar helpers para encapsular a atribuição com `Set`.

---

## 3) Convenções de estrutura do código (módulos e responsabilidades)
### 3.1 Convenção de nomes de módulos
- Módulos seguem padrão `MNN_...` (ex.: `M07_Painel_Pipelines`).
- Não renomear `Attribute VB_Name` sem necessidade (impacta import/export e referências).

### 3.2 Separação por camadas
Evitar misturar:
- UI Excel (ler/escrever folhas, status bar) com lógica pura.
- Construção de payload com execução HTTP.
- Parsing de “Config extra” com validação de folhas.

Regra: funções “puras” (sem I/O) devem existir quando possível:
- mais fáceis de testar,
- menos regressões.

### 3.3 Interfaces internas (contratos)
Para cada subsistema, manter contratos estáveis:
- Leitura de catálogo por ID: entradas/saídas determinísticas.
- Parsing de Config extra: chaves proibidas e validações.
- FILES: resolução de caminhos e effective_mode.
- Logging: formato e severidades.

---

## 4) Estilo e legibilidade (consistência > preferências)
### 4.1 Indentação e formatação
- Indentar consistentemente (4 espaços recomendado).
- `With ... End With` sempre completo.
- Evitar linhas demasiado longas; usar `_` correctamente.

### 4.2 Tipos (evitar Variant por omissão)
- Declarar tipos sempre que possível.
- Preferir `Long` a `Integer` (Excel usa muito Long para rows/cols).
- Evitar `Variant` excepto quando:
  - interage com APIs que devolvem Variant,
  - precisa de arrays variant para Range.Value.

### 4.3 Strings e encoding
- Ser explícito quando precisa de:
  - escapes JSON,
  - normalização de CRLF/LF,
  - sanitização ASCII_SAFE para filenames.

---

## 5) Padrão oficial de tratamento de erros (entry points)
### 5.1 Regra base
Entry points públicos (macros de botão, orquestrações principais) devem usar:
- `On Error GoTo EH`
- Secção `EH:` com reporte mínimo:
  - nome da rotina,
  - `Err.Number` e `Err.Description`,
  - contexto essencial (p.ex. ID actual, pipeline, step),
  - (opcional) `Erl` se houver linhas numeradas.

### 5.2 “Try/Finally” em VBA
Sempre que alterares estado global do Excel:
- `Application.ScreenUpdating`
- `Application.EnableEvents`
- `Application.Calculation`
- `Application.StatusBar`

…deves:
1) guardar valores anteriores,
2) aplicar alterações,
3) garantir reposição em `Finally:` (mesmo em erro).

Exemplo de estrutura (padrão, não copiar sem adaptar):
- `On Error GoTo EH`
- `...`
- `GoTo Finally`
- `EH: ...`
- `Finally: ... restore ...`

### 5.3 Guardas com Resume Next (permitido, com disciplina)
Usar apenas para:
- testar se uma propriedade/método existe no objecto,
- ou quando uma chamada falha legitimamente em certos tipos.

Regras:
- `On Error Resume Next`
- executar 1–3 linhas
- capturar `Err.Number`
- `Err.Clear`
- `On Error GoTo 0`
- fallback determinístico

Nunca:
- engolir erros sem log quando impacta execução.
- deixar `Resume Next` activo.

---

## 6) Logging e auditoria (DEBUG vs Seguimento)
### 6.1 Objetivo dos logs
- Seguimento: log por passo (auditoria funcional).
- DEBUG: alertas, erros, INFO de diagnóstico e troubleshooting.

### 6.2 Regras de conteúdo (anti-lixo)
- Logs devem permitir diagnosticar sem abrir o VBA.
- Não registar:
  - API keys,
  - payload completo com segredos,
  - dados pessoais desnecessários.
- Truncar conteúdo longo (ex.: previews de JSON).

### 6.3 Estrutura recomendada (consistência)
Sempre que possível, incluir:
- timestamp,
- severidade (INFO/WARN/ERROR/PASS/FAIL),
- componente (M09_FILES, M05_API, M04_PARSER),
- contexto (pipeline, prompt_id, step),
- mensagem curta e accionável.

---

## 7) Excel Object Model — boas práticas (evitar bugs clássicos)
### 7.1 Qualificar tudo
Evitar:
- `ActiveWorkbook`, `ActiveSheet`, `Selection`
Preferir:
- `ThisWorkbook.Worksheets("...")`
- objectos passados como parâmetro.

### 7.2 Evitar Select/Activate
- Proibido usar `.Select`/`.Activate` para lógica.
- Só permitido para UX deliberada (ex.: focar Seguimento!A1 no arranque).

### 7.3 Leitura/escrita eficiente
- Para milhares de células:
  - ler Range para array Variant,
  - processar em memória,
  - escrever de volta em bloco.

### 7.4 Sem merges em catálogos
- Código deve assumir “sem merges” nas zonas de catálogo e Next PROMPT.
- Se detectar merges, logar alerta e falhar de forma clara.

---

## 8) Parser de “Config extra” (amigável → JSON)
### 8.1 Gramática suportada (resumo operacional)
- `chave: valor` por linha
- nesting com pontos: `a.b.c: valor`
- listas: `[a, b, c]`
- objectos: `{k: v}`
- `input:` com linhas indentadas `role:` e `content:`

### 8.2 Validação e alertas
- Linhas inválidas → ignorar + alerta em DEBUG.
- Chaves proibidas → ignorar + alerta em DEBUG:
  - model, temperature, max_output_tokens, store, tools
- “conversation” vs “previous_response_id”:
  - não permitir simultâneo; escolher comportamento determinístico e logar.

### 8.3 Robustez JSON
- Nunca construir JSON “a olho” sem escaping.
- Usar utilitários de escape e builders consistentes.
- Se existir “preview do payload” em DEBUG:
  - truncar,
  - mascarar segredos,
  - garantir CRLF escapado.

---

## 9) Construção de pedidos à API (robustez)
### 9.1 Responsabilidades mínimas
- Construir request determinístico a partir de:
  - Config global (defaults),
  - overrides do catálogo (modelo, tools, store),
  - Config extra convertido para JSON,
  - anexos (FILES),
  - encadeamento (previous_response_id).

### 9.2 HTTP: timeouts, status e retries
- Definir timeouts razoáveis (connect/send/receive).
- Logar sempre:
  - endpoint,
  - HTTP status,
  - erro de transporte (se HTTP 0),
  - snippet curto da resposta em erro (sem segredos).

### 9.3 Gestão de erros comuns (mínimo esperado)
- 401/403: credenciais/permissões → mensagem clara.
- 413: payload demasiado grande → recomendar PDF vs text_embed, reduzir limites.
- 415: multipart inválido → sugerir ASCII_SAFE e alternar engine/perfil.
- 429/5xx: sugerir retry controlado (se existir política), com logs.

---

## 10) FILES: resolução, upload e “effective_mode”
### 10.1 Segurança de caminhos
- Resolver sempre paths dentro do INPUT Folder.
- Proibir path traversal (`..\`, `:` e afins) — usar apenas nome do ficheiro quando necessário.

### 10.2 Resolução determinística
- wildcard `*`:
  - resolver candidatos na pasta,
  - se `latest`, escolher o mais recente (critério claro),
  - se ambíguo, logar AMBIGUOUS e aplicar regra/política.

### 10.3 Transport modes
- `FILE_ID`:
  - upload /v1/files quando necessário,
  - reutilização via cache (hash + modo),
  - validação do file_id (GET /v1/files/<id>) quando aplicável.
- `INLINE_BASE64`:
  - respeitar limite de MB,
  - falhar com mensagem accionável se exceder.

### 10.4 Office “as is” e compatibilidade com Responses
- DOCX/PPTX/XLSX como `input_file` pode ser incompatível.
- Implementar `effective_mode`:
  - por defeito converter para PDF (AUTO_AS_PDF),
  - fallback para text_embed,
  - ou ERROR (parar) conforme config.

### 10.5 text_embed: limites e overflow
- Respeitar `FILES_TEXT_EMBED_MAX_CHARS`.
- Em overflow:
  - aplicar acção definida: ALERT_ONLY/TRUNCATE/RETRY_AS_PDF/STOP
- Logar claramente:
  - tamanho extraído,
  - limite,
  - acção tomada.

---

## 11) File Output (download/gravação)
- Validar OUTPUT Folder:
  - existe,
  - permissões,
  - caminho simples (evitar falhas de sync/rede).
- Em colisão de nomes:
  - usar estratégia previsível (suffix timestamp) se configurado.
- Registar em FILES_MANAGEMENT:
  - path completo,
  - tipo/extensão,
  - hash/tamanho,
  - response_id (se disponível),
  - status.

---

## 12) Idempotência (quando aplicável)
Se o código:
- cria ficheiros,
- cria folhas,
- escreve estruturas no DEBUG/Seguimento,
- cria caches (_pdf_cache),
- cria registos FILES_MANAGEMENT,

…deve ser idempotente:
- não duplicar estruturas em execuções repetidas,
- apagar/recriar apenas o que o próprio código criou (por prefixo/tag),
- manter logs claros.

---

## 13) SelfTests (obrigatório manter e actualizar quando tocar em subsistemas)
- SelfTest_RunAll deve:
  - ser rápido,
  - ser idempotente,
  - escrever PASS/FAIL no DEBUG,
  - validar sanitização, multipart e engines (mínimo).
- Se mexeres em:
  - sanitização de filename → adicionar teste específico,
  - upload/profile/engine → adicionar teste e logs,
  - parser Config extra → adicionar casos de erro/edge cases.

---

## 14) Checklist pré-entrega (antes de commit/PR)
1) Compila com Option Explicit.
2) Não há referências a membros inexistentes no host.
3) Não há object assignments sem Set.
4) Não há On Error Resume Next fora de guardas curtas.
5) Logs:
   - DEBUG tem alertas úteis,
   - Seguimento mantém auditoria por passo,
   - sem segredos.
6) FILES:
   - effective_mode não quebra compatibilidade,
   - overflow controlado.
7) Se aplicável:
   - SelfTest_RunAll passa,
   - execução curta (MaxSteps baixo) funciona.

---

## 15) Anti-padrões (proibidos ou altamente desencorajados)
- Uso generalizado de `Select`/`Activate` para lógica.
- “Hardcode” de paths pessoais ou dependência de OneDrive/UNC sem fallback.
- Guardar API key em texto no repo.
- Alterar nomes de folhas/cabeçalhos sem migração coordenada.
- Construir JSON por concatenação sem escape.
- Silenciar erros sem logging quando afectam execução.

---

## 16) Nota final para agentes
Se uma regra for impossível de cumprir:
- explica no PR (curto e objectivo),
- propõe fallback,
- e adiciona logging/guardas para reduzir risco.
