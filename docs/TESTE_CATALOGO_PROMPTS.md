# Testes rápidos do PIPELINER + layout de catálogo (folha `TESTE`)

Este guia dá uma sequência de testes práticos e um bloco de catálogo pronto para colar no Excel.

## Testes que pode executar (checklist)

| ID | Objetivo | Preparação | Passos | Resultado esperado |
|---|---|---|---|---|
| T1 | Validar execução base (1 passo) | Pipeline no `PAINEL` com `TESTE/01/Entrada/A`; `Max Steps >= 1`; API key válida. | Clicar **INICIAR**. | 1 linha no `Seguimento`; `DEBUG` sem ERRO; pipeline termina em `STOP`. |
| T2 | Validar append de `INPUTS` em `RAW` | Em `Config`: `INPUTS_APPEND_MODE=RAW`; `AUTO_INJECT_INPUT_VARS=TRUE`. | Executar `TESTE/01/Entrada/A`. | Prompt final inclui bloco `### INPUTS_RESOLVIDOS` com todas as linhas de `INPUTS` (exceto vazias). |
| T3 | Validar filtro técnico em `SAFE` | Em `Config`: `INPUTS_APPEND_MODE=SAFE`. | Executar `TESTE/01/Entrada/A` com linha `FILES:` no `INPUTS`. | Bloco `INPUTS_RESOLVIDOS` **não** inclui linhas técnicas `FILES:/FICHEIROS:`. |
| T4 | Validar modo `OFF` | Em `Config`: `INPUTS_APPEND_MODE=OFF`. | Executar `TESTE/01/Entrada/A`. | Não aparece bloco `INPUTS_RESOLVIDOS` no prompt enviado. |
| T5 | Validar extração de variáveis (`:` e `=`) | Em `Config`: `AUTO_INJECT_INPUT_VARS=TRUE`. | Em `INPUTS` usar `URLS_ENTRADA: ...` e `MODO_DE_VERIFICACAO=Padrao`; executar. | `DEBUG` mostra `INPUTS_VARS` INFO; `Seguimento.captured_vars` guarda chaves normalizadas. |
| T6 | Validar conflito de chave | No `INPUTS`, repetir chave com valores diferentes (`MODO: A` e `MODO=B`). | Executar. | `DEBUG` com ALERTA de conflito e manutenção do 1.º valor. |
| T7 | Validar linha ignorada | No `INPUTS`, incluir linha sem `:` nem `=` (ex.: texto livre). | Executar. | `DEBUG` com ALERTA de linha ignorada para extração. |
| T8 | Validar persistência em `captured_vars_meta` | `AUTO_INJECT_INPUT_VARS=TRUE`. | Executar um passo. | `captured_vars_meta` contém `{"source":"inputs_extract","mode":"normalized_kv"}`. |

---

## Layout para copy-paste no catálogo `TESTE`

> Regras respeitadas: prefixo do ID igual ao nome da folha (`TESTE`), bloco de 5 linhas por prompt, `Next PROMPT` em coluna B, documentação `Descrição textual / INPUTS / OUTPUTS` em C-D.

### Tabela Markdown (referência visual)

| A (ID) | B | C | D | E (Modelo) | F (Modos) | G (Storage) | H (Config extra) | I | J | K |
|---|---|---|---|---|---|---|---|---|---|---|
| TESTE/01/Entrada/A | Entrada | Teste de INPUTS híbrido + extração KV | És um assistente de teste. Resume os INPUTS recebidos e confirma o modo de execução. | gpt-5.2 | Web search | TRUE | output_kind: file\nprocess_mode: metadata | Prompt de validação | Sem dependências externas | v1 |
|  | Next PROMPT: STOP | Descrição textual: | Prompt de smoke test para INPUTS. |  |  |  |  |  |  |  |
|  | Next PROMPT default: STOP | INPUTS: | URLS_ENTRADA: https://www.rtp.pt/noticias/mundo/europeus-apontam-falta-de-preparacao-para-lidar-com-alteracoes-climaticas-que-afetam-a-maioria-dos-cidadaos-da-ue_n1716225\nMODO_DE_VERIFICACAO=Padrao\nFILES: GUIA_DE_ESTILO.pdf (latest) (as pdf) |  |  |  |  |  |  |  |
|  | Next PROMPT allowed: STOP | OUTPUTS: | Texto curto com confirmação de leitura de INPUTS e fim em STOP. |  |  |  |  |  |  |  |
|  |  |  |  |  |  |  |  |  |  |  |

### Bloco TSV (recomendado para colar diretamente no Excel)

Copiar da linha abaixo para o Excel (começando na célula `A2` da folha `TESTE`):

```tsv
TESTE/01/Entrada/A	Entrada	Teste de INPUTS híbrido + extração KV	És um assistente de teste. Resume os INPUTS recebidos e confirma o modo de execução.	gpt-5.2	Web search	TRUE	output_kind: file
process_mode: metadata	Prompt de validação	Sem dependências externas	v1
	Next PROMPT: STOP	Descrição textual:	Prompt de smoke test para INPUTS.							
	Next PROMPT default: STOP	INPUTS:	URLS_ENTRADA: https://www.rtp.pt/noticias/mundo/europeus-apontam-falta-de-preparacao-para-lidar-com-alteracoes-climaticas-que-afetam-a-maioria-dos-cidadaos-da-ue_n1716225
MODO_DE_VERIFICACAO=Padrao
FILES: GUIA_DE_ESTILO.pdf (latest) (as pdf)							
	Next PROMPT allowed: STOP	OUTPUTS:	Texto curto com confirmação de leitura de INPUTS e fim em STOP.							
										
```

## Config mínima para estes testes

Adicionar/validar na folha `Config` (coluna A/B/C):

- `INPUTS_APPEND_MODE` = `RAW` (ou `SAFE`/`OFF` conforme o teste)
- `AUTO_INJECT_INPUT_VARS` = `TRUE`

Descrição sugerida (coluna C):

- `INPUTS_APPEND_MODE`: define se `INPUTS` é anexado ao prompt final e com que filtro.
- `AUTO_INJECT_INPUT_VARS`: extrai `CHAVE:valor` / `CHAVE=valor` de `INPUTS` para `captured_vars`.
