# Conferências de qualidade dos dados

| # | Conferência | Regra | Status se falhar |
|---|---|---|---|
| 1 | Mês de referência cadastrado | Mês existe em `Cadastros` | ERRO |
| 2 | Linhas do mês na base | Mais de zero linhas | ERRO |
| 3 | Linhas sem instituição ou modalidade | Zero | ERRO |
| 4 | Soma por fonte × total | Diferença de R$ 0,00 | ERRO |
| 5 | Instituições não cadastradas | Zero linhas | ERRO |
| 6 | Modalidades não cadastradas | Zero linhas | ERRO |
| 7 | Movimentação que não fecha | Zero linhas com diferença > R$ 0,01 | ERRO |
| 8 | Continuidade do saldo | Saldo inicial do mês = saldo final do mês anterior (até a tolerância) | ALERTA |
| 9 | Data-base original em dia útil | Segunda a sexta | ALERTA |
| 10 | Rendimento negativo | Zero aplicações | ALERTA |
| 11 | Mês anterior disponível | Existe na base | ALERTA |
| 12 | CDI do mês disponível | Preenchido e diferente de zero | ERRO |

**ERRO** bloqueia o PDF final (só permite a versão PRELIMINAR).
**ALERTA** não bloqueia, mas deve ser justificado no comentário da analista.

## Casos de teste incluídos

| Arquivo | Resultado esperado |
|---|---|
| `planilha_mensal_2025-04_data_fim_de_semana.xlsx` | 1 ALERTA (data-base em 26/04/2025, sábado) |
| `planilha_mensal_2025-05_com_erro.xlsx` | 1 ERRO (instituição "Banco Zeta" não cadastrada) |
