# Arquitetura

## Princípios

1. **A planilha mensal não muda.** A base abre a planilha de origem somente leitura e "puxa" os dados.
2. **Uma base, várias saídas.** A tabela `tblBase` (uma linha por aplicação por mês) alimenta o resumo, o PDF e, no futuro, um painel.
3. **Tudo por fórmula.** O resumo e as conferências recalculam ao trocar o mês de referência (`Config!B3`).
4. **Publicar só o que foi conferido.** ERRO bloqueia o PDF final; ALERTA exige justificativa.
5. **Rastreabilidade.** Cada importação gera uma linha na aba `Log`.

## Abas da base

| Aba | Função |
|---|---|
| LEIAME | Instruções de uso |
| Resumo | Relatório executivo (área de impressão A4) e botões |
| Conferencias | 12 verificações OK/ALERTA/ERRO e contadores `QtdErros` / `QtdAlertas` |
| Config | Mês de referência, limite de concentração, meta, tolerância, pasta do PDF |
| Cadastros | Instituições, modalidades, fontes e meses válidos |
| Indicadores | CDI e IPCA do mês e data-base original |
| Log | Histórico de importações |
| Base | Tabela `tblBase`: 52 colunas da planilha + 2 auxiliares |

## Colunas auxiliares da base

| Coluna | Fórmula | Uso |
|---|---|---|
| `Mov_Diferenca` | Saldo final − (anterior + aplicações − resgates − juros + rendimento − IR/IOF retido se saldo líquido) | Conferência de fechamento |
| `Saldo_Medio` | (saldo anterior + saldo final) ÷ 2 | Rentabilidade e efeito volume |

## Nomes definidos

`MesRef`, `MesAnt`, `LimiteConc`, `MetaCDI`, `TolContinuidade`, `PastaPDF`, `StatusRelatorio`, `QtdErros`, `QtdAlertas`, `col_*` (colunas da base), `cad_*` (cadastros) e `ind_*` (indicadores).

## Decomposição do rendimento

- Efeito volume = (saldo médio atual − saldo médio anterior) × taxa anterior
- Efeito taxa = variação total − efeito volume
- Taxa = rendimento bruto ÷ saldo médio simples

É uma decomposição gerencial: explica se a variação veio de "ter mais ou menos dinheiro aplicado" ou de "o dinheiro render mais ou menos".
