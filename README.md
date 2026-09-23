# Gerador de Relatório Executivo de Carteira de Renda Fixa

Automação em **Excel + VBA** que transforma uma planilha mensal de apuração de investimentos em um relatório executivo em PDF, com base histórica, conferências automáticas de qualidade dos dados e rastreabilidade de cada importação.

> **Todos os dados deste repositório são fictícios.** Nomes de instituições, valores, taxas e indicadores de mercado foram gerados por script (`python/gerar_dados_ficticios.py`) apenas para demonstração.

## O problema

Relatórios mensais de carteira costumam depender de trabalho manual: copiar dados, refazer somas, conferir saldos e montar gráficos. Isso consome tempo, cria dependência de poucas pessoas e abre espaço para erros silenciosos (saldo que não fecha, instituição nova sem cadastro, data-base em fim de semana).

## A solução

```
Planilha mensal (não é alterada)
        │  abertura somente leitura
        ▼
[Importar mês] ── valida 52 cabeçalhos ── normaliza data-base ── grava log
        │
        ▼
Base histórica (tblBase: uma linha por aplicação por mês)
        │  fórmulas SUMIFS / SUMPRODUCT filtradas pelo mês de referência
        ▼
Resumo executivo  +  Conferências (OK / ALERTA / ERRO)
        │
        ▼
[Gerar PDF] ── ERRO bloqueia a versão final (permite PRELIMINAR)
```

**Principais recursos**

- Importação validada: aborta se o layout da planilha mudar.
- Base histórica única, reaproveitável por outras ferramentas (ex.: Power BI).
- 12 conferências automáticas, com status OK, ALERTA ou ERRO.
- Log de importação: quando, quem, qual arquivo, quantas linhas e observações.
- Resumo com resultado do mês, meta em % do CDI, distribuição por fonte, concentração por instituição, modalidades e pontos de atenção gerados por fórmula.
- Decomposição da variação do rendimento em **efeito volume** e **efeito taxa**.
- Fórmulas independentes do idioma do Excel (sem `TEXT()` com códigos de formato) e sem funções de matriz dinâmica.

## Estrutura

| Pasta | Conteúdo |
|---|---|
| `excel/` | `base_relatorio_exemplo.xlsx`: base com jan–mar/2025 pré-carregados |
| `exemplos/` | Planilhas mensais fictícias para testar a importação (inclui casos de ALERTA e ERRO) |
| `vba/` | Módulos `modImportacao`, `modPDF` e `modUtil` |
| `python/` | Geração dos dados fictícios, construção da base e conferência independente |
| `docs/` | Arquitetura, regras das conferências, aprendizados técnicos e roteiro |

## Como testar (10 minutos)

1. Baixe o repositório (**Code → Download ZIP**) e descompacte.
2. Abra `excel/base_relatorio_exemplo.xlsx` e salve como **Pasta de Trabalho Habilitada para Macro (.xlsm)**.
3. Pressione `Alt+F11`, **selecione o arquivo da base no painel da esquerda** e use *Arquivo → Importar arquivo* para os três `.bas` da pasta `vba/`.
4. Feche o editor, pressione `Alt+F8` e execute `CriarBotoes`.
5. Clique em **Importar mês** e escolha `exemplos/planilha_mensal_2025-04_data_fim_de_semana.xlsx` → deve surgir **1 ALERTA**.
6. Importe `exemplos/planilha_mensal_2025-05_com_erro.xlsx` → deve surgir **1 ERRO** (instituição não cadastrada).
7. Clique em **Gerar PDF** e observe o bloqueio da versão final.

## Conferência independente (opcional, Python)

```bash
pip install -r python/requirements.txt
python python/analise_volume_taxa.py
```

## Limitações conhecidas

- Dia útil considera apenas fins de semana (feriados não são tratados).
- Instituições e modalidades são listas fixas na aba `Cadastros`.
- Nesta versão o PDF contém as abas `Resumo` e `Conferencias`; as páginas com layout de relatório estão no roteiro (`docs/roteiro.md`).

## Aviso

Projeto de demonstração. Qualquer uso com dados reais exige validação humana e das áreas responsáveis antes de decisões ou divulgação formal.

## Licença

A definir. Enquanto não houver licença, todos os direitos são reservados.
