# Aprendizados técnicos

- **`TEXT()` depende do idioma do Excel.** `TEXT(A1;"mmm/yyyy")` exibe resultado errado em Excel em português. Alternativas independentes: `FIXED()` para números e `CHOOSE(MONTH())` para nomes de mês.
- **Nomes definidos não podem parecer endereços de célula** (ex.: `dR12`, `d2099`). Use prefixos como `col_` ou `_`.
- **Erro 9 ("subscrito fora do intervalo")** ao acessar abas geralmente indica módulo importado no projeto VBA errado: `ThisWorkbook` é sempre o arquivo onde o código está. A função `ArquivoCorreto()` detecta e explica.
- **`Shapes.AddShape` pode falhar (erro 1004)** em algumas configurações. `CriarBotoes` tenta a forma colorida e cai para `Buttons.Add`.
- **Codificação dos `.bas`:** o editor VBA lê Windows-1252 e o GitHub exibe UTF-8. Para evitar acentos corrompidos, os módulos deste repositório usam apenas ASCII.
- **Fundos com saldo informado líquido de IR/IOF** não fecham pela fórmula padrão de movimentação. Não é erro: exigem desconto do imposto retido na conciliação (coluna `Saldo_Liquido_de_IR`).
- **Queda de rendimento nem sempre é perda de eficiência.** Menos dias úteis e menor volume costumam explicar a maior parte; a decomposição volume × taxa ajuda a separar.
- **Concentração percentual pode subir com saldo menor**, quando os resgates se concentram em outras instituições.
- **openpyxl não grava valores calculados.** O arquivo gerado usa `fullCalcOnLoad` para o Excel recalcular ao abrir; a validação é feita com LibreOffice headless.
