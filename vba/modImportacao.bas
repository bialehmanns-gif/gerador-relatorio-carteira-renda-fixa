Attribute VB_Name = "modImportacao"
Option Explicit

' =====================================================================
' modImportacao - importa a aba "Tabela" de uma planilha mensal para a
' tabela historica tblBase (aba Base).
'
' - A planilha de origem e aberta SOMENTE LEITURA e nunca e alterada.
' - Aborta se os 52 cabecalhos nao forem identicos aos da base.
' - Normaliza a data-base para o ultimo dia do mes (guarda a original).
' - Se o mes ja existir na base, as linhas sao substituidas.
' - Grava CDI/IPCA do mes, cadastra o mes e registra log.
' - NAO salva a base: revise e salve manualmente.
'
' Obs.: textos sem acentos de proposito (compatibilidade de codificacao
' entre o editor VBA e o GitHub).
' =====================================================================

Private Const ABA_ORIGEM As String = "Tabela"
Private Const QTD_COLUNAS As Long = 52
Private Const COL_CDI As Long = 30
Private Const COL_IPCA As Long = 33

Public Sub ImportarMes()
    Dim caminho As String
    Dim wbOrig As Workbook, wsOrig As Worksheet
    Dim lo As ListObject
    Dim dados As Variant
    Dim ultLinha As Long, n As Long, i As Long, j As Long, c As Long
    Dim dataOriginal As Date, dataNorm As Date
    Dim cdi As Variant, ipca As Variant
    Dim formulasAux() As String, qtdAux As Long
    Dim primeiraNova As Long, removidas As Long
    Dim celulasErro As Long, datasDivergentes As Long
    Dim msgCab As String, obs As String
    Dim calcAnterior As XlCalculation
    Dim erros As Long, alertas As Long

    If Not ArquivoCorreto() Then Exit Sub

    caminho = EscolherArquivo()
    If caminho = "" Then Exit Sub

    Set lo = ThisWorkbook.Worksheets("Base").ListObjects("tblBase")
    If lo.DataBodyRange Is Nothing Then
        MsgBox "A tabela tblBase esta vazia. Mantenha ao menos o historico de exemplo " & _
               "(ou uma linha) para preservar as formulas auxiliares.", vbCritical
        Exit Sub
    End If

    calcAnterior = Application.Calculation
    On Error GoTo Falha
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wbOrig = Workbooks.Open(Filename:=caminho, ReadOnly:=True, UpdateLinks:=0)
    Set wsOrig = wbOrig.Worksheets(ABA_ORIGEM)

    ' 1) Validacao do layout
    msgCab = ValidarCabecalhos(wsOrig, lo)
    If msgCab <> "" Then
        wbOrig.Close SaveChanges:=False
        Restaurar calcAnterior
        MsgBox "Importacao cancelada: o layout da planilha e diferente do esperado." & _
               vbCrLf & vbCrLf & msgCab, vbCritical
        Exit Sub
    End If

    ' 2) Leitura dos dados
    ultLinha = wsOrig.Cells(wsOrig.Rows.Count, 1).End(xlUp).Row
    n = ultLinha - 1
    If n < 1 Then
        wbOrig.Close SaveChanges:=False
        Restaurar calcAnterior
        MsgBox "A aba '" & ABA_ORIGEM & "' nao tem linhas de dados.", vbExclamation
        Exit Sub
    End If
    dados = wsOrig.Range(wsOrig.Cells(2, 1), wsOrig.Cells(ultLinha, QTD_COLUNAS)).Value

    ' 3) Data-base
    If Not IsDate(dados(1, 1)) Then
        wbOrig.Close SaveChanges:=False
        Restaurar calcAnterior
        MsgBox "A data-base (coluna A, linha 2) nao e uma data valida.", vbCritical
        Exit Sub
    End If
    dataOriginal = CDate(dados(1, 1))
    dataNorm = DateSerial(Year(dataOriginal), Month(dataOriginal) + 1, 0)
    cdi = dados(1, COL_CDI)
    ipca = dados(1, COL_IPCA)
    If IsError(cdi) Then cdi = Empty
    If IsError(ipca) Then ipca = Empty

    For i = 1 To n
        For j = 1 To QTD_COLUNAS
            If IsError(dados(i, j)) Then
                dados(i, j) = Empty
                celulasErro = celulasErro + 1
            End If
        Next j
        If IsDate(dados(i, 1)) Then
            If Year(dados(i, 1)) <> Year(dataNorm) Or Month(dados(i, 1)) <> Month(dataNorm) Then
                datasDivergentes = datasDivergentes + 1
            End If
        End If
        dados(i, 1) = dataNorm
    Next i

    wbOrig.Close SaveChanges:=False
    Set wbOrig = Nothing

    ' 4) Guarda as formulas das colunas auxiliares (53 em diante)
    qtdAux = lo.ListColumns.Count - QTD_COLUNAS
    If qtdAux > 0 Then
        ReDim formulasAux(1 To qtdAux)
        For c = 1 To qtdAux
            formulasAux(c) = lo.ListColumns(QTD_COLUNAS + c).DataBodyRange.Cells(1, 1).FormulaR1C1
        Next c
    End If

    ' 5) Remove o mes, se ja existir
    removidas = RemoverMes(lo, dataNorm)

    ' 6) Acrescenta as novas linhas
    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add
        primeiraNova = lo.DataBodyRange.Row
        If n > 1 Then lo.Resize lo.Range.Resize(lo.Range.Rows.Count + n - 1)
    Else
        primeiraNova = lo.DataBodyRange.Row + lo.ListRows.Count
        lo.Resize lo.Range.Resize(lo.Range.Rows.Count + n)
    End If

    With lo.Parent
        If primeiraNova > lo.DataBodyRange.Row Then
            lo.DataBodyRange.Rows(1).Copy
            .Range(.Cells(primeiraNova, lo.Range.Column), _
                   .Cells(primeiraNova + n - 1, lo.Range.Column + lo.ListColumns.Count - 1)).PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
        End If
        .Range(.Cells(primeiraNova, lo.Range.Column), _
               .Cells(primeiraNova + n - 1, lo.Range.Column + QTD_COLUNAS - 1)).Value = dados
        For c = 1 To qtdAux
            .Range(.Cells(primeiraNova, lo.Range.Column + QTD_COLUNAS + c - 1), _
                   .Cells(primeiraNova + n - 1, lo.Range.Column + QTD_COLUNAS + c - 1)).FormulaR1C1 = formulasAux(c)
        Next c
    End With

    ' 7) Cadastros, indicadores, mes de referencia e log
    CadastrarMes dataNorm
    GravarIndicadores dataNorm, cdi, ipca, dataOriginal
    ThisWorkbook.Names("MesRef").RefersToRange.Value = dataNorm

    obs = ""
    If removidas > 0 Then obs = obs & "Substituiu " & removidas & " linha(s) do mes. "
    If celulasErro > 0 Then obs = obs & celulasErro & " celula(s) com erro limpas. "
    If datasDivergentes > 0 Then obs = obs & datasDivergentes & " linha(s) com data-base de outro mes. "
    If CLng(dataOriginal) <> CLng(dataNorm) Then obs = obs & "Data-base original normalizada para o fim do mes. "
    If Weekday(dataOriginal, vbMonday) > 5 Then obs = obs & "Data-base original em fim de semana. "
    RegistrarLog NomeDoArquivo(caminho), dataNorm, n, dataOriginal, Trim(obs)

    Restaurar calcAnterior
    Application.Calculate
    erros = CLng(ThisWorkbook.Names("QtdErros").RefersToRange.Value)
    alertas = CLng(ThisWorkbook.Names("QtdAlertas").RefersToRange.Value)
    ThisWorkbook.Worksheets("Conferencias").Activate

    MsgBox "Importacao concluida: " & Format(dataNorm, "mm/yyyy") & vbCrLf & _
           n & " linha(s) importada(s)." & vbCrLf & vbCrLf & _
           "Conferencias: " & erros & " ERRO(s) e " & alertas & " ALERTA(s)." & vbCrLf & vbCrLf & _
           "A base NAO foi salva. Revise e salve manualmente.", _
           IIf(erros > 0, vbExclamation, vbInformation)
    Exit Sub

Falha:
    Dim msg As String, numErro As Long
    numErro = Err.Number
    msg = "Erro " & numErro & ": " & Err.Description
    On Error Resume Next
    If Not wbOrig Is Nothing Then wbOrig.Close SaveChanges:=False
    Application.CutCopyMode = False
    Restaurar calcAnterior
    If numErro = 9 Then
        msg = msg & vbCrLf & vbCrLf & "Verifique se a planilha tem a aba '" & ABA_ORIGEM & "'."
    End If
    MsgBox "A importacao foi interrompida." & vbCrLf & msg, vbCritical
End Sub

Private Function EscolherArquivo() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Escolha a planilha mensal a importar"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Planilhas Excel", "*.xlsx;*.xlsm;*.xls"
        If .Show = -1 Then EscolherArquivo = .SelectedItems(1)
    End With
End Function

Private Function ValidarCabecalhos(wsOrig As Worksheet, lo As ListObject) As String
    Dim j As Long, qtd As Long, esperado As String, recebido As String, txt As String
    For j = 1 To QTD_COLUNAS
        esperado = Trim(CStr(lo.HeaderRowRange.Cells(1, j).Value))
        recebido = Trim(CStr(wsOrig.Cells(1, j).Value))
        If StrComp(esperado, recebido, vbTextCompare) <> 0 Then
            qtd = qtd + 1
            If qtd <= 5 Then
                txt = txt & "Coluna " & j & ": esperado '" & esperado & "', encontrado '" & recebido & "'" & vbCrLf
            End If
        End If
    Next j
    If Trim(CStr(wsOrig.Cells(1, QTD_COLUNAS + 1).Value)) <> "" Then
        qtd = qtd + 1
        txt = txt & "Ha coluna extra apos a coluna " & QTD_COLUNAS & "." & vbCrLf
    End If
    If qtd > 5 Then txt = txt & "... e mais " & (qtd - 5) & " divergencia(s)."
    ValidarCabecalhos = txt
End Function

Private Function RemoverMes(lo As ListObject, dataNorm As Date) As Long
    Dim i As Long, v As Variant
    If lo.DataBodyRange Is Nothing Then Exit Function
    For i = lo.ListRows.Count To 1 Step -1
        v = lo.DataBodyRange.Cells(i, 1).Value
        If IsDate(v) Then
            If CLng(CDate(v)) = CLng(dataNorm) Then
                lo.ListRows(i).Delete
                RemoverMes = RemoverMes + 1
            End If
        End If
    Next i
End Function

Private Sub CadastrarMes(dataNorm As Date)
    Dim lo As ListObject, cel As Range
    Set lo = ThisWorkbook.Worksheets("Cadastros").ListObjects("tblMeses")
    If Not lo.DataBodyRange Is Nothing Then
        For Each cel In lo.ListColumns(1).DataBodyRange.Cells
            If IsDate(cel.Value) Then
                If CLng(CDate(cel.Value)) = CLng(dataNorm) Then Exit Sub
            End If
        Next cel
    End If
    lo.ListRows.Add.Range.Cells(1, 1).Value = dataNorm
End Sub

Private Sub GravarIndicadores(dataNorm As Date, cdi As Variant, ipca As Variant, dataOriginal As Date)
    Dim lo As ListObject, lr As ListRow, i As Long
    Set lo = ThisWorkbook.Worksheets("Indicadores").ListObjects("tblIndicadores")
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            If IsDate(lo.DataBodyRange.Cells(i, 1).Value) Then
                If CLng(CDate(lo.DataBodyRange.Cells(i, 1).Value)) = CLng(dataNorm) Then
                    Set lr = lo.ListRows(i)
                    Exit For
                End If
            End If
        Next i
    End If
    If lr Is Nothing Then Set lr = lo.ListRows.Add
    lr.Range.Cells(1, 1).Value = dataNorm
    lr.Range.Cells(1, 2).Value = cdi
    lr.Range.Cells(1, 3).Value = ipca
    lr.Range.Cells(1, 4).Value = dataOriginal
End Sub

Private Sub RegistrarLog(arquivo As String, dataNorm As Date, linhas As Long, _
                         dataOriginal As Date, obs As String)
    Dim lr As ListRow
    Set lr = ThisWorkbook.Worksheets("Log").ListObjects("tblLog").ListRows.Add
    With lr.Range
        .Cells(1, 1).Value = Now
        .Cells(1, 1).NumberFormat = "dd/mm/yyyy hh:mm"
        .Cells(1, 2).Value = Environ("USERNAME")
        .Cells(1, 3).Value = arquivo
        .Cells(1, 4).Value = dataNorm
        .Cells(1, 4).NumberFormat = "mm/yyyy"
        .Cells(1, 5).Value = linhas
        .Cells(1, 6).Value = dataOriginal
        .Cells(1, 6).NumberFormat = "dd/mm/yyyy"
        .Cells(1, 7).Value = IIf(obs = "", "Sem observacoes", obs)
    End With
End Sub

Private Function NomeDoArquivo(caminho As String) As String
    ' Guarda so o nome do arquivo (sem o caminho) para nao expor pastas no log.
    Dim p As Long
    p = InStrRev(caminho, "\")
    If p = 0 Then p = InStrRev(caminho, "/")
    NomeDoArquivo = Mid(caminho, p + 1)
End Function

Private Sub Restaurar(calcAnterior As XlCalculation)
    Application.Calculation = calcAnterior
    Application.ScreenUpdating = True
End Sub
