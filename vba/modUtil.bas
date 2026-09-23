Attribute VB_Name = "modUtil"
Option Explicit

' =====================================================================
' modUtil - botoes, navegacao e diagnostico.
' =====================================================================

Public Sub VerConferencias()
    If Not ArquivoCorreto() Then Exit Sub
    Application.Calculate
    ThisWorkbook.Worksheets("Conferencias").Activate
    ThisWorkbook.Worksheets("Conferencias").Range("A1").Select
End Sub

' Executar UMA vez (Alt+F8 > CriarBotoes). Pode ser repetida sem duplicar botoes.
Public Sub CriarBotoes()
    Dim ws As Worksheet, i As Long
    Dim rotulos As Variant, macros As Variant, cores As Variant

    If Not ArquivoCorreto() Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Resumo")

    rotulos = Array("Importar mes", "Ver conferencias", "Gerar PDF")
    macros = Array("ImportarMes", "VerConferencias", "GerarPDF")
    cores = Array(RGB(31, 58, 95), RGB(68, 84, 106), RGB(0, 112, 96))

    RemoverBotoes ws
    For i = 0 To 2
        CriarUmBotao ws, CStr(rotulos(i)), CStr(macros(i)), CLng(cores(i)), i
    Next i
    MsgBox "Botoes criados na aba Resumo (coluna I). Eles ficam fora da area de impressao.", vbInformation
End Sub

Private Sub CriarUmBotao(ws As Worksheet, rotulo As String, macro As String, cor As Long, ordem As Long)
    Dim esq As Double, topo As Double, sh As Shape, bt As Object
    esq = ws.Range("I1").Left + 6
    topo = ws.Range("I1").Top + 10 + ordem * 42

    ' Tentativa 1: forma colorida
    On Error Resume Next
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, esq, topo, 150, 32)
    If Err.Number = 0 And Not sh Is Nothing Then
        With sh
            .Name = "btn_" & macro
            .Fill.ForeColor.RGB = cor
            .Line.Visible = msoFalse
            .TextFrame2.TextRange.Text = rotulo
            .TextFrame2.TextRange.Font.Size = 11
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = macro
            .Placement = xlFreeFloating
        End With
    Else
        ' Tentativa 2 (fallback): botao de formulario
        Err.Clear
        If Not sh Is Nothing Then sh.Delete
        Set bt = ws.Buttons.Add(esq, topo, 150, 32)
        If Err.Number = 0 Then
            bt.Name = "btn_" & macro
            bt.Caption = rotulo
            bt.OnAction = macro
        Else
            MsgBox "Nao foi possivel criar o botao '" & rotulo & "'. " & _
                   "Use Alt+F8 e execute a macro " & macro & " diretamente.", vbExclamation
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub RemoverBotoes(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i
End Sub

' Verifica se o codigo esta no arquivo certo. "ThisWorkbook" e sempre o
' arquivo onde o modulo foi importado: se as abas nao existem, o modulo
' provavelmente foi importado no projeto VBA errado (erro 9).
Public Function ArquivoCorreto() As Boolean
    Dim faltando As String, nome As Variant, lo As ListObject

    For Each nome In Array("Base", "Resumo", "Conferencias", "Config", "Cadastros", "Indicadores", "Log")
        If Not AbaExiste(CStr(nome)) Then faltando = faltando & " - aba " & nome & vbCrLf
    Next nome
    If AbaExiste("Base") Then
        On Error Resume Next
        Set lo = ThisWorkbook.Worksheets("Base").ListObjects("tblBase")
        On Error GoTo 0
        If lo Is Nothing Then faltando = faltando & " - tabela tblBase" & vbCrLf
    End If

    If faltando <> "" Then
        MsgBox "Este codigo esta no arquivo '" & ThisWorkbook.Name & "', que nao e a base do relatorio." & _
               vbCrLf & vbCrLf & "Itens nao encontrados:" & vbCrLf & faltando & vbCrLf & _
               "Provavel causa: o modulo foi importado no projeto errado. No editor VBA (Alt+F11), " & _
               "selecione o arquivo da base no painel da esquerda antes de importar.", vbCritical
        Exit Function
    End If
    ArquivoCorreto = True
End Function

Private Function AbaExiste(nome As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nome)
    AbaExiste = Not ws Is Nothing
End Function

Public Sub Diagnostico()
    Dim txt As String, nm As Variant, ok As Boolean
    txt = "Arquivo com o codigo: " & ThisWorkbook.Name & vbCrLf
    txt = txt & "Arquivo ativo: " & ActiveWorkbook.Name & vbCrLf
    txt = txt & "Calculo automatico: " & IIf(Application.Calculation = xlCalculationAutomatic, "sim", "NAO") & vbCrLf & vbCrLf

    For Each nm In Array("MesRef", "MesAnt", "QtdErros", "QtdAlertas", "PastaPDF", "StatusRelatorio")
        ok = NomeExiste(CStr(nm))
        txt = txt & IIf(ok, "[ok] ", "[FALTA] ") & "nome " & nm & vbCrLf
    Next nm

    If ArquivoCorreto() Then
        Application.Calculate
        txt = txt & vbCrLf & "Mes de referencia: " & Format(ThisWorkbook.Names("MesRef").RefersToRange.Value, "mm/yyyy")
        txt = txt & vbCrLf & "Linhas na base: " & ThisWorkbook.Worksheets("Base").ListObjects("tblBase").ListRows.Count
        txt = txt & vbCrLf & "Conferencias: " & ThisWorkbook.Names("QtdErros").RefersToRange.Value & " erro(s), " & _
              ThisWorkbook.Names("QtdAlertas").RefersToRange.Value & " alerta(s)"
    End If
    MsgBox txt, vbInformation, "Diagnostico"
End Sub

Private Function NomeExiste(nome As String) As Boolean
    Dim n As Name
    On Error Resume Next
    Set n = ThisWorkbook.Names(nome)
    NomeExiste = Not n Is Nothing
End Function
