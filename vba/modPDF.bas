Attribute VB_Name = "modPDF"
Option Explicit

' =====================================================================
' modPDF - exporta as abas do relatorio para um unico PDF.
'
' - ERRO nas conferencias bloqueia o PDF final (permite PRELIMINAR).
' - ALERTA nao bloqueia, mas pede confirmacao.
' - Avisa se o comentario da analista nao foi preenchido.
' - Nome padronizado: Relatorio_Investimentos_AAAA-MM[_PRELIMINAR].pdf
' =====================================================================

Private Const TEXTO_PADRAO As String = "Digite aqui a analise do mes antes de gerar o PDF."
Private Const CELULA_COMENTARIO As String = "A67"

Public Sub GerarPDF()
    Dim erros As Long, alertas As Long, prelim As Boolean
    Dim pasta As String, arquivo As String, coment As String
    Dim mesRef As Date
    Dim abas As Variant

    If Not ArquivoCorreto() Then Exit Sub
    On Error GoTo Falha

    Application.Calculate
    erros = CLng(ThisWorkbook.Names("QtdErros").RefersToRange.Value)
    alertas = CLng(ThisWorkbook.Names("QtdAlertas").RefersToRange.Value)

    If erros > 0 Then
        If MsgBox("Ha " & erros & " conferencia(s) com ERRO. O PDF final esta bloqueado." & _
                  vbCrLf & vbCrLf & "Deseja gerar uma versao PRELIMINAR?", _
                  vbYesNo + vbExclamation, "Conferencias") = vbNo Then
            ThisWorkbook.Worksheets("Conferencias").Activate
            Exit Sub
        End If
        prelim = True
    ElseIf alertas > 0 Then
        If MsgBox("Ha " & alertas & " ALERTA(s). Eles nao bloqueiam o PDF, mas devem ser " & _
                  "justificados no comentario." & vbCrLf & vbCrLf & "Continuar?", _
                  vbYesNo + vbQuestion, "Conferencias") = vbNo Then Exit Sub
    End If

    coment = Trim(CStr(ThisWorkbook.Worksheets("Resumo").Range(CELULA_COMENTARIO).Value))
    If coment = "" Or StrComp(coment, TEXTO_PADRAO, vbTextCompare) = 0 Or _
       InStr(1, coment, "Digite aqui", vbTextCompare) = 1 Then
        If MsgBox("O comentario da analista nao foi preenchido. Gerar mesmo assim?", _
                  vbYesNo + vbQuestion, "Comentario") = vbNo Then
            ThisWorkbook.Worksheets("Resumo").Activate
            ThisWorkbook.Worksheets("Resumo").Range(CELULA_COMENTARIO).Select
            Exit Sub
        End If
    End If

    pasta = DefinirPasta()
    If pasta = "" Then Exit Sub

    mesRef = ThisWorkbook.Names("MesRef").RefersToRange.Value
    arquivo = pasta & Application.PathSeparator & "Relatorio_Investimentos_" & _
              Format(mesRef, "yyyy-mm") & IIf(prelim, "_PRELIMINAR", "") & ".pdf"

    If Dir(arquivo) <> "" Then
        If MsgBox("Ja existe um arquivo com este nome:" & vbCrLf & arquivo & vbCrLf & vbCrLf & _
                  "Substituir?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    ThisWorkbook.Names("StatusRelatorio").RefersToRange.Value = IIf(prelim, "PRELIMINAR", "FINAL")
    Application.Calculate

    abas = Array("Resumo", "Conferencias")
    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(abas).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=arquivo, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
    ThisWorkbook.Worksheets("Resumo").Select

    ThisWorkbook.Names("StatusRelatorio").RefersToRange.Value = "FINAL"
    MsgBox "PDF gerado:" & vbCrLf & arquivo, vbInformation
    Exit Sub

Falha:
    Dim msg As String
    msg = "Erro " & Err.Number & ": " & Err.Description
    On Error Resume Next
    ThisWorkbook.Worksheets("Resumo").Select
    ThisWorkbook.Names("StatusRelatorio").RefersToRange.Value = "FINAL"
    MsgBox "Nao foi possivel gerar o PDF." & vbCrLf & msg & vbCrLf & vbCrLf & _
           "Se o arquivo estiver aberto em outro programa, feche-o e tente novamente.", vbCritical
End Sub

Private Function DefinirPasta() As String
    Dim p As String
    p = Trim(CStr(ThisWorkbook.Names("PastaPDF").RefersToRange.Value))
    If p = "" Then p = ThisWorkbook.Path

    ' Arquivos em OneDrive/SharePoint podem ter caminho "https://...": pergunta a pasta.
    If p = "" Or LCase(Left(p, 4)) = "http" Or Dir(p, vbDirectory) = "" Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Escolha a pasta onde o PDF sera salvo"
            If .Show <> -1 Then Exit Function
            p = .SelectedItems(1)
        End With
    End If
    If Right(p, 1) = Application.PathSeparator Then p = Left(p, Len(p) - 1)
    DefinirPasta = p
End Function
