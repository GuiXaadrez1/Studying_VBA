Attribute VB_Name = "Módulo3"
Option Explicit

' ===========================================================
' CONFIGURAÇÃO
' ===========================================================

Const SourceSheetName As String = "Itens das NF-es Recebidas - Aut"
Const SourceColumn As String = "G"
Const OutputColumn As String = "M"
Const StartRowSource As Long = 4

Const ReductionSheetName As String = "ReducaoNCM"
Const ReductionCodeColumn As String = "A"
Const ReductionTaxColumn As String = "G"
Const StartRowReduction As Long = 2

Const SheetCName As String = "PlanilhaC"
Const Ignore9DigitsInSheetC As Boolean = True

' ===========================================================
' FUNÇÕES AUXILIARES
' ===========================================================

Private Function SomenteDigitos(ByVal s As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = "\D"
        .Global = True
    End With
    SomenteDigitos = re.Replace(Trim$(s), "")
End Function


Private Function GerarNiveisNCM(ByVal codigo As String) As Collection
    Dim c As New Collection
    Dim t As Long: t = Len(codigo)

    ' Níveis do mais específico ao mais genérico
    If t >= 8 Then c.Add Left$(codigo, 8)  ' completo
    If t >= 7 Then c.Add Left$(codigo, 7)
    If t >= 6 Then c.Add Left$(codigo, 6)
    If t >= 5 Then c.Add Left$(codigo, 5)  ' << NÍVEL 5 — agora aplicado corretamente
    If t >= 4 Then c.Add Left$(codigo, 4)
    If t >= 2 Then c.Add Left$(codigo, 2)
    If t >= 1 Then c.Add Left$(codigo, 1)

    Set GerarNiveisNCM = c
End Function


Private Function BuildReductionDict(ws As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, ReductionCodeColumn).End(xlUp).Row
    
    Dim i As Long, raw As String, norm As String, taxa As Variant
    
    For i = StartRowReduction To lastRow
        raw = CStr(ws.Cells(i, ReductionCodeColumn).Value)
        norm = SomenteDigitos(raw)
        
        If Len(norm) > 0 Then
            If ws.Name = SheetCName And Ignore9DigitsInSheetC And Len(norm) = 9 Then
                ' ignora
            Else
                taxa = ws.Cells(i, ReductionTaxColumn).Value
                If Not dict.Exists(norm) Then dict.Add norm, taxa
            End If
        End If
    Next i
    
    Set BuildReductionDict = dict
End Function


' ===========================================================
' SUB PRINCIPAL
' ===========================================================

Public Sub CruzarNcm_PorNiveis()
    On Error GoTo ErrHandler
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsSrc As Worksheet: Set wsSrc = wb.Worksheets(SourceSheetName)
    Dim wsRed As Worksheet: Set wsRed = wb.Worksheets(ReductionSheetName)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim dict As Object: Set dict = BuildReductionDict(wsRed)
    
    Dim last As Long
    last = wsSrc.Cells(wsSrc.Rows.Count, SourceColumn).End(xlUp).Row
    
    Dim i As Long, raw As String, norm As String
    Dim niveis As Collection, nivel As Variant
    Dim found As Boolean, result As Variant
    
    For i = StartRowSource To last
        
        raw = CStr(wsSrc.Cells(i, SourceColumn).Value)
        norm = SomenteDigitos(raw)
        
        ' --- REGRAS ESPECIAIS PARA PLANILHA C ---
        If wsSrc.Name = SheetCName And Ignore9DigitsInSheetC And Len(norm) = 9 Then
            wsSrc.Cells(i, OutputColumn).Value = "Ignorado (9 dígitos)"
            GoTo NextRow
        End If
        
        If Len(norm) = 0 Then
            wsSrc.Cells(i, OutputColumn).Value = "0%"
            GoTo NextRow
        End If
        
        ' --- GERAR NÍVEIS DO NCM (inclui nível 5) ---
        Set niveis = GerarNiveisNCM(norm)
        found = False
        
        ' -----------------------------------------
        ' APLICA A REDUÇÃO DO MAIS ESPECÍFICO PARA O MAIS GENÉRICO
        ' Inclui nível 5 de forma correta
        ' -----------------------------------------
        For Each nivel In niveis
            If dict.Exists(CStr(nivel)) Then
                result = dict(CStr(nivel))
                found = True
                Exit For
            End If
        Next nivel
        
        If found Then
            wsSrc.Cells(i, OutputColumn).Value = result
        Else
            wsSrc.Cells(i, OutputColumn).Value = "0%"
        End If
        
NextRow:
    Next i
    
    MsgBox "Processamento concluído.", vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical
    Resume CleanExit
End Sub

