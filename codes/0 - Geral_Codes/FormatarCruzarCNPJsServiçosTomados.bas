VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function NormalizaCNPJ(ByVal valor As String) As String
    ' Remove espa�os normais e invis�veis
    valor = Trim(Replace(valor, Chr(160), ""))
    
    ' Garante que est� como texto
    valor = CStr(valor)
    
    NormalizaCNPJ = valor
End Function

Sub FormatarCruzarCNPJsServicosTomados()
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim ultima1 As Long, ultima2 As Long
    Dim dict As Object
    Dim i As Long
    Dim valuecel As String
    
    Dim cnpj As String
    Dim cpf As String
    
    ' Materializando um Objeto Regex
    Dim regexCnpj As New RegExp
    Dim regexCpf As New RegExp
    Dim match As Object ' Definindo Objeto que vai Armazenar o nosso grupo de Objetos e ocorr�ncias
    
    ' Configurando o nosso Regex
    With regexCnpj
        .Global = False
        .Pattern = "(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})"
    End With
    
    With regexCpf
        .Global = False
        .Pattern = "(\d{3})(\d{3})(\d{3})(\d{2})"
    End With
    
    Set ws1 = ThisWorkbook.Sheets("Servicos Tomados")   ' ajuste o nome se necess�rio
    Set ws2 = ThisWorkbook.Sheets("Consta_SN_ServicosTomados")   ' ajuste o nome se necess�rio

    Set dict = CreateObject("Scripting.Dictionary")

    ' ---- LER A PLANILHA 2 ----
    ultima2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To ultima2
        Dim chave As String
        chave = NormalizaCNPJ(ws2.Cells(i, "A").Value)
        
        If Not dict.exists(chave) Then
            dict.Add chave, ws2.Cells(i, "B").Value
        End If
    Next i

    ' ---- PREENCHER NA PLANILHA 1 ----
    ultima1 = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row
    
    For i = 5 To ultima1
        
        valuecel = ws1.Cells(i, "D").Value
               
        If Trim(valuecel) <> "" And Len(valuecel) = 11 Then
        
            If regexCpf.Test(valuecel) Then
                
                Set match = regexCpf.Execute(valuecel)(0)
                cpf = CStr(match.SubMatches(0) & "." & match.SubMatches(1) & "." & match.SubMatches(2) & "-" & match.SubMatches(3))
                
                ws1.Cells(i, "D").Value = cpf
                
            End If
        
        ElseIf Trim(valuecel) <> "" And Len(valuecel) = 14 Then
            
            If regexCnpj.Test(valuecel) Then
                
                Set match = regexCnpj.Execute(valuecel)(0)
                cnpj = CStr(match.SubMatches(0) & "." & match.SubMatches(1) & "." & match.SubMatches(2) & "/" & match.SubMatches(3) & "-" & match.SubMatches(4))
                
                ws1.Cells(i, "D").Value = cnpj
            
            End If
        
        End If
        
        cnpj = NormalizaCNPJ(ws1.Cells(i, "D").Value)
        
        If dict.exists(cnpj) Then
            ws1.Cells(i, "E").Value = dict(cnpj)
        Else
            ws1.Cells(i, "E").Value = "N�O ENCONTRADO"
        End If
    Next i

    MsgBox "Processo conclu�do com sucesso!", vbInformation

End Sub

