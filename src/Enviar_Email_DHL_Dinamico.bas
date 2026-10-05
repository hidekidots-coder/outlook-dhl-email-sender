Option Explicit

' ============================================================
' OUTLOOK DHL EMAIL SENDER (DYNAMIC)
' ============================================================
'
' Purpose:
'   Create a single Outlook email based on the first row of
'   the sheet "ENVIO AWB SEM LOOP", then delete the row.
'
' Technologies:
'   VBA
'   Microsoft Outlook Object Model
'   Microsoft Excel
'
' ============================================================

Public Sub Enviar_Email_DHL_Dinamico()

    On Error GoTo ErroHandler

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    
    Dim EmailPara As String
    Dim EmailCC As String
    Dim assunto As String
    Dim corpo As String
    
    Dim ValorA As String
    Dim ValorB As String
    Dim ValorC As String
    Dim ValorD As String

    ' --------------------------------------------------------
    ' Configuration
    ' --------------------------------------------------------
    
    Set ws = ThisWorkbook.Sheets("ENVIO AWB SEM LOOP")
    
    ' --------------------------------------------------------
    ' Validate if there is data
    ' --------------------------------------------------------
    
    If WorksheetFunction.CountA(ws.Rows(1)) = 0 Then
        MsgBox "A linha 1 está vazia. Nada para enviar.", vbExclamation
        Exit Sub
    End If
    
    ValorA = Trim(CStr(ws.Range("A1").Value))
    ValorB = Trim(CStr(ws.Range("B1").Value))
    ValorC = Trim(CStr(ws.Range("C1").Value))
    ValorD = Trim(CStr(ws.Range("D1").Value))
    
    assunto = "TE BRASIL - " & ValorA & " - " & ValorC & " - AWB " & ValorB
    
    ' --------------------------------------------------------
    ' Example emails (replace with real ones)
    ' --------------------------------------------------------
    
    EmailPara = "destinatario1@empresa.com; destinatario2@empresa.com; destinatario3@empresa.com"
    
    EmailCC = "copia1@empresa.com; copia2@empresa.com; copia3@empresa.com"
    
    corpo = "Olá, pessoal!" & vbNewLine & vbNewLine & _
            "Segue envio DHL Express Formal." & vbNewLine & vbNewLine & _
            "Time responsável," & vbNewLine & vbNewLine & _
            ValorD & vbNewLine & vbNewLine & _
            "Segue referência TE para este embarque: " & ValorC & "."
    
    ' --------------------------------------------------------
    ' Initialize Outlook
    ' --------------------------------------------------------
    
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    On Error GoTo ErroHandler
    
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    
    ' --------------------------------------------------------
    ' Create and display email
    ' --------------------------------------------------------
    
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    With OutlookMail
        .To = EmailPara
        .CC = EmailCC
        .Subject = assunto
        .Body = corpo
        .Display          ' Use .Send to send automatically
    End With
    
    ' --------------------------------------------------------
    ' Delete the processed row
    ' --------------------------------------------------------
    
    ws.Rows(1).Delete Shift:=xlUp

Saida:
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    Set ws = Nothing
    Exit Sub

ErroHandler:
    MsgBox "Ocorreu um erro: " & Err.Description, vbCritical
    Resume Saida
    
End Sub
