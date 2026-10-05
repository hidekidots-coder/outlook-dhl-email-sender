Option Explicit

' ============================================================
' OUTLOOK DHL EMAIL SENDER (LOOP)
' ============================================================
'
' Purpose:
'   Automatically create Outlook emails for DHL shipments
'   based on data from an Excel sheet, attach matching PDF/XML
'   files and manage the process in a loop.
'
' Technologies:
'   VBA
'   Microsoft Outlook Object Model
'   Microsoft Excel
'
' ============================================================

Public Sub Enviar_Email_DHL_Loop()

    On Error GoTo ErroHandler

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim wsSobras As Worksheet
    
    Dim EmailPara As String
    Dim EmailCC As String
    Dim assunto As String
    Dim corpo As String
    
    Dim AWB As String
    Dim ReferenciaTE As String
    Dim TextoExtra As String
    Dim ValorA2 As String
    
    Dim CaminhoPasta As String
    Dim Arquivo As String
    Dim Extensao As String
    Dim ApenasNumeros As String
    
    Dim ArquivosAnexados As Collection
    Dim ItemArquivo As Variant
    Dim i As Long
    Dim LinhaSobras As Long

    ' --------------------------------------------------------
    ' Configuration
    ' --------------------------------------------------------
    
    Set ws = ThisWorkbook.Sheets("ENVIO AWB LOOP")
    Set wsSobras = ThisWorkbook.Sheets("ENVIO AWB SEM LOOP")
    
    CaminhoPasta = "C:\Temp\HAWB\"          ' <-- Altere para o caminho desejado
    
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
    ' Main Loop
    ' --------------------------------------------------------
    
    Do While Trim(ws.Range("A2").Value) <> ""
        
        ValorA2 = Trim(CStr(ws.Range("A2").Value))
        AWB = Replace(Trim(CStr(ws.Range("B2").Text)), " ", "")
        ReferenciaTE = Trim(CStr(ws.Range("C2").Text))
        TextoExtra = Trim(CStr(ws.Range("D2").Text))
        
        assunto = "TE BRASIL - " & ValorA2 & " - " & ReferenciaTE & " - AWB " & AWB
        
        ' ----------------------------------------------------
        ' Example emails (replace with real ones)
        ' ----------------------------------------------------
        
        EmailPara = "destinatario1@empresa.com; destinatario2@empresa.com"
        
        EmailCC = "copia1@empresa.com; copia2@empresa.com; copia3@empresa.com"
        
        corpo = "Olá, pessoal!" & vbNewLine & vbNewLine & _
                "Segue envio DHL Express Formal." & vbNewLine & vbNewLine & _
                "Time responsável," & vbNewLine & vbNewLine & _
                TextoExtra & vbNewLine & vbNewLine & _
                "Segue referência TE para este embarque: " & ReferenciaTE & "."
        
        Set ArquivosAnexados = New Collection
        Set OutlookMail = OutlookApp.CreateItem(0)
        
        With OutlookMail
            .To = EmailPara
            .CC = EmailCC
            .Subject = assunto
            .Body = corpo
            
            ' ------------------------------------------------
            ' Search and attach matching PDF / XML files
            ' ------------------------------------------------
            
            Arquivo = Dir(CaminhoPasta & "*.*")
            
            Do While Arquivo <> ""
                
                Extensao = LCase(Mid(Arquivo, InStrRev(Arquivo, ".") + 1))
                
                If Extensao = "pdf" Or Extensao = "xml" Then
                    
                    ApenasNumeros = ""
                    
                    For i = 1 To Len(Arquivo)
                        If Mid(Arquivo, i, 1) Like "[0-9]" Then
                            ApenasNumeros = ApenasNumeros & Mid(Arquivo, i, 1)
                        End If
                    Next i
                    
                    If InStr(1, ApenasNumeros, AWB, vbTextCompare) > 0 Then
                        .Attachments.Add CaminhoPasta & Arquivo
                        ArquivosAnexados.Add Arquivo
                    End If
                    
                End If
                
                Arquivo = Dir
            Loop
            
        End With
        
        ' ----------------------------------------------------
        ' If no attachments found → move to "Sobras" sheet
        ' ----------------------------------------------------
        
        If ArquivosAnexados.Count = 0 Then
            
            LinhaSobras = wsSobras.Cells(wsSobras.Rows.Count, 1).End(xlUp).Row + 1
            ws.Rows(2).Cut Destination:=wsSobras.Rows(LinhaSobras)
            
        Else
            
            OutlookMail.Display
            
            ' Delete only the attached files
            For Each ItemArquivo In ArquivosAnexados
                Kill CaminhoPasta & ItemArquivo
            Next ItemArquivo
            
            ws.Rows(2).Delete Shift:=xlUp
            
        End If
        
    Loop
    
    MsgBox "Processo finalizado! Planilha zerada.", vbInformation

Saida:
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    Set ws = Nothing
    Set wsSobras = Nothing
    Set ArquivosAnexados = Nothing
    Exit Sub

ErroHandler:
    MsgBox "Erro: " & Err.Description, vbCritical
    Resume Saida
    
End Sub
