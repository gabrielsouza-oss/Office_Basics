Public mes As String

Private Function sendEmail()

mes = 1 'Mês do aquivo a ser enviado

'Salva o arquivo aberto no formato XLSX
ActiveWorkbook.SaveAs Filename:="C:\temp\Arquivo.xlsx", FileFormat:=51

'Gera o email final
    Set olApp = getOutlookInstance
    Dim olMail As Outlook.MailItem
    Set olMail = olApp.CreateItem(olMailItem)
    
    With olMail
        .BodyFormat = olFormatHTML
        .Display
        .HTMLBody = getHTML & .HTMLBody
        .Subject = "Meu Arquivo " & getNomeMes(mes)
        .To = "Gabriel d Souza <manoalo420@gmail.com>; Jose Manito <josemanito@gmail.com>"
        .CC = ""
        .Attachments.Add "C:\temp\Arquivo.xlsx"
        .BCC = ""
    End With
    
End Function

Private Function getOutlookInstance() As Outlook.Application
'Encontra e Define a instância ativa do Outlook
    Dim olApp As Outlook.Application
    Dim olNameSpace As Outlook.Namespace
    Dim olInbox As Outlook.Folder
    On Error Resume Next
    Set getOutlookInstance = GetObject(Class:="Outlook.Application")
    Err.Clear
    On Error GoTo 0
    If getOutlookInstance Is Nothing Then
        Set getOutlookInstance = CreateObject("Outlook.Application")
        Set olNameSpace = getOutlookInstance.GetNamespace("MAPI")
        Set olInbox = olNameSpace.GetDefaultFolder(olFolderInbox)
        olInbox.Display
    End If
    Set olNameSpace = Nothing
    Set olInbox = Nothing
End Function

Private Function getHTML() As String
'Edita mensagem enviada no corpo do Email
            getHTML = "<font size='3' family='calibri'>" _
            & getSaudacao & " !<br>" _
            & "<br>Seguem anexas as apresentações" _
            & "<br>Quaisquer dúvidas, fico à disposição.</font>"
End Function

Private Function getSaudacao() As String
'Define a Saudação do E-mail com base no período do dia
    Dim hr As Integer
    hr = Hour(Now)
    Select Case hr
        Case Is < 12
            getSaudacao = "Bom dia"
        Case Is < 18
            getSaudacao = "Boa tarde"
        Case Else
            getSaudacao = "Boa noite"
    End Select
End Function

Private Function getNomeMes(ByVal mes As String) As String
'Converte um mês numerico para String
    Select Case (mes)
        Case 1
            getNomeMes = "Janeiro"
        Case 2
            getNomeMes = "Fevereiro"
        Case 3
            getNomeMes = "Março"
        Case 4
            getNomeMes = "Abril"
        Case 5
            getNomeMes = "Maio"
        Case 6
            getNomeMes = "Junho"
        Case 7
            getNomeMes = "Julho"
        Case 8
            getNomeMes = "Agosto"
        Case 9
            getNomeMes = "Setembro"
        Case 10
            getNomeMes = "Outubro"
        Case 11
            getNomeMes = "Novembro"
        Case 12
            getNomeMes = "Dezembro"
    End Select
End Function
