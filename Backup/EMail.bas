Attribute VB_Name = "EMail"
Option Explicit

' Erzeugt eine Outlook-Mail und zeigt sie an. Die Mail wird nicht gesendet
Public Function GenerateEMail( _
            ByVal sendTo As String, _
            ByVal sendCC As String, _
            ByVal sendBCC As String, _
            ByVal betreff As String, _
            ByVal Text As String, _
            Optional ByVal filenames As Collection = Nothing) As MailItem
    Dim olApp As Outlook.Application
    Set olApp = New Outlook.Application
    
    Dim olMail As MailItem
    Set olMail = olApp.CreateItem(olMailItem)
    olMail.Display
    olMail.HTMLBody = Text + olMail.HTMLBody
    
    olMail.To = sendTo
    If sendCC <> "" Then olMail.CC = sendCC
    If sendBCC <> "" Then olMail.BCC = sendBCC
    olMail.Subject = betreff
    
    If Not (filenames Is Nothing) Then
        Dim fname As Variant
        For Each fname In filenames
            olMail.Attachments.Add fname
        Next fname
    End If
    
    Set GenerateEMail = olMail
End Function

