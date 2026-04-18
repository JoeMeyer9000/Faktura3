Attribute VB_Name = "JMail"
Option Explicit

' *****************************************************
' Module JMail
' *****************************************************

' Dieses Modul erwartet, dass die Microsoft Outlook 16.0 Library in den
' Verweisen referenziert wird.
' Die beiden Funktionen erzeugen aus den übergebenen Paranmetern eine
' Outlook E-Mail und zeigt sie an. Die Mail wird nur angezeigt, aber nicht versendet

' Erzeugt eine Outlook-Mail und zeigt sie an. Die Mail wird nicht gesendet
' Mehrere Dateinamen trennt man mit Semikolon
Public Sub GenerateMail( _
            ByVal sendTo As String, _
            ByVal sendCC As String, _
            ByVal sendBCC As String, _
            ByVal betreff As String, _
            ByVal Text As String, _
            Optional ByVal filenames As String = "")
    ' Optimierung: wenn keine Anhänge übergeben wurden, kann man auch die eigentliche
    ' GenerateMailWithAttm aufrufen und einfach keine Filename-Collection angeben
    If filenames = "" Then
        GenerateMailWithAttm sendTo, sendCC, sendBCC, betreff, Text
        Exit Sub
    End If
    
    ' Wenn mehrere Filenames übergebn wurden, müssen wir sie in eine Collection aufteilen,
    ' damit wir GenerateMailWithAttm aufrufen können
    Dim fnameArray As Variant
    Dim fname As Variant
    Dim coll As New Collection
    
    fnameArray = Split(filenames, ";")
    For Each fname In fnameArray
        coll.Add fname
    Next fname
    
    GenerateMailWithAttm sendTo, sendCC, sendBCC, betreff, Text, coll
End Sub

' Erzeugt eine Outlook-Mail und zeigt sie an. Die Mail wird nicht gesendet
Public Sub GenerateMailWithAttm( _
            ByVal sendTo As String, _
            ByVal sendCC As String, _
            ByVal sendBCC As String, _
            ByVal betreff As String, _
            ByVal Text As String, _
            Optional ByVal filenames As Collection = Nothing)
    Dim olApp As Outlook.Application
    Set olApp = New Outlook.Application
    
    Dim olMail As MailItem
    Set olMail = olApp.CreateItem(olMailItem)
    olMail.Display
    olMail.HTMLBody = Text + olMail.HTMLBody
    
    olMail.Categories = "Rechnung"
    olMail.To = sendTo
    If sendCC <> "" Then olMail.cc = sendCC
    If sendBCC <> "" Then olMail.BCC = sendBCC
    olMail.Subject = betreff
    
    Dim fname As Variant
    For Each fname In filenames
        olMail.attachments.Add fname
    Next fname
End Sub


