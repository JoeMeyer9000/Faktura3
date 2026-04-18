VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MessageBox 
   Caption         =   "Message Box"
   ClientHeight    =   2730
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6555
   OleObjectBlob   =   "MessageBox.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "MessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ***********************************************
' Formular MessageBox
' ***********************************************

' Man kann auch MsgBox verwenden, aber dann sieht die MessageBox echt besch***en aus
' Dieser Dialog ist auch keinen Designer-Preis wert, aber ich finde ihn schon besser
' als die eingebaute MsgBox.
' Die MessageBox kann einen Titel haben, das ist ein String über der Message in Fettschrift
' Der Text steht unter dem Titel

' MessageBoxStyle bestimmt, welches Icon angezeigt wird
Public Enum MessageBoxStyle
    mbsMessage = 1
    mbsQuestion
    mbsWarning
    mbsError
End Enum

' MessageBoxButtons bestimmt, welche Buttons angezeigt werden
' Die Werte sind Bitmasken, können also mit Or kombiniert werden. Z. Bsp.:
' [mbYes or mbCancel] zeigen Yes-Button und Cancel-Button an
Public Enum MessageBoxButtons
    mbbYes = 1
    mbbNo = 2
    mbbOk = 4
    mbbCancel = 8
    mbbOkCancel = mbbOk + mbbCancel
    mbbYesNo = mbbYes + mbbNo
    mbbYesNoCancel = mbbYes + mbbNo + mbbCancel
End Enum

' ModalResult gibt an, mit welchem Button der Dialog geschlossen wurde
Public ModalResult As VbMsgBoxResult

Private ButtonLeft As Double
Private ButtonWidth As Double
Private Const ButtonGap As Integer = 12

' Setzt die Titelzeile
Public Property Let Title(ByVal value As String)
    lblTitle.Caption = value
End Property

' Setzt den Text unterhalb des Titel
Public Property Let Text(ByVal value As String)
    lblText.Caption = value
End Property

' Setzt den Style und bestimmt damit, welches Icon angezeigt wird
Public Property Let Style(ByVal value As MessageBoxStyle)
    ' das gewünschte Icon auf sichtbar setzen
    imgMessage.visible = (value = mbsMessage)
    imgQuestion.visible = (value = mbsQuestion)
    imgWarning.visible = (value = mbsWarning)
    imgError.visible = (value = mbsError)
End Property

' Bestimmt, welche Buttons erscheinen. Es dürfen auch mehr als 1 Button sein. Beispiele:
' mbbCancel zeigt den Cancel-Button an
' mbbYes mbbCancel zeigt Yes-Button und Cancel-Button an
Public Property Let Buttons(ByVal value As MessageBoxButtons)
    ' erstmal alle Buttons auf unsichtbar setzen
    ResetButtons
    
    ' je nachdem welche Buttons in der Bitmaske value stecken, diese anzeigen
    ShowButton btnCancel, (value And mbbCancel) > 0
    ShowButton btnNo, (value And mbbNo) > 0
    ShowButton btnOk, (value And mbbOk) > 0
    ShowButton btnYes, (value And mbbYes) > 0

    ' mal sehen, wer der Cancel Button wird, damit man den Dialog mit der
    ' ESC-Taste schliessen kann
    If btnCancel.visible Then
        btnCancel.Cancel = True
    ElseIf btnOk.visible Then
        btnOk.Cancel = True
    ElseIf btnYes.visible Then
        btnYes.Cancel = True
    End If
End Property

' Zeigt die MessageBox an
' In [titleAndText] übergeben wir Titelzeile und den Text, betrennt durch Pipe-Symbol |
' [titleAndText] kann \n enthalten, daraus wird ein Zeilenumbruch
Public Function ShowMessage(ByVal titleAndText As String, _
    Optional ByVal msgStyle As MessageBoxStyle = mbsMessage, _
    Optional ByVal msgButtons As MessageBoxButtons = mbbOk) As VbMsgBoxResult
    
    titleAndText = Replace(titleAndText, "\r\n", vbCrLf)
    titleAndText = Replace(titleAndText, "\n", vbCrLf)
    titleAndText = Replace(titleAndText, "\r", vbCrLf)
    
    ' Titel und Text trennen
    Dim p As Integer
    Dim aTitle As String
    Dim aText As String
    p = InStr(1, titleAndText, "|")
    If p > 0 Then
        aTitle = Left(titleAndText, p - 1)
        aText = Right(titleAndText, Len(titleAndText) - p)
    Else
        aText = titleAndText
    End If
    
    ' Parameter des Dialogs setzen
    Title = aTitle
    Text = aText
    Style = msgStyle
    Buttons = msgButtons
    
    ' Dialog anzeigen
    Show vbModal
    
    ' das Ergebnis ist vom Typ VbMsgBoxResult
    ShowMessage = Me.ModalResult
End Function

' Beim Instanziieren wird die Überschrift des Fensters gesetzt
Private Sub UserForm_Initialize()
    Me.Caption = ThisWorkbook.Name
    ResetButtons
End Sub

' Alle Buttons auf unsichtbar
Private Sub ResetButtons()
    btnYes.visible = False
    btnNo.visible = False
    btnOk.visible = False
    btnCancel.visible = False
    
    ButtonWidth = btnCancel.Width
    ButtonLeft = Me.Width - ButtonWidth - 18
End Sub

' Button anzeigen und seine Position berechnen
Private Sub ShowButton(ByVal button As CommandButton, ByVal makeVisible As Boolean)
    button.visible = makeVisible
    If Not makeVisible Then Exit Sub
        
    button.Left = ButtonLeft
    ButtonLeft = ButtonLeft - ButtonGap - ButtonWidth
End Sub

Private Sub btnCancel_Click()
    ModalResult = vbCancel
    Me.Hide
End Sub

Private Sub btnNo_Click()
    ModalResult = vbNo
    Me.Hide
End Sub

Private Sub btnOk_Click()
    ModalResult = vbOK
    Me.Hide
End Sub

Private Sub btnYes_Click()
    ModalResult = vbYes
    Me.Hide
End Sub
