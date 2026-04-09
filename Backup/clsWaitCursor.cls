VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWaitCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *****************************************************
' Klasse JWaitCursor
' *****************************************************

' Diese Klasse implementiert einen Wait Cursor, der automatisch
' wieder zum normalen Cursor wird, sobald der Gültigkeitsbereich
' des Objekts verlassen wird

' Anwendung:
' dim wait as new JWaitCursor
' wait.show

' =====================================================
' Public
' =====================================================

' Der Cursor ändert sich auf xlWait
Public Sub Show(Optional ByVal Kind As XlMousePointer = XlMousePointer.xlWait)
    If Not TESTMODE Then Application.Cursor = xlWait
End Sub

' Der Cursor ändert sich zurück auf den normalen Zeiger
Public Sub Hide()
    If Not TESTMODE Then Application.Cursor = xlDefault
End Sub

' =====================================================
' Private
' =====================================================

' Der Cursor ändert sich automatisch auf den normalen Zeiger, wenn die Instanz
' ihren Gültigkeitsbereich verlässt
Private Sub Class_Terminate()
    Hide
End Sub

