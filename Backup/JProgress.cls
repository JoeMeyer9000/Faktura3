VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *****************************************************
' Klasse JProgress
' *****************************************************

' Einfache Handhabung eines Fortschrittsbalkens.
' Wird der Fortschrittsbalken mit maxValue < 1 initialisiert, wird er gar nicht
' erst angezeigt und alle folgenden Aufrufe des Fortschrittsbalken haben keine Wirkung

' Dim progress As New JProgress
' progress.Initialize zeilen.count

Private m_Visible As Boolean

' liefert den aktuellen Wert des Fortschrittsbalkens, wenn er sichtbar ist
' Wenn der Fortschrittsbalken nicht sichtbar ist, liefert die Property -1
Public Property Get value() As Long
    If m_Visible Then
        value = ProgressBar.value
    Else
        value = -1
    End If
End Property

' Setzt den Wert des Fortschrittsbalkens, aber nur wenn er sichtbar ist
Public Property Let value(ByVal currentValue As Long)
    If m_Visible Then ProgressBar.value = currentValue
End Property

' Zeigt den Fortschrittsbalken an, wenn [maxCount] >= Setting(xsetConfirmLoop) ist
Public Sub Initialize(ByVal maxCount As Long)
    If maxCount >= Setting(xsetConfirmLoop) Then
        ProgressBar.Create maxCount
        m_Visible = True
    End If
End Sub

' Destruktor. Entfernt den Fortschrittsbalken, wenn er sichtbar war
Private Sub Class_Terminate()
    If m_Visible Then
        Unload ProgressBar
    End If
End Sub
