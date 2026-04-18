VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Fortschritt..."
   ClientHeight    =   285
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4005
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************
' Formular Progressbar
' ************************************************

' Progressbar zeigt ein nicht-modales Fenster mit einem Fortschrittsbalken an
' Siehe hierzu auch Klasse JProgress

Private Const m_BalkenWidth As Double = 200
Private m_MaxValue As Long

' Liefert den aktuellen Wert des Balkens
Public Property Get value() As Variant
    value = Balken.Width
End Property

' Setzt den aktuellen Wert und damit die Breite des Balkens.
' Der Wert wird nie größer als die maximale Breite
Public Property Let value(ByVal newValue As Variant)
    newValue = Round(newValue, 0)
    If newValue > m_MaxValue Then newValue = m_MaxValue
    Balken.Width = Round(m_BalkenWidth * newValue / m_MaxValue, 0)
    Caption = newValue & " von " & m_MaxValue & " erledigt (" & Round(newValue * 100 / m_MaxValue, 0) & "%)"
    DoEvents
End Property

' Zeigt das nicht-modale Fenster an
Public Sub Create(Optional ByVal MaxValue As Long = 100)
    m_MaxValue = MaxValue
    ProgressBar.Show vbModeless
    value = 0
End Sub

' schliesst das Fenster
Public Sub CloseBar()
    Unload ProgressBar
End Sub

