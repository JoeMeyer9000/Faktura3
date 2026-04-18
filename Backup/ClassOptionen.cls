VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassOptionen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==================================================
' ClassOptionen
' ==================================================

Implements IDataEntity

' Ceolumn Indexes
Public Enum ColIdxOption
    xblgDatum = 1
    xblgKundenname
    xblgBelegArt
    xblgNetto
    xblgSteuer
    xblgMitarbeiter
    xblgBrutto
    xblgBelegNr
    xblgReferenz
    xblgDebitor
    xblgBelegID
    xblgPeriode
End Enum

' Data Variables, filled in Populate()
Public Datum As Date
Public Kundenname As String
Public BelegArt As String
Public Netto As Currency
Public Steuer As Double
Public Mitarbeiter As String
Public Brutto As Currency
Public BelegNr As String
Public Referenz As String
Public Debitor As Long
Public BelegID As Long
Public Periode As String

' -----------------------------------------------
' Private property data variables
' -----------------------------------------------
Private m_Row As Long

' -----------------------------------------------
' Public properties
' -----------------------------------------------

Public Property Get IDataEntity_Table() As ListObject
    Set IDataEntity_Table = TOptionen
End Property

Public Property Get IDataEntity_Row() As Long
    IDataEntity_Row = m_Row
End Property

' -----------------------------------------------
' Public methods
' -----------------------------------------------

Public Sub IDataEntity_Populate(ByVal aRow As Long)
    m_Row = aRow
    If aRow < 1 Then Exit Sub
    
    Dim values As Variant
    values = IDataEntity_Table.ListRows(m_Row).Range.Value2
    
    Datum = values(1, xblgDatum)
    Kundenname = values(1, xblgKundenname)
    BelegArt = values(1, xblgBelegArt)
    Netto = values(1, xblgNetto)
    Steuer = values(1, xblgSteuer)
    Mitarbeiter = values(1, xblgMitarbeiter)
    Brutto = values(1, xblgBrutto)
    BelegNr = values(1, xblgBelegNr)
    Referenz = values(1, xblgReferenz)
    Debitor = values(1, xblgDebitor)
    BelegID = values(1, xblgBelegID)
    Periode = values(1, xblgPeriode)
End Sub


