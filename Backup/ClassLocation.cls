VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================
' ClassLocation
' ============================================

Private m_Row As Long
Private m_Column As Long

' --------------------------------------------
' Properties
' --------------------------------------------

Public Property Get row() As Long
    row = m_Row
End Property

Public Property Let row(Value As Long)
    m_Row = Value
End Property

Public Property Get column() As Long
    column = m_Column
End Property

Public Property Let column(Value As Long)
    m_Column = Value
End Property

Public Property Get address() As String
    address = ConvertR1C1ToA1(m_Row, m_Column)
End Property

Public Property Let address(Value As String)
    ConvertA1ToR1C1 Value, m_Row, m_Column
End Property

Public Property Get AddressAtRow(ByVal aRow As Long) As String
    AddressAtRow = ConvertR1C1ToA1(aRow, m_Column)
End Property

Public Property Get AddressAtColumn(ByVal aColumn As Long) As String
    AddressAtColumn = ConvertR1C1ToA1(m_Row, aColumn)
End Property

Public Property Get ValueAtAddress(ByRef Values As Variant, ByVal aAddress As String) As Variant
    Dim r As Long, c As Long
    ConvertA1ToR1C1 aAddress, r, c
    ValueAtAddress = Values(r, c)
End Property

' --------------------------------------------
' Public
' --------------------------------------------

Public Sub NextRow()
    m_Row = m_Row + 1
End Sub

Public Sub NextColumn()
    m_Column = m_Column + 1
End Sub

