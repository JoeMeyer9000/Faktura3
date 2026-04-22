VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFileListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==================================================
' ClassFileListItem
' ==================================================

Implements IEntity

' Column Indexes
Public Enum ColIdxKarteikarte
    xfliKarteikarte = 1
    xfliFileDate
    xfliSize
    xfliLink
    xfliFullname
End Enum

' Data Variables, filled in Populate()
Public Karteikarte As String
Public FileDate As Date
Public Size As Long
Public Link As String
Public Fullname As String

' -----------------------------------------------
' Private property data variables
' -----------------------------------------------
Private m_Row As Long

' -----------------------------------------------
' Public properties
' -----------------------------------------------

Public Property Get Table() As ListObject
    Set Table = shFiles.ListObjects(1)
End Property

Public Property Get Row() As Long
    Row = m_Row
End Property

' -----------------------------------------------
' Public methods
' -----------------------------------------------

Public Sub IEntity_Populate(Optional ByVal aRow As Long = 0)
    If aRow > 0 Then m_Row = aRow
    
    Dim values As Variant
    values = Table.ListRows(m_Row).Range.Value2
    
    Karteikarte = values(1, xfliKarteikarte)
    FileDate = values(1, xfliFileDate)
    Size = values(1, xfliSize)
    Link = values(1, xfliLink)
    Fullname = values(1, xfliFullname)

End Sub

Public Sub Convert()
    Dim kkOld As New ClassKKOld
    kkOld.LoadFromFile Me
End Sub
