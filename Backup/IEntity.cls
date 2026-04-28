VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==================================================
 ' Interface IEntity
' ==================================================

' This interface is for classes that represent a row in a ListObject
' The Populate method is automatically being invoked from CreateFromFactory()
'
' Usage:
'
'   dim tableClass as IEntity
'   set tableClass = CreateFromFactory(aTable, aRow)
'
'or
'
'   dim tableClass as MyTableClass
'   set tableClass = CreateFromFactory(aTable, aRow)


' IEntity_Populate reads the data from the table and calls Assignments which must put the values into the data variables
Public Sub Populate(Optional ByVal aRow As Long = 0)
End Sub
