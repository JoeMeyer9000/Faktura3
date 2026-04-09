VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XRXml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' Class XRXml
' =========================================

' Helper for creating structured XML hierarchies

Private lines As Collection
Private Stack As Collection
Private Indent As String

' =============================================
' public
' =============================================

Public Property Get xml() As String
    Dim result As String, line
    
    For Each line In lines
        If result = "" Then
            result = line
        Else
            result = result & vbCrLf & line
        End If
    Next line
    
    xml = result
End Property

Public Sub StartGroup(ByVal Tag As String, Optional ByVal Attributes As String = "")
    Dim value As String
    
    If Attributes <> "" Then
        value = Indent & "<" & Tag & " " & Attributes & ">"
    Else
        value = Indent & "<" & Tag & ">"
    End If
    
    lines.Add value
    
    Stack.Add Tag
    Indent = String(Stack.Count * 2, " ")
End Sub

Public Sub CloseGroup(Optional ByVal ExpectedTag As String = "")
    Dim Tag As String
    Dim index As Integer
    
    index = Stack.Count
    Tag = Stack(index)
    If ExpectedTag <> "" And Tag <> ExpectedTag Then
        Err.Raise vbObjectError + 1, "XRXml.CloseGroup", "Expected Tag " & ExpectedTag
    End If
    Stack.remove index
    Indent = String(Stack.Count * 2, " ")
    
    lines.Add Indent & "</" & Tag & ">"
   
End Sub

Public Sub AddString(ByVal Tag As String, ByVal data As String, Optional ByVal Attributes As String = "")
    If data = "" Then Exit Sub
    
    Dim value As String
    
    If Attributes <> "" Then
        value = Indent & "<" & Tag & " " & Attributes & ">" & data & "</" & Tag & ">"
    Else
        value = Indent & "<" & Tag & ">" & data & "</" & Tag & ">"
    End If
    
    lines.Add value
End Sub

Public Sub Add(ByVal Tag As String, ByVal data As String, Optional ByVal Attributes As String = "")
    AddString Tag, Replace(data, vbCrLf, " "), Attributes
End Sub

Public Sub AddDate(ByVal Tag As String, ByVal aDate As Date)
    If aDate < DateSerial(2000, 1, 1) Then
        Exit Sub
    End If
    StartGroup Tag
    AddString "udt:DateTimeString", Format(aDate, "YYYYMMDD"), "format=""102"""
    CloseGroup
End Sub

Public Sub AddNumber(ByVal Tag As String, ByVal aDouble As Double, Optional ByVal Attributes As String = "")
    AddString Tag, Replace(CStr(WorksheetFunction.Round(aDouble, 2)), ",", "."), Attributes
End Sub

Public Sub AddInt(ByVal Tag As String, ByVal aInt As Integer, Optional ByVal Attributes As String = "")
    AddString Tag, CStr(aInt), Attributes
End Sub

' =============================================
' Private
' =============================================

Private Sub Class_Initialize()
    Set lines = New Collection
    Set Stack = New Collection
    Indent = ""
    
    lines.Add "<?xml version=""1.0"" encoding=""UTF-8""?>"
    lines.Add ""
End Sub

