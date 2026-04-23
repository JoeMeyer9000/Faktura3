VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassKarteikarte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==================================================
' ClassKarteikarte
' ==================================================

Public Function CreateKarteikarte(ByVal filename As String) As Boolean
   
    On Error GoTo ErrorExit
    
    Dim app As New ClassApplication
    
    ' open source and get it's path
    Dim wbSource As Workbook
    Set wbSource = app.Workbooks.Open(filename, , True)
    
    Dim wbTarget As Workbook
    Set wbTarget = app.Workbooks.Add(PathCombine(GetWorkbookDirectory(), "KKTemplate.xltm"))
    
    Dim shSource As Worksheet
    Dim shTarget As Worksheet
    
    Set shSource = wbTarget.Worksheets(2)
    Dim i As Long
    For i = 1 To GetNumberOfVerträgeSheets(wbSource)
        shSource.Copy wbTarget.Worksheets(wbTarget.Worksheets.count)
    Next i
    
    Dim sheetName As String
    Dim sheetIndex As Long
    For sheetIndex = 1 To wbSource.Worksheets.count
        shSource = wbSource.Worksheets(sheetIndex)
        sheetName = shSource.Range("A1").value
        If StrEqual(sheetName, "Kunden-Karteikarte") Then
            TransferKunde shSource, wbTarget.Worksheets(sheetIndex)
        ElseIf StrEqual(sheetName, "Verträge von Kunde:") Then
            Set shTarget = IIf(sheetIndex > 2, wbTarget.Worksheets.Add, wbTarget.Worksheets(sheetIndex))
            TransferVerträge shSource, shTarget
        End If
    Next sheetIndex
    
    wbTarget.SaveAs Replace(filename, "Karteikarte", "KK", Compare:=vbTextCompare)
    
ErrorExit:
    If Not wbSource Is Nothing Then wbSource.Close (False)
    If Not wbTarget Is Nothing Then wbTarget.Close (False)
    
End Function

Private Sub TransferKunde(ByVal shSource As Worksheet, ByVal shTarget As Worksheet)

End Sub

Private Sub TransferVerträge(ByVal shSource As Worksheet, ByVal shTarget As Worksheet)

End Sub

Private Function GetNumberOfVerträgeSheets(ByVal wb As Workbook) As Long
    Dim sum As Long: sum = 0
    Dim i As Long
    For i = 1 To wb.Worksheets.count
        If StrEqual(wb.Worksheets(i).Range("A1").value, "Verträge von Kunde:") Then sum = sum + 1
    Next i
    GetNumberOfVerträgeSheets = sum
End Function

