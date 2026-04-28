VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassKKCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================
' ClassKKCollector
' ============================================

' --------------------------------------------
' Public
' --------------------------------------------

Public Sub RunConversion(ByVal SourceFilename As String, ByVal TargetFilename As String)
    
    If Not FSO.FileExists(SourceFilename) Then Err.Raise ERR_FileNotFound, "ClassKundenDaten.Populate", "ClassKundenDaten.Populate: Source file not found"
    If FSO.FileExists(TargetFilename) Then Err.Raise ERR_FileExists, "ClassKundenDaten.Populate", "ClassKundenDaten.Populate: Target file exists already"
    
    On Error GoTo ErrorExit
    
    ' invisible application to avoid flickering on screen
    Dim App As New ClassApplication
    
    ' create new workbook
    Dim tgtBook As Workbook
    Set tgtBook = App.Instance.Workbooks.Add("X:\Faktura3\KKTemplate.xltm")
    ' TODO: get rid of hard coded template name
    
    ' create 8 more sheets
    ' Would have been easier to just create the template with 8 more sheets but sheets 2 to 8 are
    ' identical so if there would be any change it would have to be applied 9 times manually.
    ' Having only 1 sheet to apply is easy and this small code snippet copies the original sheet
    ' 8 times including the changes
    Dim i As Long
    For i = 1 To 8
        ' sheet 2 exists already in the template, take it as source for the copy
        tgtBook.Worksheets(2).Copy , tgtBook.Worksheets(i + 1)
        tgtBook.Worksheets(i + 2).Name = CStr(i + 1)
    Next i
    tgtBook.Worksheets(1).Activate
    
    ' open source workbook in R/O mode
    Dim srcBook As Workbook
    Set srcBook = App.Instance.Workbooks.Open(SourceFilename, , True)
    
    ' transfer the customer data to the new file
    TransferKunde srcBook.Worksheets(1), tgtBook.Worksheets(1)
    
    ' loop over subsequent sheets
    For i = 2 To srcBook.Worksheets.count
        TransferVerträge srcBook.Worksheets(i), tgtBook.Worksheets(i)
    Next i
    
    ' finished
    tgtBook.SaveAs TargetFilename
    
ErrorExit:
    ' make sure the books are close to free the memory occupied by the invisible application
    ' The invisible application is automagically removed from memory in the class destructor by calling Quit()
    If Not (srcBook Is Nothing) Then srcBook.Close False
    If Not (tgtBook Is Nothing) Then tgtBook.Close False
End Sub

' --------------------------------------------
' Private
' --------------------------------------------

Private Sub TransferKunde(ByVal srcSheet As Worksheet, ByVal tgtSheet As Worksheet)
    Const MaxRow = 30
    Dim srcValues As Variant
    srcValues = srcSheet.Range("A1:M" & MaxRow).Value2
    tgtSheet.Range("B4").Value = srcValues(4, 5)    ' VP Name1
    tgtSheet.Range("B5").Value = srcValues(5, 5)    ' VP Name2
    tgtSheet.Range("B6").Value = srcValues(6, 5)    ' VP Strasse
    tgtSheet.Range("B7").Value = srcValues(7, 5)    ' VP Plz/Ort
    
    tgtSheet.Range("B8").Value = srcValues(9, 5)    ' VP Tel
    tgtSheet.Range("B9").Value = srcValues(10, 5)   ' VP Email
    
    tgtSheet.Range("B17").Value = srcValues(12, 5)  ' HR Nr
    tgtSheet.Range("B18").Value = srcValues(13, 5)  ' HR Ort
    
    tgtSheet.Range("B13").Value = srcValues(18, 5)  ' Bank1
    tgtSheet.Range("B14").Value = srcValues(19, 5)  ' Bank1 BIC
    tgtSheet.Range("B12").Value = srcValues(20, 5)  ' Bank1 IBAN
    
    tgtSheet.Range("C13").Value = srcValues(18, 10) ' Bank2
    tgtSheet.Range("C14").Value = srcValues(19, 10) ' Bank2 BIC
    tgtSheet.Range("C12").Value = srcValues(20, 10) ' Bank2 IBAN
    
    tgtSheet.Range("C4").Value = srcValues(4, 10)   ' RE Name1
    tgtSheet.Range("C5").Value = srcValues(5, 10)   ' RE Name2
    tgtSheet.Range("C6").Value = srcValues(6, 10)   ' RE Strasse
    tgtSheet.Range("C7").Value = srcValues(7, 10)   ' RE Plz/Ort
    
    tgtSheet.Range("B22").Value = srcValues(4, 3)   ' Rechnung erhalten
    tgtSheet.Range("B23").Value = srcValues(5, 3)   ' Analyse zugeschickt
    tgtSheet.Range("B25").Value = srcValues(6, 3)   ' Kartenanzahl
    tgtSheet.Range("B24").Value = srcValues(7, 3)   ' Kunde seit
    tgtSheet.Range("B26").Value = srcValues(8, 3)   ' Vertriebler
    tgtSheet.Range("B27").Value = srcValues(11, 3)  ' Kundenkennwort
    
    ' Ansprechpartner
    Dim aspRow As Long
    Dim tgtCol As Long: tgtCol = 5
    Dim hasAnyValue As Boolean
    For aspRow = 24 To MaxRow
        hasAnyValue = srcValues(aspRow, 3) <> "" Or srcValues(aspRow, 4) <> "" Or srcValues(aspRow, 5) <> "" Or _
                    srcValues(aspRow, 6) <> "" Or srcValues(aspRow, 7) <> "" Or srcValues(aspRow, 8) <> "" Or _
                    srcValues(aspRow, 10) <> ""
        If hasAnyValue Then
            tgtSheet.Cells(4, tgtCol).Value = srcValues(aspRow, 3)   ' Anrede
            tgtSheet.Cells(5, tgtCol).Value = srcValues(aspRow, 4)   ' Nachname
            tgtSheet.Cells(6, tgtCol).Value = srcValues(aspRow, 5)   ' Vorname
            tgtSheet.Cells(7, tgtCol).Value = srcValues(aspRow, 6)   ' Titel
            tgtSheet.Cells(8, tgtCol).Value = srcValues(aspRow, 7)   ' Durchwahl
            tgtSheet.Cells(10, tgtCol).Value = srcValues(aspRow, 8)  ' Email
            tgtSheet.Cells(14, tgtCol).Value = srcValues(aspRow, 10) ' Kommentar
            tgtCol = tgtCol + 1
            If tgtCol > 8 Then Err.Raise ERR_Failed, "TransferKunde()", "TransferKunde: mehr als 4 Ansprechpartner"
        End If
    Next aspRow
End Sub

Private Sub TransferVerträge(ByVal srcSheet As Worksheet, ByVal tgtSheet As Worksheet)
    If Not StrStartsWith(srcSheet.Range("A1").Value2, "Verträge von Kunde") Then Exit Sub
End Sub

