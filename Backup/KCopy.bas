Attribute VB_Name = "KCopy"
Option Explicit

' ==================================================
' KKCopy
' ==================================================

Public Sub ListFiles(ByVal subDir As String)
    ' get the subdirectory that the user added
    Dim path As String
    Dim OneDrivePath As String
    shFiles.Range("SubDirectory").value = UCase$(subDir)
    
    ' make it a path and exit if the path doesn't exist
    OneDrivePath = FindOnedrivePath("TalkServ GmbH", "Verwaltung", "Kartei\Kunden")
    path = PathCombine(OneDrivePath, subDir)
    If ExitOnError(Not FSO.FolderExists(path), "Das Verzeichnis exitiert nicht:" & vbCrLf & path) Then Exit Sub


    ' clear the table and refill it
    Dim files As New Collection
    If Not TFiles.DataBodyRange Is Nothing Then TFiles.DataBodyRange.Delete
    CollectFiles files, path

    ' in case something goes wrong make sure the optimize settings are restored
    On Error GoTo ErrorExit
    OptimizedMode True
    
    ' write the found files into the table
    With TFiles
        Dim fileObj As Object
        Dim newRow As ListRow
        For Each fileObj In files
            Set newRow = .ListRows.Add
            .DataBodyRange(newRow.index, 1) = fileObj.Name
            .DataBodyRange(newRow.index, 2) = fileObj.DateLastModified
            .DataBodyRange(newRow.index, 3) = fileObj.Size
            .DataBodyRange(newRow.index, 5) = fileObj.path
        Next fileObj
    End With
    
ErrorExit:
    OptimizedMode False

End Sub

Public Sub CopyToNewFormat(ByVal aFilename As String)
    If ExitOnError(Not FSO.FileExists(aFilename), "File not found|" & aFilename) Then Exit Sub
    
    ' overwrite prompt
    Dim newFilename As String
    newFilename = Replace(aFilename, "Karteikarte", "KK")
    If FSO.FileExists(newFilename) Then
        If Ask("Karteikarte existiert|Die Karteikarte im neuen Format existiert bereits, möchte Sie diese überschreiben?") <> vbOK Then Exit Sub
    End If
    
    On Error GoTo ErrorExit
    Dim app As New ClassApplication
'    Dim app As Application
'    Set app = Application
    
    ' create new workbook
    Dim tgtBook As Workbook
    Set tgtBook = app.Instance.Workbooks.Add("X:\Faktura3\KKTemplate.xltm")
    
    ' create 8 more sheets
    ' Would have been easier to just create the template with 8 more sheets but
    ' sheets 2 to 8 are identical so if there would be any change it would have
    ' to be applied 9 times manually. Having only 1 sheet to apply is easy and
    ' this small code snippet copies the original sheet 8 times including the
    ' changes
    Dim i As Long
    For i = 1 To 8
        tgtBook.Worksheets(2).Copy , tgtBook.Worksheets(i + 1)
        tgtBook.Worksheets(i + 2).Name = CStr(i + 1)
    Next i
    tgtBook.Worksheets(1).Activate
    
    ' open source workbook
    Dim srcBook As Workbook
    Set srcBook = app.Instance.Workbooks.Open(aFilename, , True)
    
    TransferKunde srcBook.Worksheets(1), tgtBook.Worksheets(1)
    For i = 2 To srcBook.Worksheets.count
        TransferVerträge srcBook.Worksheets(i), tgtBook.Worksheets(i)
    Next i
    ' save new workbook
    tgtBook.SaveAs newFilename
ErrorExit:
    If Not (srcBook Is Nothing) Then srcBook.Close False
    If Not (tgtBook Is Nothing) Then tgtBook.Close False
End Sub

' ==================================================
' private
' ==================================================

Private Sub TransferKunde(ByVal srcSheet As Worksheet, ByVal tgtSheet As Worksheet)
    Const MaxRow = 30
    Dim srcValues As Variant
    srcValues = srcSheet.Range("A1:M" & MaxRow).Value2
    tgtSheet.Range("B4").value = srcValues(4, 5)    ' VP Name1
    tgtSheet.Range("B5").value = srcValues(5, 5)    ' VP Name2
    tgtSheet.Range("B6").value = srcValues(6, 5)    ' VP Strasse
    tgtSheet.Range("B7").value = srcValues(7, 5)    ' VP Plz/Ort
    
    tgtSheet.Range("B8").value = srcValues(9, 5)    ' VP Tel
    tgtSheet.Range("B9").value = srcValues(10, 5)   ' VP Email
    
    tgtSheet.Range("B17").value = srcValues(12, 5)  ' HR Nr
    tgtSheet.Range("B18").value = srcValues(13, 5)  ' HR Ort
    
    tgtSheet.Range("B13").value = srcValues(18, 5)  ' Bank1
    tgtSheet.Range("B14").value = srcValues(19, 5)  ' Bank1 BIC
    tgtSheet.Range("B12").value = srcValues(20, 5)  ' Bank1 IBAN
    
    tgtSheet.Range("C13").value = srcValues(18, 10) ' Bank2
    tgtSheet.Range("C14").value = srcValues(19, 10) ' Bank2 BIC
    tgtSheet.Range("C12").value = srcValues(20, 10) ' Bank2 IBAN
    
    tgtSheet.Range("C4").value = srcValues(4, 10)   ' RE Name1
    tgtSheet.Range("C5").value = srcValues(5, 10)   ' RE Name2
    tgtSheet.Range("C6").value = srcValues(6, 10)   ' RE STrasse
    tgtSheet.Range("C7").value = srcValues(7, 10)   ' RE Plz/Ort
    
    tgtSheet.Range("B22").value = srcValues(4, 3)   ' Rechnung erhalten
    tgtSheet.Range("B23").value = srcValues(5, 3)   ' Analyse zugeschickt
    tgtSheet.Range("B25").value = srcValues(6, 3)   ' Kartenanzahl
    tgtSheet.Range("B24").value = srcValues(7, 3)   ' Kunde seit
    tgtSheet.Range("B26").value = srcValues(8, 3)   ' Vertriebler
    tgtSheet.Range("B27").value = srcValues(11, 3)  ' Kundenkennwort
    
    ' Ansprechpartner
    Dim aspRow As Long
    Dim tgtCol As Long: tgtCol = 5
    Dim hasAnyValue As Boolean
    For aspRow = 24 To MaxRow
        hasAnyValue = srcValues(aspRow, 3) <> "" Or srcValues(aspRow, 4) <> "" Or srcValues(aspRow, 5) <> "" Or _
                    srcValues(aspRow, 6) <> "" Or srcValues(aspRow, 7) <> "" Or srcValues(aspRow, 8) <> "" Or _
                    srcValues(aspRow, 10) <> ""
        If hasAnyValue Then
            tgtSheet.Cells(4, tgtCol).value = srcValues(aspRow, 3)   ' Anrede
            tgtSheet.Cells(5, tgtCol).value = srcValues(aspRow, 4)   ' Nachname
            tgtSheet.Cells(6, tgtCol).value = srcValues(aspRow, 5)   ' Vorname
            tgtSheet.Cells(7, tgtCol).value = srcValues(aspRow, 6)   ' Titel
            tgtSheet.Cells(8, tgtCol).value = srcValues(aspRow, 7)   ' Durchwahl
            tgtSheet.Cells(10, tgtCol).value = srcValues(aspRow, 8)  ' Email
            tgtSheet.Cells(14, tgtCol).value = srcValues(aspRow, 10) ' Kommentar
            tgtCol = tgtCol + 1
            If tgtCol > 8 Then Err.Raise ERR_Failed, "TransferKunde()", "TransferKunde: mehr als 4 Ansprechpartner"
        End If
    Next aspRow
End Sub

Private Sub TransferVerträge(ByVal srcSheet As Worksheet, ByVal tgtSheet As Worksheet)
    If Not StrStartsWith(srcSheet.Range("A1").Value2, "Verträge von Kunde") Then Exit Sub
End Sub

Private Sub CollectFiles(ByVal fileObjects As Collection, ByVal path As String)
    ' returns a list of FileSystemObject.FileObject
    ' see https://t1p.de/w63fn
    
    ' Property          Description
    ' ----------------  ------------------------------------------------------------------------
    ' Attributes        Sets or returns the attributes of a specified file.
    ' DateCreated       Returns the date and time when a specified file was created.
    ' DateLastAccessed  Returns the date and time when a specified file was last accessed.
    ' DateLastModified  Returns the date and time when a specified file was last modified.
    ' Drive             Returns the drive letter of the drive where a specified file or folder resides.
    ' Name              Sets or returns the name of a specified file.
    ' ParentFolder      Returns the folder object for the parent of the specified file.
    ' Path              Returns the path for a specified file.
    ' ShortName         Returns the short name of a specified file (the 8.3 naming convention).
    ' ShortPath         Returns the short path of a specified file (the 8.3 naming convention).
    ' Size              Returns the size, in bytes, of a specified file.
    ' Type              Returns the type of a specified file.

    ' get the data of the current folder
    Dim folder As Object
    Set folder = FSO.GetFolder(path)
    
    ' loop over all files
    Dim fileObj As Object
    For Each fileObj In folder.files
        If StrStartsWith(fileObj.Name, "Karteikarte") And Not StrContains(fileObj.path, "(alt)", "Kein Bedarf", "Dummy", "Datensätze") Then
            fileObjects.Add fileObj
        End If
    Next fileObj
    
    ' loop over subdirectories recursively
    For Each fileObj In folder.SubFolders
        If Not StrContains(fileObj.path, "(alt)", "Kein Bedarf", "Dummy", "Datensätze") Then
            CollectFiles fileObjects, PathCombine(path, fileObj.Name)
        End If
    Next fileObj
End Sub
