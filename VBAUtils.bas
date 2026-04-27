Attribute VB_Name = "VBAUtils"
Option Explicit

' ==================================================
' Module VBAUtils
' ==================================================

' This module contains a couple of subs and functions that are
' commonly used in one place

Public Const TESTMODE = False

Public Const ERR_Offset = vbObjectError + 100

Public Const ERR_OutOfRange = ERR_Offset + 0
Public Const ERR_FormatError = ERR_Offset + 1
Public Const ERR_SyntaxError = ERR_Offset + 2
Public Const ERR_UnsupportedItem = ERR_Offset + 3
Public Const ERR_NotFound = ERR_Offset + 4
Public Const ERR_MissingParameter = ERR_Offset + 5
Public Const ERR_InvalidParameter = ERR_Offset + 6
Public Const ERR_Failed = ERR_Offset + 7
Public Const ERR_AccessDenied = ERR_Offset + 8
Public Const ERR_FileNotFound = ERR_Offset + 9
Public Const ERR_FileExists = ERR_Offset + 10
Public Const ERR_PathNotFound = ERR_Offset + 11

' m_FSO is used in property FSO
Private m_FSO As FileSystemObject

' --------------------------------------------------
' Message
' --------------------------------------------------

Public Function ExitOnError(ByVal errorCondition As Boolean, Optional ByVal errorMessage As String = "Fehler|Aktion fehlgeschlagen") As Boolean
    ' helper to show error and return false. Use like this:
    ' if ExitOnError(myErrorCondition, "It's an error") then Exit Sub
    ExitOnError = errorCondition
    If errorCondition Then ShowError (errorMessage)
End Function

Public Sub ShowNotImplemented(Optional titleAndMessage As String = "Not Implemented|Diese Funktion ist noch nicht implementiert")
    ' shows "Not Implemented" message, optionally with your on text
    If Not (InStr(1, titleAndMessage, "|") > 0) Then titleAndMessage = "Not Implemented|" & titleAndMessage
    MessageBox.ShowMessage titleAndMessage, mbsWarning
End Sub

Public Sub ShowError(ByVal titleAndMessage As String)
    ' shows an error box
    If Not (InStr(1, titleAndMessage, "|") > 0) Then titleAndMessage = "Fehler|" & titleAndMessage
    MessageBox.ShowMessage titleAndMessage, mbsError
End Sub

Public Sub ShowHint(titleAndMessage As String)
    ' shows a box with a hint text
    If Not (InStr(1, titleAndMessage, "|") > 0) Then titleAndMessage = "Hinweis|" & titleAndMessage
    MessageBox.ShowMessage titleAndMessage, mbsMessage
End Sub

Public Function Ask(titleAndMessage As String, Optional ByVal msgButtons As MessageBoxButtons = mbbOkCancel) As VbMsgBoxResult
    ' shows a message box with a question
    If Not (InStr(1, titleAndMessage, "|") > 0) Then titleAndMessage = "Bestätigung|" & titleAndMessage
    Ask = MessageBox.ShowMessage(titleAndMessage, mbsMessage, msgButtons)
End Function

' -----------------------------------------------
' Arrays
' -----------------------------------------------

Public Property Get ArrayWidth(ByRef TwoDimArray As Variant) As Long
    ArrayWidth = ArraySize(TwoDimArray, 2)
End Property

Public Property Get ArrayHeight(ByRef TwoDimArray As Variant) As Long
    ArrayHeight = ArraySize(TwoDimArray, 1)
End Property

Public Property Get ArraySize(ByRef aArray As Variant, ByVal aDimension As Byte) As Long
    If IsEmpty(aArray) Then
        ArraySize = 0
    Else
        ArraySize = UBound(aArray, aDimension) - LBound(aArray, aDimension) + 1
    End If
End Property

' --------------------------------------------------
' Excel
' --------------------------------------------------

Public Sub ConvertA1ToR1C1(ByVal aAddress As String, ByRef aRow As Long, ByRef aColumn As Long)
    With ActiveSheet
        aColumn = .Range(aAddress).column
        aRow = .Range(aAddress).row
    End With
End Sub

Public Function ConvertR1C1ToA1(ByVal aRow As Long, ByVal aColumn As Long) As String
    ConvertR1C1ToA1 = ActiveSheet.Cells(aRow, aColumn)
End Function

Public Function GetWorkbookDirectory(Optional aWorkbook As Workbook = Nothing) As String
    If aWorkbook Is Nothing Then
        GetWorkbookDirectory = Application.ActiveWorkbook.path
    Else
        GetWorkbookDirectory = aWorkbook.path
    End If
End Function

Public Function GetWorkbookFilename(Optional aWorkbook As Workbook = Nothing) As String
    If aWorkbook Is Nothing Then
        GetWorkbookFilename = Application.ActiveWorkbook.Fullname
    Else
        GetWorkbookFilename = aWorkbook.Fullname
    End If
End Function

Public Function FindRowIndex(ByVal tableColumn As ListColumn, ByVal searchTerm As Variant) As Long
    ' searches a list column for a search term and returns the row nummber (or 0 if not found)
    Dim result As Long
    result = Application.Match(searchTerm, tableColumn.DataBodyRange, 0)
    If IsError(result) Then result = 0
    FindRowIndex = result
End Function

Public Sub OptimizedMode(ByVal enable As Boolean, Optional ByVal aApp As Application = Nothing)
    ' In optimized mode several time costly things are switched off
    ' This is obsolete since we have ClassAplication which is way more efficient
    ' see also: https://vbacompiler.com/optimize-vba-code/
    If aApp Is Nothing Then Set aApp = Application
    aApp.EnableEvents = Not enable
    'aApp.Calculation = IIf(enable, xlCalculationManual, xlCalculationAutomatic)
    aApp.ScreenUpdating = Not enable
    aApp.EnableAnimations = Not enable
    aApp.DisplayStatusBar = Not enable
    aApp.PrintCommunication = Not enable
End Sub

Public Function CollectSelectedRows(Table As ListObject, Optional multiSelect As Boolean = False) As Collection
    ' returns a list of indexes of the selected rows
    ' an error is displayed if the number of selected rows does not match the multiSelect parameter
    Dim result As Collection
    Set result = SelectedRows(Table)
    
    If result.count = 0 Then
        ShowHint StrFormat("Bitte markieren Sie %1 Zelle in der gewünschten Zeile", IIf(multiSelect, "mindestens 1", "genau 1"))
    Else
        If Not multiSelect And result.count > 1 Then ShowHint "Bitte markieren Sie nur eine einzige Zeile"
    End If
    
    Set CollectSelectedRows = result
End Function

Private Property Get SelectedRows(aTable As ListObject) As Collection
    ' returns a list of indexes of the selected rows
    Dim rowColl As New Collection
    Dim i As Long
    Dim j As Long
    Dim selectedArea As Range
    Dim selectedRow As Range
    Dim foundRow As Long
        
    With Application.Selection
        ' alle Areas müssen geprüft werden
        For i = 1 To .Areas.count
            Set selectedArea = .Areas(i)
            ' Alle Zeilen der Area
            For j = 1 To selectedArea.Rows.count
                ' Die Tabellenzeile ist nicht identisch mit der Zeilennummer des Worksheet
                ' daher mappen wir die gefundenen Sheet-Zeile in eine Table-Zeile
                foundRow = MapRowToTable(aTable, selectedArea.Rows(j).row)
                ' Wenn die Zeile außerhalb der Table liegt, sind wir fertig
                If foundRow > aTable.ListRows.count Then Exit For
                If foundRow > 0 Then
                    ' Collection.Add liefert einen Fehler, wenn der Key bereits exitiert.
                    ' Dies nutzen wir indem wir die Fehlermeldung ignorieren so dass der
                    ' Mechanismus Duplikate selbst aussortiert
                    On Error Resume Next
                    rowColl.Add foundRow, CStr(foundRow)
                    On Error GoTo 0
                End If
            Next j
        
        Next i
        
    End With
    
    Set SelectedRows = rowColl
End Property

Public Function MapRowToSheet(aTable As ListObject, ByVal rowIndex As Long) As Long
    ' takes a table's row index and returns the according screen row
    MapRowToSheet = rowIndex + aTable.HeaderRowRange.row
End Function

Public Function MapRowToTable(aTable As ListObject, ByVal rowIndex As Long) As Long
    ' takes a screen's row index and returns the according table row
    ' If you have to use this, you're probablky doing something wrong because
    ' all data resided in ListObjects and there's usually only 1 table object per sheet
    MapRowToTable = rowIndex - aTable.HeaderRowRange.row
End Function

Public Function IsFiltered(aTable As ListObject) As Boolean
    ' returns true if a ListObject's rows are filtered
    With aTable.ListColumns(1).DataBodyRange.Rows
        IsFiltered = .count <> .SpecialCells(xlCellTypeVisible).count
    End With
End Function

Public Function NextId(tableColumn As ListColumn, Optional minId As Long = 1)
    ' returns the next highest number in the ListColumn
    ' This is good for list columns with auto incrementing numbers
    Dim newId As Long
    newId = WorksheetFunction.Max(tableColumn.DataBodyRange.Rows) + 1
    NextId = IIf(newId < minId, minId, newId)
End Function

' --------------------------------------------------
' Dates
' --------------------------------------------------

Public Function QuarterOf(ByVal Datum As Date) As Integer
    ' returns the quarter of a date as a number from 1 to 4
    QuarterOf = IIf(Datum = 0, 0, (Month(Datum) - 1) \ 3 + 1)
End Function

Public Function StartOfMonth(Optional ByVal dt As Date = 0) As Date
    ' returns the 1st day in the month of the given date
    ' omit the date parameter to get the start of the current month
    If dt = 0 Then dt = Now
    StartOfMonth = DateSerial(Year(dt), Month(dt), 1)
End Function

' --------------------------------------------------
' Files
' --------------------------------------------------

Public Property Get FSO() As FileSystemObject
    ' On 1st call this function creates the FileSystemObject and stores it for
    ' later use. Subsequent calls use the previously stored FileSystemObject.
    If m_FSO Is Nothing Then Set m_FSO = New FileSystemObject
    Set FSO = m_FSO
End Property

Public Property Get UserProfilePath() As String
    ' returns the current user's home directory, like C:\UsersJoachimMeyer
    UserProfilePath = PathCombine("C:\users", Environ("Username"))
End Property

Public Function FindOnedrivePath(ByVal organization As String, ByVal Site As String, Optional ByVal subDirectory As String = "") As String
    ' Finds a path on OneDrive using UserProfilePath, organization, channel and subdirectory
    ' Example:  FindOnedrivePath("TalkServ GmbH", "Verwaltung","Rechnungen") finds
    '       C:\Users\JoachimMeyer\TalkServ GmbH\Verwaltung - Dokumente\General\Rechnungen\
    '       or C:\Users\JoachimMeyer\TalkServ GmbH\Verwaltung - General\Rechnungen\
    FindOnedrivePath = ""
    
    If organization = "" Then Err.Raise ERR_MissingParameter, "RootPath", "RootPath: missing parameter 'organization'"
    If Site = "" Then Err.Raise ERR_MissingParameter, "RootPath", "RootPath: missing parameter 'Site'"
    
    Dim path As String
    If TESTMODE Then
        path = Application.ActiveWorkbook.path
        If subDirectory <> "" Then path = PathCombine(path, subDirectory)
        FindOnedrivePath = AddBackslash(path)
        Exit Function
    End If
    
    path = PathCombine(UserProfilePath, organization, Site & " - General")
    If Not FSO.FolderExists(path) Then
        path = PathCombine(UserProfilePath, organization, Site & " - Dokumente\General")
        If Not FSO.FolderExists(path) Then
            Err.Raise ERR_PathNotFound, "RootPath()", "RootPath(): Verzeichnis nicht gefunden"
        End If
    End If
    
    If subDirectory <> "" Then path = PathCombine(path, subDirectory)
    FindOnedrivePath = AddBackslash(path)
    
End Function

Public Function PathCombine(ParamArray parts() As Variant) As String
    ' takes a list of strings and combines them into a path by adding backslashes where needed
    ' PathCombine("UserProfilePath", "MyDir", "MySubDur") creates C:\Users\JoachimMeyer\MyDir\MySubDir\
    Dim result As String: result = ""
    Dim part As Variant
    For Each part In parts
        part = CStr(part)
        If Left$(part, 1) = "\" Then part = Right$(part, Len(part) - 1)
        If Len(part) > 0 Then
            If result = "" Then
                result = part
            Else
                result = AddBackslash(result) & part
            End If
        End If
    Next part
    PathCombine = result
End Function

Public Function AddBackslash(ByVal folder As String) As String
    ' Adds a backslash to the end of the string if it doesn't already exist
    AddBackslash = folder
    If folder = "" Then Exit Function
    If Right$(folder, 1) <> "\" Then AddBackslash = folder & "\"
End Function

Public Function ChangeExtension(ByVal filename As String, ByVal newExtension As String) As String
    ' filename cannot be empty
    If filename = "" Then Err.Raise ERR_MissingParameter, "ChangeExtension", "ChangeExtension: missing filename"
    
    ' newExtension cannot be empty
    If newExtension = "" Then Err.Raise ERR_MissingParameter, "ChangeExtension", "ChangeExtension: missing new extension"
    If Left$(newExtension, 1) = "." Then newExtension = Right$(newExtension, Len(newExtension) - 1)
    
    ' find last point in filename
    Dim p As Long
    p = InStrRev(filename, ".")
    If p = 0 Then Err.Raise ERR_FormatError, "ChangeExtension", "ChangeExtension: no extension in filename"
    
    ' new filename
    ChangeExtension = Left$(filename, p) & newExtension
End Function

Public Sub CreateFullDirectory(strPath As String)
    ' Creates a directory structure if it doesn't already exist
    ' Die Routine kann mehrere Verzeichniseben auf einmal erzeugen
    Dim part As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    For Each part In Split(strPath, "\")
        strCheckPath = strCheckPath & part & "\"
        If InStr(1, part, ":") = 0 Then
            If dir(strCheckPath, vbDirectory) = "" Then
                MkDir strCheckPath
            End If
        End If
    Next
End Sub

Public Sub WriteTextToUTF8File(ByVal filename As String, ByVal text As String)
    ' Creates a file and writes text to it. The file will have UTF-8 format which is required for XRechnung
    ' see https://t1p.de/04gyt
    With CreateObject("ADODB.Stream")
        .Type = 2 ' adTypeText = 2
        .Charset = "UTF-8"
        .Open
        .WriteText text
        .SaveToFile filename, 2 ' adSaveCreateOverWrite = 2
        .Close
    End With
End Sub

Public Sub WriteTextToFile(ByVal filename As String, ByVal text As String)
    ' Creates a file and writes text to it
    Dim oFile As Object
    Set oFile = FSO.CreateTextFile(filename)
    oFile.Write text
    oFile.Close
End Sub

Public Function SelectFolder(Optional ByVal Title As String = "", Optional ByVal initialFolder As String = "") As String
    ' Returns the folder that the user selected or an empty string if the user clicked Cancel
    SelectFolder = ""
    With Application.FileDialog(msoFileDialogFolderPicker)
        If Title <> "" Then .Title = Title
        If initialFolder <> "" Then .InitialFileName = initialFolder
        If .Show = -1 Then
            SelectFolder = .SelectedItems(1)
        End If
    End With
End Function

Public Function SelectFile(Optional ByVal Title As String = "", Optional ByVal initialFolder As String = "") As String
    ' Returns the name of the file that the user selected or an empty string if the user clicked Cancel
    SelectFile = ""
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = IIf(Title <> "", Title, "Datei wählen")
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel-Dateien", "*.xls;*.xlsx"
        .Filters.Add "Alle Dateien", "*.*"
        If initialFolder <> "" Then .InitialFileName = initialFolder
        If .Show = True Then
            SelectFile = .SelectedItems(1)
        End If
    End With
End Function

Public Sub ShellExecute(ByVal fname As String)
    ' Launches a file based on the file extension
    Shell StrFormat("explorer.exe ""%1""", fname)
End Sub

Public Sub SplitOnedrivePath(ByVal value, ByRef organization, ByRef siteName As String, ByRef subDirectory As String)
    ' splits the following paths into their parts to extract Teams channel ("siteName")
    ' and directory part.
    ' For exampe, any of the following:
    '   C:\Users\JoachimMeyer\TalkServ GmbH\Verwaltung - General\Kartei\Kunden\A\ABPE
    '   C:\Users\JoachimMeyer\TalkServ GmbH\Verwaltung - Dokumente\General\Kartei\Kunden\A\ABPE
    '   https://talkserv.sharepoint.com/sites/Verwaltung/Freigegebene Dokumente/General/Kartei/Kunden/A/ABPE
    ' will be split into
    '   siteName = "Verwaltung"
    '   subDirectory = "Kartei\Kunden\A\ABPE\"
    
    siteName = ""
    subDirectory = ""
    If value = "" Then Exit Sub
    
    Dim isHTTP As Boolean
    isHTTP = StrStartsWith(value, "https:")
    If Not isHTTP Then value = Replace(value, "\", "/")
    value = Replace(value, "//", "/")
    
    Dim parts() As String
    parts = Split(value, "/")
    
    If Not StrStartsWith(value, "https:") Then organization = parts(3)
    
    Dim state As Integer: state = 0
    Dim p As Integer
    Dim i As Integer
    Dim part As String
    For i = LBound(parts) To UBound(parts)
        part = parts(i)
        Select Case state
            Case 0 ' found nothing yet
                If (Not isHTTP) And (i = 3) Then
                    organization = part
                    state = 1 ' next action:
                ElseIf StrEqual(part, "sites") Then
                    state = 3
                ElseIf StrEqual(part, "General") Then
                    state = 4
                ElseIf StrContains(part, ".sharepoint.com") Then
                    organization = Left$(part, Len(part) - Len(".sharepoint.com"))
                End If
            Case 1
                ' Position is on site name
                p = InStr(1, part, " - ")
                If p > 0 Then
                    siteName = Left$(part, p - 1)
                Else
                    siteName = part
                End If
                state = 4
            Case 3
                siteName = part
                state = 0
            Case 4
                p = InStr(1, value, "General")
                If Not (p > 0) Then Err.Raise ERR_FormatError, "SplitOnedrivePath", "SplitOnedrivePath: invalid path"
                subDirectory = Mid$(value, p + Len("General") + 1)
                Exit Sub
        End Select
    Next i
    
    subDirectory = AddBackslash(subDirectory)
End Sub

' --------------------------------------------------
' Strings
' --------------------------------------------------

Public Function StrFormat(ByVal formatStr As String, ParamArray values() As Variant) As String
    ' Setzt parameter in [formatStr] ein. Die Anzahl der Parameter ist variabel und die Funktion ist
    ' einigermaßen tolerant, was die Typen der Parameter angeht. Die Funktion wandelt auch
    ' "\t" in vbTab und "\n" in vbLf um. Beispiele:
    ' StrFormat("Rechnung %2 vom %1", Date, RechnungsNr) liefert "Rechnung R25-12345 vom 01.02.2025"
    ' StrFormat("Rechnung %2 vom %1", Date, RechnungsNr) liefert "Rechnung R25-12345 vom 01.02.2025"
    Dim i As Long
    Dim value As Variant
    For i = LBound(values) To UBound(values)
        value = values(i)
        Select Case TypeName(value)
            Case "Integer", "Long", "Double"
                value = CStr(value)
            Case "Currency"
                value = Format(value, "Currency")
            Case "Date"
                value = Format(value, "ShortDate")
            Case "String"
                ' leave as is
            Case Else
                Err.Raise ERR_FormatError, "StrFormat: parameter " & i & " has an unsupported type"
        End Select
        formatStr = Replace(formatStr, "%" & CStr(i + 1), values(i))
    Next i
    formatStr = Replace(formatStr, "\t", vbTab)
    formatStr = Replace(formatStr, "\n", vbLf)
    StrFormat = formatStr
End Function

Public Function StrEqual(ByVal s1 As String, ByVal s2 As String) As Boolean
    ' compares 2 strings and returns ture if they are equal.
    ' comparison is case-insensitive and ignores vbCr and vbLF
    s1 = Replace(s1, vbCr, "")
    s1 = Replace(s1, vbLf, "")
    s2 = Replace(s2, vbCr, "")
    s2 = Replace(s2, vbLf, "")
    StrEqual = (StrComp(s1, s2, vbTextCompare) = 0)
End Function

Public Function StrAppend(ByVal str1 As String, ByVal str2 As String, Optional separator As String = vbLf) As String
    ' concatenates 2 strings, separated by a separator
    If str1 = "" Then
        StrAppend = str2
    ElseIf str2 = "" Then
        StrAppend = str1
    Else
        StrAppend = str1 & separator & str2
    End If
End Function

Public Function StrStartsWith(ByVal value As String, ByVal startValue As String) As Boolean
    ' returns true if value starts with startValue, case-insensitive
    StrStartsWith = StrComp(Left$(value, Len(startValue)), startValue, vbTextCompare) = 0
End Function

Public Function StrContains(ByVal value As String, ParamArray values() As Variant) As Boolean
    ' checks if value contains any of the string in values. Case-insensitive
    Dim s As Variant
    StrContains = True
    For Each s In values
        If InStr(1, value, s, vbTextCompare) > 0 Then Exit Function
    Next s
    StrContains = False
End Function

