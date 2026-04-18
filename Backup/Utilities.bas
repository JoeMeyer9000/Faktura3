Attribute VB_Name = "Utilities"
Option Explicit

' *****************************************************
' Module Utilities
' *****************************************************

' Dieses Modul enthält allgemein gebräuchliche Routinen, die keiner Klasse zugeordnet werden können

' Interne Variablen, die für Lazy Loading gebraucht werden
Private m_FSO As FileSystemObject
Private m_RootPath As String

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Public Function ReadBeleg(ByVal aRowIndex As Long) As clsBeleg
    Dim result As New clsBeleg
    result.Populate aRowIndex
    Set ReadBeleg = result
End Function

' =====================================================
' Datum
' =====================================================

' Liefert die Zahl des Quartals zu einem Datum
Public Function QuarterOf(ByVal datum As Date) As Integer
    QuarterOf = IIf(datum = 0, 0, (month(datum) - 1) \ 3 + 1)
End Function

' Liefert den 1. Tag des Monats
Public Function StartOfMonth(Optional ByVal dt As Date = 0) As Date
    If dt = 0 Then dt = Now
    StartOfMonth = DateSerial(Year(dt), month(dt), 1)
End Function

' Liefert [dt] plus [number] Tage. [number] darf auch negativ sein
Public Function AddDays(ByVal dt As Date, Optional ByVal number As Double = 1) As Date
    AddDays = DateAdd("d", number, dt)
End Function

' Liefert [dt] plus [number] Monate. [number] darf auch negativ sein
Public Function AddMonths(ByVal dt As Date, Optional ByVal number As Double = 1) As Date
    AddMonths = DateAdd("m", number, dt)
End Function

' Liefert [dt] plus [number] Jahre. [number] darf auch negativ sein
Public Function AddYears(ByVal dt As Date, Optional ByVal number As Double = 1) As Date
    AddYears = DateAdd("y", number, dt)
End Function

' Liefert [dt] plus [number] Wochen. [number] darf auch negativ sein
Public Function AddWeeks(ByVal dt As Date, Optional ByVal number As Double = 1) As Date
    AddWeeks = DateAdd("d", number * 7, dt)
End Function



' =====================================================
' Excel
' =====================================================

Public Function SelectedRows(Table As ListObject) As Collection
    Dim rowColl As New Collection
    Dim i As Long
    Dim j As Long
    Dim selectedArea As Range
    Dim selectedRow As Range
    Dim foundRow As Long
        
    With Application.selection
        ' alle Areas müssen geprüft werden
        For i = 1 To .Areas.count
            Set selectedArea = .Areas(i)
            ' Alle Zeilen der Area
            For j = 1 To selectedArea.Rows.count
                ' Die Tabellenzeile ist nicht identisch mit der Zeilennummer des Worksheet
                ' daher mappen wir die gefundenen Sheet-Zeile in eine Table-Zeile
                foundRow = MapRowToTable(Table, selectedArea.Rows(j).row)
                ' Wenn die Zeile außerhalb der Table liegt, sind wir fertig
                If foundRow > Table.ListRows.count Then Exit For
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
End Function

' Wandelt eine Spaltennummer in eine Buchstaben um
Public Function ColumnToStr(ByVal columnIndex As Integer) As String
    If columnIndex < 2 Then
        ColumnToStr = "A"
        Exit Function
    End If
    Dim Rtn As String
    Rtn = ""
    Dim value As Integer
    value = columnIndex - 1
    While value >= 0
        Rtn = Chr(65 + (value Mod 26)) & Rtn
        value = Fix(value / 26) - 1
    Wend
    ColumnToStr = Rtn
End Function

Public Function FindRow(tableColumn As ListColumn, findValue As Variant) As Long
    Dim result As Variant
    result = Application.Match(findValue, tableColumn.DataBodyRange, 0)
    FindRow = IIf(IsError(result), 0, result)
End Function

Public Function FindTable(Optional aSheet As Worksheet = Nothing, Optional aTableName As String = "") As ListObject
    If aSheet Is Nothing Then
        Set aSheet = ActiveSheet
    End If
    If IsMissing(aTableName) Or aTableName = "" Then
        Set FindTable = aSheet.ListObjects(1)
    Else
        Set FindTable = aSheet.ListObjects(aTableName)
    End If
End Function

' Rechnet eine Tabellenzeile in eine Worksheet-Zeile um
Public Function MapRowToSheet(Table As ListObject, ByVal rowIndex As Long) As Long
    MapRowToSheet = rowIndex + Table.HeaderRowRange.row
End Function

' Rechnet eine Worksheet-Zeile in eine Tabellenzeile um
Public Function MapRowToTable(Table As ListObject, ByVal rowIndex As Long) As Long
    MapRowToTable = rowIndex - Table.HeaderRowRange.row
End Function

Public Function IsFiltered(Table As ListObject) As Boolean
    With Table.ListColumns(1).DataBodyRange.Rows
        IsFiltered = .count <> .SpecialCells(xlCellTypeVisible).count
    End With
End Function

Public Function NextValue(tableColumn As ListColumn, Optional ByVal minNr As Long = 1)
    Dim nr As Long
    nr = WorksheetFunction.Max(tableColumn.DataBodyRange.Rows) + 1
    NextValue = IIf(nr < minNr, minNr, nr)
End Function

' Wenn [enable] = True, wird alles mögliche abgeschaltet, was Performance kostet
' Das ist aber veraltet, weil es mittlerweile FKApplication gibt, die eine unsichtbare
' Instanz der Application-Klasse erzeugt und deutlich mehr Effekt hat
Public Sub OptimizedMode(ByVal enable As Boolean, Optional app As Excel.Application = Nothing)
    ' see https://vbacompiler.com/optimize-vba-code/
    If app Is Nothing Then Set app = Application
    app.EnableEvents = Not enable
    'app.Calculation = IIf(enable, xlCalculationManual, xlCalculationAutomatic)
    app.ScreenUpdating = Not enable
    app.EnableAnimations = Not enable
    app.DisplayStatusBar = Not enable
    app.PrintCommunication = Not enable
End Sub


' =====================================================
' Files
' =====================================================

' Ändert die Namenserweiterung eines Dateinamens. Die neue Erweiterung kann den Punkt
' enthalten, muss aber nicht
Public Function ChangeExtension(ByVal Filename As String, ByVal newExtension As String) As String
    If newExtension <> "" And Left$(newExtension, 1) <> "." Then
        newExtension = "." & newExtension
    End If
    Dim oldExtension As String
    oldExtension = FSO.GetExtensionName(Filename)
    ChangeExtension = Left$(Filename, Len(Filename) - Len(oldExtension) - 1) & newExtension
End Function

' erzeugt einen kompletten Directory-Pfad, wenn es ihn nocth nicht gibt
' Die Routine kann mehrere Verzeichniseben auf einmal erzeugen
Public Sub CreateFullDirectory(strPath As String)
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

' Liefert eine Instanz des FileSystemObjekts zurück. Das Objekt wird erst dann instanziiert,
' wenn es zum ersten Mal angesprochen wird ("Lazy Loading")
' In den Verweisen muss Microsoft Scripting Library hinzugefügt werden
Public Property Get FSO() As FileSystemObject
    If m_FSO Is Nothing Then
        Set m_FSO = New FileSystemObject
    End If
    Set FSO = m_FSO
End Property

Public Function RootPath() As String
    If TESTMODE Then m_RootPath = Application.ActiveWorkbook.path
    
    If m_RootPath <> "" Then
        RootPath = m_RootPath
        Exit Function
    End If
    
    ' C:\Users\JoachimMeyer\TalkServ GmbH\Verwaltung - General\Rechnungen
    ' C:\Users\JoachimMeyer\TalkServ GmbH\Verwaltung - Dokumente\General\Rechnungen
    Dim UserProfile As String: UserProfile = Environ("UserProfile")
    
    m_RootPath = PathCombine(UserProfile, "TalkServ GmbH\Verwaltung - General\Rechnungen")
    If FSO.FolderExists(m_RootPath) Then
        RootPath = m_RootPath
        Exit Function
    End If
    
    m_RootPath = PathCombine(UserProfile, "TalkServ GmbH\Verwaltung - Dokumente\General\Rechnungen")
    If FSO.FolderExists(m_RootPath) Then
        RootPath = m_RootPath
        Exit Function
    End If
    
    RootPath = ""
    
End Function

Public Function MakeBelegPath(ByVal datum As Date) As String
    MakeBelegPath = ""
    If datum = 0 Then Exit Function
    
    Dim path As String
    path = PathCombine(RootPath, Year(datum), "Q" & QuarterOf(datum))
    If Not FSO.FolderExists(path) Then
        CreateFullDirectory (path)
    End If
    
    MakeBelegPath = path
End Function

Public Function MakeBelegFilename(ByVal datum As Date, ByVal rechId As Long, Optional extension As String = "xlsx") As String
    MakeBelegFilename = ""
    If datum = 0 Or rechId < OptionValue(xoptMinRechId) Then Exit Function
    
    Dim path As String
    path = MakeBelegPath(datum)
    If Left$(extension, 1) <> "." Then extension = "." & extension
    If path <> "" Then MakeBelegFilename = PathCombine(path, StrFormat("%1-%2%3", Year(datum) - 2000, rechId, extension))
End Function

' Diese Funktion fügt Fragmente eines Dateinamens zusammen, durch Backslash getrennt
' Die Anzahl der Fragmenbte ist beliebig. Typischer AUfruf:
' filename = PathCombine(MyWorkPath, aSubDirectory, "MyFile.xlsx")
Public Function PathCombine(ParamArray params() As Variant) As String
    Dim result As String
    Dim tmp As String
    Dim i As Long
    
    result = ""
    For i = LBound(params) To UBound(params)
        tmp = StrTrim(params(i), "\")
        If tmp <> "" Then
            If result = "" Then
                result = tmp
            Else
                result = result & "\" & tmp
            End If
        End If
    Next i
    
    PathCombine = result
End Function

' Liefert das Userverzeichnis
Public Property Get UserPath() As String
    UserPath = Environ("USERPROFILE")
End Property

' Speichert [text] in einer Datei mit dem Format UTF-8
' see https://t1p.de/04gyt
Public Sub WriteTextToUTF8File(ByVal Filename As String, ByVal Text As String)
    With CreateObject("ADODB.Stream")
        .Type = 2 ' adTypeText = 2
        .Charset = "UTF-8"
        .Open
        .WriteText Text
        .SaveToFile Filename, 2 ' adSaveCreateOverWrite = 2
        .Close
    End With
End Sub

' Speichert [text] in einer Datei
Public Sub WriteTextToFile(ByVal Filename As String, ByVal Text As String)
    Dim oFile As Object
    Set oFile = FSO.CreateTextFile(Filename)
    oFile.Write Text
    oFile.Close
End Sub

' Verzeichnis wählen. Wenn der Benutzer auf Abbrechen klickt, liefert die Funktion einen leeren String
Public Function SelectFolder(Optional ByVal Title As String = "", Optional ByVal initialFolder As String = "") As String
    SelectFolder = ""
    With Application.FileDialog(msoFileDialogFolderPicker)
        If Title <> "" Then .Title = Title
        If initialFolder <> "" Then .InitialFileName = initialFolder
        If .Show = -1 Then
            SelectFolder = .SelectedItems(1)
        End If
    End With
End Function

' 1 Datei wählen. Wenn der Benutzer auf Abbrechen klickt, liefert die Funktion einen leeren String
Public Function SelectFile(Optional ByVal Title As String = "", Optional ByVal initialFolder As String = "") As String
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

' Startet eine Datei über die Zuordnung der Datei-Erweiterung.
' xlsx öffnet die Datei in Excel, docx öffnet die Datei in Word, u.s.w.
Public Sub ShellExecute(ByVal fname As String)
    Shell StrFormat("explorer.exe ""%1""", fname), vbMaximizedFocus
End Sub


' =====================================================
' Strings
' =====================================================

Public Function SameAnschrift(anschrift1 As String, anschrift2 As String) As Boolean
    anschrift1 = Replace(anschrift1, vbLf, "|")
    anschrift1 = Replace(anschrift1, vbCr, "")
    anschrift1 = Replace(anschrift1, vbCrLf, "")
    anschrift2 = Replace(anschrift2, vbLf, "|")
    anschrift2 = Replace(anschrift2, vbCr, "")
    anschrift2 = Replace(anschrift2, vbCrLf, "")
    SameAnschrift = StrComp(anschrift1, anschrift2, vbTextCompare) = 0
End Function

' Verbindet 2 Strings getrennt, durch [separator]
' StrAppend("AAA", "BBB", ", ") liefert "AAA, BBB"
' StrAppend("AAA", "", ", ") liefert "AAA"
' StrAppend("", "BBB", ", ") liefert "BBB"
' StrAppend("", "", ", ") liefert ""
' Separator kann z. Bsp. auch vbLf sein für Zeilenumbruch innerhalb eines Feldes
Public Function StrAppend(ByVal valueA As String, ByVal valueB As String, Optional ByVal separator As String = vbLf) As String
    If valueA = "" Then
        StrAppend = valueB
    ElseIf valueB = "" Then
        StrAppend = valueA
    Else
        StrAppend = valueA & separator & valueB
    End If
End Function

' LIefert True, wenn 2 Strings gleich sind (case-insensitive)
Public Function StrEqual(ByVal s1 As String, ByVal s2 As String) As Boolean
    StrEqual = (StrComp(s1, s2, vbTextCompare) = 0)
End Function

' Ersetzt bestimmte Zeichen eines Strings in ihr HTML-Äquivalent
' Wird hauptsächlich in XRechnung benötigt
Public Function StrExpandToHTML(ByVal value As String) As String
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    value = Replace(value, vbCr, "&#13;")
    value = Replace(value, vbLf, "&#10;")
    value = Replace(value, vbCrLf, "&#13;&#10;")
    value = Replace(value, vbTab, "&tab;")
    StrExpandToHTML = value
End Function

' Setzt parameter in [formatStr] ein. Die Anzahl der Parameter ist variabel und die Funktion ist
' einigermaßen tolerant, was die Typen der Parameter angeht. Die Funktion wandelt auch
' "\t" in vbTab und "\n" in vbLf um. Beispiele:
' StrFormat("Rechnung %2 vom %1", Date, RechnungsNr) liefert "Rechnung R25-12345 vom 01.02.2025"
' StrFormat("Rechnung %2 vom %1", Date, RechnungsNr) liefert "Rechnung R25-12345 vom 01.02.2025"
Public Function StrFormat(ByVal formatStr As String, ParamArray values() As Variant) As String
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
                Err.Raise 100, "StrFormat: parameter " & i & " has an unsupported type"
        End Select
        formatStr = Replace(formatStr, "%" & CStr(i + 1), values(i))
    Next i
    formatStr = Replace(formatStr, "\t", vbTab)
    formatStr = Replace(formatStr, "\n", vbLf)
    StrFormat = formatStr
End Function

' Übernimmt eine Collection von Strings und fügt diese zuammen, getrennt durch [separator]
Public Function StrFromCollection(coll As Collection, Optional separator As String = vbCrLf, Optional ignoreEmptyItems As Boolean = False) As String
    ' assumes that collection contains only strings
    Dim result As String
    Dim item As Variant
    For Each item In coll
        If item <> "" Or Not ignoreEmptyItems Then
            If result = "" Then
                result = item
            Else
                result = result & separator & item
            End If
        End If
    Next item
    
    StrFromCollection = result
End Function

' Füllt einen String linksbündig auf. Beispiel:
' Aus StrPadLeft(CStr(125), 5, "0") wird "00123"
Public Function StrPadLeft(ByVal value As String, ByVal length As Integer, Optional padChar As String = " ") As String
    StrPadLeft = Right$(String(length, padChar) & value, length)
End Function

' Füllt einen String rechtsbündig auf. Beispiel:
' Aus StrPadRight("ABC", 6, "X") wird "ABCXXX"
Public Function StrPadRight(ByVal value As String, ByVal length As Integer, Optional padChar As String = " ") As String
    StrPadRight = Left$(value & String(length, padChar), length)
End Function

' Trennt einen String in 2 Teile auf. Beispiel:
' StrSplit("28307 Bremen", " ", sLeft, sRight) liefert sLeft="28307" und sRight="Bremen"
' Das erste Auftauchen von [separator] bestimmt, an welcher Stelle der String getrennt wird
Public Sub StrSplit(ByVal value As String, ByVal separator As String, ByRef leftValue As String, ByRef rightValue As String)
    Dim p As Integer
    p = InStr(value, separator)
    If p > 0 Then
        leftValue = Left$(value, p - 1)
        rightValue = Right$(value, Len(value) - Len(separator) - p + 1)
    Else
        leftValue = value
        rightValue = ""
    End If
End Sub

' Liefert die Position eines Textes in einem anderen Text
' Zeilenumbruch in Feldern mit vbLf und nicht vbCrLf
' StrIndexOf("Joe Meyer", "mey") liefert 5
Public Function StrIndexOf(ByVal value As String, ByVal findValue As String) As Long
    StrIndexOf = 0
    If value = "" Then Exit Function
    If findValue = "" Then Exit Function
    StrIndexOf = InStr(1, value, findValue, vbTextCompare)
End Function

' Trim() funcktioniert mit Leerzeichen, bei StrTrim() kann man das Zeichen angeben, das an
' Anfang und Ende entfernt werden soll
' TODO: bis jetzt entfernt die Funktion nur jeweils 1 Vorkommen von [remove] an Anfang un Ende
Public Function StrTrim(ByVal value As String, Optional ByVal remove As String = " ") As String
    If value = "" Or remove = "" Then
        StrTrim = value
        Exit Function
    End If
    
    Dim removeLen As Long
    removeLen = Len(remove)
    
    ' führende Zeichen entfernen
    If Len(value) >= removeLen And Left$(value, removeLen) = remove Then
        value = Right$(value, Len(value) - removeLen)
    End If
    
    ' anhängende Zeichen entfernen
    If Len(value) >= removeLen And Right$(value, removeLen) = remove Then
        value = Left$(value, Len(value) - removeLen)
    End If
    
    StrTrim = value
End Function

' =====================================================
' Message Boxes
' =====================================================

' Kleiner Helper :-)
' Ich hab immer wieder denselben Code geschrieben, um eine Bedingung zu prüfen und bei
' negativem Ergebnis eine Fehlermeldung anzuzeigen und die aufrufende Routine zu verlassen:
' Statt diesem Code:
'   If myCondition <> correct then
'       ShowError "Fehler|Das hat nicht geklappt, weil..."
'       Exit Sub
'   End If
' kann man jetzt folgendes schreiben:
'   If ExitOnError(myCondition <> correct, "Fehler|Das hat nicht geklappt, weil...") then Exit Sub
Public Function ExitOnError(ByVal errorCondition As Boolean, ByVal errorMsg As String) As Boolean
    If errorCondition Then ShowFailed errorMsg
    ExitOnError = errorCondition
End Function

' Zeigt den internen Message-Dialog an, der etwas hübscher ist als die MsgBox des Systems
' titleAndText kann einen Titel und die Message enthalten. getrennt durch ein Pipe-Symbol |
' Style gibt an, welches Icon angezeigt wird
' Buttons bestimmt, welche Buttons angezeigt werden
Public Function ShowMessage(ByVal titleAndText As String, _
            Optional ByVal style As MessageBoxStyle = MessageBoxStyle.mbsMessage, _
            Optional ByVal Buttons As MessageBoxButtons = MessageBoxButtons.mbbOk) As VbMsgBoxResult
    ShowMessage = MessageBox.ShowMessage(titleAndText, style, Buttons)
End Function

' Zeigt eine Messagebox mit Style = mbsError und mbbOk an
Public Sub ShowError(ByVal titleAndText As String)
    If InStr(1, titleAndText, "|") = 0 Then titleAndText = "Fehler|" & titleAndText
    ShowMessage titleAndText, mbsError, mbbOk
End Sub

' Zeigt eine Messagebox mit Style = mbsError und mbbOk an
Public Sub ShowAccessDenied(ByVal titleAndText As String)
    If InStr(1, titleAndText, "|") = 0 Then titleAndText = "Zugriff verweigert|" & titleAndText
    ShowMessage titleAndText, mbsError, mbbOk
End Sub

' Zeigt eine Messagebox mit Style = mbsError und mbbOk an
Public Sub ShowFailed(ByVal titleAndText As String)
    If InStr(1, titleAndText, "|") = 0 Then titleAndText = "Aktion fehlgeschlagen|" & titleAndText
    ShowMessage titleAndText, mbsError, mbbOk
End Sub

' Zeigt eine Messagebox mit Style = mbsError und mbbOk an
Public Sub ShowDevError(ByVal source As String, ByVal errorMsg As String)
    errorMsg = StrFormat("Interner Fehler|%1\n\nQuelle: %2", errorMsg, source)
    ShowMessage errorMsg, mbsError, mbbOk
End Sub

' Zeigt eine Messagebox mit Style = mbsQuestion und mbbOkCancel an
' liefert true, wenn der Benutzer Ok gewählt hat
Public Function Confirm(ByVal titleAndText As String) As Boolean
    If InStr(1, titleAndText, "|") = 0 Then titleAndText = "Bestätigung|" & titleAndText
    If InStr(1, titleAndText, "Fortsetzen") = 0 Then titleAndText = titleAndText & vbLf & vbLf & "Fortsetzen?"
    Confirm = ShowMessage(titleAndText, mbsQuestion, mbbOkCancel) = vbOK
End Function

' Zeigt eine MessageBox mit Style = mbsWarniing und Titel "Nicht implementiert" an
Public Sub NotImplemented(ByVal titleAndText As String)
    If InStr(1, titleAndText, "|") = 0 Then titleAndText = "Nicht implementiert|" & titleAndText
    Call ShowMessage(titleAndText, mbsMessage)
End Sub

' =====================================================
' Miscellaneous
' =====================================================

' Liefert die Anzahl Elemente in einem Array
Public Function ArraySize(arr As Variant, Optional level As Integer = 1) As Long
    ArraySize = UBound(arr, level) - LBound(arr, level) + 1
End Function

Public Function ArrayFrom(ParamArray values() As Variant) As Variant
    Dim result() As Variant
    ReDim result(LBound(values, 1) To UBound(values, 1))
    Dim i As Long
    For i = LBound(values, 1) To UBound(values, 1)
        result(i) = values(i)
    Next i
    ArrayFrom = result
End Function

