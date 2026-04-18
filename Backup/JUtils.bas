Attribute VB_Name = "JUtils"
Option Explicit

' *****************************************************
' Module JUtils
' *****************************************************

' Dieses Modul enthält allgemein gebräuchliche Routinen, die keiner Klasse zugeordnet werden können

' Interne Variablen, die für Lazy Loading gebraucht werden
Private m_FSO As FileSystemObject
Private m_AppPath As String

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' =====================================================
' Datum
' =====================================================

' Liefert die Zahl des Quartals zu einem Datum
Public Function QuarterOf(ByVal Datum As Date) As Integer
    QuarterOf = IIf(Datum = 0, 0, (Month(Datum) - 1) \ 3 + 1)
End Function

' Liefert den 1. Tag des Monats
Public Function StartOfMonth(Optional ByVal dt As Date = 0) As Date
    If dt = 0 Then dt = Now
    StartOfMonth = DateSerial(Year(dt), Month(dt), 1)
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

' Ermittelt die markierte Zeile aus einem Table Objekt und gibt sie in [selectedRow] zurück
' Ist keine Zelle innerhalb der Tabelle markiert, liefert die Funktion False
' Eine Zeile gilt als markiert, wenn mindestens 1 Zelle in der Zeile markiert ist
' Die Funktion gibt auch eine Fehlermeldung aus, wenn nicht genau 1 Zeile mariert ist
Public Function CollectSelectedRow(table As ListObject, selectedRow As Long) As Boolean
    ' Ich bin faul und nutze einfach CollectSelectedRows, um anschließend zu prüfen,
    ' ob auch wirklich nur genau 1 Zeile markiert ist
    Dim selectedRows As Collection
    Set selectedRows = GetSelectedRows(table)
    If selectedRows.Count <> 1 Then
        ShowMessage "Aktion unzulässig|Bitte markieren Sie genau 1 Zeile", mbsWarning
        CollectSelectedRow = False
    Else
        CollectSelectedRow = True
        selectedRow = selectedRows(1)
    End If
End Function

' Sammelt eine Collection von Zeilennummern der markierten Zeilen.
' Eine Zeile gilt als markiert, wenn mindestens 1 Zelle in der Zeile markiert ist
' Die Funktion gibt auch eine Fehlermeldung aus, wenn keine Zeile mariert ist
Public Function CollectSelectedRows(table As ListObject, selectedRows As Collection) As Boolean
    Set selectedRows = GetSelectedRows(table)
    If selectedRows.Count = 0 Then
        ShowMessage "Aktion unzulässig|Bitte markieren Sie 1 oder mehrere Zeilen", mbsWarning
        CollectSelectedRows = False
    Else
        CollectSelectedRows = True
    End If
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

' Sucht [searchValue] in Spalte [tableColumn] und liefert die Zeile der Tabelle, wenn gefunden
' Wurde die Zeile nicht gefunden, diefert die Funktion 0
Public Function FindRow(tableColumn As ListColumn, ByVal searchValue As Variant) As Long
    FindRow = 0
    Dim cell As Range
    Set cell = tableColumn.DataBodyRange.Rows.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If Not (cell Is Nothing) Then FindRow = MapRowToTable(tableColumn.Parent, cell.row)
End Function

'Public Function GetFirstEmptyRow(ByVal sheet As Worksheet, Optional columnIndex As Long = 1) As Long
'    GetFirstEmptyRow = GetLastRow(sheet, columnIndex) + 1
'End Function
'
'Public Function GetLastRow(ByVal sheet As Worksheet, Optional columnIndex As Long = 1) As Long
'    GetLastRow = sheet.Range(ColumnToStr(columnIndex) & "1000000").Rows.End(xlUp).row
'End Function

' Sucht in einer Tabellenspalte die höchste Nummer und gibt diese plus 1 zurück.
' Wurde keine Nummer gefunden, liefert die Funktion den Wert [minNr].
' Wenn die kleinste Bestellnummer also z. Bsp. 80001 ist ruft man die Funktion so
' auf: nr = GetNextID(TableGeräte, Setting(xsetBestellNr))
' Gibt es noch keine Bestellnummer, liefert die Funktion den ert aus Setting(xsetBestellNr)
Public Function GetNextID(tableColumn As ListColumn, Optional ByVal minNr As Long = 1)
    Dim nr As Long
    nr = WorksheetFunction.Max(tableColumn.DataBodyRange.Rows) + 1
    GetNextID = IIf(nr < minNr, minNr, nr)
End Function

' Markierte Zellen müssen nicht unbveding zusammenhängen sondern kännen über die gesamte
' Tabelle verteilt sein.Application.Selection enthält in diesem Fall mehrere Areas. Diese
' Funktion findet alle Zellen und liefert deren Zeilennummern in einer Collection zurück.
' Zeilennummern können nicht doppelt in der Collection vorkommen
' Eine Zeile gilt als markiert, wenn mindestens 1 Zelle in der Zeile markiert ist
Public Function GetSelectedRows(table As ListObject) As Collection
    Dim rowColl As New Collection
    Dim i As Long
    Dim j As Long
    Dim selectedArea As Range
    Dim selectedRow As Range
    Dim foundRow As Long
        
    With Application.selection
        ' alle Areas müssen geprüft werden
        For i = 1 To .Areas.Count
            Set selectedArea = .Areas(i)
            ' Alle Zeilen der Area
            For j = 1 To selectedArea.Rows.Count
                ' Die Tabellenzeile ist nicht identisch mit der Zeilennummer des Worksheet
                ' daher mappen wir die gefundenen Sheet-Zeile in eine Table-Zeile
                foundRow = MapRowToTable(table, selectedArea.Rows(j).row)
                ' Wenn die Zeile außerhalb der Table liegt, sind wir fertig
                If foundRow > table.ListRows.Count Then Exit For
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
    
    Set GetSelectedRows = rowColl
End Function

' LIefert die Zeilennummer der Überschriftenzeile einer Tabelle
Public Function GetTableHeaderRow(table As ListObject) As Long
    GetTableHeaderRow = table.ListRows(1).Range.row - 1
End Function

' LIefert True wenn die Tabelle gefiltert ist
Public Function IsTableFiltered(ByVal tbl As ListObject) As Boolean
    With tbl.ListColumns(1).DataBodyRange.Rows
        IsTableFiltered = .Count <> .SpecialCells(xlCellTypeVisible).Count
    End With
End Function

' Rechnet eine Tabellenzeile in eine Worksheet-Zeile um
Public Function MapRowToSheet(table As ListObject, ByVal rowIndex As Long)
    MapRowToSheet = rowIndex + GetTableHeaderRow(table)
End Function

' Rechnet eine Worksheet-Zeile in eine Tabellenzeile um
Public Function MapRowToTable(table As ListObject, ByVal rowIndex As Long)
    MapRowToTable = rowIndex - GetTableHeaderRow(table)
End Function

' Wenn [enable] = True, wird alles mögliche abgeschaltet, was Performance kostet
' Das ist aber veraltet, weil es mittlerweile FKApplication gibt, die eine unsichtbare
' Instanz der Application-Klasse erzeugt und deutlich mehr Effekt hat
Public Sub OptimizedMode(ByVal enable As Boolean)
    ' see https://vbacompiler.com/optimize-vba-code/
     Application.EnableEvents = Not enable
     Application.Calculation = IIf(enable, xlCalculationManual, xlCalculationAutomatic)
     Application.ScreenUpdating = Not enable
     Application.EnableAnimations = Not enable
     Application.DisplayStatusBar = Not enable
     Application.PrintCommunication = Not enable
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

Public Function FindPath() As String
    ' C:\Users\JoachimMeyer\TalkServ GmbH\Verwaltung - General\Rechnungen
    ' C:\Users\JoachimMeyer\TalkServ GmbH\Verwaltung - Dokumente\General\Rechnungen
    Dim UserProfile As String: UserProfile = Environ("UserProfile")
    
    Dim Path As String
    
    Path = PathCombine(UserProfile, "TalkServ GmbH\Verwaltung - General\Rechnungen")
    If FSO.FolderExists(Path) Then
        FindPath = Path
        Exit Function
    End If
    
    Path = PathCombine(UserProfile, "TalkServ GmbH\Verwaltung - Dokumente\General\Rechnungen")
    If FSO.FolderExists(Path) Then
        FindPath = Path
        Exit Function
    End If
    
    FindPath = ""
    
End Function

Public Property Get GetApplicationPath() As String
    GetApplicationPath = FindPath()
End Property

Public Function GetAppPath(Optional defaultRoot As String = "TalkServ GmbH\Geschäftsführung - General") As String
    GetAppPath = FindPath()
End Function

' Veraltet, nutzen Sie GetAPpPath stattdesseen
'Public Property Get GetApplicationPath() As String
'    ' COMPUTERNAME=DESKTOP-FBF5V3R
'    Dim aPath As String
'    aPath = IIf(Environ("COMPUTERNAME") = "DESKTOP-FBF5V3R", "TalkServ GmbH\Verwaltung - Dokumente\General", "TalkServ GmbH\Verwaltung - General")
'     GetApplicationPath = GetAppPath(aPath)
'End Property

' Mit OneDrive gibt es immer wieder Verzeichnis-Probleme. So liefert ThisWorkbook.Path einen seltsamen
' URL auf eine Sharepoint Website zurück.Diese Funktion versucht, Daraus ein normales Verzeichnis zu
' machen. Leider muss man einen Teil des Verzeichnis als Parameter übergeben, denn der URL generiert
' eindeutige Verzeichnisnamen und wenn man in Teams ein Team löscht und danach neu anlegt, stimmt der
' URL nicht mehr mit dem erwarteten Teams-Namen überein.
' Wenn man selbst UNterverezichniss anlegen will (zum Beispiel für Belege), kommt man um diesen Mist hier
' wohl nicht herum. Ich habe jedenfalls keine bessere Lösung gefunden.
'Public Function GetAppPath(Optional defaultRoot As String = "TalkServ GmbH\Geschäftsführung - General") As String
'    ' Lazy Loading
'    If m_AppPath <> "" Then
'        GetAppPath = m_AppPath
'        Exit Function
'    End If
'
'    ' Pfad des aktiven Workbooks ermitteln
'    Dim workbookPath As String
'    m_AppPath = Replace(ThisWorkbook.Path, "/", "\")
'    defaultRoot = Replace(defaultRoot, "/", "\")
'
'    ' wenn an 2. Stelle ein Doppelpunkt steht, ist es ein lokaler Pfad, den wir nicht bearbeiten müssen
'    If Mid$(m_AppPath, 2, 1) = ":" Then
'        GetAppPath = m_AppPath
'        Exit Function
'    End If
'
'    ' es ist ein Pfad auf OneDrive, z. Bsp.
'    ' "https:\\talkserv.sharepoint.com\sites\Geschftsfhrung2\Freigegebene Dokumente\General\Tools\Excel\TS - Rechnungen"
'    ' Wir ersetzen alles bis "General\" durch {user profile}}{[defaultRoot]}
'    Dim p As Long
'    Dim magicString As String
'
'    magicString = "General\"
'    p = InStr(1, m_AppPath, magicString, vbTextCompare)
'    m_AppPath = Right$(m_AppPath, Len(m_AppPath) - p - Len(magicString) + 1)
'
'    m_AppPath = PathCombine(UserPath, defaultRoot, m_AppPath)
'    If Right$(m_AppPath, 1) <> "\" Then m_AppPath = m_AppPath & "\"
'
'    GetAppPath = m_AppPath
'End Function
    
' Liefert denOneDrive Path aus den Umgebungsvariablen. Totally useless :-(
Public Property Get OneDrivePath() As String
    ' liefert z. Bsp. "C:\Users\{user name}\OneDrive - {company name}"
    OneDrivePath = Environ("OneDrive")
End Property

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
    Shell StrFormat("explorer.exe ""%1""", fname)
End Sub


' =====================================================
' Strings
' =====================================================

' Verbindet 2 Strings getrennt, durch [separator]
' StrAppend("AAA", "BBB", ", ") liefert "AAA, BBB"
' StrAppend("AAA", "", ", ") liefert "AAA"
' StrAppend("", "BBB", ", ") liefert "BBB"
' StrAppend("", "", ", ") liefert ""
' Separator kann z. Bsp. auch vbCrlf sein
Public Function StrAppend(ByVal valueA As String, ByVal valueB As String, Optional ByVal separator As String = vbCrLf) As String
    If valueA = "" Then
        StrAppend = valueB
    ElseIf valueB = "" Then
        StrAppend = valueA
    Else
        StrAppend = valueA & separator & valueB
    End If
End Function

Public Function StrEqualMultiLine(ByVal s1 As String, ByVal s2 As String) As Boolean
    s1 = Replace(s1, vbCr, "")
    s1 = Replace(s1, vbLf, "")
    s2 = Replace(s1, vbCr, "")
    s2 = Replace(s1, vbLf, "")
    StrEqualMultiLine = (StrComp(s1, s2, vbTextCompare) = 0)
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
' "\t" in vbZab und "\n" in vbCrLf um. Beispiele:
' StrFormat("Rechnung %2 vom %1", Date, RechnungsNr) liefert "Rechnung R25-12345 vom 01.02.2025"
' StrFormat("Rechnung %2 vom %1", Date, RechnungsNr) liefert "Rechnung R25-12345 vom 01.02.2025"
Public Function StrFormat(ByVal formatStr As String, ParamArray values() As Variant) As String
    Dim i As Long
    Dim value As Variant
    For i = LBound(values) To UBound(values)
        value = values(i)
        Select Case TypeName(values(i))
            Case "Integer", "Long", "Double"
                value = CStr(value)
            Case "Currency"
                value = Format(value, "Currency")
            Case "Date"
                value = Format(value, "ShortDate")
            Case "String"
                ' leave as is
            Case Else
                Err.Raise 100, "StrFormat: parameter " & i & " must be of type String"
        End Select
        formatStr = Replace(formatStr, "%" & CStr(i + 1), values(i))
    Next i
    formatStr = Replace(formatStr, "\t", vbTab)
    formatStr = Replace(formatStr, "\n", vbCrLf)
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
' Miscellaneous
' =====================================================

' Liefert die Anzahl Elemente in einem Array
Public Function ArraySize(arr As Variant, Optional level As Integer = 1) As Long
    ArraySize = UBound(arr, level) - LBound(arr, level) + 1
End Function

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
    If errorCondition Then ShowError errorMsg
    ExitOnError = errorCondition
End Function

' Zeigt den internen Message-Dialog an, der etwas hübscher ist als die MsgBox des Systems
' titleAndText kann einen Titel und die Message enthalten. getrennt durch ein Pipe-Symbol |
' Style gibt an, welches Icon angezeigt wird
' Buttons bestimmt, welche Buttons angezeigt werden
Public Function ShowMessage(ByVal titleAndText As String, _
            Optional ByVal Style As MessageBoxStyle = MessageBoxStyle.mbsMessage, _
            Optional ByVal Buttons As MessageBoxButtons = MessageBoxButtons.mbbOk) As VbMsgBoxResult
    ShowMessage = MessageBox.ShowMessage(titleAndText, Style, Buttons)
End Function

' Zeigt eine Messagebox mit Style = mbsError und mbbOk an
Public Sub ShowError(ByVal titleAndText As String)
    If InStr(1, titleAndText, "|") = 0 Then titleAndText = "Fehler|" & titleAndText
    ShowMessage titleAndText, mbsError, mbbOk
End Sub

' Zeigt eine Messagebox mit Style = mbsError und mbbOk an
Public Sub DevError(ByVal source As String, ByVal errorMsg As String)
    errorMsg = StrFormat("Interner Fehler|%1\n\nQuelle: %2", errorMsg, source)
    ShowMessage errorMsg, mbsError, mbbOk
End Sub

' Zeigt eine Messagebox mit Style = mbsQuestion und mbbOkCancel an
' liefert true, wenn der Benutzer Ok gewählt hat
Public Function Confirm(ByVal titleAndText As String) As Boolean
    If InStr(1, titleAndText, "|") = 0 Then titleAndText = "Bestätigung|" & titleAndText
    Confirm = ShowMessage(titleAndText, mbsQuestion, mbbOkCancel) = vbOK
End Function

' Zeigt eine MessageBox mit Style = mbsWarniing und Titel "Nicht implementiert" an
Public Sub NotImplemented(ByVal titleAndText As String)
    If InStr(1, titleAndText, "|") = 0 Then titleAndText = "Nicht implementiert|" & titleAndText
    Call ShowMessage(titleAndText, mbsWarning)
End Sub
