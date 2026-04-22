VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassBelegDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==================================================
' ClassBelegDoc
' ==================================================

Public Bearbeiter As String
Public Brutto As Currency
Public Email As String
Public KdNr As Double
Public Kunde As String
Public LieferDatum As Date
Public Nachspann1 As String
Public Nachspann2 As String
Public Nachspann3 As String
Public Nachspann4 As String
Public Nachspann5 As String
Public Netto As Currency
Public RCText As String
Public ReDatum As Date
Public ReNr As String
Public Steuer As Currency
Public Steuersatz As Double

Public Positionen As Variant

Private Const GUTSCHRIFT_TEXT = "Diese Gutschrift ersetzt Rechnung"

Public Property Get Kind() As BelegKind
    Kind = xbkVoid
    
    If ReNr <> "" Then
        Select Case Left$(ReNr, 1)
            Case "R"
                Kind = xbkRechnung
            Case "G"
                Kind = IIf(Left$(Nachspann1, Len(GUTSCHRIFT_TEXT)) = GUTSCHRIFT_TEXT, xbkStornoGutschrift, xbkGutschrift)
        End Select
    End If
End Property

Public Function CalcNetto() As Currency
    Dim menge As Double
    Dim text As String
    Dim stkpreis As Currency
    Dim preis As Currency
    
    Dim summe As Currency
    summe = 0
    
    Dim aRow As Long
    For aRow = LBound(Positionen, 1) To UBound(Positionen, 1)
        menge = Positionen(aRow, 1)
        text = Positionen(aRow, 2)
        stkpreis = Positionen(aRow, 3)
        preis = Positionen(aRow, 4)
        If stkpreis <> 0 Then summe = summe + stkpreis
    Next aRow
    
    CalcNetto = summe
End Function

Public Function CreateFile(ByRef aBeleg As ClassBeleg) As Workbook
    Set CreateFile = Nothing
    
    If ExitOnError(FSO.FileExists(aBeleg.Filename), "Die Beleg-Datei existiert bereits") Then Exit Function
    
    On Error GoTo ErrorExit
    
    OptimizedMode True
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add(PathCombine(FindOnedrivePath(ONEDRIVE_Organization, ONEDRIVE_Channel, ONEDRIVE_Subdir), "Template.xltx"))
    Dim sheet As Worksheet
    Set sheet = wb.Worksheets(1)
    
    sheet.Range("Bearbeiter").value = CurrentUser.Kürzel
    sheet.Range("Email").value = aBeleg.Kunde.Email
    sheet.Range("KdNr").value = aBeleg.Kunde.Debitor
    sheet.Range("Kunde").value = aBeleg.Kunde.Anschrift
    sheet.Range("ReDatum").value = aBeleg.Datum
    sheet.Range("ReNr").value = aBeleg.BelegNr
    sheet.Range("Steuersatz").value = aBeleg.Steuer
    
    Select Case aBeleg.Kind
        Case BelegKind.xbkRechnung
            sheet.Range("LieferDatum").value = aBeleg.Datum
            sheet.Range("Nachspann1").value = StrFormat("Bitte zahlen Sie bis spätestens %1 auf folgendes Konto:", FormatDateTime(aBeleg.Datum + OptionValue(xoptDefaultDue)))
            sheet.Range("Nachspann2").value = OptionValue(xoptFirmaBank) & ", IBAN " & OptionValue(xoptFirmaIBAN)
            sheet.Range("Nachspann3").value = "Verwendungszweck: " & aBeleg.BelegNr
            sheet.Range("Nachspann5").value = "Beachten Sie unsere neue Bankverbindung!"
        Case BelegKind.xbkGutschrift
            sheet.Range("Nachspann1").value = "Bitte ziehen Sie die Gutschrift nicht von Rechnungen ab, denn"
            sheet.Range("Nachspann2").value = "wir überweisen Ihnen den Betrag in wenigen Tagen auf folgendes Konto:"
            sheet.Range("Nachspann3").value = aBeleg.Kunde.IBAN
            sheet.Range("Nachspann4").value = aBeleg.Kunde.Bank
        Case BelegKind.xbkStornoGutschrift
            sheet.Range("Nachspann1").value = StrFormat("%3 %1 vom %2", aBeleg.BelegNr, FormatDateTime(aBeleg.Datum, vbShortDate), GUTSCHRIFT_TEXT)
    End Select

    wb.SaveAs aBeleg.Filename
    Set CreateFile = wb
    Exit Function
    
ErrorExit:
    If Not (wb Is Nothing) Then wb.Close (False)
    OptimizedMode False
End Function

Public Function LoadFromFile(ByRef aBeleg As ClassBeleg) As Boolean
    LoadFromFile = False
    
    Dim Filename As String
    Filename = aBeleg.Filename
    If Not FSO.FileExists(Filename) Then Exit Function
    
    On Error GoTo ErrorExit
    
    Dim app As New ClassApplication
    OptimizedMode True, app.Instance
    Dim wb As Workbook: Set wb = app.Instance.Workbooks.Open(Filename)
    Dim sheet As Worksheet: Set sheet = wb.Worksheets(1)
    
    Bearbeiter = sheet.Range("Bearbeiter").value
    Brutto = sheet.Range("Brutto").value
    Email = sheet.Range("Email").value
    KdNr = sheet.Range("KdNr").value
    Kunde = sheet.Range("Kunde").value
    LieferDatum = sheet.Range("LieferDatum").value
    Nachspann1 = sheet.Range("Nachspann1").value
    Nachspann2 = sheet.Range("Nachspann2").value
    Nachspann3 = sheet.Range("Nachspann3").value
    Nachspann4 = sheet.Range("Nachspann4").value
    Nachspann5 = sheet.Range("Nachspann5").value
    Netto = sheet.Range("Netto").value
    RCText = sheet.Range("RCText").value
    ReDatum = sheet.Range("ReDatum").value
    ReNr = sheet.Range("ReNr").value
    Steuer = sheet.Range("Steuer").value
    Steuersatz = sheet.Range("Steuersatz").value
        
    Positionen = sheet.ListObjects(1).DataBodyRange
    
    LoadFromFile = True
    Exit Function
    
ErrorExit:
    If Not (wb Is Nothing) Then wb.Close (False)
    OptimizedMode False, app.Instance
End Function
