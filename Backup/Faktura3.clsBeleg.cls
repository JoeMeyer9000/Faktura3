VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBeleg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *****************************************************
' Class clsBeleg
' *****************************************************

' class enums
Public Enum ColIdxBeleg
    xblgDatum = 1
    xblgKunde
    xblgBelegArt
    xblgNetto
    xblgSteuersatz
    xblgBrutto
    xblgBelegNr
    xblgMitarbeiter
    xblgReferenz
    xblgBelegId
    xblgDebitor
End Enum

Public Enum BelegKind
    xbkVoid
    xbkRechnung
    xbkGutschrift
    xbkStornoGutschrift
End Enum


' class fields
Public datum As Date
Public Kunde As String
Public BelegArt As String
Public Netto As Currency
Public Steuersatz As Double
Public Brutto As Currency
Public belegNr As String
Public Mitarbeiter As String
Public Referenz As String
Public BelegID As Integer
Public Debitor As Double

Private m_Kunde As clsKunde

Public MyRow As Long

' =====================================
' Public
' =====================================

Public Property Get SheetRowIndex() As Long
    SheetRowIndex = MapRowToSheet(TBelege, MyRow)
End Property

Public Property Get Filename(Optional extension As String = "xlsx") As String
    Filename = MakeBelegFilename(datum, BelegID, extension)
End Property

Public Property Get IsReverseCharge() As Boolean
    IsReverseCharge = (Steuersatz < 0.01) And (Abs(Netto) >= 5000)
End Property

Public Property Get KundeData() As clsKunde
    If m_Kunde Is Nothing Then
        Dim kdRow As Long
        kdRow = FindRow(TKunden.ListColumns(xkdDebitor), Debitor)
        If ExitOnError(kdRow = 0, "Ungültige Debitor-Nummer") Then Exit Property
        Set m_Kunde = New clsKunde
        Call m_Kunde.Populate(TKunden, kdRow)
    End If
    Set KundeData = m_Kunde
End Property

Public Property Get Kind() As BelegKind
    If BelegArt = "Rechnung" Then
        Kind = xbkRechnung
    ElseIf BelegArt = "Gutschrift" Then
        Kind = IIf(Referenz = "", xbkGutschrift, xbkStornoGutschrift)
    Else
        Kind = xbkVoid
    End If
End Property

' Read a row into the data fields
Public Sub Populate(ByVal aRowIndex As Long)
    MyRow = aRowIndex
    Refresh
End Sub

Public Sub Refresh()
    Dim values As Variant
    values = TBelege.DataBodyRange.Rows(MyRow)
    datum = values(1, xblgDatum)
    Kunde = values(1, xblgKunde)
    BelegArt = values(1, xblgBelegArt)
    Netto = values(1, xblgNetto)
    Steuersatz = values(1, xblgSteuersatz)
    Brutto = values(1, xblgBrutto)
    belegNr = values(1, xblgBelegNr)
    Mitarbeiter = values(1, xblgMitarbeiter)
    Referenz = values(1, xblgReferenz)
    BelegID = values(1, xblgBelegId)
    Debitor = values(1, xblgDebitor)
End Sub

Public Sub CreateOrOpen()
    Dim row As Variant
    Set row = TBelege.DataBodyRange.Rows(MyRow)
    
    If BelegID = 0 Then
        ' keine Beleg-Id angegeben, also muss die Datei neu erzeugt werden
        
        ' KO-Kriterien
        If ExitOnError(Kunde = "", "Bitte wählen Sie einen Kunden aus") Then Exit Sub
        If ExitOnError(Debitor < OptionValue(xoptMinDebitor), "Die Debitor-Nummer ist ungültig") Then Exit Sub
        
        ' fehlende Daten automatisch setzen
        If row.Cells(xblgDatum) = 0 Then row.Cells(xblgDatum) = Date
        If IsEmpty(row.Cells(xblgBelegArt)) Then row.Cells(xblgBelegArt) = "Rechnung"
        If IsEmpty(row.Cells(xblgSteuersatz)) Then row.Cells(xblgSteuersatz) = OptionValue(xoptDefaultTax)
        row.Cells(xblgBelegId) = NewBelegId()
        row.Cells(xblgMitarbeiter) = CurrentUser.Kürzel
        
        ' Zellen wurden verändert, also neu einlesen
        Refresh
        
        ' erst jetzt kann der Filename generiert werden, nachdem die BelegId gesetzt ist
        If ExitOnError(Filename = "", "Der Dateiname konnte nicht generiert werden, weil Datum oder Rechnungs-Id ungültig sind") Then Exit Sub
        If ExitOnError(FSO.FileExists(Filename), StrFormat("Eine %1 mit der Nummer existiert bereits", BelegArt)) Then Exit Sub
        
        ' ... und zum Schluß das Dokument erzeugen
        Dim doc As New clsBelegDoc
        doc.CreateFile Me
    ElseIf ExitOnError(Not FSO.FileExists(Filename), StrFormat("Der Beleg %1 wurde nicht gefunden", BelegID)) Then Exit Sub
    Else
        ShellExecute Filename
    End If
End Sub

Public Sub SendByMail()
    If Not CreatePDF() Then Exit Sub
    If Not CreateXRechnung() Then Exit Sub
    
    Dim files As New Collection
    files.Add Filename("pdf")
    files.Add Filename("xml")
    
    Dim betreff As String
    betreff = StrFormat("Ihre %1 %2 - %3", BelegArt, belegNr, Kunde)
    
    Dim sendTo As String
    sendTo = IIf(KundeData.RechEMail = "", KundeData.Email, KundeData.RechEMail)
    If ExitOnError(sendTo = "", "Eine E-Mail kann nicht generiert werden, weil zu der Adresse keine E-Mail-Adresse angegeben wurde") Then Exit Sub
    
    Dim sendBCC As String
    sendBCC = IIf(Kind = xbkGutschrift, OptionValue(xoptGutschriftBCC), "")
    
    Call GenerateEMail(sendTo, "", sendBCC, betreff, GenerateEmailBody, files)
End Sub

Public Function CreatePDF() As Boolean
    CreatePDF = False
    If ExitOnError(Not FSO.FileExists(Filename), StrFormat("Beleg %1 existiert nicht", belegNr)) Then Exit Function
    
    Dim app As New clsApplication
    Dim wb As Workbook
    Set wb = app.Instance.Workbooks.Open(Filename, ReadOnly = True)
    wb.ExportAsFixedFormat xlTypePDF, Filename("pdf")
    wb.Close False
    
    CreatePDF = True
End Function

Public Function CreateXRechnung() As Boolean
    CreateXRechnung = False
    If ExitOnError(Not FSO.FileExists(Filename), StrFormat("Beleg %1 existiert nicht", belegNr)) Then Exit Function
    
    Dim xml As String
    Dim XR As New clsXRechnung
    xml = XR.Generate(MyRow)
    If xml = "" Then Exit Function
    
    Dim fname As String: fname = Filename(".xml")
    If FSO.FileExists(fname) Then Kill fname
    ' XRechnungen müssen in UTF-8 kodiert sein, sonst sind sie nicht KOSIT konform
    WriteTextToUTF8File fname, xml
     
     CreateXRechnung = True
End Function

Public Function Synchronize() As Boolean
    Synchronize = False
    
    Dim doc As New clsBelegDoc
    If ExitOnError(Not doc.LoadFromFile(Me), StrFormat("Die Beleg %1 existiert nicht", belegNr)) Then Exit Function
    
    With TBelege.DataBodyRange
        If ExitOnError(Right$(belegNr, 8) <> Right$(doc.ReNr, 8), StrFormat("Die Rechnungsnummer %1 im Beleg stimmt nicht mit der Rechnungsnummer %2 in der Liste überein", doc.ReNr, belegNr)) Then Exit Function
        If ExitOnError(doc.KdNr <> Debitor, StrFormat("Die Debitor-Nummer %1 im Beleg ist nicht identisch mit der Debitor-Nummer %2 in der Liste", doc.ReNr, belegNr)) Then Exit Function
        If ExitOnError(Not SameAnschrift(doc.Kunde, KundeData.Anschrift), StrFormat("Die Kundenanschrift im Beleg %1 ist nicht identisch mit der Anschrift des Kunden in der Kundenliste", belegNr)) Then Exit Function
        If ExitOnError(doc.Kind = xbkVoid, StrFormat("Die Belegart von Beleg %1 kann nicht ermittelt werden", belegNr)) Then Exit Function
            
        ' xbkVoid hat bereits einen Fehler verursacht
        ' xbkStornoGutschrift ist ebenfalls eine Gutschrift
        .Cells(MyRow, xblgBelegArt) = IIf(doc.Kind = xbkRechnung, "Rechnung", "Gutschrift")
        .Cells(MyRow, xblgDatum) = doc.ReDatum
        .Cells(MyRow, xblgNetto) = doc.Netto
        .Cells(MyRow, xblgSteuersatz) = doc.Steuersatz
        .Cells(MyRow, xblgMitarbeiter) = doc.Bearbeiter
    End With
    
    Refresh
    
    Synchronize = True
    
End Function

Public Function Check(errors As Collection) As Boolean
    Check = False
    
    Dim doc As New clsBelegDoc
    If Not doc.LoadFromFile(Me) Then
        AddError errors, "Beleg existiert nicht"
        Exit Function
    End If
    
    With TBelege.DataBodyRange
        If Right$(belegNr, 8) <> Right$(doc.ReNr, 8) Then AddError errors, "Rechnungsnummer nicht synchron"
        If doc.KdNr <> Debitor Then AddError errors, "Debitor-Nummer nicht synchron"
        If Not SameAnschrift(doc.Kunde, KundeData.Anschrift) Then AddError errors, "Kundenanschrift nicht synchron"
        If doc.ReDatum <> datum Then AddError errors, "Rechnungsdatum nicht synchron"
        If (Left$(doc.ReNr, 1) = "R" And Kind <> xbkRechnung) Or (Left$(doc.ReNr, 1) = "G" And Kind = xbkRechnung) Then AddError errors, "Beleg-Art nicht synchron"
        If doc.Netto <> Netto Then AddError errors, "Netto-Betrag nicht synchron"
        If doc.Steuersatz <> Steuersatz Then AddError errors, "Steuersatz nicht synchron"
        If doc.Bearbeiter <> Mitarbeiter Then AddError errors, "Bearbeiter nicht synchron"
    End With
    
    Refresh
    Check = True
End Function

Public Sub PrintBeleg()
    If ExitOnError(BelegID = 0, StrFormat("Der Beleg in Zeile %1 wurde noch nicht erzeugt", MapRowToSheet(TBelege, MyRow))) Then Exit Sub
    If ExitOnError(Not FSO.FileExists(Filename), StrFormat("%1 %2 wurde nicht gefunden", BelegArt, belegNr)) Then Exit Sub

    On Error GoTo ErrorExit
    Dim app As New clsApplication
    Dim wb As Workbook
    Set wb = app.Instance.Workbooks.Open(Filename, ReadOnly:=True)
    wb.PrintOut

ErrorExit:
    If Not (wb Is Nothing) Then wb.Close False
End Sub

Public Sub Storno()
    If ExitOnError(Kind <> xbkRechnung, "Es können nur Rechnungen storniert werden") Then Exit Sub
    If ExitOnError(BelegID < OptionValue(xoptMinRechId), "Die Rechnung wurde noch nicht geschrieben") Then Exit Sub
    If ExitOnError(Not FSO.FileExists(Filename), "Die Rechnungs-Datei existiert nicht") Then Exit Sub
    If ExitOnError(Referenz <> "", "Die Rechnung wurde bereits storniert und kann nur 1x storniert werden") Then Exit Sub
    If ExitOnError(Netto = 0, "Die Rechnung enthält einen Netto-Betrag von 0,00 €, eine Stornierung macht keinen Sinn") Then Exit Sub
    
    Dim newRowIndex As Long: newRowIndex = TBelege.ListRows.Add.index
    Dim newId As Long: newId = NewBelegId
    Dim Gutschrift As New clsBeleg: Gutschrift.Populate newRowIndex
    
    With TBelege.DataBodyRange.Rows(newRowIndex)
        .Cells(xblgDatum) = Date
        .Cells(xblgKunde) = Kunde
        .Cells(xblgBelegArt) = "Gutschrift"
        .Cells(xblgNetto) = -Netto
        .Cells(xblgSteuersatz) = Steuersatz
        .Cells(xblgMitarbeiter) = CurrentUser.Kürzel
        .Cells(xblgBelegId) = newId
        .Cells(xblgReferenz) = belegNr
        
    End With
    
    TBelege.DataBodyRange.Rows(newRowIndex).Select
    
    'shBelege.Calculate
    Gutschrift.Refresh
    TBelege.DataBodyRange.Cells(MyRow, xblgReferenz) = Gutschrift.belegNr
    Refresh
    
    On Error GoTo ErrorExit
    
    FileCopy Filename, Gutschrift.Filename
        
    Dim app As New clsApplication
    Dim wb As Workbook
    Dim sheet As Worksheet
    
    OptimizedMode True, app.Instance
    Set wb = app.Instance.Workbooks.Open(Gutschrift.Filename)
    Set sheet = wb.Worksheets(1)
    
    With sheet
        .Range("Bearbeiter").value = Gutschrift.Mitarbeiter
        .Range("ReDatum").value = Gutschrift.datum
        .Range("ReNr").value = Gutschrift.belegNr
        .Range("LieferDatum").Clear
        .Range("Nachspann1").value = StrFormat("Diese Gutschrift ersetzt Rechnung %1", Gutschrift.belegNr)
        .Range("Nachspann2").ClearContents
        .Range("Nachspann3").ClearContents
        .Range("Nachspann4").ClearContents
        .Range("Nachspann5").ClearContents
    End With
    
    Dim tbl As ListObject: Set tbl = sheet.ListObjects(1)
    
    With tbl.DataBodyRange
        Dim i As Long
        Dim value As Currency
        For i = 1 To .Rows.count
            value = .Cells(i, 3).value
            If value <> 0 Then .Cells(i, 3).value = -value
        Next i
    End With
    
    wb.Save
    
ErrorExit:
    If Not (wb Is Nothing) Then wb.Close (False)
    
End Sub

' =====================================
' Private
' =====================================

Private Function NewBelegId() As Long
    NewBelegId = NextValue(TBelege.ListColumns("Beleg-ID"), OptionValue(xoptMinRechId))
End Function

Private Sub AddError(errors As Collection, message As String)
    errors.Add StrFormat("[%1] %2", belegNr, message)
End Sub

' Erzeugt aus den Belegdaten einen E-Mail-Text
Private Function GenerateEmailBody() As String
    Dim HTML As String
    
    AddHtmlLine HTML, "<p style=""|FONT_STYLE|"">Sehr geehrte Damen und Herren,</p>"
    AddHtmlLine HTML, StrFormat("<p style=""|FONT_STYLE|"">Im Anhang erhalten Sie %1 %2 vom %3 über %4</p>", BelegArt, belegNr, Format(datum, "dd/mm/yyyy"), Format(Abs(Brutto), "Currency"))
    Select Case Kind
    
    Case BelegKind.xbkRechnung
        AddHtmlLine HTML, StrFormat("<p style=""|FONT_STYLE|"">Bitte überweisen Sie den offenen Betrag bis spätestens %1 auf folgendes Konto:<br>", Format(datum + OptionValue(xoptDefaultDue), "dd/mm/yyyy"))
        AddHtmlLine HTML, StrFormat("%1 bei der %2</p>", OptionValue(xoptFirmaIBAN), OptionValue(xoptFirmaBank))
    Case BelegKind.xbkStornoGutschrift
        AddHtmlLine HTML, StrFormat(" <p style=""|FONT_STYLE|"">Diese Gutschrift ersetzt unsere Rechnung %1, die damit hinfällig wird</p>", Referenz)
    Case Else
        AddHtmlLine HTML, StrFormat(" <p style=""|FONT_STYLE|"">Bitte ziehen Sie den Betrag nicht von offenen Rechnungen ab<br>")
        AddHtmlLine HTML, StrFormat("Wir überweisen Ihnen den Betrag in den nächsten Tagen auf folgendes Konto:<br>")
        AddHtmlLine HTML, StrFormat("%1, %2</p>", KundeData.IBAN, KundeData.Bank)
    End Select
    
    GenerateEmailBody = HTML
End Function

' Hilfsroutine, um E-Mail-Text zu erzeugen
Private Sub AddHtmlLine(ByRef HTML As String, ByVal line As String)
    line = Replace(line, "|FONT_STYLE|", "font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 11pt", Compare:=vbTextCompare)
    If HTML <> "" Then HTML = HTML & vbCrLf
    HTML = HTML & line
End Sub

