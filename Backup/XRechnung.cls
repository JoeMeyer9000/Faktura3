VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XRechnung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ********************************************
' Class XRechnung
' ********************************************

' XRechnung: Generator einer XRechnung

' Wie man eine XRechnung erzeugt:
' Füllen Sie die Variablen Invoice, Seller, Buyer und Items. Danach rufen Sie
' Generate() auf und erhalten den XML Code der XRechnung. Den Code speichern
' Sie dann mit JUtils.WriteTextToUTF8File() in eine Datei

' Diese Daten müssen von extern befüllt werden, bevor man Generate() aufrufen kann
Public Invoice As New XRInvoice
Public Seller As New XRAddress
Public Buyer As New XRAddress
Public Items As New Collection

' =========================================
' Public
' =========================================

' Füllen Sie die Variablen Invoice, Seller, Buyer und Items. Danach rufen Sie diese
' Fujnktion auf, um den XML Code zu erhalten
Public Function Generate() As String
    Generate = ""
    If Not IsRequiredInfoComplete() Then Exit Function
    
    Dim total As Currency
    total = Abs(CalcTotalAmount)
    
    Dim taxTotal As Currency
    taxTotal = Abs(CalcTotalTax)
    
    Dim X As New XRXml
    X.StartGroup "rsm:CrossIndustryInvoice", "xmlns:rsm=""urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100"" xmlns:qdt=""urn:un:unece:uncefact:data:standard:QualifiedDataType:100"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:ram=""urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100"" xmlns:udt=""urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100"""
        
        ' ExchangedDocumentContext
        X.StartGroup "rsm:ExchangedDocumentContext"
            X.StartGroup "ram:BusinessProcessSpecifiedDocumentContextParameter"
                X.Add "ram:ID", "urn:fdc:peppol.eu:2017:poacc:billing:01:1.0"
            X.CloseGroup
            X.StartGroup "ram:GuidelineSpecifiedDocumentContextParameter"
                X.Add "ram:ID", "urn:cen.eu:en16931:2017#compliant#urn:xeinkauf.de:kosit:xrechnung_3.0"
            X.CloseGroup
        X.CloseGroup "rsm:ExchangedDocumentContext"
        
        ' ExchangedDocument
        X.StartGroup "rsm:ExchangedDocument"
            X.Add "ram:ID", Invoice.ID
            X.AddInt "ram:TypeCode", Invoice.TypeCode
            X.AddDate "ram:IssueDateTime", Invoice.IssuedDate
            If Invoice.Note <> "" Then
                X.StartGroup "ram:IncludedNote"
                    X.Add "ram:Content", Invoice.Note
                X.CloseGroup
            End If
        X.CloseGroup "rsm:ExchangedDocument"
        
        ' SupplyChainTradeTransaction
        X.StartGroup "rsm:SupplyChainTradeTransaction"
        
            Dim i As Integer
            Dim item As XRItem
            For i = 1 To Items.Count
                Set item = Items(i)
                
                ' IncludedSupplyChainTradeLineItem
                X.StartGroup "ram:IncludedSupplyChainTradeLineItem"
                    
                    X.StartGroup "ram:AssociatedDocumentLineDocument"
                        X.AddInt "ram:LineID", i
                    X.CloseGroup
                    X.StartGroup "ram:SpecifiedTradeProduct"
                        X.Add "ram:Name", item.Name
                        X.Add "ram:Description", item.Description
                    X.CloseGroup
                    X.StartGroup "ram:SpecifiedLineTradeAgreement"
                        X.StartGroup "ram:NetPriceProductTradePrice"
                            X.AddNumber "ram:ChargeAmount", Abs(item.Price)
                            X.AddNumber "ram:BasisQuantity", 1, "unitCode=""H87"""
                        X.CloseGroup
                    X.CloseGroup
                    X.StartGroup "ram:SpecifiedLineTradeDelivery"
                        X.AddNumber "ram:BilledQuantity", item.Quantity, "unitCode=""H87"""
                    X.CloseGroup
                    X.StartGroup "ram:SpecifiedLineTradeSettlement"
                        X.StartGroup "ram:ApplicableTradeTax"
                            X.Add "ram:TypeCode", "VAT"
                            If Invoice.TaxExemptionReason <> "" Then
                                X.Add "ram:ExemptionReason", Invoice.TaxExemptionReason
                            End If
                            X.Add "ram:CategoryCode", Invoice.TaxCategoryCode
                            X.AddNumber "ram:RateApplicablePercent", Invoice.TaxPercent
                        X.CloseGroup
                        X.StartGroup "ram:SpecifiedTradeSettlementLineMonetarySummation"
                            X.AddNumber "ram:LineTotalAmount", Abs(item.LineAmount)
                        X.CloseGroup
                    X.CloseGroup "ram:SpecifiedLineTradeSettlement"
                    
                X.CloseGroup "ram:IncludedSupplyChainTradeLineItem"
                
            Next i
            
            ' ApplicableHeaderTradeAgreement
            X.StartGroup "ram:ApplicableHeaderTradeAgreement"
                X.Add "ram:BuyerReference", Invoice.BuyerReference
                GenerateAddress X, Seller, "ram:SellerTradeParty"
                GenerateAddress X, Buyer, "ram:BuyerTradeParty"
                X.StartGroup "ram:BuyerOrderReferencedDocument"
                    X.Add "ram:IssuerAssignedID", Invoice.BuyerOrderRefDoc
                X.CloseGroup
            X.CloseGroup "ram:ApplicableHeaderTradeAgreement"
            
            ' ApplicableHeaderTradeDelivery
            X.StartGroup "ram:ApplicableHeaderTradeDelivery"
                If Invoice.OccurrenceDate <> 0 Then
                X.StartGroup "ram:ActualDeliverySupplyChainEvent"
                    X.AddDate "ram:OccurrenceDateTime", Invoice.OccurrenceDate
                X.CloseGroup
                End If
            X.CloseGroup "ram:ApplicableHeaderTradeDelivery"
        
            ' ApplicableHeaderTradeSettlement
            X.StartGroup "ram:ApplicableHeaderTradeSettlement"
                X.Add "ram:PaymentReference", Invoice.PaymentRef
                X.Add "ram:InvoiceCurrencyCode", "EUR"
                X.StartGroup "ram:SpecifiedTradeSettlementPaymentMeans"
                    X.AddNumber "ram:TypeCode", Invoice.PaymentTypeCode
                    X.Add "ram:Information", Invoice.PaymentCodeInfo
                    X.StartGroup "ram:PayeePartyCreditorFinancialAccount"
                        X.Add "ram:IBANID", Invoice.IBAN
                        X.Add "ram:AccountName", Seller.Name
                    X.CloseGroup
                    If Invoice.BIC <> "" Then
                        X.StartGroup "ram:PayeeSpecifiedCreditorFinancialInstitution"
                            X.Add "ram:BICID", Invoice.BIC
                        X.CloseGroup
                    End If
                X.CloseGroup "ram:SpecifiedTradeSettlementPaymentMeans"
                X.StartGroup "ram:ApplicableTradeTax"
                    X.AddNumber "ram:CalculatedAmount", taxTotal
                    X.Add "ram:TypeCode", "VAT"
                    If Invoice.TaxExemptionReason <> "" Then
                        X.Add "ram:ExemptionReason", Invoice.TaxExemptionReason
                    End If
                    X.AddNumber "ram:BasisAmount", total
                    X.Add "ram:CategoryCode", Invoice.TaxCategoryCode
                    X.AddNumber "ram:RateApplicablePercent", Invoice.TaxPercent
                X.CloseGroup "ram:ApplicableTradeTax"
                X.StartGroup "ram:SpecifiedTradePaymentTerms"
                    X.Add "ram:Description", Invoice.PaymentDescription
                    X.AddDate "ram:DueDateDateTime", Invoice.dueDate
                X.CloseGroup "ram:SpecifiedTradePaymentTerms"
                X.StartGroup "ram:SpecifiedTradeSettlementHeaderMonetarySummation"
                    X.AddNumber "ram:LineTotalAmount", total
                    X.AddNumber "ram:ChargeTotalAmount", 0
                    X.AddNumber "ram:AllowanceTotalAmount", 0
                    X.AddNumber "ram:TaxBasisTotalAmount", total
                    X.AddNumber "ram:TaxTotalAmount", taxTotal, "currencyID=""EUR"""
                    X.AddNumber "ram:GrandTotalAmount", total + taxTotal
                    X.AddNumber "ram:TotalPrepaidAmount", 0
                    X.AddNumber "ram:DuePayableAmount", total + taxTotal
                X.CloseGroup "ram:SpecifiedTradeSettlementHeaderMonetarySummation"
            X.CloseGroup "ram:ApplicableHeaderTradeSettlement"
        
        X.CloseGroup "rsm:SupplyChainTradeTransaction"
    
    X.CloseGroup
    Generate = X.xml
End Function

' =========================================
' Private
' =========================================

' Generiert den Code eine Adresse (Seller und Buyer)
Private Sub GenerateAddress(X As XRXml, Adr As XRAddress, GroupName As String)
    X.StartGroup GroupName
        X.Add "ram:Name", Adr.Name
        X.StartGroup "ram:SpecifiedLegalOrganization"
            X.Add "ram:ID", Adr.HR
            X.Add "ram:TradingBusinessName", Adr.TradingName
        X.CloseGroup
        X.StartGroup "ram:DefinedTradeContact"
            X.Add "ram:PersonName", Adr.ContactName
            X.StartGroup "ram:TelephoneUniversalCommunication"
                X.Add "ram:CompleteNumber", Adr.ContactPhone
            X.CloseGroup
            X.StartGroup "ram:EmailURIUniversalCommunication"
                X.Add "ram:URIID", Adr.ContactEmail
            X.CloseGroup
        X.CloseGroup
        X.StartGroup "ram:PostalTradeAddress"
            X.Add "ram:PostcodeCode", Adr.ZipCode
            X.Add "ram:LineOne", Adr.Street1
            X.Add "ram:LineTwo", Adr.Street2
            X.Add "ram:CityName", Adr.City
            X.Add "ram:CountryID", "DE"
        X.CloseGroup
        X.StartGroup "ram:URIUniversalCommunication"
            X.Add "ram:URIID", Adr.Email, "schemeID=""EM"""
        X.CloseGroup
        If Adr.VATID <> "" Then
            X.StartGroup "ram:SpecifiedTaxRegistration"
                X.Add "ram:ID", Adr.VATID, "schemeID=""VA"""
            X.CloseGroup
        End If
        If Adr.TaxNumber <> "" Then
            X.StartGroup "ram:SpecifiedTaxRegistration"
                X.Add "ram:ID", Adr.TaxNumber, "schemeID=""FC"""
            X.CloseGroup
        End If
    X.CloseGroup
End Sub

' Liefert True, wenn alle erforderlichen Daten vorhanden sind, um eine XRechnung zu erzeugen
Private Function IsRequiredInfoComplete() As Boolean
    IsRequiredInfoComplete = False
    Const UNDEFINED As String = " is undefined"
    
    If Invoice Is Nothing Then
        MsgBox "Invoice object" & UNDEFINED, vbOKOnly Or vbCritical, "XRechnung"
        Exit Function
    End If
    
    If Seller Is Nothing Then
        MsgBox "Seller object" & UNDEFINED, vbOKOnly Or vbCritical, "XRechnung"
        Exit Function
    End If
    
    If Buyer Is Nothing Then
        MsgBox "Buyer object" & UNDEFINED, vbOKOnly Or vbCritical, "XRechnung"
        Exit Function
    End If
    
    If Items Is Nothing Then
        MsgBox "Items collection" & UNDEFINED, vbOKOnly Or vbCritical, "XRechnung"
        Exit Function
    End If
    
    If Items.Count = 0 Then
        MsgBox "Items collection is empty", vbOKOnly Or vbCritical, "XRechnung"
        Exit Function
    End If
    
    IsRequiredInfoComplete = True
End Function

' Berechnet die Gesamtsumme (netto) der Rechnung
Private Function CalcTotalAmount() As Currency
    Dim sum As Currency
    Dim item As XRItem
    For Each item In Items
        sum = sum + item.LineAmount
    Next item
    CalcTotalAmount = Abs(sum)
End Function

' Berechnet den Gesamt-Steuerbetrag der Rechnung
Private Function CalcTotalTax() As Currency
    Dim sum As Currency
    Dim item As XRItem
    For Each item In Items
        sum = sum + item.LineTaxAmount
    Next item
    CalcTotalTax = Abs(sum)
End Function

