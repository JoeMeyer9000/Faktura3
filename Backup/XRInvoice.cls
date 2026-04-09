VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XRInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' Class XRInvoice
' =========================================

' XRechnung: Daten einer Rechnung

Public Enum XDocTypeCode
    TypeRechnung = 380
    TypeGutschrift = 381
    TypeKorrektur = 384
End Enum

Public Enum XPaymentTypeCodes
    SEPACreditTransfer = 58
End Enum

Public TypeCode As XDocTypeCode
Public IssuedDate As Date
Public OccurrenceDate As Date
Public PaymentRef As String
Public PaymentTypeCode As XPaymentTypeCodes
Public IBAN As String
Public BIC As String
Public TaxCategoryCode As String
Public TaxPercent As Double

Private PrivateID As String
Private PrivateBuyerReference As String
Private PrivateBuyerOrderRefDoc As String
Private PrivateNote As String
Private PrivateTaxExemptionReason As String
Private PrivatePaymentDescription As String

' Mit den Properties wird sichergestellt, dass beim Setzen der Daten spezielle Zeichen
' XML-konform ersetzt werden. Siehe auch:StrExpandToHTML

Public Property Get ID() As Variant
    ID = PrivateID
End Property

Public Property Let ID(ByVal value As Variant)
    PrivateID = StrExpandToHTML(value)
    PaymentRef = PrivateID
End Property

Public Property Get Note() As Variant
    Note = PrivateNote
End Property

Public Property Let Note(ByVal value As Variant)
    PrivateNote = StrExpandToHTML(value)
End Property

Public Property Get BuyerReference() As Variant
    BuyerReference = PrivateBuyerReference
End Property

Public Property Let BuyerReference(ByVal value As Variant)
    PrivateBuyerReference = StrExpandToHTML(value)
End Property

Public Property Get BuyerOrderRefDoc() As Variant
    BuyerOrderRefDoc = PrivateBuyerOrderRefDoc
End Property

Public Property Let BuyerOrderRefDoc(ByVal value As Variant)
    PrivateBuyerOrderRefDoc = StrExpandToHTML(value)
End Property

Public Property Get TaxExemptionReason() As Variant
    TaxExemptionReason = PrivateTaxExemptionReason
End Property

Public Property Let TaxExemptionReason(ByVal value As Variant)
    PrivateTaxExemptionReason = StrExpandToHTML(value)
End Property

Public Property Get PaymentDescription() As Variant
    PaymentDescription = PrivatePaymentDescription
End Property

Public Property Let PaymentDescription(ByVal value As Variant)
    PrivatePaymentDescription = StrExpandToHTML(value)
End Property

Public Property Get PaymentCodeInfo() As Variant
    Select Case PaymentTypeCode
        Case XPaymentTypeCodes.SEPACreditTransfer
            PaymentCodeInfo = "SEPA credit transfer"
        Case Else
            PaymentCodeInfo = "n/a"
    End Select
End Property

Public Property Get dueDate() As Variant
    dueDate = DateAdd("d", 10, IssuedDate)
End Property


Private Sub Class_Initialize()
    PaymentTypeCode = XPaymentTypeCodes.SEPACreditTransfer
End Sub
