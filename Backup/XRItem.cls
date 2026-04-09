VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XRItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' Class XRItem
' =========================================

' XRechnung: Daten einer Rechnungsposition

Public Price As Currency
Public Quantity As Double
Public TaxPercent As Double

Private PrivateName As String
Private PrivateDescription As String

' automatisches Berechnen der Summe
Public Property Get LineAmount() As Variant
     LineAmount = IIf(Quantity <> 0 And Price <> 0, Quantity * Price, 0)
End Property

' automatisches Berechnen der Steuer
Public Property Get LineTaxAmount() As Variant
    LineTaxAmount = IIf(TaxPercent <> 0, Round(LineAmount * (TaxPercent / 100), 2), 0)
End Property

' Mit den Properties wird sichergestellt, dass beim Setzen der Daten spezielle Zeichen
' XML-konform ersetzt werden. Siehe auch:StrExpandToHTML
Public Property Get Name() As Variant
    Name = PrivateName
End Property

Public Property Let Name(ByVal value As Variant)
    PrivateName = StrExpandToHTML(value)
End Property

Public Property Get Description() As Variant
    Description = PrivateDescription
End Property

Public Property Let Description(ByVal value As Variant)
    PrivateDescription = StrExpandToHTML(value)
End Property

