VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XRAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' Class XRAddress
' =========================================

' XRechnung: Adressdaten (Seller und Buyer)

Public ZipCode As String
Public Email As String
Public VATID As String
Public TaxNumber As String

Private PrivateName As String
Private PrivateTradingName As String
Private PrivateHR As String
Private PrivateStreet1 As String
Private PrivateStreet2 As String
Private PrivateCity As String
Private PrivateContactName As String
Private PrivateContactPhone As String
Private PrivateContactEmail As String

' Mit den Properties wird sichergestellt, dass beim Setzen der Daten spezielle Zeichen
' XML-konform ersetzt werden. Siehe auch:StrExpandToHTML
Public Property Get Name() As Variant
    Name = PrivateName
End Property

Public Property Let Name(ByVal value As Variant)
    PrivateName = StrExpandToHTML(value)
End Property

Public Property Get TradingName() As Variant
    TradingName = PrivateTradingName
End Property

Public Property Let TradingName(ByVal value As Variant)
    PrivateTradingName = StrExpandToHTML(value)
End Property

Public Property Get HR() As Variant
    HR = PrivateHR
End Property

Public Property Let HR(ByVal value As Variant)
    PrivateHR = StrExpandToHTML(value)
End Property

Public Property Get Street1() As Variant
    Street1 = PrivateStreet1
End Property

Public Property Let Street1(ByVal value As Variant)
    PrivateStreet1 = StrExpandToHTML(value)
End Property

Public Property Get Street2() As Variant
    Street2 = PrivateStreet2
End Property

Public Property Let Street2(ByVal value As Variant)
    PrivateStreet2 = StrExpandToHTML(value)
End Property

Public Property Get City() As Variant
    City = PrivateCity
End Property

Public Property Let City(ByVal value As Variant)
    PrivateCity = StrExpandToHTML(value)
End Property

Public Property Get ContactName() As Variant
    ContactName = PrivateContactName
End Property

Public Property Let ContactName(ByVal value As Variant)
    PrivateContactName = StrExpandToHTML(value)
End Property

Public Property Get ContactPhone() As Variant
    ContactPhone = PrivateContactPhone
End Property

Public Property Let ContactPhone(ByVal value As Variant)
    PrivateContactPhone = StrExpandToHTML(value)
End Property

Public Property Get ContactEmail() As Variant
    ContactEmail = PrivateContactEmail
End Property

Public Property Let ContactEmail(ByVal value As Variant)
    PrivateContactEmail = StrExpandToHTML(value)
End Property
