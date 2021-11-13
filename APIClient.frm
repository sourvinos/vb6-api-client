VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Begin VB.Form APIClient 
   Caption         =   "API Destinations"
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8340
   Icon            =   "APIClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtId 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1725
      TabIndex        =   4
      Top             =   675
      Width           =   540
   End
   Begin VB.TextBox txtResults 
      Appearance      =   0  'Flat
      Height          =   3390
      Left            =   3525
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "APIClient.frx":014A
      Top             =   150
      Width           =   4665
   End
   Begin Dacara_dcButton.dcButton cmdGet 
      Height          =   465
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonStyle     =   8
      Caption         =   "Get"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmdCreate 
      Height          =   465
      Left            =   150
      TabIndex        =   2
      Top             =   1200
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonStyle     =   8
      Caption         =   "Create"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmdGetById 
      Height          =   465
      Left            =   150
      TabIndex        =   3
      Top             =   675
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonStyle     =   8
      Caption         =   "GetById"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton createXML 
      Height          =   465
      Left            =   150
      TabIndex        =   5
      Top             =   3075
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonStyle     =   8
      Caption         =   "Create XML"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmdCreateMyData 
      Height          =   465
      Left            =   1725
      TabIndex        =   6
      Top             =   1200
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonStyle     =   8
      Caption         =   "Create myDATA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "APIClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents winH As WinHttp.WinHttpRequest
Attribute winH.VB_VarHelpID = -1


Private Function CreateCounterPart()

    Dim myArray(8) As String
    
    myArray(0) = "000000000"
    myArray(1) = "GR"
    myArray(2) = "0"
    myArray(3) = "ÅÐÙÍÕÌÉÁ"
    myArray(4) = "ÄÉÅÕÈÕÍÓÇ"
    myArray(5) = ""
    myArray(6) = ""
    myArray(7) = "ÊÅÑÊÕÑÁ"

    CreateCounterPart = myArray

End Function

Private Function CreateInvoiceDetails()

    Dim myArray(1, 5) As String
    
    myArray(0, 0) = "1" 'Á/Á ãñáììÞò
    myArray(0, 1) = "1" 'ÊáèáñÞ áîßá
    myArray(0, 2) = "5" 'Êáôçãïñßá ÖÐÁ
    myArray(0, 3) = "2.4" 'Ðïóü ÖÐÁ
    myArray(0, 4) = "E3_106" 'Êùäéêüò ÷áñáêôçñéóìïý
    
    CreateInvoiceDetails = myArray

End Function

Private Function CreateInvoiceHeader()

    Dim myArray(5) As String
    
    myArray(0) = "-" 'ÓåéñÜ
    myArray(1) = "1" 'Íï ðáñáóôáôéêïý
    myArray(2) = "2021-01-01" 'Çìåñïìçíßá
    myArray(3) = "2.1" 'Ôýðïò ðáñáóôáôéêïý
    myArray(4) = "EUR"
    
    CreateInvoiceHeader = myArray

End Function

Private Function CreateInvoiceSummary()

    Dim myArray(8) As String
    
    myArray(0) = "1" 'TotalNetValue
    myArray(1) = "0.24" 'TotalVatAmount
    myArray(2) = "0" 'TotalWithheldAmount
    myArray(3) = "0" 'TotalFeesAmount
    myArray(4) = "0" ''TotalStampDutyAmount
    myArray(5) = "0" 'TotalOtherTaxesAmount
    myArray(6) = "0" 'TotalDeductionsAmount
    myArray(7) = "1.24" 'TotalGrossValue
    
    CreateInvoiceSummary = myArray

End Function

Private Function CreateIssuer()

    Dim myArray(8) As String
    
    myArray(0) = "099863549"
    myArray(1) = "GR"
    myArray(2) = "0"
    myArray(3) = "ÊÑÏÔÓÇÓ Ì.Å.Ð.Å."
    myArray(4) = "ÅÈÍ. ËÅÕÊÉÌÌÇÓ"
    myArray(5) = "17Á"
    myArray(6) = "491 00"
    myArray(7) = "ÊÅÑÊÕÑÁ"
    
    CreateIssuer = myArray
    
End Function

Private Function CreatePaymentMethod()

    Dim myArray(2) As String
    
    myArray(0) = "1" 'Ôñüðïò ðëçñùìÞò: Ðßíáêáò 8.12
    myArray(1) = "124.00"
    
    CreatePaymentMethod = myArray

End Function


Private Function CreatePaymentMethodDetails()

    Dim myArray(1) As String
    
    myArray(0) = "3" 'Ðßíáêáò 8,12
    myArray(1) = "1.24" 'Ðïóü
    
    CreatePaymentMethodDetails = myArray

End Function

Private Function CreateRequestBody(issuerArray, counterPartArray, paymentMethodArray, paymentMethodDetailsArray, invoiceHeaderArray, invoiceDetailsArray, invoiceSummaryArray)

    Dim dom As DOMDocument
    Dim rootElement As IXMLDOMElement
    Dim objElement As IXMLDOMElement
    Dim objAttribute As IXMLDOMAttribute
    
    Dim root As IXMLDOMElement
    Dim invoice As IXMLDOMElement
    Dim issuer As IXMLDOMElement
    Dim address As IXMLDOMElement
    Dim element As IXMLDOMElement
    Dim counterPart As IXMLDOMElement
    Dim invoiceHeader As IXMLDOMElement
    Dim invoicePaymentMethod As IXMLDOMElement
    Dim invoicePaymentMethodDetails As IXMLDOMElement
    Dim invoiceDetails As IXMLDOMElement
    Dim incomeClassifications As IXMLDOMElement
    
    Dim invoiceSummary As IXMLDOMElement
    
    Set dom = New DOMDocument
    
    'Creates root element
    Set rootElement = dom.createElement("InvoicesDoc")
    dom.appendChild rootElement
    
    'Creates Attribute to the Root Element
    Set objAttribute = dom.createAttribute("xmlns")
    objAttribute.nodeValue = "http://www.aade.gr/myDATA/invoice/v1.0"
    rootElement.setAttributeNode objAttribute
    
    'Creates Attribute to the Root Element
    Set objAttribute = dom.createAttribute("xmlns:xsi")
    objAttribute.nodeValue = "http://www.w3.org/2001/XMLSchema-instance"
    rootElement.setAttributeNode objAttribute
    
    'Creates Attribute to the Root Element
    Set objAttribute = dom.createAttribute("xsi:schemaLocation")
    objAttribute.nodeValue = "http://www.aade.gr/my DATA/invoice/v1.0 schema.xsd"
    rootElement.setAttributeNode objAttribute
    
    'Invoice
    Set invoice = dom.createElement("invoice")
    rootElement.appendChild invoice
    
    'Invoice > Éssuer
    Set issuer = dom.createElement("issuer")
    invoice.appendChild issuer
    Set element = dom.createElement("vatNumber")
    issuer.appendChild element
    element.Text = issuerArray(0)
    Set element = dom.createElement("country")
    issuer.appendChild element
    element.Text = issuerArray(1)
    Set element = dom.createElement("branch")
    issuer.appendChild element
    element.Text = issuerArray(2)
    Set element = dom.createElement("name")
    issuer.appendChild element
    element.Text = issuerArray(3)
    
    'Invoice > Éssuer > Address
    Set address = dom.createElement("address")
    issuer.appendChild address
    Set element = dom.createElement("street")
    address.appendChild element
    element.Text = issuerArray(4)
    Set element = dom.createElement("number")
    address.appendChild element
    element.Text = issuerArray(5)
    Set element = dom.createElement("postalCode")
    address.appendChild element
    element.Text = issuerArray(6)
    Set element = dom.createElement("city")
    address.appendChild element
    element.Text = issuerArray(7)
    
    'Invoice > Counterpart
    Set counterPart = dom.createElement("counterpart")
    invoice.appendChild counterPart
    Set element = dom.createElement("vatNumber")
    counterPart.appendChild element
    element.Text = counterPartArray(0)
    Set element = dom.createElement("country")
    counterPart.appendChild element
    element.Text = counterPartArray(1)
    Set element = dom.createElement("branch")
    counterPart.appendChild element
    element.Text = counterPartArray(2)
    Set element = dom.createElement("name")
    counterPart.appendChild element
    element.Text = counterPartArray(3)
    
    'Invoice > Counterpart > Address
    Set address = dom.createElement("address")
    counterPart.appendChild address
    Set element = dom.createElement("street")
    address.appendChild element
    element.Text = counterPartArray(4)
    Set element = dom.createElement("number")
    address.appendChild element
    element.Text = counterPartArray(5)
    Set element = dom.createElement("postalCode")
    address.appendChild element
    element.Text = counterPartArray(6)
    Set element = dom.createElement("city")
    address.appendChild element
    element.Text = counterPartArray(7)
    
    'Invoice > Invoice header
    Set invoiceHeader = dom.createElement("invoiceHeader")
    invoice.appendChild invoiceHeader
    Set element = dom.createElement("series")
    invoiceHeader.appendChild element
    element.Text = invoiceHeaderArray(0)
    Set element = dom.createElement("aa")
    invoiceHeader.appendChild element
    element.Text = invoiceHeaderArray(1)
    Set element = dom.createElement("issueDate")
    invoiceHeader.appendChild element
    element.Text = invoiceHeaderArray(2)
    Set element = dom.createElement("invoiceType")
    invoiceHeader.appendChild element
    element.Text = invoiceHeaderArray(3)
    Set element = dom.createElement("currency")
    invoiceHeader.appendChild element
    element.Text = invoiceHeaderArray(4)
    
    'Invoice > Payment method
    Set invoicePaymentMethod = dom.createElement("paymentMethods")
    invoice.appendChild invoicePaymentMethod
    
    'Invoice > Payment method > Details
    Set invoicePaymentMethodDetails = dom.createElement("paymentMethodDetails")
    invoicePaymentMethod.appendChild invoicePaymentMethodDetails
    Set element = dom.createElement("type")
    invoicePaymentMethodDetails.appendChild element
    element.Text = paymentMethodDetailsArray(0)
    Set element = dom.createElement("amount")
    invoicePaymentMethodDetails.appendChild element
    element.Text = paymentMethodDetailsArray(1)
        
    'Invoice > Details
    Dim detailLine As Integer
    For detailLine = 0 To UBound(invoiceDetailsArray) - 1
        
        Set invoiceDetails = dom.createElement("invoiceDetails")
        invoice.appendChild invoiceDetails
        
        Set element = dom.createElement("lineNumber")
        invoiceDetails.appendChild element
        element.Text = invoiceDetailsArray(detailLine, 0)
        
        Set element = dom.createElement("netValue")
        invoiceDetails.appendChild element
        element.Text = invoiceDetailsArray(detailLine, 1)
        
        Set element = dom.createElement("vatCategory")
        invoiceDetails.appendChild element
        element.Text = invoiceDetailsArray(detailLine, 2)
        
        Set element = dom.createElement("vatAmount")
        invoiceDetails.appendChild element
        element.Text = invoiceDetailsArray(detailLine, 3)
        
        Set incomeClassifications = dom.createElement("incomeClassification")
        invoiceDetails.appendChild incomeClassifications
        
        Set element = dom.createElement("classificationType")
        incomeClassifications.appendChild element
        element.Text = "E3_106"
        
    Next detailLine
    
    'Invoice > Summary
    Set invoiceSummary = dom.createElement("invoiceSummary")
    invoice.appendChild invoiceSummary
    Set element = dom.createElement("totalNetValue")
    invoiceSummary.appendChild element
    element.Text = invoiceSummaryArray(0)
    Set element = dom.createElement("totalVatAmount")
    invoiceSummary.appendChild element
    element.Text = invoiceSummaryArray(1)
    Set element = dom.createElement("totalWithheldAmount")
    invoiceSummary.appendChild element
    element.Text = invoiceSummaryArray(2)
    Set element = dom.createElement("totalFeesAmount")
    invoiceSummary.appendChild element
    element.Text = invoiceSummaryArray(3)
    Set element = dom.createElement("totalStampDutyAmount")
    invoiceSummary.appendChild element
    element.Text = invoiceSummaryArray(4)
    Set element = dom.createElement("totalOtherTaxesAmount")
    invoiceSummary.appendChild element
    element.Text = invoiceSummaryArray(5)
    Set element = dom.createElement("totalDeductionsAmount")
    invoiceSummary.appendChild element
    element.Text = invoiceSummaryArray(6)
    Set element = dom.createElement("totalGrossValue")
    invoiceSummary.appendChild element
    element.Text = invoiceSummaryArray(7)

    dom.save ("d:\API Client\Export.xml")
    
    CreateRequestBody = dom.xml
   
End Function


Private Sub cmdCreate_Click()

    Dim myMSXML
    Dim textJSON
    
    Set myMSXML = CreateObject("Microsoft.XmlHttp")
    
    textJSON = "{ ""abbreviation"":""new.."", ""description"":""new from vb.."",""isActive"":""true"",""userId"":""e7e014fd-5608-4936-866e-ec11fc8c16da""}"
    
    myMSXML.Open "POST", "https://www.appcorfucruises.com/api/destinations", False
    myMSXML.SetRequestHeader "Content-Type", "application/json"
    myMSXML.Send textJSON
    
    MsgBox myMSXML.ResponseText

End Sub

Private Sub cmdCreateMyData_Click()

    Dim myMSXML
    Dim textJSON
    Dim body As String
    
    Set myMSXML = CreateObject("Microsoft.XmlHttp")
    
    myMSXML.Open "POST", "https://mydatapi.aade.gr/mydata/SendInvoices", False
    
    myMSXML.SetRequestHeader "aade-user-id", "krotsismepe"
    myMSXML.SetRequestHeader "ocp-apim-subscription-key", "e3ab4ffa43f64fc2baee668890d1c804"
    
    body = CreateRequestBody(CreateIssuer, CreateCounterPart, CreatePaymentMethod, CreatePaymentMethodDetails, CreateInvoiceHeader, CreateInvoiceDetails, CreateInvoiceSummary)
    
    Debug.Print body
    
    'strSend = "<?xml version='1.0' encoding='utf-8'?><sysbus><auth><key>ABC123</key></auth></sysbus>"
    
    'myMSXML.Send textJSON
    
    'MsgBox myMSXML.ResponseText

End Sub


Private Sub cmdGetById_Click()

    Dim response As String
    
    winH.Open "get", "https://www.appcorfucruises.com/api/destinations/" + txtId.Text
    winH.Send
    
    response = winH.ResponseText
    
    txtResults.Text = response

End Sub

Private Sub cmdGet_Click()

    Dim response As String
    
    winH.Open "get", "https://www.appcorfucruises.com/api/destinations"
    winH.Send
    
    response = winH.ResponseText
    
    txtResults.Text = response


End Sub

Private Sub createXML_Click()

    Dim result As String
    
    result = CreateRequestBody(CreateIssuer, CreateCounterPart, CreatePaymentMethod, CreatePaymentMethodDetails, CreateInvoiceHeader, CreateInvoiceDetails, CreateInvoiceSummary)
    
    Debug.Print result

End Sub


Private Sub Form_Load()

    Set winH = New WinHttp.WinHttpRequest

End Sub


