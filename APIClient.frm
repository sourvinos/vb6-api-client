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
   Begin Dacara_dcButton.dcButton dcButton1 
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
    
    Set myMSXML = CreateObject("Microsoft.XmlHttp")
    
    textJSON = "{ ""abbreviation"":""new.."", ""description"":""new from vb.."",""isActive"":""true"",""userId"":""e7e014fd-5608-4936-866e-ec11fc8c16da""}"
    
    myMSXML.Open "POST", "https://mydata-dev.azure-api.net/SendInvoices", False
    
    myMSXML.SetRequestHeader "aade-user-id", "krotsisepe"
    myMSXML.SetRequestHeader "ocp-apim-subscription-key", "1f8476ce37534742886b2009739bd6ad"
    
    'strSend = "<?xml version='1.0' encoding='utf-8'?><sysbus><auth><key>ABC123</key></auth></sysbus>"
    
    myMSXML.Send textJSON
    
    MsgBox myMSXML.ResponseText


End Sub


Private Sub cmdCreateXML_Click()

    Dim dom As DOMDocument
    
    Dim root As IXMLDOMElement
    
    Dim invoice As IXMLDOMElement
    Dim issuer As IXMLDOMElement
    Dim counterpart As IXMLDOMElement
    Dim address As IXMLDOMElement
    Dim invoiceHeader As IXMLDOMElement
    
    Dim element As IXMLDOMElement
    Dim objAttribute As IXMLDOMAttribute
    Dim objMemberElem As IXMLDOMElement
    
    Set dom = New DOMDocument
    
    'Root
    Set root = dom.createElement("InvoicesDoc")
    dom.appendChild root
    Set objAttribute = dom.createAttribute("Relationship")
    objAttribute.nodeValue = "Father"
    objMemberElem.setAttributeNode objAttribute
    'Invoice
    Set invoice = dom.createElement("invoice")
    root.appendChild invoice
    'Invoice > Éssuer
    Set issuer = dom.createElement("issuer")
    invoice.appendChild issuer
    Set element = dom.createElement("vatNumber")
    issuer.appendChild element
    element.Text = "099863549"
    Set element = dom.createElement("country")
    issuer.appendChild element
    element.Text = "GR"
    Set element = dom.createElement("branch")
    issuer.appendChild element
    element.Text = "0"
    Set element = dom.createElement("name")
    issuer.appendChild element
    element.Text = "ÊÑÏÔÓÇÓ Ì.Å.Ð.Å."
    'Invoice > Éssuer > Address
    Set address = dom.createElement("address")
    issuer.appendChild address
    Set element = dom.createElement("street")
    address.appendChild element
    element.Text = "ÅÈÍÉÊÇ ÏÄÏÓ ÊÅÑÊÕÑÁÓ - ËÅÕÊÉÌÌÇÓ"
    Set element = dom.createElement("number")
    address.appendChild element
    element.Text = "17A"
    Set element = dom.createElement("postalCode")
    address.appendChild element
    element.Text = "491 00"
    Set element = dom.createElement("city")
    address.appendChild element
    element.Text = "ÊÅÑÊÕÑÁ"
    'Invoice > Counterpart
    Set counterpart = dom.createElement("counterpart")
    invoice.appendChild counterpart
    Set element = dom.createElement("vatNumber")
    counterpart.appendChild element
    element.Text = "99999999"
    Set element = dom.createElement("country")
    counterpart.appendChild element
    element.Text = "EL"
    Set element = dom.createElement("branch")
    counterpart.appendChild element
    element.Text = "1"
    Set element = dom.createElement("name")
    counterpart.appendChild element
    element.Text = "Ç ÅÐÙÍÕÌÉÁ ÔÏÕ ÁËËÏÕ!"
    'Invoice > CounterPart > Address
    Set address = dom.createElement("address")
    counterpart.appendChild address
    Set element = dom.createElement("street")
    address.appendChild element
    element.Text = "ÊÁÂÏÓ"
    Set element = dom.createElement("number")
    address.appendChild element
    element.Text = "-"
    Set element = dom.createElement("postalCode")
    address.appendChild element
    element.Text = "490 84"
    Set element = dom.createElement("city")
    address.appendChild element
    element.Text = "ÊÁÂÏÓ"
    'Invoice > InvoiceHeader
    Set invoiceHeader = dom.createElement("invoiceHeader")
    invoice.appendChild invoiceHeader
    Set element = dom.createElement("series")
    invoiceHeader.appendChild element
    element.Text = "53"
       
    Debug.Print dom
    dom.save ("d:\API Client\Export.xml")
    
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

Private Sub dcButton1_Click()

    Dim dom As DOMDocument
    Dim rootElement As IXMLDOMElement
    Dim objElement As IXMLDOMElement
    Dim objAttribute As IXMLDOMAttribute
    
    Dim root As IXMLDOMElement
    Dim invoice As IXMLDOMElement
    Dim issuer As IXMLDOMElement
    Dim address As IXMLDOMElement
    Dim element As IXMLDOMElement
    Dim counterpart As IXMLDOMElement
    Dim invoiceHeader As IXMLDOMElement
    Dim invoiceDetails As IXMLDOMElement
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
    element.Text = "099863549"
    Set element = dom.createElement("country")
    issuer.appendChild element
    element.Text = "GR"
    Set element = dom.createElement("branch")
    issuer.appendChild element
    element.Text = "0"
    Set element = dom.createElement("name")
    issuer.appendChild element
    element.Text = "ÊÑÏÔÓÇÓ Ì.Å.Ð.Å."
    
    'Invoice > Éssuer > Address
    Set address = dom.createElement("address")
    issuer.appendChild address
    
    Set element = dom.createElement("street")
    address.appendChild element
    element.Text = "ÏÄÏÓ"
    
    Set element = dom.createElement("number")
    address.appendChild element
    element.Text = "74"
    
    Set element = dom.createElement("postalcode")
    address.appendChild element
    element.Text = "491 00"
    
    Set element = dom.createElement("city")
    address.appendChild element
    element.Text = "ÊÅÑÊÕÑÁ"
    
    'Invoice > Counterpart
    Set counterpart = dom.createElement("counterpart")
    invoice.appendChild counterpart
    
    Set element = dom.createElement("vatNumber")
    counterpart.appendChild element
    element.Text = "099863549"
    Set element = dom.createElement("country")
    counterpart.appendChild element
    element.Text = "GR"
    Set element = dom.createElement("branch")
    counterpart.appendChild element
    element.Text = "0"
    Set element = dom.createElement("name")
    counterpart.appendChild element
    element.Text = "ÊÑÏÔÓÇÓ Ì.Å.Ð.Å."
    
    'Invoice > Counterpart > Address
    Set address = dom.createElement("address")
    counterpart.appendChild address
    
    Set element = dom.createElement("street")
    address.appendChild element
    element.Text = "ÏÄÏÓ"
    
    Set element = dom.createElement("number")
    address.appendChild element
    element.Text = "74"
    
    Set element = dom.createElement("postalcode")
    address.appendChild element
    element.Text = "491 00"
    
    Set element = dom.createElement("city")
    address.appendChild element
    element.Text = "ÊÅÑÊÕÑÁ"
    
    'Invoice > InvoiceHeader
    Set invoiceHeader = dom.createElement("invoiceHeader")
    invoice.appendChild invoiceHeader
    
    Set element = dom.createElement("series")
    invoiceHeader.appendChild element
    element.Text = "35A"
    
    Set element = dom.createElement("aa")
    invoiceHeader.appendChild element
    element.Text = "1"
    
    Set element = dom.createElement("issueDate")
    invoiceHeader.appendChild element
    element.Text = "2021-11-01"
    
    Set element = dom.createElement("invoiceType")
    invoiceHeader.appendChild element
    element.Text = "2.1"
    
    'Invoice > Details
    Set invoiceDetails = dom.createElement("invoiceDetails")
    invoice.appendChild invoiceDetails
   
    Set element = dom.createElement("lineNumber")
    invoiceDetails.appendChild element
    element.Text = "1"
    
    Set element = dom.createElement("netValue")
    invoiceDetails.appendChild element
    element.Text = "0.00"
    
    Set element = dom.createElement("vatCategory")
    invoiceDetails.appendChild element
    element.Text = "1"
    
    Set element = dom.createElement("vatAmount")
    invoiceDetails.appendChild element
    element.Text = "0.00"
   
    'Invoice > Summary
    Set invoiceSummary = dom.createElement("invoiceSummary")
    invoice.appendChild invoiceSummary
   
    Set element = dom.createElement("totalNetValue")
    invoiceSummary.appendChild element
    element.Text = "100.00"
    
    Set element = dom.createElement("totalVatAmount")
    invoiceSummary.appendChild element
    element.Text = "24"
    
    Set element = dom.createElement("totalWithheldAmount")
    invoiceSummary.appendChild element
    element.Text = "0"
    
    Set element = dom.createElement("totalFeesAmount")
    invoiceSummary.appendChild element
    element.Text = "0"
    
    Set element = dom.createElement("totalStampDutyAmount")
    invoiceSummary.appendChild element
    element.Text = "0"
    
    Set element = dom.createElement("totalOtherTaxesAmount")
    invoiceSummary.appendChild element
    element.Text = "0"
    
    Set element = dom.createElement("totalDeductionsAmount")
    invoiceSummary.appendChild element
    element.Text = "0"
    
    Set element = dom.createElement("totalGrossValue")
    invoiceSummary.appendChild element
    element.Text = "124.00"
   
    'Saves XML data to disk.
    'Debug.Print dom.xml
    
    dom.save ("d:\API Client\Export.xml")
    
End Sub

Private Sub dcButton2_Click()

End Sub

Private Sub Form_Load()

    Set winH = New WinHttp.WinHttpRequest

End Sub


