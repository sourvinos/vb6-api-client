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
   Begin Dacara_dcButton.dcButton cmdCreateXML 
      Height          =   465
      Left            =   150
      TabIndex        =   5
      Top             =   3075
      Width           =   1515
      _ExtentX        =   2672
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
   Begin Dacara_dcButton.dcButton dcButton1 
      Height          =   465
      Left            =   1725
      TabIndex        =   6
      Top             =   3075
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonStyle     =   8
      Caption         =   "Create original XML"
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

Private Sub cmdCreateXML_Click()

    Dim objDom As DOMDocument
    Dim root As IXMLDOMElement
    Dim invoice As IXMLDOMElement
    Dim issuer As IXMLDOMElement
    Dim counterPart As IXMLDOMElement
    Dim address As IXMLDOMElement
    Dim element As IXMLDOMElement
    Dim objMemberRel As IXMLDOMAttribute
    Dim objMemberElem As IXMLDOMElement
    
    Set objDom = New DOMDocument
    
    'Root
    Set root = objDom.createElement("InvoicesDoc")
    objDom.appendChild root
    Set objMemberRel = objDom.createAttribute("Relationship")
    objMemberRel.nodeValue = "Father"
    objMemberElem.setAttributeNode objMemberRel
    'Invoice
    Set invoice = objDom.createElement("invoice")
    root.appendChild invoice
    'Invoice > …ssuer
    Set issuer = objDom.createElement("issuer")
    invoice.appendChild issuer
    Set element = objDom.createElement("vatNumber")
    issuer.appendChild element
    element.Text = "099863549"
    Set element = objDom.createElement("country")
    issuer.appendChild element
    element.Text = "GR"
    Set element = objDom.createElement("branch")
    issuer.appendChild element
    element.Text = "0"
    Set element = objDom.createElement("name")
    issuer.appendChild element
    element.Text = " —œ‘”«” Ã.≈.–.≈."
    'Invoice > …ssuer > Address
    Set address = objDom.createElement("address")
    issuer.appendChild address
    Set element = objDom.createElement("street")
    address.appendChild element
    element.Text = "≈»Õ… « œƒœ”  ≈— ’—¡” - À≈’ …ÃÃ«”"
    Set element = objDom.createElement("number")
    address.appendChild element
    element.Text = "17A"
    Set element = objDom.createElement("postalCode")
    address.appendChild element
    element.Text = "491 00"
    Set element = objDom.createElement("city")
    address.appendChild element
    element.Text = " ≈— ’—¡"
    'Invoice > Counterpart
    Set counterPart = objDom.createElement("counterpart")
    invoice.appendChild counterPart
    Set element = objDom.createElement("vatNumber")
    counterPart.appendChild element
    element.Text = "99999999"
    Set element = objDom.createElement("country")
    counterPart.appendChild element
    element.Text = "EL"
    Set element = objDom.createElement("branch")
    counterPart.appendChild element
    element.Text = "1"
    Set element = objDom.createElement("name")
    counterPart.appendChild element
    element.Text = "« ≈–ŸÕ’Ã…¡ ‘œ’ ¡ÀÀœ’!"
    'Invoice > CounterPart > Address
    Set address = objDom.createElement("address")
    counterPart.appendChild address
    Set element = objDom.createElement("street")
    address.appendChild element
    element.Text = " ¡¬œ”"
    Set element = objDom.createElement("number")
    address.appendChild element
    element.Text = "-"
    Set element = objDom.createElement("postalCode")
    address.appendChild element
    element.Text = "490 84"
    Set element = objDom.createElement("city")
    address.appendChild element
    element.Text = " ¡¬œ”"
       
    objDom.save ("d:\API Client\Export.xml")
    
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

    Dim objDom As DOMDocument
    Dim objRootElem As IXMLDOMElement
    Dim objMemberElem As IXMLDOMElement
    Dim objMemberRel As IXMLDOMAttribute
    Dim objMemberName As IXMLDOMElement
    
    Set objDom = New DOMDocument
    
    ' Creates root element
    Set objRootElem = objDom.createElement("InvoicesDoc")
    objDom.appendChild objRootElem
    
    ' Creates Attribute to the Root Element
    Set objMemberRel = objDom.createAttribute("xmlns")
    objMemberRel.nodeValue = "http://www.aade.gr/myDATA/invoice/v1.0"
    objRootElem.setAttributeNode objMemberRel
    
    ' Creates Attribute to the Root Element
    Set objMemberRel = objDom.createAttribute("xmlns:xsi")
    objMemberRel.nodeValue = "http://www.w3.org/2001/XMLSchema-instance"
    objRootElem.setAttributeNode objMemberRel
    
    ' Creates Attribute to the Root Element
    Set objMemberRel = objDom.createAttribute("xsi:schemaLocation")
    objMemberRel.nodeValue = "http://www.aade.gr/myDATA/invoice/v1.0 schema.xsd"
    objRootElem.setAttributeNode objMemberRel
    
    ' Creates Member element
    'Set objMemberElem = objDom.createElement("Member")
    'objRootElem.appendChild objMemberElem
    
    ' Creates Attribute to the Member Element
    'Set objMemberRel = objDom.createAttribute("xmlns")
    'objMemberRel.nodeValue = "http://www.aade.gr/myDATA/invoice/v1.0"
    'objMemberElem.setAttributeNode objMemberRel
    
    ' Saves XML data to disk.
    objDom.save ("d:\API Client\Export.xml")
    
End Sub

Private Sub Form_Load()

    Set winH = New WinHttp.WinHttpRequest

End Sub


