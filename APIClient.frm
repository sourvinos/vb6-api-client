VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Begin VB.Form APIClient 
   Caption         =   "API Client"
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8340
   Icon            =   "APIClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
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
   Begin Dacara_dcButton.dcButton cmdRead 
      Height          =   465
      Left            =   150
      TabIndex        =   0
      Top             =   675
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonStyle     =   8
      Caption         =   "Read"
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
   Begin Dacara_dcButton.dcButton cmdLogin 
      Height          =   465
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonStyle     =   8
      Caption         =   "Login"
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

Private Sub cmdLogin_Click()

    Dim response As String
    
    winH.Open "post", "https://www.appcorfucruises.com/login"
    winH.Send
    
    response = winH.ResponseText
    
    txtResults.Text = response


End Sub

Private Sub cmdRead_Click()

    Dim response As String
    
    winH.Open "get", "https://www.appcorfucruises.com/api/customers"
    winH.Send
    
    response = winH.ResponseText
    
    txtResults.Text = response

End Sub

Private Sub Form_Load()

    Set winH = New WinHttp.WinHttpRequest

End Sub


