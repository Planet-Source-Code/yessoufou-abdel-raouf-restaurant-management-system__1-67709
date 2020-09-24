VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOrderDetails 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lv 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Unit Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Discount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   5070
      TabIndex        =   3
      Top             =   4290
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   16777215
      cBhover         =   16711680
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmOrderDetails.frx":0000
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   435
      Left            =   720
      TabIndex        =   4
      Top             =   4290
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      Caption         =   "&Print"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   16777215
      cBhover         =   16711680
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmOrderDetails.frx":031A
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   8175
   End
   Begin VB.Label lblNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "THE DETAILS OF ORDER NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00400000&
      Height          =   4215
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmOrderDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************************
'Title: RestauSys                                             *
'application domain: Restaurant Management System             *
'Description:                                                 *
'Version: 1.0                                                 *
'Author: YESSOUFOU Abdel Raouf                                *
'Copyright: All right reserved by the author Abdel Raouf.     *
'   No part of this source code shall be copied or reused     *
'   without the author's notice.2005-2006                     *
'**************************************************************
Option Explicit
    Dim sql As String
Private Sub cmdClose_Click()

    If blOrderDetails = True Then
        Load frmCustomerAndOrder
        frmCustomerAndOrder.Show
        Unload Me
    End If
    If blOrderDetails1 = True Then
        Load frmCredit
        frmCredit.Show
        Unload Me
    End If
    If blInvoice = True Then
        Load frmAllOrderPayment
        frmAllOrderPayment.Show
        Unload Me
    End If
End Sub


Private Sub cmdPrint_Click()
On Error GoTo abdel
    If blOrderDetails1 = True And blInvoice = False Then
        CreditDetailsPrint.Sections("section2").Controls.Item("lblID").Caption = Trim(lblNo.Caption)
        sql = "select * from orderdetails where orderid = " & Trim(lblNo.Caption)
        Set CreditDetailsPrint.DataSource = con.Execute(sql)
        CreditDetailsPrint.Show
        Set CreditDetailsPrint = Nothing
        Exit Sub
    End If
    If blOrderDetails1 = False And blInvoice = True Then
        CreditDetailsPrint.Sections("section2").Controls.Item("lblID").Caption = Trim(lblNo.Caption)
        sql = "select * from orderdetails where orderid = " & Trim(lblNo.Caption)
        Set CreditDetailsPrint.DataSource = con.Execute(sql)
        CreditDetailsPrint.Show
        Set CreditDetailsPrint = Nothing
        Exit Sub
    End If
    
    Exit Sub
abdel:
    MsgBox "Report not available for now...", vbExclamation, title
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    blOrderDetails = False
    blOrderDetails1 = False
    blInvoice = False
    frmScreen.Enabled = True
    
End Sub

