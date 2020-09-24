VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSalesInvoice 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "PAIEMENT"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   6120
      TabIndex        =   14
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4680
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3240
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox txtAmountDue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   3600
   End
   Begin VB.TextBox txtChange 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox txtCash 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      MaxLength       =   10
      TabIndex        =   2
      Top             =   4560
      Width           =   1455
   End
   Begin MSComctlLib.ListView lv 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Name"
         Object.Width           =   5469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Unit Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   435
      Left            =   2940
      TabIndex        =   15
      Top             =   6360
      Width           =   2115
      _ExtentX        =   3731
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
      mIcon           =   "frmSalesInvoice.frx":0000
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   6120
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Shape Shape5 
      Height          =   255
      Left            =   120
      Top             =   4320
      Width           =   7455
   End
   Begin VB.Shape Shape4 
      Height          =   375
      Left            =   120
      Top             =   5040
      Width           =   7455
   End
   Begin VB.Shape Shape3 
      Height          =   375
      Left            =   120
      Top             =   4560
      Width           =   7455
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   120
      Top             =   3960
      Width           =   7455
   End
   Begin VB.Label lblInvoiceId 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Due"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Paid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INVOICE NO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      Height          =   6975
      Left            =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmSalesInvoice"
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
    Dim reply As Integer
    Dim sql As String
Private Sub cmdPrint_Click()
On Error GoTo abdel
    If txtCash.Text = "" Then
        MsgBox "What is the amount paid by the customer ?", vbQuestion, title
        txtCash.SetFocus
        Exit Sub
    End If
    
    recSalesInvoice!invoiceid = lblInvoiceId.Caption
    recSalesInvoice!amountdue = txtAmountDue.Text
    recSalesInvoice!amountpaid = txtCash.Text
    recSalesInvoice!changes = txtChange.Text
    recSalesInvoice!Date = lblDate.Caption
    recSalesInvoice!Time = lblTime.Caption
    recSalesInvoice.UpdateBatch adAffectCurrent
    
    SalesPrint.Sections("section2").Controls.Item("lblID").Caption = lblInvoiceId.Caption
    SalesPrint.Sections("section5").Controls.Item("lblAmountDue").Caption = txtAmountDue.Text
    SalesPrint.Sections("section5").Controls.Item("lblAmountPaid").Caption = txtCash.Text
    SalesPrint.Sections("section5").Controls.Item("lblChanges").Caption = txtChange.Text
    recSale.Requery
    sql = "select * from sale where invoiceid = " & Trim(SalesPrint.Sections("section2").Controls.Item("lblID").Caption)
    Set SalesPrint.DataSource = con.Execute(sql)
    SalesPrint.Show
    Set SalesPrint = Nothing
    
    Unload Me
    Unload frmSale
    Exit Sub
abdel:
    MsgBox "Invoice not available for now...", vbExclamation, title
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
End Sub

Private Sub Timer1_Timer()
    lblDate.Caption = Date
    lblTime.Caption = Time
End Sub

Private Sub txtAmountDue_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCash_Change()
    If txtCash.Text = "" Then
        txtChange.Text = ""
        Else
            txtChange.Text = Val(txtCash.Text) - Val(txtAmountDue.Text)
    End If
End Sub

Private Sub txtCash_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtChange_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub
