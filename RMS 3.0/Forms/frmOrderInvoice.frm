VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOrderInvoice 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "PAIEMENT"
   ClientHeight    =   7920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List5 
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
      Left            =   6240
      TabIndex        =   24
      Top             =   3840
      Width           =   1575
   End
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
      Left            =   4800
      TabIndex        =   23
      Top             =   3840
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
      Left            =   3360
      TabIndex        =   22
      Top             =   3840
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
      Left            =   2040
      TabIndex        =   21
      Top             =   3840
      Width           =   1335
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
      Left            =   360
      TabIndex        =   20
      Top             =   3840
      Width           =   1695
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
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   2280
   End
   Begin VB.ComboBox com 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmOrderInvoice.frx":0000
      Left            =   4200
      List            =   "frmOrderInvoice.frx":000A
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtChange 
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
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtCheckNo 
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
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtReste 
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
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtAmountPaid 
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
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtOrderID 
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
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin MSComctlLib.ListView lv 
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   3480
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Name"
         Object.Width           =   2540
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
         Text            =   "Discount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   435
      Left            =   3030
      TabIndex        =   25
      Top             =   7170
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
      mIcon           =   "frmOrderInvoice.frx":001B
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
      Left            =   5640
      TabIndex        =   12
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label lblPaymentID 
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
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4320
      TabIndex        =   17
      Top             =   360
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
      Left            =   4560
      TabIndex        =   16
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Order ID"
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
      Left            =   720
      TabIndex        =   13
      Top             =   840
      Width           =   1455
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
      Left            =   600
      TabIndex        =   11
      Top             =   240
      Width           =   1935
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
      Left            =   4440
      TabIndex        =   10
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Reste To Pay"
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
      Left            =   4440
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
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
      Left            =   600
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Check No"
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
      Left            =   600
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Mode"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INVOICE NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   7575
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderWidth     =   4
      Height          =   7935
      Left            =   0
      Top             =   0
      Width           =   8415
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   6720
      Width           =   2295
   End
End
Attribute VB_Name = "frmOrderInvoice"
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
    If com.Text = "" Then
        MsgBox "Kindly select the payment mode", vbCritical, title
        com.SetFocus
        Exit Sub
    End If
    
    If txtAmountPaid.Text = "" Then
        MsgBox "Kindly specify the amount paid by the customer", vbCritical, title
        txtAmountPaid.SetFocus
        Exit Sub
    End If
    
    recPayment.AddNew
    recPayment!PaymentID = lblPaymentID.Caption & ""
    recPayment!orderid = txtOrderID.Text & ""
    recPayment!PaymentMode = com.Text & ""
    recPayment!checkNo = txtCheckNo.Text & ""
    recPayment!amountpaid = txtAmountPaid.Text & ""
    If txtChange.Text = "" Then
        recPayment!Change = 0
        Else
            recPayment!Change = txtChange.Text & ""
    End If
    recPayment!dates = lblDate.Caption & ""
    recPayment!Time = lblTime.Caption & ""
    recPayment.Update
    
    If txtReste > 0 Then
        recCredit.AddNew
        recCredit(0) = autogeneration
        recCredit(1) = txtOrderID.Text
        recCredit(2) = lblName.Caption
        recCredit(3) = txtAmountDue.Text
        recCredit(4) = txtAmountPaid.Text
        recCredit(5) = txtReste.Text
        recCredit.Update
    End If
    
    OrderPrint.Sections("section2").Controls.Item("lblID").Caption = txtOrderID.Text
    OrderPrint.Sections("section5").Controls.Item("lblAmountDue").Caption = txtAmountDue.Text
    OrderPrint.Sections("section5").Controls.Item("lblCredit").Caption = txtReste.Text
    OrderPrint.Sections("section5").Controls.Item("lblAmountPaid").Caption = txtAmountPaid.Text
    OrderPrint.Sections("section5").Controls.Item("lblChanges").Caption = txtChange.Text
    OrderPrint.Sections("section5").Controls.Item("lblMode").Caption = com.Text
    OrderPrint.Sections("section5").Controls.Item("lblCheckNo").Caption = txtCheckNo.Text
    recOrder.Requery
    sql = "select * from orderdetails where OrderID = " & Trim(txtOrderID.Text)
    Set OrderPrint.DataSource = con.Execute(sql)
    OrderPrint.Show
    Set OrderPrint = Nothing
    
    Unload Me
    Unload frmOrder
    
    Exit Sub
abdel:
    MsgBox "Invoice not available for now...", vbExclamation, title
End Sub

Private Sub com_Click()

    If com.Text = "CASH" Then
        txtCheckNo.Enabled = False
        txtAmountPaid.SetFocus
    End If
    If com.Text = "CHECK" Then
        txtCheckNo.Enabled = True
        txtCheckNo.SetFocus
    End If
End Sub

Private Sub com_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
End Sub

Private Sub Form_Load()

    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    lblName.Visible = False
    lblPaymentID.Caption = autogen
    txtOrderID.Text = recOrder(0)
    recOrderDetails.MoveFirst
    txtAmountDue.Text = frmOrder.total.Caption
    txtCheckNo.Enabled = False
    
End Sub

Public Function autogen()

    Dim recGen As New Recordset
    
    recGen.Open "select max(paymentid) from payment", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogen = 1
        
        Else
        
        autogen = Val(recGen(0) + 1)
    End If
End Function

Public Function autogeneration()


    Dim recGen As New Recordset
    
    recGen.Open "select max(creditID) from credit", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogeneration = 1
        
        Else
        
        autogeneration = Val(recGen(0) + 1)
    End If
End Function

Private Sub Timer1_Timer()

    lblDate.Caption = Date
    lblTime.Caption = Time
End Sub

Private Sub txtChange_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub


Private Sub txtReste_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtAmountPaid_Change()

    If Val(txtAmountDue.Text) > Val(txtAmountPaid.Text) Then
        txtReste.Text = Val(txtAmountDue.Text) - Val(txtAmountPaid.Text)
        txtChange.Text = 0
        Else
            txtChange.Text = Val(txtAmountPaid.Text) - Val(txtAmountDue.Text)
            txtReste.Text = 0
    End If
End Sub

Private Sub txtAmountPaid_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub
