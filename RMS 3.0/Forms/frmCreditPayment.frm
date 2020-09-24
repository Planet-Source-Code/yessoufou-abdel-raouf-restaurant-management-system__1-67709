VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCreditPayment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
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
      Height          =   360
      Left            =   1800
      TabIndex        =   17
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtID 
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
      Height          =   360
      Left            =   2280
      TabIndex        =   5
      Top             =   1200
      Width           =   735
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
      Height          =   360
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2400
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
      Height          =   360
      Left            =   1800
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
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
      Height          =   360
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
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
      ItemData        =   "frmCreditPayment.frx":0000
      Left            =   3960
      List            =   "frmCreditPayment.frx":000A
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   2640
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
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   435
      Left            =   4410
      TabIndex        =   19
      Top             =   5340
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   767
      Caption         =   "Cancel"
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
      mIcon           =   "frmCreditPayment.frx":001B
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   435
      Left            =   1770
      TabIndex        =   20
      Top             =   5340
      Width           =   1665
      _ExtentX        =   2937
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
      mIcon           =   "frmCreditPayment.frx":0335
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2760
      TabIndex        =   18
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label I 
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
      Left            =   2400
      TabIndex        =   16
      Top             =   600
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
      Left            =   1800
      TabIndex        =   15
      Top             =   1800
      Width           =   2055
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
      Left            =   240
      TabIndex        =   14
      Top             =   2400
      Width           =   1455
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
      TabIndex        =   13
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Owed"
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
      Left            =   4200
      TabIndex        =   12
      Top             =   2400
      Width           =   1575
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
      Left            =   4200
      TabIndex        =   11
      Top             =   3120
      Width           =   975
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
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   5040
      Width           =   1935
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
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
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
      Left            =   4200
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblInvoiceNo 
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
      Left            =   3960
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   5895
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   7815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   6615
      Left            =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmCreditPayment"
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
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error GoTo abdel
    If com.Text = "" Then
        MsgBox "Please select the payment mode", vbCritical, title
        com.SetFocus
        Exit Sub
    End If
    
    If txtAmountPaid.Text = "" Then
        MsgBox "What is the amount paid by the customer", vbCritical, title
        txtAmountPaid.SetFocus
        Exit Sub
    End If
    recPayment.AddNew
    recPayment!PaymentID = lblInvoiceNo.Caption & ""
    recPayment!orderid = txtID.Text & ""
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
        sql = "update credit set amountowed = " & Val(txtReste.Text)
        sql = sql & " where orderid = " & Trim(txtID.Text)
        con.Execute sql
        sql1 = "update credit set amountpaid = " & Val(recCredit!amountpaid) + Val(txtAmountPaid.Text)
        sql1 = sql1 & " where orderid = " & Trim(txtID.Text)
        con.Execute sql1
        Else
            con.Execute "delete from credit where orderid = " & Trim(txtID.Text)
    End If
    CreditPaymentInvoice.Sections("section5").Controls.Item("lblInvoiceNo").Caption = lblInvoiceNo.Caption
    CreditPaymentInvoice.Sections("section5").Controls.Item("lblOrderNo").Caption = txtID.Text
    CreditPaymentInvoice.Sections("section5").Controls.Item("lblAmountDue").Caption = txtAmountDue.Text
    CreditPaymentInvoice.Sections("section5").Controls.Item("lblCredit").Caption = txtReste.Text
    CreditPaymentInvoice.Sections("section5").Controls.Item("lblAmountPaid").Caption = txtAmountPaid.Text
    CreditPaymentInvoice.Sections("section5").Controls.Item("lblChanges").Caption = txtChange.Text
    CreditPaymentInvoice.Sections("section5").Controls.Item("lblMode").Caption = com.Text
    CreditPaymentInvoice.Sections("section5").Controls.Item("lblCheckNo").Caption = txtCheckNo.Text
    recOrder.Requery
    sql = "select * from orderdetails where OrderID = " & Trim(txtID.Text)
    Set CreditPaymentInvoice.DataSource = con.Execute(sql)
    CreditPaymentInvoice.Show
    Set CreditPaymentInvoice = Nothing
    
    Unload Me
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
    recCredit.Requery
    Call ConnectMe
    lblInvoiceNo.Caption = autogen
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
