VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllOrderPayment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lv 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double click to view details"
      Top             =   720
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8281
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Invoice No"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Order ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Payment Mode"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Check No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Amount Paid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   6630
      TabIndex        =   2
      Top             =   5850
      Width           =   2115
      _ExtentX        =   3731
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
      mIcon           =   "frmAllOrderPayment.frx":0000
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   435
      Left            =   3360
      TabIndex        =   3
      Top             =   5850
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
      mIcon           =   "frmAllOrderPayment.frx":031A
   End
   Begin lvButton.lvButtons_H cmdView 
      Height          =   435
      Left            =   240
      TabIndex        =   4
      Top             =   5850
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   767
      Caption         =   "&View Order Details"
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
      mIcon           =   "frmAllOrderPayment.frx":0634
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   9015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALL PREVIOUSLY PRINTED RECEIPT OR INVOICE"
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
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   7455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5775
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmAllOrderPayment"
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
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub cmdPrint_Click()
On Error GoTo abdel
    sql = "select * from payment"
    Set PaymentReport.DataSource = con.Execute(sql)
    PaymentReport.Show
    Set PaymentReport = Nothing
    Exit Sub
abdel:
    MsgBox "Report not available for now...", vbExclamation, title
End Sub

Private Sub cmdView_Click()

    blInvoice = True
    recOrderDetails.Requery
    recOrderDetails.MoveFirst
    Do While Not recOrderDetails.EOF
        If recOrderDetails!orderid = lv.SelectedItem.ListSubItems(1).Text Then
            Set lstItem = frmOrderDetails.lv.ListItems.Add(, , recOrderDetails!orderDetailsID)
                lstItem.ListSubItems.Add , , recOrderDetails!ProductName
                lstItem.ListSubItems.Add , , recOrderDetails!unitprice
                lstItem.ListSubItems.Add , , recOrderDetails!quantity
                lstItem.ListSubItems.Add , , recOrderDetails!discount
                lstItem.ListSubItems.Add , , recOrderDetails!total
                frmOrderDetails.lblNo.Caption = lv.SelectedItem.ListSubItems(1).Text
        End If
        recOrderDetails.MoveNext
    Loop
    Unload Me
    frmOrderDetails.Show
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    recPayment.Requery
    If recPayment.BOF And recPayment.EOF Then
        Label1.Caption = "NO INVOICE OR RECEIPT PRINTED"
        cmdView.Enabled = False
        cmdPrint.Enabled = False
        Exit Sub
    End If
    recPayment.MoveFirst
    Do While Not recPayment.EOF
        Set lstItem = lv.ListItems.Add(, , recPayment!PaymentID)
                lstItem.ListSubItems.Add , , recPayment!orderid & ""
                lstItem.ListSubItems.Add , , recPayment!PaymentMode & ""
                lstItem.ListSubItems.Add , , recPayment!checkNo & ""
                lstItem.ListSubItems.Add , , recPayment!amountpaid & ""
                'lstItem.ListSubItems.Add , , recPayment!amountowed & ""
                lstItem.ListSubItems.Add , , recPayment!dates & ""
                lstItem.ListSubItems.Add , , recPayment!Time & ""
        recPayment.MoveNext
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub

Private Sub lv_DblClick()

    blInvoice = True
    recOrderDetails.Requery
    recOrderDetails.MoveFirst
    Do While Not recOrderDetails.EOF
        If recOrderDetails!orderid = lv.SelectedItem.ListSubItems(1).Text Then
            Set lstItem = frmOrderDetails.lv.ListItems.Add(, , recOrderDetails!orderDetailsID)
                lstItem.ListSubItems.Add , , recOrderDetails!ProductName
                lstItem.ListSubItems.Add , , recOrderDetails!unitprice
                lstItem.ListSubItems.Add , , recOrderDetails!quantity
                lstItem.ListSubItems.Add , , recOrderDetails!discount
                lstItem.ListSubItems.Add , , recOrderDetails!total
                frmOrderDetails.lblNo.Caption = lv.SelectedItem.ListSubItems(1).Text
        End If
        recOrderDetails.MoveNext
    Loop
    Unload Me
End Sub
