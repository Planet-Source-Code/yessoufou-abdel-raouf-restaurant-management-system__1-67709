VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSalesInvoiceReport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lv 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Double click to view sales details"
      Top             =   600
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Invoice ID"
         Object.Width           =   2172
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Amount Due"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amount Paid"
         Object.Width           =   2716
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Changes"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   6510
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
      mIcon           =   "frmSalesInvoiceReport.frx":0000
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   435
      Left            =   3240
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
      mIcon           =   "frmSalesInvoiceReport.frx":031A
   End
   Begin lvButton.lvButtons_H cmdView 
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   5850
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   767
      Caption         =   "&View Sales Details"
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
      mIcon           =   "frmSalesInvoiceReport.frx":0634
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SALE INVOICES RECORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   8775
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5775
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmSalesInvoiceReport"
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
Dim sql As String
Private Sub cmdClose_Click()
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub


Private Sub cmdPrint_Click()
'On Error GoTo abdel
    sql = "select * from salesinvoice"
    Set SaleInvoicesRecord.DataSource = con.Execute(sql)
    SaleInvoicesRecord.Show
    Set SaleInvoicesRecord = Nothing
'    Exit Sub
'abdel:
'    MsgBox "Report not available for now...", vbExclamation, title

End Sub

Private Sub cmdView_Click()
    recSale.Requery
    recSale.MoveFirst
    Do While Not recSale.EOF
        If Trim(recSale!invoiceid) = Trim(lv.SelectedItem.Text) Then
            'Unload Me
            Set lst = frmSalesDetails.lv.ListItems.Add(, , recSale!saleID)
                lst.ListSubItems.Add , , recSale!ProductName
                lst.ListSubItems.Add , , recSale!unitprice
                lst.ListSubItems.Add , , recSale!quantity
                lst.ListSubItems.Add , , recSale!total
                lst.ListSubItems.Add , , recSale!Date
                lst.ListSubItems.Add , , recSale!Time
                
        End If
        recSale.MoveNext
    Loop
    frmSalesDetails.Show
    frmSalesDetails.Label2.Caption = lv.SelectedItem.Text
    Unload Me
End Sub

Private Sub Form_Load()
    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    recSalesInvoice.Requery
    If recSalesInvoice.BOF And recSalesInvoice.EOF Then
        Label1.Caption = "No Sales Invoice Has Been Registered"
        cmdView.Enabled = False
        Exit Sub
    End If
    
    recSalesInvoice.MoveFirst
    Do While Not recSalesInvoice.EOF
        Set lst = lv.ListItems.Add(, , recSalesInvoice!invoiceid)
            lst.ListSubItems.Add , , recSalesInvoice!amountdue
            lst.ListSubItems.Add , , recSalesInvoice!amountpaid
            lst.ListSubItems.Add , , recSalesInvoice!changes
            lst.ListSubItems.Add , , recSalesInvoice!Date & ""
            lst.ListSubItems.Add , , recSalesInvoice!Time & ""
        recSalesInvoice.MoveNext
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub

Private Sub lv_DblClick()
    recSale.Requery
    recSale.MoveFirst
    Do While Not recSale.EOF
        If Trim(recSale!invoiceid) = Trim(lv.SelectedItem.Text) Then
            'Unload Me
            Set lst = SalesDetails.lv.ListItems.Add(, , recSale!saleID)
                lst.ListSubItems.Add , , recSale!ProductName
                lst.ListSubItems.Add , , recSale!unitprice
                lst.ListSubItems.Add , , recSale!quantity
                lst.ListSubItems.Add , , recSale!total
                lst.ListSubItems.Add , , recSale!Date
                lst.ListSubItems.Add , , recSale!Time
                
        End If
        recSale.MoveNext
    Loop
    frmSalesDetails.Show
    frmSalesDetails.Label2.Caption = lv.SelectedItem.Text
    Unload Me
End Sub
