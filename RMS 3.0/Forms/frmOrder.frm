VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtName 
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   480
      Width           =   2055
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2295
      Left            =   240
      TabIndex        =   18
      ToolTipText     =   "Double click to edit a record"
      Top             =   3000
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4048
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
      MouseIcon       =   "frmOrder.frx":0000
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Name"
         Object.Width           =   4234
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Unit Price"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Discount%"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Discount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker picker2 
      Height          =   375
      Left            =   6480
      TabIndex        =   17
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   46530561
      CurrentDate     =   38667
   End
   Begin MSComCtl2.DTPicker picker1 
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   2400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   46530561
      CurrentDate     =   38667
   End
   Begin VB.TextBox txtAmountDue 
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
      Left            =   6480
      TabIndex        =   10
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtDiscount 
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
      Left            =   6480
      TabIndex        =   9
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtTotal 
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
      Left            =   6480
      TabIndex        =   8
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtDisc 
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
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   7
      ToolTipText     =   "Discount"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtQuantity 
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
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   6
      ToolTipText     =   "Quantity Or Number Of Product Bought"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtUnitPrice 
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
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   1935
   End
   Begin VB.ComboBox com 
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
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   21
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtOrderID 
      Height          =   405
      Left            =   2520
      TabIndex        =   25
      Top             =   3360
      Width           =   2175
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   435
      Left            =   3840
      TabIndex        =   29
      Top             =   6210
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Caption         =   "&Delete"
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
      mIcon           =   "frmOrder.frx":031A
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   435
      Left            =   2040
      TabIndex        =   30
      Top             =   6210
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Caption         =   "&Add"
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
      mIcon           =   "frmOrder.frx":0634
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   435
      Left            =   5670
      TabIndex        =   31
      Top             =   6210
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Caption         =   "&Invoice"
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
      mIcon           =   "frmOrder.frx":094E
   End
   Begin lvButton.lvButtons_H cmdCustomer 
      Height          =   435
      Left            =   240
      TabIndex        =   27
      Top             =   6210
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Caption         =   "&Customer"
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
      mIcon           =   "frmOrder.frx":0C68
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   7530
      TabIndex        =   28
      Top             =   6210
      Width           =   1335
      _ExtentX        =   2355
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
      mIcon           =   "frmOrder.frx":0F82
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "(%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
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
      Left            =   4800
      TabIndex        =   23
      Top             =   480
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   9135
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   4680
      TabIndex        =   22
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000009&
      Caption         =   "Amount Due"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Label total 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   19
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   8655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
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
      Left            =   4800
      TabIndex        =   16
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Required Date"
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
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   4800
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Left            =   4800
      TabIndex        =   12
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   4800
      TabIndex        =   11
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
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
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
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
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   6135
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmOrder"
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
Private Sub cmdCustomer_Click()
    blOrder = True
    Me.Hide
    Load frmCustomerOrder
    frmCustomerOrder.Show
End Sub

Private Sub cmddelete_Click()

    If MsgBox("Are you sure you want to delete details of '" & lv.SelectedItem.Text & "' ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            total.Caption = Val(total.Caption) - Val(lv.SelectedItem.SubItems(5))
            lv.ListItems.Remove lv.SelectedItem.Index
            If lv.ListItems.Count = 0 Then
                cmdDelete.Enabled = False
                cmdPrint.Enabled = False
            End If
    End If
    
End Sub
Private Sub cmdPrint_Click()

    If txtName.Text = "" Then
        MsgBox "Please select the customer that is ordering.", vbExclamation, title
        txtName.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Are you sure want to register all details to the database ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
                recOrder.AddNew
                recOrder(0) = autogen
                recOrder(1) = txtID.Text
                recOrder(2) = Date
                recOrder(3) = Picker1.Value
                recOrder(4) = Picker2.Value
                recOrder(5) = total.Caption
                recOrder.Update
            Dim ctr As Integer
            For ctr = 1 To lv.ListItems.Count
                recOrderDetails.AddNew
                recOrderDetails(0) = autogeneration
                recOrderDetails(1) = recOrder(0)
                recOrderDetails(2) = lv.ListItems(ctr).Text
                recOrderDetails(3) = lv.ListItems(ctr).ListSubItems(1).Text
                recOrderDetails(4) = lv.ListItems(ctr).ListSubItems(2).Text
                recOrderDetails(5) = lv.ListItems(ctr).ListSubItems(4).Text
                recOrderDetails(6) = lv.ListItems(ctr).ListSubItems(5).Text
                recOrderDetails.Update
            Next
    End If
    
    If blCredit = True And blCash = False Then
        recCredit.AddNew
        recCredit(0) = autogeneCredit
        recCredit(1) = recOrder(0)
        recCredit(2) = txtName.Text
        recCredit(3) = total.Caption
        recCredit(4) = 0
        recCredit(5) = total.Caption
        recCredit.Update
        CreditOrder.Sections("section2").Controls.Item("lblOrderNo").Caption = txtOrderID.Text
        CreditOrder.Sections("section5").Controls.Item("lblAmountDue").Caption = total.Caption
        sql = "select * from orderdetails where OrderID = " & Trim(txtOrderID.Text)
        Set CreditOrder.DataSource = con.Execute(sql)
        CreditOrder.Show
        Set CreditOrder = Nothing
        Unload Me
        Exit Sub
    End If
    
    If blCash = True And blCredit = False Then
        recOrderDetails.Requery
        recOrderDetails.MoveFirst
        Do While Not recOrderDetails.EOF
            If Trim(recOrderDetails!orderid) = Trim(frmOrderInvoice.txtOrderID.Text) Then
                frmOrderInvoice.List1.AddItem recOrderDetails!ProductName
                frmOrderInvoice.List2.AddItem recOrderDetails!unitprice
                frmOrderInvoice.List3.AddItem recOrderDetails!quantity
                frmOrderInvoice.List4.AddItem recOrderDetails!discount
                frmOrderInvoice.List5.AddItem recOrderDetails!total
                
            End If
            recOrderDetails.MoveNext
        Loop
    
        frmOrderInvoice.lblName.Caption = txtName.Text
        Unload Me
    End If
    
    Exit Sub
abdel:
    MsgBox "Invoice not available for now...", vbExclamation, title
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blCash = False
    blCredit = False
    frmScreen.Enabled = True
End Sub

Private Sub lv_DblClick()
    com.Text = lv.SelectedItem.Text
    txtUnitPrice.Text = lv.SelectedItem.SubItems(1)
    txtQuantity.Text = lv.SelectedItem.SubItems(2)
    txtDisc.Text = lv.SelectedItem.SubItems(3)
    txtDiscount.Text = lv.SelectedItem.SubItems(4)
    total = Val(total.Caption) - Val(lv.SelectedItem.SubItems(5))
    lv.ListItems.Remove lv.SelectedItem.Index
    If lv.ListItems.Count = 0 Then
        cmdPrint.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub txtDisc_Change()

        If txtDisc.Text = "" Then
            txtDiscount.Text = 0
            txtAmountDue.Text = Val(txtTotal.Text)
            Else
                txtDiscount.Text = Val(Val(txtTotal.Text) * Val(txtDisc.Text)) / 100
                txtAmountDue.Text = Val(txtTotal.Text) - Val(txtDiscount.Text)
        End If
End Sub

Private Sub txtQuantity_GotFocus()

    txtQuantity.SelStart = 0
    txtQuantity.SelLength = Len(txtQuantity.Text)
    txtQuantity.SetFocus
    txtDisc.Text = ""
    txtDiscount.Text = ""
    txtAmountDue.Text = ""
End Sub

Private Sub txtunitprice_Change()
    txtTotal.Text = Val(txtUnitPrice.Text)
End Sub

Private Sub txtQuantity_Change()

    If txtQuantity.Text = "" Then
        txtTotal.Text = Val(txtUnitPrice.Text)
        Else
            txtTotal.Text = Val(txtUnitPrice.Text) * Val(txtQuantity.Text)
            txtAmountDue.Text = Val(txtTotal.Text)
    End If
End Sub

Private Sub cmdAdd_Click()

   If com.Text = "" Then
        MsgBox "Select the product or item name", vbExclamation, title
        com.SetFocus
        Exit Sub
    End If
    
    If txtQuantity.Text = "" Then
        MsgBox "What is the quantity or number of product bought", vbExclamation, title
        txtQuantity.SetFocus
        Exit Sub
    End If
    
    If Val(txtDisc.Text) > 100 Then
        MsgBox "The discount percentage can not be more than 100.", vbExclamation, title
        txtDisc.SelStart = 0
        txtDisc.SelLength = Len(txtRabais.Text)
        txtDisc.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to add the details to the list?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
                    
            Set lst = lv.ListItems.Add(, , com)
            lst.ListSubItems.Add , , txtUnitPrice.Text
            lst.ListSubItems.Add , , txtQuantity.Text
            If txtDisc.Text = "" Then
                lst.ListSubItems.Add , , 0
                lst.ListSubItems.Add , , 0
                Else
                lst.ListSubItems.Add , , txtDisc.Text
                lst.ListSubItems.Add , , txtDiscount.Text
            End If

            lst.ListSubItems.Add , , txtAmountDue.Text
            total.Caption = Val(total.Caption) + Val(txtAmountDue.Text)
            cmdDelete.Enabled = True
            cmdPrint.Enabled = True
            com.Text = ""
            txtUnitPrice.Text = ""
            txtQuantity.Text = ""
            txtDisc.Text = ""
            txtDiscount.Text = ""
            txtAmountDue.Text = ""
    End If
End Sub

Private Sub cmdClose_Click()
   If cmdDelete.Enabled = True Or cmdPrint.Enabled = True Then
        MsgBox "Please, delete or save the data before you close", vbExclamation, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub com_Click()

    position = com.ListIndex
    recProduct.MoveFirst
    recProduct.Move position
    txtUnitPrice.Text = recProduct!unitprice & ""
    txtQuantity.Text = ""
    txtDisc.Text = ""
    txtDiscount.Text = ""
    txtAmountDue.Text = ""
End Sub

Private Sub com_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False

    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    
    Picker1.Value = Date
    Picker2.Value = Date
    txtOrderID.Text = autogen
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    'lblName.Caption = "Veuillez Sélectionner Le Client svp"
    com.Clear
    If recProduct.BOF And recProduct.EOF Then
        MsgBox "There is no product or item in the database.", vbInformation, title
        cmdAdd.Enabled = False
        Exit Sub
    End If
    recProduct.Requery
    recProduct.MoveFirst
    Do While Not recProduct.EOF
        com.AddItem recProduct!ProductName
        recProduct.MoveNext
    Loop
    recProduct.MoveFirst

End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub
Public Function autogeneration()
'On Error Resume Next

    Dim recGen As New Recordset
    
    recGen.Open "select max(OrderDetailsID) from OrderDetails", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogeneration = 1
        
        Else
        
        autogeneration = Val(recGen(0) + 1)
    End If
End Function
Public Function autogen()


    Dim recGen As New Recordset
    
    recGen.Open "select max(OrderID) from orders", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogen = 1
        
        Else
        
        autogen = Val(recGen(0) + 1)
    End If
End Function

Public Function autogeneCredit()

    Dim recGen As New Recordset
    
    recGen.Open "select max(CreditID) from credit", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogeneCredit = 1
        
        Else
        
        autogeneCredit = Val(recGen(0) + 1)
    End If
End Function
