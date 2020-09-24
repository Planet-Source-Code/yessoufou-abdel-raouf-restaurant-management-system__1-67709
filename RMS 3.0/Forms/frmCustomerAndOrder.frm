VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomerAndOrder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "LES CLIENTS ET LES COMMANDES"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox list 
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
      Height          =   3150
      Left            =   240
      TabIndex        =   18
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtFax 
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
      Height          =   345
      Left            =   7680
      TabIndex        =   17
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtPhone 
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
      Height          =   285
      Left            =   7680
      TabIndex        =   16
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtCountry 
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
      Height          =   285
      Left            =   7680
      TabIndex        =   15
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtPostalCode 
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
      Height          =   285
      Left            =   7680
      TabIndex        =   14
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtcity 
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
      Height          =   285
      Left            =   4200
      TabIndex        =   13
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtContactTitle 
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
      Height          =   285
      Left            =   4200
      TabIndex        =   12
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtContactName 
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
      Height          =   285
      Left            =   4200
      TabIndex        =   11
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtCompanyName 
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
      Height          =   285
      Left            =   4200
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtID 
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
      Height          =   285
      Left            =   6120
      TabIndex        =   9
      ToolTipText     =   "Numéro du client"
      Top             =   840
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   2055
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Double click to view order details"
      Top             =   4200
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3625
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Order ID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Order Date"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Delivery Date"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount Due"
         Object.Width           =   4410
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdPrevious 
      Height          =   435
      Left            =   5400
      TabIndex        =   23
      Top             =   3330
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      Caption         =   "&Previous"
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
      mIcon           =   "frmCustomerAndOrder.frx":0000
   End
   Begin lvButton.lvButtons_H cmdFirst 
      Height          =   435
      Left            =   3960
      TabIndex        =   24
      Top             =   3330
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      Caption         =   "&First"
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
      mIcon           =   "frmCustomerAndOrder.frx":031A
   End
   Begin lvButton.lvButtons_H cmdNext 
      Height          =   435
      Left            =   6840
      TabIndex        =   25
      Top             =   3330
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      Caption         =   "&Next"
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
      mIcon           =   "frmCustomerAndOrder.frx":0634
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   435
      Left            =   2520
      TabIndex        =   26
      Top             =   3330
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      Caption         =   "&New"
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
      mIcon           =   "frmCustomerAndOrder.frx":094E
   End
   Begin lvButton.lvButtons_H cmdLast 
      Height          =   435
      Left            =   8310
      TabIndex        =   27
      Top             =   3330
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      Caption         =   "&Last"
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
      mIcon           =   "frmCustomerAndOrder.frx":0C68
   End
   Begin lvButton.lvButtons_H cmdNewOrder 
      Height          =   435
      Left            =   2910
      TabIndex        =   28
      Top             =   6570
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   767
      Caption         =   "&New Order"
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
      mIcon           =   "frmCustomerAndOrder.frx":0F82
   End
   Begin lvButton.lvButtons_H cmdOrderDetails 
      Height          =   435
      Left            =   240
      TabIndex        =   29
      Top             =   6570
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   767
      Caption         =   "&Order Details"
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
      mIcon           =   "frmCustomerAndOrder.frx":129C
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   435
      Left            =   5550
      TabIndex        =   30
      Top             =   6570
      Width           =   1545
      _ExtentX        =   2725
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
      mIcon           =   "frmCustomerAndOrder.frx":15B6
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   8040
      TabIndex        =   22
      Top             =   6570
      Width           =   1545
      _ExtentX        =   2725
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
      mIcon           =   "frmCustomerAndOrder.frx":18D0
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Width           =   9975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT A CUSTOMER TO VIEW HIS ORDERS AND DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   9495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ORDERS"
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
      Left            =   3840
      TabIndex        =   20
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Postal Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   1470
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Shape Shape13 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000006&
      Height          =   6495
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "frmCustomerAndOrder"
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
    Dim position As Integer
    Dim control As Object
    Dim lstItem As ListItem
    Dim sql As String

Private Sub cmdClose_Click()
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub cmdOrderDetails_Click()

    If ListView.ListItems.Count = 0 Then
        MsgBox "Please select a customer and the order which details you want to view", vbCritical, title
        Exit Sub
    End If
    
    blOrderDetails = True
    blOrderDetails1 = True
    recOrderDetails.Requery
    recOrderDetails.MoveFirst
    Do While Not recOrderDetails.EOF
        If recOrderDetails!orderid = ListView.SelectedItem.Text Then
            Set lstItem = frmOrderDetails.lv.ListItems.Add(, , recOrderDetails!orderDetailsID)
                lstItem.ListSubItems.Add , , recOrderDetails!ProductName
                lstItem.ListSubItems.Add , , recOrderDetails!unitprice
                lstItem.ListSubItems.Add , , recOrderDetails!quantity
                lstItem.ListSubItems.Add , , recOrderDetails!discount
                lstItem.ListSubItems.Add , , recOrderDetails!total
                frmOrderDetails.lblNo.Caption = ListView.SelectedItem.Text
                
        End If
        recOrderDetails.MoveNext
    Loop
    Me.Hide
    frmOrderDetails.Show
End Sub

Private Sub cmdNew_Click()

    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see system administrator", vbCritical, title
        Exit Sub
    End If
    If frmScreen.lblRole.Caption = "Administrator" Then
        blAllOrders = True
        blOrd = False
        Unload Me
        Load frmCustomer
        frmCustomer.Show
        'recCustomer.AddNew
        Call frmCustomer.LockMe(False)
        Call frmCustomer.EnableCmd
        frmCustomer.txtID.Text = frmCustomer.autogen
        frmCustomer.txtCompanyName.Text = ""
        frmCustomer.txtContactName.Text = ""
        frmCustomer.txtContactTitle.Text = ""
        frmCustomer.txtAddress.Text = ""
        frmCustomer.txtCity.Text = ""
        frmCustomer.txtPostalCode.Text = ""
        frmCustomer.txtCountry.Text = ""
        frmCustomer.txtPhone.Text = ""
        frmCustomer.txtFax.Text = ""
        Else
            MsgBox "Permission denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub cmdNewOrder_Click()

    blCash = True
    Load frmOrder
    frmOrder.Show
    Unload Me
End Sub

Private Sub cmdLast_Click()

    recCustomer.MoveLast
    Call display
    CallView
End Sub

Private Sub cmdPrevious_Click()
    
    recCustomer.MovePrevious
    If recCustomer.BOF Then
        recCustomer.MoveFirst
    End If
    Call display
    CallView
End Sub

Private Sub cmdFirst_Click()

    recCustomer.MoveFirst
    Call display
    CallView
End Sub

Private Sub cmdNext_Click()

    recCustomer.MoveNext
    If recCustomer.EOF Then
        recCustomer.MoveLast
    End If
    Call display
    CallView
    
End Sub
Private Sub cmddelete_Click()
On Error GoTo abdel
    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see system administrator...", vbCritical, title
        Exit Sub
    End If
    If ListView.ListItems.Count = 0 Then
        MsgBox "Please select a customer and the order you want to delete", vbCritical, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete the order no '" & ListView.SelectedItem.Text & "' ?", vbYesNo + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            sql = "delete from Orders where orderID = " & ListView.SelectedItem.Text
            con.Execute sql
            ListView.ListItems.Remove ListView.SelectedItem.Index
    End If
Exit Sub
abdel:
    MsgBox "Sorry, data could not be deleted", vbExclamation, title
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    Call LockMe(True)
    Call DisableThem
    list.Clear
    recCustomer.Requery
    If Not recCustomer.EOF Then
        Do While Not recCustomer.EOF
            list.AddItem recCustomer!contactname
            recCustomer.MoveNext
        Loop
        recCustomer.MoveFirst
        'Call display
        Else
            MsgBox "There is no customer record in the database", vbCritical, title
            cmdNext.Enabled = False
            cmdPrevious.Enabled = False
            cmdFirst.Enabled = False
            cmdLast.Enabled = False
            cmdOrderDetails.Enabled = False
            cmdDelete.Enabled = False
            Exit Sub
    End If
    
End Sub

Public Sub display()
On Error Resume Next
    
    txtID.Text = recCustomer!customerid & ""
    txtCompanyName.Text = recCustomer!CompanyName & ""
    txtContactName.Text = recCustomer!contactname & ""
    txtContactTitle.Text = recCustomer!ContactTitle & ""
    txtCity.Text = recCustomer!city & ""
    txtPostalCode.Text = recCustomer!PostalCode & ""
    txtCountry.Text = recCustomer!Country & ""
    txtPhone.Text = recCustomer!phone & ""
    txtFax.Text = recCustomer!fax & ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub

Private Sub list_Click()

    position = list.ListIndex
    recCustomer.MoveFirst
    recCustomer.Move position
    Call display
    CallView
End Sub
Public Sub LockMe(LockUnlock As Boolean)

    For Each control In Me
        If TypeOf control Is TextBox Then
            control.Locked = LockUnlock
        End If
    Next

End Sub

Public Function autogen()


    Dim recGen As New Recordset
    
    recGen.Open "select max(Noduclient) from client", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogen = 1
        
        Else
        
        autogen = Val(recGen(0) + 1)
    End If
End Function

Public Sub EnabledThem()

    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    list.Enabled = False
End Sub

Public Sub DisableThem()
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    list.Enabled = True
End Sub


Private Sub ListView_DblClick()
    If ListView.ListItems.Count = 0 Then
        MsgBox "Please select a customer and the order which details you want to view", vbCritical, title
        Exit Sub
    End If
    
    blOrderDetails = True
    recOrderDetails.Requery
    recOrderDetails.MoveFirst
    Do While Not recOrderDetails.EOF
        If recOrderDetails!orderid = ListView.SelectedItem.Text Then
            Set lstItem = frmOrderDetails.lv.ListItems.Add(, , recOrderDetails!orderDetailsID)
                lstItem.ListSubItems.Add , , recOrderDetails!ProductName
                lstItem.ListSubItems.Add , , recOrderDetails!unitprice
                lstItem.ListSubItems.Add , , recOrderDetails!quantity
                lstItem.ListSubItems.Add , , recOrderDetails!discount
                lstItem.ListSubItems.Add , , recOrderDetails!total
                frmOrderDetails.lblNo.Caption = ListView.SelectedItem.Text
        End If
        recOrderDetails.MoveNext
    Loop
    Me.Hide
    frmScreen.Enabled = True
    frmOrderDetails.Show
End Sub

Private Sub txtCompanyname_Validate(Cancel As Boolean)
    txtCompanyName.Text = StrConv(txtCompanyName.Text, vbProperCase)
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack, vbKeySpace
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCountactName_Validate(Cancel As Boolean)
    txtContactName.Text = StrConv(txtContactName.Text, vbProperCase)
End Sub

Private Sub txtCountry_Validate(Cancel As Boolean)
    txtCountry.Text = StrConv(txtCountry.Text, vbProperCase)
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack, vbKeySpace
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtContactTitle_Validate(Cancel As Boolean)
    txtContactTitle.Text = StrConv(txtContactTitle.Text, vbProperCase)
End Sub

Private Sub txtcity_Validate(Cancel As Boolean)
    txtCity.Text = StrConv(txtCity.Text, vbProperCase)
End Sub

Public Sub CallView()
    ListView.ListItems.Clear
    recOrder.Requery
    If Not recOrder.EOF Then
        recOrder.MoveFirst
        Do While Not recOrder.EOF
            If Trim(recOrder!customerid) = Trim(txtID.Text) Then
                Set lstItem = ListView.ListItems.Add(, , recOrder!orderid)
                    lstItem.ListSubItems.Add , , recOrder!Orderdate & ""
                    lstItem.ListSubItems.Add , , recOrder!deliverydate & ""
                    lstItem.ListSubItems.Add , , recOrder!amountdue & ""
            End If
            recOrder.MoveNext
        Loop
        Else
    End If
End Sub
