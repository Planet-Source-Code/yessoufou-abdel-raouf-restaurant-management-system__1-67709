VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSale 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   240
   End
   Begin VB.TextBox total 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5040
      Width           =   2895
   End
   Begin MSComctlLib.ListView list 
      Height          =   2295
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Double click to edit any record"
      Top             =   2640
      Width           =   6855
      _ExtentX        =   12091
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Unit Price"
         Object.Width           =   2119
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Total"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date"
         Object.Width           =   2472
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Time"
         Object.Width           =   2472
      EndProperty
   End
   Begin VB.TextBox txtDate 
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtTime 
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtTotal 
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   840
      Width           =   1815
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
      Height          =   360
      Left            =   1440
      TabIndex        =   5
      ToolTipText     =   "Quantity or Number Sold"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtUnitPrice 
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
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
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Select The Product"
      Top             =   840
      Width           =   2055
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   435
      Left            =   2070
      TabIndex        =   19
      Top             =   5850
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
      mIcon           =   "frmSale.frx":0000
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   435
      Left            =   210
      TabIndex        =   20
      Top             =   5850
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
      mIcon           =   "frmSale.frx":031A
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   5730
      TabIndex        =   17
      Top             =   5850
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
      mIcon           =   "frmSale.frx":0634
   End
   Begin lvButton.lvButtons_H cmdFinish 
      Height          =   435
      Left            =   3900
      TabIndex        =   18
      Top             =   5850
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Caption         =   "&Finish"
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
      mIcon           =   "frmSale.frx":094E
   End
   Begin VB.Label lbl 
      Caption         =   "Label9"
      Height          =   615
      Left            =   2760
      TabIndex        =   16
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SALES"
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
      Height          =   495
      Left            =   2040
      TabIndex        =   15
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   615
      Left            =   240
      Top             =   4920
      Width           =   6855
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label4 
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
      Left            =   3840
      TabIndex        =   6
      Top             =   840
      Width           =   1095
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
      Top             =   2040
      Width           =   1095
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
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Products"
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
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   7335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5775
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmSale"
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
    Dim RS As New Recordset
    Dim position As Integer
    Dim prompt As String
Private Sub cmdAdd_Click()
    If txtUnitPrice.Text = "" Then
        MsgBox "Kindly select a product that you are selling.", vbExclamation, title
        com.SetFocus
        Exit Sub
    End If
    
    If txtQuantity.Text = "" Then
        MsgBox "What is the quantity or number of '" & Trim(com.Text) & " ' Sold ?", vbExclamation, title
        txtQuantity.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to add these details to the list ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            Set lst = list.ListItems.Add(, , com.Text)
                lst.ListSubItems.Add , , txtUnitPrice.Text
                lst.ListSubItems.Add , , txtQuantity.Text
                lst.ListSubItems.Add , , txtTotal.Text
                lst.ListSubItems.Add , , txtTime.Text
                lst.ListSubItems.Add , , txtDate.Text

                total.Text = Val(total.Text) + Val(txtTotal.Text)
                cmdDelete.Enabled = True
                cmdFinish.Enabled = True
                    
                com.Text = ""
                txtUnitPrice.Text = ""
                txtQuantity.Text = ""
                txtTotal.Text = ""
    End If
End Sub

Private Sub cmddelete_Click()
    If MsgBox("Are you sure you want to delete the details of '" & list.SelectedItem.Text & "' ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            total.Text = Val(total.Text) - Val(list.SelectedItem.SubItems(3))
            list.ListItems.Remove list.SelectedItem.Index
            If list.ListItems.Count = 0 Then
                cmdDelete.Enabled = False
                cmdFinish.Enabled = False
            End If
    End If
End Sub

Private Sub cmdFinish_Click()
'On Error GoTo abdel
    If MsgBox("Are you sure you want to register all the details in the list to the database ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            recSalesInvoice.AddNew
            recSalesInvoice!invoiceid = lbl.Caption
            recSalesInvoice!amountdue = total.Text
            recSalesInvoice.Update
            Dim ctr As Integer
            RS.Open "Select * from sale", con, adOpenDynamic, adLockOptimistic
            For ctr = 1 To list.ListItems.Count
                RS.AddNew
                RS(0) = autogen
                RS(1) = lbl.Caption
                RS(2) = list.ListItems(ctr).Text
                RS(3) = list.ListItems(ctr).ListSubItems(1).Text
                RS(4) = list.ListItems(ctr).ListSubItems(2).Text
                RS(5) = list.ListItems(ctr).ListSubItems(3).Text
                RS(7) = list.ListItems(ctr).ListSubItems(4).Text
                RS(6) = list.ListItems(ctr).ListSubItems(5).Text
                RS.Update
                
            Next
            
        'Load SalesInvoice
        frmSalesInvoice.Show
        frmSalesInvoice.lblInvoiceId.Caption = lbl.Caption
        frmSalesInvoice.txtAmountDue.Text = total.Text
        recSale.Requery
        recSale.MoveFirst
        Do While Not recSale.EOF
            If Trim(recSale!invoiceid) = Trim(frmSalesInvoice.lblInvoiceId.Caption) Then
                frmSalesInvoice.List1.AddItem recSale!ProductName
                frmSalesInvoice.List2.AddItem recSale!unitprice
                frmSalesInvoice.List3.AddItem recSale!quantity
                frmSalesInvoice.List4.AddItem recSale!total
            
            End If
            recSale.MoveNext
         Loop
    End If
Unload Me
'    Exit Sub
'abdel:
'    MsgBox "Invoice not available for now", vbExclamation, title
End Sub

Private Sub com_Click()
    position = com.ListIndex
    recProduct.MoveFirst
    recProduct.Move position
    txtUnitPrice.Text = recProduct!unitprice
    txtQuantity.Text = ""
End Sub

Private Sub com_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    lbl.Caption = autogeneration
    cmdDelete.Enabled = False
    cmdFinish.Enabled = False
    recProduct.Requery
    If Not recProduct.EOF Then
        recProduct.MoveFirst
        Do While Not recProduct.EOF
            com.AddItem recProduct!ProductName
            recProduct.MoveNext
        Loop
        recProduct.MoveFirst
        Else
            MsgBox "There is no product in the database yet.", vbExclamation, title
            cmdDelete.Enabled = False
            cmdFinish.Enabled = False
    End If
End Sub

Private Sub cmdClose_Click()
   If cmdDelete.Enabled = True Or cmdFinish.Enabled = True Then
        MsgBox "Please, finish or delete the details", vbExclamation, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to close this window?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack, vbKeyDelete, vbKey0 To vbKey9
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RS.Close
    frmScreen.Enabled = True
End Sub

Private Sub list_DblClick()
    com.Text = list.SelectedItem.Text
    txtUnitPrice.Text = list.SelectedItem.SubItems(1)
    txtQuantity.Text = list.SelectedItem.SubItems(2)
    txtTotal.Text = list.SelectedItem.SubItems(3)
    total = Val(total.Text) - Val(list.SelectedItem.SubItems(3))
    list.ListItems.Remove list.SelectedItem.Index
    If list.ListItems.Count = 0 Then
        cmdFinish.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    txtTime.Text = Time
    txtDate.Text = Date
End Sub

Private Sub txtQuantity_Change()
    If txtQuantity.Text = "" Then
        txtTotal.Text = Val(txtUnitPrice.Text)
        Else
            txtTotal.Text = Val(txtUnitPrice.Text) * Val(txtQuantity.Text)
    End If
End Sub

Private Sub txtunitprice_Change()
    txtTotal.Text = Val(txtUnitPrice.Text)
End Sub
Public Function autogen()

    Dim recGen As New Recordset
    
    recGen.Open "select max(saleID) from sale", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogen = 1
        
        Else
        
        autogen = Val(recGen(0) + 1)
    End If
End Function

Public Function autogeneration()

    Dim recGen As New Recordset
    
    recGen.Open "select max(InvoiceID) from salesInvoice", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogeneration = 1
        
        Else
        
        autogeneration = Val(recGen(0) + 1)
    End If
End Function
