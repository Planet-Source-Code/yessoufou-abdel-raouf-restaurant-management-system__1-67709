VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPurchase 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lv 
      Height          =   2175
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Double click to modify any record"
      Top             =   2520
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3836
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Unit Price"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Supplier"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtProduct 
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
      TabIndex        =   0
      ToolTipText     =   "Product Purchased"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtSupplier 
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
      TabIndex        =   3
      ToolTipText     =   "The Supplier"
      Top             =   1920
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
      TabIndex        =   1
      ToolTipText     =   "Unit Price"
      Top             =   960
      Width           =   2055
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
      TabIndex        =   2
      ToolTipText     =   "Quantity Or Number Purchased"
      Top             =   1440
      Width           =   2055
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
      TabIndex        =   4
      Top             =   480
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
      TabIndex        =   5
      Top             =   960
      Width           =   1815
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
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
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
      Height          =   300
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   0
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   435
      Left            =   2040
      TabIndex        =   18
      Top             =   5610
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
      mIcon           =   "frmPurchase.frx":0000
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   435
      Left            =   240
      TabIndex        =   19
      Top             =   5610
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
      mIcon           =   "frmPurchase.frx":031A
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   5730
      TabIndex        =   20
      Top             =   5610
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
      mIcon           =   "frmPurchase.frx":0634
   End
   Begin lvButton.lvButtons_H cmdFinish 
      Height          =   435
      Left            =   3870
      TabIndex        =   21
      Top             =   5610
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
      mIcon           =   "frmPurchase.frx":094E
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE"
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
      Left            =   2160
      TabIndex        =   17
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
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
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   7335
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
      TabIndex        =   13
      Top             =   480
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
      TabIndex        =   12
      Top             =   960
      Width           =   1215
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
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
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
      TabIndex        =   10
      Top             =   480
      Width           =   1095
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
      TabIndex        =   9
      Top             =   960
      Width           =   855
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
      Top             =   1440
      Width           =   735
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
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   495
      Left            =   240
      Top             =   4680
      Width           =   6855
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5535
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmPurchase"
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
    Dim reply As Integer
    Dim control As Object
    Dim lst As ListItem
    Dim RS As New Recordset
Private Sub cmdAdd_Click()

    If txtProduct.Text = "" Then
        MsgBox "Kindly type in the product purchased.", vbExclamation, title
        txtProduct.SetFocus
        Exit Sub
    End If
    If txtUnitPrice.Text = "" Then
        MsgBox "What is the unit price of '" & Trim(txtProduct.Text) & "' ?", vbQuestion, title
        txtUnitPrice.SetFocus
        Exit Sub
    End If
    If txtQuantity.Text = "" Then
        MsgBox "What is the quantity or number bought.", vbCritical, title
        txtQuantity.SetFocus
        Exit Sub
    End If
    If MsgBox("Are you sure you want to add the details to the list ?", 4 + vbQuestion, title) = vbNo Then
            Exit Sub
            Else
                Set lst = lv.ListItems.Add(, , txtProduct.Text)
                    lst.ListSubItems.Add , , txtUnitPrice.Text
                    lst.ListSubItems.Add , , txtQuantity.Text
                    lst.ListSubItems.Add , , txtTotal.Text
                    lst.ListSubItems.Add , , txtSupplier.Text
                    lst.ListSubItems.Add , , txtDate.Text
                    lst.ListSubItems.Add , , txtTime.Text
                    
                    total.Text = Val(total.Text) + Val(txtTotal.Text)
                    txtProduct.Text = ""
                    txtUnitPrice.Text = ""
                    txtQuantity.Text = ""
                    txtTotal.Text = ""
                    txtSupplier.Text = ""
                    txtProduct.SetFocus
                    cmdDelete.Enabled = True
                    cmdFinish.Enabled = True
    End If
End Sub

Private Sub cmdClose_Click()
   If cmdDelete.Enabled = True Or cmdFinish.Enabled = True Then
        MsgBox "Please, register or clear the list to be able to quit", vbExclamation, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub cmddelete_Click()

    If MsgBox("Are you want to delete the details of '" & lv.SelectedItem.Text & "' ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            total = Val(total.Text) - Val(lv.SelectedItem.SubItems(3))
            lv.ListItems.Remove lv.SelectedItem.Index
            
            If lv.ListItems.Count = 0 Then
                cmdFinish.Enabled = False
                cmdDelete.Enabled = False
            End If
    End If
End Sub

Private Sub cmdFinish_Click()
On Error GoTo abdel
    If MsgBox("Are you sure you want to register the details to the database ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            Dim ctr As Integer
            RS.Open "Select * from purchase", con, adOpenDynamic, adLockOptimistic
            For ctr = 1 To lv.ListItems.Count
                RS.AddNew
                RS(0) = autogen
                RS(1) = lv.ListItems(ctr).Text
                RS(2) = lv.ListItems(ctr).ListSubItems(1).Text
                RS(3) = lv.ListItems(ctr).ListSubItems(2).Text
                RS(4) = lv.ListItems(ctr).ListSubItems(3).Text
                RS(5) = lv.ListItems(ctr).ListSubItems(4).Text
                RS(6) = lv.ListItems(ctr).ListSubItems(6).Text
                RS(7) = lv.ListItems(ctr).ListSubItems(5).Text
                RS.Update
                
            Next
    lv.ListItems.Clear
    total.Text = ""
    cmdFinish.Enabled = False
    cmdDelete.Enabled = False
    End If
Unload Me
    Exit Sub
abdel:
    MsgBox "Sorry, transactions no successfully registered", vbExclamation, title

End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    cmdDelete.Enabled = False
    cmdFinish.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

    RS.Close
    frmScreen.Enabled = True
End Sub

Private Sub lv_DblClick()

    txtProduct.Text = lv.SelectedItem.Text
    txtUnitPrice.Text = lv.SelectedItem.SubItems(1)
    txtQuantity.Text = lv.SelectedItem.SubItems(2)
    txtTotal.Text = lv.SelectedItem.SubItems(3)
    txtSupplier.Text = lv.SelectedItem.SubItems(4)
    total = Val(total.Text) - Val(lv.SelectedItem.SubItems(3))
    lv.ListItems.Remove lv.SelectedItem.Index
    If lv.ListItems.Count = 0 Then
        cmdFinish.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()

    txtDate.Text = Date
    txtTime.Text = Time
End Sub

Private Sub txtProduct_GotFocus()
    txtProduct.SelStart = 0
    txtProduct.SelLength = Len(txtProduct.Text)
    txtProduct.SetFocus
    
    txtUnitPrice.Text = ""
    txtQuantity.Text = ""
    txtSupplier.Text = ""
End Sub

Private Sub txtunitprice_Change()

    txtTotal.Text = Val(txtUnitPrice.Text)
End Sub

Private Sub txtUnitPrice_GotFocus()
    txtUnitPrice.SelStart = 0
    txtUnitPrice.SelLength = Len(txtUnitPrice.Text)
    txtUnitPrice.SetFocus
    
    txtQuantity.Text = ""
    txtSupplier.Text = ""
End Sub

Private Sub txtunitprice_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack, vbKeyDelete
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtProduct_Validate(Cancel As Boolean)
    txtProduct.Text = StrConv(txtProduct.Text, vbProperCase)
End Sub

Private Sub txtQuantity_Change()
    If txtQuantity.Text = "" Then
        txtTotal.Text = Val(txtUnitPrice.Text)
        Else
            txtTotal.Text = Val(txtUnitPrice.Text) * Val(txtQuantity.Text)
    End If
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack, vbKeyDelete
        Case Else
            KeyAscii = 0
    End Select
End Sub
Public Function autogen()

    Dim recGen As New Recordset
    
    recGen.Open "select max(purchaseid) from purchase", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogen = 1
        
        Else
        
        autogen = Val(recGen(0) + 1)
    End If
End Function

