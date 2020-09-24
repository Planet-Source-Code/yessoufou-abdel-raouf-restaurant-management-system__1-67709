VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmProduct 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
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
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   2295
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
      Height          =   375
      Left            =   4800
      MaxLength       =   10
      TabIndex        =   10
      ToolTipText     =   "Unit Price"
      Top             =   4080
      Width           =   2655
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
      Left            =   4800
      MaxLength       =   20
      TabIndex        =   8
      ToolTipText     =   "Quantity"
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txtCategoryID 
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
      Left            =   4800
      TabIndex        =   6
      Top             =   2520
      Width           =   2655
   End
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
      Left            =   4800
      MaxLength       =   20
      TabIndex        =   4
      ToolTipText     =   "Product Name"
      Top             =   1800
      Width           =   2655
   End
   Begin VB.ListBox list 
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
      Height          =   3630
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtProductID 
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
      Left            =   4800
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   435
      Left            =   2670
      TabIndex        =   13
      Top             =   5370
      Width           =   1005
      _ExtentX        =   1773
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
      mIcon           =   "frmProduct.frx":0000
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   435
      Left            =   1380
      TabIndex        =   14
      Top             =   5370
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      Caption         =   "&Save"
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
      mIcon           =   "frmProduct.frx":031A
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   435
      Left            =   3990
      TabIndex        =   15
      Top             =   5370
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      Caption         =   "&Update"
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
      mIcon           =   "frmProduct.frx":0634
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   435
      Left            =   120
      TabIndex        =   16
      Top             =   5370
      Width           =   1005
      _ExtentX        =   1773
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
      mIcon           =   "frmProduct.frx":094E
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   435
      Left            =   5250
      TabIndex        =   17
      Top             =   5370
      Width           =   1005
      _ExtentX        =   1773
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
      mIcon           =   "frmProduct.frx":0C68
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   6450
      TabIndex        =   18
      Top             =   5370
      Width           =   1005
      _ExtentX        =   1773
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
      mIcon           =   "frmProduct.frx":0F82
   End
   Begin lvButton.lvButtons_H cmdCategory 
      Height          =   435
      Left            =   2430
      TabIndex        =   19
      Top             =   660
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   767
      Caption         =   "..."
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
      mIcon           =   "frmProduct.frx":129C
   End
   Begin lvButton.lvButtons_H cmdOk 
      Height          =   435
      Left            =   3990
      TabIndex        =   20
      Top             =   5370
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      Caption         =   "&Ok"
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
      mIcon           =   "frmProduct.frx":15B6
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTS"
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
      Left            =   2280
      TabIndex        =   12
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   7575
   End
   Begin VB.Label Label5 
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
      Left            =   3000
      TabIndex        =   9
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label4 
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
      Left            =   3000
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Category ID"
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
      Left            =   3000
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5295
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmProduct"
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
Dim position As Integer
Dim sql As String
Private Sub CmdCancel_Click()

    If MsgBox("Are you sure you want to cancel ?", 4, title) = vbNo Then
        Exit Sub
        Else
            Call LockMe(True)
            cmdSave.Enabled = False
            cmdCancel.Enabled = False
            cmdNew.Enabled = True
            cmdUpdate.Enabled = True
            cmdDelete.Enabled = True
            list.Enabled = True
            list.Clear
            recProduct2.Requery
            If Not recProduct2.EOF Then
                recProduct2.MoveFirst
                Do While Not recProduct2.EOF
                    list.AddItem recProduct2!ProductName
                    recProduct2.MoveNext
                Loop
                recProduct2.MoveFirst
                Call display
                cmdUpdate.Visible = True
                cmdOK.Visible = False
                Else
                    txtProductID.Text = ""
                    txtName.Text = ""
                    txtQuantity.Text = ""
                    txtUnitPrice.Text = ""
                    cmdUpdate.Visible = True
                    cmdOK.Visible = False
            End If
    End If
End Sub

Private Sub cmdCategory_Click()

    blCategory = True
    Unload Me
    Load frmCategory
    frmCategory.Show
End Sub

Private Sub cmdClose_Click()
    If cmdSave.Enabled = True Or cmdCancel.Enabled = True Then
        MsgBox "Please, cancel or save the new product records", vbExclamation, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            If blProduct = True Then
                Load frmCategory
                frmCategory.Show
            End If
    End If
    Unload Me
End Sub

Private Sub combo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmddelete_Click()
On Error GoTo abdel
    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see system Administrator...", vbCritical, title
        Exit Sub
    End If
    If com.Text = "" Then
        MsgBox "Kindly select a category in the combo box", vbExclamation, title
        com.SetFocus
        Exit Sub
    End If
    
    If txtProductID.Text = "" Then
        MsgBox "Kindly select the product to be deleted ", vbExclamation, title
        Exit Sub
    End If

    If MsgBox("Are you sure you want to delete the product '" & Trim(txtName.Text) & "' ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
        recProduct2.Delete adAffectCurrent
        list.Clear
        recProduct2.Requery
        If Not recProduct2.EOF Then
            recProduct2.MoveFirst
            Do While Not recProduct2.EOF
                list.AddItem recProduct2!nom
                recProduct2.MoveNext
            Loop
            recProduct2.MoveFirst
            Else
        End If
    End If
    If list.Text = "" Then
        txtProductID.Text = ""
        txtName.Text = ""
        txtQuantity.Text = ""
        txtUnitPrice.Text = ""
        Else
            Call display
    End If
    Exit Sub
abdel:
    MsgBox "Data can not be deleted", vbExclamation, title

End Sub

Private Sub cmdNew_Click()

    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see system administrator...", vbCritical, title
        Exit Sub
    End If
    If com.Text = "" Then
        MsgBox "Kindly select a category in the combo box", vbExclamation, title
        com.SetFocus
        Exit Sub
    End If

    If MsgBox("Are you sure you want to add new record ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Call LockMe(False)
            cmdSave.Enabled = True
            cmdCancel.Enabled = True
            cmdNew.Enabled = False
            cmdUpdate.Enabled = False
            cmdDelete.Enabled = False
            list.Enabled = False
            'recProduct2.AddNew
            txtProductID.Text = autogen
            txtName.Text = ""
            txtQuantity.Text = ""
            txtUnitPrice.Text = ""
    End If
End Sub

Private Sub cmdOk_Click()

    If txtName.Text = "" Then
        MsgBox "The product name field can not be left blank.", vbExclamation, title
        txtName.SetFocus
        Exit Sub
    End If
    
    If txtUnitPrice.Text = "" Then
        MsgBox "The unit price field can not be left blank.", vbExclamation, title
        txtUnitPrice.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Are you sure want to save the updated record ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            
            'recProduct2!productID = txtProductID.Text & ""
            recProduct2!ProductName = txtName.Text & ""
            recProduct2!categoryid = txtCategoryID.Text & ""
            recProduct2!quantity = txtQuantity.Text & ""
            recProduct2!unitprice = Val(txtUnitPrice.Text) & ""
            recProduct2.UpdateBatch adAffectCurrent
            MsgBox "New product data successfully saved.", vbInformation, title
            list.Clear
            recProduct2.Requery
            If Not recProduct2.EOF Then
                recProduct2.MoveFirst
                Do While Not recProduct2.EOF
                    list.AddItem recProduct2!ProductName
                    recProduct2.MoveNext
                Loop
                recProduct2.MoveFirst
                Call display
                Call LockMe(True)
                cmdSave.Enabled = False
                cmdCancel.Enabled = False
                cmdNew.Enabled = True
                cmdUpdate.Enabled = True
                cmdDelete.Enabled = True
                list.Enabled = True
            End If
        End If
    Exit Sub
abdel:
    MsgBox "sorry, transactions not successfully saved", vbExclamation, title

End Sub

Private Sub cmdSave_Click()
On Error GoTo abdel
    If txtName.Text = "" Then
        MsgBox "The product name field can not be left blank.", vbExclamation, title
        txtName.SetFocus
        Exit Sub
    End If
    
    If txtUnitPrice.Text = "" Then
        MsgBox "The unit price field can not be left blank.", vbExclamation, title
        txtUnitPrice.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to save the new record ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            recProduct2.AddNew
            recProduct2!productID = txtProductID.Text & ""
            recProduct2!ProductName = txtName.Text & ""
            recProduct2!categoryid = txtCategoryID.Text & ""
            recProduct2!quantity = txtQuantity.Text & ""
            recProduct2!unitprice = Val(txtUnitPrice.Text) & ""
            recProduct2.Update
            MsgBox "New record successfully saved.", vbInformation, title
            list.Clear
            recProduct2.Requery
            If Not recProduct2.EOF Then
                recProduct2.MoveFirst
                Do While Not recProduct2.EOF
                    list.AddItem recProduct2!ProductName
                    recProduct2.MoveNext
                Loop
                recProduct2.MoveFirst
                Call display
                Call LockMe(True)
                cmdSave.Enabled = False
                cmdCancel.Enabled = False
                cmdNew.Enabled = True
                cmdUpdate.Enabled = True
                cmdDelete.Enabled = True
                list.Enabled = True
            End If
        End If
    Exit Sub
abdel:
    MsgBox "Sorry, transactions not successfully saved", vbExclamation, title

End Sub

Private Sub cmdUpdate_Click()
On Error GoTo abdel
    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see system administrator", vbCritical, title
        Exit Sub
    End If
    If com.Text = "" Then
        MsgBox "Kindly select a category in the combo box", vbExclamation, title
        com.SetFocus
        Exit Sub
    End If
    
    If txtProductID.Text = "" Then
        MsgBox "Kindly select the product to be updated ", vbExclamation, title
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to update the details of product " & Trim(txtName.Text) & " ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Call LockMe(False)
            cmdUpdate.Visible = False
            cmdOK.Visible = True
            cmdSave.Enabled = False
            cmdCancel.Enabled = True
            cmdNew.Enabled = False
            cmdDelete.Enabled = False
            list.Enabled = True
    End If
    Exit Sub
abdel:
    MsgBox "Sorry, products data coul not be updated", vbExclamation, title

End Sub

Private Sub com_Click()

    sql = "select * from product where categoryid = " & com.ItemData(com.ListIndex) & " order by productname"
    txtCategoryID.Text = com.ItemData(com.ListIndex)
    txtProductID.Text = ""
    txtName.Text = ""
    txtQuantity.Text = ""
    txtUnitPrice.Text = ""

    recProduct2.Close
    recProduct2.Open sql, con, adOpenDynamic, adLockOptimistic
    list.Clear
    recProduct2.Requery
    If Not recProduct2.EOF Then
        recProduct2.MoveFirst
        Do While Not recProduct2.EOF
            list.AddItem recProduct2!ProductName
            recProduct2.MoveNext
        Loop
        recProduct2.MoveFirst
        Else
            MsgBox "There is no product of " & Trim(com.Text) & " category in the database.", vbExclamation, title
    End If
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    Call LockMe(True)
    Label1.Visible = False
    Label3.Visible = False
    txtProductID.Visible = False
    txtCategoryID.Visible = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    com.Clear
    recCategory.Requery
    If Not recCategory.EOF Then
        recCategory.MoveFirst
        Do While Not recCategory.EOF
            com.AddItem recCategory!categoryname
            com.ItemData(com.NewIndex) = recCategory!categoryid
            recCategory.MoveNext
        Loop
        recCategory.MoveFirst
        Else
            MsgBox "There is no category record in the database", vbExclamation, title
            Call disableAll
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    recProduct2.Close
    blProduct = False
    frmScreen.Enabled = True
End Sub

Private Sub list_Click()

    position = list.ListIndex
    recProduct2.MoveFirst
    recProduct2.Move position
    Call display
End Sub
Public Sub display()

    txtProductID.Text = recProduct2!productID & ""
    txtName.Text = recProduct2!ProductName & ""
    txtCategoryID.Text = recProduct2!categoryid & ""
    txtQuantity.Text = recProduct2!quantity & ""
    txtUnitPrice.Text = recProduct2!unitprice & ""
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
    
    recGen.Open "select max(productID) from product", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogen = 1
        
        Else
        
        autogen = Val(recGen(0) + 1)
    End If
End Function
Private Sub txtName_Validate(Cancel As Boolean)

    txtName.Text = StrConv(txtName.Text, vbProperCase)
End Sub
Public Sub disableAll()
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    
End Sub
