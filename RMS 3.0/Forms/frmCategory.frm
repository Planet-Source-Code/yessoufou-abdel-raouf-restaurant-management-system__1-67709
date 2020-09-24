VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView list 
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Double click to modify a category details"
      Top             =   720
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5106
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Category Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   5557
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdProduct 
      Height          =   435
      Left            =   330
      TabIndex        =   2
      Top             =   3990
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   767
      Caption         =   "&View Product"
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
      mIcon           =   "frmCategory.frx":0000
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   435
      Left            =   1800
      TabIndex        =   3
      Top             =   3990
      Width           =   1215
      _ExtentX        =   2143
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
      mIcon           =   "frmCategory.frx":031A
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   435
      Left            =   3240
      TabIndex        =   4
      Top             =   3990
      Width           =   1215
      _ExtentX        =   2143
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
      mIcon           =   "frmCategory.frx":0634
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   435
      Left            =   4740
      TabIndex        =   5
      Top             =   3990
      Width           =   1215
      _ExtentX        =   2143
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
      mIcon           =   "frmCategory.frx":094E
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   6210
      TabIndex        =   6
      Top             =   3990
      Width           =   1215
      _ExtentX        =   2143
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
      mIcon           =   "frmCategory.frx":0C68
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORIES"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   7695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   3855
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmCategory"
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
Private Sub cmdClose_Click()
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            If blCategory = True Then
                Load frmProduct
                frmProduct.Show
            End If
    End If
    Unload Me
End Sub

Private Sub cmddelete_Click()
On Error GoTo abdel
    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see system administrator...", vbCritical, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete the record of category " & Trim(list.SelectedItem.ListSubItems(1)) & " ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete from category where categoryid = " & Trim(list.SelectedItem.Text)
            list.ListItems.Remove list.SelectedItem.Index
            MsgBox "Category records successfully deleted", vbInformation, title
            
            If list.ListItems.Count = 0 Then
                cmdUpdate.Enabled = False
                cmdDelete.Enabled = False
            End If
    End If
Exit Sub
abdel:
    MsgBox "Sorry, category record could not be deleted", vbExclamation, title
End Sub

Private Sub cmdNew_Click()

    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see the administrator...", vbCritical, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to add new category record?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            blAddCategory = True
            blUpdateCategory = False
            Unload Me
            Load frmAddCategory
            frmAddCategory.Show
            frmAddCategory.txtID.Text = frmAddCategory.autogen
    End If
End Sub

Private Sub cmdProduct_Click()

    blProduct = True
    Unload Me
    Load frmProduct
    frmProduct.Show
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo abdel
    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see system administrator", vbCritical, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to modify the record of " & Trim(list.SelectedItem.SubItems(1)) & " ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            blAddCategory = False
            blUpdateCategory = True
            recCategory.Requery
            If Not recCategory.EOF Then
                recCategory.MoveFirst
                Do While Not recCategory.EOF
                    If Trim(recCategory!categoryid) = Trim(list.SelectedItem.Text) Then
                        Unload Me
                        Load AddCategory
                        AddCategory.Show
                        AddCategory.txtID = recCategory!categoryid & ""
                        AddCategory.txtName.Text = recCategory!categoryname & ""
                        AddCategory.txtDescription.Text = recCategory!Description & ""
                        Exit Sub
                    End If
                    recCategory.MoveNext
                Loop
                Else
            End If
    End If
Exit Sub
abdel:
    MsgBox "Sorry, category records could not be updated", vbExclamation, title
End Sub

Private Sub cmdView_Click()

End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    
    recCategory.Requery
        If Not recCategory.EOF Then
            recCategory.MoveFirst
            Do While Not recCategory.EOF
                Set lst = list.ListItems.Add(, , recCategory!categoryid)
                    lst.ListSubItems.Add , , recCategory!categoryname
                    lst.ListSubItems.Add , , recCategory!Description & ""
                    recCategory.MoveNext
            Loop
            Else
                MsgBox "There is no category record in the database.", vbExclamation, title
                cmdUpdate.Enabled = False
                cmdDelete.Enabled = False
                cmdProduct.Enabled = False
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blCategory = False
    frmScreen.Enabled = True
End Sub
Private Sub list_DblClick()
On Error GoTo abdel
    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see system administrator", vbCritical, title
        Exit Sub
    End If
    blAddCategory = False
    blUpdateCategory = True
    recCategory.Requery
    If Not recCategory.EOF Then
        recCategory.MoveFirst
        Do While Not recCategory.EOF
            If Trim(recCategory!categoryid) = Trim(list.SelectedItem.Text) Then
                Unload Me
                Load AddCategory
                AddCategory.Show
                AddCategory.txtID = recCategory!categoryid & ""
                AddCategory.txtName.Text = recCategory!categoryname & ""
                AddCategory.txtDescription.Text = recCategory!Description & ""
                Exit Sub
            End If
            recCategory.MoveNext
        Loop
        Else
    End If
Exit Sub
abdel:
    MsgBox "Sorry, category records could not be updated", vbExclamation, title
End Sub
