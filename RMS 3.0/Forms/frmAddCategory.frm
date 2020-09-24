VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddCategory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDescription 
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
      Height          =   1095
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Description"
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtName 
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
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   "Category Name"
      Top             =   1320
      Width           =   3015
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
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "ID"
      Top             =   720
      Width           =   735
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   435
      Left            =   3300
      TabIndex        =   7
      Top             =   3450
      Width           =   1215
      _ExtentX        =   2143
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
      mIcon           =   "frmAddCategory.frx":0000
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   435
      Left            =   540
      TabIndex        =   8
      Top             =   3450
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Caption         =   "&OK"
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
      mIcon           =   "frmAddCategory.frx":031A
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
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Category Name"
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
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   3375
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmAddCategory"
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
    If MsgBox("Are you sure you want to close this window?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
            Load frmCategory
            frmCategory.Show
    End If
End Sub

Private Sub cmdOk_Click()
On Error GoTo abdel
    If txtName.Text = "" Then
        MsgBox "The category name field can not be left blank.", vbExclamation, title
        txtName.SetFocus
        Exit Sub
    End If
    
    If blAddCategory = True And blUpdateCategory = False Then
        recCategory.AddNew
        recCategory!categoryid = Trim(txtID.Text) & ""
        recCategory!categoryname = Trim(txtName.Text) & ""
        recCategory!Description = Trim(txtDescription.Text) & ""
        recCategory.Update
        Unload Me
        Load frmCategory
        frmCategory.Show
    End If
    
    If blUpdateCategory = True And blAddCategory = False Then
        recCategory!categoryid = Trim(txtID.Text) & ""
        recCategory!categoryname = Trim(txtName.Text) & ""
        recCategory!Description = Trim(txtDescription.Text) & ""
        recCategory.UpdateBatch adAffectCurrent
        Unload Me
        Load frmCategory
        frmCategory.Show
    End If
Exit Sub
abdel:
    MsgBox "Transactions not successfully saved", vbExclamation, title
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
End Sub
Public Function autogen()

    Dim recGen As New Recordset
    
    recGen.Open "select max(categoryID) from category", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogen = 1
        
        Else
        
        autogen = Val(recGen(0) + 1)
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub

Private Sub txtDescription_Validate(Cancel As Boolean)
    txtDescription.Text = StrConv(txtDescription.Text, vbProperCase)
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    txtName.Text = StrConv(txtName.Text, vbProperCase)
End Sub
