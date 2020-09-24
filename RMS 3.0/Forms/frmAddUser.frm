VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox comboRole 
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
      ItemData        =   "frmAddUser.frx":0000
      Left            =   2040
      List            =   "frmAddUser.frx":000A
      TabIndex        =   7
      ToolTipText     =   "Select A Role"
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   6
      ToolTipText     =   "Retype The User Password"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtPass 
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
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   5
      ToolTipText     =   "Password"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtLoginName 
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
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   4
      ToolTipText     =   "Login Name"
      Top             =   600
      Width           =   2535
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   435
      Left            =   3000
      TabIndex        =   10
      Top             =   3210
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
      mIcon           =   "frmAddUser.frx":0027
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   435
      Left            =   360
      TabIndex        =   11
      Top             =   3210
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
      mIcon           =   "frmAddUser.frx":0341
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERS"
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
      Left            =   840
      TabIndex        =   9
      Top             =   0
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Role"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Retype Password"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
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
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   3135
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label lbl 
      Caption         =   "Label5"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddUser"
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
Dim control As Object

Private Sub CmdCancel_Click()
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
            Load frmUsers
            frmUsers.Show
    End If
End Sub

Private Sub cmdOk_Click()
On Error GoTo abdel
    For Each control In Me
        If TypeOf control Is TextBox Then
            If control.Text = "" Then
                MsgBox "No fill should be left blank !", vbExclamation, title
                Exit Sub
            End If
        End If
    Next
    
    If comboRole.Text = "" Then
        MsgBox "Kindly choose your role", vbExclamation, title
        comboRole.SetFocus
        Exit Sub
    End If
    If blAddUser = True And blUpdateUser = False Then
        recUsers.Requery
        recUsers.MoveFirst
        Do While Not recUsers.EOF
            If Trim(recUsers!loginname) = Trim(txtLoginName.Text) Then
                MsgBox "Login name already exists, kindly change it", vbExclamation, title
                txtLoginName.SelStart = 0
                txtLoginName.SelLength = Len(txtLoginName.Text)
                txtLoginName.SetFocus
                Exit Sub
            End If
            recUsers.MoveNext
        Loop
        
        If recUsers.EOF Then
            If txtPass.Text <> txtPassword.Text Then
                MsgBox "Your passwords are not correct, please check them", vbExclamation, title
                txtPass.SelStart = 0
                txtPass.SelLength = Len(txtPass.Text)
                txtPass.SetFocus
                Exit Sub
                Else
                    recUsers.AddNew
                    recUsers!userid = Trim(lbl.Caption)
                    recUsers!loginname = Trim(txtLoginName.Text)
                    recUsers!password = Trim(txtPassword.Text)
                    recUsers!role = Trim(comboRole.Text)
                    recUsers.Update
                    Unload Me
                    Load frmUsers
                    frmUsers.Show
                    Exit Sub
            End If
        End If
    End If
    
    If blUpdateUser = True And blAddUser = False Then
        'recUsers.MoveFirst
            If Trim(txtPass.Text) <> Trim(txtPassword.Text) Then
                MsgBox "Your passwords are not correct, please check them", vbExclamation, title
                txtPass.SelStart = 0
                txtPass.SelLength = Len(txtPass.Text)
                txtPass.SetFocus
                Exit Sub
            End If
            
            recUsers!userid = lbl.Caption
            recUsers!loginname = Trim(txtLoginName.Text)
            recUsers!password = Trim(txtPassword.Text)
            recUsers!role = Trim(comboRole.Text)
            recUsers.UpdateBatch adAffectCurrent
            Unload Me
            Load frmUsers
            frmUsers.Show
    End If
Exit Sub
abdel:
    MsgBox "Sorry, transactions not successfully saved", vbExclamation, title
End Sub

Private Sub comboRole_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blAddUser = False
    blUpdateUser = False
    frmScreen.Enabled = True
End Sub

Private Sub txtLoginName_Validate(Cancel As Boolean)
    txtLoginName.Text = StrConv(txtLoginName.Text, vbProperCase)
End Sub
Public Function autogen()

    Dim recGen As New Recordset
    
    recGen.Open "select max(UserID) from Users", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogen = 1
        
        Else
        
        autogen = Val(recGen(0) + 1)
    End If
End Function
