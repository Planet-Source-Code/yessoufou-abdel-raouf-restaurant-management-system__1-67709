VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   435
      Left            =   2550
      TabIndex        =   6
      Top             =   2130
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Caption         =   "&Cancel"
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
      mIcon           =   "frmLogin.frx":0000
   End
   Begin lvButton.lvButtons_H cmdLogin 
      Default         =   -1  'True
      Height          =   435
      Left            =   480
      TabIndex        =   5
      Top             =   2130
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Caption         =   "&Login"
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
      mIcon           =   "frmLogin.frx":031A
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
      Left            =   1800
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Password"
      Top             =   1440
      Width           =   2415
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
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   2
      ToolTipText     =   "Login ID"
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   4455
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
      Top             =   1440
      Width           =   1815
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
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmLogin"
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
Dim prompt As String
Dim counter As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub CmdCancel_Click()
    If MsgBox("Are you sure you want to quit the application ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            End
    End If
End Sub

Private Sub cmdLogin_Click()

    If txtPassword.Text = "" Then
        MsgBox "The Password field can not be left blank ", vbExclamation, title
        txtPassword.SetFocus
        Exit Sub
    End If
    txtPassword.Text = LCase(txtPassword.Text)
    recUsers.MoveFirst
    Do While Not recUsers.EOF
        If Trim(recUsers!loginname) = Trim(txtLoginName.Text) And _
        Trim(recUsers!password) = Trim(txtPassword.Text) Then
            sndPlaySound App.Path & "\Media\reminder.wav", &H1
            frmScreen.lblRole.Caption = recUsers!role
            frmScreen.lblName.Caption = txtLoginName.Text
            Call UserLogin
            frmScreen.Enabled = True
            Me.Hide
            counter = 1
            App.HelpFile = App.Path & "\abdelsoft.hlp"
            Exit Sub
        End If
        recUsers.MoveNext
    Loop
    
    prompt = "This Is The Fourth Time You Typed In Wrong Password."
    prompt = prompt & Chr(10) & Chr(13) & " One More Mistake And You Will Be Logged off"
    If recUsers.EOF Then
        MsgBox "Invalid password, kindly retry", vbExclamation, title
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
        txtPassword.SetFocus
        If counter = 3 Then
            MsgBox prompt, vbExclamation, title
            ElseIf counter = 4 Then
                End
        End If
        counter = counter + 1
        Exit Sub
    End If
End Sub

Private Sub Form_Load()

    frmScreen.Show
    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    cmdLogin.Enabled = False
    If recUsers.BOF And recUsers.EOF Then
        MsgBox "There's no user in the Database", vbExclamation, title
        txtLoginName.Enabled = False
        txtPassword.Enabled = False
        Exit Sub
    End If
    
End Sub

Private Sub txtLoginName_Validate(Cancel As Boolean)
    txtLoginName.Text = StrConv(txtLoginName.Text, vbProperCase)
End Sub

Private Sub txtPassword_GotFocus()

    recUsers.MoveFirst
    Do While Not recUsers.EOF
        If Trim(recUsers!loginname) = Trim(txtLoginName.Text) Then
            cmdLogin.Enabled = True
            Exit Sub
        End If
        recUsers.MoveNext
    Loop
    
    If recUsers.EOF Then
        MsgBox "Invalid login name, kindly retry", vbExclamation, title
        txtLoginName.SelStart = 0
        txtLoginName.SelLength = Len(txtLoginName.Text)
        txtLoginName.SetFocus
        Exit Sub
    End If
End Sub


