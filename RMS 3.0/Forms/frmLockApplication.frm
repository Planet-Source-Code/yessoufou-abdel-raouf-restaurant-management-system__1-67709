VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLockApplication 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Enter Password"
      Top             =   960
      Width           =   2895
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   435
      Left            =   2580
      TabIndex        =   3
      Top             =   1890
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
      mIcon           =   "frmLockApplication.frx":0000
   End
   Begin lvButton.lvButtons_H cmdOK 
      Default         =   -1  'True
      Height          =   435
      Left            =   510
      TabIndex        =   4
      Top             =   1890
      Width           =   1335
      _ExtentX        =   2355
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
      mIcon           =   "frmLockApplication.frx":031A
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOCKED"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmLockApplication"
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
        txtPassword.SetFocus
        Exit Sub
        Else
            sndPlaySound App.Path & "\Media\reminder.wav", &H1
            Call UserLogout
            End
    End If
End Sub


Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
End Sub

Private Sub cmdOk_Click()

    If txtPassword.Text = "" Then
        MsgBox "The Password field can not be left blank.", vbExclamation, title
        txtPassword.SetFocus
        Exit Sub
    End If
    txtPassword.Text = LCase(txtPassword.Text)
    If Trim(txtPassword.Text) = Trim(frmLogin.txtPassword.Text) Then
        Unload Me
        frmScreen.Enabled = True
        Else
            MsgBox "Incorrect password, please retry", vbExclamation, title
            txtPassword.SelStart = 0
            txtPassword.SelLength = Len(txtPassword.Text)
            txtPassword.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub
