VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   5115
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   4000
      Left            =   1080
      Top             =   4680
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   600
      Top             =   4680
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   4680
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   240
      Picture         =   "frmAbout.frx":4926
      ScaleHeight     =   1515
      ScaleWidth      =   1185
      TabIndex        =   3
      Top             =   240
      Width           =   1185
      Begin VB.Line Line1 
         X1              =   0
         X2              =   1200
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2835
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   7605
      Begin VB.PictureBox picCredits 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         Height          =   2490
         Left            =   120
         ScaleHeight     =   162
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   487
         TabIndex        =   1
         Top             =   240
         Width           =   7365
         Begin VB.TextBox txtCredits 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   3255
            Left            =   1560
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   2
            Text            =   "frmAbout.frx":622E
            Top             =   3000
            Width           =   4410
         End
      End
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   7500
      TabIndex        =   7
      Top             =   60
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   767
      Caption         =   "X"
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
      mIcon           =   "frmAbout.frx":64B1
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NOM:      YESSOUFOU ABDEL RAOUF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   1680
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      Height          =   1215
      Left            =   480
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "yessoufouabdel@yahoo.fr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2880
      MouseIcon       =   "frmAbout.frx":67CB
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Contacter Abdel"
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Height          =   5115
      Left            =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmAbout"
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
    Dim success As Long
    Public SW_SHOWNORMAL
    Dim email As String
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Dim prompt As String

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    TextEffect Me, "", 12, 12, , 128, 0, RGB(&H80, 0, 0)
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Me.Top = 1700
    Me.Move (Screen.Width - Width) / 2
    cmdClose.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Label3.ForeColor = &HFF0000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub

Private Sub Label3_Click()
    email = "yessoufouabdel@yahoo.fr"
    success = ShellExecute(0&, vbNullString, "mailto:" & email, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label3.ForeColor = &HC0&
End Sub

Private Sub Timer1_Timer()
If txtCredits.Top > 0 - (txtCredits.Height) Then
    txtCredits.Top = txtCredits.Top - 1
Else
    txtCredits.Visible = False
    Timer2.Enabled = True
    Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
    txtCredits.Top = 170
    txtCredits.Visible = True
    Timer1.Enabled = True
    Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
    cmdClose.Enabled = True
End Sub
