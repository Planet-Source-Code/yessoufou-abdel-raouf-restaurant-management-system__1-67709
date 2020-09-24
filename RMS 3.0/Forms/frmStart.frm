VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStart 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   6960
   ClientLeft      =   1320
   ClientTop       =   2265
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStart.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar bar 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   6120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer3 
      Interval        =   60
      Left            =   120
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Interval        =   60
      Left            =   120
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   240
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   8
      Height          =   6975
      Left            =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   75
   End
End
Attribute VB_Name = "frmStart"
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

    Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub Form_Load()

    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    'sndPlaySound App.Path & "/" & "beeth5th.wav", &H1
    Timer3.Enabled = False
    bar.Visible = False
    
End Sub

Private Sub Timer2_Timer()

    Timer3.Enabled = True
    bar.Visible = True
    
End Sub

Private Sub Timer3_Timer()

    bar.Value = bar.Value + 1
    Select Case bar.Value
        Case 5
            Label9.Caption = "Loading file_00100"
        Case 10
            Label9.Caption = "Loading file_98765"
        Case 15
            Label9.Caption = "Setting Pictures"
        Case 20
            Label9.Caption = "Loading files Loop 564"
        Case 25
            Label9.Caption = "Loading file ABDEL.html"
        Case 30
            Label9.Caption = "Setting Images_123100"
        Case 35
            Label9.Caption = "Loading file_990100"
        Case 40
            Label9.Caption = "Loading file_001"
        Case 45
            Label9.Caption = "Loading file_11100"
        Case 50
            Label9.Caption = "Loading file_00100"
        Case 55
            Label9.Caption = "Loading file_98765"
        Case 60
            Label9.Caption = "Setting Pictures"
        Case 65
            Label9.Caption = "Loading files Loop 564"
        Case 70
            Label9.Caption = "Loading file ABDEL.html"
        Case 75
            Label9.Caption = "Setting Images_123100"
        Case 80
            Label9.Caption = "Loading Flash.DLL"
        Case 85
            Label9.Caption = "Loading PIL.DLL"
        Case 90
            Label9.Caption = "Loading Report Session"
        
            Label9.Caption = "Loading All Session"
    
    End Select
    
    If bar.Value = 100 Then
        Unload Me
        frmLogin.Show
    End If
    
End Sub
