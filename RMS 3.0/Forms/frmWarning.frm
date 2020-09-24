VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmWarning 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   3750
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
      mIcon           =   "frmWarning.frx":0000
   End
   Begin lvButton.lvButtons_H cmdContinue 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "&Continue"
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
      mIcon           =   "frmWarning.frx":031A
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "However, to proceed click continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWarning.frx":0634
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmWarning"
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
    If MsgBox("Are you sure you want to close this window ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub cmdContinue_Click()

    Unload Me
    Load frmEmptyDatabase
    frmEmptyDatabase.Show
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub
