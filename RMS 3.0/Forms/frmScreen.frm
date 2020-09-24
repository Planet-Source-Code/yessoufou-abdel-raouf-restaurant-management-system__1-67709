VERSION 5.00
Object = "{57D851A8-1A2B-40A8-8D6F-9845A46F7124}#1.0#0"; "PullDownMenu.ocx"
Begin VB.Form frmScreen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin PullDownMenu.ctrl_PullDownMenu ctrl_PullDownMenu 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   19995
      _ExtentX        =   35269
      _ExtentY        =   661
      ForeColor       =   16777215
      BackColor       =   16711680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   11160
      Width           =   735
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   6360
      Top             =   11040
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   2760
      Top             =   11040
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   0
      Top             =   11040
      Width           =   19995
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7560
      TabIndex        =   5
      Top             =   11160
      Width           =   13005
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   6360
      Picture         =   "frmScreen.frx":0000
      Top             =   11040
      Width           =   480
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   11160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   11160
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   1050
      Left            =   2760
      Picture         =   "frmScreen.frx":0442
      Top             =   11040
      Width           =   900
   End
   Begin VB.Label lblRole 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   11160
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmScreen.frx":2066
      Top             =   11040
      Width           =   480
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   9240
      TabIndex        =   1
      Top             =   11160
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   45
      Left            =   10500
      TabIndex        =   0
      Top             =   8640
      Width           =   15
   End
   Begin VB.Image Image2 
      Height          =   1275
      Left            =   0
      Picture         =   "frmScreen.frx":7C78
      Top             =   11040
      Width           =   42525
   End
   Begin VB.Image Image1 
      Height          =   1275
      Left            =   0
      Picture         =   "frmScreen.frx":1327F
      Top             =   0
      Width           =   42525
   End
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Unload frmScreen
    Unload frmMain
End Sub
Private Sub ctrl_PullDownMenu_Click(Index As Integer)

    Select Case Index
        Case 1:
            PopupMenu frmMain.mnuFile, , frmScreen.ctrl_PullDownMenu.Left + frmScreen.ctrl_PullDownMenu.pSelectionLeft, frmScreen.ctrl_PullDownMenu.Top + frmScreen.ctrl_PullDownMenu.pSelectionBottom
        Case 2:
            PopupMenu frmMain.mnuTrans, , frmScreen.ctrl_PullDownMenu.Left + frmScreen.ctrl_PullDownMenu.pSelectionLeft, frmScreen.ctrl_PullDownMenu.Top + frmScreen.ctrl_PullDownMenu.pSelectionBottom
        Case 3:
            PopupMenu frmMain.mnuReport, , frmScreen.ctrl_PullDownMenu.Left + frmScreen.ctrl_PullDownMenu.pSelectionLeft, frmScreen.ctrl_PullDownMenu.Top + frmScreen.ctrl_PullDownMenu.pSelectionBottom
        Case 4:
            PopupMenu frmMain.mnuAdmin, , frmScreen.ctrl_PullDownMenu.Left + frmScreen.ctrl_PullDownMenu.pSelectionLeft, frmScreen.ctrl_PullDownMenu.Top + frmScreen.ctrl_PullDownMenu.pSelectionBottom
        Case 5:
            PopupMenu frmMain.mnuData, , frmScreen.ctrl_PullDownMenu.Left + frmScreen.ctrl_PullDownMenu.pSelectionLeft, frmScreen.ctrl_PullDownMenu.Top + frmScreen.ctrl_PullDownMenu.pSelectionBottom
        Case 6:
            PopupMenu frmMain.mnuTools, , frmScreen.ctrl_PullDownMenu.Left + frmScreen.ctrl_PullDownMenu.pSelectionLeft, frmScreen.ctrl_PullDownMenu.Top + frmScreen.ctrl_PullDownMenu.pSelectionBottom
        Case 7:
            PopupMenu frmMain.mnuHelp, , frmScreen.ctrl_PullDownMenu.Left + frmScreen.ctrl_PullDownMenu.pSelectionLeft, frmScreen.ctrl_PullDownMenu.Top + frmScreen.ctrl_PullDownMenu.pSelectionBottom
    End Select
    
End Sub

Private Sub Form_Load()
'Call connect
    Call Initialize
    lblDate.Caption = Format(Date, "Medium Date")
    lblTime.Caption = Time
    Call ConnectMe
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
    Unload frmScreen
End Sub

Private Sub Timer1_Timer()
    lblTime.Caption = Time
End Sub
