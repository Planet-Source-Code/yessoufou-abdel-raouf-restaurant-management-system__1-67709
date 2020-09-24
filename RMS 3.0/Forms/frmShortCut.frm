VERSION 5.00
Begin VB.Form frmShortCut 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   120
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   720
   End
   Begin VB.Frame fr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   6735
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   15000
         Left            =   1080
         TabIndex        =   2
         Top             =   0
         Width           =   4215
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+F6    : Creditors"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   40
            Top             =   6600
            Width           =   3735
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+F7    : All Order Invoices"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   39
            Top             =   6960
            Width           =   3975
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSACTIONS"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   120
            TabIndex        =   38
            Top             =   2280
            Width           =   3345
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+P     : Purchase"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   37
            Top             =   2760
            Width           =   3735
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+S     : Sales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   36
            Top             =   3120
            Width           =   3735
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+F8   : Credit Orders"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   35
            Top             =   3840
            Width           =   3975
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+F7   : Cash Orders"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   34
            Top             =   3480
            Width           =   3975
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+F1    : Payslip Report"
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
            Left            =   600
            TabIndex        =   33
            Top             =   4800
            Width           =   2535
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+M     : Management"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   32
            Top             =   11880
            Width           =   3375
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+D     : Date And Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   31
            Top             =   14040
            Width           =   2775
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+F5    : Orders Report"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   26
            Top             =   6240
            Width           =   3975
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+F4    : All Sales Invoices"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   25
            Top             =   5880
            Width           =   3975
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+F3    : Sales Report"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   24
            Top             =   5520
            Width           =   3975
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+F2    : Purchase Report"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   23
            Top             =   5160
            Width           =   3975
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+U     : All Users"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   22
            Top             =   10320
            Width           =   3975
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "REPORT"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   120
            TabIndex        =   21
            Top             =   4200
            Width           =   2490
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "F11        : All Products"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   20
            Top             =   9960
            Width           =   3975
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "F9          : All Categories"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   19
            Top             =   9600
            Width           =   3975
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+C     : All Customers"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   18
            Top             =   9240
            Width           =   3975
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "F8          : Payslip"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   17
            Top             =   8880
            Width           =   3975
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "F7          : Employee Salaries"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   16
            Top             =   8520
            Width           =   3975
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+E     : View Employees"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   15
            Top             =   8160
            Width           =   3975
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "ADMINISTRATION"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   240
            TabIndex        =   14
            Top             =   7560
            Width           =   4035
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+B     : Back Up"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   13
            Top             =   11400
            Width           =   2535
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "DATABASE"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   120
            TabIndex        =   12
            Top             =   10800
            Width           =   2175
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Shift+F3  : Notepad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   11
            Top             =   13680
            Width           =   3255
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Shift+F2  : Calendar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   10
            Top             =   13320
            Width           =   3255
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Shift+F1  : Calculator"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   9
            Top             =   12960
            Width           =   3255
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+L     : Lock Application"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "FILE"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2040
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "TOOLS"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   120
            TabIndex        =   6
            Top             =   12240
            Width           =   2340
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "F4          : Quit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   5
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "F3          : Log Off"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   4
            Top             =   1200
            Width           =   2895
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "F2          : Login"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   3
            Top             =   840
            Width           =   3135
         End
      End
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      MouseIcon       =   "frmShortCut.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   6000
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   6240
      MouseIcon       =   "frmShortCut.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "frmShortCut.frx":0614
      Top             =   5520
      Width           =   480
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTINUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      MouseIcon       =   "frmShortCut.frx":0A56
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      MouseIcon       =   "frmShortCut.frx":0D60
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "STOP"
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
      Left            =   480
      MouseIcon       =   "frmShortCut.frx":106A
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   6000
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   4320
      MouseIcon       =   "frmShortCut.frx":1374
      MousePointer    =   99  'Custom
      Picture         =   "frmShortCut.frx":167E
      Top             =   5520
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2400
      MouseIcon       =   "frmShortCut.frx":1AC0
      MousePointer    =   99  'Custom
      Picture         =   "frmShortCut.frx":1DCA
      Top             =   5520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      MouseIcon       =   "frmShortCut.frx":220C
      MousePointer    =   99  'Custom
      Picture         =   "frmShortCut.frx":2516
      Top             =   5520
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   480
      X2              =   7200
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   5535
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   7695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   7695
   End
End
Attribute VB_Name = "frmShortCut"
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
    Dim currentLength As Byte
    Const msg As String = "SHORTCUT KEYS"
    
Private Sub Form_Load()
    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub

Private Sub Image1_Click()
    Timer1.Enabled = False
End Sub

Private Sub Image2_Click()
    Timer1.Enabled = True
    Frame2.Top = 120
End Sub

Private Sub Image3_Click()
    Timer1.Enabled = True
End Sub

Private Sub Image4_Click()
    If MsgBox("Are you sure you want to close this window ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub Label36_Click()
    Timer1.Enabled = False
End Sub

Private Sub Label37_Click()
    Timer1.Enabled = True
    Frame2.Top = 120
End Sub

Private Sub Label38_Click()
    Timer1.Enabled = True
End Sub

Private Sub Label39_Click()
    If MsgBox("Are you sure you want to close this window ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub Timer1_Timer()
    Frame2.Top = Frame2.Top - 20
    If Frame2.Top <= -9000 Then
        Frame2.Top = 120
    End If
End Sub

Private Sub Timer2_Timer()
    Label1.Caption = Left(msg, currentLength)
    currentLength = (currentLength + 1) Mod (Len(msg) + 1)
End Sub


