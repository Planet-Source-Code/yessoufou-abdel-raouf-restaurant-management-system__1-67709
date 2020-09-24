VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUsersLog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   19660801
      CurrentDate     =   38562
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   19660801
      CurrentDate     =   38562
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   435
      Left            =   2430
      TabIndex        =   4
      Top             =   2010
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
      mIcon           =   "frmUsersLog.frx":0000
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   435
      Left            =   360
      TabIndex        =   5
      Top             =   2010
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
      mIcon           =   "frmUsersLog.frx":031A
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TO"
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
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FROM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   660
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   1935
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmUsersLog"
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
Dim reply As Integer
Dim sql As String

Private Sub cmdClose_Click()
    If MsgBox("Are you sure you want to close this window ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub



Private Sub cmdOk_Click()
On Error GoTo abdel
   If AllUsers = True And admin = False And emp = False Then
    
    sql = "SELECT * FROM userslog where logindate Between #" & DTPicker1.Value & "# AND #" & DTPicker2.Value & "#"
    Set UsersReport.DataSource = con.Execute(sql)
    UsersReport.Sections("Section2").Controls.Item("lblstatus").Caption = "All Users (Administrator and Employee) Time Record"
    UsersReport.Sections("Section2").Controls.Item("lblasof").Caption = "as of " & DTPicker1.Value & " to " & DTPicker2.Value
    UsersReport.Show
    Unload Me
    
ElseIf admin = True And AllUsers = False And emp = False Then
    sql = "SELECT * FROM UsersLog where logindate Between #" & DTPicker1.Value & "# AND #" & DTPicker2.Value & "# AND role = 'Administrator'"
    Set UsersReport.DataSource = con.Execute(sql)
    UsersReport.Sections("Section2").Controls.Item("lblstatus").Caption = "Administrator Time Record"
    UsersReport.Sections("Section2").Controls.Item("lblasof").Caption = "as of " & DTPicker1.Value & " to " & DTPicker2.Value
    UsersReport.Show
    Unload Me
    
  ElseIf emp = True And AllUsers = False And admin = False Then
    sql = "SELECT * FROM userslog where logindate Between #" & DTPicker1.Value & "# AND #" & DTPicker2.Value & "# AND role = 'Employee'"
    Set UsersReport.DataSource = con.Execute(sql)
    UsersReport.Sections("Section2").Controls.Item("lblstatus").Caption = "Employees Time Record"
    UsersReport.Sections("Section2").Controls.Item("lblasof").Caption = "as of " & DTPicker1.Value & " to " & DTPicker2.Value
    UsersReport.Show
    Unload Me
End If
    Exit Sub
abdel:
    MsgBox "Report not available for now...", vbExclamation, title
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = True
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    'Call Connection
    DTPicker1.Value = Date
    DTPicker2.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AllUsers = False
    admin = False
    emp = False
    frmScreen.Enabled = True
End Sub
