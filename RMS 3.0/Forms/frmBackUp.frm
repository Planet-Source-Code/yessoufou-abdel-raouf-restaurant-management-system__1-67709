VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBackUp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Last Back Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7335
      Begin VB.Label lblLastTime 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblLastDate 
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
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblLastPath 
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
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   1680
         Width           =   5535
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1185
      End
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   3600
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   4080
      TabIndex        =   8
      Top             =   3450
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   767
      Caption         =   "Close"
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
      mIcon           =   "frmBackUp.frx":0000
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   435
      Left            =   960
      TabIndex        =   9
      Top             =   3450
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   767
      Caption         =   "New Backup"
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
      mIcon           =   "frmBackUp.frx":031A
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   7575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATABASE BACKUP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   3675
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   3375
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmBackUp"
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
    Dim FileSyst As New FileSystemObject
    Dim backUpFile As File
    Dim databaseName As String
    Dim destination As String
    Dim source As String
    Dim currDate As String
    Dim currtime As String

Private Sub cmdClose_Click()
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub cmdSave_Click()
On Error GoTo abdel
    Dim fso As Object, MyFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    dialog.DialogTitle = " Back Up "
    dialog.Filter = "Microsoft Access Database (*.mdb)|*.mdb"
    dialog.ShowSave
    
    If dialog.FileName = "" Or dialog.FileName = " " Then
        MsgBox " You chose to cancel the back up, therefore nothing is done"
        Exit Sub
        Else
            Set MyFile = fso.GetFile(App.Path & "\AbdelSoft.mdb")
            MyFile.Copy dialog.FileName
    'FileCopy App.Path & "\inventry.mdb", Inventrymain.dialog.FileName
    
            MsgBox " Database successfully saved to " & dialog.FileName, vbInformation
    End If
    
    destination = dialog.FileName
    currDate = Date
    currtime = Time

    SaveSetting App.title, "Settings", "BackupPath", destination
    SaveSetting App.title, "Settings", "BackupDate", currDate
    SaveSetting App.title, "Settings", "BackupTime", currtime
    Unload Me
Exit Sub
abdel:
    MsgBox "Sorry, an error occured while saving database", vbExclamation, title
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Dim lastPath As String
    Dim lastdate As String
    Dim lasttime As String
    
    'Read Registry for previous settings stored
    lastPath = GetSetting(App.title, "Settings", "BackupPath")
    lastdate = GetSetting(App.title, "Settings", "BackupDate")
    lasttime = GetSetting(App.title, "Settings", "BackupTime")
    
    If lastPath = "" Then
        lblLastPath.Caption = "No backUp had been previously made"
        lblLastDate.Caption = " "
        lblLastTime.Caption = " "
    Else
        lblLastPath.Caption = lastPath
        lblLastDate.Caption = lastdate & "  (Month-Date-year)"
        lblLastTime.Caption = lasttime
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub
