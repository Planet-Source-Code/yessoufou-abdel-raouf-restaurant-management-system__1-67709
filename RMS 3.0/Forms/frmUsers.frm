VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView list 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Double click to modify a uer record"
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Login Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Role"
         Object.Width           =   3969
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   435
      Left            =   1680
      TabIndex        =   2
      Top             =   3330
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      Caption         =   "&Update"
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
      mIcon           =   "frmUsers.frx":0000
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   435
      Left            =   270
      TabIndex        =   3
      Top             =   3330
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      Caption         =   "&New"
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
      mIcon           =   "frmUsers.frx":031A
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   435
      Left            =   3030
      TabIndex        =   4
      Top             =   3330
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      Caption         =   "&Delete"
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
      mIcon           =   "frmUsers.frx":0634
   End
   Begin lvButton.lvButtons_H cmdQuit 
      Height          =   435
      Left            =   4500
      TabIndex        =   5
      Top             =   3330
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
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
      mIcon           =   "frmUsers.frx":094E
   End
   Begin VB.Label Label4 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   3255
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmUsers"
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
Private Sub cmddelete_Click()
On Error GoTo abdel
    If list.SelectedItem.Text = frmScreen.lblName.Caption Then
        MsgBox "Sorry, You Can Not Delete This User's details ", vbExclamation, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete user " & list.SelectedItem.Text & " ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete from users where loginname = '" & list.SelectedItem.Text & "'"
            list.ListItems.Remove list.SelectedItem.Index
    End If
    Exit Sub
abdel:
    MsgBox "Sorry, users data could not deleted", vbExclamation, title

End Sub

Private Sub cmdNew_Click()

    If MsgBox("Are you sure you want to add a new user ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            blAddUser = True
            blUpdateUser = False
            Unload Me
            Load frmAddUser
            frmAddUser.Show
            frmAddUser.lbl.Caption = frmAddUser.autogen
    End If
    
End Sub

Private Sub cmdQuit_Click()
    If MsgBox("Are you sure you want to close this window ?", 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo abdel
    If list.SelectedItem.Text = frmScreen.lblName.Caption Then
        MsgBox "Sorry, you can not modify this user's details ", vbExclamation, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to update the details of user " & list.SelectedItem.Text & " ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            blAddUser = False
            blUpdateUser = True
            recUsers.Requery
            recUsers.MoveFirst
            Do While Not recUsers.EOF
                If Trim(recUsers!loginname) = Trim(list.SelectedItem.Text) Then
                    Unload Me
                    AddUser.Show
                    AddUser.lbl.Caption = recUsers!userid
                    AddUser.txtLoginName.Text = recUsers!loginname
                    AddUser.txtPass.Text = recUsers!password
                    AddUser.txtPassword.Text = recUsers!password
                    AddUser.comboRole.Text = recUsers!role
                    Exit Sub
                End If
                recUsers.MoveNext
            Loop
    End If
    Exit Sub
abdel:
    MsgBox "Sorry, users data could not be updated", vbExclamation, title

End Sub

Private Sub Form_Load()
    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    
    recUsers.MoveFirst
    Do While Not recUsers.EOF
        Set lst = list.ListItems.Add(, , recUsers!loginname)
        lst.ListSubItems.Add , , recUsers!role
        recUsers.MoveNext
    Loop
    recUsers.MoveFirst
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub

Private Sub list_DblClick()
On Error GoTo abdel
    If list.SelectedItem.Text = frmScreen.lblName.Caption Then
        MsgBox "Sorry, you can not modify this user's details ", vbExclamation, title
        Exit Sub
    End If
    blAddUser = False
    blUpdateUser = True
    recUsers.Requery
    recUsers.MoveFirst
    Do While Not recUsers.EOF
        If Trim(recUsers!loginname) = Trim(list.SelectedItem.Text) Then
            Unload Me
            frmAddUser.Show
            frmAddUser.lbl.Caption = recUsers!userid
            frmAddUser.txtLoginName.Text = recUsers!loginname
            frmAddUser.txtPass.Text = recUsers!password
            frmAddUser.txtPassword.Text = recUsers!password
            frmAddUser.comboRole.Text = recUsers!role
            Exit Sub
        End If
        recUsers.MoveNext
    Loop
    Exit Sub
abdel:
    MsgBox "Sorry, users data could not be updated", vbExclamation, title

End Sub
