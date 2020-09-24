VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSalary 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView list 
      Height          =   3135
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Double click to modify an employee salary"
      Top             =   600
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "First Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Last Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Salary"
         Object.Width           =   3881
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   3630
      TabIndex        =   2
      Top             =   4050
      Width           =   1935
      _ExtentX        =   3413
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
      mIcon           =   "frmSalary.frx":0000
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   435
      Left            =   870
      TabIndex        =   3
      Top             =   4050
      Width           =   1935
      _ExtentX        =   3413
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
      mIcon           =   "frmSalary.frx":031A
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SALARIES"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   6975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   3975
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmSalary"
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
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo abdel
    recEmployee.Requery
    If Not recEmployee.EOF Then
        recEmployee.MoveFirst
        Do While Not recEmployee.EOF
            If Trim(recEmployee!firstname) = Trim(list.SelectedItem.Text) And _
                Trim(recEmployee!lastname) = Trim(list.SelectedItem.ListSubItems(1)) Then
                    Unload Me
                    Load frmUpdateSalary
                    frmUpdateSalary.Show
                    frmUpdateSalary.txtFirstName.Text = Trim(recEmployee!firstname)
                    frmUpdateSalary.txtLastName.Text = Trim(recEmployee!lastname)
                    frmUpdateSalary.txtSalary.Text = Trim(recEmployee!Salary)
                    Exit Sub
            End If
            recEmployee.MoveNext
        Loop
        Else
    End If
    Exit Sub
abdel:
    MsgBox "Sorry, salary data could not be updated", vbExclamation, title

End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    
    recEmployee.Requery
    If Not recEmployee.EOF Then
        recEmployee.MoveFirst
        Do While Not recEmployee.EOF
            Set lst = list.ListItems.Add(, , recEmployee!firstname)
                lst.ListSubItems.Add , , recEmployee!lastname
                lst.ListSubItems.Add , , recEmployee!Salary
            recEmployee.MoveNext
        Loop
        Else
            MsgBox "There is no employee record in the database.", vbExclamation, title
            cmdUpdate.Enabled = False
    End If
                
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub

Private Sub list_DblClick()
On Error GoTo abdel
    recEmployee.Requery
    If Not recEmployee.EOF Then
        recEmployee.MoveFirst
        Do While Not recEmployee.EOF
            If Trim(recEmployee!firstname) = Trim(list.SelectedItem.Text) And _
                Trim(recEmployee!lastname) = Trim(list.SelectedItem.ListSubItems(1)) Then
                    Unload Me
                    Load frmUpdateSalary
                    frmUpdateSalary.Show
                    frmUpdateSalary.txtFirstName.Text = Trim(recEmployee!firstname)
                    frmUpdateSalary.txtLastName.Text = Trim(recEmployee!lastname)
                    frmUpdateSalary.txtSalary.Text = Trim(recEmployee!Salary)
                    Exit Sub
            End If
            recEmployee.MoveNext
        Loop
        Else
    End If
    Exit Sub
abdel:
    MsgBox "Sorry, salary data could not be updated", vbExclamation, title

End Sub
