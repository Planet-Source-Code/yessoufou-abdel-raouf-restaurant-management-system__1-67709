VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSearchEmployee 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton op1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EMPLOYEE ID"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   1935
   End
   Begin VB.OptionButton op2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EMPLOYEE FIRST AND LAST NAME"
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
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox txtLastName 
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
      Height          =   360
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   4
      ToolTipText     =   "Last Name"
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox txtFirstName 
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
      Height          =   360
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   2
      ToolTipText     =   "First Name"
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox txtID 
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
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   0
      ToolTipText     =   "EmployeeID"
      Top             =   1320
      Width           =   975
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   2790
      TabIndex        =   8
      Top             =   3810
      Width           =   1755
      _ExtentX        =   3096
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
      mIcon           =   "frmSearchEmployee.frx":0000
   End
   Begin lvButton.lvButtons_H cmdSearch 
      Height          =   435
      Left            =   600
      TabIndex        =   9
      Top             =   3810
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   767
      Caption         =   "&Search"
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
      mIcon           =   "frmSearchEmployee.frx":031A
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH"
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
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   5295
   End
   Begin VB.Shape Shape3 
      Height          =   1695
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
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
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
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
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   3735
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmSearchEmployee"
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
    Dim counter As Integer
Private Sub cmdClose_Click()
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            If blSearchEmployee = True Then
                Unload Me
                Load Employees
                Employees.Show
                ElseIf blSearchEmployee = False Then
                    Unload Me
            End If
    End If
End Sub

Private Sub cmdSearch_Click()
On Error GoTo abdel
    counter = 0
    If op1.Value = True And op2.Value = False Then
        If txtID.Text = "" Then
            MsgBox "The employee Id field can not be left blank.", vbExclamation, title
            txtID.SetFocus
            Exit Sub
        End If
        
        If recEmployee.BOF And recEmployee.EOF Then
            MsgBox "Sorry, there is no record in the database.", vbExclamation, title
            op1.Value = False
            op2.Value = False
            txtID.Text = ""
            txtFirstName.Text = ""
            txtLastName.Text = ""
            Exit Sub
        End If
    
        recEmployee.Requery
        recEmployee.MoveFirst
        Do While Not recEmployee.EOF
            If Trim(recEmployee!employeeid) = Trim(txtID.Text) Then
                Unload Me
                Load frmEmployees
                frmEmployees.Show
                frmEmployees.list.ListIndex = counter
                Exit Sub
            End If
            counter = counter + 1
            recEmployee.MoveNext
        Loop
    
        If recEmployee.EOF Then
            MsgBox "There is no employee with ID : " & Trim(txtID.Text), vbExclamation, title
            txtID.SelStart = 0
            txtID.SelLength = Len(txtID.Text)
            txtID.SetFocus
            Exit Sub
        End If
    End If
    
    If op2.Value = True And op1.Value = False Then
        If txtFirstName.Text = "" Then
            MsgBox "The first name field can not be left blank.", vbExclamation, title
            txtFirstName.SetFocus
            Exit Sub
        End If
        If txtLastName.Text = "" Then
            MsgBox "The last name field can not be left blank.", vbExclamation, title
            txtLastName.SetFocus
            Exit Sub
        End If
        If recEmployee.BOF And recEmployee.EOF Then
            MsgBox "Sorry, there is no record in the database.", vbExclamation, title
            Exit Sub
        End If
        recEmployee.Requery
        recEmployee.MoveFirst
        Do While Not recEmployee.EOF
            If Trim(recEmployee!firstname) = Trim(txtFirstName.Text) And _
              Trim(recEmployee!lastname) = Trim(txtLastName.Text) Then
                Unload Me
                Load frmEmployees
                frmEmployees.Show
                frmEmployees.list.ListIndex = counter
                Exit Sub
            End If
            counter = counter + 1
            recEmployee.MoveNext
        Loop
    
        If recEmployee.EOF Then
            MsgBox "There is no employee with name : " & Trim(txtFirstName.Text) & " " & Trim(txtLastName.Text), vbExclamation, title
            txtFirstName.SelStart = 0
            txtFirstName.SelLength = Len(txtFirstName.Text)
            txtFirstName.SetFocus
            Exit Sub
        End If
    End If
    Exit Sub
abdel:
    MsgBox "Searching failed", vbExclamation, title

End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    op1.Value = False
    txtID.Enabled = False
    op2.Value = False
    txtFirstName.Enabled = False
    txtLastName.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blSearchEmployee = False
    frmScreen.Enabled = True
End Sub

Private Sub op1_Click()

    txtID.Enabled = True
    txtID.SetFocus
    op2.Value = False
    txtFirstName.Enabled = False
    txtFirstName.Text = ""
    txtLastName.Enabled = False
    txtLastName.Text = ""
    cmdSearch.Enabled = True
End Sub

Private Sub op2_Click()
    txtFirstName.Enabled = True
    txtFirstName.SetFocus
    txtLastName.Enabled = True
    op1.Value = False
    txtID.Enabled = False
    txtID.Text = ""
    cmdSearch.Enabled = True
End Sub

Private Sub txtFirstName_Validate(Cancel As Boolean)
    txtFirstName.Text = StrConv(txtFirstName.Text, vbProperCase)
End Sub


Private Sub txtID_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack, vbKey0 To vbKey9, vbKeyDelete
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtLastName_Validate(Cancel As Boolean)
    txtLastName.Text = StrConv(txtLastName.Text, vbProperCase)
End Sub
