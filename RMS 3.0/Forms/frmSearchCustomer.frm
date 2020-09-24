VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSearchCustomer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton op2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact Name"
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
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   2055
   End
   Begin VB.OptionButton op1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer ID"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtName 
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
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Contact Name"
      Top             =   2760
      Width           =   2415
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
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "Customer ID"
      Top             =   1320
      Width           =   1695
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   2550
      TabIndex        =   7
      Top             =   3690
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
      mIcon           =   "frmSearchCustomer.frx":0000
   End
   Begin lvButton.lvButtons_H cmdSearch 
      Height          =   435
      Left            =   360
      TabIndex        =   8
      Top             =   3690
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
      mIcon           =   "frmSearchCustomer.frx":031A
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
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   0
      Width           =   3015
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   4815
   End
   Begin VB.Shape Shape3 
      Height          =   1215
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Name"
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
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
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
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   3615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmSearchCustomer"
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
            If blSearchCustomer = True Then
                Unload Me
                Load Customer
                Customer.Show
                ElseIf blSearchCustomer = False Then
                    Unload Me
            End If
    End If
End Sub

Private Sub cmdSearch_Click()
On Error GoTo abdel
    counter = 0
    If op1.Value = True And op2.Value = False Then
        If txtID.Text = "" Then
            MsgBox "The id can not not be left blank.", vbExclamation, title
            txtID.SetFocus
            Exit Sub
        End If
        
        If recCustomer.BOF And recCustomer.EOF Then
            MsgBox "Sorry, there is no customer record in the database.", vbExclamation, title
            Exit Sub
        End If
    
        recCustomer.Requery
        recCustomer.MoveFirst
        Do While Not recCustomer.EOF
            If Trim(recCustomer!customerid) = Trim(txtID.Text) Then
                Unload Me
                Load frmCustomer
                frmCustomer.Show
                frmCustomer.list.ListIndex = counter
                Exit Sub
            End If
            counter = counter + 1
            recCustomer.MoveNext
        Loop
    
        If recCustomer.EOF Then
            MsgBox "There is no customer with ID : " & Trim(txtID.Text), vbExclamation, title
            txtID.SelStart = 0
            txtID.SelLength = Len(txtID.Text)
            txtID.SetFocus
            Exit Sub
        End If
    End If
    
    If op2.Value = True And op1.Value = False Then
        If txtName.Text = "" Then
            MsgBox "The contact name field can not be left blank", vbExclamation, title
            txtName.SetFocus
            Exit Sub
        End If
        If recCustomer.BOF And recCustomer.EOF Then
            MsgBox "Sorry, there is no customer record in the database.", vbExclamation, title
            op1.Value = False
            op2.Value = False
            txtID.Text = ""
            txtName.Text = ""
            Exit Sub
        End If
        recCustomer.Requery
        recCustomer.MoveFirst
        Do While Not recCustomer.EOF
            If Trim(recCustomer!contactname) = Trim(txtName.Text) Then
                Unload Me
                Load frmCustomer
                frmCustomer.Show
                frmCustomer.list.ListIndex = counter
                Exit Sub
            End If
            counter = counter + 1
            recCustomer.MoveNext
        Loop
    
        If recCustomer.EOF Then
            MsgBox "There is no customer with name : " & Trim(txtName.Text), vbExclamation, title
            txtName.SelStart = 0
            txtName.SelLength = Len(txtName.Text)
            txtName.SetFocus
            Exit Sub
        End If
    End If
    Exit Sub
abdel:
    MsgBox "Searching failed", vbExclamation, title

End Sub

Private Sub Form_Load()
    frmScreen.Enabled = False
    Me.Top = 1800
    Move (Screen.Width - Width) / 2
    Call ConnectMe
    
    op1.Value = False
    op2.Value = False
    txtID.Enabled = False
    txtName.Enabled = False
    cmdSearch.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blSearchCustomer = False
    frmScreen.Enabled = True
End Sub

Private Sub op1_Click()
On Error Resume Next
    txtID.Enabled = True
    txtID.SetFocus
    
    op2.Value = False
    txtName.Enabled = False
    txtName.Text = ""
    cmdSearch.Enabled = True
End Sub

Private Sub op2_Click()
    txtName.Enabled = True
    txtName.SetFocus
    
    op1.Value = False
    txtID.Enabled = False
    txtID.Text = ""
    cmdSearch.Enabled = True
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    txtName.Text = StrConv(txtName.Text, vbProperCase)
End Sub
