VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPaySlip 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   4350
      TabIndex        =   19
      Top             =   4890
      Width           =   1935
      _ExtentX        =   3413
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
      mIcon           =   "frmPaySlip.frx":0000
   End
   Begin lvButton.lvButtons_H cmdRegister 
      Height          =   435
      Left            =   510
      TabIndex        =   20
      Top             =   4890
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      Caption         =   "&Print"
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
      mIcon           =   "frmPaySlip.frx":031A
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   360
   End
   Begin VB.TextBox txtNo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   17
      Top             =   4890
      Width           =   615
   End
   Begin VB.TextBox txtNetIncome 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtDeduction 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4800
      MaxLength       =   7
      TabIndex        =   5
      ToolTipText     =   "Deduction"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Time"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Date"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtGrossPay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Gross Pay"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtBonus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   4
      ToolTipText     =   "Bonus"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "txtRole"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox com 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Choisisser le nom de l'employe en question"
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   6735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAYSLIP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1080
      TabIndex        =   16
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label8 
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
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label7 
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
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Income"
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
      Left            =   1920
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction"
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
      Left            =   3600
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonus"
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
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Pay"
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
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404000&
      Height          =   4815
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmPaySlip"
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
    Dim currentLength As Byte
    Const msg As String = "LES CASSEROLES DU SHERIF"
    Dim reply As Integer
    Dim position As Integer
    Dim salBrut As Double
    Dim prim As Double
    Dim penalit As Double
    Dim salNet As Double
    Dim prompt As String
    Dim prompt1 As String
    Dim sql As String

Private Sub cmdClose_Click()

    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub cmdRegister_Click()
On Error GoTo abdel
    If com.Text = "" Then
        MsgBox "Please select the employee name", vbExclamation, title
        com.SetFocus
        Exit Sub
    End If
    
    prompt = "Please, check wether everything is in order."
    prompt = prompt & Chr(10) & Chr(13) & " Click 'Yes' To Register The Details"
    If MsgBox(prompt, 4 + vbQuestion, title) = vbNo Then
        Exit Sub
        Else
            'MsgBox prompt, vbYesNo + vbQuestion, title
            If txtBonus.Text = "" Then
                txtBonus.Text = 0
            End If
            If txtDeduction.Text = "" Then
                txtDeduction.Text = 0
            End If
            recPayslip.AddNew
            recPayslip!payslipid = txtNo.Text & ""
            recPayslip!employeeid = txtID.Text & ""
            recPayslip!employeename = com.Text & ""
            recPayslip!title = txtTitle.Text & ""
            recPayslip!GrossPay = txtGrossPay.Text & ""
            recPayslip!bonus = txtBonus.Text
            recPayslip!deduction = txtDeduction.Text
            recPayslip!netincome = txtNetIncome.Text & ""
            recPayslip!dates = txtDate.Text & ""
            recPayslip!Time = txtTime.Text & ""
            recPayslip.Update
            
            sql = "select * from payslip where payslipid = " & Trim(txtNo.Text)
            Set PayslipBill.DataSource = con.Execute(sql)
            PayslipBill.Show
            Set PayslipBill = Nothing

        End If
        Unload Me
    Exit Sub
abdel:
    MsgBox "Payslip not available for now...", vbExclamation, title

End Sub

Private Sub com_Click()

    position = com.ListIndex
    recEmployee.MoveFirst
    recEmployee.Move position
    Call display
    txtNetIncome.Text = Trim(recEmployee!Salary)
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    txtNo.Text = autogen
    txtNetIncome.Text = txtGrossPay.Text
    com.Clear
    recEmployee.Requery
    If Not recEmployee.EOF Then
        recEmployee.MoveFirst
        Do While Not recEmployee.EOF
            com.AddItem recEmployee!firstname & "  " & recEmployee!lastname
            recEmployee.MoveNext
        Loop
        recEmployee.MoveFirst
    End If
    
End Sub
Public Function autogen()

    Dim recGen As New Recordset
    
    recGen.Open "select max(PaySlipID) from PaySlip", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogen = 1
        
        Else
        
        autogen = Val(recGen(0) + 1)
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub

Private Sub Timer2_Timer()

    txtDate.Text = Date
    txtTime.Text = Time
End Sub

Public Sub display()
On Error Resume Next
    txtID.Text = recEmployee!employeeid & ""
    txtTitle.Text = recEmployee!title & ""
    txtGrossPay.Text = recEmployee!Salary & ""

End Sub

Private Sub txtDeduction_Change()

    txtNetIncome.Text = Val(txtGrossPay.Text) - Val(txtDeduction.Text) + Val(txtBonus.Text)
End Sub

Private Sub txtDeduction_KeyPress(KeyAscii As Integer)

 Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack, vbKeyDelete
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtbonus_Change()

    txtNetIncome.Text = Val(txtGrossPay.Text) + Val(txtBonus.Text)
End Sub

Private Sub txtbonus_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack, vbKeyDelete
        Case Else
            KeyAscii = 0
    End Select
End Sub

Public Sub clearThem()

    txtTitle.Text = ""
    txtGrossPay.Text = ""
    txtBonus.Text = ""
    txtDeduction.Text = ""
    txtNetIncome.Text = ""
End Sub
