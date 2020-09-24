VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEmptyDatabase 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   2610
      TabIndex        =   10
      Top             =   5370
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
      mIcon           =   "frmEmptyDatabase.frx":0000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DATABASE MANAGEMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale"
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
      Left            =   3240
      MouseIcon       =   "frmEmptyDatabase.frx":031A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Double Click"
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase"
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
      Left            =   360
      MouseIcon       =   "frmEmptyDatabase.frx":0624
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "Double Click"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
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
      Left            =   6000
      MouseIcon       =   "frmEmptyDatabase.frx":092E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Double Click"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Payslip"
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
      Left            =   3120
      MouseIcon       =   "frmEmptyDatabase.frx":0C38
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Double Click"
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Users Time record"
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
      Left            =   5400
      MouseIcon       =   "frmEmptyDatabase.frx":0F42
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Double Click"
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Orders"
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
      Left            =   360
      MouseIcon       =   "frmEmptyDatabase.frx":124C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Double Click"
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee"
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
      Left            =   6000
      MouseIcon       =   "frmEmptyDatabase.frx":1556
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Double Click"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
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
      Left            =   3120
      MouseIcon       =   "frmEmptyDatabase.frx":1860
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Double Click"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Left            =   360
      MouseIcon       =   "frmEmptyDatabase.frx":1B6A
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Double Click"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Image Image11 
      Height          =   1005
      Left            =   3240
      MouseIcon       =   "frmEmptyDatabase.frx":1E74
      MousePointer    =   99  'Custom
      Picture         =   "frmEmptyDatabase.frx":217E
      ToolTipText     =   "Double Click"
      Top             =   3720
      Width           =   1635
   End
   Begin VB.Image Image3 
      Height          =   1005
      Left            =   360
      MouseIcon       =   "frmEmptyDatabase.frx":359C
      MousePointer    =   99  'Custom
      Picture         =   "frmEmptyDatabase.frx":38A6
      ToolTipText     =   "Double Click"
      Top             =   3720
      Width           =   1635
   End
   Begin VB.Image Image10 
      Height          =   1005
      Left            =   6000
      MouseIcon       =   "frmEmptyDatabase.frx":4CC4
      MousePointer    =   99  'Custom
      Picture         =   "frmEmptyDatabase.frx":4FCE
      ToolTipText     =   "Double Click"
      Top             =   2040
      Width           =   1635
   End
   Begin VB.Image Image9 
      Height          =   1005
      Left            =   3120
      MouseIcon       =   "frmEmptyDatabase.frx":63EC
      MousePointer    =   99  'Custom
      Picture         =   "frmEmptyDatabase.frx":66F6
      ToolTipText     =   "Double Click"
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Image Image7 
      Height          =   1005
      Left            =   6000
      MouseIcon       =   "frmEmptyDatabase.frx":7B14
      MousePointer    =   99  'Custom
      Picture         =   "frmEmptyDatabase.frx":7E1E
      ToolTipText     =   "Double Click"
      Top             =   3720
      Width           =   1635
   End
   Begin VB.Image Image6 
      Height          =   1005
      Left            =   360
      MouseIcon       =   "frmEmptyDatabase.frx":923C
      MousePointer    =   99  'Custom
      Picture         =   "frmEmptyDatabase.frx":9546
      ToolTipText     =   "Double Click"
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Image Image5 
      Height          =   1005
      Left            =   6000
      MouseIcon       =   "frmEmptyDatabase.frx":A964
      MousePointer    =   99  'Custom
      Picture         =   "frmEmptyDatabase.frx":AC6E
      Top             =   600
      Width           =   1635
   End
   Begin VB.Image Image4 
      Height          =   1005
      Left            =   3120
      MouseIcon       =   "frmEmptyDatabase.frx":C08C
      MousePointer    =   99  'Custom
      Picture         =   "frmEmptyDatabase.frx":C396
      ToolTipText     =   "Double Click"
      Top             =   600
      Width           =   1635
   End
   Begin VB.Image Image1 
      Height          =   1005
      Left            =   360
      MouseIcon       =   "frmEmptyDatabase.frx":D7B4
      MousePointer    =   99  'Custom
      Picture         =   "frmEmptyDatabase.frx":DABE
      ToolTipText     =   "Double Click"
      Top             =   600
      Width           =   1635
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   7815
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5295
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmEmptyDatabase"
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
    Dim prompt As String
    
Private Sub cmdClose_Click()
    If MsgBox("Are you sure you want to close this window?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub Form_Load()
    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub

Private Sub Image1_DblClick()
    If MsgBox("Are you sure you want to empty the category table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from category"
            MsgBox "Transaction Successfull...", vbInformation, title
    End If
End Sub
Private Sub Image10_DblClick()

    If MsgBox("Are you sure you want to empty the product table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from product"
            MsgBox "Transaction Successfull...", vbInformation, title
    End If
End Sub
Private Sub Image11_DblClick()

    If MsgBox("Are you sure you want to empty the sale table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from sale"
            MsgBox "Transaction Successfull...", vbInformation, title
    End If
End Sub
Private Sub Image3_DblClick()

    If MsgBox("Are you sure you want to empty the purchase table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from purchase"
            MsgBox "Transaction Successfull...", vbInformation, title
    End If
End Sub
Private Sub Image4_DblClick()

    If MsgBox("Are you sure you want to empty the customer table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from Customer"
            MsgBox "Transaction Successfull...", vbInformation, title
    End If
End Sub
Private Sub Image5_DblClick()

    If MsgBox("Are you sure you want to empty the employee table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from employees"
            MsgBox "Transaction Successfull...", vbInformation, title
    End If
End Sub
Private Sub Image6_DblClick()

    If MsgBox("Are you sure you want to empty the order table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from orders"
            MsgBox "Transaction Successfull...", vbInformation, title
    End If
End Sub

Private Sub Image7_Click()

    If MsgBox("Are you sure you want to empty the users time record Table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from userslog"
            MsgBox "Transaction Successfull...", vbInformation, title
    End If
End Sub

Private Sub Image9_DblClick()

    If MsgBox("Are you sure you want to empty the payslip table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from payslip"
            MsgBox "Transaction Successfull...", vbInformation, title
    End If
End Sub
Private Sub Label1_DblClick()

    If MsgBox("Are you sure you want to empty the category table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from category"
            MsgBox "Transaction Successfull...", vbInformation, title
    End If
End Sub
Private Sub Label10_DblClick()

    If MsgBox("Are you sure you want to empty the purchase table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from purchase"
            MsgBox "Transaction Successfull...", vbInformation, title
    End If
End Sub
Private Sub Label11_DblClick()

    If MsgBox("Are you sure you want to empty the sale table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from Sale"
            MsgBox "Transaction Successfull...", vbInformation, title
    End If
End Sub
Private Sub Label3_DblClick()

    If MsgBox("Are you sure you want to empty the customer table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from Customer"
            MsgBox "Transaction successfull...", vbInformation, title
    End If
End Sub
Private Sub Label4_DblClick()

    If MsgBox("Are you sure you want to empty the Employee table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from employees"
            MsgBox "Transaction successfull...", vbInformation, title
    End If
End Sub
Private Sub Label5_DblClick()

    If MsgBox("Are you sure you want to empty the order table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from orders"
            MsgBox "Transaction successfull...", vbInformation, title
    End If
End Sub

Private Sub Label6_Click()

    If MsgBox("Are you sure you want to empty the users time record Table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from userslog"
            MsgBox "Transaction successfull...", vbInformation, title
    End If
End Sub

Private Sub Label8_DblClick()

    If MsgBox("Are you sure you want to empty the payslip table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from payslip"
            MsgBox "Transaction successfull...", vbInformation, title
    End If
End Sub
Private Sub Label9_DblClick()

    If MsgBox("Are you sure you want to empty the product table?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            con.Execute "delete * from product"
            MsgBox "Transaction successfull...", vbInformation, title
    End If
End Sub
