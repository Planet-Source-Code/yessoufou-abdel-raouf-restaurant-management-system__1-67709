VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCustomerOrder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSearch 
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
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.ListBox list 
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
      Height          =   4350
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   4215
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   435
      Left            =   1800
      TabIndex        =   3
      Top             =   6000
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   767
      Caption         =   "&New Customer"
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
      mIcon           =   "frmCustomerOrder.frx":0000
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   1545
      _ExtentX        =   2725
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
      mIcon           =   "frmCustomerOrder.frx":031A
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   3600
      TabIndex        =   5
      Top             =   6000
      Width           =   1545
      _ExtentX        =   2725
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
      mIcon           =   "frmCustomerOrder.frx":0634
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   5895
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   5295
   End
End
Attribute VB_Name = "frmCustomerOrder"
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
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Const LB_FINDSTRING = &H18F
Private Sub cmdClose_Click()
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
            Load Order
            Order.Show
    End If
End Sub

Private Sub cmdNew_Click()

    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see system administrator", vbCritical, title
        Exit Sub
    End If
    
    blNewOrder = True
        Unload Me
        Load frmCustomer
        frmCustomer.Show
        'recCustomer.AddNew
        Call frmCustomer.LockMe(False)
        
        frmCustomer.txtID.Text = frmCustomer.autogen
        frmCustomer.txtCompanyName.Text = ""
        frmCustomer.txtContactName.Text = ""
        frmCustomer.txtContactTitle.Text = ""
        frmCustomer.txtAddress.Text = ""
        frmCustomer.txtCity.Text = ""
        frmCustomer.txtPostalCode.Text = ""
        frmCustomer.txtCountry.Text = ""
        frmCustomer.txtPhone.Text = ""
        frmCustomer.txtFax.Text = ""
        Call frmCustomer.EnableCmd
End Sub

Private Sub cmdNewCustomer_Click()

End Sub

Private Sub cmdOk_Click()

    recCustomer.MoveFirst
    recCustomer.Move list.ListIndex
    frmOrder.Show
    frmOrder.txtName.Text = recCustomer!contactname
    frmOrder.txtID.Text = recCustomer!customerid
    Unload Me
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    
    If recCustomer.BOF And recCustomer.EOF Then
        MsgBox "There is no customer in the dataBase, kindly add some.", vbExclamation, title
        cmdOk.Enabled = False
        Exit Sub
    End If
        
    recCustomer.Requery
    recCustomer.MoveFirst
    Do While Not recCustomer.EOF
        list.AddItem recCustomer!contactname
        recCustomer.MoveNext
    Loop
    recCustomer.MoveFirst
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blOrder = False
    frmScreen.Enabled = True
End Sub

Private Sub list_DblClick()

    recCustomer.MoveFirst
    recCustomer.Move list.ListIndex
    frmOrder.Show
    frmOrder.txtName.Text = recCustomer!contactname
    frmOrder.txtID.Text = recCustomer!customerid
    Unload Me
End Sub

Private Sub TxtSearch_Change()

If txtSearch <> "" Then
    A = SendMessage(list.hwnd, LB_FINDSTRING, -1, ByVal txtSearch.Text)
    If A >= 0 Then list.ListIndex = A
End If
End Sub
