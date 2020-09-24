VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lv 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double click to view details"
      Top             =   720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5318
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Customer Name"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Order"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amount Due"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount Paid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Amount Owed"
         Object.Width           =   2540
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   6570
      TabIndex        =   2
      Top             =   4050
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
      mIcon           =   "frmCredit.frx":0000
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   435
      Left            =   2220
      TabIndex        =   3
      Top             =   4050
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
      mIcon           =   "frmCredit.frx":031A
   End
   Begin lvButton.lvButtons_H cmdOrderDetails 
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   4050
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      Caption         =   "&View Details"
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
      mIcon           =   "frmCredit.frx":0634
   End
   Begin lvButton.lvButtons_H cmdPayment 
      Height          =   435
      Left            =   4380
      TabIndex        =   5
      Top             =   4050
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      Caption         =   "&Payment"
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
      mIcon           =   "frmCredit.frx":094E
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "THE LIST OF ALL CUSTOMER(S) WHO ARE OWING US"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   6495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   3975
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "frmCredit"
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
    Dim sql As String
Private Sub cmdClose_Click()
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Unload Me
    End If
End Sub

Private Sub cmdOrderDetails_Click()

    blOrderDetails1 = True
    recOrderDetails.Requery
    recOrderDetails.MoveFirst
    Do While Not recOrderDetails.EOF
        If recOrderDetails!orderid = lv.SelectedItem.ListSubItems(1).Text Then
            Set lstItem = frmOrderDetails.lv.ListItems.Add(, , recOrderDetails!orderDetailsID)
                lstItem.ListSubItems.Add , , recOrderDetails!ProductName
                lstItem.ListSubItems.Add , , recOrderDetails!unitprice
                lstItem.ListSubItems.Add , , recOrderDetails!quantity
                lstItem.ListSubItems.Add , , recOrderDetails!discount
                lstItem.ListSubItems.Add , , recOrderDetails!total
                frmOrderDetails.lblNo.Caption = lv.SelectedItem.ListSubItems(1).Text
        End If
        recOrderDetails.MoveNext
    Loop
    frmOrderDetails.Show
    Unload Me
    
End Sub

Private Sub cmdPayment_Click()

    frmCreditPayment.lblName.Caption = lv.SelectedItem.Text
    frmCreditPayment.txtID.Text = lv.SelectedItem.ListSubItems(1).Text
    frmCreditPayment.txtAmountDue = lv.SelectedItem.ListSubItems(4).Text
    frmCreditPayment.Show
    Unload Me
    
End Sub


Private Sub cmdPrint_Click()
    On Error GoTo abdel
    sql = "select * from credit"
    Set CreditPrint.DataSource = con.Execute(sql)
    CreditPrint.Show
    Set CreditPrint = Nothing
    Exit Sub
abdel:
    MsgBox "Report not available for now...", vbExclamation, title

End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    
    recCredit.Requery
    If Not recCredit.EOF Then
        recCredit.MoveFirst
        Do While Not recCredit.EOF
            Set lstItem = lv.ListItems.Add(, , recCredit!customername)
                lstItem.ListSubItems.Add , , recCredit!orderid
                lstItem.ListSubItems.Add , , recCredit!amountdue
                lstItem.ListSubItems.Add , , recCredit!amountpaid
                lstItem.ListSubItems.Add , , recCredit!AmountOwed
            recCredit.MoveNext
        Loop
        Else
            Label1.Caption = "Nobody Is Owing The Company"
            cmdPayment.Enabled = False
            cmdOrderDetails.Enabled = False
            cmdPrint.Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub
Private Sub lv_DblClick()

    blOrderDetails1 = True
    recOrderDetails.Requery
    recOrderDetails.MoveFirst
    Do While Not recOrderDetails.EOF
        If recOrderDetails!orderid = lv.SelectedItem.ListSubItems(1).Text Then
            Set lstItem = frmOrderDetails.lv.ListItems.Add(, , recOrderDetails!orderDetailsID)
                lstItem.ListSubItems.Add , , recOrderDetails!ProductName
                lstItem.ListSubItems.Add , , recOrderDetails!unitprice
                lstItem.ListSubItems.Add , , recOrderDetails!quantity
                lstItem.ListSubItems.Add , , recOrderDetails!discount
                lstItem.ListSubItems.Add , , recOrderDetails!total
                frmOrderDetails.lblNo.Caption = lv.SelectedItem.ListSubItems(1).Text
        End If
        recOrderDetails.MoveNext
    Loop
    Unload Me
End Sub
