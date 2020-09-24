VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomerSearch 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   LinkTopic       =   "Form2"
   ScaleHeight     =   6885
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lv 
      Height          =   4455
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Double-click to view record"
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   7858
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.TextBox txtSearch 
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
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Enter the first character"
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox com 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCustomerSearch.frx":0000
      Left            =   360
      List            =   "frmCustomerSearch.frx":000D
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   540
      TabIndex        =   5
      Top             =   6330
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
      mIcon           =   "frmCustomerSearch.frx":003C
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000007&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6255
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmCustomerSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rec1 As New Recordset
Dim rec2 As New Recordset
Dim rec3 As New Recordset

Dim rec4 As New Recordset
Dim rec5 As New Recordset
Dim rec6 As New Recordset

Dim lst As ListItem
Dim sql As String
Dim ctr As Integer
Private Sub cmdClose_Click()

    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            If custSearch = True Then
                Unload Me
                Else
                    Load Customer
                    Customer.Show
                    Unload Me
            End If
    End If
    
End Sub

Private Sub com_Click()

    If com.Text = "Customer ID" Then
        lv.ListItems.Clear
        rec1.Open "select customerid from customer order by customerid", con, adOpenDynamic, adLockOptimistic
        rec1.Requery
        rec1.MoveFirst
        Do While Not rec1.EOF
            Set lst = lv.ListItems.Add(, , rec1!customerid)
            rec1.MoveNext
        Loop
        Set rec1 = Nothing
    End If
    
    If com.Text = "Customer City" Then
        lv.ListItems.Clear
        rec2.Open "select city from customer order by city", con, adOpenDynamic, adLockOptimistic
        rec2.Requery
        rec2.MoveFirst
        Do While Not rec2.EOF
            Set lst = lv.ListItems.Add(, , rec2!city)
            rec2.MoveNext
        Loop
        Set rec2 = Nothing
    End If
    
    If com.Text = "Customer Name" Then
        lv.ListItems.Clear
        rec3.Open "select contactname from customer order by contactname", con, adOpenDynamic, adLockOptimistic
        rec3.Requery
        rec3.MoveFirst
        Do While Not rec3.EOF
            Set lst = lv.ListItems.Add(, , rec3!contactname)
            rec3.MoveNext
        Loop
        Set rec3 = Nothing
    End If
    txtSearch.Text = ""
    txtSearch.SetFocus
End Sub

Private Sub com_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    com.Text = "Customer Name"
    lbl.Caption = com.Text
    
    recCustomer.Requery
    recCustomer.MoveFirst
    Do While Not recCustomer.EOF
        Set lst = lv.ListItems.Add(, , recCustomer!contactname)
        recCustomer.MoveNext
    Loop
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
    cust = False
End Sub
Private Sub lv_DblClick()
On Error GoTo abdel
    If lv.ListItems.Count = 0 Then
        MsgBox "There is no record selected or record does not exist", vbExclamation, title
        Exit Sub
    End If
    
    If com.Text = "Customer ID" Then
        ctr = 0
        recCustomer.Requery
        recCustomer.MoveFirst
        Do While Not recCustomer.EOF
            If Trim(recCustomer!customerid) = lv.SelectedItem.Text Then
                Unload Me
                Load frmCustomer
                frmCustomer.Show
                frmCustomer.list.ListIndex = ctr
                Exit Sub
            End If
            ctr = ctr + 1
            recCustomer.MoveNext
        Loop
    
    ElseIf com.Text = "Customer City" Then
        ctr = 0
        recCustomer.Requery
        recCustomer.MoveFirst
        Do While Not recCustomer.EOF
            If Trim(recCustomer!city) = lv.SelectedItem.Text Then
                Unload Me
                Load frmCustomer
                frmCustomer.Show
                frmCustomer.list.ListIndex = ctr
                Exit Sub
            End If
            ctr = ctr + 1
            recCustomer.MoveNext
        Loop
        
    ElseIf com.Text = "Customer Name" Then
        ctr = 0
        recCustomer.Requery
        recCustomer.MoveFirst
        Do While Not recCustomer.EOF
            If Trim(recCustomer!contactname) = lv.SelectedItem.Text Then
                Unload Me
                Load frmCustomer
                frmCustomer.Show
                frmCustomer.list.ListIndex = ctr
                Exit Sub
            End If
            ctr = ctr + 1
            recCustomer.MoveNext
        Loop
    End If
    Exit Sub
abdel:
    MsgBox "Sorry! An error occured...", vbExclamation, title
End Sub

Private Sub TxtSearch_Change()
On Error GoTo abdel
    If com.Text = "Customer ID" Then
        sql = "select customerid from customer where customerid like '" & Trim(txtSearch.Text) & "%'"
        rec4.Open sql, con, adOpenDynamic, adLockOptimistic
        lv.ListItems.Clear
        Do While Not rec4.EOF
            Set lst = lv.ListItems.Add(, , rec4!customerid)
            rec4.MoveNext
        Loop
        Set rec4 = Nothing
    End If

    If com.Text = "Customer City" Then
        sql = "select city from customer where city like '" & Trim(txtSearch.Text) & "%'"
        rec5.Open sql, con, adOpenDynamic, adLockOptimistic
        lv.ListItems.Clear
        Do While Not rec5.EOF
            Set lst = lv.ListItems.Add(, , rec5!city)
            rec5.MoveNext
        Loop
        Set rec5 = Nothing
    End If
    
    If com.Text = "Customer Name" Then
        sql = "select contactname from customer where contactname like '" & Trim(txtSearch.Text) & "%'"
        rec6.Open sql, con, adOpenDynamic, adLockOptimistic
        lv.ListItems.Clear
        Do While Not rec6.EOF
            Set lst = lv.ListItems.Add(, , rec6!contactname)
            rec6.MoveNext
        Loop
        Set rec6 = Nothing
    End If
Exit Sub
abdel:
    MsgBox "Invalid input. Kindly retry", vbExclamation, title
    Unload Me
    Load frmCustomer
    frmCustomer.Show
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)

    If com.Text = "Customer ID" Then
    Select Case KeyAscii
        Case vbKeyDelete, vbKeyBack, vbKey0 To vbKey9
        Case Else
            KeyAscii = 0
    End Select
    End If
End Sub

