VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesReport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker Picker2 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   47120385
      CurrentDate     =   38665
   End
   Begin MSComCtl2.DTPicker Picker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   47120385
      CurrentDate     =   38665
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   2460
      TabIndex        =   5
      Top             =   2370
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
      mIcon           =   "frmSalesReport.frx":0000
   End
   Begin lvButton.lvButtons_H cmdView 
      Height          =   435
      Left            =   270
      TabIndex        =   6
      Top             =   2370
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   767
      Caption         =   "&View"
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
      mIcon           =   "frmSalesReport.frx":031A
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2295
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmSalesReport"
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

Private Sub cmdView_Click()
On Error GoTo abdel
    If Picker1.Value > Picker2.Value Then
        MsgBox "The first date should not be less than the second one.", vbExclamation, title
        Exit Sub
    End If
    If blSaleRep = True And blPurchaseRep = False Then
        sql = "select * from sale where date between #" & Picker1.Value & "# and #" & Picker2.Value & "#"
        SalesRep.Sections("section2").Controls.Item("lbldate").Caption = "From " & Picker1.Value & " To " & Picker2.Value
        Set SalesRep.DataSource = con.Execute(sql)
        SalesRep.Show
        Set SalesRep = Nothing
    End If
    
    If blPurchaseRep = True And blSaleRep = False Then
        sql = "select * from purchase where date between #" & Picker1.Value & "# and #" & Picker2.Value & "#"
        PurchaseRep.Sections("section2").Controls.Item("lbldate").Caption = "From " & Picker1.Value & " To " & Picker2.Value
        Set PurchaseRep.DataSource = con.Execute(sql)
        PurchaseRep.Show
        Set PurchaseRep = Nothing
    End If
    Exit Sub
abdel:
    MsgBox "Report not available for now...", vbExclamation, title
End Sub

Private Sub Form_Load()
    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    Picker1.Value = Date
    Picker2.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blSaleRep = False
    blPurchaseRep = False
    frmScreen.Enabled = True
End Sub
