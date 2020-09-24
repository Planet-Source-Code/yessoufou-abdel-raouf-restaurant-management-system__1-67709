VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCustomer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
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
      Height          =   4830
      Left            =   360
      TabIndex        =   20
      Top             =   630
      Width           =   2415
   End
   Begin VB.TextBox txtFax 
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
      Left            =   4920
      MaxLength       =   20
      TabIndex        =   19
      ToolTipText     =   "Fax"
      Top             =   5310
      Width           =   2895
   End
   Begin VB.TextBox txtPhone 
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
      Left            =   4920
      MaxLength       =   20
      TabIndex        =   18
      ToolTipText     =   "Phone"
      Top             =   4830
      Width           =   2895
   End
   Begin VB.TextBox txtCountry 
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
      Left            =   4920
      MaxLength       =   20
      TabIndex        =   17
      ToolTipText     =   "Country"
      Top             =   4350
      Width           =   2895
   End
   Begin VB.TextBox txtPostalCode 
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
      Left            =   4920
      MaxLength       =   20
      TabIndex        =   16
      ToolTipText     =   "Postal Code"
      Top             =   3870
      Width           =   2895
   End
   Begin VB.TextBox txtCity 
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
      Left            =   4920
      MaxLength       =   20
      TabIndex        =   15
      ToolTipText     =   "City"
      Top             =   3390
      Width           =   2895
   End
   Begin VB.TextBox txtAddress 
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
      Height          =   615
      Left            =   4920
      TabIndex        =   14
      ToolTipText     =   "Address"
      Top             =   2670
      Width           =   2895
   End
   Begin VB.TextBox txtContactTitle 
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
      Left            =   4920
      MaxLength       =   20
      TabIndex        =   13
      ToolTipText     =   "His Title"
      Top             =   2070
      Width           =   2895
   End
   Begin VB.TextBox txtContactName 
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
      Left            =   4920
      MaxLength       =   20
      TabIndex        =   12
      ToolTipText     =   "His Contact Name"
      Top             =   1590
      Width           =   2895
   End
   Begin VB.TextBox txtCompanyName 
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
      Left            =   4920
      MaxLength       =   20
      TabIndex        =   11
      ToolTipText     =   "Company Name"
      Top             =   1110
      Width           =   2895
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   4920
      TabIndex        =   10
      ToolTipText     =   "Customer ID"
      Top             =   630
      Width           =   855
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   435
      Left            =   2220
      TabIndex        =   22
      Top             =   6240
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   767
      Caption         =   "Cancel"
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
      mIcon           =   "frmCustomer.frx":0000
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   435
      Left            =   1260
      TabIndex        =   23
      Top             =   6240
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   767
      Caption         =   "&Save"
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
      mIcon           =   "frmCustomer.frx":031A
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   435
      Left            =   3180
      TabIndex        =   24
      Top             =   6240
      Width           =   915
      _ExtentX        =   1614
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
      mIcon           =   "frmCustomer.frx":0634
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   435
      Left            =   330
      TabIndex        =   25
      Top             =   6240
      Width           =   915
      _ExtentX        =   1614
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
      mIcon           =   "frmCustomer.frx":094E
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   435
      Left            =   4140
      TabIndex        =   26
      Top             =   6240
      Width           =   915
      _ExtentX        =   1614
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
      mIcon           =   "frmCustomer.frx":0C68
   End
   Begin lvButton.lvButtons_H cmdSearch 
      Height          =   435
      Left            =   5100
      TabIndex        =   27
      Top             =   6240
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   767
      Caption         =   "Search"
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
      mIcon           =   "frmCustomer.frx":0F82
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   435
      Left            =   6060
      TabIndex        =   28
      Top             =   6240
      Width           =   915
      _ExtentX        =   1614
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
      mIcon           =   "frmCustomer.frx":129C
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   7050
      TabIndex        =   29
      Top             =   6240
      Width           =   915
      _ExtentX        =   1614
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
      mIcon           =   "frmCustomer.frx":15B6
   End
   Begin lvButton.lvButtons_H cmdNext 
      Height          =   435
      Left            =   1560
      TabIndex        =   30
      Top             =   5520
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   767
      Caption         =   ">"
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
      mIcon           =   "frmCustomer.frx":18D0
   End
   Begin lvButton.lvButtons_H cmdPrevious 
      Height          =   435
      Left            =   960
      TabIndex        =   31
      Top             =   5520
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   767
      Caption         =   "<"
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
      mIcon           =   "frmCustomer.frx":1BEA
   End
   Begin lvButton.lvButtons_H cmdLast 
      Height          =   435
      Left            =   2160
      TabIndex        =   32
      Top             =   5520
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   767
      Caption         =   ">|"
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
      mIcon           =   "frmCustomer.frx":1F04
   End
   Begin lvButton.lvButtons_H cmdFirst 
      Height          =   435
      Left            =   360
      TabIndex        =   33
      Top             =   5520
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   767
      Caption         =   "|<"
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
      mIcon           =   "frmCustomer.frx":221E
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   435
      Left            =   3180
      TabIndex        =   34
      Top             =   6240
      Width           =   915
      _ExtentX        =   1614
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
      mIcon           =   "frmCustomer.frx":2538
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMERS"
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
      Left            =   2760
      TabIndex        =   21
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   8175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   5310
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      Left            =   3000
      TabIndex        =   8
      Top             =   4830
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
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
      Left            =   3000
      TabIndex        =   7
      Top             =   4350
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Postal Code"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   3870
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Left            =   3000
      TabIndex        =   5
      Top             =   3390
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   3000
      TabIndex        =   4
      Top             =   2670
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Title"
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
      Left            =   3000
      TabIndex        =   3
      Top             =   2070
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   3000
      TabIndex        =   2
      Top             =   1590
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   1110
      Width           =   1935
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
      Left            =   3000
      TabIndex        =   0
      Top             =   630
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   6135
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmCustomer"
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
    Dim control As Object
    Dim position As Integer
    Dim sql As String
Private Sub cmdClose_Click()
    If cmdSave.Enabled = True Or cmdCancel.Enabled = True Then
        MsgBox "Please save or cancel the new customer details.", vbExclamation, title
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            If custSearch = True Then
                Load frmCustomerSearch
                frmCustomerSearch.Show
                Else
                
            End If
    End If
    Unload Me
End Sub

Private Sub cmdFirst_Click()

    If recCustomer.BOF And recCustomer.EOF Then
        MsgBox "There is no customer record in the database.", vbExclamation, title
        Exit Sub
    End If
    
    recCustomer.MoveFirst
    Call display
End Sub

Private Sub cmdLast_Click()

    If recCustomer.BOF And recCustomer.EOF Then
        MsgBox "There is no customer record in the database.", vbExclamation, title
        Exit Sub
    End If
    
    recCustomer.MoveLast
    Call display
End Sub

Private Sub cmdNew_Click()

    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see system administrator...", vbCritical, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to add new customer record ? ", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            'recCustomer.AddNew
            Call LockMe(False)
            Call EnableCmd
            txtID.Text = autogen
            txtCompanyName.Text = ""
            txtContactName.Text = ""
            txtContactTitle.Text = ""
            txtAddress.Text = ""
            txtCity.Text = ""
            txtPostalCode.Text = ""
            txtCountry.Text = ""
            txtPhone.Text = ""
            txtFax.Text = ""
    End If
End Sub

Private Sub cmdNext_Click()

    If recCustomer.BOF And recCustomer.EOF Then
        MsgBox "There is no customer record in the database.", vbExclamation, title
        Exit Sub
    End If
    
    recCustomer.MoveNext
    If recCustomer.EOF Then
        recCustomer.MoveLast
        MsgBox "This is the last customer record", vbExclamation, title
    End If
    Call display
End Sub

Private Sub cmdOk_Click()
On Error GoTo abdel
    If txtContactName.Text = "" Then
        MsgBox "The contact name field can not be left blank.", vbExclamation, title
        txtContactName.SetFocus
        Exit Sub
    End If
    
    If txtCity.Text = "" Then
        MsgBox "The city field can not be left blank.", vbExclamation, title
        txtCity.SetFocus
        Exit Sub
    End If
    If txtCountry.Text = "" Then
        MsgBox "The country field can not be left blank.", vbExclamation, title
        txtCountry.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to add the new customer record to the database?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            recCustomer!customerid = txtID.Text & ""
            recCustomer!CompanyName = txtCompanyName.Text & ""
            recCustomer!contactname = txtContactName.Text & ""
            recCustomer!ContactTitle = txtContactTitle.Text & ""
            recCustomer!Address = txtAddress.Text & ""
            recCustomer!city = txtCity.Text & ""
            recCustomer!PostalCode = txtPostalCode.Text & ""
            recCustomer!Country = txtCountry.Text & ""
            recCustomer!phone = txtPhone.Text & ""
            recCustomer!fax = txtFax.Text & ""
            recCustomer.UpdateBatch adAffectCurrent
            
            MsgBox "Customer record successfully saved.", vbInformation, title
            Call DisableCmd
            Call LockMe(True)
            list.Clear
            recCustomer.Requery
            recCustomer.MoveFirst
            Do While Not recCustomer.EOF
                list.AddItem recCustomer!contactname
                recCustomer.MoveNext
            Loop
            recCustomer.MoveFirst
            Call display
            cmdOk.Visible = False
            cmdUpdate.Visible = True
    End If
Exit Sub
abdel:
    MsgBox "Sorry, transactions not successfull", vbExclamation, title
End Sub

Private Sub cmdPrevious_Click()

    If recCustomer.BOF And recCustomer.EOF Then
        MsgBox "There is no record in the database.", vbExclamation, title
        Exit Sub
    End If
    
    recCustomer.MovePrevious
    If recCustomer.BOF Then
        recCustomer.MoveFirst
        MsgBox "This is the first customer record", vbExclamation, title
    End If
    Call display
End Sub

Private Sub CmdCancel_Click()

    If MsgBox("Are you sure you want to cancel the new record ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            'recCustomer.CancelUpdate
            Call LockMe(True)
            list.Clear
            recCustomer.Requery
            If Not recCustomer.EOF Then
                recCustomer.MoveFirst
                Do While Not recCustomer.EOF
                    list.AddItem recCustomer!contactname
                    recCustomer.MoveNext
                Loop
                recCustomer.MoveFirst
                Call display
                Call DisableCmd
                cmdOk.Visible = False
                cmdUpdate.Visible = True
                Else
                    For Each control In Me
                        If TypeOf control Is TextBox Then
                            control.Text = ""
                        End If
                    Next
                    Call disableAll
                    cmdNew.Enabled = True
                    cmdOk.Visible = False
                    cmdUpdate.Visible = True
            End If
    End If
    
    If blNewOrder = True Then
        Unload Me
        Load frmCustomerOrder
        frmCustomerOrder.Show
    End If
End Sub
Private Sub cmddelete_Click()
On Error GoTo abdel
    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see system administrator", vbCritical, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete the records of " & Trim(txtContactName.Text) & " ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            recCustomer.Delete adAffectCurrent
            list.Clear
            recCustomer.Requery
            If Not recCustomer.EOF Then
                recCustomer.MoveFirst
                Do While Not recCustomer.EOF
                    list.AddItem recCustomer!contactname
                    recCustomer.MoveNext
                Loop
                recCustomer.MoveFirst
                Call display
                Else
                    For Each control In Me
                        If TypeOf control Is TextBox Then
                            control.Text = ""
                        End If
                        Call disableAll
                    Next
            End If
    End If
Exit Sub
abdel:
    MsgBox "Sorry, custeomer records could not be deleted", vbExclamation, title
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo abdel
        sql = "select * from customer"
        Set CustomerReport.DataSource = con.Execute(sql)
        CustomerReport.Show
        Set CustomerReport = Nothing
    Exit Sub
abdel:
    MsgBox "Customer report not available for now", vbCritical, title
End Sub

Private Sub cmdSave_Click()
On Error GoTo abdel
    If txtContactName.Text = "" Then
        MsgBox "The contact name field can not be left blank.", vbExclamation, title
        txtContactName.SetFocus
        Exit Sub
    End If
    
    If txtCity.Text = "" Then
        MsgBox "The city field can not be left blank.", vbExclamation, title
        txtCity.SetFocus
        Exit Sub
    End If
    If txtCountry.Text = "" Then
        MsgBox "The country field can not be left blank.", vbExclamation, title
        txtCountry.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to add the new customer record to the database?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            recCustomer.AddNew
            recCustomer!customerid = txtID.Text & ""
            recCustomer!CompanyName = txtCompanyName.Text & ""
            recCustomer!contactname = txtContactName.Text & ""
            recCustomer!ContactTitle = txtContactTitle.Text & ""
            recCustomer!Address = txtAddress.Text & ""
            recCustomer!city = txtCity.Text & ""
            recCustomer!PostalCode = txtPostalCode.Text & ""
            recCustomer!Country = txtCountry.Text & ""
            recCustomer!phone = txtPhone.Text & ""
            recCustomer!fax = txtFax.Text & ""
            recCustomer.Update
            
            MsgBox "Customer records successfully saved.", vbInformation, title
            Call DisableCmd
            Call LockMe(True)
            list.Clear
            recCustomer.Requery
            recCustomer.MoveFirst
            Do While Not recCustomer.EOF
                list.AddItem recCustomer!contactname
                recCustomer.MoveNext
            Loop
            recCustomer.MoveFirst
            Call display
    End If
    If blNewOrder = True Then
        Unload Me
        Load frmCustomerOrder
        frmCustomerOrder.Show
    End If
Exit Sub
abdel:
    MsgBox "Sorry, transactions not successfull", vbExclamation, title
End Sub

Private Sub cmdSearch_Click()

    blSearchCustomer = True
    Load frmCustomerSearch
    frmCustomerSearch.Show
    Unload Me
End Sub

Private Sub cmdUpdate_Click()

    If frmScreen.lblRole.Caption <> "Administrator" Then
        MsgBox "Access denied, see system administrator", vbCritical, title
        Exit Sub
    End If
    If MsgBox("Are you sure you want to update the records of " & Trim(txtContactName.Text) & " ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Call LockMe(False)
            cmdUpdate.Visible = False
            cmdOk.Visible = True
            Call disableAll
            cmdNew.Enabled = False
            cmdCancel.Enabled = True
    End If
End Sub

Private Sub Form_Load()

    frmCustomer.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
    If blNewOrder = True Then
        Exit Sub
    End If
    recCustomer.Requery
    If recCustomer.BOF And recCustomer.EOF Then
        MsgBox "There is no customer record in the database!", vbExclamation, title
        Call LockMe(True)
        disableAll
        Exit Sub
    End If
    
    Call DisableCmd
    Call LockMe(True)
    list.Clear
    
    recCustomer.Requery
    recCustomer.MoveFirst
    Do While Not recCustomer.EOF
        list.AddItem recCustomer!contactname
        recCustomer.MoveNext
    Loop
    recCustomer.MoveFirst
    Call display
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blNewOrder = False
    frmCustomer.Enabled = True
End Sub

Private Sub list_Click()

    position = list.ListIndex
    recCustomer.MoveFirst
    recCustomer.Move position
    Call display
End Sub

Public Sub LockMe(LockUnlock As Boolean)

    For Each control In Me
        If TypeOf control Is TextBox Then
            control.Locked = LockUnlock
        End If
    Next
    
    For Each control In Me
        If TypeOf control Is ComboBox Then
            control.Locked = LockUnlock
        End If
    Next
End Sub
Public Function autogen()


    Dim recGen As New Recordset
    
    recGen.Open "select max(customerID) from customer", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogen = 1
        
        Else
        
        autogen = Val(recGen(0) + 1)
    End If
End Function

Public Sub EnableCmd()
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdNew.Enabled = False
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    cmdSearch.Enabled = False
    list.Enabled = False
End Sub

Public Sub DisableCmd()
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdNew.Enabled = True
    cmdUpdate.Enabled = True
    cmdPrint.Enabled = True
    cmdDelete.Enabled = True
    cmdSearch.Enabled = True
    list.Enabled = True
End Sub

Public Sub disableAll()
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdUpdate.Enabled = False
    cmdPrint.Enabled = False
    cmdDelete.Enabled = False
    cmdSearch.Enabled = False
End Sub

Public Sub display()

    txtID.Text = recCustomer!customerid & ""
    txtCompanyName.Text = recCustomer!CompanyName & ""
    txtContactName.Text = recCustomer!contactname & ""
    txtContactTitle.Text = recCustomer!ContactTitle & ""
    txtAddress.Text = recCustomer!Address & ""
    txtCity.Text = recCustomer!city & ""
    txtPostalCode.Text = recCustomer!PostalCode & ""
    txtCountry.Text = recCustomer!Country & ""
    txtPhone.Text = recCustomer!phone & ""
    txtFax.Text = recCustomer!fax & ""
End Sub
Private Sub txtcity_Validate(Cancel As Boolean)
    txtCity.Text = StrConv(txtCity.Text, vbProperCase)
End Sub
Private Sub txtCompanyname_Validate(Cancel As Boolean)
    txtCompanyName.Text = StrConv(txtCompanyName.Text, vbProperCase)
End Sub

Private Sub txtContactName_Validate(Cancel As Boolean)
    txtContactName.Text = StrConv(txtContactName.Text, vbProperCase)
End Sub
Private Sub txtContactTitle_Validate(Cancel As Boolean)
    txtContactTitle.Text = StrConv(txtContactTitle.Text, vbProperCase)
End Sub
Private Sub txtCountry_Validate(Cancel As Boolean)
    txtCountry.Text = StrConv(txtCountry.Text, vbProperCase)
End Sub
