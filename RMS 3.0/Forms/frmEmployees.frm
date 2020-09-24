VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmployees 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   10095
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
      Height          =   4110
      Left            =   120
      TabIndex        =   30
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtSalary 
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
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   5
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox txtNotes 
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
      Height          =   855
      Left            =   7440
      TabIndex        =   11
      ToolTipText     =   "Education Background"
      Top             =   4320
      Width           =   2415
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
      Left            =   7440
      TabIndex        =   10
      ToolTipText     =   "Phone"
      Top             =   3600
      Width           =   2415
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
      Left            =   7440
      TabIndex        =   9
      ToolTipText     =   "Country"
      Top             =   3000
      Width           =   2415
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
      Left            =   7440
      TabIndex        =   8
      ToolTipText     =   "Postal Code"
      Top             =   2400
      Width           =   2415
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
      Left            =   7440
      TabIndex        =   7
      ToolTipText     =   "City"
      Top             =   1800
      Width           =   2415
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
      Height          =   735
      Left            =   7440
      TabIndex        =   6
      ToolTipText     =   "Address"
      Top             =   840
      Width           =   2415
   End
   Begin VB.ComboBox comboGender 
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
      ItemData        =   "frmEmployees.frx":0000
      Left            =   3720
      List            =   "frmEmployees.frx":000A
      TabIndex        =   4
      ToolTipText     =   "Gender"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtTitle 
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
      Left            =   3720
      MaxLength       =   25
      TabIndex        =   3
      ToolTipText     =   "Title"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtLastName 
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
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Last Name"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtFirstName 
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
      Left            =   3720
      TabIndex        =   1
      ToolTipText     =   "First Name"
      Top             =   1320
      Width           =   2055
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
      Left            =   3720
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin MSComCtl2.DTPicker HirePicker 
      Height          =   375
      Left            =   3720
      TabIndex        =   28
      Top             =   4320
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   46661633
      CurrentDate     =   38677
   End
   Begin MSComCtl2.DTPicker BirthPicker 
      Height          =   375
      Left            =   3720
      TabIndex        =   29
      Top             =   3600
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   46661633
      CurrentDate     =   38677
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   435
      Left            =   2610
      TabIndex        =   31
      Top             =   6090
      Width           =   1095
      _ExtentX        =   1931
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
      mIcon           =   "frmEmployees.frx":0014
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   435
      Left            =   1380
      TabIndex        =   32
      Top             =   6090
      Width           =   1095
      _ExtentX        =   1931
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
      mIcon           =   "frmEmployees.frx":032E
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   435
      Left            =   3840
      TabIndex        =   33
      Top             =   6090
      Width           =   1095
      _ExtentX        =   1931
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
      mIcon           =   "frmEmployees.frx":0648
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   435
      Left            =   120
      TabIndex        =   34
      Top             =   6090
      Width           =   1095
      _ExtentX        =   1931
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
      mIcon           =   "frmEmployees.frx":0962
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   435
      Left            =   5070
      TabIndex        =   35
      Top             =   6090
      Width           =   1095
      _ExtentX        =   1931
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
      mIcon           =   "frmEmployees.frx":0C7C
   End
   Begin lvButton.lvButtons_H cmdSearch 
      Height          =   435
      Left            =   6330
      TabIndex        =   36
      Top             =   6090
      Width           =   1095
      _ExtentX        =   1931
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
      mIcon           =   "frmEmployees.frx":0F96
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   435
      Left            =   7530
      TabIndex        =   37
      Top             =   6090
      Width           =   1095
      _ExtentX        =   1931
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
      mIcon           =   "frmEmployees.frx":12B0
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   435
      Left            =   8760
      TabIndex        =   38
      Top             =   6090
      Width           =   1095
      _ExtentX        =   1931
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
      mIcon           =   "frmEmployees.frx":15CA
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   435
      Left            =   3840
      TabIndex        =   39
      Top             =   6090
      Width           =   1095
      _ExtentX        =   1931
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
      mIcon           =   "frmEmployees.frx":18E4
   End
   Begin lvButton.lvButtons_H cmdNext 
      Height          =   435
      Left            =   1140
      TabIndex        =   40
      Top             =   5010
      Width           =   465
      _ExtentX        =   820
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
      mIcon           =   "frmEmployees.frx":1BFE
   End
   Begin lvButton.lvButtons_H cmdPrevious 
      Height          =   435
      Left            =   600
      TabIndex        =   41
      Top             =   5010
      Width           =   465
      _ExtentX        =   820
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
      mIcon           =   "frmEmployees.frx":1F18
   End
   Begin lvButton.lvButtons_H cmdLast 
      Height          =   435
      Left            =   1680
      TabIndex        =   42
      Top             =   5010
      Width           =   465
      _ExtentX        =   820
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
      mIcon           =   "frmEmployees.frx":2232
   End
   Begin lvButton.lvButtons_H cmdFirst 
      Height          =   435
      Left            =   90
      TabIndex        =   43
      Top             =   5010
      Width           =   465
      _ExtentX        =   820
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
      mIcon           =   "frmEmployees.frx":254C
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEES"
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
      Left            =   3840
      TabIndex        =   27
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
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
      Left            =   2280
      TabIndex        =   26
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   10095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Education Background"
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
      Left            =   6000
      TabIndex        =   24
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label12 
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
      Left            =   6000
      TabIndex        =   23
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label11 
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
      Left            =   6000
      TabIndex        =   22
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label10 
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
      Left            =   6000
      TabIndex        =   21
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label9 
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
      Left            =   6000
      TabIndex        =   20
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label8 
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
      Left            =   6000
      TabIndex        =   19
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Hire Date"
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
      Left            =   2280
      TabIndex        =   18
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Birth"
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
      Left            =   2280
      TabIndex        =   17
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Left            =   2280
      TabIndex        =   16
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label4 
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
      Left            =   2280
      TabIndex        =   15
      Top             =   2400
      Width           =   1695
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
      Left            =   2280
      TabIndex        =   14
      Top             =   1800
      Width           =   1695
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
      Left            =   2280
      TabIndex        =   13
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
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
      Left            =   2280
      TabIndex        =   12
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblRecordNo 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   720
      TabIndex        =   25
      Top             =   240
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   6015
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmEmployees"
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
        MsgBox "Please save or cancel the new employee record .", vbExclamation, title
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            If empSearch = True Then
                Load frmEmployeeSearch
                frmEmployeeSearch.Show
                Else
                    
            End If
    End If
    Unload Me
End Sub
Private Sub cmdFirst_Click()

    If recEmployee.BOF And recEmployee.EOF Then
        MsgBox "There is no record in the database.", vbExclamation, title
        Exit Sub
    End If
    
    recEmployee.MoveFirst
    Call display
End Sub

Private Sub cmdLast_Click()

    If recEmployee.BOF And recEmployee.EOF Then
        MsgBox "There is no record in the database.", vbExclamation, title
        Exit Sub
    End If
    
    recEmployee.MoveLast
    Call display
End Sub
Private Sub cmdNext_Click()

    If recEmployee.BOF And recEmployee.EOF Then
        MsgBox "There is no record in the database.", vbExclamation, title
        Exit Sub
    End If
    
    recEmployee.MoveNext
    If recEmployee.EOF Then
        recEmployee.MoveLast
        MsgBox "This is the last employee", vbExclamation, title
    End If
    Call display
End Sub
Private Sub cmdOk_Click()
On Error GoTo abdel
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
 
    If txtTitle.Text = "" Then
        MsgBox "The title (Role) field can not be left blank.", vbExclamation, title
        txtTitle.SetFocus
        Exit Sub
    End If
    
    If comboGender.Text = "" Then
        MsgBox "The gender field can not be left blank.", vbExclamation, title
        comboGender.SetFocus
        Exit Sub
    End If

    If txtSalary.Text = "" Then
        MsgBox "The salary field can not be left blank.", vbExclamation, title
        txtSalary.SetFocus
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
    If BirthPicker.Value > Date Then
        MsgBox "Your date of birth can be in future", vbExclamation, title
        Exit Sub
    End If
    If BirthPicker.Value = HirePicker.Value Then
        MsgBox "Your date of birth can not be the same as the hire date", vbExclamation, title
        Exit Sub
    End If
    If BirthPicker.Value > HirePicker.Value Then
        MsgBox "Your date of birth can be after the hire date", vbExclamation, title
        Exit Sub
    End If
    
    recEmployee!employeeid = txtID.Text & ""
    recEmployee!firstname = txtFirstName.Text & ""
    recEmployee!lastname = txtLastName.Text & ""
    recEmployee!title = txtTitle.Text & ""
    recEmployee!Gender = comboGender.Text & ""
    recEmployee!birthdate = BirthPicker.Value & ""
    recEmployee!hiredate = HirePicker.Value & ""
    recEmployee!Address = txtAddress.Text & ""
    recEmployee!city = txtCity.Text & ""
    recEmployee!PostalCode = txtPostalCode.Text & ""
    recEmployee!Country = txtCountry.Text & ""
    recEmployee!phone = txtPhone.Text & ""
    recEmployee!Notes = txtNotes.Text & ""
    recEmployee.UpdateBatch adAffectCurrent
    
    list.Clear
    recEmployee.Requery
    recEmployee.MoveFirst
    Do While Not recEmployee.EOF
        list.AddItem recEmployee!firstname & "  " & recEmployee!lastname
        recEmployee.MoveNext
    Loop
    recEmployee.MoveFirst
    Call display
    cmdOk.Visible = False
    cmdUpdate.Visible = True
    Call DisableCmd
Exit Sub
abdel:
    MsgBox "Sorry, transactions not successfull", vbExclamation, title
End Sub

Private Sub cmdPrevious_Click()

    If recEmployee.BOF And recEmployee.EOF Then
        MsgBox "There is no record in the database.", vbExclamation, title
        Exit Sub
    End If
    
    recEmployee.MovePrevious
    If recEmployee.BOF Then
        recEmployee.MoveFirst
        MsgBox "This is the first employee", vbExclamation, title
    End If
    Call display
End Sub
Private Sub cmdNew_Click()

    If MsgBox("Are you sure you want to add new record ? ", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Call LockMe(False)
            Call EnableCmd
            BirthPicker.Value = Date
            HirePicker.Value = Date
'            txtDateOfBirth.Visible = False
 '           txtHireDate.Visible = False
            txtID.Text = autogen
            txtFirstName.Text = ""
            txtLastName.Text = ""
            txtTitle.Text = ""
            comboGender.Text = ""
            txtSalary.Text = ""
            txtAddress.Text = ""
            txtCity.Text = ""
            txtPostalCode.Text = ""
            txtCountry.Text = ""
            txtPhone.Text = ""
            txtNotes.Text = ""
            
            txtFirstName.SetFocus
    End If
        
End Sub

Private Sub cmdPrint_Click()
On Error GoTo abdel
        sql = "select * from employees"
        Set EmployeeReport.DataSource = con.Execute(sql)
        EmployeeReport.Show
        Set EmployeeReport = Nothing
    Exit Sub
abdel:
    MsgBox "Employee report not available for now", vbCritical, title
    
End Sub

Private Sub cmdSave_Click()
On Error GoTo abdel
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
    
    If comboGender.Text = "" Then
        MsgBox "The gender field can not be left blank.", vbExclamation, title
        comboGender.SetFocus
        Exit Sub
    End If
    
    If txtSalary.Text = "" Then
        MsgBox "The salary field can not be left blank.", vbExclamation, title
        txtSalary.SetFocus
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
    If BirthPicker.Value > Date Then
        MsgBox "Your date of birth can not be in future", vbExclamation, title
        Exit Sub
    End If
    If BirthPicker.Value = HirePicker.Value Then
        MsgBox "Your date of birth can not be the same as the hire date", vbExclamation, title
        Exit Sub
    End If
    If BirthPicker.Value > HirePicker.Value Then
        MsgBox "Your date of birth can not be after the hire date", vbExclamation, title
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to add the new record to the database?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            recEmployee.AddNew
            recEmployee!employeeid = txtID.Text & ""
            recEmployee!firstname = txtFirstName.Text & ""
            recEmployee!lastname = txtLastName.Text & ""
            recEmployee!title = txtTitle.Text & ""
            recEmployee!Gender = comboGender.Text & ""
            recEmployee!birthdate = BirthPicker.Value & ""
            recEmployee!hiredate = HirePicker.Value & ""
            recEmployee!Salary = txtSalary.Text & ""
            recEmployee!Address = txtAddress.Text & ""
            recEmployee!city = txtCity.Text & ""
            recEmployee!PostalCode = txtPostalCode.Text & ""
            recEmployee!Country = txtCountry.Text & ""
            recEmployee!phone = txtPhone.Text & ""
            recEmployee!Notes = txtNotes.Text & ""
            recEmployee.Update
            
            MsgBox "Transactions successfully saved.", vbInformation, title
            Call DisableCmd
            Call LockMe(True)
            list.Clear
            recEmployee.Requery
            recEmployee.MoveFirst
            Do While Not recEmployee.EOF
                list.AddItem recEmployee!firstname & "  " & recEmployee!lastname
                recEmployee.MoveNext
            Loop
            recEmployee.MoveFirst
            Call display
        End If
        
Exit Sub
abdel:
    MsgBox "Sorry, transactions not successfully saved", vbExclamation, title
End Sub
Private Sub CmdCancel_Click()

    If MsgBox("Are you sure you want to cancel new employee record ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            'recEmployee.CancelUpdate
            Call DisableCmd
            Call LockMe(True)
            cmdOk.Visible = False
            cmdUpdate.Visible = True
            list.Clear
            recEmployee.Requery
            If Not recEmployee.EOF Then
                recEmployee.MoveFirst
                Do While Not recEmployee.EOF
                    list.AddItem recEmployee!firstname & "  " & recEmployee!lastname
                    recEmployee.MoveNext
                Loop
                recEmployee.MoveFirst
                Call display
                Else
                    For Each control In Me
                        If TypeOf control Is TextBox Then
                            control.Text = ""
                        End If
                    Next
                    Call disableAll
                    comboGender.Text = ""
            End If
    End If
End Sub
Private Sub cmddelete_Click()
On Error GoTo abdel
    If MsgBox("Are you sure you want to delete the records of " & Trim(txtFirstName.Text) & " " & Trim(txtLastName.Text) & " ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            recEmployee.Delete adAffectCurrent
            list.Clear
            recEmployee.Requery
            If Not recEmployee.EOF Then
                recEmployee.MoveFirst
                Do While Not recEmployee.EOF
                    list.AddItem recEmployee!firstname & "  " & recEmployee!lastname
                    recEmployee.MoveNext
                Loop
                recEmployee.MoveFirst
                Call display
                Else
                    For Each control In Me
                        If TypeOf control Is TextBox Then
                            control.Text = ""
                        End If
                        Call disableAll
                        comboGender.Text = ""
                    Next
            End If
    End If
Exit Sub
abdel:
    MsgBox "Sorry, employee records could not be deleted", vbExclamation, title
End Sub

Private Sub cmdSearch_Click()

    Load frmEmployeeSearch
    frmEmployeeSearch.Show
    
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo abdel
    If MsgBox("Are you sure you want to update the records of " & Trim(txtFirstName.Text) & " " & Trim(txtLastName.Text) & " ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            Call LockMe(False)
            cmdUpdate.Visible = False
            cmdOk.Visible = True
            Call disableAll
            cmdNew.Enabled = False
            cmdCancel.Enabled = True
            HirePicker.Visible = True
            list.Enabled = False
    End If
Exit Sub
abdel:
    MsgBox "Sorry, employee records could not be updated", vbExclamation, title
End Sub

Private Sub comboGender_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()

    frmScreen.Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call ConnectMe
'    cal.Visible = False
'    cmdCal.Visible = False
    If recEmployee.BOF And recEmployee.EOF Then
        MsgBox "There is no employee in the database!", vbExclamation, title
        Call LockMe(True)
        disableAll
        Exit Sub
    End If
    
    Call DisableCmd
    Call LockMe(True)
    list.Clear
    recEmployee.Requery
    recEmployee.MoveFirst
    Do While Not recEmployee.EOF
        list.AddItem recEmployee!firstname & "  " & recEmployee!lastname
        recEmployee.MoveNext
    Loop
    recEmployee.MoveFirst
    Call display
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmScreen.Enabled = True
End Sub

Private Sub list_Click()

    position = list.ListIndex
    recEmployee.MoveFirst
    recEmployee.Move position
    Call display
End Sub
Public Sub display()

    txtID.Text = recEmployee!employeeid & ""
    txtFirstName.Text = recEmployee!firstname & ""
    txtLastName.Text = recEmployee!lastname & ""
    txtTitle.Text = recEmployee!title & ""
    comboGender.Text = recEmployee!Gender & ""
    BirthPicker.Value = recEmployee!birthdate & ""
    HirePicker.Value = recEmployee!hiredate & ""
    txtSalary.Text = recEmployee!Salary & ""
    txtAddress.Text = recEmployee!Address & ""
    txtCity.Text = recEmployee!city & ""
    txtPostalCode.Text = recEmployee!PostalCode & ""
    txtCountry.Text = recEmployee!Country & ""
    txtPhone.Text = recEmployee!phone & ""
    txtNotes.Text = recEmployee!Notes & ""
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
    
    recGen.Open "select max(EmployeeID) from Employees", con, adOpenDynamic, adLockOptimistic
    
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
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
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
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    cmdSearch.Enabled = False
End Sub
Private Sub txtcity_Validate(Cancel As Boolean)
    txtCity.Text = StrConv(txtCity.Text, vbProperCase)
End Sub

Private Sub txtCountry_Validate(Cancel As Boolean)
    txtCountry.Text = StrConv(txtCountry.Text, vbProperCase)
End Sub

Private Sub txtFirstName_Validate(Cancel As Boolean)
    txtFirstName.Text = StrConv(txtFirstName.Text, vbProperCase)
End Sub
Private Sub txtLastName_Validate(Cancel As Boolean)
    txtLastName.Text = StrConv(txtLastName.Text, vbProperCase)
End Sub

Private Sub txtNotes_Validate(Cancel As Boolean)
    txtNotes.Text = StrConv(txtNotes.Text, vbProperCase)
End Sub
Private Sub txtSalary_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtTitle_Validate(Cancel As Boolean)
    txtTitle.Text = StrConv(txtTitle.Text, vbProperCase)
End Sub







