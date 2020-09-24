VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   1155
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dialog 
      Left            =   720
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      HelpCommand     =   3
      HelpFile        =   "\Help\AbdelSoft"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLock 
         Caption         =   "Lock application"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Log out"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Transaction"
      Begin VB.Menu mnuPurchase 
         Caption         =   "Purchase"
      End
      Begin VB.Menu mnuSales 
         Caption         =   "Sales"
      End
      Begin VB.Menu mnuOrders 
         Caption         =   "Orders"
         Begin VB.Menu mnuCredit 
            Caption         =   "Orders on credit"
         End
         Begin VB.Menu mnuCash 
            Caption         =   "Cash Orders"
         End
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu mnuPays 
         Caption         =   "Payslip"
         Begin VB.Menu mnupaysliprep 
            Caption         =   "Payslip report"
         End
      End
      Begin VB.Menu mnuPur 
         Caption         =   "Purchase"
         Begin VB.Menu mnuPurchaseRep 
            Caption         =   "Purchase report"
         End
      End
      Begin VB.Menu mnuSalerep 
         Caption         =   "Sales"
         Begin VB.Menu mnuSalesRep 
            Caption         =   "Sales report"
         End
         Begin VB.Menu mnuSalesInv 
            Caption         =   "All invoices"
         End
      End
      Begin VB.Menu mnuOrdersRep 
         Caption         =   "Orders"
         Begin VB.Menu mnuAllOrders 
            Caption         =   "All orders"
         End
         Begin VB.Menu mnuAllInvoices 
            Caption         =   "All orders invoices"
         End
         Begin VB.Menu mnuAllCust 
            Caption         =   "All customers owing"
         End
      End
      Begin VB.Menu mnuTime 
         Caption         =   "Users time record"
         Begin VB.Menu mnuAllUsers 
            Caption         =   "All users"
         End
         Begin VB.Menu mnuAdminTime 
            Caption         =   "Administration"
         End
         Begin VB.Menu mnuOthers 
            Caption         =   "Other users"
         End
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Administration"
      Begin VB.Menu mnuPers 
         Caption         =   "Personnel"
         Begin VB.Menu mnuEmp 
            Caption         =   "View all employees"
         End
         Begin VB.Menu mnuSearchEmp 
            Caption         =   "Search"
         End
      End
      Begin VB.Menu mnuSalaries 
         Caption         =   "Salaries"
         Begin VB.Menu mnuAllSal 
            Caption         =   "View all salaries"
         End
         Begin VB.Menu mnuPayslip 
            Caption         =   "Payslip"
         End
      End
      Begin VB.Menu mnuCustomers 
         Caption         =   "Customers"
         Begin VB.Menu mnuAllCustomers 
            Caption         =   "View all customers"
         End
         Begin VB.Menu mnuSearch 
            Caption         =   "Search"
         End
      End
      Begin VB.Menu mnuCat 
         Caption         =   "Categories"
         Begin VB.Menu mnuAllcat 
            Caption         =   "View all categories"
         End
         Begin VB.Menu mnuNewCat 
            Caption         =   "New"
         End
         Begin VB.Menu mnuUpdateCat 
            Caption         =   "Update"
         End
         Begin VB.Menu mnuDeleteCat 
            Caption         =   "Delete"
         End
      End
      Begin VB.Menu mnuProd 
         Caption         =   "Products"
         Begin VB.Menu mnuAllProd 
            Caption         =   "View all products"
         End
         Begin VB.Menu mnuNewP 
            Caption         =   "New"
         End
         Begin VB.Menu mnuUpdateP 
            Caption         =   "Update"
         End
         Begin VB.Menu mnuDeleteP 
            Caption         =   "Delete"
         End
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Users"
         Begin VB.Menu mnuViewAllU 
            Caption         =   "View all users"
         End
         Begin VB.Menu mnuNewU 
            Caption         =   "New"
         End
         Begin VB.Menu mnuUpdateU 
            Caption         =   "Update"
         End
         Begin VB.Menu mnuDeleteU 
            Caption         =   "Delete"
         End
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Database"
      Begin VB.Menu mnuBack 
         Caption         =   "Back Up"
      End
      Begin VB.Menu mnuMan 
         Caption         =   "Database Management"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuCalc 
         Caption         =   "Calculatrice"
      End
      Begin VB.Menu mnuCal 
         Caption         =   "Calendrier"
      End
      Begin VB.Menu mnuDT 
         Caption         =   "Date and Time"
      End
      Begin VB.Menu mnuNote 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuTech 
         Caption         =   "Technical support"
         Begin VB.Menu mnuCont 
            Caption         =   "Contents"
         End
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About RestauSys"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub mnuAbout_Click()

    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mnuAdminTime_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmUsersLog
        frmUsersLog.Show
        admin = True
        AllUsers = False
        emp = False
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuAllcat_Click()

    Load frmCategory
    frmCategory.Show
    
End Sub

Private Sub mnuAllCust_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
    Load frmCredit
    frmCredit.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
    
End Sub

Private Sub mnuAllCustomers_Click()

    Load frmCustomer
    frmCustomer.Show
End Sub

Private Sub mnuAllInvoices_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
    Load frmAllOrderPayment
    frmAllOrderPayment.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuAllOrders_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
    Load frmCustomerAndOrder
    frmCustomerAndOrder.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuAllProd_Click()

    Load frmProduct
    frmProduct.Show
    
End Sub

Private Sub mnuAllSal_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmSalary
        frmSalary.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuAllUsers_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmUsersLog
        frmUsersLog.Show
        AllUsers = True
        admin = False
        emp = False
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuBack_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmBackUp
        frmBackUp.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbExclamation, title
    End If
End Sub

Private Sub mnuCal_Click()
On Error GoTo abdel
    Load frmCalendar
    frmCalendar.Show
    Exit Sub
abdel:
    MsgBox "The calendar is not available for now...", vbExclamation, title
End Sub

Private Sub mnuCalc_Click()
On Error GoTo abdel
    Shell "calc"
    Exit Sub
abdel:
    MsgBox "Calculator is not available for now...", vbExclamation, title
End Sub

Private Sub mnuCash_Click()

    blCredit = False
    blCash = True
    Load frmOrder
    frmOrder.Show
    frmOrder.Label11.Caption = "Cash Order"
    
End Sub

Private Sub mnuCont_Click()
On Error GoTo abdel
    dialog.ShowHelp
    Exit Sub
abdel:
    MsgBox "Help is not available for now", vbExclamation, title
End Sub

Private Sub mnuCredit_Click()

    blCredit = True
    blCash = False
    Load frmOrder
    frmOrder.Show
    frmOrder.Label11.Caption = "Credit Order"
    
End Sub

Private Sub mnuDeleteCat_Click()

    Load frmCategory
    frmCategory.Show
    
End Sub

Private Sub mnuDeleteP_Click()

    Load frmProduct
    frmProduct.Show
End Sub

Private Sub mnuDeleteU_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmUsers
        frmUsers.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuDT_Click()
On Error GoTo abdel
    Dim dblReturn As Double
    If frmScreen.lblRole.Caption = "Administrator" Then
        dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbExclamation, title
    End If
    Exit Sub
abdel:
    MsgBox "The date and time is not available for now", vbExclamation, title

End Sub

Private Sub mnuEmp_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmEmployees
        frmEmployees.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuExit_Click()
    If MsgBox("Are you sure you want to quit the application ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
        sndPlaySound App.Path & "\Media\reminder.wav", &H1
        Call UserLogout
        End
    End If
End Sub

Private Sub mnuLock_Click()
    Load frmLockApplication
    frmLockApplication.Show
End Sub

Private Sub mnuLogout_Click()
    If MsgBox("Are you sure you want to log out ?", vbQuestion + 4, title) = vbNo Then
        Exit Sub
        Else
            frmScreen.lblRole.Caption = ""
            frmScreen.lblName.Caption = ""
            Call UserLogout
            Unload frmLogin
            frmLogin.Show
    End If
End Sub

Private Sub mnuMan_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmWarning
        frmWarning.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuNewCat_Click()

    Load frmCategory
    frmCategory.Show
    
End Sub

Private Sub mnuNewP_Click()

    Load frmProduct
    frmProduct.Show
End Sub

Private Sub mnuNewU_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmUsers
        frmUsers.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuNote_Click()
On Error GoTo abdel
    Shell "notepad.exe", vbNormalFocus
    Exit Sub
abdel:
    MsgBox "Notepad not available for now", vbExclamation, title
End Sub

Private Sub mnuOthers_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmUsersLog
        frmUsersLog.Show
        emp = True
        admin = False
        AllUsers = False
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuPayslip_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmPaySlip
        frmPaySlip.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnupaysliprep_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmPayslipReport
        frmPayslipReport.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuPurchase_Click()
    Load frmPurchase
    frmPurchase.Show
End Sub

Private Sub mnuPurchaseRep_Click()
 recPurchase.Requery
    If recPurchase.BOF And recPurchase.EOF Then
        MsgBox "No purchase transactions available", vbExclamation, title
        Exit Sub
    End If
    
    If frmScreen.lblRole.Caption = "Administrator" Then
        blSaleRep = False
        blPurchaseRep = True
        Load frmSalesReport
        frmSalesReport.Show
        frmSalesReport.lbl.Caption = "PURCHASE REPORT"
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuSales_Click()

    Load frmSale
    frmSale.Show
    
End Sub

Private Sub mnuSalesInv_Click()
    If frmScreen.lblRole.Caption = "Administrator" Then
    Load frmSalesInvoiceReport
    frmSalesInvoiceReport.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuSalesRep_Click()

 recSale.Requery
    If recSale.BOF And recSale.EOF Then
        MsgBox "No Sale transactions available", vbExclamation, title
        Exit Sub
    End If
    
    If frmScreen.lblRole.Caption = "Administrator" Then
        blSaleRep = True
        blPurchaseRep = False
        Load frmSalesReport
        frmSalesReport.Show
        frmSalesReport.lbl.Caption = "SALES REPORT"
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuSearch_Click()

    custSearch = True
    Load frmCustomerSearch
    frmCustomerSearch.Show
End Sub

Private Sub mnuSearchEmp_Click()
    If frmScreen.lblRole.Caption = "Administrator" Then
        empSearch = True
        Load frmEmployeeSearch
        frmEmployeeSearch.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuUpdateCat_Click()

    Load frmCategory
    frmCategory.Show
End Sub

Private Sub mnuUpdateP_Click()

    Load frmProduct
    frmProduct.Show
End Sub

Private Sub mnuUpdateU_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmUsers
        frmUsers.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub

Private Sub mnuViewAllU_Click()

    If frmScreen.lblRole.Caption = "Administrator" Then
        Load frmUsers
        frmUsers.Show
        Exit Sub
        Else
            MsgBox "Access denied, see system administrator", vbCritical, title
    End If
End Sub
