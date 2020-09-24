Attribute VB_Name = "MdlConnection"
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
Public recCategory As New Recordset
Public recCredit As New Recordset
Public recCustomer As New Recordset
Public recEmployee As New Recordset
Public recOrder As New Recordset
Public recOrderDetails As New Recordset
Public recPayment As New Recordset
Public recPayslip As New Recordset
Public recProduct As New Recordset
Public recProduct2 As New Recordset
Public recPurchase As New Recordset
Public recSale As New Recordset
Public recSalesInvoice As New Recordset
Public recUsers As New Recordset
Public recUsersLog As New Recordset


Public con As New Connection
Public cmd As New Command

Public blAddUser As Boolean
Public blUpdateUser As Boolean
Public blSearchEmployee As Boolean
Public blSearchCustomer As Boolean
Public blAddCategory As Boolean
Public blUpdateCategory As Boolean
Public blSaleRep As Boolean
Public blPurchaseRep As Boolean
Public blOrder As Boolean
Public blNewOrder As Boolean
Public blCredit As Boolean
Public blCash As Boolean
Public blCategory As Boolean
Public blProduct As Boolean
Public blAllOrders As Boolean
Public blOrderDetails As Boolean
Public blOrd As Boolean
Public blInvoice As Boolean
Public blOrderDetails1 As Boolean
Public blOrderRep As Boolean
Public AllUsers As Boolean
Public admin As Boolean
Public emp As Boolean
Public custSearch As Boolean
Public empSearch As Boolean

Public lst As ListItem
Public lstItem As ListItem
Public title As String
Public Sub ConnectMe()
On Error Resume Next
    title = "AbdelSoft"
    con.Open "provider = microsoft.jet.oledb.4.0;data source = " & App.Path & "\Database\abdelsoft.mdb"
        
    recCategory.Open "select * from category order by categoryname", con, adOpenDynamic, adLockOptimistic
    recCredit.Open "select * from credit", con, adOpenDynamic, adLockOptimistic
    recCustomer.Open "select * from customer order by contactname", con, adOpenDynamic, adLockOptimistic
    recEmployee.Open "select * from employees order by employeeID", con, adOpenDynamic, adLockOptimistic
    recOrder.Open "select * from Orders", con, adOpenDynamic, adLockOptimistic
    recOrderDetails.Open "select * from OrderDetails", con, adOpenDynamic, adLockOptimistic
    recPayment.Open "select * from payment", con, adOpenDynamic, adLockOptimistic
    recPayslip.Open "select * from payslip", con, adOpenDynamic, adLockOptimistic
    recProduct.Open "select * from product order by productname", con, adOpenDynamic, adLockOptimistic
    recProduct2.Open "select * from product order by productname", con, adOpenDynamic, adLockOptimistic
    recPurchase.Open "select * from purchase order by productname", con, adOpenDynamic, adLockOptimistic
    recSale.Open "select * from sale order by saleid", con, adOpenDynamic, adLockOptimistic
    recSalesInvoice.Open "select * from salesInvoice ", con, adOpenDynamic, adLockOptimistic
    recUsers.Open "select * from users order by loginName", con, adOpenDynamic, adLockOptimistic
    

End Sub
Public Sub UserLogin()
    recUsersLog.Open "select * from usersLog ", con, adOpenDynamic, adLockOptimistic
    recUsersLog.AddNew
    
    recUsersLog!userslogID = autogen
    recUsersLog!role = frmScreen.lblRole.Caption
    recUsersLog!loginname = frmScreen.lblName.Caption
    recUsersLog!logintime = CDate(frmScreen.lblTime.Caption)
    recUsersLog!logindate = Date
    recUsersLog.Update
End Sub

Public Sub UserLogout()
On Error Resume Next
    recUsersLog!logoutdate = Date
    recUsersLog!logouttime = Time
    recUsersLog.UpdateBatch adAffectCurrent
    Set recUsersLog = Nothing
End Sub

Public Function autogen()
'On Error Resume Next

    Dim recGen As New Recordset
    
    recGen.Open "select max(userslogid) from UsersLog", con, adOpenDynamic, adLockOptimistic
    
    recGen.MoveFirst
    If IsNull(recGen(0)) Then
        autogen = 1
        
        Else
        
        autogen = Val(recGen(0) + 1)
    End If
End Function
