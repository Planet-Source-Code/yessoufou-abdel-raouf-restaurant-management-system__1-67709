VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} SalesInvoiceEnv 
   ClientHeight    =   7695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8610
   _ExtentX        =   15187
   _ExtentY        =   13573
   FolderFlags     =   1
   TypeLibGuid     =   "{DC3F057A-AF54-4F1A-AF73-AE402F5D00BB}"
   TypeInfoGuid    =   "{EB4F104A-4C7B-4738-A7AF-96B82F20E820}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "con"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=AbdelSoft.mdb;Persist Security Info=False"
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   2
   BeginProperty Recordset1 
      CommandName     =   "cmdInvoice"
      CommDispId      =   1007
      RsDispId        =   1015
      CommandText     =   "select invoiceID, amountdue, amountpaid,changes from salesInvoice"
      ActiveConnectionName=   "con"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "invoiceID"
         Caption         =   "invoiceID"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "amountdue"
         Caption         =   "amountdue"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "amountpaid"
         Caption         =   "amountpaid"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "changes"
         Caption         =   "changes"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "cmdSale"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "select * from sale"
      ActiveConnectionName=   "con"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdInvoice"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "SaleID"
         Caption         =   "SaleID"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "InvoiceID"
         Caption         =   "InvoiceID"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ProductName"
         Caption         =   "ProductName"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "UnitPrice"
         Caption         =   "UnitPrice"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Quantity"
         Caption         =   "Quantity"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Total"
         Caption         =   "Total"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "Date"
         Caption         =   "Date"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "Time"
         Caption         =   "Time"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "InvoiceID"
         ChildField      =   "InvoiceID"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "SalesInvoiceEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

