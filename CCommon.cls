VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum KeysGroup
   Integers = 1
   Floats = 2
   Alphabets = 3
   AlphaNumeric = 4
   AnyKey = 0
   NoKey = 5
End Enum

Enum RceordAction
   AddingNew = 1
   Editing = 2
   Deleting = 3
End Enum

Enum SecurityObjects
   soChartOfAccounts = 1
   soListOfProducts = 2
   soProductGroup = 3
   soProductCategories = 4
   soPurchaseInvoice = 5
   soSaleInvoice = 6
   soCashBook = 7
   soOpeningStock = 8
   soOpeningBalance = 9
   soPurchaseReturn = 10
   soSaleReturn = 11
   soJournalVoucher = 12
'   soIssueToProduction = 13
'   soProduction = 14
End Enum

Enum enumRights
  CanView = 1
  CanAdd = 2
  CanEdit = 3
  CanDelete = 4
End Enum

Enum enumControlGroup
  eGrpBANK = 1
  eGrpPARTIES = 2
  eGrpCUSTOMERS = 3
  eGrpSUPPLIERS = 4
  eGrpSALES = 5
  eGrpPURCHASES = 6
  eGrpEXPENSE = 7
  eGrpLCACCTS = 8
  egrpCASH = 9
End Enum
