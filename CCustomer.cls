VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAcctSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Class CAcctSub
Option Explicit
Private m_lID As Long
Private m_sAddress As String
Private m_sCity As String
Private m_sAreaID As String
Private m_sPhoneNumber1 As String
Private m_sPhoneNumber2 As String
Private m_sMobileNumber As String
Private m_sFaxNumber As String
Private m_sEmail As String
Private m_sContactName As String
Private m_sContactTitle As String
Private m_dTradeDiscount As Double
Private m_sSTaxRegNumber As String
Private m_sNTN As String
Private m_iCreditDays As Variant
Private m_dCustBalance As Double
Private m_sBalanceType As String
Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private m_bIsActive As Boolean
Public Event DataValidated(bValid As Boolean)

Public Property Get ID() As Long
    ID = m_lID
End Property

Public Property Let Active(ByVal vData As Boolean)
    m_bIsActive = vData
    ValidateData
End Property
Public Property Get Active() As Boolean
    Active = m_bIsActive
End Property

Public Property Let Address(ByVal vData As String)
    m_sAddress = vData
    ValidateData
End Property
Public Property Get Address() As String
    Address = m_sAddress
End Property

Public Property Let City(ByVal vData As String)
    m_sCity = vData
    ValidateData
End Property
Public Property Get City() As String
    City = m_sCity
End Property

Public Property Let AreaID(ByVal vData As String)
    m_sAreaID = vData
    ValidateData
End Property
Public Property Get AreaID() As String
    AreaID = m_sAreaID
End Property

Public Property Let FaxNumber(ByVal vData As String)
    m_sFaxNumber = vData
    ValidateData
End Property
Public Property Get FaxNumber() As String
    FaxNumber = m_sFaxNumber
End Property

Public Property Let PhoneNumber1(ByVal vData As String)
    m_sPhoneNumber1 = vData
    ValidateData
End Property
Public Property Get PhoneNumber1() As String
    PhoneNumber1 = m_sPhoneNumber1
End Property

Public Property Let PhoneNumber2(ByVal vData As String)
    m_sPhoneNumber2 = vData
    ValidateData
End Property
Public Property Get PhoneNumber2() As String
    PhoneNumber2 = m_sPhoneNumber2
End Property

Public Property Let MobileNumber(ByVal vData As String)
    m_sMobileNumber = vData
    ValidateData
End Property
Public Property Get MobileNumber() As String
    MobileNumber = m_sMobileNumber
End Property

Public Property Let Email(ByVal vData As String)
    m_sEmail = vData
    ValidateData
End Property
Public Property Get Email() As String
    Email = m_sEmail
End Property
Public Property Let ContactName(ByVal vData As String)
    m_sContactName = vData
    ValidateData
End Property
Public Property Get ContactName() As String
    ContactName = m_sContactName
End Property

Public Property Let ContactTitle(ByVal vData As String)
    m_sContactTitle = vData
    ValidateData
End Property
Public Property Get ContactTitle() As String
    ContactTitle = m_sContactTitle
End Property

Public Property Let TradeDiscount(ByVal vData As Double)
    m_dTradeDiscount = vData
    ValidateData
End Property
Public Property Get TradeDiscount() As Double
    TradeDiscount = m_dTradeDiscount
End Property

Public Property Let STaxRegNumber(ByVal vData As String)
    m_sSTaxRegNumber = vData
    ValidateData
End Property
Public Property Get STaxRegNumber() As String
    STaxRegNumber = m_sSTaxRegNumber
End Property

Public Property Let NTN(ByVal vData As String)
    m_sNTN = vData
    ValidateData
End Property
Public Property Get NTN() As String
    NTN = m_sNTN
End Property

Public Property Let CreditDays(ByVal vData As Variant)
    m_iCreditDays = vData
    ValidateData
End Property
Public Property Get CreditDays() As Variant
    CreditDays = m_iCreditDays
End Property

Public Property Let CustomerName(ByVal sData As String)
    m_sCustomerName = sData
    ValidateData
End Property

Public Property Get CustomerName() As String
   CustomerName = m_sCustomerName
End Property

Public Property Get Balance() As Double
    Balance = m_dCustBalance
End Property

Public Property Get BalanceType() As String
    BalanceType = m_sBalanceType
End Property


Private Sub ValidateData()
   Dim bValid As Boolean
   bValid = True
   
   'bValid = bValid And m_sCustID <> ""
   bValid = bValid And m_sCustomerName <> ""
   'bValid = bValid And m_sContactName <> ""
   
   '
   ' Depending on your situation,
   ' the contact's title may or may not
   ' be required. For now, it's not.
   '
   'bValid = bValid And m_sContactTitle <> ""
   bValid = bValid And m_sAddress <> ""
   bValid = bValid And m_sCity <> ""
   'bValid = bValid And m_iCreditDays <> 0
   'bValid = bValid And m_dTradeDiscount <> 0
'   '
'   ' Depending on your situation,
'   ' the Region may or may not
'   ' be required. For now, it's not.
'   '
'   'bValid = bValid And m_sRegion <> ""
'   ' US States must be two letters long
'   'bValid = bValid And Len(m_sRegion) = 2
'
'   bValid = bValid And m_sPostalCode <> ""
'   bValid = bValid And m_sCountry <> ""
'   bValid = bValid And m_sPhoneNumber <> ""
'   '
   ' Depending on your situation,
   ' the fax number may or may not
   ' be required. For now, it's not.
   '
   'bValid = bValid And m_sFaxNumber <> ""
   
   m_bIsDataValid = bValid
   RaiseEvent DataValidated(bValid)
End Sub

Private Sub Class_Initialize()
    ValidateData
    Set m_rsCustomer = New ADODB.Recordset
End Sub

Public Sub Init(Optional lID As Long = 0)
   
   ' If the ID is missing, the object
   ' is being instantiated to be filled
   ' with new data.
   OpenRs
   If lID = 0 Then
      m_bIsNewRecord = True
      Exit Sub
    Else
'       Dim cmdCustomer As ADODB.Command
'       Dim prmCustomerID As ADODB.Parameter
'       Set cmdCustomer = New ADODB.Command
'       Set prmCustomerID = New ADODB.Parameter
'       Dim m_RS As ADODB.Recordset
'       Set m_RS = New ADODB.Recordset
'       Set prmCustomerID = cmdCustomer.CreateParameter("paramCustID", adInteger, adParamInput, , lID)
'       cmdCustomer.Parameters.Append prmCustomerID
'       cmdCustomer.ActiveConnection = m_objConnectDB.cnnMyshop
'       cmdCustomer.CommandType = adCmdStoredProc
'       cmdCustomer.CommandText = "usp_GetCustomer"
'       Set m_RS = cmdCustomer.Execute
       m_lID = m_rsCustomer("ID") & ""
       m_sCustID = m_rsCustomer("CustomerID") & ""
       m_sCustomerName = m_rsCustomer("CustomerName") & ""
       m_sAddress = m_rsCustomer("Address") & ""
       m_sCity = m_rsCustomer("City") & ""
       m_sAreaID = m_rsCustomer("AreaID") & ""
       m_sPhoneNumber1 = m_rsCustomer("PhoneNumber1") & ""
       m_sPhoneNumber2 = m_rsCustomer("PhoneNumber2") & ""
       m_sMobileNumber = m_rsCustomer("MobileNumber") & ""
       m_sFaxNumber = m_rsCustomer("FaxNumber") & ""
       m_sEmail = m_rsCustomer("Email") & ""
       m_sContactName = m_rsCustomer("ContactName") & ""
       m_sContactTitle = m_rsCustomer("ContactTitle") & ""
       m_dTradeDiscount = m_rsCustomer("TradeDiscount") & ""
       m_sSTaxRegNumber = m_rsCustomer("STaxRegNumber") & ""
       m_sNTN = m_rsCustomer("NTN") & ""
       m_iCreditDays = m_rsCustomer("CreditDays") & ""
    
       
       m_bIsActive = m_rsCustomer("Active")
       ' Clean up
       m_rsCustomer.Close
       Set m_rsCustomer = Nothing
       'Get Customer Balance from query
       Dim cmdCustBalance As ADODB.Command
       Dim prmCustomerID As ADODB.Parameter
       Dim rsTemp As ADODB.Recordset
       Set cmdCustBalance = New ADODB.Command
       Set rsTemp = New ADODB.Recordset
       Set prmCustomerID = cmdCustBalance.CreateParameter("paramCustID", adInteger, adParamInput, , lID)
       cmdCustBalance.Parameters.Append prmCustomerID
       cmdCustBalance.ActiveConnection = m_objConnectDB.cnnMyshop
       cmdCustBalance.CommandText = "usp_GetAccountBalance"
       cmdCustBalance.CommandType = adCmdStoredProc
       rsTemp.CursorType = adOpenDynamic
       Set rsTemp = cmdCustBalance.Execute
       If rsTemp.BOF Or rsTemp.EOF Then
           m_dCustBalance = 0
           m_sBalanceType = ""
       Else
        If (rsTemp("SumOfDebit").Value - rsTemp("SumOfCredit").Value) > 0 Then
           m_dCustBalance = rsTemp("SumOfDebit").Value - rsTemp("SumOfCredit").Value
           m_sBalanceType = "Debit"
        Else
           m_dCustBalance = rsTemp("SumOfCredit").Value - rsTemp("SumOfDebit").Value
           m_sBalanceType = "Credit"
        End If
      End If
        rsTemp.Close
        Set rsTemp = Nothing
     End If
End Sub

Public Sub Save()
   If m_bIsNewRecord Then
      OpenRs
      m_rsCustomer.MoveLast
      m_rsCustomer.AddNew
   Else
      m_rsCustomer.MoveFirst
      m_rsCustomer.Find "ID=" & m_lID 'To Be replaced with better code.
   End If
       m_rsCustomer("CustomerID") = m_sCustID
       m_rsCustomer("CustomerName") = m_sCustomerName
       m_rsCustomer("Address") = m_sAddress
       m_rsCustomer("City") = m_sCity
       m_rsCustomer("AreaID") = m_sAreaID
       m_rsCustomer("PhoneNumber1") = m_sPhoneNumber1
       m_rsCustomer("PhoneNumber2") = m_sPhoneNumber2
       m_rsCustomer("MobileNumber") = m_sMobileNumber
       m_rsCustomer("FaxNumber") = m_sFaxNumber
       m_rsCustomer("Email") = m_sEmail
       m_rsCustomer("ContactName") = m_sContactName
       m_rsCustomer("ContactTitle") = m_sContactTitle
       m_rsCustomer("TradeDiscount") = m_dTradeDiscount
       m_rsCustomer("STaxRegNumber") = m_sSTaxRegNumber
       m_rsCustomer("NTN") = m_sNTN
       m_rsCustomer("Active") = m_bIsActive
      'm_rsCustomer("CreditDays") = m_iCreditDays
       m_rsCustomer.Update
       m_rsCustomer.Requery
       m_lID = m_rsCustomer("ID")
       m_bIsNewRecord = False
End Sub

Public Sub OpenRs(Optional lID As Long = 0)
    Dim strRecordset As String
    If lID = 0 Then
    strRecordset = "Select * from Customers order by CustomerName"
    Else
    strRecordset = "Select * from Customers WHERE Customers.ID=" & lID
    End If
    If m_rsCustomer.State <> adStateOpen Then
        m_rsCustomer.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub
Public Sub UpdateFormCombo(ByRef m_objCmb As ComboBox)
    Dim a As Integer
    m_objCmb.Clear
    a = 0
    OpenRs
    m_rsCustomer.Filter = "Active = 1"
    m_rsCustomer.MoveFirst
    m_objCmb.Clear
    Do While Not m_rsCustomer.EOF
        m_objCmb.AddItem m_rsCustomer("CustomerName") + " , " + m_rsCustomer("City")
        m_objCmb.ItemData(a) = m_rsCustomer("ID")
        a = a + 1
        m_rsCustomer.MoveNext
    Loop
    m_rsCustomer.Filter = ""
End Sub

