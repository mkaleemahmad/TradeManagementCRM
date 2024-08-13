VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccounts2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Basic Info
Dim m_bIsDataValid As Boolean
Dim m_lAccountID As Long
Dim m_sAccountName As String
Dim m_lDebit As Long
Dim m_lCredit As Long
Dim m_bEditable As Boolean
Dim m_lGroupID As Long
Dim m_lSubGroupID As Long
Dim m_sbalancetype As String
Dim m_bActive As Boolean
'Details
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

Dim m_rsAccts As New ADODB.Recordset
 
Dim m_bValid As Boolean

' AccountID
Public Property Get AccountID() As Long
AccountID = m_lAccountID
End Property

' Account Name
Public Property Get AccountName() As String
AccountName = m_sAccountName
End Property

Public Property Let AccountName(ByVal NewAccountName As String)
m_sAccountName = NewAccountName
End Property
' Debit
Public Property Get Debit() As Long
Debit = m_lDebit
End Property

Public Property Let Debit(ByVal NewDebit As Long)
m_lDebit = NewDebit
End Property
' Credit
Public Property Get Credit() As Long
Credit = m_lCredit
End Property

Public Property Let Credit(ByVal NewCredit As Long)
m_lCredit = NewCredit
End Property
' GroupID
Public Property Get GroupID() As Long
GroupID = m_lGroupID
End Property

Public Property Let GroupID(ByVal NewGroupID As Long)
m_lGroupID = NewGroupID
End Property
' Editable
Public Property Get Editable() As Boolean
Editable = m_bEditable
End Property

Public Property Let Editable(ByVal IsEditable As Boolean)
 m_bEditable = IsEditable
End Property
'SubGroupID
Public Property Get SubGroupID() As Long
SubGroupID = m_lSubGroupID
End Property

Public Property Let SubGroupID(ByVal lNewValue As Long)
m_lSubGroupID = lNewValue
End Property
'Details'
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

Public Property Get BalanceType() As String
    BalanceType = m_sbalancetype
End Property

Sub OpenRs(sCriteria As String)
    Dim sSQL As String
    sSQL = "SELECT * FROM Accounts Where GroupID IN (Select GroupID From AccountsRights Where CanView=1 and UserID=" & frmLogin.UserID & ")"      '& IIf(sCriteria = "", "", " WHERE " & sCriteria)
    sSQL = sSQL & IIf(sCriteria = "", "", " AND " & sCriteria)
   'sSQL = "SELECT * FROM Accounts" & IIf(sCriteria = "", "", " WHERE " & sCriteria)
    sSQL = sSQL & " Order By AccountName"
    With m_rsAccts
      If .State = adStateOpen Then .Close
      .Open sSQL, m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
      If .EOF Then
        m_bValid = False
      Else
        m_bValid = True
      End If
    End With
End Sub

Public Sub Init(lAccountID As Long)
OpenRs "ID=" & lAccountID
With m_rsAccts
  If m_bValid Then
    m_lGroupID = NumVal(SoftVal(!GroupID))
    m_lSubGroupID = NumVal(SoftVal(!SubGroupID))
    m_lDebit = NumVal(SoftVal(!Debit))
    m_lCredit = NumVal(SoftVal(!Credit))
    m_bEditable = NumVal(SoftVal(!Editable))
    m_sAccountName = !AccountName
    m_lAccountID = !ID
    m_bActive = NumVal(SoftVal(!Active))
    'Details
    m_sAddress = SoftVal(!Address)
    m_sCity = SoftVal(!City)
    m_sAreaID = SoftVal(!AreaID)
    m_sPhoneNumber1 = SoftVal(!PhoneNumber1)
    m_sPhoneNumber2 = SoftVal(!PhoneNumber2)
    m_sMobileNumber = SoftVal(!MobileNumber)
    m_sFaxNumber = SoftVal(!FaxNumber)
    m_sEmail = SoftVal(!Email)
    m_sContactName = SoftVal(!ContactName)
    m_sContactTitle = SoftVal(!ContactTitle)
    m_dTradeDiscount = SoftVal(!TradeDiscount)
    m_sSTaxRegNumber = SoftVal(!STaxRegNumber)
    m_sNTN = SoftVal(!NTN)
    m_iCreditDays = SoftVal(!CreditDays)
  Else
    m_lGroupID = 0
    m_lSubGroupID = 0
    m_lDebit = 0
    m_lCredit = 0
    m_bEditable = False
    m_bActive = False
    m_sAccountName = ""
    m_lAccountID = 0
  End If
End With
End Sub

Public Sub Save(bNew As Boolean)
With m_rsAccts
  If bNew Then
    If .State = adStateClosed Then OpenRs ("") ' .Open " S", m_objConnectDB.cnnMyshop
    .AddNew
  End If
    !GroupID = m_lGroupID
    '!SubGroupID = m_lSubGroupID
    !Debit = m_lDebit
    !Credit = m_lCredit
    !Editable = True
    !Active = True
    
    !AccountName = m_sAccountName
    'Details
    !Address = m_sAddress
    !City = m_sCity
    !AreaID = m_sAreaID
    !PhoneNumber1 = m_sPhoneNumber1
    !PhoneNumber2 = m_sPhoneNumber2
    !MobileNumber = m_sMobileNumber
    !FaxNumber = m_sFaxNumber
    !Email = m_sEmail
    !ContactName = m_sContactName
    !ContactTitle = m_sContactTitle
    !TradeDiscount = m_dTradeDiscount
    !STaxRegNumber = m_sSTaxRegNumber
    !NTN = m_sNTN
    !CreditDays = NumVal(SoftVal(m_iCreditDays))
  .Update
  If bNew Then
    Dim sSQL As String
    sSQL = "ID = (Select Max(ID) From Accounts)"
    OpenRs sSQL
    m_lGroupID = SoftVal(!GroupID)
    m_lSubGroupID = SoftVal(SubGroupID)
    m_lDebit = !Debit
    m_lCredit = !Credit
   ' m_bEditable = !Editable
    m_sAccountName = !AccountName
    m_lAccountID = !ID
    'm_bActive = !Active
  End If
  m_bValid = True
End With
End Sub

Public Sub Remove()
Dim sSQL As String
sSQL = "DELETE FROM Accounts WHERE ID=" & m_lAccountID
m_objConnectDB.cnnMyshop.Execute sSQL
m_bValid = False
m_rsAccts.Close
End Sub

Public Function IsValid() As Boolean
IsValid = m_bValid
End Function

Sub PopListView(refLvw As ListView)
Dim k As Long, rs As New ADODB.Recordset, sSQL As String
With rs
  If .State = adStateOpen Then .Close
  sSQL = "Select Accounts.[ID], AccountName, Accounts.GroupID, Accounts.SubGroupID, Active, MobileNumber, PhoneNumber1, City, ContactName, Address" & _
  " Editable , AccountGroups.[Description] As AGDesc, Accountsubgroups.[Description] As SAGDesc" & _
  " From " & _
  " (Accounts left join AccountGroups on Accounts.Groupid = AccountGroups.Groupid)" & _
  " left join Accountsubgroups on Accounts.subgroupid=AccountSubGroups.Subgroupid" & _
  " order by Accounts.AccountName"
  .Open sSQL, m_objConnectDB.cnnMyshop ', adOpenStatic, adLockOptimistic
  With refLvw
    .ColumnHeaders.Clear
    .ListItems.Clear
      With .ColumnHeaders
        .Add , "AccountName", "Account Name"
        .Add , "AccountID", "Account #"
        .Add , "GroupID", "Account Group"
        .Add , "ContactName", "Contact Name"
        .Add , "MobileNumber", "Mobile Number"
        .Add , "PhoneNumber1", "Phone Number"
 '       .Add , "address", "Address"
        .Add , "City", "City"
'       .Add , "SubGroupID", "Account Sub Group"
        '.Add , "Active", "Active"
        '.Add , "Editable", "Editable"
      End With
  End With
  
  Dim li As ListItem
  Dim lis As ListSubItem
  Do Until .EOF
    
    Set li = refLvw.ListItems.Add(, "AccountID=" & !ID, "" & SoftVal(!AccountName))
    li.Tag = "" & SoftVal(!ID)
    Set lis = li.ListSubItems.Add(, "AID", "" & SoftVal(!ID))
    Set lis = li.ListSubItems.Add(, "GID", "" & SoftVal(!AGDesc))
    Set lis = li.ListSubItems.Add(, "ContactName", "" & SoftVal(!ContactName))
    Set lis = li.ListSubItems.Add(, "MobileNumber", "" & SoftVal(!MobileNumber))
    Set lis = li.ListSubItems.Add(, "PhoneNumber1", "" & SoftVal(!PhoneNumber1))
'    Set lis = li.ListSubItems.Add(, "address", "" & SoftVal(!Address))
    Set lis = li.ListSubItems.Add(, "City", "" & SoftVal(!City))
    'Set lis = li.ListSubItems.Add(, "E", Map2YN(SoftVal(!Editable)))
    .MoveNext
   Loop
   If refLvw.ListItems.Count > 0 Then Set refLvw.SelectedItem = refLvw.ListItems.Item(1)
   .Close
End With
End Sub

Public Sub Add2ListView(refLvw As ListView, Optional sGDesc As String = "", Optional sSGDesc As String = "")
    Dim li As ListItem
    Dim lis As ListSubItem
    
    Set li = refLvw.ListItems.Add(, "AccountID=" & m_lAccountID, "" & SoftVal(m_sAccountName))
    li.Tag = "" & SoftVal(m_lAccountID)
    Set lis = li.ListSubItems.Add(, "AID", "" & SoftVal(m_lAccountID))
    Set lis = li.ListSubItems.Add(, "GID", "" & SoftVal(sGDesc))
    Set lis = li.ListSubItems.Add(, "ContactName", "" & SoftVal(m_sContactName))
    Set lis = li.ListSubItems.Add(, "MobileNumber", "" & SoftVal(m_sMobileNumber))
    Set lis = li.ListSubItems.Add(, "PhoneNumber1", "" & SoftVal(m_sPhoneNumber1))
'    Set lis = li.ListSubItems.Add(, "address", "" & SoftVal(!Address))
    Set lis = li.ListSubItems.Add(, "City", "" & SoftVal(m_sCity))
    If refLvw.ListItems.Count > 0 Then Set refLvw.SelectedItem = li: li.EnsureVisible
    
End Sub
Public Sub UpdateInListView(refLvw As ListView, li As ListItem, Optional sGDesc As String = "", Optional sSGDesc As String = "")
'    Dim li As ListItem
    Dim lis As ListSubItem
'    Set li = refLvw.FindItem("" & m_lAccountID, lvwText, , lvwWhole)
    ' Set li = refLvw.ListItems.Add(, "AccountID=" & m_lAccountID, m_lAccountID)
    If li Is Nothing Then MsgBox "Unexpected error !", vbExclamation: Exit Sub
    ' li.ListSubItems.Clear
    li.Text = m_sAccountName
    li.ListSubItems.Item("ContactName").Text = m_sContactName
    li.ListSubItems.Item("MobileNumber").Text = m_sMobileNumber
    li.ListSubItems.Item("PhoneNumber1").Text = m_sPhoneNumber1
    li.ListSubItems.Item("City").Text = m_sCity
    li.ListSubItems.Item("GID").Text = sGDesc
'   li.ListSubItems.Item("SGID").Text = sSGDesc
    'li.ListSubItems.Item("A").Text = Map2YN(m_bActive)
    'li.ListSubItems.Item("E").Text = Map2YN(m_bEditable)
    If refLvw.ListItems.Count > 0 Then Set refLvw.SelectedItem = li: li.EnsureVisible
End Sub
Public Sub RemoveFromListView(refLvw As ListView, li As ListItem)
'    Dim li As ListItem
    Dim lis As ListSubItem
'    Set li = refLvw.FindItem("" & m_lAccountID, lvwText, , lvwWhole)
    If li Is Nothing Then MsgBox "Unexpected error !", vbExclamation: Exit Sub
    refLvw.ListItems.Remove li.Key
End Sub
Public Property Get Active() As Boolean
Active = m_bActive
End Property

Public Property Let Active(ByVal bNewValue As Boolean)
m_bActive = bNewValue
End Property

Public Sub UpdateFormCombo(ByRef m_objCmb As ComboBox, Optional lAcctGroup As Long = 0, Optional lAcctSubGroup As Long = 0, Optional sCrit As String = "")
    Dim a As Integer
    Dim s As String
    a = 0
    If lAcctGroup <> 0 And lAcctSubGroup = 0 Then
        s = "GroupID=" & lAcctGroup
    ElseIf lAcctSubGroup <> 0 And lAcctSubGroup = 0 Then
        s = "SubGroupID=" & lAcctGroup
    ElseIf lAcctGroup <> 0 And lAcctSubGroup <> 0 Then
        s = "GroupID=" & lAcctGroup & " AND " & "SubGroupID=" & lAcctSubGroup
    ElseIf sCrit <> "" Then
        s = sCrit
    Else
        s = ""
    End If
    OpenRs s
    m_objCmb.Clear
    Do While Not m_rsAccts.EOF
        m_objCmb.AddItem m_rsAccts("AccountName") & ", " & m_rsAccts("ContactName") & ", " & m_rsAccts("City")
        m_objCmb.ItemData(a) = m_rsAccts("ID")
        a = a + 1
        m_rsAccts.MoveNext
    Loop
End Sub

Sub ValidateData()
   Dim bValid As Boolean
   bValid = True
   bValid = bValid And m_sAccountName <> ""
   'bValid = bValid And m_sContactName <> ""
   'bValid = bValid And m_sContactTitle <> ""
   'bValid = bValid And m_iCreditDays <> 0
   'bValid = bValid And m_dTradeDiscount <> 0
'   bValid = bValid And m_sPostalCode <> ""
'   bValid = bValid And m_sCountry <> ""
'   bValid = bValid And m_sPhoneNumber <> ""
   'bValid = bValid And m_sFaxNumber <> ""
   m_bIsDataValid = bValid
'   RaiseEvent DataValidated(bValid)
End Sub

Public Function Balance(AccountNo As Long, BalanceDate As Date) As Double
    Dim rsAccountBalance As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "SELECT Sum(Debit) as TotDebit,Sum(Credit) as TotCredit "
    strSQL = strSQL + "FROM Trans INNER JOIN TransDet ON (Trans.TransID = TransDet.TransID) AND (Trans.TransType = TransDet.TransType) "
    strSQL = strSQL + "WHERE TRANS.TRANSDATE <= '" & Format(BalanceDate, "mm/dd/yyyy") & "' AND TransDet.AccountNumber =" & AccountNo
    'Debug.Print strSQL
    rsAccountBalance.Open strSQL, m_objConnectDB.cnnMyshop
    On Error Resume Next
    rsAccountBalance.MoveFirst
    On Error GoTo 0
    If rsAccountBalance.EOF Then
       Balance = 0
    Else
       Balance = NumVal(SoftVal(rsAccountBalance!TotDebit)) - NumVal(SoftVal(rsAccountBalance!TotCredit))
       If Balance > 0 Then
            m_sbalancetype = "Dr"
       Else
            m_sbalancetype = "Cr"
       End If
            Balance = Abs(Balance)
    End If
    rsAccountBalance.Close
    Set rsAccountBalance = Nothing
End Function

Function CanDelete(lID As Long) As Boolean
Dim sSQL As String
sSQL = "SELECT COUNT(*) AS TotRecs FROM TransDet WHERE AccountNumber=" & lID
Dim rsTmp As ADODB.Recordset
Set rsTmp = m_objConnectDB.cnnMyshop.Execute(sSQL)
CanDelete = IsNull2(rsTmp!TotRecs, 0) = 0
End Function


