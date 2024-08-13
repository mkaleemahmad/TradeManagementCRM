VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CAccount

Option Explicit
Private m_lID As Long
Private m_sAccountID As String
Private m_sAccountName As String
Private m_dAccountBalance As Double
Private m_sBalanceType As String
Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private m_bIsActive As Boolean
Private m_rsAccount As ADODB.Recordset
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

'Public Property Let GroupID(ByVal vData As String)
'    m_sGroupID = vData
'    ValidateData
'End Property
'Public Property Get GroupID() As String
'    GroupID = m_sGroupID
'End Property

Public Property Let AccountID(ByVal vData As String)
    m_sAccountID = vData
    ValidateData
End Property
Public Property Get AccountID() As String
    AccountID = m_sAccountID
End Property

Public Property Let AccountName(ByVal sData As String)
    m_sAccountName = sData
    ValidateData
End Property

Public Property Get AccountName() As String
   AccountName = m_sAccountName
End Property

Public Property Get Balance() As Double
    Balance = m_dAccountBalance
End Property

Public Property Get BalanceType() As String
    BalanceType = m_sBalanceType
End Property

Public Property Get IsValid() As Boolean
   '
   ' This property is available when having
   ' an event trigger is not always helpful.
   ' It can also be used within objects that
   ' contain this one. This object can be
   ' queried as part of the other object's
   ' validation routines.
   '
   IsValid = m_bIsDataValid
End Property

Private Sub ValidateData()
   Dim bValid As Boolean
   bValid = True
   'bValid = bValid And m_sAccountID <> ""
   bValid = bValid And m_sAccountName <> ""
   
   m_bIsDataValid = bValid
   RaiseEvent DataValidated(bValid)
End Sub

Private Sub Class_Initialize()
    ValidateData
    Set m_rsAccount = New ADODB.Recordset
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
       Dim cmdAccount As ADODB.Command
       Dim prmAccountID As ADODB.Parameter
       Set cmdAccount = New ADODB.Command
       Set prmAccountID = New ADODB.Parameter
       Dim m_RS As ADODB.Recordset
       Set m_RS = New ADODB.Recordset
       Set prmAccountID = cmdAccount.CreateParameter("paramAccountID", adInteger, adParamInput, , lID)
       cmdAccount.Parameters.Append prmAccountID
       cmdAccount.ActiveConnection = m_objConnectDB.cnnMyshop
       cmdAccount.CommandType = adCmdStoredProc
       cmdAccount.CommandText = "qryGetAccount"
       Set m_RS = cmdAccount.Execute
       m_lID = m_RS("ID") & ""
       m_sAccountID = m_RS("AccountID") & ""
       m_sAccountName = m_RS("AccountName") & ""
       'm_sGroupID = m_RS("GroupID") & ""
       m_bIsActive = m_RS("Active")
       ' Clean up
       m_RS.Close
       Set m_RS = Nothing
       'Get Account Balance from query
       Dim cmdAccountBalance As ADODB.Command
       Dim rsTemp As ADODB.Recordset
       Set cmdAccountBalance = New ADODB.Command
       Set rsTemp = New ADODB.Recordset
       Set prmAccountID = cmdAccountBalance.CreateParameter("paramAccountID", adInteger, adParamInput, , lID)
       cmdAccountBalance.Parameters.Append prmAccountID
       cmdAccountBalance.ActiveConnection = m_objConnectDB.cnnMyshop
       cmdAccountBalance.CommandText = "qryGetAccountBalance"
       cmdAccountBalance.CommandType = adCmdStoredProc
       rsTemp.CursorType = adOpenDynamic
       Set rsTemp = cmdAccountBalance.Execute
       If rsTemp.BOF Or rsTemp.EOF Then
           m_dAccountBalance = 0
           m_sBalanceType = ""
       Else
        If (rsTemp("SumOfDebit").Value - rsTemp("SumOfCredit").Value) > 0 Then
           m_dAccountBalance = rsTemp("SumOfDebit").Value - rsTemp("SumOfCredit").Value
           m_sBalanceType = "Debit"
        Else
           m_dAccountBalance = rsTemp("SumOfCredit").Value - rsTemp("SumOfDebit").Value
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
      m_rsAccount.MoveLast
      m_rsAccount.AddNew
   Else
      m_rsAccount.MoveFirst
      m_rsAccount.Find "ID=" & m_lID 'To Be replaced with better code.
   End If
       m_rsAccount("AccountID") = m_sAccountID
       m_rsAccount("AccountName") = m_sAccountName
       'm_rsAccount("GroupID") = m_sGroupID
       m_rsAccount("Active") = m_bIsActive
       m_rsAccount.Update
       m_rsAccount.Requery
       m_lID = m_rsAccount("ID")
       m_bIsNewRecord = False
End Sub

Public Sub OpenRs()
    Dim strRecordset As String
    strRecordset = "Select * from Accounts order by AccountName"
    If m_rsAccount.State <> adStateOpen Then
        m_rsAccount.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub
Public Sub UpdateFormCombo(ByRef m_objCmb As ComboBox)
    Dim a As Integer
    m_objCmb.Clear
    a = 0
    OpenRs
    m_rsAccount.Filter = "Active = 1"
    m_rsAccount.MoveFirst
    m_objCmb.Clear
    Do While Not m_rsAccount.EOF
        m_objCmb.AddItem m_rsAccount("AccountName")
        m_objCmb.ItemData(a) = m_rsAccount("ID")
        a = a + 1
        m_rsAccount.MoveNext
    Loop
    m_rsAccount.Filter = ""
End Sub

