VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_lAccountID As Long
Dim m_sAccountName As String
Dim m_lDebit As Long
Dim m_lCredit As Long
Dim m_lGroupID As Long
Dim m_bEditable As Boolean
Dim m_bIsGroup As Boolean
Dim m_lStartAccountID As Long
Dim m_lEndAccountID As Long
Dim m_sbalancetype As String

Dim m_bAddingNew As Boolean
Dim m_bEditing As Boolean

Dim m_rsAccounts As New ADODB.Recordset
 
'Dim m_objconnectdb As New ConnectDB
Dim m_bCanSave As Boolean

' AccountID
Public Property Get AccountID() As Long
AccountID = m_lAccountID
End Property

Public Property Let AccountID(ByVal NewAccountID As Long)
m_lAccountID = NewAccountID
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
' IsGroup
Public Property Get IsGroup() As Boolean
IsGroup = m_bIsGroup
End Property

Public Property Let IsGroup(ByVal WhetherGroup As Boolean)
m_bIsGroup = WhetherGroup
End Property
' From Account ID
Public Property Get FromAccountID() As Long
FromAccountID = m_lStartAccountID
End Property

Public Property Let FromAccountID(ByVal AccountID As Long)
m_lStartAccountID = AccountID
End Property
' To Account ID
Public Property Get ToAccountID() As Long
ToAccountID = m_lEndAccountID
End Property

Public Property Let ToAccountID(ByVal AccountID As Long)
m_lEndAccountID = AccountID
End Property

Public Property Get BalanceType() As String
    BalanceType = m_sbalancetype
End Property


Public Sub AddNew()
m_lAccountID = 0
m_sAccountName = ""
m_lDebit = 0
m_lCredit = 0
m_lGroupID = 0
m_bEditable = True
m_bIsGroup = True
m_lStartAccountID = 0
m_lEndAccountID = 0

m_rsAccounts.Open "Select * From Accounts", m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
m_rsAccounts.AddNew
m_bCanSave = True
m_bAddingNew = True
m_bEditing = False
End Sub

Public Sub Edit(FieldArray, ValueArray)
If Not IsArray(FieldArray) Then
   Err.Raise vbObjectError + 1, "Function: Edit", "Parameter FieldArray is not a valid array."
   Exit Sub
End If
If Not IsArray(ValueArray) Then
   Err.Raise vbObjectError + 1, "Function: Edit", "Parameter ValueArray is not a valid array."
   Exit Sub
End If

Dim sCrit As String
Dim I As Long
sCrit = ""
For I = LBound(FieldArray) To UBound(ValueArray)
   If sCrit <> "" Then sCrit = sCrit & " AND "
   sCrit = sCrit & " " & FieldArray(I) & "=" & ValueArray(I) & " "
Next
With m_rsAccounts
 .Open "Select * From Accounts Where " & sCrit, m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
 If .EOF Or .BOF Then
   .Close
   m_bCanSave = False
   Exit Sub
End If

 m_lAccountID = !ID
 m_sAccountName = !AccountName
 m_lDebit = IIf(IsNull(!Debit), 0, !Debit)
 m_lCredit = IIf(IsNull(!Credit), 0, !Credit)
 m_lGroupID = !GroupID
 m_bEditable = IIf(IsNull(!Editable), 0, !Editable)
 m_bIsGroup = IIf(IsNull(!IsGroup), 0, !IsGroup)
 m_lStartAccountID = IIf(IsNull(!FromAccountID), 0, !FromAccountID)
 m_lEndAccountID = IIf(IsNull(!ToAccountID), 0, !ToAccountID)
End With
m_bCanSave = True
m_bAddingNew = False
m_bEditing = True
End Sub

Public Sub Save()
If Not m_bCanSave Then
   Err.Raise vbObjectError + 1, "Sub: Save", "Can not save"
End If

With m_rsAccounts
'   !AccountID = m_lAccountID
   !AccountName = m_sAccountName
   !Debit = m_lDebit
   !Credit = m_lCredit
   !GroupID = m_lGroupID
   !Editable = m_bEditable
   !IsGroup = m_bIsGroup
   !FromAccountID = m_lStartAccountID
   !ToAccountID = m_lEndAccountID
   .Update
   .Close
End With
m_bCanSave = False
m_bAddingNew = False
m_bEditing = False
End Sub

Public Sub Delete()
With m_rsAccounts
   .Delete
   .Close
End With
m_bCanSave = False
m_bAddingNew = False
m_bEditing = False
End Sub

Private Sub Class_Initialize()

m_bCanSave = False
End Sub

Private Sub Class_Terminate()
'm_objconnectdb.Disconnect
End Sub

Public Sub Cancel()
m_rsAccounts.Close
m_bCanSave = False
m_bAddingNew = False
m_bEditing = False
End Sub

Public Property Get CanSave() As Boolean
CanSave = m_bCanSave
End Property

Public Function GetAccountName(AccountNo As Long) As String
Dim rsA As New ADODB.Recordset
rsA.Open "Select AccountName From Accounts Where ID=" & AccountNo, m_objConnectDB.cnnMyshop
If rsA.EOF Then
   GetAccountName = ""
Else
   GetAccountName = rsA!AccountName
End If
End Function

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
            m_sbalancetype = "Debit"
       Else
            m_sbalancetype = "Credit"
       End If
    End If
    rsAccountBalance.Close
    Set rsAccountBalance = Nothing
End Function

Public Function NextAccountID(FromAccountID As Long, ToAccountID As Long) As Long
Dim rs As New ADODB.Recordset
Dim sSQL As String
sSQL = "Select MAX(ID) as NAID From Accounts Where ID>" & FromAccountID & " and ID<" & ToAccountID & ";"
rs.Open sSQL, m_objConnectDB.cnnMyshop
If IsNull(rs!NAID) Then
  NextAccountID = FromAccountID
Else
  NextAccountID = SoftVal(rs!NAID) + 1
End If
End Function
