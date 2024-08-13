VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_rsCA As ADODB.Recordset
Dim m_rsA As ADODB.Recordset

Public Property Get AccountNo(AccountDesc As String) As Long
If m_rsCA.EOF Then
  AccountNo = -1
  Exit Sub
End If
m_rsCA.MoveFirst
m_rsCA.Find "AccountDesc='" & AccountDesc & "'"
If m_rsCA.EOF Then
   AccountNo = -1
Else
   AccountNo = IIf(IsNull(m_rsCA!AccountNumber), -1, m_rsCA!AccountNumber)
End If
End Property

Public Property Let AccountNo(AccountDesc As String, ByVal NewAccountNo As Long)
m_rsCA.MoveFirst
m_rsCA.Find "AccountDesc='" & AccountDesc & "'"
If m_rsCA.EOF Then
   m_rsCA.AddNew Array("AccountDesc", "AccountNumber"), Array(AccountDesc, NewAccountNo)
   m_rsCA.Update
Else
  m_rsCA.Move 0
   m_rsCA!AccountNumber = NewAccountNo
   m_rsCA.Update
End If
End Property

Public Property Get AccountName(AccountDesc As String) As String
If m_rsCA.EOF Then
  GoTo lblNoRecord
End If
m_rsCA.MoveFirst
m_rsCA.Find "AccountDesc='" & AccountDesc & "'"
lblNoRecord:
If m_rsCA.EOF Then
   m_rsCA.AddNew Array("AccountDesc", "AccountNumber"), Array(AccountDesc, -1)
   m_rsCA.Update
   m_rsCA.Requery
End If
If Not m_rsA.EOF Then
m_rsA.MoveFirst
m_rsA.Find "AccountNumber=" & m_rsCA!AccountNumber
End If
If m_rsA.EOF Then
   AccountName = ""
Else
   AccountName = m_rsA!AccountName
End If
End Property


Public Sub ChangeAccountDesc(OldDesc As String, NewDesc As String)
m_rsCA.MoveFirst
m_rsCA.Find "AccountDesc='" & OldDesc & "'"
If m_rsCA.EOF Then
   m_rsCA.AddNew Array("AccountDesc"), Array(OldDesc)
   m_rsCA.Update
   m_rsCA.Requery
Else
   m_rsCA!AccountDesc = NewDesc
   m_rsCA.Update
End If
End Sub

Public Sub Initialize()
Set m_rsCA = New ADODB.Recordset
Set m_rsA = New ADODB.Recordset
Dim sSQL As String
sSQL = "Select ControlAccounts.* From ControlAccounts "
m_rsCA.Open sSQL, m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
sSQL = "Select ID as AccountNumber,AccountName From Accounts"
m_rsA.Open sSQL, m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
On Error Resume Next
 If Not m_rsCA.EOF And Not m_rsCA.BOF Then
    m_rsCA.MoveNext
    m_rsCA.MovePrevious
    m_rsA.MoveNext
    m_rsA.MovePrevious
  End If
End Sub

Public Sub Purge()
m_rsCA.Close
Set m_rsCA = Nothing
m_rsA.Close
Set m_rsA = Nothing
End Sub

Public Function ExistInTransDet(lAccountID As Long) As Boolean
Dim rsAccts As New ADODB.Recordset
Dim sSQL As String
sSQL = "Select Top 1 * From Transdet Where AccountNumber=" & lAccountID
rsAccts.Open sSQL, m_objConnectDB.cnnMyshop
If rsAccts.EOF Then
  ExistInTransDet = False
ElseIf IsNull(rsAccts!AccountNumber) Then
  ExistInTransDet = False
Else
  ExistInTransDet = True
End If
rsAccts.Close
Set rsAccts = Nothing
End Function
