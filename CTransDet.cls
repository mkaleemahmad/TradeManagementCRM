VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTransDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CTransDet

Option Explicit

Private m_objAccount As cAccounts
Private m_lID As Long
Private m_lAccountID As Long
Private m_sTransType As String
Private m_sDescription As String
Private m_dDebit As Double
Private m_dCredit As Double
'Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private m_rsTransDet As ADODB.Recordset
'Private m_objConnectDB As New ConnectDB
Public Event DataValidated(bValid As Boolean)

Public Property Get AccountID() As Variant
    AccountID = m_lAccountID
End Property

Public Property Let AccountID(ByVal vNewValue As Variant)
    m_lAccountID = vNewValue
    ValidateData
End Property

'Public Property Get AccountID() As Long
'    AccountID = m_lAccountID
'End Property
'
'Public Property Let AccountID(ByVal vData As Long)
'    m_lAccountID = AccountID
'End Property

Public Property Let ID(ByVal vData As Long)
    m_lID = vData
    ValidateData
End Property

Public Property Get ID() As Long
    ID = m_lID
End Property

Public Property Let TransType(ByVal vData As String)
    m_sTransType = vData
    ValidateData
End Property
Public Property Get TransType() As String
    TransType = m_sTransType
End Property

Public Property Let Description(ByVal vData As String)
    m_sDescription = vData
    ValidateData
End Property
Public Property Get Description() As String
    Description = m_sDescription
End Property

Public Property Let Debit(ByVal vData As Double)
    m_dDebit = vData
    ValidateData
End Property
Public Property Get Debit() As Double
    Debit = m_dDebit
End Property

Public Property Let Credit(ByVal vData As Double)
    m_dCredit = vData
    ValidateData
End Property
Public Property Get Credit() As Double
    Credit = m_dCredit
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
   
   bValid = bValid And m_lID <> 0
   bValid = bValid And m_sTransType <> ""
   bValid = bValid And (m_dDebit <> 0 Or m_dCredit <> 0)
      
   m_bIsDataValid = bValid
   RaiseEvent DataValidated(bValid)
End Sub

Private Sub Class_Initialize()
    Set m_rsTransDet = New ADODB.Recordset
    ValidateData
End Sub

'Public Sub Init(Optional lID As Long = 0, Optional sTType As String = "", Optional objRs As ADODB.Recordset)
'   'OpenRs
'   'Dim qdData As QueryDef
'   'Set m_DB = objDB
'   ' If the ID is missing, the object
'   ' is being instantiated to be filled
'   ' with new data.
'   If lID = 0 Then
'  '    m_bIsNewRecord = True
'      Exit Sub
'      Else
'        If Not objRs Is Nothing Then
'            m_lID = objRs("TransID") & ""
'            m_sTransType = objRs("TransType") & ""
'            m_sDescription = objRs("Description") & ""
'            m_dDebit = objRs("Debit") & ""
'            m_dCredit = objRs("Credit") & ""
'        End If
'   End If
'
'End Sub

Public Sub Save(IsNewRecord As Boolean, lID As Long)

   If IsNewRecord Then
    OpenRs
    m_rsTransDet.AddNew
   End If
 
   m_rsTransDet("TransID") = lID
   m_rsTransDet("TransType") = m_sTransType
   m_rsTransDet("AccountNumber") = m_lAccountID
   m_rsTransDet("Description") = Left(m_sDescription, 50)
   m_rsTransDet("Debit") = m_dDebit
   m_rsTransDet("Credit") = m_dCredit
   m_rsTransDet.Update
   IsNewRecord = False
End Sub

Public Sub OpenRs()
    m_objConnectDB.Connect
    Set m_rsTransDet = New ADODB.Recordset
    Dim strRecordset As String
    strRecordset = "Select * from TransDet"
    If m_rsTransDet.State <> adStateOpen Then
        m_rsTransDet.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub
Public Sub DeleteRec(ByVal lVoucherID As Long, ByVal strVoucherType As String)
    OpenRs
    Dim cmdUpdateInStock As ADODB.Command
    Set cmdUpdateInStock = New ADODB.Command
    cmdUpdateInStock.CommandType = adCmdText
    cmdUpdateInStock.CommandText = "DELETE FROM TransDet WHERE (((TransDet.TransType) = " & "'" & strVoucherType & "'" & ") AND ((TransDet.TransID) = " & lVoucherID & "))"
    cmdUpdateInStock.ActiveConnection = m_objConnectDB.cnnMyshop
    cmdUpdateInStock.Execute
End Sub

Public Function FindRec(lPTransID As Long, strTransType As String) As Boolean
    OpenRs
   m_rsTransDet.MoveFirst
    Do While Not m_rsTransDet.EOF
        If m_rsTransDet("transID") = lPTransID And m_rsTransDet("TransType") = strTransType Then
            FindRec = True
            Exit Function
        Else
        m_rsTransDet.MoveNext
        End If
    Loop
    FindRec = False
    
End Function

