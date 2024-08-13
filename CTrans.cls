VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CTrans

Option Explicit
Private m_lID As Long
Private m_lLastID As Long
Private m_sTransType As String
Private m_dtTDate As Date
Private m_dtEDate As Date
Private m_lSalesManID As Long
Private m_sDescription As String
Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private m_sRefNo As String
Private m_bValidRecord As Boolean
Private rsTrans As ADODB.Recordset
Dim m_objCTransDet As CTransDet
Public Event DataValidated(bValid As Boolean)

Public Property Let ID(ByVal vData As Long)
    m_lID = vData
    ValidateData
End Property
Public Property Get ID() As Long
    ID = m_lID
End Property
Public Property Get ValidRecord() As Boolean
    ValidRecord = m_bValidRecord
End Property
Public Property Let RefNo(ByVal vData As String)
    m_sRefNo = vData
    ValidateData
End Property
Public Property Get RefNo() As String
    RefNo = m_sRefNo
End Property


Public Property Let LastID(ByVal vData As Long)
    m_lLastID = vData
    ValidateData
End Property
Public Property Get LastID() As Long
    LastID = m_lLastID
End Property

Public Property Let TransType(ByVal vData As String)
    m_sTransType = vData
    ValidateData
End Property
Public Property Get TransType() As String
    TransType = m_sTransType
End Property

Public Property Let TDate(ByVal vData As Date)
    m_dtTDate = vData
    ValidateData
End Property
Public Property Get TDate() As Date
    TDate = m_dtTDate
End Property

'Public Property Let EDate(ByVal vData As Date)
'    m_dtEDate = vData
'    ValidateData
'End Property
Public Property Get EDate() As Date
    EDate = m_dtEDate
End Property

Public Property Let SalesManID(ByVal vData As Long)
    m_lSalesManID = vData
    ValidateData
End Property
Public Property Get SalesManID() As Long
    SalesManID = m_lSalesManID
End Property

Public Property Let Description(ByVal vData As String)
    m_sDescription = vData
    ValidateData
End Property
Public Property Get Description() As String
    Description = m_sDescription
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
 '  bValid = bValid And m_dtPTDate <> ""
   
   m_bIsDataValid = bValid
   RaiseEvent DataValidated(bValid)
End Sub

Private Sub Class_Initialize()
   ValidateData
End Sub

Public Sub Init(Optional lID As Long = 0, Optional sTType As String = "", Optional dDate As Variant)
   If IsMissing(dDate) Then
      OpenRs lID, sTType
   Else
      OpenRs lID, sTType, Date
   End If
   If lID = 0 Then
      m_bIsNewRecord = True
      Exit Sub
'      Else
'      rsTrans.Find "TransID=" & lID & " AND " & "TransType=" & "'" & sTType & "'"
   End If
   If rsTrans.BOF = True And rsTrans.EOF = True Then
    m_lID = 0
     m_bValidRecord = False
   Else
    m_lID = rsTrans("TransID") & ""
    m_sTransType = rsTrans("TransType") & ""
    m_dtTDate = rsTrans("TransDate") & ""
    'm_dtEDate = rsTrans("EntryDate") & ""
 '   m_lSalesManID = rsTrans("SalesManID") & ""
    m_sDescription = rsTrans("Description") & ""
     m_sRefNo = IsNull2(rsTrans("RefNo"), "")
     m_bValidRecord = True
   End If
End Sub

Public Sub Save(IsNewRecord As Boolean)
   If IsNewRecord Then
    OpenRs
    rsTrans.AddNew
   End If
   AddToEvntLg IIf(IsNewRecord, "New", "Edit"), m_sTransType, m_lID
   rsTrans("TransID") = m_lID
   rsTrans("TransType") = m_sTransType
   rsTrans("TransDate") = m_dtTDate
'   rsTrans("EntryDate") = Date
'   rsTrans("SalesManID") = m_lSalesManID
   rsTrans("Description") = m_sDescription
   rsTrans("RefNo") = m_sRefNo
   m_bIsNewRecord = False
   rsTrans.Update
End Sub

Public Sub OpenRs(Optional lTID As Long = 0, Optional mvTransType As Variant = "", Optional dDate As Variant)
    Set rsTrans = New ADODB.Recordset
    Dim strRecordset As String
    Dim sSQL1 As String
    If IsMissing(dDate) Then
        sSQL1 = ""
    Else
        sSQL1 = " Where TransDate='" & Format(dDate, "mm/dd/yyyy") & "' "
    End If
    Dim sSQL2 As String

    If IsMissing(dDate) Then
        sSQL2 = ""
    Else
        sSQL2 = " AND TransDate='" & Format(dDate, "mm/dd/yyyy") & "'"
    End If

    If mvTransType = "" And lTID = 0 Then
        strRecordset = "Select  * from Trans " & sSQL1 & " Order By TransID"
    ElseIf mvTransType <> "" And lTID = 0 Then
        strRecordset = "Select  * from Trans Where TransType=" & "'" & mvTransType & "' " & sSQL2 & " Order By TransID"
    Else
        strRecordset = "Select  * from Trans Where TransID=" & lTID & " AND TransType = " & "'" & mvTransType & "'" & sSQL2 & "Order By TransID"
    End If
    If rsTrans.State <> adStateOpen Then
         rsTrans.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub

Public Sub UpdateFormCombo(ByRef m_objCmb As ComboBox, m_sTransType As Variant)
    m_objCmb.Clear
    Dim a As Integer
    a = 0
    OpenRs 0, m_sTransType
    Do While Not rsTrans.EOF
       If rsTrans("TransType").value = m_sTransType Then
            m_objCmb.AddItem rsTrans("TransID")
            m_objCmb.ItemData(a) = rsTrans("TransID")
            a = a + 1
        End If
        rsTrans.MoveNext
    Loop
End Sub

Public Sub DeleteRec(ByVal lVoucherID As Long, ByVal strVoucherType As String)
'    AddToEvntLg "Delete", strVoucherType, lVoucherID
'    OpenRs
    Dim cmdUpdateInStock As ADODB.Command
    Set cmdUpdateInStock = New ADODB.Command
    cmdUpdateInStock.CommandType = adCmdText
    cmdUpdateInStock.CommandText = "DELETE FROM Trans WHERE (((Trans.TransType) = " & "'" & strVoucherType & "'" & ") AND ((TRANS.TransID) = " & lVoucherID & "))"
    cmdUpdateInStock.ActiveConnection = m_objConnectDB.cnnMyshop
    cmdUpdateInStock.Execute
    m_objConnectDB.cnnMyshop.Execute "DELETE FROM TransDet WHERE (((TransDet.TransType) = " & "'" & strVoucherType & "'" & ") AND ((TransDet.TransID) = " & lVoucherID & "))"
End Sub
Public Function FindRec(ByVal lPTransID As Long, ByVal strTransType As String) As Boolean
   OpenRs
   If Not (rsTrans.EOF And rsTrans.BOF) Then rsTrans.MoveFirst
    Do While Not rsTrans.EOF
        If rsTrans("transID") = lPTransID And rsTrans("TransType") = strTransType Then
            FindRec = True
            Exit Function
        Else
        rsTrans.MoveNext
        End If
    Loop
    FindRec = False
End Function

Public Function LastTrans(ByVal strTransType As String)
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "SELECT MAX([TransID]) as LASTID FROM Trans WHERE (((Trans.TransType) = " & "'" & strTransType & "'" & "))"
    rsTemp.Open strSQL, m_objConnectDB.cnnMyshop, adOpenStatic
'    cmdLastTrans.CommandType = adCmdText
'    cmdLastTrans.CommandText = "SELECT MAX(TRANS_ID) as LASTID FROM Trans WHERE ((Trans.TransType) = " & "'" & strTransType & "'" & ")"
'    cmdLastTrans.ActiveConnection = m_objConnectDB.cnnMyshop
'    cmdLastTrans.Execute
    'If rsTemp.BOF = True And rsTemp.EOF = True Then
    'If rsTemp("LastID").Value =  Then
    If IsNull(rsTemp("LastID").value) Then
        LastID = 0
    Else
        LastID = rsTemp("LastID").value
    End If
End Function
