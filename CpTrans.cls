VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CPTrans
Option Explicit
Private m_lID As Long
Private m_sPTransType As String
Private m_dtPTDate As String
Private m_dtPTDateTime As Date
Private m_iSenderReceiverID As Integer
Private m_sSalesManID As Double
Private m_sShippingInfo As String
Private m_sRefNo As String
Private m_dFreight As Double
Private m_dDiscountPCent As Double
Private m_dDiscountValue As Double
Private m_sDescription As String
Private m_bIsCash As Boolean
Private m_dCashReceived As Double
Private m_dPrevBal As Double
Private m_sPrevBalType As String
Private m_lCashTransID As Long
Private m_sCashTransType As String
Private m_bNoRecords As Boolean
Private m_objTrans As CTrans
Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private rsPTrans As ADODB.Recordset
'Dim clsConnectDB As ConnectDB
Public Event DataValidated(bValid As Boolean)

Public Property Let ID(ByVal vData As Long)
    m_lID = vData
    ValidateData
End Property
Public Property Get ID() As Long
    ID = m_lID
End Property
Public Property Let PTransType(ByVal vData As String)
    m_sPTransType = vData
    ValidateData
End Property
Public Property Get PTransType() As String
    PTransType = m_sPTransType
End Property

Public Property Let PTDate(ByVal vData As Date)
    m_dtPTDate = vData
    ValidateData
End Property
Public Property Get PTDate() As Date
    PTDate = m_dtPTDate
End Property
Public Property Let PTDateTime(ByVal vData As Date)
    m_dtPTDateTime = vData
    ValidateData
End Property
Public Property Get PTDateTime() As Date
    PTDateTime = m_dtPTDateTime
End Property

Public Property Let SenderReceiverID(ByVal vData As Integer)
    m_iSenderReceiverID = vData
    ValidateData
End Property
Public Property Get SenderReceiverID() As Integer
    SenderReceiverID = m_iSenderReceiverID
End Property

Public Property Let SalesManID(ByVal vData As String)
    m_sSalesManID = vData
    ValidateData
End Property
Public Property Get SalesManID() As String
    SalesManID = m_sSalesManID
End Property

Public Property Let ShippingInfo(ByVal vData As String)
    m_sShippingInfo = vData
    ValidateData
End Property
Public Property Get ShippingInfo() As String
    ShippingInfo = m_sShippingInfo
End Property

Public Property Let RefNo(ByVal vData As String)
    m_sRefNo = vData
    ValidateData
End Property
Public Property Get RefNo() As String
    RefNo = m_sRefNo
End Property

Public Property Let Freight(ByVal vData As Double)
    m_dFreight = vData
    ValidateData
End Property
Public Property Get Freight() As Double
    Freight = m_dFreight
End Property

Public Property Let DiscountPCent(ByVal vData As Double)
    m_dDiscountPCent = vData
    ValidateData
End Property
Public Property Get DiscountPCent() As Double
    DiscountPCent = m_dDiscountPCent
End Property

Public Property Let DiscountValue(ByVal vData As Double)
    m_dDiscountValue = vData
    ValidateData
End Property
Public Property Get DiscountValue() As Double
    DiscountValue = m_dDiscountValue
End Property

Public Property Let Description(ByVal vData As String)
    m_sDescription = vData
    ValidateData
End Property
Public Property Get Description() As String
    Description = m_sDescription
End Property

Public Property Let IsCash(ByVal vData As Boolean)
    m_bIsCash = vData
    ValidateData
End Property
Public Property Get IsCash() As Boolean
    IsCash = m_bIsCash
End Property

Public Property Let CashReceived(ByVal vData As Double)
    m_dCashReceived = vData
    ValidateData
End Property
Public Property Get CashReceived() As Double
    CashReceived = m_dCashReceived
End Property


Public Property Let PrevBal(ByVal vData As Double)
    m_dPrevBal = vData
    ValidateData
End Property
Public Property Get PrevBal() As Double
    PrevBal = m_dPrevBal
End Property
Public Property Let PrevBalType(ByVal vData As String)
    m_sPrevBalType = vData
    ValidateData
End Property
Public Property Get PrevBalType() As String
    PrevBalType = m_sPrevBalType
End Property

Public Property Let CashTransID(ByVal vData As Long)
    m_lCashTransID = vData
    ValidateData
End Property
Public Property Get CashTransID() As Long
    CashTransID = m_lCashTransID
End Property
Public Property Let CashTransType(ByVal vData As String)
    m_sCashTransType = vData
    ValidateData
End Property
Public Property Get CashTransType() As String
    CashTransType = m_sCashTransType
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
   bValid = bValid And m_sPTransType <> ""
 '  bValid = bValid And m_dtPTDate <> ""
   
   m_bIsDataValid = bValid
   RaiseEvent DataValidated(bValid)
End Sub

Private Sub Class_Initialize()
   Set m_objTrans = New CTrans
   ValidateData
End Sub


Public Sub Init(Optional lID As Long = 0, Optional sPTType As String = "")
   If lID = 0 Then
      m_bIsNewRecord = True
      OpenRs
      Exit Sub
   Else
      OpenRs "Where PTransID=" & lID & " AND " & "PTransType='" & sPTType & "'"
   End If
   If Not rsPTrans.BOF And Not rsPTrans.EOF Then
    m_bNoRecords = False
    m_lID = rsPTrans("PTransID") & ""
    m_sPTransType = rsPTrans("PTransType") & ""
    m_dtPTDate = rsPTrans("PTDate") & ""
    m_iSenderReceiverID = rsPTrans("SenderReceiverID") & ""
    m_sSalesManID = rsPTrans("SaleManID") & ""
    m_sShippingInfo = rsPTrans("ShippingInfo") & ""
    m_sRefNo = rsPTrans("RefNo") & ""
    m_dFreight = rsPTrans("Freight") & ""
    m_dDiscountPCent = rsPTrans("DiscountPCent") & ""
    m_dDiscountValue = rsPTrans("DiscountValue") & ""
    m_sDescription = rsPTrans("Description") & ""
    m_bIsCash = rsPTrans("IsCash") & ""
    m_dCashReceived = rsPTrans("CashReceived") & ""
    m_dtPTDateTime = IsNull2(rsPTrans("PTDateTime"), Now)
    m_sPrevBalType = rsPTrans("PrevBalType") & ""
    m_dPrevBal = rsPTrans("PrevBal") & ""
    m_lCashTransID = rsPTrans("CashTransID") & ""
    m_sCashTransType = rsPTrans("CashTransType") & ""
    
    If sPTType <> "OS" And sPTType <> "PD" Then
     m_objTrans.Init lID, sPTType
    End If
   Else
    m_bNoRecords = True
   End If
End Sub

Public Sub Save(IsNewRecord As Boolean)
   '
   If IsNewRecord Then
    OpenRs
    rsPTrans.AddNew
   End If
   AddToEvntLg IIf(IsNewRecord, "New", "Edit"), m_sPTransType, m_lID
   rsPTrans("PTransID") = m_lID
   rsPTrans("PTransType") = m_sPTransType
   rsPTrans("PTDate") = m_dtPTDate
   rsPTrans("SenderReceiverID") = m_iSenderReceiverID
   rsPTrans("SaleManID") = m_sSalesManID
   rsPTrans("ShippingInfo") = m_sShippingInfo
   rsPTrans("Description") = m_sDescription
   rsPTrans("RefNo") = Left(m_sRefNo, 10)
   rsPTrans("Freight") = m_dFreight
   rsPTrans("DiscountPCent") = m_dDiscountPCent
   rsPTrans("DiscountValue") = m_dDiscountValue
   rsPTrans("IsCash") = m_bIsCash
   rsPTrans("PTDateTime") = m_dtPTDateTime
   rsPTrans("CashReceived") = m_dCashReceived
   rsPTrans("PrevBalType") = m_sPrevBalType
   rsPTrans("PrevBal") = m_dPrevBal
   rsPTrans("CashTransID") = m_lCashTransID
   rsPTrans("CashTransType") = m_sCashTransType
   
   If m_bIsNewRecord Then
    rsPTrans("PTDateTime") = CVDate(IsNull2(m_dtPTDate, Date) & " " & Time) 'Now
    Else
    rsPTrans("PTDateTime") = CVDate(IsNull2(m_dtPTDate, Date) & " " & TimeValue(m_dtPTDateTime)) 'Now
   End If
   m_bIsNewRecord = False
   rsPTrans.Update
   If m_sPTransType <> "OS" And m_sPTransType <> "PD" Then
    m_objTrans.ID = m_lID
    m_objTrans.TransType = m_sPTransType
    m_objTrans.TDate = m_dtPTDate
    m_objTrans.Description = m_sDescription
    m_objTrans.Save IsNewRecord
   End If
End Sub

Public Sub OpenRs(Optional sCriteria As String = "")
    'Set clsConnectDB = New ConnectDB
    
    Set rsPTrans = New ADODB.Recordset
    Dim strRecordset As String
    If sCriteria <> "" Then
        strRecordset = "Select * from ProdTrans " & sCriteria
    Else
        strRecordset = "Select * from ProdTrans "
    End If
    If rsPTrans.State <> adStateOpen Then
    rsPTrans.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub
Public Sub DeleteRec(ByVal lPTransID As Long, ByVal strTransType As String)
'    AddToEvntLg "Delete", strTransType, lPTransID
'    OpenRs
    Dim cmdUpdateInStock As ADODB.Command
    Set cmdUpdateInStock = New ADODB.Command
    cmdUpdateInStock.CommandType = adCmdText
    cmdUpdateInStock.CommandText = "DELETE FROM ProdTrans WHERE (((ProdTrans.PTransType) = " & "'" & strTransType & "'" & ") AND ((ProdTrans.PTransID) = " & lPTransID & "))"
    cmdUpdateInStock.ActiveConnection = m_objConnectDB.cnnMyshop
    cmdUpdateInStock.Execute
    m_objConnectDB.cnnMyshop.Execute "DELETE FROM ProdTransDet WHERE (((ProdTransDet.PTransType) = " & "'" & strTransType & "'" & ") AND ((ProdTransDet.PTransID) = " & lPTransID & "))"
    cmdUpdateInStock.ActiveConnection = m_objConnectDB.cnnMyshop
    cmdUpdateInStock.Execute
    cmdUpdateInStock.CommandText = "DELETE FROM Trans WHERE TransType = " & "'" & strTransType & "'" & " AND TransID = " & lPTransID
    cmdUpdateInStock.ActiveConnection = m_objConnectDB.cnnMyshop
    cmdUpdateInStock.Execute
    cmdUpdateInStock.CommandText = "DELETE FROM TransDet WHERE TransType = " & "'" & strTransType & "'" & " AND TransID = " & lPTransID
    cmdUpdateInStock.ActiveConnection = m_objConnectDB.cnnMyshop
    cmdUpdateInStock.Execute
End Sub

'Public Function FindRec(ByVal lPTransID As Long, ByVal strTransType As String) As Boolean
'    OpenRs
'    rsPTrans.MoveFirst
'    Do While Not rsPTrans.EOF
'        If rsPTrans("PTransID") = lPTransID And rsPTrans("PTransType") = strTransType Then
'            FindRec = True
'            Exit Function
'        Else
'            rsPTrans.MoveNext
'        End If
'    Loop
'    FindRec = False
'End Function

Public Function LastPTrans(ByVal strPTransType As String) As Double
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "SELECT MAX([PTransID]) as LASTID FROM ProdTrans WHERE (((ProdTrans.PTransType) = " & "'" & strPTransType & "'" & "))"
    rsTemp.Open strSQL, m_objConnectDB.cnnMyshop, adOpenStatic
    If IsNull(rsTemp("LastID").value) Then
        LastPTrans = 0
    Else
       LastPTrans = rsTemp("LastID").value
    End If
End Function


Public Property Get NoRecords() As Boolean
    NoRecords = m_bNoRecords
End Property

