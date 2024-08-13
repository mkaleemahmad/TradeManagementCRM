VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustomerRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CCustomerRate
Option Explicit

Private m_iCustomerID As Integer
Private m_iProductID As Integer
Private m_sPTotalDescription As String
Private m_dRate As Double
Private m_bIsDataValid As Boolean
Private m_rsCustomerRates As ADODB.Recordset
'Private m_objConnDB As ConnectDB
'Dim m_objConnectDB As New ConnectDB
Public Event DataValidated(bValid As Boolean)

Public Property Get ProductID() As Integer
    ProductID = m_iProductID
End Property
Public Property Let ProductID(ByVal vNewValue As Integer)
    m_iProductID = vNewValue
End Property

Public Property Get ProductTotalDescription() As String
    ProductTotalDescription = m_sPTotalDescription
End Property
Public Property Let ProductTotalDescription(ByVal vNewValue As String)
    m_sPTotalDescription = vNewValue
End Property

Public Property Get CustomerID() As Integer
    CustomerID = m_iCustomerID
End Property
Public Property Let CustomerID(ByVal vNewValue As Integer)
    m_iCustomerID = vNewValue
End Property

Public Property Get Rate() As Double
    Rate = m_dRate
End Property

Public Property Let Rate(ByVal vNewValue As Double)
    m_dRate = vNewValue
End Property

Public Sub OpenRs()
    Set m_rsCustomerRates = New ADODB.Recordset
    Dim strRecordset As String
    strRecordset = "Select * from CustomerRates"
    If m_rsCustomerRates.State <> adStateOpen Then
        m_rsCustomerRates.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub

Public Sub Init(iCustID As Long, iProdID As Long)
    
'    Dim m_RS As ADODB.Recordset
'
'    Dim cmdCustProdRate As ADODB.Command
'    Dim paramCustId As ADODB.Parameter
'    Dim paramProdID As ADODB.Parameter
'
'    Set m_RS = New ADODB.Recordset
'    Set cmdCustProdRate = New ADODB.Command
'    Set paramCustId = New ADODB.Parameter
'    Set paramProdID = New ADODB.Parameter
'
'    With cmdCustProdRate
'        .ActiveConnection = m_objConnectDB.cnnMyshop
'        .CommandType = adCmdStoredProc
'        .CommandText = "qryGetCustProdRate"
'    End With
'
'    Set paramCustId = cmdCustProdRate.CreateParameter("CustID", adInteger, adParamInput, , iCustID)
'    Set paramProdID = cmdCustProdRate.CreateParameter("ProdID", adInteger, adParamInput, , iProdID)
'
'    cmdCustProdRate.Parameters.Append paramCustId
'    cmdCustProdRate.Parameters.Append paramProdID
'
'    Set m_RS = cmdCustProdRate.Execute
'    If Not m_RS.BOF And Not m_RS.EOF Then
'        m_dRate = m_RS("Rate")
'    Else
'        m_dRate = 0
'    End If
'    m_RS.Close
'    Set m_RS = Nothing
'    m_iCustomerID = iCustID
'    m_iProductID = iProdID
'   ' m_objCustomer.Init iCustID
'   ' m_objProduct.Init iProdID
End Sub

Public Sub Save(iCustID As Integer)
'IsNewRecord As Boolean,
   OpenRs
   m_rsCustomerRates.AddNew
   m_rsCustomerRates("CustomerID") = iCustID
   m_rsCustomerRates("ProductID") = ProductID  'Product.ID
   m_rsCustomerRates("Rate") = m_dRate

   m_rsCustomerRates.Update
   m_rsCustomerRates.Close
'   IsNewRecord = False

End Sub

Private Sub Class_Initialize()
'    m_objConnectDB.Connect
  '  Set m_objCustomer = New CCustomer
  '  Set m_objProduct = New CProduct
End Sub

