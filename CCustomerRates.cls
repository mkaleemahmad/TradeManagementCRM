VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustomerRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CCustomerRates
Option Explicit

Private m_Col As Collection
Private m_RS As Recordset
Private m_objConnectDB As ConnectDB

'Private m_sPTransID As String
'Private m_sPTransType As String

Private Const ITEMPREFIX = "Item"

Public Sub AddC(objNewMember As CCustomerRate, iLoc As Integer)
   m_Col.Add objNewMember, ITEMPREFIX & Str(iLoc)
End Sub

Public Property Get Item(vntIndexKey As Variant) As CCustomerRate
   Set Item = m_Col(vntIndexKey)
End Property

Public Property Get Count() As Long
   Count = m_Col.Count
End Property

Public Sub Remove(iColIndex As Long)
   m_Col.Remove ITEMPREFIX & Str(iColIndex)
End Sub

Public Property Get NewEnum() As IUnknown
   Set NewEnum = m_Col.[_NewEnum]
End Property

Private Sub Class_Initialize()
   Set m_Col = New Collection
   Set m_objConnectDB = New ConnectDB
   m_objConnectDB.Connect
End Sub

Private Sub Class_Terminate()

   '
   ' Just to make sure the memory is cleaned up,
   ' each object in the collection is cleared
   ' before the collection object itself is
   ' cleared.
   '
   Dim objLoopDetail As CCustomerRate
   
   For Each objLoopDetail In m_Col
      Set objLoopDetail = Nothing
   Next objLoopDetail
   Set m_Col = Nothing

End Sub

Public Sub Init(iCustID As Long)
    Dim objDetail As CCustomerRate
    Dim cmdCustRates As ADODB.Command
    Dim paramCustId As ADODB.Parameter
    Dim rsCustRates As ADODB.Recordset
    Set cmdCustRates = New ADODB.Command
    Set paramCustId = New ADODB.Parameter
    With cmdCustRates
        .ActiveConnection = m_objConnectDB.cnnMyshop
        .CommandType = adCmdStoredProc
        .CommandText = "qryGetCustomerRatesbyCustomer"
    End With
    Set paramCustId = cmdCustRates.CreateParameter("CustID", adInteger, adParamInput, , iCustID)
    cmdCustRates.Parameters.Append paramCustId
    Set rsCustRates = cmdCustRates.Execute
    Do While Not rsCustRates.EOF
      Set objDetail = New CCustomerRate
     ' objDetail.Init rsCustRates("CustomerID"), rsCustRates("ProductID")
      objDetail.CustomerID = rsCustRates("CustomerID")
      objDetail.ProductID = rsCustRates("ProductID")
      objDetail.ProductTotalDescription = Trim(rsCustRates("Description")) + " " + Trim(rsCustRates("Size")) + " " + Trim(rsCustRates("HorsePower")) + " " + Trim(rsCustRates("Size")) + " " + Trim(rsCustRates("Quality"))
      objDetail.Rate = rsCustRates("Rate")
      AddC objDetail, rsCustRates("ProductID")
      rsCustRates.MoveNext
    Loop
    rsCustRates.Close
    Set rsCustRates = Nothing
End Sub

Public Function GetProductRate(lProdID As Long) As Single
    Dim obj As CCustomerRate
    For Each obj In m_Col
      If obj.ProductID = lProdID Then
        GetProductRate = obj.Rate
        Exit For
      End If
    Next obj
    Set obj = Nothing
End Function


Public Sub Save(iCustomerID As Integer)
   
   Dim cmdDelCustRates As ADODB.Command
   Set cmdDelCustRates = New ADODB.Command
   cmdDelCustRates.ActiveConnection = m_objConnectDB.cnnMyshop
   cmdDelCustRates.CommandType = adCmdStoredProc
   cmdDelCustRates.CommandText = "qryDelCustomerRates"
   'cmdDelCustRates.Parameters("paramCustomerID") = iCustomerID
   cmdDelCustRates.Execute , iCustomerID
   
   Dim objDetail As CCustomerRate
   '
   ' This object's parent can simply call
   ' Save on the collection instead of
   ' calling Save on each individual item.
   '
   For Each objDetail In m_Col
      objDetail.Save iCustomerID
   Next objDetail
   
End Sub
