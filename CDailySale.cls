VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDailySale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CPTrans

Option Explicit

Private m_lID As Long
Private m_lSalemanID As Long
Private m_dtTransDate As String
Private m_lProductID As Long
Private m_lWholeSale As Long
Private m_lRetail As Long
Private m_lScheme As Long
Private m_lQuantity As Long

Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private rsDailySale As ADODB.Recordset
Dim CConnectDB As ConnectDB
Public Event DataValidated(bValid As Boolean)

Public Property Let ID(ByVal vData As Long)
    m_lID = vData
    ValidateData
End Property
Public Property Get ID() As Long
    ID = m_lID
End Property
Public Property Let SalemanID(ByVal vData As Long)
    m_lSalemanID = vData
    ValidateData
End Property
Public Property Get SalemanID() As Long
    SalemanID = m_lSalemanID
End Property
Public Property Let TransDate(ByVal vData As String)
    m_dtTransDate = vData
    ValidateData
End Property
Public Property Get TransDate() As String
    TransDate = m_dtTransDate
End Property
Public Property Let ProductID(ByVal vData As Long)
    m_lProductID = vData
    ValidateData
End Property
Public Property Get ProductID() As Long
    ProductID = m_lProductID
End Property
Public Property Let WholeSale(ByVal vData As Long)
    m_lWholeSale = vData
    ValidateData
End Property
Public Property Get WholeSale() As Long
    WholeSale = m_lWholeSale
End Property
Public Property Let Retail(ByVal vData As Long)
    m_lRetail = vData
    ValidateData
End Property
Public Property Get Retail() As Long
    Retail = m_lRetail
End Property
Public Property Let Scheme(ByVal vData As Long)
    m_lScheme = vData
    ValidateData
End Property
Public Property Get Scheme() As Long
    Scheme = m_lScheme
End Property
Public Property Let Quantity(ByVal vData As Long)
    m_lQuantity = vData
    ValidateData
End Property
Public Property Get Quantity() As Long
    Quantity = m_lQuantity
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
   bValid = bValid And m_lSalemanID <> 0
   bValid = bValid And m_dtTransDate <> ""
   
   m_bIsDataValid = bValid
   RaiseEvent DataValidated(bValid)
End Sub

Public Sub Init(Optional lID As Long = 0, Optional dtTrDate As Date)
   OpenRs
   'Dim qdData As QueryDef
   'Set m_DB = objDB
   ' If the ID is missing, the object
   ' is being instantiated to be filled
   ' with new data.
   If lID = 0 Then
      m_bIsNewRecord = True
      Exit Sub
      Else
      rsDailySale.Find "SaleManID=" & lID & " AND " & "TransDate=" & dtTrDate
   End If
   
   m_lID = rsDailySale("ID") & ""
   m_lSalemanID = rsDailySale("SalemanID") & ""
   m_dtTransDate = rsDailySale("TransDate") & ""
   m_lProductID = rsDailySale("ProductID") & ""
   m_lWholeSale = rsDailySale("WholeSale") & ""
   m_lRetail = rsDailySale("Retail") & ""
   m_lScheme = rsDailySale("Scheme") & ""
   m_lQuantity = rsDailySale("TotalQuantity") & ""
   
   ' Clean up
   '
   'm_RS.Close
   'Set m_RS = Nothing

End Sub

Public Sub Save(IsNewRecord As Boolean)
   '
   If IsNewRecord Then
    OpenRs
    rsDailySale.AddNew
   End If

   rsDailySale("SalemanID") = m_lSalemanID
   rsDailySale("TransDate") = m_dtTransDate
   rsDailySale("ProductID") = m_lProductID
   rsDailySale("WholeSale") = m_lWholeSale
   rsDailySale("Retail") = m_lRetail
   rsDailySale("Scheme") = m_lScheme
   rsDailySale("TotalQuantity") = m_lQuantity
   
   
   m_bIsNewRecord = False
   rsDailySale.Update
End Sub

Public Sub OpenRs(Optional lSalemanID As Long = 0, Optional sTrDate As String = "")
    Set CConnectDB = New ConnectDB
    CConnectDB.Connect
    Set rsDailySale = New ADODB.Recordset
    Dim strRecordset As String
    If lSalemanID = 0 And sTrDate = "" Then
        strRecordset = "Select * from DailySale"
    Else
        strRecordset = "Select * from DailySale WHERE DailySale.SalemanID= " & lSalemanID & " AND DailySale.TransDate= " & sTrDate
    End If
    If rsDailySale.State <> adStateOpen Then
    rsDailySale.Open strRecordset, CConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub
Public Sub DeleteRec(ByVal lSalemanID As Long, ByVal sTrDate As String)
    OpenRs
    Dim cmdUpdateInStock As ADODB.Command
    Set cmdUpdateInStock = New ADODB.Command
    cmdUpdateInStock.CommandType = adCmdText
    cmdUpdateInStock.CommandText = "DELETE FROM DailySale WHERE (((DailySale.lSalemanID) = " & strTransType & ") AND ((DailySale.TransDate) = " & sTrDate & "))"
    cmdUpdateInStock.ActiveConnection = m_objConnectDB.cnnMyshop
    cmdUpdateInStock.Execute
End Sub

Public Function FindRec(ByVal lPTransID As Long, ByVal strTransType As String) As Boolean
    OpenRs
    rsDailySale.MoveFirst
    Do While Not rsDailySale.EOF
        If rsDailySale("PTransID") = lPTransID And rsDailySale("lSalemanID") = strTransType Then
            FindRec = True
            Exit Function
        Else
            rsDailySale.MoveNext
        End If
    Loop
    FindRec = False
    
End Function

