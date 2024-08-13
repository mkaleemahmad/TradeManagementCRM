VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' Class CProduct
Private m_lID As Long
Private m_sProdID As String
Private m_sDescription As String
Private m_sPackingMode As String
Private m_sPacking As String
Private m_dSalePrice As Double
Private m_dWholeSalePrice As Double
Private m_dPurchaseRate As Double
Private m_dPRateInDollar As Double
Private m_dBalance As Double
Private m_sReorderLevel As Single
Private m_lGroupID As Long
Private m_lCategoryID As Long
Private m_sProductDetails As String
Private m_sUnit As String
Private m_dMinSalePrice As Double
Private m_sStoreLocation As String
Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private m_bIsFound As Boolean
Private m_bIsActive As Boolean
Private m_sStockStatus As String
Private m_rsProduct As ADODB.Recordset
Private m_colProd As Collection
Public Event DataValidated(bValid As Boolean)

'Public Property Let ID(ByVal vData As Long)
'    m_lID = vData
'    ValidateData
'End Property
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

Public Property Let ProductID(ByVal vData As String)
    m_sProdID = vData
    ValidateData
End Property
Public Property Get ProductID() As String
    ProductID = m_sProdID
End Property

Public Property Let GroupID(ByVal vData As Long)
    m_lGroupID = vData
    ValidateData
End Property
Public Property Get GroupID() As Long
    GroupID = m_lGroupID
End Property

Public Property Let CategoryID(ByVal vData As Long)
    m_lCategoryID = vData
    ValidateData
End Property
Public Property Get CategoryID() As Long
    CategoryID = m_lCategoryID
End Property
Public Property Let Description(ByVal vData As String)
    m_sDescription = vData
    ValidateData
End Property
Public Property Get Description() As String
    Description = m_sDescription
End Property
Public Property Let StoreLocation(ByVal vData As String)
    m_sStoreLocation = vData
    ValidateData
End Property
Public Property Get StoreLocation() As String
    StoreLocation = m_sStoreLocation
End Property

Public Property Let PackingMode(ByVal vData As String)
    m_sPackingMode = vData
    ValidateData
End Property
Public Property Get PackingMode() As String
    PackingMode = m_sPackingMode
End Property

Public Property Let Packing(ByVal vData As String)
    m_sPacking = vData
End Property

Public Property Get Packing() As String
    Packing = m_sPacking
End Property

Public Property Let SalePrice(ByVal vData As Double)
    m_dSalePrice = vData
    ValidateData
End Property
Public Property Get SalePrice() As Double
    SalePrice = m_dSalePrice
End Property
Public Property Let MinSalePrice(ByVal vData As Double)
    m_dMinSalePrice = vData
    ValidateData
End Property
Public Property Get MinSalePrice() As Double
    MinSalePrice = m_dMinSalePrice
End Property

Public Property Let WholeSalePrice(ByVal vData As Double)
    m_dWholeSalePrice = vData
    ValidateData
End Property
Public Property Get WholeSalePrice() As Double
    WholeSalePrice = m_dWholeSalePrice
End Property

Public Property Let PurchaseRate(ByVal vData As Double)
    m_dPurchaseRate = vData
    ValidateData
End Property
Public Property Get PurchaseRate() As Double
    PurchaseRate = m_dPurchaseRate
End Property

Public Property Let Balance(ByVal vData As Double)
    m_dBalance = vData
    ValidateData
End Property
Public Property Get Balance() As Double
    Balance = m_dBalance
End Property

Public Property Let ReorderLevel(ByVal vData As Single)
    m_sReorderLevel = vData
    ValidateData
End Property
Public Property Get ReorderLevel() As Single
    ReorderLevel = m_sReorderLevel
End Property

Public Property Get StockStatus() As String
    StockStatus = m_sStockStatus
End Property

Public Property Let ProductDetails(ByVal vData As String)
    m_sProductDetails = vData
    ValidateData
End Property
Public Property Get ProductDetails() As String
    ProductDetails = m_sProductDetails
End Property

Public Property Let Unit(ByVal vData As String)
    m_sUnit = vData
    ValidateData
End Property
Public Property Get Unit() As String
    Unit = m_sUnit
End Property

Public Property Get IsValid() As Boolean
   IsValid = m_bIsDataValid
End Property

Public Property Get IsFound() As Boolean
   IsFound = m_bIsFound
End Property

Private Sub ValidateData()
   Dim bValid As Boolean
   bValid = True
   'bValid = bValid And m_sProdID <> ""
   bValid = bValid And m_sDescription <> ""
   bValid = bValid And m_dSalePrice <> 0
'   bValid = bValid And m_dWholeSalePrice <> 0
   m_bIsDataValid = bValid
   RaiseEvent DataValidated(bValid)
End Sub

Private Sub Class_Initialize()
   Set m_colProd = New Collection
   m_bIsFound = False
   ValidateData
End Sub

Public Sub Init(Optional lID As Long = 0)
  If lID = 0 Then
      m_bIsNewRecord = True
      OpenRs
      Exit Sub
    Else
        '      Dim m_RS As ADODB.Recordset
        '      Dim strRecordset As String
        '      Set m_RS = New ADODB.Recordset
        '      strRecordset = "Select * From Products where Products.ID=" & lID
        '      m_RS.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
      OpenRs , , , lID
   End If
       ' Load each of the private member
       ' variables with the corresponding
       ' data from the database
       m_lID = m_rsProduct("ID") & ""
       m_sProdID = m_rsProduct("ProductID") & ""
       m_sDescription = m_rsProduct("Description") & ""
       m_sPackingMode = m_rsProduct("PackingMode") & ""
       m_sPacking = m_rsProduct("Packing") & ""
       m_dSalePrice = IIf(IsNull(m_rsProduct("SalePrice")), 0, m_rsProduct("SalePrice"))
       m_dWholeSalePrice = IIf(IsNull(m_rsProduct("WholeSalePrice")), 0, m_rsProduct("WholeSalePrice"))
       m_dPurchaseRate = IIf(IsNull(m_rsProduct("PurchaseRate")), 0, m_rsProduct("PurchaseRate"))
       m_dBalance = IIf(IsNull(m_rsProduct("Balance")), 0, m_rsProduct("Balance"))
       m_sReorderLevel = IIf(IsNull(m_rsProduct("ReorderLevel")), 0, m_rsProduct("ReorderLevel"))
       m_sStockStatus = IIf(IsNull(m_rsProduct("StockStatus")), "", m_rsProduct("StockStatus"))
       m_lGroupID = IIf(IsNull(m_rsProduct("GroupID")), "", m_rsProduct("GroupID"))
       m_lCategoryID = IIf(IsNull(m_rsProduct("CategoryID")), "", m_rsProduct("CategoryID"))
       m_sProductDetails = IIf(IsNull(m_rsProduct("ProductDetails")), " ", m_rsProduct("ProductDetails"))
       m_sUnit = IIf(IsNull(m_rsProduct("Unit")), " ", m_rsProduct("Unit"))
       m_dPRateInDollar = IsNull2(m_rsProduct("PRateInDollar"), 0)
       m_bIsActive = m_rsProduct("Active") & ""
       m_dMinSalePrice = IsNull2(m_rsProduct("MinSalePrice"), 0)
       m_sStoreLocation = IsNull2(m_rsProduct("StoreLocation"), "")
       ' Clean up
'       m_RS.Close
'       Set m_RS = Nothing
End Sub

Public Sub Save(m_bIsNewRecord As Boolean)
  If m_bIsNewRecord Then
     OpenRs
'     m_rsProduct.MoveLast
     m_rsProduct.AddNew
   Else
     'OpenRs , , m_lID
     m_rsProduct.MoveFirst
   End If
     m_rsProduct("ProductID") = m_sProdID
     m_rsProduct("Description") = m_sDescription
     m_rsProduct("PackingMode") = m_sPackingMode
     m_rsProduct("Packing") = m_sPacking
     m_rsProduct("SalePrice") = m_dSalePrice
     m_rsProduct("WholeSalePrice") = m_dWholeSalePrice
     m_rsProduct("PurchaseRate") = m_dPurchaseRate
     m_rsProduct("Balance") = m_dBalance
     m_rsProduct("ReorderLevel") = m_sReorderLevel
     m_rsProduct("StockStatus") = m_sStockStatus
     m_rsProduct("GroupID") = m_lGroupID
     m_rsProduct("CategoryID") = m_lCategoryID
     m_rsProduct("ProductDetails") = m_sProductDetails
     m_rsProduct("Unit") = m_sUnit '= IIf(IsNull(m_rsProduct("Unit")), " ", m_rsProduct("Unit"))
     m_rsProduct("Active") = m_bIsActive
     m_rsProduct("PRateInDollar") = m_dPRateInDollar
     m_rsProduct("MinSalePrice") = m_dMinSalePrice
     m_rsProduct("StoreLocation") = m_sStoreLocation
     m_rsProduct.Update
     m_rsProduct.Requery
     m_rsProduct.MoveLast
     m_lID = m_rsProduct("ID")
     m_bIsNewRecord = False
End Sub

Private Sub OpenRs(Optional strCriteria As String = "", Optional lCategoryID As Long = 0, Optional lGroupID As Long = 0, Optional lProdID As Long = 0, Optional sOrderBy As String = "")
    Set m_rsProduct = New ADODB.Recordset
    Dim strRecordset As String
    'strRecordset = "Select * from Products"
    If strCriteria = "" And lCategoryID = 0 And lGroupID = 0 And lProdID = 0 And sOrderBy = "" Then
        strRecordset = "Select * from Products where Active = 1 ORDER BY ID"
    ElseIf strCriteria = "" And lCategoryID = 0 And lGroupID = 0 And lProdID = 0 And sOrderBy <> "" Then
        strRecordset = "Select * from Products where Active = 1 " & sOrderBy
    ElseIf strCriteria <> "" And lCategoryID = 0 And lGroupID = 0 And lProdID = 0 Then 'Product Code given
        strRecordset = "Select * from Products Where Active = 1 AND ProductID = " & "'" & strCriteria & "'"
    ElseIf strCriteria = "" And lCategoryID <> 0 And lGroupID = 0 And lProdID = 0 Then 'Product Category given
        strRecordset = "Select * from Products Where Active = 1 AND CategoryID = " & lCategoryID & " Order By Description"
    ElseIf strCriteria = "" And lCategoryID = 0 And lGroupID <> 0 And lProdID = 0 Then
        strRecordset = "Select * from Products Where Active = 1 and GroupID = " & lGroupID & " Order By Description"
    ElseIf strCriteria = "" And lCategoryID <> 0 And lGroupID <> 0 And lProdID = 0 Then
        strRecordset = "Select * from Products Where Active = 1 AND CategoryID = " & lCategoryID & " and GroupID = " & lGroupID & " Order By Description"
    ElseIf strCriteria = "" And lCategoryID = 0 And lGroupID = 0 And lProdID <> 0 Then 'Product Code given
        strRecordset = "Select * from Products Where ID = " & lProdID
    End If
    '"Order By Description"
    If m_rsProduct.State <> adStateOpen Then
        m_rsProduct.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenKeyset, adLockOptimistic
        'adOpenStatic, adLockOptimistic
    End If
End Sub

Public Sub UpdateFormCombo(ByRef m_objCmb As ComboBox, Optional ByRef m_objCmb2 As ComboBox, Optional lCategoryID As Long = 0, Optional lGroupID As Long = 0, Optional bTwoCombos As Boolean = False)
    
    OpenRs , lCategoryID, lGroupID, , " Order By Description"
    Dim a As Integer
    a = 0
    m_objCmb.Clear
    Do While Not m_rsProduct.EOF
        m_objCmb.AddItem (m_rsProduct("Description"))
        m_objCmb.ItemData(a) = m_rsProduct("ID")
        If bTwoCombos = True Then
            m_objCmb2.AddItem (m_rsProduct("SalePrice"))
            m_objCmb2.ItemData(a) = m_rsProduct("ID")
        End If
        a = a + 1
        m_rsProduct.MoveNext
    Loop
End Sub

Public Function GetProductID(PCode As String) As Integer
    OpenRs
    Dim counter As Integer
    counter = 0
    Do While Not m_rsProduct.EOF
        If m_rsProduct.Fields("productid") = PCode Then
            GetProductID = m_rsProduct.Fields("ID")
            Exit Function
        End If
        counter = counter + 1
        m_rsProduct.MoveNext
    Loop
End Function

Public Function GetProductCode(PID As Long) As String
    OpenRs , , , PID
    Dim counter As Integer
    If Not m_rsProduct.EOF Then
         GetProductCode = m_rsProduct.Fields("productid")
    Else
        GetProductCode = ""
    End If
End Function

Public Function GetProdName(sPCode As String) As String
    OpenRs "ProductID = '" & sPCode & "'"
    If Not m_rsProduct.EOF Then
         GetProdName = m_rsProduct.Fields("Description")
    Else
        GetProdName = ""
    End If
End Function

Public Function ProductCodeExist(PCode1 As String, PID As Integer) As Boolean
    Dim RS1 As New ADODB.Recordset
    Dim cmd1 As New ADODB.Command
    With cmd1
        .ActiveConnection = m_objConnectDB.cnnMyshop
        .CommandType = adCmdStoredProc
        .CommandText = "spIsProductCodeExists"
        .Parameters("@ProductCode") = PCode1
        .Parameters("@ProductID") = PID
        Set RS1 = .Execute
    End With
    If Not RS1.BOF And Not RS1.EOF Then
        ProductCodeExist = True
    Else
        ProductCodeExist = False
    End If
    RS1.Close
    Set RS1 = Nothing
End Function


Public Function FindRec(sFindCriteria As String)
    OpenRs sFindCriteria
    If m_rsProduct.BOF = False And m_rsProduct.EOF = False Then
       m_lID = m_rsProduct("ID") & ""
       m_sProdID = m_rsProduct("ProductID") & ""
       m_sDescription = m_rsProduct("Description") & ""
       'm_sPackingMode = m_rsProduct("PackingMode") & ""
       'm_sPacking = m_rsProduct("Packing") & ""
       m_dSalePrice = m_rsProduct("SalePrice") & ""
       'm_dWholeSalePrice = m_rsProduct("WholeSalePrice") & ""
       m_dPurchaseRate = m_rsProduct("PurchaseRate") & ""
       m_dBalance = m_rsProduct("Balance") & ""
       m_sStockStatus = m_rsProduct("StockStatus") & ""
       m_lGroupID = m_rsProduct("GroupID") & ""
       'm_lCategoryID = m_rsProduct("CategoryID") & ""
       'm_sProductDetails = m_rsProduct("ProductDetails") & ""
       'm_bIsActive = m_rsProduct("Active") & ""
       'm_dPRateInDollar = IsNull2(m_rsProduct("PRateInDollar"), 0)
       'm_sUnit = IIf(IsNull(m_rsProduct("Unit")), " ", m_rsProduct("Unit"))
        m_bIsFound = True
    Else
'        m_lID = 0
'        m_sProdID = ""
'        m_sDescription = ""
'        m_sPackingMode = ""
'        m_dSalePrice = 0
'        m_dBalance = 0
'        m_lGroupID = ""
'        m_bIsActive = False
        m_bIsFound = False
    End If
End Function

Public Function UpdateStock(lProductID As Long, dBalance As Double, sChange As String, Optional dPurchRate As Double = 0)
    Dim cmdUpdateStock As New ADODB.Command
    Dim strCommand As String
    strCommand = "UPDATE products SET products.balance = [products].[balance]" & sChange & dBalance & ""
    If dPurchRate > 0 Then
        strCommand = strCommand & " , Products.PurchaseRate = " & dPurchRate
    End If
    strCommand = strCommand & " WHERE ([products].[id]= " & lProductID & ")"
    cmdUpdateStock.CommandText = strCommand
    cmdUpdateStock.ActiveConnection = m_objConnectDB.cnnMyshop
    cmdUpdateStock.Execute
End Function
Public Function UpdateColorStock(lProductID As Long, lColorID As Long, dBalance As Double, sChange As String)
    
    Dim rs As New ADODB.Recordset, sSQL As String
       sSQL = "Select Count(*) as rcd from ProductColors where ([ProductColors].[ProductID]= " & lProductID & ") and ([ProductColors].[ColorID]= " & lColorID & ") "
   
    rs.Open sSQL, m_objConnectDB.cnnMyshop, adOpenForwardOnly
    Dim cmdUpdateStock As New ADODB.Command
    Dim strCommand As String
    If rs.Fields("rcd") > 0 Then
    strCommand = "UPDATE ProductColors SET ProductColors.balance = [ProductColors].[balance]" & sChange & dBalance & ""
    strCommand = strCommand & " WHERE ([ProductColors].[ProductID]= " & lProductID & ") and ([ProductColors].[ColorID]= " & lColorID & ") "
    cmdUpdateStock.CommandText = strCommand
    cmdUpdateStock.ActiveConnection = m_objConnectDB.cnnMyshop
    cmdUpdateStock.Execute
    
    Else
    strCommand = "Insert into ProductColors ([ProductColors].[ProductID], [ProductColors].[ColorID], [ProductColors].[balance]) values(" & lProductID & "," & lColorID & "," & dBalance & ")"
    cmdUpdateStock.CommandText = strCommand
    cmdUpdateStock.ActiveConnection = m_objConnectDB.cnnMyshop
    cmdUpdateStock.Execute
    End If
    
End Function

Sub PopListView(refLvw As ListView)
    Dim k As Long, rs As New ADODB.Recordset, sSQL As String
    Dim cmd1 As New ADODB.Command
    With cmd1
        .ActiveConnection = m_objConnectDB.cnnMyshop
        .CommandText = "ProductsFull"
        .CommandType = adCmdStoredProc
        Set rs = .Execute
    End With
  Dim li As ListItem
  Dim lis As ListSubItem
  Do Until rs.EOF
    Set li = refLvw.ListItems.Add(, "I" & rs.Fields("ID"), "" & SoftVal(rs.Fields("Description")))
    Set lis = li.ListSubItems.Add(, "PCode", "" & SoftVal(rs.Fields("ProductID")))
    'Set lis = li.ListSubItems.Add(, "Pack", "" & SoftVal(rs.Fields("Packing")))
    Set lis = li.ListSubItems.Add(, "Bal", "" & SoftVal(rs.Fields("Balance")))
    'Set lis = li.ListSubItems.Add(, "Cp", "" & SoftVal(rs.Fields("PrCategDesc")))
    Set lis = li.ListSubItems.Add(, "Gr", "" & SoftVal(rs.Fields("PrGroupDesc")))
'    Set lis = li.ListSubItems.Add(, "A", Map2YN(SoftVal(rs.Fields("Active"))))
    rs.MoveNext
   Loop
   If refLvw.ListItems.Count > 0 Then Set refLvw.SelectedItem = refLvw.ListItems.Item(1)
   rs.Close
End Sub

Sub Remove(lID As Long)
Dim sSQL As String
sSQL = "Delete From Products Where ID=" & lID
m_objConnectDB.cnnMyshop.Execute sSQL
End Sub

Function CanDelete(lID As Long) As Boolean
    Dim sSQL As String
    sSQL = "SELECT COUNT(*) AS TotRecs FROM ProdTransDet WHERE ProductID=" & lID
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = m_objConnectDB.cnnMyshop.Execute(sSQL)
    CanDelete = IsNull2(rsTmp!TotRecs, 0) = 0
End Function

Public Property Get PurchaseRateInDollar() As Double
PurchaseRateInDollar = m_dPRateInDollar
End Property

Public Property Let PurchaseRateInDollar(ByVal dNV As Double)
m_dPRateInDollar = dNV
End Property

Public Function GetLastID() As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select ProductID from Products Where ID = (Select Max(ID) From Products)", m_objConnectDB.cnnMyshop, adOpenForwardOnly, adLockReadOnly
    GetLastID = IsNull2(rs("ProductID"), "")
End Function

Public Function NewBarCode() As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select ProdBarCode from Company", m_objConnectDB.cnnMyshop, adOpenForwardOnly, adLockReadOnly
    NewBarCode = IsNull2(rs("ProdBarCode"), 0)
End Function

Public Sub UpdateBarCode(sBarCode As String)
    Dim rs As New ADODB.Recordset
    rs.Open "Update Company Set ProdBarCode = '" & Val(sBarCode) + 1 & "'", m_objConnectDB.cnnMyshop, adOpenForwardOnly, adLockReadOnly
End Sub
