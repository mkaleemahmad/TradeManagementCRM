VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPTransDets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class CPTransDets
Option Explicit
Private m_objProduct As CProduct
Private m_objTransDet As CTransDet
Private m_objTransDets As CTransDets
Private m_objTrans As CTrans
Private m_Col As Collection
Private m_sPTransID As String
'Private m_objConnectDB As ConnectDB
Private Const ITEMPREFIX = "Item"

Public Sub AddC(objNewMember As CPTransDet, iLoc As Integer)
   m_Col.Add objNewMember, ITEMPREFIX & Str(iLoc)
End Sub

Public Property Get Item(vntIndexKey As Variant) As CPTransDet
   If vntIndexKey > 0 Then
 
   Set Item = m_Col(vntIndexKey)
   End If
End Property

Public Property Get Count() As Long
   Count = m_Col.Count
End Property

Public Sub Remove(iColIndex As Long, detID As Long)

If detID > 0 Then
 m_Col(iColIndex).Rate = -1
 Else
   m_Col.Remove ITEMPREFIX & Str(iColIndex)
End If

End Sub

Public Property Get NewEnum() As IUnknown
   Set NewEnum = m_Col.[_NewEnum]
End Property

Private Sub Class_Initialize()
   'Set m_objConnectDB = New ConnectDB
   m_objConnectDB.Connect
   Set m_objTransDets = New CTransDets
   Set m_objProduct = New CProduct
   Set m_Col = New Collection
End Sub

Private Sub Class_Terminate()
   '
   ' Just to make sure the memory is cleaned up,
   ' each object in the collection is cleared
   ' before the collection object itself is
   ' cleared.
   '
   Dim objLoopDetail As CPTransDet
   
   For Each objLoopDetail In m_Col
      Set objLoopDetail = Nothing
   Next objLoopDetail
   Set m_Col = Nothing

End Sub

Public Sub Init(lPTransID As Long, sPTransType As String)
    Dim objDetail As CPTransDet
'    Dim cmdCPTransDets As ADODB.Command
'    Dim paramlPTransID As ADODB.Parameter
'    Dim paramsPTransType As ADODB.Parameter
     'Set m_objConnectDB = New ConnectDB
     m_objConnectDB.Connect
     Dim rsCPTransDets As ADODB.Recordset
     Dim strRecordset As String
     Set rsCPTransDets = New ADODB.Recordset
     strRecordset = "SELECT ProdTransDet.detID, ProdTransDet.PTransID, ProdTransDet.PTransType, ProdTransDet.ProductID, ProdTransDet.ColorID, ProdTransDet.Quantity, ProdTransDet.rate, ProdTransDet.DiscountPCent, ProdTransDet.DiscountValue From ProdTransDet WHERE ProdTransDet.PTransID= " & lPTransID & " AND ProdTransDet.PTransType = '" & sPTransType & "'"
     rsCPTransDets.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    
'    Set cmdCPTransDets = New ADODB.Command
'    Set paramlPTransID = New ADODB.Parameter
'    Set paramsPTransType = New ADODB.Parameter
'
'    With cmdCPTransDets
'        .ActiveConnection = m_objConnectDB.cnnMyshop
'        .CommandType = adCmdStoredProc
'        .CommandText = "qryGetPTransDets"
'    End With
'
'    Set paramlPTransID = cmdCPTransDets.CreateParameter("PTransID", adInteger, adParamInput, , lPTransID)
'    Set paramsPTransType = cmdCPTransDets.CreateParameter("PTransType", adChar, adParamInput, , sPTransType)
'
'    cmdCPTransDets.Parameters.Append paramlPTransID
'    cmdCPTransDets.Parameters.Append paramsPTransType
    
   
    Dim a As Integer
    a = 1
    Do While Not rsCPTransDets.EOF
      Set objDetail = New CPTransDet
   '   objDetail.Init rsCPTransDets("PTransID"), rsCPTransDets("PTransType")
        objDetail.detID = rsCPTransDets("detID")
        objDetail.ID = rsCPTransDets("PTransID")
        objDetail.PTransType = rsCPTransDets("PTransType")
        objDetail.ProductID = rsCPTransDets("ProductID")
        objDetail.ColorID = rsCPTransDets("ColorID")
        objDetail.Quantity = rsCPTransDets("Quantity")
        objDetail.Rate = rsCPTransDets("Rate")
        objDetail.DiscountPerc = IsNull2(rsCPTransDets("DiscountPCent"), 0)
      AddC objDetail, a
      a = a + 1
      rsCPTransDets.MoveNext
    Loop
    m_objTransDets.Init lPTransID, sPTransType
    rsCPTransDets.Close
    Set rsCPTransDets = Nothing
End Sub

Public Sub Save(IsNewRecord As Boolean, lID As Long, sPTransType As String, lAccountID As Long, gNetAmount As Single, gDiscAmount As Single)
   Dim objDetail As CPTransDet
   Dim a As Integer
   Dim s As String
   
   
   For Each objDetail In m_Col
   
    If objDetail.detID = 0 Then
      
      objDetail.Save (lID)
      
      m_objProduct.Init objDetail.ProductID
      's = s & Trim(Str(objDetail.Quantity)) & " " & Trim(objDetail.Unit) & " " & Left(m_objProduct.Description, 20) & "@" & Trim(Str(objDetail.Rate))
          If sPTransType = "OS" Or sPTransType = "PH" Or sPTransType = "SR" Or sPTransType = "PD" Then
            m_objProduct.UpdateStock objDetail.ProductID, objDetail.Quantity, "+", objDetail.Rate
           
            If objDetail.ColorID > 0 Then
                m_objProduct.UpdateColorStock objDetail.ProductID, objDetail.ColorID, objDetail.Quantity, "+"
            End If
            
          ElseIf sPTransType = "SL" Or sPTransType = "PR" Or sPTransType = "IP" Then
            m_objProduct.UpdateStock objDetail.ProductID, objDetail.Quantity, "-"
            
            If objDetail.ColorID > 0 Then
                m_objProduct.UpdateColorStock objDetail.ProductID, objDetail.ColorID, objDetail.Quantity, "-"
            End If
    
          End If
      
      ElseIf objDetail.Rate = -1 Then
            
              objDetail.DeleteRecByDetID (objDetail.detID)
              
              If sPTransType = "OS" Or sPTransType = "PH" Or sPTransType = "SR" Or sPTransType = "PD" Then
                m_objProduct.UpdateStock objDetail.ProductID, objDetail.Quantity, "-", objDetail.Rate
               
                If objDetail.ColorID > 0 Then
                    m_objProduct.UpdateColorStock objDetail.ProductID, objDetail.ColorID, objDetail.Quantity, "-"
                End If
                
              ElseIf sPTransType = "SL" Or sPTransType = "PR" Or sPTransType = "IP" Then
                m_objProduct.UpdateStock objDetail.ProductID, objDetail.Quantity, "+"
                
                If objDetail.ColorID > 0 Then
                    m_objProduct.UpdateColorStock objDetail.ProductID, objDetail.ColorID, objDetail.Quantity, "+"
                End If
        
              End If
      
    End If
            a = a + 1
  
'      If Trim(s) <> "" Then
'       s = s & ","
'      End If
        'Accounting transaction (DETAILS ENTRY) in Sales OR PURCHASE Account
        If sPTransType <> "OS" And sPTransType <> "PD" And sPTransType <> "IP" Then
             Set m_objTransDet = New CTransDet
             m_objTransDet.ID = lID
             m_objTransDet.TransType = sPTransType
             m_objProduct.Init objDetail.ProductID
             m_objTransDet.Description = Trim(Str(objDetail.Quantity)) & Trim(objDetail.Unit) & " " & m_objProduct.Description & "@" & Trim(Str(objDetail.Rate))
            Select Case sPTransType
                Case "PH"
                    m_objTransDet.AccountID = SelAcntIDs.Purchase
                    m_objTransDet.Debit = (objDetail.Quantity * objDetail.Rate)
                Case "PR"
                    m_objTransDet.AccountID = SelAcntIDs.PurchaseReturn
                    m_objTransDet.Credit = (objDetail.Quantity * objDetail.Rate)
                Case "SL"
                    m_objTransDet.AccountID = SelAcntIDs.Sales
                    m_objTransDet.Credit = (objDetail.Quantity * objDetail.Rate)
                Case "SR"
                    m_objTransDet.AccountID = SelAcntIDs.SalesReturn
                    m_objTransDet.Debit = (objDetail.Quantity * objDetail.Rate)
            End Select
            m_objTransDets.AddC m_objTransDet, a
         End If
   Next objDetail
'   s = Left(s, Len(s) - 1)
'         If sPTransType <> "OS" And sPTransType <> "PD" And sPTransType <> "IP" Then
'             Set m_objTransDet = New CTransDet
'             m_objTransDet.ID = lID
'             m_objTransDet.TransType = sPTransType
'             m_objTransDet.AccountID = lAccountID
'             m_objTransDet.Description = s
'             If sPTransType = "SL" Or sPTransType = "PR" Then
'              m_objTransDet.Debit = gNetAmount + gDiscAmount
'             ElseIf sPTransType = "SR" Or sPTransType = "PH" Then
'              m_objTransDet.Credit = gNetAmount + gDiscAmount
'             End If
'             m_objTransDets.AddC m_objTransDet, a
'          End If
    ''Discount Entry
        If gDiscAmount > 0 And (sPTransType <> "OS" And sPTransType <> "PD" And sPTransType <> "IP") Then
            a = a + 1
            Set m_objTransDet = New CTransDet
            m_objTransDet.ID = lID
            m_objTransDet.TransType = sPTransType
             Select Case sPTransType
                Case "SL"
                    m_objTransDet.AccountID = SelAcntIDs.DiscountsOffered
                    m_objTransDet.Description = "Discount - Sales No. " & lID
                    m_objTransDet.Debit = gDiscAmount
                Case "SR"
                    m_objTransDet.AccountID = SelAcntIDs.DiscountsOffered
                    m_objTransDet.Description = "Discount - S. Return No. " & lID
                    m_objTransDet.Credit = gDiscAmount
                Case "PH"
                    m_objTransDet.AccountID = SelAcntIDs.DiscountsReceived
                    m_objTransDet.Description = "Discount - Purchase No. " & lID
                    m_objTransDet.Credit = gDiscAmount
                Case "PR"
                    m_objTransDet.AccountID = SelAcntIDs.DiscountsReceived
                    m_objTransDet.Description = "Discount - P. Return No. " & lID
                    m_objTransDet.Debit = gDiscAmount
                End Select
                    m_objTransDets.AddC m_objTransDet, a
            'Discount Second Entry
                 a = a + 1
                 Set m_objTransDet = New CTransDet
                 m_objTransDet.ID = lID
                 m_objTransDet.TransType = sPTransType
                 m_objTransDet.AccountID = lAccountID
                 m_objTransDet.Description = "Discount - Entry No. " & sPTransType & "-" & lID
                 If sPTransType = "SL" Or sPTransType = "PR" Then
                  m_objTransDet.Credit = gDiscAmount
                 ElseIf sPTransType = "SR" Or sPTransType = "PH" Then
                  m_objTransDet.Debit = gDiscAmount
                 End If
                 m_objTransDets.AddC m_objTransDet, a
        End If
        'Accounting transaction TO CUSTOMER, SUPPLIER OR CASH ACCOUNT
         If sPTransType <> "OS" And sPTransType <> "PD" And sPTransType <> "IP" Then
             a = a + 1
             Set m_objTransDet = New CTransDet
             m_objTransDet.ID = lID
             m_objTransDet.TransType = sPTransType
             m_objTransDet.AccountID = lAccountID
             m_objTransDet.Description = s
             If sPTransType = "SL" Or sPTransType = "PR" Then
              m_objTransDet.Debit = gNetAmount + gDiscAmount
             ElseIf sPTransType = "SR" Or sPTransType = "PH" Then
              m_objTransDet.Credit = gNetAmount + gDiscAmount
             End If
             m_objTransDets.AddC m_objTransDet, a
             m_objTransDets.Save IsNewRecord, lID
             Set m_Col = New Collection
         End If
End Sub

Property Get Total() As Currency
   Dim cTemp As Currency
   Dim objDetail As CPTransDet
   cTemp = 0
   For Each objDetail In m_Col
      cTemp = cTemp + objDetail.Total
   Next objDetail
   Total = cTemp
End Property

Property Get TotalQuantity() As Currency
   Dim dTemp As Double
   Dim objDetail As CPTransDet
   dTemp = 0
   For Each objDetail In m_Col
      dTemp = dTemp + objDetail.Quantity
   Next objDetail
   TotalQuantity = dTemp
End Property
Property Get TotalDiscount() As Currency
   Dim dTemp As Double
   Dim objDetail As CPTransDet
   dTemp = 0
   For Each objDetail In m_Col
      dTemp = dTemp + (objDetail.Total * objDetail.DiscountPerc / 100)
   Next objDetail
   TotalDiscount = dTemp
End Property
