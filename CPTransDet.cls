VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPTransDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CPTransDet
Option Explicit

'Private m_objPTrans As CPTrans
Private m_detID As Long
Private m_lID As Long
Private m_sPTransType As String
Private m_iProductID As Integer
Private m_iColorID As Integer
Private m_dQuantity As Double
Private m_dRate As Double
Private m_sUnit As String
Private m_sPacking As String
Private m_dDisPerc As Double
Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private m_rsPTransDet As ADODB.Recordset
'Private m_clsConnectDB As New ConnectDB
Public Event DataValidated(bValid As Boolean)

Public Property Let detID(ByVal vData As Integer)
    m_detID = vData
    ValidateData
End Property
Public Property Get detID() As Integer
    detID = m_detID
End Property


Public Property Let ID(ByVal vData As Integer)
    m_lID = vData
    ValidateData
End Property
Public Property Get ID() As Integer
    ID = m_lID
End Property

Public Property Let PTransType(ByVal vData As String)
    m_sPTransType = vData
    ValidateData
End Property
Public Property Get PTransType() As String
    PTransType = m_sPTransType
End Property

Public Property Let ProductID(ByVal vData As Integer)
    m_iProductID = vData
    ValidateData
End Property
Public Property Get ProductID() As Integer
    ProductID = m_iProductID
End Property

Public Property Let ColorID(ByVal vData As Integer)
    m_iColorID = vData
    ValidateData
End Property
Public Property Get ColorID() As Integer
    ColorID = m_iColorID
End Property

Public Property Let Quantity(ByVal vData As Double)
    m_dQuantity = vData
    ValidateData
End Property
Public Property Get Quantity() As Double
    Quantity = m_dQuantity
End Property

Public Property Let Rate(ByVal vData As Double)
    m_dRate = vData
    ValidateData
End Property
Public Property Get Rate() As Double
    Rate = m_dRate
End Property
Public Property Let DiscountPerc(ByVal vData As Double)
    m_dDisPerc = vData
    ValidateData
End Property
Public Property Get DiscountPerc() As Double
    DiscountPerc = m_dDisPerc
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
   bValid = bValid And m_iProductID <> 0
   bValid = bValid And m_dQuantity <> 0
   bValid = bValid And m_dRate <> 0
   
   m_bIsDataValid = bValid
   RaiseEvent DataValidated(bValid)
End Sub

Private Sub Class_Initialize()
    Set m_rsPTransDet = New ADODB.Recordset
    ValidateData
End Sub

Public Sub Init(Optional lID As Long = 0, Optional sPTType As String = "", Optional objRs As ADODB.Recordset)
   'OpenRs
   'Dim qdData As QueryDef
   'Set m_DB = objDB
   ' If the ID is missing, the object
   ' is being instantiated to be filled
   ' with new data.
   If lID = 0 Then
      m_bIsNewRecord = True
      Exit Sub
      Else
        If Not objRs Is Nothing Then
            m_lID = objRs("PTransID") & ""
            m_sPTransType = objRs("PTransType") & ""
            m_iProductID = objRs("ProductID") & ""
            m_iColorID = objRs("ColorID") & ""
            m_dQuantity = objRs("Quantity") & ""
            m_dRate = objRs("Rate") & ""
            m_sUnit = objRs("Unit") & ""
            m_dDisPerc = IsNull2(objRs("DiscountPCent"), 0)
        End If
   End If
   
End Sub

Public Sub Save(lID As Long)
   ' Dim qdData As QueryDef
   '
   ' New records have to be added with
   ' the AddNew method instead of Edit.
   ' The m_bIsNewRecord flag is set in
   ' the Init method.

    OpenRs
    m_rsPTransDet.AddNew
   
   ' Update data in each field. If you
   ' are using an AutoNumber field, never
   ' attempt to add/edit data in it. You
   ' will get an error.
   m_rsPTransDet("PTransID") = lID
   m_rsPTransDet("PTransType") = m_sPTransType
   m_rsPTransDet("ProductID") = m_iProductID
   m_rsPTransDet("ColorID") = m_iColorID
   m_rsPTransDet("Quantity") = m_dQuantity
   m_rsPTransDet("Rate") = m_dRate
   m_rsPTransDet("DiscountPCent") = m_dDisPerc
   m_rsPTransDet("Unit") = m_sUnit
   m_rsPTransDet("Packing") = m_sPacking
   
   m_rsPTransDet.Update
   
End Sub

Public Sub OpenRs()
    
    Set m_rsPTransDet = New ADODB.Recordset
    Dim strRecordset As String
    strRecordset = "Select * from ProdTransDet"
    If m_rsPTransDet.State <> adStateOpen Then
        m_rsPTransDet.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub

Public Property Get Total() As Currency
      Total = Round(m_dQuantity * m_dRate, 2) '* (1 - m_fDiscount)
End Property

Public Sub DeleteRec(ByVal lPTransID As Long, ByVal strTransType As String)
    OpenRs
    Dim cmdUpdateInStock As ADODB.Command
    Set cmdUpdateInStock = New ADODB.Command
    cmdUpdateInStock.CommandType = adCmdText
    cmdUpdateInStock.CommandText = "DELETE FROM ProdTransDet WHERE (((ProdTransDet.PTransType) = " & "'" & strTransType & "'" & ") AND ((ProdTransDet.PTransID) = " & lPTransID & "))"
    cmdUpdateInStock.ActiveConnection = m_objConnectDB.cnnMyshop
    cmdUpdateInStock.Execute
End Sub
Public Sub DeleteRecByDetID(ByVal detID As Long)
    OpenRs
    Dim cmdUpdateInStock As ADODB.Command
    Set cmdUpdateInStock = New ADODB.Command
    cmdUpdateInStock.CommandType = adCmdText
    cmdUpdateInStock.CommandText = "DELETE FROM ProdTransDet WHERE detID = " & detID & ""
    cmdUpdateInStock.ActiveConnection = m_objConnectDB.cnnMyshop
    cmdUpdateInStock.Execute
End Sub

Public Function FindRec(ByVal lPTransID As Long, ByVal strTransType As String) As Boolean
    Dim icounter As Integer
    OpenRs
    m_rsPTransDet.MoveFirst
    Do While Not m_rsPTransDet.EOF
        If m_rsPTransDet("PTransID") = lPTransID And m_rsPTransDet("PTransType") = strTransType Then
            FindRec = True
            Exit Function
        Else
        m_rsPTransDet.MoveNext
        End If
    Loop
    FindRec = False
End Function

Public Property Get Unit() As String
Unit = m_sUnit
End Property

Public Property Let Unit(ByVal sNV As String)
m_sUnit = sNV
End Property

Public Property Get Packing() As String
Packing = m_sPacking
End Property

Public Property Let Packing(ByVal sNV As String)
m_sPacking = sNV
End Property
