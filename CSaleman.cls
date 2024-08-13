VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSaleman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CWorkers
Option Explicit

Private m_lID As Long
Private m_sSalemanID As String
Private m_sName As String
Private m_sIDCard As String
Private m_dBasicSalary As Double
Private m_dCommission As Double
Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private m_rsSaleman As ADODB.Recordset
'Dim m_objConnectDB As ConnectDB
Public Event DataValidated(bValid As Boolean)

Public Property Get ID() As Long
    ID = m_lID
End Property

Public Property Let SalemanID(ByVal vData As String)
    m_sSalemanID = vData
    ValidateData
End Property
Public Property Get SalemanID() As String
    SalemanID = m_sSalemanID
End Property
Public Property Let Name(ByVal vData As String)
    m_sName = vData
    ValidateData
End Property
Public Property Get Name() As String
    Name = m_sName
End Property
Public Property Let BasicSalary(ByVal vData As Double)
    m_dBasicSalary = vData
    ValidateData
End Property
Public Property Get BasicSalary() As Double
    BasicSalary = m_dBasicSalary
End Property
Public Property Let IDCard(ByVal vData As String)
    m_sIDCard = vData
    ValidateData
End Property
Public Property Get IDCard() As String
    IDCard = m_sIDCard
End Property
Public Property Let Commission(ByVal vData As Double)
    m_dCommission = vData
    ValidateData
End Property
Public Property Get Commission() As Double
    Commission = m_dCommission
End Property

Public Property Get IsValid() As Boolean
   IsValid = m_bIsDataValid
End Property

Private Sub ValidateData()
   Dim bValid As Boolean
   bValid = True
   
   bValid = bValid And m_sName <> ""
   m_bIsDataValid = bValid
   RaiseEvent DataValidated(bValid)
End Sub

Private Sub Class_Initialize()
'Set m_objConnectDB = New ConnectDB
   m_objConnectDB.Connect
   ValidateData
End Sub

Public Sub Init(Optional lID As Long = 0)
'    Set m_objConnectDB = New ConnectDB
   m_objConnectDB.Connect
   If lID = 0 Then
      m_bIsNewRecord = True
      OpenRs
      
      Exit Sub
    Else
      Dim m_RS As ADODB.Recordset
      Dim strRecordset As String
      Set m_RS = New ADODB.Recordset
      strRecordset = "Select * From Saleman where Saleman.ID=" & lID
      m_RS.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic

      
   End If
   
   m_lID = m_RS("ID") & ""
   m_sSalemanID = m_RS("SalemanID") & ""
   m_sName = m_RS("Name") & ""
   m_sIDCard = m_RS("IDCardNo") & ""
   m_dBasicSalary = m_RS("BasicSalary") & ""
   m_dCommission = m_RS("Commission") & ""
   m_RS.Close
   Set m_RS = Nothing
End Sub

Public Sub Save()
    OpenRs
   If m_bIsNewRecord Then
      m_rsSaleman.MoveLast
      m_rsSaleman.AddNew
   Else
      m_rsSaleman.Find "ID=" & m_lID 'To Be replaced with better code.
   End If
    m_rsSaleman("SalemanID") = m_sSalemanID
    m_rsSaleman("Name") = m_sName
    m_rsSaleman("IDCardNo") = m_sIDCard
    m_rsSaleman("BasicSalary") = m_dBasicSalary
    m_rsSaleman("Commission") = m_dCommission
    m_rsSaleman.Update
    m_rsSaleman.Requery
    m_lID = m_rsSaleman("ID")
    m_bIsNewRecord = False
End Sub

Public Sub OpenRs()
    Set m_rsSaleman = New ADODB.Recordset
    Dim strRecordset As String
    strRecordset = "Select * From Saleman Order By SalemanID"
    If m_rsSaleman.State <> adStateOpen Then
        m_rsSaleman.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub
Public Sub UpdateFormCombo(ByRef m_objCmb As ComboBox)
    Dim a As Integer
    m_objCmb.Clear
    a = 0
    OpenRs
    Do While Not m_rsSaleman.EOF
        m_objCmb.AddItem m_rsSaleman("Name")
        m_objCmb.ItemData(a) = m_rsSaleman("ID")
        a = a + 1
        m_rsSaleman.MoveNext
    Loop
End Sub
Public Function Get_SalemanName(ByVal lSaleman_ID As Long) As String
    OpenRs
    m_rsSaleman.Find "ID=" & lSaleman_ID
    Get_SalemanName = m_rsSaleman("Name")
End Function



