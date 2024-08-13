VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CWorkers
Option Explicit

Private m_lID As Long
Private m_sAreaID As String
Private m_sDescription As String
Private m_lZoneID As Long
Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private m_rsAreas As ADODB.Recordset
'Dim m_objConnectDB As ConnectDB
Public Event DataValidated(bValid As Boolean)

Public Property Get ID() As Long
    ID = m_lID
End Property

Public Property Let AreaID(ByVal vData As String)
    m_sAreaID = vData
    ValidateData
End Property
Public Property Get AreaID() As String
    AreaID = m_sAreaID
End Property
Public Property Let Description(ByVal vData As String)
    m_sDescription = vData
    ValidateData
End Property
Public Property Get Description() As String
    Description = m_sDescription
End Property
Public Property Let ZoneID(ByVal vData As Long)
    m_lZoneID = vData
    ValidateData
End Property
Public Property Get ZoneID() As Long
    ZoneID = m_lZoneID
End Property

Public Property Get IsValid() As Boolean
   IsValid = m_bIsDataValid
End Property

Private Sub ValidateData()
   Dim bValid As Boolean
   bValid = True
   
   bValid = bValid And m_sDescription <> ""
   m_bIsDataValid = bValid
   RaiseEvent DataValidated(bValid)
End Sub

Private Sub Class_Initialize()
'Set m_objConnectDB = New ConnectDB
   m_objConnectDB.Connect
   ValidateData
End Sub

Public Sub Init(Optional lID As Long = 0)
    'Set m_objConnectDB = New ConnectDB
   m_objConnectDB.Connect
   If lID = 0 Then
      m_bIsNewRecord = True
      OpenRs
      
      Exit Sub
    Else
      Dim m_RS As ADODB.Recordset
      Dim strRecordset As String
      Set m_RS = New ADODB.Recordset
      strRecordset = "Select * From Areas where Areas.ID=" & lID
      m_RS.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic

      
   End If
   
   m_lID = m_RS("ID") & ""
   m_sAreaID = m_RS("AreaID") & ""
   m_sDescription = m_RS("Description") & ""
   m_lZoneID = m_RS("ZoneID") & ""
   m_RS.Close
   Set m_RS = Nothing
End Sub

Public Sub Save()
    OpenRs
   If m_bIsNewRecord Then
      m_rsAreas.MoveLast
      m_rsAreas.AddNew
   Else
      m_rsAreas.Find "ID=" & m_lID 'To Be replaced with better code.
   End If
    m_rsAreas("AreaID") = m_sAreaID
    m_rsAreas("Description") = m_sDescription
     m_rsAreas("ZoneID") = m_lZoneID
    m_rsAreas.Update
    m_rsAreas.Requery
    m_lID = m_rsAreas("ID")
    m_bIsNewRecord = False
End Sub

Public Sub OpenRs()
    Set m_rsAreas = New ADODB.Recordset
    Dim strRecordset As String
    strRecordset = "Select * From Areas Order By AreaID"
    If m_rsAreas.State <> adStateOpen Then
        m_rsAreas.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub
Public Sub UpdateFormCombo(ByRef m_objCmb As ComboBox)
    Dim a As Integer
    m_objCmb.Clear
    a = 0
    OpenRs
    Do While Not m_rsAreas.EOF
        m_objCmb.AddItem m_rsAreas("Description")
        m_objCmb.ItemData(a) = m_rsAreas("ID")
        a = a + 1
        m_rsAreas.MoveNext
    Loop
End Sub



