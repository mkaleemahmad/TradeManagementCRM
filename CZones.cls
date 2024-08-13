VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CZones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CWorkers
Option Explicit

Private m_lID As Long
Private m_sZoneID As String
Private m_sDescription As String
Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private m_rsZones As ADODB.Recordset
'Dim m_objConnectDB As ConnectDB
Public Event DataValidated(bValid As Boolean)

Public Property Get ID() As Long
    ID = m_lID
End Property

Public Property Let ZoneID(ByVal vData As String)
    m_sZoneID = vData
    ValidateData
End Property
Public Property Get ZoneID() As String
    ZoneID = m_sZoneID
End Property
Public Property Let Description(ByVal vData As String)
    m_sDescription = vData
    ValidateData
End Property
Public Property Get Description() As String
    Description = m_sDescription
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
      strRecordset = "Select * From Zones where Zones.ID=" & lID
      m_RS.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic

      
   End If
   
   m_lID = m_RS("ID") & ""
   m_sZoneID = m_RS("ZoneID") & ""
   m_sDescription = m_RS("Description") & ""
   m_RS.Close
   Set m_RS = Nothing
End Sub

Public Sub Save()
    OpenRs
   If m_bIsNewRecord Then
      m_rsZones.MoveLast
      m_rsZones.AddNew
   Else
      m_rsZones.Find "ID=" & m_lID 'To Be replaced with better code.
   End If
    m_rsZones("ZoneID") = m_sZoneID
    m_rsZones("Description") = m_sDescription
    m_rsZones.Update
    m_rsZones.Requery
    m_lID = m_rsZones("ID")
    m_bIsNewRecord = False
End Sub

Public Sub OpenRs()
    Set m_rsZones = New ADODB.Recordset
    Dim strRecordset As String
    strRecordset = "Select * From Zones Order By ZoneID"
    If m_rsZones.State <> adStateOpen Then
        m_rsZones.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub
Public Sub UpdateFormCombo(ByRef m_objCmb As ComboBox)
    Dim a As Integer
    m_objCmb.Clear
    a = 0
    OpenRs
    Do While Not m_rsZones.EOF
        m_objCmb.AddItem m_rsZones("Description")
        m_objCmb.ItemData(a) = m_rsZones("ID")
        a = a + 1
        m_rsZones.MoveNext
    Loop
End Sub
Public Function Get_ZoneDescription(ByVal lZone_ID As Long) As String
    OpenRs
    m_rsZones.Find "ID=" & lZone_ID
    Get_ZoneDescription = m_rsZones("Description")
End Function


