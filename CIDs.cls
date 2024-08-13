VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CIDs
Option Explicit

Private m_sID As String
Private m_sDescription As String
Private m_bIsDataValid As Boolean
Private m_bPrintVoucher As Boolean
Private m_rsIDs As ADODB.Recordset
Public Event DataValidated(bValid As Boolean)

Public Property Get ID() As String
    ID = m_sID
End Property
Public Function IsPrintVoucher(sVoucherType As String) As Boolean
    OpenRs sVoucherType
    IsPrintVoucher = IsNull2(m_rsIDs("PrintVoucher"), 0)
End Function

Public Sub OpenRs(Optional sDescription As String)
    Set m_rsIDs = New ADODB.Recordset
    Dim strRecordset As String
    If sDescription = "" Then
        strRecordset = "Select * from IDs order by IDDescription"
    Else
        strRecordset = "Select * from IDs WHERE IDs.IDDescription = '" & sDescription & "'"
    End If
    If m_rsIDs.State <> adStateOpen Then
        m_rsIDs.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenKeyset, adLockOptimistic
    End If
End Sub

Public Sub NewID(IDDesc As String)
    OpenRs IDDesc
'    m_rsIDs.MoveFirst
'    m_rsIDs.Find "IDDescription=" & IDDesc
'    Select Case IDDesc
'        Case "OS"
'            m_rsIDs.MoveFirst
'        Case "PH"
'            m_rsIDs.Move (2)
'        Case "PR"
'            m_rsIDs.Move (3)
'        Case "SL"
'            m_rsIDs.Move (3)
'        Case "SR"
'            m_rsIDs.Move (4)
'        Case "CV"
'            m_rsIDs.Move (5)
'    End Select
'    m_rsIDs.Find "TRIM(IDDescription)=" & IDDesc
    If Not m_rsIDs.BOF And Not m_rsIDs.EOF Then
        m_sID = Trim(m_rsIDs("IDVALUE"))
        m_bPrintVoucher = IsNull2(m_rsIDs("PrintVoucher"), 0)
     '   m_rs("IDVALUE") = Trim(Str(Val(m_sID) + 1))
      '  m_rs.Update
      OpenRs
      Do While Not m_rsIDs.EOF
        If m_rsIDs("IDDescription") = IDDesc Then
            m_rsIDs("IDVALUE") = Trim(Str(Val(m_sID) + 1))
            m_rsIDs.Update
            Exit Do
        Else
            m_rsIDs.MoveNext
        End If
      Loop
    End If
End Sub

Public Sub UpDatePrintState(sVoucherType As String, bVal As Boolean)
    OpenRs sVoucherType
    m_rsIDs("PrintVoucher") = bVal
    m_rsIDs.Update
    m_rsIDs.Close
End Sub
