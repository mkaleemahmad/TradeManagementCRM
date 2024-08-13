VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTransLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim m_rsLgEvnt As New ADODB.Recordset
Dim m_bValidInDB As Boolean

Dim m_lSessionID As Long
Dim m_dtLogDateTime As Date
Dim m_sEventName As String
Dim m_lUserID As Long
Dim m_sTransType As String
Dim m_lTransID As Long

' int - 4 - auto
Public Property Get SessionID() As Long
SessionID = m_lSessionID
End Property

' Session ID can not be modified
'Public Property Let SessionID(ByVal lNV As Long)
'm_lSessionID = lNV
'End Property
' datetime - 8
Public Property Get LogDateTime() As Date
LogDateTime = m_dtLogDateTime
End Property

'Public Property Let LogDateTime(ByVal dtNV As Date)
'm_dtLogDateTime = dtNV
'End Property

' nvarchar - 50
Public Property Get EventName() As String
EventName = m_sEventName
End Property

Public Property Let EventName(ByVal sNV As String)
m_sEventName = sNV
End Property

' int - 4
Public Property Get UserID() As Long
UserID = m_lUserID
End Property

'Public Property Let UserID(ByVal lNV As Long)
'm_lUserID = lNV
'End Property

' nvarchar - 300
Public Property Get TransType() As String
TransType = m_sTransType
End Property

Public Property Let TransType(ByVal sNV As String)
m_sTransType = sNV
End Property

' nvarchar - 150
Public Property Get TransID() As String
TransID = m_lTransID
End Property

Public Property Let TransID(ByVal sNV As String)
m_lTransID = sNV
End Property

Sub OpenRs(Optional sCriteria As String = "")
Dim sSQL As String
sSQL = "SELECT * FROM TransLog"
If sCriteria = "" Then
Else
    sSQL = sSQL & " Where " & sCriteria
End If
sSQL = sSQL & " Order by SessionID Desc;"
With m_rsLgEvnt
    If .State = adStateOpen Then .Close
    .CacheSize = 5
    .Open sSQL, m_objConnectDB.cnnMyshop, adOpenKeyset, adLockOptimistic
End With
End Sub

Private Function HasRecord() As Boolean
With m_rsLgEvnt
    If .RecordCount > 0 Then
        HasRecord = True
    Else
        HasRecord = False
    End If
End With
End Function


Public Sub Init(lUserID As Long)
Dim sCriteria As String
sCriteria = "UserID=" & lUserID
OpenRs sCriteria
If HasRecord Then
    UpdateData True
    m_bValidInDB = True
Else
    m_bValidInDB = False
End If
End Sub

Private Sub UpdateData(bUpdate_Variables As Boolean)
With m_rsLgEvnt
    If bUpdate_Variables Then
        m_lSessionID = NumVal(!SessionID)
        If Not IsNull(!LogDateTime) Then
            m_dtLogDateTime = !LogDateTime
        Else
            m_dtLogDateTime = CDate(#1/1/1900#) + Time
        End If
        m_sEventName = SoftVal(!EventName)
        
        m_sTransType = SoftVal(!TransType)
        m_lTransID = SoftVal(!TransID)
        
        m_lUserID = SoftVal(!UserID)
        
        m_bValidInDB = True
    Else
        '!SessionID = m_lSessionID
        !EventName = m_sEventName

        !TransType = m_sTransType
        !TransID = m_lTransID

            ' Self Detecting
        !LogDateTime = Date + Time ' m_dtLogDateTime
        !UserID = frmLogin.UserID
        
    End If
End With
End Sub


Public Sub Save(bNew As Boolean)
If bNew Then
    m_rsLgEvnt.AddNew
End If
UpdateData False
m_rsLgEvnt.Update
m_bValidInDB = True
End Sub

Public Sub Remove()
If HasRecord Then
    m_rsLgEvnt.Delete
    m_bValidInDB = False
Else
    Err.Raise vbObjectError + 1, "Remove Failed"
End If
End Sub

Public Function IsValidInDB() As Boolean
IsValidInDB = m_bValidInDB
End Function

Private Sub Class_Terminate()
If m_rsLgEvnt.State = adStateOpen Then
    m_rsLgEvnt.Close
End If
Set m_rsLgEvnt = Nothing
End Sub


Sub PopListView(refLvw As ListView)
Dim k As Long, rs As ADODB.Recordset, sSQL As String
Dim fld As ADODB.Field
Set rs = m_rsLgEvnt.Clone
With rs
  With refLvw
    .ColumnHeaders.Clear
    .ListItems.Clear
    Dim clmnHr As ColumnHeader
      With .ColumnHeaders
        Set clmnHr = .Add(, "EventID", "Event No.")
        clmnHr.Width = 0.66 * clmnHr.Width
'        clmnHr.Alignment = lvwColumnRight
        Set clmnHr = .Add(, "EventDate", "Event Date")
        clmnHr.Width = 1.24 * clmnHr.Width
        Set clmnHr = .Add(, "EventTime", "Event Time")
        clmnHr.Width = 0.8 * clmnHr.Width

        .Add , "EventName", "Event"
        .Add , "TransType", "Transaction"
        Set clmnHr = .Add(, "TransID", "Trans No.")
        clmnHr.Alignment = lvwColumnRight
      End With
  End With
  
  Dim li As ListItem
  Dim lis As ListSubItem
  Dim bFirst As Boolean
  If Not (.BOF And .EOF) Then .MoveFirst
  Do Until .EOF
    Set li = refLvw.ListItems.Add(, "SessionID=" & !SessionID, "" & SoftVal(!SessionID))
    If IsNull(!LogDateTime) Then
        li.ListSubItems.Add , "LogDate", ""
    Else
        li.ListSubItems.Add , "LogDate", FormatDateTime(!LogDateTime, vbLongDate)
    End If
    li.ListSubItems.Add , "LogTime", FormatDateTime(!LogDateTime, vbLongTime)
    li.ListSubItems.Add , "EventName", SoftVal(!EventName)
    li.ListSubItems.Add , "TransType", MnTransType(SoftVal(!TransType))
    li.ListSubItems.Add , "TransID", SoftVal(!TransID)
    .MoveNext
   Loop
   If refLvw.ListItems.Count > 0 Then Set refLvw.SelectedItem = refLvw.ListItems.Item(1)
   .Close
End With
End Sub

Function MnTransType(sTransType As String) As String
Dim s As String
Select Case UCase(sTransType)
Case "OS"
    s = "Opening Stock"
Case "OB"
    s = "Opening Balance"
Case "PH"
    s = "Purchase"
Case "PR"
    s = "Purchase Return"
Case "SL"
    s = "Sales Invoice"
Case "SR"
    s = "Sales Return"
Case "CV"
    s = "Customer Voucher"
Case "AV"
    s = "Accounts Voucher"
Case "CP"
    s = "Cash Payment"
Case "CR"
    s = "Cash Receipt"
Case "BD"
    s = "Bank Deposit"
Case "BP"
    s = "Bank Payment"
Case "PD"
    s = "Production Voucher"
Case Else
    s = UCase(sTransType)
End Select
MnTransType = s
End Function
