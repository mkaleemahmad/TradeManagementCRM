VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSubGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_lGroupID As Long
Dim m_lSubGroupID As Long
Dim m_sDesc As String

Dim m_rsSubAcctsGrp As New ADODB.Recordset
Dim m_bValid As Boolean

Public Property Get GroupID() As Long
GroupID = m_lGroupID
End Property

Public Property Let GroupID(ByVal lNewValue As Long)
m_lGroupID = lNewValue
End Property

Public Property Get Description() As String
Description = m_sDesc
End Property

Public Property Let Description(ByVal sNewValue As String)
m_sDesc = sNewValue
End Property

Sub OpenRs(sCriteria As String)
Dim sSQL As String
sSQL = "SELECT * FROM AccountSubGroups" & IIf(sCriteria = "", "", " WHERE " & sCriteria)
With m_rsSubAcctsGrp
  If .State = adStateOpen Then .Close
  .Open sSQL, m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
  If .EOF Then
    m_bValid = False
  Else
    m_bValid = True
  End If
End With
End Sub

Public Sub Init(lSubGroupID As Long, Optional lGroupID)
Dim sSubSQL As String
If Not IsMissing(lGroupID) Then
  sSubSQL = " AND GroupID=" & lGroupID
Else
  sSubSQL = ""
End If
OpenRs "SubGroupID=" & lSubGroupID & sSubSQL
With m_rsSubAcctsGrp
  If m_bValid Then
    m_lGroupID = !GroupID
    m_sDesc = !Description
    m_lSubGroupID = !SubGroupID
  Else
    m_lGroupID = 0
    m_sDesc = ""
    m_lSubGroupID = 0
  End If
End With
End Sub

Public Sub Save(bNew As Boolean)
With m_rsSubAcctsGrp
  If bNew Then
    If .State = adStateClosed Then OpenRs ("")
    .AddNew
  End If
    !GroupID = m_lGroupID
    !Description = m_sDesc
    .Update
  If bNew Then
    Dim sSQL As String
    sSQL = " SubGroupID=(Select Max(SubGroupID) From AccountSubGroups)"
    OpenRs sSQL
    m_lGroupID = !GroupID
    m_sDesc = !Description
    m_lSubGroupID = !SubGroupID
  End If
  m_bValid = True
End With
End Sub

Public Sub Remove(Optional lGroupID)
Dim sSQL As String
Dim sSubSQL As String
If Not IsMissing(lGroupID) Then
  sSubSQL = " AND GroupID=" & lGroupID
Else
  sSubSQL = ""
End If
sSQL = "DELETE FROM AccountSubGroups WHERE SubGroupID=" & m_lSubGroupID & sSubSQL
m_objConnectDB.cnnMyshop.Execute sSQL
m_bValid = False
m_rsSubAcctsGrp.Close
End Sub

Public Function IsValid() As Boolean
IsValid = m_bValid
End Function

Sub PopComboBox(refCmbBx As ComboBox, Optional lParentGroupID)
Dim k As Long, rs As New ADODB.Recordset, bMissing As Boolean
Dim sSubSQL As String
bMissing = IsMissing(lParentGroupID)
If Not bMissing Then
  sSubSQL = " WHERE GroupID=" & lParentGroupID
End If
With rs
   If .State = adStateOpen Then .Close
   .Open "Select * From AccountSubGroups" & IIf(bMissing, "", sSubSQL) & " Order By Description", m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
   k = 0
   refCmbBx.Clear
   Do Until .EOF
      refCmbBx.AddItem "" & !Description, k
      refCmbBx.ItemData(k) = !SubGroupID
      k = k + 1
      .MoveNext
   Loop
   If k > 0 Then refCmbBx.ListIndex = 0
   .Close
End With

End Sub



Public Property Get SubGroupID() As Long
SubGroupID = m_lSubGroupID
End Property

Sub Add2ComboBox(refCmbBx As ComboBox)
refCmbBx.AddItem m_sDesc
refCmbBx.ItemData(refCmbBx.NewIndex) = m_lSubGroupID
End Sub

Sub UpdateComboBox(refCmbBx As ComboBox)
Dim lInd As Long
lInd = TellIndexInDataItem(refCmbBx, m_lSubGroupID)
If lInd = -1 Then MsgBox "Unexpected error !", vbInformation: Exit Sub
refCmbBx.List(lInd) = m_sDesc
End Sub
Sub RemoveFromComboBox(refCmbBx As ComboBox)
If refCmbBx.ListCount < 1 Then MsgBox "Unexpected error !", vbInformation: Exit Sub
Dim lInd As Long
lInd = TellIndexInDataItem(refCmbBx, m_lSubGroupID)
refCmbBx.RemoveItem lInd
End Sub


Function CanDelete(lID As Long) As Boolean
Dim sSQL As String
sSQL = "SELECT COUNT(*) AS TotRecs FROM Accounts WHERE SubGroupID=" & lID
Dim rsTmp As ADODB.Recordset
Set rsTmp = m_objConnectDB.cnnMyshop.Execute(sSQL)
CanDelete = IsNull2(rsTmp!TotRecs, 0) = 0
End Function
