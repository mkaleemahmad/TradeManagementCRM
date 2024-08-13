VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_lGroupID As Long
Dim m_sDesc As String
Dim m_rsAcctsGrp As New ADODB.Recordset
Dim m_bValid As Boolean

Public Property Get GroupID() As Long
GroupID = m_lGroupID
End Property

Public Property Get Description() As String
Description = m_sDesc
End Property

Public Property Let Description(ByVal sNewValue As String)
m_sDesc = sNewValue
End Property

Sub OpenRs(sCriteria As String)
Dim sSQL As String
sSQL = "SELECT * FROM AccountGroups" & IIf(sCriteria = "", "", " WHERE " & sCriteria)
With m_rsAcctsGrp
  If .State = adStateOpen Then .Close
  .Open sSQL, m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
  If .EOF Then
    m_bValid = False
  Else
    m_bValid = True
  End If
End With
End Sub

Public Sub Init(lGroupID As Long)
OpenRs "GroupID=" & lGroupID
With m_rsAcctsGrp
  If m_bValid Then
    m_lGroupID = !GroupID
    m_sDesc = !Description
  Else
    m_lGroupID = 0
    m_sDesc = ""
  End If
End With
End Sub

Public Sub Save(bNew As Boolean)
With m_rsAcctsGrp
  If bNew Then
    If .State = adStateClosed Then OpenRs ("")
    .AddNew
  End If
  ' !GroupID = m_lGroupID
  !Description = m_sDesc
  .Update
  If bNew Then
    OpenRs " GroupID = (Select Max(GroupID) From AccountGroups)"
    m_lGroupID = !GroupID
    m_sDesc = !Description
  End If
  m_bValid = True
End With
End Sub

Public Sub Remove()
Dim sSQL As String
sSQL = "DELETE FROM AccountGroups WHERE GroupID=" & m_lGroupID
m_objConnectDB.cnnMyshop.Execute sSQL
m_bValid = False
m_rsAcctsGrp.Close
End Sub

Public Function IsValid() As Boolean
IsValid = m_bValid
End Function

Sub PopComboBox(refCmbBx As ComboBox)
Dim k As Long, rs As New ADODB.Recordset
With rs
   If .State = adStateOpen Then .Close
   .Open "Select * From AccountGroups Order By Description", m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
   k = 0
   refCmbBx.Clear
   'refCmbBx.Text = ""
   Do Until .EOF
      refCmbBx.AddItem "" & !Description, k
      refCmbBx.ItemData(k) = !GroupID
      k = k + 1
      .MoveNext
   Loop
   If k > 0 Then refCmbBx.ListIndex = 0
   .Close
End With

End Sub

Sub Add2ComboBox(refCmbBx As ComboBox)
refCmbBx.AddItem m_sDesc
refCmbBx.ItemData(refCmbBx.NewIndex) = m_lGroupID
End Sub

Sub UpdateComboBox(refCmbBx As ComboBox)
Dim lInd As Long
lInd = TellIndexInDataItem(refCmbBx, m_lGroupID)
If lInd = -1 Then MsgBox "Unexpected error !", vbInformation: Exit Sub
refCmbBx.List(lInd) = m_sDesc
End Sub
Sub RemoveFromComboBox(refCmbBx As ComboBox)
If refCmbBx.ListCount < 1 Then MsgBox "Unexpected error !", vbInformation: Exit Sub
refCmbBx.RemoveItem TellIndexInDataItem(refCmbBx, m_lGroupID)
End Sub

Function CanDelete(lID As Long) As Boolean
Dim sSQL As String
sSQL = "SELECT COUNT(*) AS TotRecs FROM Accounts WHERE GroupID=" & lID
Dim rsTmp As ADODB.Recordset
Set rsTmp = m_objConnectDB.cnnMyshop.Execute(sSQL)
CanDelete = IsNull2(rsTmp!TotRecs, 0) = 0
End Function

