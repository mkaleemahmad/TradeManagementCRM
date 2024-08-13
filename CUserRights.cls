VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUserRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_bAdd As Boolean
Dim m_bView As Boolean
Dim m_bEdit As Boolean
Dim m_bDelete As Boolean

Dim m_sRightDesc As String
Dim m_lUserID As Long

Dim m_rsUR As New ADODB.Recordset

Dim m_bEIDb As Boolean

Public Property Get ExistInDB() As Boolean
ExistInDB = m_bEIDb
End Property

Public Property Get CanAdd() As Boolean
CanAdd = m_bAdd
End Property

Public Property Let CanAdd(ByVal bNewValue As Boolean)
m_bAdd = bNewValue
End Property

Public Property Get CanView() As Boolean
CanView = m_bView
End Property

Public Property Let CanView(ByVal bNewValue As Boolean)
m_bView = bNewValue
End Property

Public Property Get CanEdit() As Boolean
CanEdit = m_bEdit
End Property

Public Property Let CanEdit(ByVal bNewValue As Boolean)
m_bEdit = bNewValue
End Property

Public Property Get CanDelete() As Boolean
CanDelete = m_bDelete
End Property

Public Property Let CanDelete(ByVal bNewValue As Boolean)
m_bDelete = bNewValue
End Property

Public Property Get ObjectDescription() As String
ObjectDescription = m_sRightDesc
End Property

Sub Init(lUserID As Long, lRightNumber As Long)
OpenRs "UserID=" & lUserID & " And RightNo=" & lRightNumber
With m_rsUR
   If .EOF Then
      m_bAdd = False
      m_bDelete = False
      m_bEdit = False
      m_bView = False
   Else
      m_bAdd = ![CanAdd]
      m_bDelete = ![CanDelete]
      m_bEdit = ![CanEdit]
      m_bView = ![CanView]
   End If
End With
'm_sRightDesc = sRight
m_lUserID = lUserID
m_rsUR.Close
End Sub

Private Sub OpenRs(sCriteria As String)
Dim strRecordset As String
With m_rsUR
   If .State = adStateOpen Then .Close
   strRecordset = "SELECT * FROM UserRights WHERE " & sCriteria
   .Open "Select * From UserRights Where " & sCriteria, m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
   If .EOF Then
      m_bEIDb = False
   Else
      m_bEIDb = True
   End If
End With
End Sub

Sub Save(bNew As Boolean, lUserID As Long, sRight As String)
OpenRs " UserID=" & lUserID & " And RightDescription='" & sRight & "';"
With m_rsUR
   If bNew Then .AddNew
   ![UserID] = lUserID
   ![RightDescription] = sRight
   ![CanAdd] = m_bAdd
   ![CanEdit] = m_bEdit
   ![CanView] = m_bView
   ![CanDelete] = m_bDelete
   .Update
   .Close
End With
End Sub

Sub Remove(lUserID As Long, sRight As String)
OpenRs "UserID=" & lUserID & " And RightDescription='" & sRight & "'"
With m_rsUR
   If .EOF Then
   Else
      .Delete
   End If
   .Close
End With
End Sub

