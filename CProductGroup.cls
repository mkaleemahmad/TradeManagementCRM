VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProductGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CProductGroup

Option Explicit
Private m_lID As Long
Private m_sGroupID As String
Private m_sDescription As String
Private m_lCategoryID As Long
Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private m_rsProductGroup As ADODB.Recordset
Public Event DataValidated(bValid As Boolean)

Public Property Let ID(ByVal vData As Integer)
    m_lID = vData
    ValidateData
End Property
Public Property Get ID() As Integer
    ID = m_lID
End Property

Public Property Let GroupID(ByVal vData As String)
    m_sGroupID = vData
    ValidateData
End Property
Public Property Get GroupID() As String
    GroupID = m_sGroupID
End Property

Public Property Let Description(ByVal vData As String)
    m_sDescription = vData
    ValidateData
End Property
Public Property Get Description() As String
    Description = m_sDescription
End Property

Public Property Let CategoryID(ByVal vData As Long)
    m_lCategoryID = vData
    ValidateData
End Property
Public Property Get CategoryID() As Long
    CategoryID = m_lCategoryID
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
   
  ' bValid = bValid And m_sGroupID <> ""
   bValid = bValid And m_sDescription <> ""
   m_bIsDataValid = bValid
   RaiseEvent DataValidated(bValid)
End Sub

Private Sub Class_Initialize()
   ValidateData
End Sub

Public Sub Init(Optional iID As Integer = 0)
   OpenRs
   If iID = 0 Then
      m_bIsNewRecord = True
      Exit Sub
      Else
      m_rsProductGroup.Find "ID=" & iID
   End If
   m_lID = m_rsProductGroup("ID") & ""
   m_sGroupID = m_rsProductGroup("GroupID") & ""
   m_sDescription = m_rsProductGroup("Description") & ""
   m_lCategoryID = IsNull2(m_rsProductGroup("CategoryID"), 0)
End Sub

Public Sub Save()
   If m_bIsNewRecord Then
    OpenRs
    m_rsProductGroup.AddNew
   End If
   
   m_rsProductGroup("GroupID") = m_sGroupID
   m_rsProductGroup("Description") = m_sDescription
   m_rsProductGroup("CategoryID") = m_lCategoryID
   m_rsProductGroup.Update
   m_rsProductGroup.Requery
   m_lID = m_rsProductGroup("ID")
   m_bIsNewRecord = False
End Sub

Public Sub OpenRs(Optional lPCategory As Long = 0)
    Set m_rsProductGroup = New ADODB.Recordset
    Dim strRecordset As String
    If lPCategory = 0 Then
    strRecordset = "Select * from ProductGroups Order By Description"
    Else
    strRecordset = "Select * from ProductGroups Where CategoryID = " & lPCategory & " Order By Description"
    End If
    If m_rsProductGroup.State <> adStateOpen Then
        m_rsProductGroup.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub

Public Sub UpdateFormCombo(ByRef m_objCmb As ComboBox, Optional lPCategory As Long = 0)
    OpenRs lPCategory
    m_objCmb.Clear
    Dim a As Integer
    Do While Not m_rsProductGroup.EOF
        m_objCmb.AddItem m_rsProductGroup("Description")
        m_objCmb.ItemData(a) = m_rsProductGroup("ID")
        a = a + 1
        m_rsProductGroup.MoveNext
    Loop
End Sub

Public Sub Remove()
If Not m_bIsNewRecord Then m_rsProductGroup.Delete
End Sub

Function CanDelete(lID As Long) As Boolean
Dim sSQL As String
sSQL = "SELECT COUNT(*) AS TotRecs FROM Products WHERE GroupID=" & lID
Dim rsTmp As ADODB.Recordset
Set rsTmp = m_objConnectDB.cnnMyshop.Execute(sSQL)
CanDelete = IsNull2(rsTmp!TotRecs, 0) = 0
End Function
