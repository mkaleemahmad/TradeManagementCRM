VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProductCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CProductCategory

Option Explicit
Private m_lID As Long
Private m_sCategoryID As String
Private m_sDescription As String
Private m_bIsNewRecord As Boolean
Private m_bIsDataValid As Boolean
Private m_rsProductCategories As ADODB.Recordset
Public Event DataValidated(bValid As Boolean)

Public Property Let ID(ByVal vData As Integer)
    m_lID = vData
    ValidateData
End Property
Public Property Get ID() As Integer
    ID = m_lID
End Property

Public Property Let CategoryID(ByVal vData As String)
    m_sCategoryID = vData
    ValidateData
End Property
Public Property Get CategoryID() As String
    CategoryID = m_sCategoryID
End Property

Public Property Let Description(ByVal vData As String)
    m_sDescription = vData
    ValidateData
End Property
Public Property Get Description() As String
    Description = m_sDescription
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
   
  ' bValid = bValid And m_sCategoryID <> ""
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
      m_rsProductCategories.Find "ID=" & iID
   End If
   m_lID = m_rsProductCategories("ID") & ""
   m_sCategoryID = m_rsProductCategories("CategoryID") & ""
   m_sDescription = m_rsProductCategories("Description") & ""
End Sub

Public Sub Save()
   If m_bIsNewRecord Then
    OpenRs
    m_rsProductCategories.AddNew
   End If
   
   ' m_rsProductCategories("CategoryID") = m_sCategoryID 'USER defined code not required
   m_rsProductCategories("Description") = m_sDescription
   'm_rsProductCategories("CategoryID") = m_sCategoryID
   m_rsProductCategories.Update
   m_rsProductCategories.Requery
   m_lID = m_rsProductCategories("ID")
   m_bIsNewRecord = False
End Sub

Public Sub OpenRs()
    Set m_rsProductCategories = New ADODB.Recordset
    Dim strRecordset As String
    strRecordset = "Select * from ProductCategories Order By Description"
    If m_rsProductCategories.State <> adStateOpen Then
        m_rsProductCategories.Open strRecordset, m_objConnectDB.cnnMyshop, adOpenDynamic, adLockOptimistic
    End If
End Sub

Public Sub UpdateFormCombo(ByRef m_objCmb As ComboBox)
    OpenRs
    m_objCmb.Clear
    Dim a As Integer
    Do While Not m_rsProductCategories.EOF
        m_objCmb.AddItem m_rsProductCategories("Description")
        m_objCmb.ItemData(a) = m_rsProductCategories("ID")
        a = a + 1
        m_rsProductCategories.MoveNext
    Loop
End Sub

Function CanDelete(lID As Long) As Boolean
Dim sSQL As String
sSQL = "SELECT COUNT(*) AS TotRecs FROM Products WHERE CategoryID=" & lID
Dim rsTmp As ADODB.Recordset
Set rsTmp = m_objConnectDB.cnnMyshop.Execute(sSQL)
CanDelete = IsNull2(rsTmp!TotRecs, 0) = 0
End Function

Sub Delete()
If Not m_bIsNewRecord Then m_rsProductCategories.Delete
End Sub

Public Function CategoryExist(sDesc As String) As Boolean
    
    Dim sSQL As String
    sSQL = "SELECT COUNT(*) AS TotRecs FROM ProductCategories WHERE Description Like'" & Trim(sDesc) & "'"
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = m_objConnectDB.cnnMyshop.Execute(sSQL)
    CategoryExist = IsNull2(rsTmp!TotRecs, 0)
    
End Function

