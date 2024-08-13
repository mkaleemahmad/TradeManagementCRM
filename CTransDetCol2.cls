VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTransDetCol2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CTransDetCol

Option Explicit
Private m_Col As Collection

Private Const ITEMPREFIX = "Item"

Public Sub AddC(objNewMember As CTransDet, iLoc As Integer)
   m_Col.Add objNewMember, ITEMPREFIX & Str(iLoc)
End Sub

Public Property Get Item(vntIndexKey As Variant) As CTransDet
   Set Item = m_Col(vntIndexKey)
End Property

Public Property Get Count() As Long
   Count = m_Col.Count
End Property

Public Sub Remove(iColIndex As Long)
   m_Col.Remove ITEMPREFIX & Str(iColIndex)
End Sub

Public Property Get NewEnum() As IUnknown
   Set NewEnum = m_Col.[_NewEnum]
End Property

Private Sub Class_Initialize()
   Set m_Col = New Collection
End Sub

Private Sub Class_Terminate()
   Dim objLoopDetail As CTransDet
   For Each objLoopDetail In m_Col
      Set objLoopDetail = Nothing
   Next objLoopDetail
   Set m_Col = Nothing
End Sub

Public Sub Init(dTransDt As Date)
    Dim objDetail As CTransDet
    Dim cmdCTransDetsCol As ADODB.Command
    Dim pmdTransDate As ADODB.Parameter
    
    Dim rsCTransDetss As ADODB.Recordset
    Set cmdCTransDetsCol = New ADODB.Command
    Set pmdTransDate = New ADODB.Parameter
    With cmdCTransDetsCol
        Set .ActiveConnection = m_objConnectDB.cnnMyshop
        .CommandType = adCmdStoredProc 'adCmdTable
        .CommandText = "qryGetTranss"
        .Parameters.Refresh
        .Parameters(1) = dTransDt
    End With
        
    Set rsCTransDetss = cmdCTransDetsCol.Execute
    
    Dim a As Integer
    a = 1
    Do While Not rsCTransDetss.EOF
      Set objDetail = New CTransDet
'      objDetail.Init rsCTransDetss("TransID"), rsCTransDetss("TransType")
      objDetail.ID = rsCTransDetss("TransID")
      objDetail.TransType = rsCTransDetss("TransType")
      objDetail.AccountID = rsCTransDetss("AccountNumber")
      objDetail.Description = IsNull2(rsCTransDetss("Description"), "")
      objDetail.Debit = rsCTransDetss("Debit")
      objDetail.Credit = rsCTransDetss("Credit")
      AddC objDetail, a
      a = a + 1
      rsCTransDetss.MoveNext
    Loop
    rsCTransDetss.Close
    Set rsCTransDetss = Nothing
End Sub

Public Sub Save(IsNewRecord As Boolean, lID As Long)
   If Count > 0 Then
        Dim objDetail As CTransDet
        For Each objDetail In m_Col
           objDetail.Save (IsNewRecord), (lID)
        Next objDetail
        
        Set m_Col = New Collection
    Else
        MsgBox "Nothing found to save", vbCritical + vbOKOnly
    End If
End Sub


