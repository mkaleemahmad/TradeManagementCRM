VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTransDets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class CTransDets

Option Explicit
Option Compare Text
Private m_Col As Collection
Private m_sTransID As String
Private m_sTransType As String

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
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = m_Col.[_NewEnum]
End Property

Private Sub Class_Initialize()
   Set m_Col = New Collection
End Sub

Private Sub Class_Terminate()

   '
   ' Just to make sure the memory is cleaned up,
   ' each object in the collection is cleared
   ' before the collection object itself is
   ' cleared.
   '
   Dim objLoopDetail As CTransDet
   
   For Each objLoopDetail In m_Col
      Set objLoopDetail = Nothing
   Next objLoopDetail
   Set m_Col = Nothing

End Sub

Public Sub Init(lTransID As Long, sTransType As String)
    Dim objDetail As CTransDet
    Dim cmdCTDets As ADODB.Command
    Dim paramlTransId As ADODB.Parameter
    Dim pmsTransType As ADODB.Parameter
    
    Dim rsCTransDets As ADODB.Recordset
    Set cmdCTDets = New ADODB.Command
    Set rsCTransDets = New ADODB.Recordset
    Set paramlTransId = New ADODB.Parameter
    Set pmsTransType = New ADODB.Parameter
    With cmdCTDets
        Set .ActiveConnection = m_objConnectDB.cnnMyshop
        .CommandType = adCmdStoredProc 'adCmdTable
        .CommandText = "qryGetTransDets"
    End With
    Set paramlTransId = cmdCTDets.CreateParameter("TransID", adInteger, adParamInput, , lTransID)
'    Set pmsTransType = cmdCTDets.CreateParameter("paramsTransType", adChar, adParamInput, , "'" & sTransType & "'")
    cmdCTDets.Parameters.Append paramlTransId
    'cmdCTDets.Parameters.Append pmsTransType
    'cmdCTDets.Parameters.Refresh
    'cmdCTDets.Parameters.Count
    
    Set rsCTransDets = cmdCTDets.Execute
    
    Dim a As Integer
    a = 1
    Do While Not rsCTransDets.EOF
        If rsCTransDets("TransType") = sTransType Then
            Set objDetail = New CTransDet
            objDetail.ID = rsCTransDets("TransID")
            objDetail.TransType = rsCTransDets("TransType")
            objDetail.AccountID = rsCTransDets("AccountNumber")
            objDetail.Description = rsCTransDets("Description")
            objDetail.Debit = rsCTransDets("Debit")
            objDetail.Credit = rsCTransDets("Credit")
            'objDetail.Init rsCTransDets("TransID"), rsCTransDets("TransType")
            AddC objDetail, a
            a = a + 1
        End If
      rsCTransDets.MoveNext
    Loop
'    rsCTransDets.Close
'    Set rsCTransDets = Nothing
End Sub

Public Sub Save(IsNewRecord As Boolean, lID As Long)
   If Count > 0 Then
        Dim objDetail As CTransDet
        If (Not IsNewRecord) And m_Col.Count > 0 Then
          Set objDetail = m_Col(1)
           objDetail.DeleteRec lID, objDetail.TransType
        End If
        For Each objDetail In m_Col
           objDetail.Save True, (lID)  '(IsNewRecord), (lID)
        Next ' objDetail
        Set m_Col = New Collection
    Else
        MsgBox "Nothing found to save", vbCritical + vbOKOnly
    End If
End Sub

Property Get TotalDebit() As Currency
   Dim cTemp As Currency
   Dim objDetail As CTransDet
   cTemp = 0
   For Each objDetail In m_Col
      cTemp = cTemp + objDetail.Debit
   Next objDetail
   TotalDebit = cTemp
End Property

Property Get TotalCredit() As Currency
   Dim cTemp As Currency
   Dim objDetail As CTransDet
   cTemp = 0
   For Each objDetail In m_Col
      cTemp = cTemp + objDetail.Credit
   Next objDetail
   TotalCredit = cTemp
End Property



