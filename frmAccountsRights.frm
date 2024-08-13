VERSION 5.00
Begin VB.Form frmAccountsRights 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AccountRights"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   4590
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstRights 
      Height          =   5235
      ItemData        =   "frmAccountsRights.frx":0000
      Left            =   0
      List            =   "frmAccountsRights.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   0
      Width           =   4545
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Left            =   1875
      TabIndex        =   1
      Top             =   5340
      Width           =   1260
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   3285
      TabIndex        =   0
      Top             =   5340
      Width           =   1260
   End
End
Attribute VB_Name = "frmAccountsRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lUserID As Long
Public sUserName As String
Dim rs As ADODB.Recordset
Private bChanged As Boolean
Private m_Rl As Long

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdSave_Click()
 SaveAccountsRights
 MsgBox "Accounts Rights saved successfully.", vbInformation
 bChanged = False
End Sub

Private Sub Form_Load()
 Dim m_Groupid As Long
 Set rs = New ADODB.Recordset
 Dim a As Integer
 Dim sSQL As String
     sSQL = "Select AG.GroupID,AG.Description,IsNull((Select Isnull(CanView,0) From AccountsRights Where GroupID = AG.GroupID and UserID =" & lUserID & "),0) as CanView from AccountGroups AG"
     rs.Open sSQL, m_objConnectDB.cnnMyshop, adOpenKeyset, adLockOptimistic
     Do Until rs.EOF
         lstRights.AddItem rs("Description") & ""
         lstRights.ItemData(a) = rs("GroupID")
         lstRights.Selected(a) = rs("CanView")
         rs.MoveNext
         a = a + 1
     Loop
   Me.Caption = "Accounts Rights for " & sUserName
   bChanged = False
End Sub

Sub IndicateAccountsRights()
   '    Dim l As Long
   '    Dim sSQL As String
   '    Dim rsTemp As New ADODB.Recordset
   '    sSQL = "Select GroupID,CanView From AccountsRights where UserID=" & lUserID
   '    rsTemp.Open sSQL, m_objConnectDB.cnnMyshop
   ''    For l = 0 To lstRights.ListCount - 1
   ''        lstRights.Selected(l) = False
   ''    Next
   '    rsTemp.MoveFirst
   '    For l = 0 To lstRights.ListCount - 1
   '        rsTemp.Find "GroupID=" & lstRights.ItemData(l)
   '        lstRights.Selected(l) = rs!CanView
   '        rs.MoveNext
   '    Next
End Sub

Sub SaveAccountsRights()
    Dim I As Long
    Dim sSQL As String
    'Delete Previous Values from Table
    sSQL = "DELETE FROM AccountsRights where UserID=" & lUserID
    m_objConnectDB.cnnMyshop.Execute sSQL
    'Add new entries in the table
    For I = 0 To lstRights.ListCount - 1
        sSQL = "INSERT INTO ACCOUNTSRIGHTS (USERID,GROUPID,CANVIEW) VALUES(" & lUserID & "," & lstRights.ItemData(I) & "," & IIf(lstRights.Selected(I) = True, 1, 0) & ")"
        m_objConnectDB.cnnMyshop.Execute sSQL
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim ret As VbMsgBoxResult
If bChanged Then
    ret = MsgBox("Accounts Rights may be changed, Save them No ?", vbQuestion + vbYesNoCancel)
    If ret = vbCancel Then
        Cancel = 1
    ElseIf ret = vbYes Then
        SaveAccountsRights
    End If
End If
End Sub

Private Sub lstRights_ItemCheck(Item As Integer)
bChanged = True
End Sub

