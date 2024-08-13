VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmAccounts 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Accounts"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7290
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   7290
   Begin VB.Timer tmrNotify 
      Interval        =   1000
      Left            =   4350
      Top             =   3555
   End
   Begin MSComctlLib.StatusBar staDetails 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4935
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Total Accounts:"
            TextSave        =   "Total Accounts:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Time Elapsed:"
            TextSave        =   "Time Elapsed:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Data Rate:"
            TextSave        =   "Data Rate:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "4:00 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvAccounts 
      Height          =   2385
      Left            =   270
      TabIndex        =   2
      Top             =   885
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   4207
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   1587
      LabelEdit       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      Appearance      =   1
   End
   Begin VB.Frame fraTitle 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7245
      Begin VB.Label Label1 
         Caption         =   "ACCOUNTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   3660
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuChange 
         Caption         =   "C&hange"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnusep_b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSep_a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenAll 
         Caption         =   "&Expand All"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Colla&pse All"
      End
      Begin VB.Menu mnuSep_c 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Right Name:  "ChartOfAccounts"
Dim m_prevNode As Node
Dim m_tStart As Single
Dim m_rsAccounts As New ADODB.Recordset
Dim m_bCanExit As Boolean

Dim m_cCAccounts As New cAccounts

Private Sub InitEdit()
If Not HasRights(ChartOfAccounts, CanEdit) Then
  MsgBox "You do not have required privileges to proceed !", vbExclamation
  Exit Sub
End If
Dim m_sAccountID As String
Dim m_sAccountName As String
Dim m_sGroupID As String
Dim m_sRefAccountID As String
Dim m_node4Edit As Node
Dim m_bOK As Boolean
Dim m_bSuccess As Boolean
Dim m_bEditable As Boolean
Dim m_bGroup As Boolean
Dim m_lFrom As Long
Dim m_lTo As Long
Dim m_lStartRange As Long
Dim m_lEndRange As Long
Dim v
Dim m_bCanBeTurnedToDetail As Boolean
Set m_node4Edit = tvAccounts.SelectedItem
If m_node4Edit Is Nothing Then
   MsgBox "No Account selected to be edited"
   Exit Sub
End If
If m_node4Edit.Parent Is Nothing Then
   m_lStartRange = 0
   m_lEndRange = 0
Else
   v = Split(m_node4Edit.Parent.Tag, ";")
   m_lStartRange = IIf(Len(v(2)) = 0, 0, v(2))
   m_lEndRange = IIf(Len(v(3)) = 0, 0, v(3))
End If
If m_node4Edit.Children > 0 Then
   m_bCanBeTurnedToDetail = False
Else
   m_bCanBeTurnedToDetail = True
End If
m_sAccountID = GetAccountId(m_node4Edit.Text)
m_sAccountName = GetAccountName(m_node4Edit.Text)

m_sRefAccountID = GetAccountId(m_node4Edit.Text)
v = Split(m_node4Edit.Tag, ";", , vbTextCompare)
m_bEditable = CBool(v(0))
m_bGroup = v(1)
m_lFrom = IIf(Len(v(2)) = 0, 0, v(2))
m_lTo = IIf(Len(v(3)) = 0, 0, v(3))
m_bOK = frmAccount.ShowAccountForm("Edit Account [" & m_sAccountName & "]", True, m_sAccountID, m_sAccountName, m_bEditable, m_bGroup, m_lFrom, m_lTo, False, m_lStartRange, m_lEndRange, m_bCanBeTurnedToDetail)

If m_bOK Then
   If tvAccounts.SelectedItem.Parent Is Nothing Then
      m_sGroupID = "'0'"
   Else
      m_sGroupID = tvAccounts.SelectedItem.Parent.Key
   End If
   m_sGroupID = Left(m_sGroupID, Len(m_sGroupID) - 1)
   m_sGroupID = Right(m_sGroupID, Len(m_sGroupID) - 1)
   With m_cCAccounts
      .Edit Array("AccountID"), Array(m_sRefAccountID)
      .AccountID = m_sAccountID
      .AccountName = m_sAccountName
      .GroupID = m_sGroupID
      .Editable = m_bEditable
      .IsGroup = m_bGroup
      .FromAccountID = m_lFrom
      .ToAccountID = m_lTo
      .Save
   End With

      m_node4Edit.Text = m_sAccountID & " - " & m_sAccountName
      m_node4Edit.Tag = m_bEditable & ";" & m_bGroup & ";" & m_lFrom & ";" & m_lTo

Else
End If
End Sub

Private Sub InitFind()
frmFindAccount.Show vbModeless, Me
End Sub

Private Sub InitNew()
If Not HasRights(ChartOfAccounts, CanAdd) Then
  MsgBox "You do not have required privileges to proceed !", vbExclamation
  Exit Sub
End If
Dim m_sAccountID As String
Dim m_sAccountName As String
Dim m_sGroupID As String
Dim m_node4Child As Node
Dim m_bOK As Boolean
Dim m_bSuccess As Boolean
Dim m_node As Node
Dim m_sCaption As String
Dim m_bEditable As Boolean
Dim m_bGroup As Boolean
Dim m_lFrom As Long
Dim m_lTo As Long
Dim m_lStartRange As Long
Dim m_lEndRange As Long
Dim v

Set m_node4Child = tvAccounts.SelectedItem
v = Split(m_node4Child.Tag, ";")
m_lStartRange = IIf(Len(v(2)) = 0, 0, v(2))
m_lEndRange = IIf(Len(v(3)) = 0, 0, v(3))
m_sCaption = "New Account under [" & GetAccountName(m_node4Child.Text) & "]"
m_bOK = frmAccount.ShowAccountForm(m_sCaption, True, m_sAccountID, m_sAccountName, m_bEditable, m_bGroup, m_lFrom, m_lTo, True, m_lStartRange, m_lEndRange, , m_cCAccounts.NextAccountID(m_lStartRange, m_lEndRange))
If m_bOK Then
   m_sGroupID = m_node4Child.Key
   m_sGroupID = Left(m_sGroupID, Len(m_sGroupID) - 1)
   m_sGroupID = Right(m_sGroupID, Len(m_sGroupID) - 1)

   With m_cCAccounts
      .Edit Array("AccountID"), Array(m_sAccountID)
      If .CanSave Then
         MsgBox "An account with this Account ID already exist."
         .Cancel
         Exit Sub
      End If
      .AddNew
      .AccountID = m_sAccountID
      .AccountName = m_sAccountName
      .Editable = m_bEditable
      .FromAccountID = m_lFrom
      .GroupID = m_sGroupID
      .IsGroup = m_bGroup
      .ToAccountID = m_lTo
      .Save
   End With

      Set m_node = tvAccounts.Nodes.Add(m_node4Child, tvwChild, "'" & m_sAccountID & "'", m_sAccountID & " - " & m_sAccountName)
      m_node.Tag = m_bEditable & ";" & m_bGroup & ";" & m_lFrom & ";" & m_lTo
      m_node4Child.Expanded = True
      tvAccounts.SelectedItem = m_node4Child

Else
End If
End Sub

Private Sub InitRefresh()
m_bCanExit = False
mnuAction.Enabled = False
m_tStart = Timer

Set m_prevNode = Nothing
tmrNotify.Enabled = True
tvAccounts.Visible = False
PopListView
If tvAccounts.Nodes.Count > 0 Then
   tvAccounts.SelectedItem = tvAccounts.Nodes.Item(1).Root
   IndicateNode tvAccounts.SelectedItem
End If
tvAccounts.Visible = True
tmrNotify.Enabled = False
tmrNotify_Timer
m_bCanExit = True
mnuAction.Enabled = True
End Sub

Private Sub Form_Load()
m_bCanExit = True
DoEvents
InitRefresh
End Sub

Private Sub PopListView()

Dim m_node As Node
Dim m_sSQL As String
tvAccounts.Nodes.Clear
m_sSQL = "Select * From Accounts Where GroupID=0 Order By AccountID"
If m_rsAccounts.State = adStateOpen Then m_rsAccounts.Close
m_rsAccounts.Open m_sSQL, m_objConnectDB.cnnMyshop ' , adOpenForwardOnly, adLockReadOnly
If m_rsAccounts.BOF Or m_rsAccounts.EOF Then
   m_rsAccounts.Close
Else
   While Not m_rsAccounts.EOF
      'DoEvents
      Set m_node = tvAccounts.Nodes.Add(, , "'" & m_rsAccounts!AccountID & "'", m_rsAccounts!AccountID & " - " & m_rsAccounts!AccountName)
      m_node.Tag = m_rsAccounts!Editable & ";" & m_rsAccounts!IsGroup & ";" & m_rsAccounts!FromAccountID & ";" & m_rsAccounts!ToAccountID
      m_rsAccounts.Move 1
   Wend
   m_rsAccounts.Close
   Set m_node = m_node.FirstSibling
   While Not m_node Is Nothing
      DoEvents
      PopWithChild m_node
      Set m_node = m_node.Next
   Wend
End If
End Sub

Private Sub PopWithChild(ParentNode As Node)

Dim m_node As Node
Dim m_sSQL As String
Dim m_sTmp As String
m_sTmp = ParentNode.Key
m_sTmp = Left(m_sTmp, Len(m_sTmp) - 1)
m_sTmp = Right(m_sTmp, Len(m_sTmp) - 1)
m_sSQL = "Select * From Accounts Where GroupID=" & m_sTmp ' & " Order By AccountID"
m_rsAccounts.Open m_sSQL, m_objConnectDB.cnnMyshop ' , adOpenForwardOnly, adLockReadOnly
If m_rsAccounts.BOF Or m_rsAccounts.EOF Then
   m_rsAccounts.Close
Else
   While Not m_rsAccounts.EOF
      'DoEvents
      Set m_node = tvAccounts.Nodes.Add(ParentNode, tvwChild, "'" & m_rsAccounts!AccountID & "'", m_rsAccounts!AccountID & " - " & m_rsAccounts!AccountName)
      m_node.Tag = m_rsAccounts!Editable & ";" & m_rsAccounts!IsGroup & ";" & m_rsAccounts!FromAccountID & ";" & m_rsAccounts!ToAccountID

      m_rsAccounts.Move 1
   Wend
   m_rsAccounts.Close
   Set m_node = m_node.FirstSibling
   While Not m_node Is Nothing
      DoEvents
      PopWithChild m_node
      Set m_node = m_node.Next
   Wend
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = Not m_bCanExit
End Sub

Private Sub Form_Resize()
On Error Resume Next
Unload frmFindAccount

If Me.WindowState = vbMinimized Then
  Exit Sub
End If
fraTitle.Left = Me.ScaleLeft
fraTitle.Width = Me.ScaleWidth

tvAccounts.Left = Me.ScaleLeft
tvAccounts.Top = fraTitle.height + Me.ScaleTop
tvAccounts.Width = Me.ScaleWidth
tvAccounts.height = Me.ScaleHeight - fraTitle.height - staDetails.height
End Sub


Private Sub mnuAction_Click()
Dim m_node As Node
Set m_node = tvAccounts.SelectedItem
Dim v
If m_node Is Nothing Then
  mnuNew.Enabled = False
  mnuChange.Enabled = False
Else
  v = Split(m_node.Tag, ";")
  If CBool(v(1)) Then
    mnuNew.Enabled = True
  Else
    mnuNew.Enabled = False
  End If
  mnuChange.Enabled = True
End If
End Sub

Private Sub mnuChange_Click()
InitEdit
End Sub

Private Sub mnuCloseAll_Click()
Dim v As Node
tvAccounts.Visible = False
For Each v In tvAccounts.Nodes
   If v.Children > 0 Then v.Expanded = False
Next
tvAccounts.Visible = True
End Sub

Private Sub mnuExit_Click()
If StrComp(mnuExit.Caption, "E&xit", vbTextCompare) = 0 Then
   Unload Me
Else
End If
End Sub

Private Sub mnuFind_Click()
InitFind
End Sub

Private Sub mnuNew_Click()
InitNew
End Sub

Private Sub mnuOpenAll_Click()
Dim v As Node
tvAccounts.Visible = False
For Each v In tvAccounts.Nodes
   If v.Children > 0 Then v.Expanded = True
Next
tvAccounts.Visible = True
End Sub

Private Sub mnuRefresh_Click()
InitRefresh
End Sub

Private Sub tmrNotify_Timer()
On Error Resume Next
Dim i As Long
Dim j As Long
Dim k As Long
i = tvAccounts.Nodes.Count
j = Timer - m_tStart
If j = 0 Then
  k = 0
Else
  k = i / j
End If
With staDetails
   .Panels(1).Text = "Total Accounts: " & i & " "
   .Panels(2).Text = "Time Elapsed: " & j & " s "
   .Panels(3).Text = "Data Rate: " & k & " "
End With
End Sub

Private Sub tvAccounts_DblClick()
If tvAccounts.SelectedItem Is Nothing Then
Else
   If tvAccounts.SelectedItem.Children < 1 Then mnuChange_Click
End If
End Sub

Private Sub tvAccounts_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If tvAccounts.SelectedItem Is Nothing Then
   Else
      If tvAccounts.SelectedItem.Children < 1 Then mnuChange_Click
   End If
End If
End Sub

Private Sub tvAccounts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      mnuExit.Caption = "&Cancel"
      PopupMenu mnuAction
      mnuExit.Caption = "E&xit"
   End If
End Sub

Private Function GetAccountId(str As String) As String
GetAccountId = Left(str, InStr(1, str, " - ") - 1)
End Function

Private Function GetAccountName(str As String) As String
GetAccountName = Right(str, Len(str) - (InStr(1, str, " - ") - 1 + Len(" - ")))
End Function

Private Sub tvAccounts_NodeClick(ByVal Node As MSComctlLib.Node)
' IndicateNode Node
End Sub

Public Sub IndicateNode(ByVal Node As MSComctlLib.Node)
If Not m_prevNode Is Nothing Then
   m_prevNode.Bold = False
End If
Set m_prevNode = Node
If Not Node.Parent Is Nothing Then
   Node.Parent.Bold = True
   Set m_prevNode = Node.Parent
End If
End Sub
