VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   HasDC           =   0   'False
   Icon            =   "frmAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraRanges 
      Caption         =   "Account ID &range for sub accounts"
      Height          =   1335
      Left            =   165
      TabIndex        =   7
      Top             =   1365
      Width           =   4380
      Begin MSComCtl2.UpDown udTo 
         Height          =   315
         Left            =   2761
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   825
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtToAccountID"
         BuddyDispid     =   196611
         OrigLeft        =   3390
         OrigTop         =   870
         OrigRight       =   3630
         OrigBottom      =   1140
         Max             =   100
         Min             =   1
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udFrom 
         Height          =   315
         Left            =   2775
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtFromAccountID"
         BuddyDispid     =   196610
         OrigLeft        =   3030
         OrigTop         =   375
         OrigRight       =   3270
         OrigBottom      =   660
         Max             =   100
         Min             =   1
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtFromAccountID 
         Height          =   315
         Left            =   1230
         MaxLength       =   9
         TabIndex        =   9
         Text            =   "999999999"
         Top             =   345
         Width           =   1530
      End
      Begin VB.TextBox txtToAccountID 
         Height          =   315
         Left            =   1230
         MaxLength       =   9
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   825
         Width           =   1530
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "T&o"
         Height          =   195
         Left            =   825
         TabIndex        =   11
         Top             =   885
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fro&m"
         Height          =   195
         Left            =   660
         TabIndex        =   8
         Top             =   405
         Width           =   345
      End
   End
   Begin VB.OptionButton optDetail 
      Caption         =   "Detail"
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   900
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton optGroup 
      Caption         =   "Group"
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   900
      Width           =   810
   End
   Begin VB.PictureBox picCommands 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   4740
      TabIndex        =   17
      Top             =   3225
      Width           =   4740
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   360
         Left            =   2370
         TabIndex        =   15
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   360
         Left            =   1305
         TabIndex        =   16
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.CheckBox chkEditable 
      Caption         =   "Account can not be &changed in future"
      Height          =   345
      Left            =   870
      TabIndex        =   14
      Top             =   2760
      Width           =   3075
   End
   Begin VB.TextBox txtAccount 
      Height          =   315
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   3
      ToolTipText     =   "Account Name"
      Top             =   480
      Width           =   3285
   End
   Begin VB.TextBox txtAccountID 
      Height          =   315
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   1
      ToolTipText     =   "Account ID"
      Top             =   105
      Width           =   3285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Account &Type"
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Account &Name"
      Height          =   195
      Left            =   165
      TabIndex        =   2
      Top             =   540
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Account &ID"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   165
      Width           =   810
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sAccountID As String
Private m_sAccountName As String
Private m_bEditable As Boolean
Private m_bOK As String
Private m_bGroup As Boolean
Private m_lFrom As Long
Private m_lTo As Long

Private m_lStartRange As Long, m_lEndRange As Long
Dim m_bDirUp As Boolean

Private Sub cmdCancel_Click()
m_bOK = False
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim m_lAID As Long
m_lAID = CLng(txtAccountID.Text)
If (m_lAID < m_lStartRange Or m_lAID > m_lEndRange) And m_lStartRange <> 0 And m_lEndRange <> 0 Then
   MsgBox "Account ID must be within (" & m_lStartRange & ", " & m_lEndRange & ")", vbOKOnly + vbCritical
   Exit Sub
End If
If txtAccount = "" Then
  MsgBox "Account Name is missing.", vbOKOnly + vbCritical
  Exit Sub
End If
If CLng(txtFromAccountID.Text) <> 0 Or CLng(txtToAccountID.Text) <> 0 Then
   If CLng(txtFromAccountID.Text) < CLng(txtToAccountID.Text) Then
      MsgBox "Can not save as" & vbCrLf & _
         "Range must be specified from a lower value to higher value.", vbOKOnly + vbCritical
   ElseIf AccountIDRangePermissible Then
      m_bOK = True
      Unload Me
   Else
      MsgBox "Can not save as " & vbCrLf & _
         "Account ID range for sub accounts is invalid", vbOKOnly + vbCritical
   End If
Else
m_bOK = True
Unload Me
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Load()
optDetail.value = True
optDetail_Click
With udFrom
  .Left = txtFromAccountID.Left + txtFromAccountID.Width - .Width
End With
With udTo
  .Left = txtToAccountID.Left + txtToAccountID.Width - .Width
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
If m_bOK Then
   m_sAccountID = txtAccountID.Text
   m_sAccountName = txtAccount.Text
   m_bEditable = IIf(chkEditable.value = vbChecked, False, True)
   If optGroup.value Then
      m_bGroup = True
      m_lFrom = txtFromAccountID.Text
      m_lTo = txtToAccountID.Text
   Else
      m_bGroup = False
      m_lFrom = 0
      m_lTo = 0
   End If
   
Else
End If
End Sub

Public Function ShowAccountForm(title As String, ByRef Save As Boolean, ByRef AccountID As String, ByRef AccountName As String, ByRef Editable As Boolean, ByRef IsGroup As Boolean, ByRef FromRange As Long, ByRef ToRange As Long, Optional NewAccount As Boolean = False, Optional StartRange As Long = 0, Optional EndRange As Long = 0, Optional CanBeTurnedToDetail As Boolean = True, Optional lNAID As Long) As Boolean
Load Me
txtAccountID.Text = AccountID
txtAccount.Text = AccountName
chkEditable.value = IIf(Editable, vbUnchecked, vbChecked)
If Save Then
   If NewAccount Then ' New
      txtAccountID.Locked = IIf(lNAID = 0, False, True)      ' True ' False
      chkEditable.Enabled = True
      chkEditable.value = vbUnchecked
      txtAccount.Locked = False
      optDetail.value = True
      txtFromAccountID.Text = 0
      txtToAccountID.Text = 0
      txtAccountID.Text = lNAID
   Else  ' Edit
      txtAccountID.Locked = True
      chkEditable.Enabled = Editable
      txtAccount.Locked = Not Editable
      If IsGroup Then
         optGroup.value = True
         txtFromAccountID.Text = FromRange
         txtToAccountID.Text = ToRange
      Else
         optDetail.value = True
         txtFromAccountID.Text = 0
         txtToAccountID.Text = 0
      End If
      optDetail.Enabled = CanBeTurnedToDetail
      If Not Editable Then
         txtAccount.Locked = True
         txtFromAccountID.Locked = True
         txtToAccountID.Locked = True
         optGroup.Enabled = False
         optDetail.Enabled = False
      End If
   End If
Else ' Delete
   cmdSave.Caption = "&Delete"
   txtAccountID.Locked = True
   txtAccount.Locked = True
   chkEditable.Enabled = False
   If IsGroup Then
      optGroup.value = True
      txtFromAccountID.Text = FromRange
      txtToAccountID.Text = ToRange
   Else
      optDetail.value = True
      txtFromAccountID.Text = 0
      txtToAccountID.Text = 0
   End If
End If
If StartRange = 0 And EndRange = 0 Then
   txtAccountID.ToolTipText = "Account ID can be any available number."
   m_lStartRange = 0
   m_lEndRange = 0
Else
   txtAccountID.ToolTipText = "Valid Range for Account ID is from " & StartRange & " to " & EndRange & "."
   m_lStartRange = StartRange
   m_lEndRange = EndRange
End If
m_bOK = False
Me.Caption = title
Me.Show vbModal
AccountID = m_sAccountID
AccountName = m_sAccountName
Editable = m_bEditable
IsGroup = m_bGroup
FromRange = m_lFrom
ToRange = m_lTo
ShowAccountForm = m_bOK
End Function

Private Sub optDetail_Click()
txtFromAccountID.Enabled = False
txtFromAccountID.Text = "0"
txtToAccountID.Enabled = False
txtToAccountID.Text = "0"
udFrom.Enabled = False
udTo.Enabled = False
End Sub

Private Sub optGroup_Click()
txtFromAccountID.Enabled = True
txtToAccountID.Enabled = True
udFrom.Enabled = True
udTo.Enabled = True
End Sub

Private Sub txtAccount_GotFocus()
HighlightText txtAccount
End Sub

Private Sub txtAccountID_GotFocus()
HighlightText txtAccountID
End Sub

Private Sub txtAccountID_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtFromAccountID_Change()
If Len(txtFromAccountID.Text) = 0 Then
   txtFromAccountID.Text = "0"
End If
End Sub

Private Sub txtFromAccountID_GotFocus()
HighlightText txtFromAccountID
End Sub

Private Sub txtFromAccountID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
   udFrom_DownClick
ElseIf KeyCode = vbKeyUp Then
   udFrom_UpClick
End If
End Sub

Private Sub txtFromAccountID_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtToAccountID_Change()
If Len(txtToAccountID.Text) = 0 Then
   txtToAccountID.Text = "0"
End If
End Sub

Private Sub txtToAccountID_GotFocus()
HighlightText txtToAccountID
End Sub

Private Sub txtToAccountID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
   udTo_DownClick
ElseIf KeyCode = vbKeyUp Then
   udTo_UpClick
End If
End Sub

Private Sub txtToAccountID_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub udFrom_DownClick()
If txtFromAccountID.Locked Then Exit Sub
On Error GoTo EH
Dim v As Long
txtFromAccountID.SetFocus
v = CLng(txtFromAccountID.Text)
v = v - 1
If v < 0 Then v = 0
txtFromAccountID.Text = v
Exit Sub
EH:
txtFromAccountID.Text = "0"

End Sub

Private Sub udFrom_UpClick()
If txtFromAccountID.Locked Then Exit Sub
On Error GoTo EH
Dim v As Long
txtFromAccountID.SetFocus
v = CLng(txtFromAccountID.Text)
v = v + 1
If v < 0 Then v = 0
txtFromAccountID.Text = v
Exit Sub
EH:
txtFromAccountID.Text = "0"
End Sub

Private Sub udTo_DownClick()
If txtToAccountID.Locked Then Exit Sub
On Error GoTo EH
Dim v As Long
txtToAccountID.SetFocus
v = CLng(txtToAccountID.Text)
v = v - 1
If v < 0 Then v = 0
txtToAccountID.Text = v
Exit Sub
EH:
txtToAccountID.Text = "0"
End Sub

Private Sub udTo_UpClick()
If txtToAccountID.Locked Then Exit Sub
On Error GoTo EH
Dim v As Long
txtToAccountID.SetFocus
v = CLng(txtToAccountID.Text)
v = v + 1
If v < 0 Then v = 0
txtToAccountID.Text = v
Exit Sub
EH:
txtToAccountID.Text = "0"
End Sub

Function AccountIDRangePermissible() As Boolean
Dim sSQL As String
Dim rsAccounts As New ADODB.Recordset
'Dim cCDB As New ConnectDB
Dim i As Long
Dim j As Long
i = CLng(txtFromAccountID.Text)
j = CLng(txtToAccountID.Text)

sSQL = "Select * From Accounts Where AccountID=" & i & " or AccountID=" & j & " or  FromAccountID=" & i & " or FromAccountID=" & i & " or FromAccountID=" & j & " or ToAccountID=" & j & _
   " or (" & i & ">FromAccountID and " & i & "<ToAccountID )" & _
   " or (" & j & ">FromAccountID and " & j & "<ToAccountID )" & _
   " or (FromAccountID>" & i & " and FromAccountID<" & j & ")" & _
   " or (ToAccountID>" & i & " and ToAccountID<" & j & ")"
Debug.Print sSQL
With rsAccounts
   .Open sSQL, m_objConnectDB.cnnMyshop
   On Error Resume Next
   .MoveNext
   .MovePrevious
   If .EOF Then
      AccountIDRangePermissible = True
   Else
      AccountIDRangePermissible = False
   End If
End With
End Function

