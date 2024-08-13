VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAccounts2 
   Caption         =   "Accounts"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12465
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   Picture         =   "frmAccounts2.frx":0000
   ScaleHeight     =   8865
   ScaleWidth      =   12465
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvwAccounts 
      Height          =   6600
      Left            =   105
      TabIndex        =   0
      Top             =   1245
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   11642
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove                  "
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnu_sep_a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnu_sep_b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "frmAccounts2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_MouseX As Single
Dim m_MouseY As Single
Dim m_Tlb As Integer

Private Sub Form_Load()
    Dim cCAccts2 As New cAccounts2
    cCAccts2.PopListView lvwAccounts
    frmMain.SetTlbLayout 3
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmFindAccountLVW
End Sub

'Private Sub Form_Resize()
'    Dim h As Single, sngMinW As Single, sngMinH As Single
'    If Me.WindowState = vbMinimized Then
'      Unload frmFindAccountLVW
'      Exit Sub
'    End If
'    sngMinW = 165
'    sngMinH = 6
'    If Me.Width < sngMinW Then
'      Me.Width = sngMinW
'    End If
'    If Me.Height < sngMinH Then
'      Me.Height = sngMinH
'    End If
'    With lvwAccounts
'      .Left = Me.ScaleLeft
'      .Top = Me.ScaleTop
'      .Width = Me.ScaleWidth
'      h = Me.ScaleHeight
'      .Height = IIf(h < 0, 0, h)
'    End With
'    SetHeaderWidth
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not frmMain.m_lOldTBID = -5 Then
       frmMain.SetTlbLayout frmMain.m_lOldTBID
    Else
       frmMain.SetTlbDefLayout
    End If
End Sub

Private Sub lvwAccounts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwAccounts.Sorted = True
    lvwAccounts.SortKey = ColumnHeader.index - 1
End Sub

Private Sub lvwAccounts_DblClick()
Dim li As ListItem
Set li = lvwAccounts.HitTest(m_MouseX, m_MouseY)
If li Is Nothing Then
    NewAccount
Else
    EditAccount
End If
End Sub

'Private Sub lvwAccounts_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyDelete Then
'  RemoveAccount
'ElseIf KeyCode = vbKeyN And Shift = vbCtrlMask Then
'  NewAccount
'End If
'End Sub

Private Sub lvwAccounts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      EditAccount
    End If
End Sub

Private Sub lvwAccounts_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 93 Then
  PopupMenu mnuAction, , , , mnuEdit
End If
End Sub

Private Sub lvwAccounts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim li As ListItem
    m_MouseX = X
    m_MouseY = Y
    Set li = lvwAccounts.HitTest(X, Y)
    If Button = vbRightButton And Shift = 0 Then
      If li Is Nothing Then
        PopupMenu mnuAction, , , , mnuNew
      Else
        PopupMenu mnuAction, , , , mnuEdit
      End If
    End If
End Sub

Private Sub mnuEdit_Click()
EditAccount
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFind_Click()
FindAccount
End Sub

Private Sub mnuNew_Click()
NewAccount
End Sub

Private Sub mnuRemove_Click()
RemoveAccount
End Sub

Sub NewAccount()
    Unload frmFindAccountLVW
    If Not HasRights(soChartOfAccounts, CanAdd) Then
        MsgBox "You do not have required privileges to proceed !", vbExclamation
        Exit Sub
    End If
    Dim cCAccts2 As New cAccounts2
    If frmAccount2.ShowForm(cCAccts2, "n", vbModal, frmMain) Then
      With cCAccts2
        .Save True
        .Add2ListView lvwAccounts, frmAccount2.GroupDescription
      End With
    End If
    lvwAccounts.SetFocus
End Sub

Sub EditAccount()
Unload frmFindAccountLVW
If Not HasRights(soChartOfAccounts, CanEdit) Then
    MsgBox "You do not have required privileges to proceed !", vbExclamation
    Exit Sub
End If
Dim cCAccts2 As New cAccounts2
Dim li As ListItem
Set li = lvwAccounts.SelectedItem
If li Is Nothing Then
  Exit Sub
End If
cCAccts2.Init CLng(li.Tag)
If frmAccount2.ShowForm(cCAccts2, "e", vbModal, frmMain) Then
  With cCAccts2
    .Save False
    .UpdateInListView lvwAccounts, li, frmAccount2.GroupDescription, frmAccount2.SubGroupDescription
  End With
End If
End Sub

Sub RemoveAccount()
Unload frmFindAccountLVW
If Not HasRights(soChartOfAccounts, CanDelete) Then
    MsgBox "You do not have required privileges to proceed !", vbExclamation
    Exit Sub
End If
Dim cCAccts2 As New cAccounts2
Dim li As ListItem
Set li = lvwAccounts.SelectedItem
If li Is Nothing Then
  Exit Sub
End If
cCAccts2.Init CLng(li.Tag)
If frmAccount2.ShowForm(cCAccts2, "d", vbModal, frmMain) Then
  With cCAccts2
    .Remove
    .RemoveFromListView lvwAccounts, li
  End With
End If
End Sub

Sub FindAccount()
Set frmFindAccountLVW.lvw = lvwAccounts
frmFindAccountLVW.Show vbModeless, frmMain
End Sub

Sub SetHeaderWidth()
Dim ch As ColumnHeader
With lvwAccounts
    .Visible = False
'
    Set ch = .ColumnHeaders.Item(4)
    ch.Width = 750
    Set ch = .ColumnHeaders.Item(5)
    ch.Width = 750
    Set ch = .ColumnHeaders.Item(6)
    ch.Width = 1100
'
    Set ch = .ColumnHeaders.Item(1)
    ch.Width = .Width * 0.333
    Set ch = .ColumnHeaders.Item(2)
    ch.Width = .Width * 0.2
    Set ch = .ColumnHeaders.Item(3)
    ch.Width = .Width * 0.2
'
    .Visible = True
End With
End Sub
