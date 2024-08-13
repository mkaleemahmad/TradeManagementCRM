VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAccounts21 
   Caption         =   "Accounts"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6435
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   6435
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3375
      Width           =   6435
      Begin VB.PictureBox picAction 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   30
         ScaleHeight     =   420
         ScaleWidth      =   6360
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         Width           =   6360
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   360
            Left            =   30
            TabIndex        =   7
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   360
            Left            =   5130
            TabIndex        =   1
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "&Remove"
            Height          =   360
            Left            =   3855
            TabIndex        =   2
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "&New"
            Height          =   360
            Left            =   1305
            TabIndex        =   4
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   360
            Left            =   2580
            TabIndex        =   3
            Top             =   0
            Width           =   1215
         End
      End
   End
   Begin MSComctlLib.ListView lvwAccounts 
      Height          =   3255
      Left            =   15
      TabIndex        =   0
      Top             =   120
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   5741
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
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnu_sep_a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnu_sep_b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmAccounts21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdEdit_Click()
EditAccount
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
FindAccount
End Sub

Private Sub cmdNew_Click()
NewAccount
End Sub

Private Sub cmdRemove_Click()
RemoveAccount
End Sub


Private Sub Form_Load()
Dim cCAccts2 As New cAccounts2
cCAccts2.PopListView lvwAccounts
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload frmFindAccountLVW
End Sub

Private Sub Form_Resize()
Dim h As Single, sngMinW As Single, sngMinH As Single
If Me.WindowState = vbMinimized Then
  Unload frmFindAccountLVW
  Exit Sub
End If
sngMinW = picAction.Width + 165
sngMinH = 6 * picAction.Height
If Me.Width < sngMinW Then
  Me.Width = sngMinW
End If
If Me.Height < sngMinH Then
  Me.Height = sngMinH
End If
With lvwAccounts
  .Left = Me.ScaleLeft
  .Top = Me.ScaleTop
  .Width = Me.ScaleWidth
  h = Me.ScaleHeight - picBottom.Height
  .Height = IIf(h < 0, 0, h)
End With
End Sub

Private Sub lvwAccounts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvwAccounts.Sorted = True
lvwAccounts.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub lvwAccounts_DblClick()
EditAccount
End Sub

Private Sub lvwAccounts_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
  RemoveAccount
ElseIf KeyCode = vbKeyN And Shift = vbCtrlMask Then
  NewAccount
End If
End Sub

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

Private Sub lvwAccounts_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim li As ListItem
Set li = lvwAccounts.HitTest(x, y)
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

Private Sub picBottom_Resize()
picAction.Left = picBottom.ScaleWidth - picAction.Width
End Sub

Sub NewAccount()
Unload frmFindAccountLVW
Dim cCAccts2 As New cAccounts2
If frmAccount2.ShowForm(cCAccts2, "n", vbModal, frmMain) Then
  With cCAccts2
    .Save True
    .Add2ListView lvwAccounts, frmAccount2.GroupDescription, frmAccount2.SubGroupDescription
  End With
End If
lvwAccounts.SetFocus
End Sub

Sub EditAccount()
Unload frmFindAccountLVW
Dim cCAccts2 As New cAccounts2
Dim li As ListItem
Set li = lvwAccounts.SelectedItem
If li Is Nothing Then
  Exit Sub
End If
cCAccts2.Init CLng(li.Text)
If frmAccount2.ShowForm(cCAccts2, "e", vbModal, frmMain) Then
  With cCAccts2
    .Save False
    .UpdateInListView lvwAccounts, frmAccount2.GroupDescription, frmAccount2.SubGroupDescription
  End With
End If
End Sub

Sub RemoveAccount()
Unload frmFindAccountLVW
Dim cCAccts2 As New cAccounts2
Dim li As ListItem
Set li = lvwAccounts.SelectedItem
If li Is Nothing Then
  Exit Sub
End If
cCAccts2.Init CLng(li.Text)
If frmAccount2.ShowForm(cCAccts2, "d", vbModal, frmMain) Then
  With cCAccts2
    .Remove
    .RemoveFromListView lvwAccounts
  End With
End If
End Sub

Sub FindAccount()
Set frmFindAccountLVW.lvw = lvwAccounts
frmFindAccountLVW.Show vbModeless, frmMain
End Sub
