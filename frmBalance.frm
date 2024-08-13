VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmBalance 
   Caption         =   "Balance Sheet Layout"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form4"
   ScaleHeight     =   3900
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picUp 
      Align           =   1  'Align Top
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   5760
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5760
      Begin VB.PictureBox picA 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   390
         Left            =   3090
         ScaleHeight     =   390
         ScaleWidth      =   2490
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   375
         Width           =   2490
         Begin VB.CommandButton cmdAssociateR 
            Height          =   330
            Left            =   1650
            Picture         =   "frmBalance.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Edit Asset Association"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdEditA 
            Height          =   330
            Left            =   330
            Picture         =   "frmBalance.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Edit Asset"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdAddA 
            Height          =   330
            Left            =   0
            Picture         =   "frmBalance.frx":0444
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "New Asset"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdDelA 
            Height          =   330
            Left            =   660
            Picture         =   "frmBalance.frx":0546
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Delete Asset"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdDownA 
            Height          =   330
            Left            =   990
            Picture         =   "frmBalance.frx":0888
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Move Down Asset"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdUpA 
            Height          =   330
            Left            =   1320
            Picture         =   "frmBalance.frx":0BCA
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Move Up Asset"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   330
         End
      End
      Begin VB.PictureBox picL 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   390
         Left            =   0
         ScaleHeight     =   390
         ScaleWidth      =   2475
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   375
         Width           =   2475
         Begin VB.CommandButton cmdAssociateL 
            Height          =   330
            Left            =   1650
            Picture         =   "frmBalance.frx":0F0C
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Edit Liability Association"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdUpL 
            Height          =   330
            Left            =   1320
            Picture         =   "frmBalance.frx":124E
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Move Up Liability"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdDownL 
            Height          =   330
            Left            =   990
            Picture         =   "frmBalance.frx":1590
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Move Down Liability"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdDelL 
            Height          =   330
            Left            =   660
            Picture         =   "frmBalance.frx":18D2
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Delete Liability"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdAddL 
            Height          =   330
            Left            =   0
            Picture         =   "frmBalance.frx":19D4
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "New Liability"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdEditL 
            Height          =   330
            Left            =   330
            Picture         =   "frmBalance.frx":1AD6
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Edit Liability"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   330
         End
      End
      Begin VB.Label lblheader 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Balance Sheet Layout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   195
         TabIndex        =   1
         Top             =   15
         Width           =   4725
      End
   End
   Begin MSComctlLib.ListView lvwLiability 
      Height          =   2895
      Left            =   75
      TabIndex        =   16
      Top             =   870
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Liabilities"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwAsset 
      Height          =   2970
      Left            =   3180
      TabIndex        =   17
      Top             =   840
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   5239
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Assets"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim cBSLi As New CBalSheet
Public cn As ADODB.Connection

Dim cBSAs As New CBalSheet

Private Sub cmdAddA_Click()
cBSAs.AddItem "<Asset>"
End Sub

Private Sub cmdAddL_Click()
cBSLi.AddItem "<Liability>"
End Sub

Private Sub cmdAssociateL_Click()
Dim f As New frmBalAssociate
Dim li As ListItem
Dim k As Long
With lvwLiability
    Set li = .SelectedItem
    If li Is Nothing Then
        MsgBox "No Liability is selected."
        Exit Sub
    End If
    k = CLng(Mid(li.Key, 1 + InStr(1, li.Key, "=")))
    Set f.cn = cn
    f.Move Me.Left, Me.Top, Me.Width, Me.Height
    f.ShowForm li.Text, k, vbModal, Me
    Set f = Nothing
    .SetFocus
End With
End Sub

Private Sub cmdAssociateR_Click()
Dim f As New frmBalAssociate
Dim li As ListItem
Dim k As Long
With lvwAsset
    Set li = .SelectedItem
    If li Is Nothing Then
        MsgBox "No Asset is selected."
        Exit Sub
    End If
    k = CLng(Mid(li.Key, 1 + InStr(1, li.Key, "=")))
    Set f.cn = cn
    f.Move Me.Left, Me.Top, Me.Width, Me.Height
    f.ShowForm li.Text, k, vbModal, Me
    Set f = Nothing
    .SetFocus
End With
End Sub

Private Sub cmdDelA_Click()
If MsgBox("Are You Sure to Delete Asset ?", vbQuestion + vbYesNo) = vbYes Then
    If Not cBSAs.Delete() Then
        MsgBox "Nothing Deleted", vbInformation
    End If
End If
lvwAsset.SetFocus
End Sub

Private Sub cmdDelL_Click()
If MsgBox("Are You Sure to Delete Liability ?", vbQuestion + vbYesNo) = vbYes Then
    If Not cBSLi.Delete() Then
        MsgBox "Nothing Deleted", vbInformation
    End If
End If
lvwLiability.SetFocus
End Sub

Private Sub cmdDownA_Click()
cBSAs.MoveDown
End Sub

Private Sub cmdDownL_Click()
cBSLi.MoveDown
End Sub

Private Sub cmdEditA_Click()
cBSAs.Edit
End Sub

Private Sub cmdEditL_Click()
cBSLi.Edit
End Sub

Private Sub cmdUpA_Click()
cBSAs.MoveUp
End Sub

Private Sub cmdUpL_Click()
cBSLi.MoveUp
End Sub

Private Sub Form_Load()
Set cn = m_objConnectDB.cnnMyshop
With cBSLi
    Set .ActiveConnection = cn
    .BalType = "LI"
    Set .ListView = lvwLiability
    .PopulateLVW
End With
With cBSAs
    Set .ActiveConnection = cn
    .BalType = "AS"
    Set .ListView = lvwAsset
    .PopulateLVW
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMain.Enabled = True
frmMain.SetFocus
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
If Me.Width < 4500 Then Me.Width = 4500
If Me.Height < 2200 Then Me.Height = 2200
lvwLiability.Move Me.ScaleLeft, picUp.Height, Me.ScaleWidth / 2 - 25, Me.ScaleHeight - picUp.Height - 100
lvwLiability.ColumnHeaders.Item(1).Width = lvwLiability.Width - 100
lvwAsset.Move 25 + Me.ScaleWidth / 2, picUp.Height, Me.ScaleWidth / 2, Me.ScaleHeight - picUp.Height - 100
lvwAsset.ColumnHeaders.Item(1).Width = lvwAsset.Width - 100
End Sub

Private Sub picUp_Resize()
Const pos As Integer = 3
Select Case pos
Case 1 ' Right
    picL.Left = lvwLiability.Left + lvwLiability.Width - picL.Width - 100
    picA.Left = lvwAsset.Left + lvwAsset.Width - picA.Width - 100
Case 2 ' Center
    picL.Left = picUp.ScaleWidth / 4 - picL.Width / 2
    picA.Left = picUp.ScaleWidth * 3 / 4 - picA.Width / 2
Case 3 ' Left
    picL.Left = picUp.ScaleLeft + 35
    picA.Left = picUp.ScaleWidth / 2
End Select

lblheader.Move picUp.ScaleLeft, picUp.ScaleTop, picUp.ScaleWidth, lblheader.Height
End Sub
