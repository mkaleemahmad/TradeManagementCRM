VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBalAssociate 
   Caption         =   "Balance Sheet Association"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMove 
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   2430
      ScaleHeight     =   1395
      ScaleWidth      =   630
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   855
      Width           =   630
      Begin VB.CommandButton cmdLeftAll 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   9
         Top             =   1005
         Width           =   576
      End
      Begin VB.CommandButton cmdLeftOne 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   8
         Top             =   675
         Width           =   576
      End
      Begin VB.CommandButton cmdRightAll 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   330
         Width           =   576
      End
      Begin VB.CommandButton cmdRightOne 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   6
         Top             =   0
         Width           =   576
      End
   End
   Begin VB.PictureBox picTitle 
      Align           =   1  'Align Top
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   6630
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6630
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Asset OR Liability"
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
         Height          =   300
         Left            =   3195
         TabIndex        =   1
         Top             =   15
         Width           =   3435
      End
   End
   Begin MSComctlLib.ListView lvwAGS 
      Height          =   3450
      Left            =   3090
      TabIndex        =   4
      Top             =   465
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   6085
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "List No."
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Associate Type"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Associate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ListBox lstAGS 
      Height          =   3090
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   3
      Top             =   870
      Width           =   2310
   End
   Begin VB.ComboBox cmbAGS 
      Height          =   315
      ItemData        =   "frmBalAssociate.frx":0000
      Left            =   60
      List            =   "frmBalAssociate.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2310
   End
End
Attribute VB_Name = "frmBalAssociate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public cn As New ADODB.Connection
Private lRunCount As Long
Public lBalSheetID As Long

Sub PopWithAccount()
Dim sSQL As String
Dim rs As New ADODB.Recordset
sSQL = "SELECT * FROM Accounts Where ID NOT IN(SELECT AssociateID From BalAssociate Where AssociateType LIKE 'Account' And BalSheetID=" & lBalSheetID & ") Order by Accounts.AccountName "
rs.Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
lstAGS.Clear
Do Until rs.EOF
    lstAGS.AddItem rs("AccountName").Value
    lstAGS.ItemData(lstAGS.NewIndex) = rs("ID").Value
    rs.MoveNext
Loop
End Sub

Sub PopWithGroup()
Dim sSQL As String
Dim rs As New ADODB.Recordset
sSQL = "SELECT * FROM AccountGroups Where GroupID NOT IN(SELECT AssociateID From BalAssociate Where AssociateType LIKE 'Group' And BalSheetID=" & lBalSheetID & ") Order by AccountGroups.Description"
rs.Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
lstAGS.Clear
Do Until rs.EOF
    lstAGS.AddItem rs("Description").Value
    lstAGS.ItemData(lstAGS.NewIndex) = rs("GroupID").Value
    rs.MoveNext
Loop
End Sub

Sub PopWithSubGroup()
Dim sSQL As String
Dim rs As New ADODB.Recordset
sSQL = "SELECT * FROM AccountSubGroups Where SubGroupID NOT IN(SELECT AssociateID From BalAssociate Where AssociateType LIKE 'SubGroup' And BalSheetID=" & lBalSheetID & ") Order By AccountSubGroups.Description"
rs.Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
lstAGS.Clear
Do Until rs.EOF
    lstAGS.AddItem rs("Description").Value
    lstAGS.ItemData(lstAGS.NewIndex) = rs("SubGroupID").Value
    rs.MoveNext
Loop
End Sub

Private Sub cmbAGS_click()
If cmbAGS.Text = "Accounts" Then
    PopWithAccount
ElseIf cmbAGS.Text = "Groups" Then
    PopWithGroup
ElseIf cmbAGS.Text = "Sub Groups" Then
    PopWithSubGroup
Else
    lstAGS.Clear
End If
End Sub

Private Sub cmdLeftAll_Click()
picMove.Enabled = False
MoveAllLeft
picMove.Enabled = True
End Sub

Private Sub cmdLeftOne_Click()
MoveOneLeft
End Sub

Private Sub cmdRightAll_Click()
picMove.Enabled = False
MoveAllRight
picMove.Enabled = True
End Sub

Private Sub cmdRightOne_Click()
MoveOneRight
End Sub

Private Sub Form_Resize()
SetFormLayout
End Sub

Sub SetFormLayout()
If Me.WindowState = vbMinimized Then Exit Sub
If Me.Width < 5580 Then Me.Width = 5580
If Me.Height < 2795 Then Me.Height = 2795
Dim k As Single
lstAGS.Height = Me.ScaleHeight - lstAGS.Top - 100
lvwAGS.Height = lstAGS.Top + lstAGS.Height - lvwAGS.Top  ' Me.ScaleHeight - lvwAGS.Top
lvwAGS.Width = Me.ScaleWidth - lvwAGS.Left - 65
k = lvwAGS.Width - lvwAGS.ColumnHeaders.Item(1).Width - lvwAGS.ColumnHeaders.Item(2).Width - 100
If k < 1400 Then k = 1400
lvwAGS.ColumnHeaders.Item(3).Width = k
End Sub

Sub ShowAssociates()
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim li As ListItem
' Show Associated Accounts
lRunCount = 1
sSQL = "SELECT BalAssociate.*,Accounts.AccountName FROM BalAssociate Left Join Accounts on AssociateID=Accounts.ID Where AssociateType='Account' And BalSheetID=" & lBalSheetID
rs.Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
lvwAGS.ListItems.Clear
Do Until rs.EOF
    Set li = lvwAGS.ListItems.Add(, "ID=" & rs("ID").Value, lRunCount)
    li.ListSubItems.Add , "Q", "Account"
    li.ListSubItems.Add , "A", rs("AccountName").Value
    li.Tag = rs("BalSheetID") & ";" & rs("AssociateID") & ";" & rs("AssociateType")
    rs.MoveNext
    lRunCount = lRunCount + 1
Loop
' Show Associated Groups
rs.Close
sSQL = "SELECT BalAssociate.*,AccountGroups.Description FROM BalAssociate Left Join AccountGroups on AssociateID=AccountGroups.GroupID Where AssociateType='Group' And BalSheetID=" & lBalSheetID
rs.Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
Do Until rs.EOF
    Set li = lvwAGS.ListItems.Add(, "ID=" & rs("ID").Value, lRunCount)
    li.ListSubItems.Add , "Q", "Group"
    li.ListSubItems.Add , "A", rs("Description").Value
    li.Tag = rs("BalSheetID") & ";" & rs("AssociateID") & ";" & rs("AssociateType")
    rs.MoveNext
    lRunCount = lRunCount + 1
Loop
' Show Associated Sub Groups
rs.Close
sSQL = "SELECT BalAssociate.*,AccountSubGroups.Description FROM BalAssociate Left Join AccountSubGroups on AssociateID=AccountSubGroups.SubGroupID Where AssociateType='SubGroup' And BalSheetID=" & lBalSheetID
rs.Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
Do Until rs.EOF
    Set li = lvwAGS.ListItems.Add(, "ID=" & rs("ID").Value, lRunCount)
    li.ListSubItems.Add , "Q", "SubGroup"
    li.ListSubItems.Add , "A", rs("Description").Value
    li.Tag = rs("BalSheetID") & ";" & rs("AssociateID") & ";" & rs("AssociateType")
    rs.MoveNext
    lRunCount = lRunCount + 1
Loop
End Sub

Sub MoveOneRight()
Dim sSQL As String
Dim li As ListItem
If lstAGS.ListIndex = -1 Then Exit Sub
If cmbAGS.Text = "Accounts" Then
    sSQL = "Insert Into BalAssociate(BalSheetID,AssociateType,AssociateID) Values(" & lBalSheetID & ",'Account'," & lstAGS.ItemData(lstAGS.ListIndex) & ")"
    cn.Execute sSQL
    Set li = lvwAGS.ListItems.Add(, "ID=" & LastIdentity("BalAssociate"), lRunCount)
    lRunCount = lRunCount + 1
    li.ListSubItems.Add , "Q", "Account"
    li.ListSubItems.Add , "A", lstAGS.Text
    li.Tag = lBalSheetID & ";" & lstAGS.ItemData(lstAGS.ListIndex) & ";Account"
    lstAGS.RemoveItem lstAGS.ListIndex
ElseIf cmbAGS.Text = "Groups" Then
    sSQL = "Insert Into BalAssociate(BalSheetID,AssociateType,AssociateID) Values(" & lBalSheetID & ",'Group'," & lstAGS.ItemData(lstAGS.ListIndex) & ")"
    cn.Execute sSQL
    Set li = lvwAGS.ListItems.Add(, "ID=" & LastIdentity("BalAssociate"), lRunCount)
    lRunCount = lRunCount + 1
    li.ListSubItems.Add , "Q", "Group"
    li.ListSubItems.Add , "A", lstAGS.Text
    li.Tag = lBalSheetID & ";" & lstAGS.ItemData(lstAGS.ListIndex) & ";Group"
    lstAGS.RemoveItem lstAGS.ListIndex
ElseIf cmbAGS.Text = "Sub Groups" Then
    sSQL = "Insert Into BalAssociate(BalSheetID,AssociateType,AssociateID) Values(" & lBalSheetID & ",'SubGroup'," & lstAGS.ItemData(lstAGS.ListIndex) & ")"
    cn.Execute sSQL
    Set li = lvwAGS.ListItems.Add(, "ID=" & LastIdentity("BalAssociate"), lRunCount)
    lRunCount = lRunCount + 1
    li.ListSubItems.Add , "Q", "SubGroup"
    li.ListSubItems.Add , "A", lstAGS.Text
    li.Tag = lBalSheetID & ";" & lstAGS.ItemData(lstAGS.ListIndex) & ";SubGroup"
    lstAGS.RemoveItem lstAGS.ListIndex
Else
    MsgBox "Unexpected Error! - Contact Program Vendor.", vbInformation
End If
End Sub

Sub MoveAllRight()
Dim sSQL As String
Dim li As ListItem
Do Until lstAGS.ListCount < 1
    lstAGS.ListIndex = 0
    MoveOneRight
Loop
End Sub

Function LastIdentity(sTableName As String) As Long
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("select Ident_current('" & sTableName & "')")
    Dim v
    v = rs.Fields(0).Value
    LastIdentity = IIf(IsNull(v), 1, v)
    Set rs = Nothing
End Function

Sub MoveOneLeft()
Dim sSQL As String
Dim li As ListItem
Set li = lvwAGS.SelectedItem
If li Is Nothing Then Exit Sub
cmbAGS.ListIndex = -1
sSQL = "Delete From BalAssociate Where " & li.Key
cn.Execute sSQL
lvwAGS.ListItems.Remove li.index
End Sub

Sub MoveAllLeft()
Dim sSQL As String
sSQL = "Delete From BalAssociate Where BalSheetID=" & lBalSheetID
cn.Execute sSQL
lvwAGS.ListItems.Clear
cmbAGS.ListIndex = -1
lstAGS.Clear
End Sub

Sub ShowForm(sBSheet As String, lBSheetID As Long, Optional Modal, Optional OwnerForm)
lBalSheetID = lBSheetID
lblTitle.Caption = sBSheet
ShowAssociates
If IsMissing(Modal) Then Modal = vbModeless
If IsMissing(OwnerForm) Then Set OwnerForm = Nothing
Me.Show Modal, OwnerForm
End Sub

Private Sub picTitle_Resize()
'lblTitle.Left = picTitle.ScaleWidth / 2 - lblTitle.Width / 2
End Sub
