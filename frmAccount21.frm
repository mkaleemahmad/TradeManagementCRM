VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAccount21 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   3225
      TabIndex        =   14
      Top             =   2250
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   1965
      TabIndex        =   15
      Top             =   2250
      Width           =   1215
   End
   Begin VB.CheckBox chkActive 
      Caption         =   "Yes/No"
      Height          =   195
      Left            =   1335
      TabIndex        =   11
      Top             =   1653
      Width           =   1005
   End
   Begin VB.TextBox txtAccountID 
      Height          =   315
      Left            =   1335
      MaxLength       =   9
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Account ID"
      Top             =   105
      Width           =   3330
   End
   Begin VB.TextBox txtAccount 
      Height          =   315
      Left            =   1335
      MaxLength       =   50
      TabIndex        =   3
      ToolTipText     =   "Account Name"
      Top             =   492
      Width           =   3330
   End
   Begin VB.CheckBox chkEditable 
      Caption         =   "Yes/No"
      Height          =   345
      Left            =   1335
      TabIndex        =   13
      Top             =   1920
      Width           =   1005
   End
   Begin VB.ComboBox cmbGroup 
      Height          =   315
      Left            =   1335
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   879
      Width           =   3330
   End
   Begin VB.ComboBox cmbSubGroup 
      Height          =   315
      Left            =   1335
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1266
      Width           =   3330
   End
   Begin MSComctlLib.Toolbar tbSubGroup 
      Height          =   330
      Left            =   4740
      TabIndex        =   9
      Top             =   1258
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   582
      ButtonWidth     =   1508
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Action"
            Key             =   "Action"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "New"
                  Text            =   "&New"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Edit"
                  Text            =   "&Edit"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Remove"
                  Text            =   "&Remove"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbGroup 
      Height          =   330
      Left            =   4740
      TabIndex        =   6
      Top             =   871
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   582
      ButtonWidth     =   1508
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Action"
            Key             =   "Action"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "New"
                  Text            =   "&New"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Edit"
                  Text            =   "&Edit"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Remove"
                  Text            =   "&Remove"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLock 
      AutoSize        =   -1  'True
      Caption         =   "Locked"
      Height          =   195
      Left            =   105
      TabIndex        =   12
      Top             =   1995
      Width           =   540
   End
   Begin VB.Label lblActive 
      AutoSize        =   -1  'True
      Caption         =   "Active"
      Height          =   195
      Left            =   105
      TabIndex        =   10
      Top             =   1653
      Width           =   450
   End
   Begin VB.Label lblAccountID 
      AutoSize        =   -1  'True
      Caption         =   "Account &ID"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   165
      Width           =   810
   End
   Begin VB.Label lblAccountName 
      AutoSize        =   -1  'True
      Caption         =   "Account &Name"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   552
      Width           =   1065
   End
   Begin VB.Label lblGroup 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   939
      Width           =   435
   End
   Begin VB.Label lblSubGroup 
      AutoSize        =   -1  'True
      Caption         =   "Sub Group"
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   1326
      Width           =   765
   End
End
Attribute VB_Name = "frmAccount21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_bOK As Boolean
Dim m_bDeleting As Boolean

Dim m_sGDesc As String
Dim m_sSubGDesc As String

Dim m_cAccounts2 As cAccounts2

Private Sub chkActive_Click()
With chkActive
  If .Value = vbChecked Then
    .Caption = "Yes"
  ElseIf .Value = vbUnchecked Then
    .Caption = "No"
  Else
    .Caption = "Yes / NO"
  End If
End With
End Sub

Private Sub chkEditable_Click()
With chkEditable
  If .Value = vbChecked Then
    .Caption = "Yes"
  ElseIf .Value = vbUnchecked Then
    .Caption = "No"
  Else
    .Caption = "Yes / NO"
  End If
End With
End Sub

Private Sub UpdateData(bUpdate_Class As Boolean)
If bUpdate_Class Then
  m_cAccounts2.AccountName = txtAccount
  With cmbGroup
    If .ListIndex < 0 Then
      m_cAccounts2.GroupID = 0
      m_sGDesc = ""
    Else
      m_cAccounts2.GroupID = .ItemData(.ListIndex)
      m_sGDesc = .Text
    End If
  End With
  With cmbSubGroup
    If .ListIndex < 0 Then
      m_cAccounts2.SubGroupID = 0
      m_sSubGDesc = ""
    Else
      m_cAccounts2.SubGroupID = .ItemData(.ListIndex)
      m_sSubGDesc = .Text
    End If
  End With
  m_cAccounts2.Active = IIf(chkActive.Value = vbChecked, True, False)
  m_cAccounts2.Editable = IIf(chkEditable.Value = vbChecked, True, False)
Else
  txtAccountID = m_cAccounts2.AccountID
  txtAccount = m_cAccounts2.AccountName
  cmbGroup.ListIndex = TellIndexInDataItem(cmbGroup, m_cAccounts2.GroupID)
  cmbSubGroup.ListIndex = TellIndexInDataItem(cmbSubGroup, m_cAccounts2.SubGroupID)
  chkActive.Value = IIf(m_cAccounts2.Active, vbChecked, vbUnchecked)
  chkEditable.Value = IIf(m_cAccounts2.Editable, vbChecked, vbUnchecked)
End If
End Sub

'Function TellIndexInDataItem(refCmbBx As ComboBox, lItem As Long) As Long
'Dim k As Long
'If refCmbBx.ListCount < 1 Then
'  TellIndexInDataItem = -1
'  Exit Function
'End If
'For k = 0 To refCmbBx.ListCount - 1
'  If refCmbBx.ItemData(k) = lItem Then Exit For
'Next
'If k > refCmbBx.ListCount - 1 Then
'  TellIndexInDataItem = -1
'Else
'  TellIndexInDataItem = k
'End If
'End Function

Private Sub cmdCancel_Click()
m_bOK = False
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim sWarn As String
sWarn = ""
If txtAccount = "" Then
  sWarn = "No Account Name is specified."
End If
'If cmbGroup.ListCount > 0 And cmbGroup.ListIndex = -1 Then
'  sWarn = sWarn & vbCrLf & "No Group is selected."
'End If
'If cmbSubGroup.ListCount > 0 And cmbSubGroup.ListIndex = -1 Then
'  sWarn = sWarn & vbCrLf & "No Sub Group is selected."
'End If
If sWarn = "" Or m_bDeleting Then
  m_bOK = True
  Unload Me
Else
  MsgBox "Following errors have found, correct them first" & vbCrLf & vbCrLf & sWarn, vbCritical
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Load()
m_bOK = False
chkActive_Click
chkEditable_Click
PopCombos
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
UpdateData True
End Sub

Public Function ShowForm(cAccounts2 As cAccounts2, ByVal sNED As String, Optional Modal, Optional OwnerForm) As Boolean
Attribute ShowForm.VB_Description = "NED is extracted from ""New"", ""Edit"",""Delete"""
sNED = LCase(sNED)
Set m_cAccounts2 = cAccounts2
Load Me
UpdateData False
m_bDeleting = False
If sNED = "n" Then
  txtAccountID.Locked = True
  txtAccountID.Text = "(New)"
  txtAccountID.ForeColor = &HC0C0C0
  txtAccount.Text = ""
  cmbGroup.ListIndex = -1
  cmbSubGroup.ListIndex = -1
  chkActive.Value = vbUnchecked
  chkEditable.Value = vbUnchecked
  
  cmdOK.Caption = "&Save"
ElseIf sNED = "e" Then
  txtAccountID.Locked = True
  cmdOK.Caption = "&Save"
ElseIf sNED = "d" Or sNED = "r" Then
  cmdOK.Caption = "&Delete"
  txtAccountID.Locked = True
  txtAccount.Locked = True
  cmbGroup.Locked = True
  cmbSubGroup.Locked = True
  chkActive.Enabled = False
  chkEditable.Enabled = False
  m_bDeleting = True

Else
End If
If IsMissing(Modal) And IsMissing(OwnerForm) Then
  Show
ElseIf IsMissing(Modal) Then
  Show , OwnerForm
ElseIf IsMissing(OwnerForm) Then
  Show Modal
Else
  Show Modal, OwnerForm
End If
ShowForm = m_bOK
End Function


Sub PopCombos()
    Dim cgrp As New CGroup
    cgrp.PopComboBox cmbGroup
    Dim cSGrp As New CSubGroup
    cSGrp.PopComboBox cmbSubGroup
End Sub

Public Property Get GroupDescription() As Variant
    GroupDescription = m_sGDesc
End Property

Public Property Get SubGroupDescription()
    SubGroupDescription = m_sSubGDesc
End Property

Private Sub tbSubGroup_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case LCase(ButtonMenu.Key)
    Case "new"
      NewSubGroup
    Case "edit"
      EditSubGroup
    Case "remove"
      RemoveSubGroup
    End Select
End Sub

Private Sub tbGroup_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case LCase(ButtonMenu.Key)
    Case "new"
      NewGroup
    Case "edit"
      EditGroup
    Case "remove"
      RemoveGroup
    End Select
End Sub

Sub NewGroup()
    Dim cgrp As New CGroup
    If frmGroup.ShowForm(cgrp, "n", vbModal, Me) Then
      cgrp.Save True
      cgrp.Add2ComboBox cmbGroup
    End If
End Sub
Sub EditGroup()
    If cmbGroup.ListIndex = -1 Then Exit Sub
    Dim cgrp As New CGroup
    cgrp.Init cmbGroup.ItemData(cmbGroup.ListIndex)
    If frmGroup.ShowForm(cgrp, "e", vbModal, Me) Then
      cgrp.Save False
      cgrp.UpdateComboBox cmbGroup
    End If
End Sub
Sub RemoveGroup()
    If cmbGroup.ListIndex = -1 Then Exit Sub
    Dim cgrp As New CGroup
    cgrp.Init cmbGroup.ItemData(cmbGroup.ListIndex)
    If frmGroup.ShowForm(cgrp, "d", vbModal, Me) Then
      cgrp.Remove
      cgrp.RemoveFromComboBox cmbGroup
    End If
End Sub
' Sub Group
Sub NewSubGroup()
    Dim cSGrp As New CSubGroup
    If frmSubGroup.ShowForm(cSGrp, "n", vbModal, Me) Then
      cSGrp.Save True
      cSGrp.Add2ComboBox cmbSubGroup
    End If
End Sub
Sub EditSubGroup()
    If cmbSubGroup.ListIndex = -1 Then Exit Sub
    Dim cSGrp As New CSubGroup
    cSGrp.Init cmbSubGroup.ItemData(cmbSubGroup.ListIndex)
    If frmSubGroup.ShowForm(cSGrp, "e", vbModal, Me) Then
      cSGrp.Save False
      cSGrp.UpdateComboBox cmbSubGroup
    End If
End Sub
Sub RemoveSubGroup()
    If cmbSubGroup.ListIndex = -1 Then Exit Sub
    Dim cSGrp As New CSubGroup
    cSGrp.Init cmbSubGroup.ItemData(cmbSubGroup.ListIndex)
    If frmSubGroup.ShowForm(cSGrp, "d", vbModal, Me) Then
      cSGrp.Remove
      cSGrp.RemoveFromComboBox cmbSubGroup
    End If
End Sub

Private Sub txtAccount_GotFocus()
    HighlightText txtAccount
End Sub

