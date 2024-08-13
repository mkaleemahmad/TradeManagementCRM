VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAccount2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8565
   Begin VB.TextBox txtAddress 
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   3
      ToolTipText     =   "Account Name"
      Top             =   1080
      Width           =   5000
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Details"
      Height          =   345
      Left            =   5880
      TabIndex        =   13
      Top             =   4245
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   4440
      TabIndex        =   12
      Top             =   4245
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   3000
      TabIndex        =   14
      Top             =   4245
      Width           =   990
   End
   Begin VB.TextBox txtAccountID 
      Height          =   315
      Left            =   2295
      MaxLength       =   9
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Account ID"
      Top             =   105
      Width           =   5000
   End
   Begin VB.TextBox txtAccount 
      Height          =   315
      Left            =   2295
      MaxLength       =   50
      TabIndex        =   2
      ToolTipText     =   "Account Name"
      Top             =   600
      Width           =   5000
   End
   Begin VB.ComboBox cmbGroup 
      Height          =   315
      Left            =   2295
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3445
      Width           =   5000
   End
   Begin MSComctlLib.Toolbar tbGroup 
      Height          =   330
      Left            =   7380
      TabIndex        =   11
      Top             =   3465
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
   Begin MSMask.MaskEdBox txtContactName 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtMobileNumber 
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   2520
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtPhone1 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   3000
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtCity 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   19
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "City:"
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   18
      Top             =   1655
      Width           =   300
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Phone Number"
      Height          =   195
      Index           =   3
      Left            =   1080
      TabIndex        =   17
      Top             =   3120
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Mobile Number"
      Height          =   195
      Index           =   14
      Left            =   1080
      TabIndex        =   16
      Top             =   2640
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Contact Name:"
      Height          =   195
      Index           =   6
      Left            =   1080
      TabIndex        =   15
      Top             =   2160
      Width           =   1065
   End
   Begin VB.Label lblAccountID 
      AutoSize        =   -1  'True
      Caption         =   "Account &ID"
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   200
      Width           =   810
   End
   Begin VB.Label lblAccountName 
      AutoSize        =   -1  'True
      Caption         =   "Account &Name"
      Height          =   195
      Left            =   1080
      TabIndex        =   8
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label lblGroup 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   3560
      Width           =   435
   End
End
Attribute VB_Name = "frmAccount2"
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


Private Sub UpdateData(bUpdate_Class As Boolean)
If bUpdate_Class Then
  m_cAccounts2.AccountName = txtAccount
  m_cAccounts2.Address = txtAddress
  m_cAccounts2.City = txtCity
  m_cAccounts2.ContactName = txtContactName
  m_cAccounts2.MobileNumber = txtMobileNumber
  m_cAccounts2.PhoneNumber1 = txtPhone1
  
  
  With cmbGroup
    If .ListIndex < 0 Then
      m_cAccounts2.GroupID = 0
      m_sGDesc = ""
    Else
      m_cAccounts2.GroupID = .ItemData(.ListIndex)
      m_sGDesc = .Text
    End If
  End With
  
Else
  txtAccountID = m_cAccounts2.AccountID
  txtAccount = m_cAccounts2.AccountName
  txtAddress = m_cAccounts2.Address
  txtCity = m_cAccounts2.City
  txtContactName = m_cAccounts2.ContactName
  txtMobileNumber = m_cAccounts2.MobileNumber
  txtPhone1 = m_cAccounts2.PhoneNumber1
  
  cmbGroup.ListIndex = TellIndexInDataItem(cmbGroup, m_cAccounts2.GroupID)
'  cmbSubGroup.ListIndex = TellIndexInDataItem(cmbSubGroup, m_cAccounts2.SubGroupID)
End If
End Sub

Private Sub cmbGroup_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbCtrlMask Then
    Select Case KeyCode
    Case vbKeyN
        NewGroup
    Case vbKeyE
        EditGroup
    Case vbKeyR, vbKeyD
        RemoveGroup
    End Select
End If
End Sub

Private Sub cmbSubGroup_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbCtrlMask Then
    Select Case KeyCode
    Case vbKeyN
        NewSubGroup
    Case vbKeyE
        'EditSubGroup
    Case vbKeyR, vbKeyD
       ' RemoveSubGroup
    End Select
End If
End Sub

Private Sub cmdCancel_Click()
m_bOK = False
Unload Me
End Sub

Private Sub cmdDetails_Click()
    frmAcctSub.ShowMe m_cAccounts2
End Sub

Private Sub cmdOK_Click()
    Dim sWarn As String
    sWarn = ""
    If txtAccount = "" Then
      sWarn = "No Account Name is specified."
    End If
    If sWarn = "" Or m_bDeleting Then
        If m_bDeleting And Not m_cAccounts2.CanDelete(m_cAccounts2.AccountID) Then
            MsgBox "This Account can not be deleted.", vbInformation
            m_bOK = False
        Else
            m_bOK = True
            Unload Me
        End If
    Else
      MsgBox "Following errors have found, correct them first" & vbCrLf & vbCrLf & sWarn, vbCritical
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
      SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()
    Me.Top = (frmMain.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmMain.ScaleWidth - Me.ScaleWidth) / 2
    m_bOK = False
   ' chkActive_Click
   ' chkEditable_Click
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
     ' cmbSubGroup.ListIndex = -1
     ' chkActive.value = vbUnchecked
     ' chkEditable.value = vbUnchecked
      cmdOK.Caption = "&Save"
    ElseIf sNED = "e" Then
      txtAccountID.Locked = True
      cmdOK.Caption = "&Save"
    ElseIf sNED = "d" Or sNED = "r" Then
      cmdOK.Caption = "&Delete"
      txtAccountID.Locked = True
      txtAccount.Locked = True
      cmbGroup.Locked = True
      'cmbSubGroup.Locked = True
      'chkActive.Enabled = False
      'chkEditable.Enabled = False
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
   ' Dim cSGrp As New CSubGroup
   ' cSGrp.PopComboBox cmbSubGroup
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
      'EditSubGroup
    Case "remove"
      'RemoveSubGroup
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
      'cSGrp.Add2ComboBox cmbSubGroup
    End If
End Sub


Private Sub txtAccount_GotFocus()
    HighlightText txtAccount
End Sub

