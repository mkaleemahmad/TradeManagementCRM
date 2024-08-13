VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAccounts3 
   Caption         =   "Accounts"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   7020
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   360
      Left            =   5700
      TabIndex        =   9
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   360
      Left            =   3090
      TabIndex        =   8
      Top             =   4095
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   360
      Left            =   4395
      TabIndex        =   7
      Top             =   4095
      Width           =   1215
   End
   Begin MSComctlLib.Toolbar tbGroup 
      Height          =   330
      Left            =   3960
      TabIndex        =   5
      Top             =   105
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   582
      ButtonWidth     =   1746
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwAccounts 
      Height          =   3135
      Left            =   60
      TabIndex        =   2
      Top             =   885
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5530
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox cmbSubGroup 
      Height          =   315
      Left            =   945
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   495
      Width           =   2970
   End
   Begin VB.ComboBox cmbGroup 
      Height          =   315
      Left            =   945
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   105
      Width           =   2970
   End
   Begin MSComctlLib.Toolbar tbSubGroup 
      Height          =   330
      Left            =   3960
      TabIndex        =   6
      Top             =   465
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   582
      ButtonWidth     =   1746
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sub Group"
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   555
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   75
      TabIndex        =   3
      Top             =   165
      Width           =   435
   End
End
Attribute VB_Name = "frmAccounts3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbGroup_Click()
Dim cCSubGrp As New CSubGroup
If cmbGroup.ListIndex <> -1 Then
  cCSubGrp.PopComboBox cmbSubGroup, cmbGroup.ItemData(cmbGroup.ListIndex)
Else
  cCSubGrp.PopComboBox cmbSubGroup
End If
End Sub

Private Sub cmbSubGroup_Click()
Dim cCA As New cAccounts2
If cmbSubGroup.ListIndex <> -1 Then
'  cCA.PopListView lvwAccounts, cmbGroup.ItemData(cmbGroup.ListIndex), cmbSubGroup.ItemData(cmbSubGroup.ListIndex)
End If
End Sub

Private Sub Form_Load()
Dim cCGrp As New CGroup
cCGrp.PopComboBox cmbGroup
End Sub
