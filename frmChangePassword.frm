VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4155
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtVerifyPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1245
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   975
      Width           =   2805
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   1575
      TabIndex        =   7
      Top             =   1410
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Save"
      Height          =   360
      Left            =   2835
      TabIndex        =   6
      Top             =   1410
      Width           =   1215
   End
   Begin VB.TextBox txtNewPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1245
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   555
      Width           =   2805
   End
   Begin VB.TextBox txtOldPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1245
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   135
      Width           =   2805
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Verify Password"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1035
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&New Password"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   615
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Old Password"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   195
      Width           =   975
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
If ChangePassword Then
    MsgBox "Password Changed Successfully.", vbInformation
    Unload Me
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  KeyAscii = 0
  SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Load()
Dim txtBx
For Each txtBx In Array(txtNewPass, txtOldPass, txtVerifyPass)
  txtBx.Text = ""
Next
Me.Caption = Me.Caption & " [" & frmLogin.UserName & "]"
End Sub

Function ChangePassword() As Boolean
Dim cU As cUser
If txtOldPass <> frmLogin.Password Then
  ChangePassword = False
  MsgBox "Old Password can not be verified", vbOKOnly + vbCritical
ElseIf txtNewPass <> txtVerifyPass Then
  ChangePassword = False
  MsgBox "New Password can not be verified.", vbOKOnly + vbCritical
ElseIf txtNewPass = "" Then
  ChangePassword = False
  MsgBox "Password can not be zero length.", vbCritical
Else
  Set cU = New cUser
  cU.InitWithName frmLogin.UserName
  If Not cU.IsValid Then
    MsgBox "Current changes can not be applied.", vbOKOnly + vbCritical
    ChangePassword = False
    Exit Function
  End If
  cU.Password = txtNewPass
  cU.Save
  Set cU = Nothing
  frmLogin.Password = txtNewPass
  ChangePassword = True
End If
End Function
