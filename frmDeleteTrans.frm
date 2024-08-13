VERSION 5.00
Begin VB.Form frmDeleteTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Transaction"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4020
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   1305
      TabIndex        =   3
      Top             =   615
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   360
      Left            =   2595
      TabIndex        =   2
      Top             =   615
      Width           =   1215
   End
   Begin VB.TextBox txtTransID 
      Height          =   315
      Left            =   1470
      MaxLength       =   9
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   135
      Width           =   2340
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Number :"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   1140
   End
End
Attribute VB_Name = "frmDeleteTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_sTransType As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Are you sure to delete ?", vbQuestion + vbYesNo) = vbYes Then
    RemoveTrans
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{Tab}"
Else
    AcceptKeys KeyAscii, Integers
End If
End Sub

Private Sub Form_Load()
txtTransID = ""
End Sub

Public Sub ShowForm(sTransType As String, Optional Modal, Optional OwnerForm)
If IsMissing(Modal) Then Modal = vbModeless
If IsMissing(OwnerForm) Then Set OwnerForm = Nothing
m_sTransType = sTransType
Me.Show Modal, OwnerForm
End Sub

Sub RemoveTrans()
If Trim(txtTransID.Text) = "" Then
    MsgBox "Can not delete, as no  Entry Number is specified.", vbExclamation
ElseIf IsNumeric(txtTransID.Text) Then
    AddToEvntLg "Delete", m_sTransType, CLng(txtTransID.Text)
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = m_objConnectDB.cnnMyshop
    cmd.CommandText = "spDeleteTrans"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters.Item("@TransID").value = CLng(txtTransID.Text)
    cmd.Parameters.Item("@TransType").value = m_sTransType
    cmd.Execute
    Set cmd = Nothing
    MsgBox "Deletion Completed.", vbInformation
    txtTransID.Text = ""
Else
    MsgBox "Can not delete, as invalid input in Entry Number", vbExclamation
End If
End Sub

Private Sub txtTransID_Change()
cmdDelete.Enabled = Len(txtTransID.Text) > 0
End Sub
