VERSION 5.00
Begin VB.Form frmCash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cash Book Entry"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   HasDC           =   0   'False
   Icon            =   "frmCash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4635
      ScaleHeight     =   270
      ScaleWidth      =   525
      TabIndex        =   10
      Top             =   210
      Width           =   525
      Begin VB.CommandButton cmdAccountID 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   15
         Picture         =   "frmCash.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.TextBox txtAccountID 
      Height          =   315
      Left            =   1245
      MaxLength       =   9
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   180
      Width           =   3705
   End
   Begin VB.TextBox txtAccountName 
      Height          =   315
      Left            =   1245
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   585
      Width           =   3705
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   1200
      TabIndex        =   9
      Top             =   1875
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   2520
      TabIndex        =   8
      Top             =   1875
      Width           =   1215
   End
   Begin VB.TextBox txtAmount 
      Height          =   315
      Left            =   1245
      MaxLength       =   15
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1470
      Width           =   3705
   End
   Begin VB.TextBox txtDescription 
      Height          =   315
      Left            =   1245
      MaxLength       =   50
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1030
      Width           =   3705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   195
      Index           =   3
      Left            =   105
      TabIndex        =   6
      Top             =   1485
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   4
      Top             =   1065
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Account Name"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   660
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Account ID"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "frmCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_sAccountID As Long
Public m_sAccountName As String
Public m_sDescription As String
Public m_sAmount As Double
Private m_bOK As Boolean

Sub ReadData()
  m_sAccountID = IIf(Len(txtAccountID.Text) = 0, 0, txtAccountID.Text)
  m_sAccountName = txtAccountName.Text
  m_sDescription = txtDescription.Text
  m_sAmount = IIf(Len(txtAmount.Text) = 0, 0, txtAmount.Text)
End Sub

Sub WriteData()
txtAccountID.Text = m_sAccountID
txtAccountName.Text = m_sAccountName
txtDescription.Text = m_sDescription
txtAmount.Text = m_sAmount
End Sub

Public Sub setValues(AccountID As Long, AccountName As String, Description As String, Amount As Double)
m_sAccountID = AccountID
m_sAccountName = AccountName
m_sDescription = Description
m_sAmount = Amount
End Sub

Public Sub getValues(AccountID As Long, AccountName As String, Description As String, Amount As Double)
AccountID = m_sAccountID
AccountName = m_sAccountName
Description = m_sDescription
Amount = m_sAmount
End Sub

Public Function ShowForm(AccountID As Long, AccountName As String, Description As String, Amount As Double, NED As String, Optional sTitle As String = "Cash Book Entry") As Boolean
Load Me
If StrComp(NED, "n", vbTextCompare) = 0 Then
   cmdOK.Caption = "&Save"
   setValues 0, "", "", 0
   WriteData
   txtAccountID.Text = ""
   txtAmount.Text = ""
ElseIf StrComp(NED, "e", vbTextCompare) = 0 Then
   cmdOK.Caption = "&Save"
   setValues AccountID, AccountName, Description, Amount
   WriteData
ElseIf StrComp(NED, "d", vbTextCompare) = 0 Then
   cmdOK.Caption = "&Delete"
   setValues AccountID, AccountName, Description, Amount
   WriteData
End If
Me.Caption = sTitle
Me.Show vbModal
getValues AccountID, AccountName, Description, Amount
ShowForm = m_bOK
End Function
Private Sub cmdAccountID_Click()
   frmFindAccount2.Show vbModal, Me
   If frmFindAccount2.m_bOK Then
      txtAccountID.Text = frmFindAccount2.m_lAccountID
      txtAccountName.Text = frmFindAccount2.m_sAccountName
   End If
   txtAccountID.SetFocus

End Sub

Private Sub cmdCancel_Click()
m_bOK = False
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim m_sWarn As String
m_sWarn = ""
Dim rscAccounts As New cAccounts
txtAccountName.Text = rscAccounts.GetAccountName(CLng(Val(txtAccountID.Text)))
If txtAccountID.Text = "" Or txtAccountID.Text = "0" Then
   m_sWarn = "AccountID must be specified." & vbCrLf
End If
If txtAccountName.Text = "" Then
   m_sWarn = m_sWarn & "Account Name is undetermined." & vbCrLf
End If
If txtDescription.Text = "" Then
   m_sWarn = m_sWarn & "Description must be specified." & vbCrLf
End If
If txtAmount.Text = "" Or txtAmount.Text = "0" Then
   m_sWarn = m_sWarn & "Amount must be specified and it can not be zero." & vbCrLf
End If
If m_sWarn = "" Then
   m_bOK = True
   Unload Me
Else
   MsgBox "Following errors has occurred, correct them." & vbCrLf & m_sWarn, vbOKOnly + vbCritical
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
   SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Load()
m_bOK = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ReadData
Cancel = 0
End Sub

Private Sub txtAccountID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   frmFindAccount2.Show vbModal, Me
   If frmFindAccount2.m_bOK Then
      txtAccountID.Text = frmFindAccount2.m_lAccountID
      txtAccountName.Text = frmFindAccount2.m_sAccountName
   End If
   txtAccountID.SetFocus
End If
End Sub

Private Sub txtAccountID_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtAccountID_Validate(Cancel As Boolean)
Dim cAccts As New cAccounts
Dim sAcctName As String
If txtAccountID.Text = "" Then
  txtAccountName = ""
  Exit Sub
End If
sAcctName = cAccts.GetAccountName(CLng(Val(txtAccountID)))
If sAcctName = "" Then
  Cancel = True
  MsgBox "Invalid Account Number.", vbOKOnly + vbCritical
  txtAccountName = ""
  HighlightText txtAccountID
Else
  txtAccountName = sAcctName
End If

End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Floats, txtAmount.Text
End Sub
