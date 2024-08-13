VERSION 5.00
Begin VB.Form frmControlAccounts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Accounts"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frmControlAccounts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6690
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Left            =   4155
      TabIndex        =   39
      Top             =   4020
      Width           =   1215
   End
   Begin VB.TextBox txtSupANo 
      Height          =   315
      Left            =   1575
      MaxLength       =   9
      TabIndex        =   35
      Top             =   3615
      Width           =   1215
   End
   Begin VB.TextBox txtSupAName 
      Height          =   315
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3615
      Width           =   3285
   End
   Begin VB.CommandButton cmdSup 
      Caption         =   "?"
      Height          =   315
      Left            =   2880
      TabIndex        =   36
      Top             =   3615
      Width           =   330
   End
   Begin VB.TextBox txtCANo 
      Height          =   315
      Left            =   1575
      MaxLength       =   9
      TabIndex        =   31
      Top             =   3210
      Width           =   1215
   End
   Begin VB.TextBox txtCAName 
      Height          =   315
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3210
      Width           =   3285
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "?"
      Height          =   315
      Left            =   2880
      TabIndex        =   32
      Top             =   3210
      Width           =   330
   End
   Begin VB.CommandButton cmdP 
      Caption         =   "?"
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      Top             =   780
      Width           =   330
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "?"
      Height          =   315
      Left            =   2880
      TabIndex        =   12
      Top             =   1185
      Width           =   330
   End
   Begin VB.CommandButton cmdPR 
      Caption         =   "?"
      Height          =   315
      Left            =   2880
      TabIndex        =   16
      Top             =   1590
      Width           =   330
   End
   Begin VB.CommandButton cmdSR 
      Caption         =   "?"
      Height          =   315
      Left            =   2880
      TabIndex        =   20
      Top             =   1995
      Width           =   330
   End
   Begin VB.CommandButton cmdDR 
      Caption         =   "?"
      Height          =   315
      Left            =   2880
      TabIndex        =   24
      Top             =   2400
      Width           =   330
   End
   Begin VB.CommandButton cmdDO 
      Caption         =   "?"
      Height          =   315
      Left            =   2880
      TabIndex        =   28
      Top             =   2805
      Width           =   330
   End
   Begin VB.CommandButton cmdCIH 
      Caption         =   "?"
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   390
      Width           =   330
   End
   Begin VB.TextBox txtCIHAName 
      Height          =   315
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   390
      Width           =   3285
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   5415
      TabIndex        =   38
      Top             =   4020
      Width           =   1215
   End
   Begin VB.TextBox txtDOAName 
      Height          =   315
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2805
      Width           =   3285
   End
   Begin VB.TextBox txtDRAName 
      Height          =   315
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2400
      Width           =   3285
   End
   Begin VB.TextBox txtSRAName 
      Height          =   315
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1995
      Width           =   3285
   End
   Begin VB.TextBox txtPRAName 
      Height          =   315
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1590
      Width           =   3285
   End
   Begin VB.TextBox txtSAName 
      Height          =   315
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1185
      Width           =   3285
   End
   Begin VB.TextBox txtPAName 
      Height          =   315
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   780
      Width           =   3285
   End
   Begin VB.TextBox txtDOANo 
      Height          =   315
      Left            =   1575
      MaxLength       =   9
      TabIndex        =   27
      Top             =   2805
      Width           =   1215
   End
   Begin VB.TextBox txtDRANo 
      Height          =   315
      Left            =   1575
      MaxLength       =   9
      TabIndex        =   23
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtSRANo 
      Height          =   315
      Left            =   1575
      MaxLength       =   9
      TabIndex        =   19
      Top             =   1995
      Width           =   1215
   End
   Begin VB.TextBox txtPRANo 
      Height          =   315
      Left            =   1575
      MaxLength       =   9
      TabIndex        =   15
      Top             =   1590
      Width           =   1215
   End
   Begin VB.TextBox txtSANo 
      Height          =   315
      Left            =   1575
      MaxLength       =   9
      TabIndex        =   11
      Top             =   1185
      Width           =   1215
   End
   Begin VB.TextBox txtPANo 
      Height          =   315
      Left            =   1560
      MaxLength       =   9
      TabIndex        =   7
      Top             =   780
      Width           =   1215
   End
   Begin VB.TextBox txtCIHANo 
      Height          =   315
      Left            =   1575
      MaxLength       =   9
      TabIndex        =   3
      Top             =   390
      Width           =   1215
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Suppliers"
      Height          =   195
      Left            =   75
      TabIndex        =   34
      Top             =   3675
      Width           =   645
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Customers"
      Height          =   195
      Left            =   75
      TabIndex        =   30
      Top             =   3270
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Account Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4365
      TabIndex        =   1
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label Label9 
      Caption         =   "Account No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Discounts &Offered"
      Height          =   195
      Left            =   75
      TabIndex        =   26
      Top             =   2865
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Discounts &Received"
      Height          =   195
      Left            =   75
      TabIndex        =   22
      Top             =   2460
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "S&ales Return"
      Height          =   195
      Left            =   75
      TabIndex        =   18
      Top             =   2055
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "P&urchase Return"
      Height          =   195
      Left            =   75
      TabIndex        =   14
      Top             =   1650
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Sales "
      Height          =   195
      Left            =   75
      TabIndex        =   10
      Top             =   1245
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Purchase"
      Height          =   195
      Left            =   75
      TabIndex        =   6
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Cash In Hand"
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   450
      Width           =   975
   End
End
Attribute VB_Name = "frmControlAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_cCA As New ControlAccounts
Dim m_bErrorOccurred As Boolean
Dim m_sWarn As String

Private Sub cmdC_Click()
With frmFindAccount2
   .Show vbModal, Me
   If Not .m_bOK Then Exit Sub
   txtCANo.Text = .m_lAccountID
   txtCANo.DataChanged = True
   txtCAName.Text = .m_sAccountName
   txtCAName.DataChanged = True
End With
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCIH_Click()
With frmFindAccount2
   .Show vbModal, Me
   If Not .m_bOK Then Exit Sub
   txtCIHANo.Text = .m_lAccountID
   txtCIHANo.DataChanged = True
   txtCIHAName.Text = .m_sAccountName
   txtCIHAName.DataChanged = True
End With
End Sub

Private Sub cmdDO_Click()
With frmFindAccount2
   .Show vbModal, Me
   If Not .m_bOK Then Exit Sub
   txtDOANo.Text = .m_lAccountID
   txtDOANo.DataChanged = True
   txtDOAName.Text = .m_sAccountName
   txtDOAName.DataChanged = True
End With
End Sub

Private Sub cmdDR_Click()
With frmFindAccount2
   .Show vbModal, Me
   If Not .m_bOK Then Exit Sub
   txtDRANo.Text = .m_lAccountID
   txtDRANo.DataChanged = True
   txtDRAName.Text = .m_sAccountName
   txtDRAName.DataChanged = True
End With
End Sub

Private Sub cmdP_Click()
With frmFindAccount2
   .Show vbModal, Me
   If Not .m_bOK Then Exit Sub
   txtPANo.Text = .m_lAccountID
   txtPANo.DataChanged = True
   txtPAName.Text = .m_sAccountName
   txtPAName.DataChanged = True
End With
End Sub

Private Sub cmdPR_Click()
With frmFindAccount2
   .Show vbModal, Me
   If Not .m_bOK Then Exit Sub
   txtPRANo.Text = .m_lAccountID
   txtPRANo.DataChanged = True
   txtPRAName.Text = .m_sAccountName
   txtPRAName.DataChanged = True
End With
End Sub

Private Sub cmdS_Click()
With frmFindAccount2
   .Show vbModal, Me
   If Not .m_bOK Then Exit Sub
   txtSANo.Text = .m_lAccountID
   txtSANo.DataChanged = True
   txtSAName.Text = .m_sAccountName
   txtSAName.DataChanged = True
End With
End Sub

Private Sub cmdSave_Click()
ValidateAll
If m_bErrorOccurred Then
    MsgBox "Can not Save as Following error(s) occured" & vbCrLf & m_sWarn, vbCritical
    Exit Sub
End If
m_sWarn = ""
SaveTextBoxes
If m_bErrorOccurred Then
  MsgBox " Following errors are encountered while saving, correct them first" & vbCrLf & _
    m_sWarn, vbOKOnly + vbCritical
  m_cCA.Purge
  m_cCA.Initialize
  
Else
    MsgBox "Control Accounts saved.", vbInformation
  ' Unload Me
End If
End Sub

Private Sub cmdSR_Click()
With frmFindAccount2
   .Show vbModal, Me
   If Not .m_bOK Then Exit Sub
   txtSRANo.Text = .m_lAccountID
   txtSRANo.DataChanged = True
   txtSRAName.Text = .m_sAccountName
   txtSRAName.DataChanged = True
End With
End Sub

Private Sub cmdSup_Click()
With frmFindAccount2
   .Show vbModal, Me
   If Not .m_bOK Then Exit Sub
   txtSupANo.Text = .m_lAccountID
   txtSupANo.DataChanged = True
   txtSupAName.Text = .m_sAccountName
   txtSupAName.DataChanged = True
End With
End Sub

Private Sub Form_Load()
m_cCA.Initialize
PopulateTextBoxes
frmMain.SetTlbLayout -1
End Sub

Sub PopulateTextBoxes()
With m_cCA
   txtCIHANo = IsMO(.AccountNo(CashInHand))
   txtCIHAName = .AccountName(CashInHand)
   If txtCIHANo <> "" Then
   If .ExistInTransDet(txtCIHANo) Then
    txtCIHANo.Locked = True
    cmdCIH.Enabled = False
   End If
   End If
   
   txtPANo = IsMO(.AccountNo(Purchase))
   txtPAName = .AccountName(Purchase)
   If txtPANo <> "" Then
   If .ExistInTransDet(txtPANo) Then
    txtPANo.Locked = True
    cmdP.Enabled = False
   End If
   End If
   
   txtSANo = IsMO(.AccountNo(Sales))
   txtSAName = .AccountName(Sales)
   If txtSANo <> "" Then
   If .ExistInTransDet(txtSANo) Then
    txtSANo.Locked = True
    cmdS.Enabled = False
   End If
   End If
   
   txtPRANo = IsMO(.AccountNo(PurchaseReturn))
   txtPRAName = .AccountName(PurchaseReturn)
   If txtPRANo <> "" Then
   If .ExistInTransDet(txtPRANo) Then
    txtPRANo.Locked = True
    cmdPR.Enabled = False
   End If
   End If
   
   txtSRANo = IsMO(.AccountNo(SalesReturn))
   txtSRAName = .AccountName(SalesReturn)
   If txtSRANo <> "" Then
   If .ExistInTransDet(txtSRANo) Then
      txtSRANo.Locked = True
      cmdSR.Enabled = False
    End If
   End If
   
   txtDRANo = IsMO(.AccountNo(DiscountsReceived))
   txtDRAName = .AccountName(DiscountsReceived)
   If txtDRANo <> "" Then
  If .ExistInTransDet(txtDRANo) Then
    txtDRANo.Locked = True
    cmdDR.Enabled = False
  End If
  End If
   
   txtDOANo = IsMO(.AccountNo(DiscountsOffered))
   txtDOAName = .AccountName(DiscountsOffered)
   If txtDOANo <> "" Then
   If .ExistInTransDet(txtDOANo) Then
    txtDOANo.Locked = True
    cmdDO.Enabled = False
  End If
  End If
   
   txtCANo = IsMO(.AccountNo(Customers))
   txtCAName = .AccountName(Customers)
   If txtCANo <> "" Then
   If .ExistInTransDet(txtCANo) Then
    txtCANo.Locked = True
    cmdC.Enabled = False
   End If
   End If
   
   txtSupANo = IsMO(.AccountNo(Suppliers))
   txtSupAName = .AccountName(Suppliers)
   If txtSupANo <> "" Then
   If .ExistInTransDet(txtSupANo) Then
    txtSupANo.Locked = True
    cmdSup.Enabled = False
   End If
   End If
   
   txtCIHANo.DataChanged = False
   txtPANo.DataChanged = False
   txtSANo.DataChanged = False
   txtPRANo.DataChanged = False
   txtSRANo.DataChanged = False
   txtDRANo.DataChanged = False
   txtDOANo.DataChanged = False
   txtCANo.DataChanged = False
   txtSupANo.DataChanged = False
End With
End Sub

Sub SaveTextBoxes()
Dim bAnyError As Boolean
m_bErrorOccurred = False
bAnyError = False
SaveCIHANo
bAnyError = m_bErrorOccurred Or bAnyError
SavePANo
bAnyError = m_bErrorOccurred Or bAnyError
SaveSANo
bAnyError = m_bErrorOccurred Or bAnyError
SavePRANo
bAnyError = m_bErrorOccurred Or bAnyError
SaveSRANo
bAnyError = m_bErrorOccurred Or bAnyError
SaveDRANo
bAnyError = m_bErrorOccurred Or bAnyError
SaveDOANo
bAnyError = m_bErrorOccurred Or bAnyError
SavecaNo
bAnyError = m_bErrorOccurred Or bAnyError
SaveSupANo
m_bErrorOccurred = bAnyError Or m_bErrorOccurred
If Not m_bErrorOccurred Then PopSelAcntIDs
End Sub

Private Sub Form_Unload(Cancel As Integer)
m_cCA.Purge
frmMain.SetTlbDefLayout
End Sub

Private Sub txtCIHAName_GotFocus()
'HighlightText txtCIHAName
End Sub

Private Sub txtCIHANo_GotFocus()
HighlightText txtCIHANo
End Sub

Private Sub txtCIHANo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   AcceptKeys KeyAscii, Integers
   Exit Sub
End If
With txtCIHANo
   If .DataChanged Then
      txtCIHANo_GotFocus
      'SendKeys "{End}+{Home}"
   Else
        KeyAscii = 0
      SendKeys "{Tab}"
   End If
End With
End Sub

Private Sub SaveCIHANo()
m_bErrorOccurred = False
With txtCIHANo
   If .DataChanged Then
      m_cCA.AccountNo(CashInHand) = IIf(Len(.Text) = 0, -1, .Text)
      txtCIHAName.Text = m_cCA.AccountName(CashInHand)
      If txtCIHAName = "" Then
        m_sWarn = m_sWarn & vbCrLf & "Cash In Hand Account No not valid."
      Else
        .DataChanged = False
        End If
   End If
   Exit Sub
   If txtCIHAName.Text = "" And .Text <> "-1" Then
      m_sWarn = m_sWarn & vbCrLf & "Cash In Hand Account No not valid."
      .Text = ""
      .DataChanged = True
      .SetFocus
      m_bErrorOccurred = True
   End If
End With
End Sub

Private Sub txtDOAName_GotFocus()
'HighlightText txtDOAName
End Sub

Private Sub txtDOANo_GotFocus()
HighlightText txtDOANo
End Sub

Private Sub txtDOANo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   AcceptKeys KeyAscii, Integers
   Exit Sub
End If
With txtDOANo
   If .DataChanged Then
      txtDOANo_GotFocus
   Else
        KeyAscii = 0
      SendKeys "{Tab}"
   End If
End With
End Sub

Private Sub SaveDOANo()
m_bErrorOccurred = False
With txtDOANo
   If .DataChanged Then
      m_cCA.AccountNo(DiscountsOffered) = IIf(Len(.Text) = 0, -1, .Text)
      txtDOAName.Text = m_cCA.AccountName(DiscountsOffered)
      .DataChanged = False
   End If
   Exit Sub
   If txtDOAName.Text = "" And .Text <> "-1" Then
      m_sWarn = m_sWarn & vbCrLf & "Discounts Offered Account No not valid."
      .Text = ""
      .DataChanged = True
      .SetFocus
      m_bErrorOccurred = True
   End If
End With
End Sub

Private Sub txtDRAName_GotFocus()
'HighlightText txtDRAName
End Sub

Private Sub txtDRANo_GotFocus()
HighlightText txtDRANo
End Sub

Private Sub txtDRANo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   AcceptKeys KeyAscii, Integers
   Exit Sub
End If
With txtDRANo
   If .DataChanged Then
      txtDRANo_GotFocus
   Else
    KeyAscii = 0
      SendKeys "{Tab}"
   End If
End With
End Sub

Private Sub SaveDRANo()
m_bErrorOccurred = False
With txtDRANo
   If .DataChanged Then
      m_cCA.AccountNo(DiscountsReceived) = IIf(Len(.Text) = 0, -1, .Text)
      txtDRAName.Text = m_cCA.AccountName(DiscountsReceived)
      .DataChanged = False
   End If
   Exit Sub
   If txtDRAName.Text = "" And .Text <> "-1" Then
      m_sWarn = m_sWarn & vbCrLf & "Discounts Received Account No not valid."
      .Text = ""
      .DataChanged = True
      .SetFocus
      m_bErrorOccurred = True
   End If
End With
End Sub

Private Sub txtPAName_GotFocus()
'HighlightText txtPAName
End Sub

Private Sub txtPANo_GotFocus()
HighlightText txtPANo
End Sub

Private Sub txtPANo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   AcceptKeys KeyAscii, Integers
   Exit Sub
End If
With txtPANo
   If .DataChanged Then
      txtPANo_GotFocus
   Else
        KeyAscii = 0
      SendKeys "{Tab}"
   End If
End With
End Sub

Private Sub SavePANo()
m_bErrorOccurred = False
With txtPANo
   If .DataChanged Then
      m_cCA.AccountNo(Purchase) = IIf(Len(.Text) = 0, -1, .Text)
      txtPAName.Text = m_cCA.AccountName(Purchase)
      .DataChanged = False
   End If
   Exit Sub
   If txtPAName.Text = "" And .Text <> "-1" Then
      m_sWarn = m_sWarn & vbCrLf & "Purchase Account No not valid."
      .Text = ""
      .DataChanged = True
      .SetFocus
      m_bErrorOccurred = True
   End If
End With
End Sub

Private Sub txtPRAName_GotFocus()
'HighlightText txtPRAName
End Sub

Private Sub txtPRANo_GotFocus()
HighlightText txtPRANo
End Sub

Private Sub txtPRANo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   AcceptKeys KeyAscii, Integers
   Exit Sub
End If
With txtPRANo
   If .DataChanged Then
      txtPRANo_GotFocus
   Else
        KeyAscii = 0
      SendKeys "{Tab}"
   End If
End With
End Sub

Private Sub SavePRANo()
m_bErrorOccurred = False
With txtPRANo
   If .DataChanged Then
      m_cCA.AccountNo(PurchaseReturn) = IIf(Len(.Text) = 0, -1, .Text)
      txtPRAName.Text = m_cCA.AccountName(PurchaseReturn)
      .DataChanged = False
   End If
   Exit Sub
   If txtPRAName.Text = "" And .Text <> "-1" Then
      m_sWarn = m_sWarn & vbCrLf & "Purchase Return Account No not valid."
      .Text = ""
      .DataChanged = True
      .SetFocus
      m_bErrorOccurred = True
   End If
End With
End Sub

Private Sub txtSAName_GotFocus()
'HighlightText txtSAName
End Sub

Private Sub txtSANo_GotFocus()
HighlightText txtSANo
End Sub

Private Sub txtSANo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   AcceptKeys KeyAscii, Integers
   Exit Sub
End If
With txtSANo
   If .DataChanged Then
      txtSANo_GotFocus
   Else
        KeyAscii = 0
      SendKeys "{Tab}"
   End If
End With
End Sub

Private Sub SaveSANo()
m_bErrorOccurred = False
With txtSANo
   If .DataChanged Then
      m_cCA.AccountNo(Sales) = IIf(Len(.Text) = 0, -1, .Text)
      txtSAName.Text = m_cCA.AccountName(Sales)
      .DataChanged = False
   End If
   Exit Sub
   If txtSAName.Text = "" And .Text <> "-1" Then
      m_sWarn = m_sWarn & vbCrLf & "Sales Account No not valid."
      .Text = ""
      .DataChanged = True
      .SetFocus
      m_bErrorOccurred = True
   End If
End With
End Sub

Private Sub txtSRAName_GotFocus()
'HighlightText txtSRAName
End Sub

Private Sub txtSRANo_GotFocus()
HighlightText txtSRANo
End Sub

Private Sub txtSRANo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   AcceptKeys KeyAscii, Integers
   Exit Sub
End If
With txtSRANo
   If .DataChanged Then
      txtSRANo_GotFocus
   Else
        KeyAscii = 0
      SendKeys "{Tab}"
   End If
End With
End Sub

Private Sub SaveSRANo()
m_bErrorOccurred = False
With txtSRANo
   If .DataChanged Then
      m_cCA.AccountNo(SalesReturn) = IIf(Len(.Text) = 0, -1, .Text)
      txtSRAName.Text = m_cCA.AccountName(SalesReturn)
      .DataChanged = False
   End If
   Exit Sub
   If txtSRAName.Text = "" And .Text <> "-1" Then
      m_sWarn = m_sWarn & vbCrLf & "Sales Return Account No not valid."
      .Text = ""
      .DataChanged = True
      .SetFocus
      m_bErrorOccurred = True
   End If
End With
End Sub

' ==
Private Sub txtcANo_GotFocus()
HighlightText txtCANo
End Sub

Private Sub txtcaNo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   AcceptKeys KeyAscii, Integers
   Exit Sub
End If
With txtCANo
   If .DataChanged Then
      txtcANo_GotFocus
   Else
        KeyAscii = 0
      SendKeys "{Tab}"
   End If
End With
End Sub

Private Sub SavecaNo()
    m_bErrorOccurred = False
    With txtCANo
       If .DataChanged Then
          m_cCA.AccountNo(Customers) = IIf(Len(.Text) = 0, -1, .Text)
          txtCAName.Text = m_cCA.AccountName(Customers)
          .DataChanged = False
       End If
       Exit Sub
       If txtCAName.Text = "" And .Text <> "-1" Then
          m_sWarn = m_sWarn & vbCrLf & "Customers Account No not valid."
      .Text = ""
      .DataChanged = True
          .SetFocus
          m_bErrorOccurred = True
       End If
    End With
End Sub

' |||||||||||||||||||||||||||||||||||||||||

Private Sub txtSupANo_GotFocus()
HighlightText txtSupANo
End Sub

Private Sub txtSupANo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   AcceptKeys KeyAscii, Integers
   Exit Sub
End If
With txtSupANo
   If .DataChanged Then
      txtSupANo_GotFocus
   Else
        KeyAscii = 0
      SendKeys "{Tab}"
   End If
End With
End Sub

Private Sub SaveSupANo()
m_bErrorOccurred = False
With txtSupANo
   If .DataChanged Then
      m_cCA.AccountNo(Suppliers) = IIf(Len(.Text) = 0, -1, .Text)
      txtSupAName.Text = m_cCA.AccountName(Suppliers)
      .DataChanged = False
   End If
   Exit Sub
   If txtSupAName.Text = "" And .Text <> "-1" Then
      m_sWarn = m_sWarn & vbCrLf & "Suppliers Account No not valid."
      .Text = ""
      .DataChanged = True
      .SetFocus
      m_bErrorOccurred = True
   End If
End With
End Sub
' Is Minus One
Function IsMO(lID As Long) As String
If lID = -1 Then
    IsMO = ""
Else
    IsMO = lID
End If
End Function

Function IsValidAcnt(lAccountID As Long) As Boolean
Dim sSQL As String
Dim rs As ADODB.Recordset
sSQL = "SELECT COUNT(*) AS IAI FROM ACCOUNTS WHERE ID=" & lAccountID
Set rs = m_objConnectDB.cnnMyshop.Execute(sSQL)
IsValidAcnt = CBool(IsNull2(rs!IAI, 0))
Set rs = Nothing
End Function

Function ValidateAll() As Boolean
Dim l As Long
m_sWarn = ""
m_bErrorOccurred = False
With txtCIHANo
    If .Text <> "" Then
        If Not IsValidAcnt(CLng(.Text)) Then
            m_sWarn = m_sWarn & vbCrLf & _
                "Cash In Hand Account is not valid."
            m_bErrorOccurred = True
        End If
    End If
End With

With txtPANo
    If .Text <> "" Then
        If Not IsValidAcnt(CLng(.Text)) Then
            m_sWarn = m_sWarn & vbCrLf & _
                "Purchase Account is not valid."
            m_bErrorOccurred = True
        End If
    End If
End With
With txtSANo
    If .Text <> "" Then
        If Not IsValidAcnt(CLng(.Text)) Then
            m_sWarn = m_sWarn & vbCrLf & _
                "Sale Account is not valid."
            m_bErrorOccurred = True
        End If
    End If
End With
With txtPRANo
    If .Text <> "" Then
        If Not IsValidAcnt(CLng(.Text)) Then
            m_sWarn = m_sWarn & vbCrLf & _
                "Purchase Return Account is not valid."
            m_bErrorOccurred = True
        End If
    End If
End With
With txtSRANo
    If .Text <> "" Then
        If Not IsValidAcnt(CLng(.Text)) Then
            m_sWarn = m_sWarn & vbCrLf & _
                "Sale Return Account is not valid."
            m_bErrorOccurred = True
        End If
    End If
End With
With txtDRANo
    If .Text <> "" Then
        If Not IsValidAcnt(CLng(.Text)) Then
            m_sWarn = m_sWarn & vbCrLf & _
                "Discounts Received Account is not valid."
            m_bErrorOccurred = True
        End If
    End If
End With
With txtDOANo
    If .Text <> "" Then
        If Not IsValidAcnt(CLng(.Text)) Then
            m_sWarn = m_sWarn & vbCrLf & _
                "Discounts Offered Account is not valid."
            m_bErrorOccurred = True
        End If
    End If
End With
With txtCANo
    If .Text <> "" Then
        If Not IsValidAcnt(CLng(.Text)) Then
            m_sWarn = m_sWarn & vbCrLf & _
                "Customers Account is not valid."
            m_bErrorOccurred = True
        End If
    End If
End With
With txtSupANo
    If .Text <> "" Then
        If Not IsValidAcnt(CLng(.Text)) Then
            m_sWarn = m_sWarn & vbCrLf & _
                "Suppliers Account is not valid."
            m_bErrorOccurred = True
        End If
    End If
End With
End Function

Public Function CashInHandAccount() As Long
Dim ca As New ControlAccounts
ca.Initialize
CashInHandAccount = IsMO(ca.AccountNo(CashInHand))
ca.Purge
End Function
