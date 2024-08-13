VERSION 5.00
Begin VB.Form frmControlGrp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Groups"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCash 
      Height          =   315
      Left            =   1620
      MaxLength       =   9
      TabIndex        =   38
      Top             =   3555
      Width           =   1215
   End
   Begin VB.TextBox txtCashDesc 
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3555
      Width           =   3285
   End
   Begin VB.CommandButton cmdCash 
      Caption         =   "?"
      Height          =   315
      Left            =   2925
      TabIndex        =   36
      Top             =   3555
      Width           =   330
   End
   Begin VB.TextBox txtCAB 
      Height          =   315
      Left            =   1620
      MaxLength       =   9
      TabIndex        =   3
      Top             =   315
      Width           =   1215
   End
   Begin VB.TextBox txtPar 
      Height          =   315
      Left            =   1605
      MaxLength       =   9
      TabIndex        =   7
      Top             =   705
      Width           =   1215
   End
   Begin VB.TextBox txtCust 
      Height          =   315
      Left            =   1605
      MaxLength       =   9
      TabIndex        =   11
      Top             =   1110
      Width           =   1215
   End
   Begin VB.TextBox txtSUP 
      Height          =   315
      Left            =   1620
      MaxLength       =   9
      TabIndex        =   15
      Top             =   1515
      Width           =   1215
   End
   Begin VB.TextBox txtSL 
      Height          =   315
      Left            =   1620
      MaxLength       =   9
      TabIndex        =   19
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtPur 
      Height          =   315
      Left            =   1620
      MaxLength       =   9
      TabIndex        =   23
      Top             =   2325
      Width           =   1215
   End
   Begin VB.TextBox txtExp 
      Height          =   315
      Left            =   1620
      MaxLength       =   9
      TabIndex        =   27
      Top             =   2730
      Width           =   1215
   End
   Begin VB.TextBox txtParDesc 
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   705
      Width           =   3285
   End
   Begin VB.TextBox txtCustDesc 
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1110
      Width           =   3285
   End
   Begin VB.TextBox txtSUPDesc 
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1515
      Width           =   3285
   End
   Begin VB.TextBox txtSLDesc 
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3285
   End
   Begin VB.TextBox txtPurDesc 
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2325
      Width           =   3285
   End
   Begin VB.TextBox txtExpDesc 
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2730
      Width           =   3285
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   5460
      TabIndex        =   34
      Top             =   3975
      Width           =   1215
   End
   Begin VB.TextBox txtCABDesc 
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   315
      Width           =   3285
   End
   Begin VB.CommandButton cmdCAB 
      Caption         =   "?"
      Height          =   315
      Left            =   2925
      TabIndex        =   4
      Top             =   315
      Width           =   330
   End
   Begin VB.CommandButton cmdExp 
      Caption         =   "?"
      Height          =   315
      Left            =   2925
      TabIndex        =   28
      Top             =   2730
      Width           =   330
   End
   Begin VB.CommandButton cmdPur 
      Caption         =   "?"
      Height          =   315
      Left            =   2925
      TabIndex        =   24
      Top             =   2325
      Width           =   330
   End
   Begin VB.CommandButton cmdSl 
      Caption         =   "?"
      Height          =   315
      Left            =   2925
      TabIndex        =   20
      Top             =   1920
      Width           =   330
   End
   Begin VB.CommandButton cmdSup 
      Caption         =   "?"
      Height          =   315
      Left            =   2925
      TabIndex        =   16
      Top             =   1515
      Width           =   330
   End
   Begin VB.CommandButton cmdCust 
      Caption         =   "?"
      Height          =   315
      Left            =   2925
      TabIndex        =   12
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdPar 
      Caption         =   "?"
      Height          =   315
      Left            =   2925
      TabIndex        =   8
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmdLcAccts 
      Caption         =   "?"
      Height          =   315
      Left            =   2925
      TabIndex        =   32
      Top             =   3135
      Width           =   330
   End
   Begin VB.TextBox txtLcAcctsDesc 
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3135
      Width           =   3285
   End
   Begin VB.TextBox txtLcAccts 
      Height          =   315
      Left            =   1620
      MaxLength       =   9
      TabIndex        =   31
      Top             =   3135
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4200
      TabIndex        =   35
      Top             =   3975
      Width           =   1215
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "C&ash"
      Height          =   195
      Left            =   120
      TabIndex        =   39
      Top             =   3615
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Bank"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Pa&rties"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   765
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Customers"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1170
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Suppliers"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1575
      Width           =   645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Sa&les"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   1980
      Width           =   390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "P&urchases"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   2385
      Width           =   750
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "E&xpense"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   2790
      Width           =   615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Group ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1725
      TabIndex        =   0
      Top             =   45
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Group Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4410
      TabIndex        =   1
      Top             =   45
      Width           =   1545
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "&L/C Accounts"
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   3195
      Width           =   990
   End
End
Attribute VB_Name = "frmControlGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_sWarn As String

Private Sub cmdCAB_Click()
With frmFindGroup
  .Show vbModal, Me
  If .m_bOK Then
    txtCAB = .m_lGroupID
    txtCABDesc = .m_sDesc
    txtCAB.DataChanged = False
  End If
End With

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCash_Click()
With frmFindGroup
  .Show vbModal, Me
  If .m_bOK Then
    txtCash = .m_lGroupID
    txtCashDesc = .m_sDesc
    txtCash.DataChanged = False
  End If
End With
End Sub

Private Sub cmdCust_Click()
With frmFindGroup
  .Show vbModal, Me
  If .m_bOK Then
    txtCust = .m_lGroupID
    txtCustDesc = .m_sDesc
    txtCust.DataChanged = False
  End If
End With

End Sub

Private Sub cmdExp_Click()
With frmFindGroup
  .Show vbModal, Me
  If .m_bOK Then
    txtExp = .m_lGroupID
    txtExpDesc = .m_sDesc
    txtExp.DataChanged = False
  End If
End With

End Sub

Private Sub cmdPar_Click()
With frmFindGroup
  .Show vbModal, Me
  If .m_bOK Then
    txtPar = .m_lGroupID
    txtParDesc = .m_sDesc
    txtPar.DataChanged = False
  End If
End With

End Sub

Private Sub cmdLcAccts_Click()
With frmFindGroup
  .Show vbModal, Me
  If .m_bOK Then
    txtLcAccts = .m_lGroupID
    txtLcAcctsDesc = .m_sDesc
    txtLcAccts.DataChanged = False
  End If
End With
End Sub

Private Sub cmdPur_Click()
With frmFindGroup
  .Show vbModal, Me
  If .m_bOK Then
    txtPur = .m_lGroupID
    txtPurDesc = .m_sDesc
    txtPur.DataChanged = False
  End If
End With

End Sub

Private Sub cmdSave_Click()
ValidateAll
If m_sWarn = "" Then
  SaveAll
  MsgBox "Control Groups saved successfully.", vbInformation
Else
  MsgBox "Follwoing errors have found, correct them first" & vbCrLf & m_sWarn
End If
End Sub

Private Sub cmdSl_Click()
With frmFindGroup
  .Show vbModal, Me
  If .m_bOK Then
    txtSL = .m_lGroupID
    txtSLDesc = .m_sDesc
    txtSL.DataChanged = False
  End If
End With

End Sub

Private Sub cmdSup_Click()
With frmFindGroup
  .Show vbModal, Me
  If .m_bOK Then
    txtSUP = .m_lGroupID
    txtSUPDesc = .m_sDesc
    txtSUP.DataChanged = False
  End If
End With

End Sub

Sub ValidateAll()
Dim cgrp As New CGroup
m_sWarn = ""
With txtCAB
If .Text <> "" Then
  cgrp.Init CLng(.Text)
  If Not cgrp.IsValid Then
    m_sWarn = m_sWarn & vbCrLf & TellCtrlGrp(eGrpBANK) & " must be valid group id"
  End If
End If
End With

With txtPar
If .Text <> "" Then
  cgrp.Init CLng(.Text)
  If Not cgrp.IsValid Then
    m_sWarn = m_sWarn & vbCrLf & TellCtrlGrp(eGrpPARTIES) & " must be valid group id"
  End If
End If
End With

With txtCust
If .Text <> "" Then
  cgrp.Init CLng(.Text)
  If Not cgrp.IsValid Then
    m_sWarn = m_sWarn & vbCrLf & TellCtrlGrp(eGrpCUSTOMERS) & " must be valid group id"
  End If
End If
End With

With txtSUP
If .Text <> "" Then
  cgrp.Init CLng(.Text)
  If Not cgrp.IsValid Then
    m_sWarn = m_sWarn & vbCrLf & TellCtrlGrp(eGrpSUPPLIERS) & " must be valid group id"
  End If
End If
End With

With txtSL
If .Text <> "" Then
  cgrp.Init CLng(.Text)
  If Not cgrp.IsValid Then
    m_sWarn = m_sWarn & vbCrLf & TellCtrlGrp(eGrpSALES) & " must be valid group id"
  End If
End If
End With

With txtPur
If .Text <> "" Then
  cgrp.Init CLng(.Text)
  If Not cgrp.IsValid Then
    m_sWarn = m_sWarn & vbCrLf & TellCtrlGrp(eGrpPURCHASES) & " must be valid group id"
  End If
End If
End With

With txtExp
If .Text <> "" Then
  cgrp.Init CLng(.Text)
  If Not cgrp.IsValid Then
    m_sWarn = m_sWarn & vbCrLf & TellCtrlGrp(eGrpEXPENSE) & " must be valid group id"
  End If
End If
End With

With txtLcAccts
If .Text <> "" Then
  cgrp.Init CLng(.Text)
  If Not cgrp.IsValid Then
    m_sWarn = m_sWarn & vbCrLf & TellCtrlGrp(eGrpLCACCTS) & " must be valid group id"
  End If
End If
End With

With txtCash
If .Text <> "" Then
  cgrp.Init CLng(.Text)
  If Not cgrp.IsValid Then
    m_sWarn = m_sWarn & vbCrLf & TellCtrlGrp(egrpCASH) & " must be valid group id"
  End If
End If
End With

With txtCash
If .Text <> "" Then
  cgrp.Init CLng(.Text)
  If Not cgrp.IsValid Then
    m_sWarn = m_sWarn & vbCrLf & TellCtrlGrp(egrpCASH) & " must be valid group id"
  End If
End If
End With

End Sub


Sub SaveAll()
Dim cgrp As New CControlGroup
With txtCAB
If .Text <> "" Then
  cgrp.Init TellCtrlGrp(eGrpBANK)
  If cgrp.IsValid Then
    cgrp.GroupID = CLng(NumVal(.Text))
    cgrp.Save False
  Else
    cgrp.GroupID = CLng(.Text)
    cgrp.Description = TellCtrlGrp(eGrpBANK)
    cgrp.Save True
  End If
Else
    cgrp.Init TellCtrlGrp(eGrpBANK)
    If cgrp.IsValid Then cgrp.Remove
End If
End With

With txtPar
If .Text <> "" Then
  cgrp.Init TellCtrlGrp(eGrpPARTIES)
  If cgrp.IsValid Then
    cgrp.GroupID = CLng(.Text)
    cgrp.Save False
  Else
    cgrp.GroupID = CLng(.Text)
    cgrp.Description = TellCtrlGrp(eGrpPARTIES)
    cgrp.Save True
  End If
Else
  cgrp.Init TellCtrlGrp(eGrpPARTIES)
  cgrp.Remove
End If
End With

With txtCust
If .Text <> "" Then
  cgrp.Init TellCtrlGrp(eGrpCUSTOMERS)
  If cgrp.IsValid Then
    cgrp.GroupID = CLng(.Text)
    cgrp.Save False
  Else
    cgrp.GroupID = CLng(.Text)
    cgrp.Description = TellCtrlGrp(eGrpCUSTOMERS)
    cgrp.Save True
  End If
Else
    cgrp.Init TellCtrlGrp(eGrpCUSTOMERS)
    cgrp.Remove
End If
End With

With txtSUP
If .Text <> "" Then
  cgrp.Init TellCtrlGrp(eGrpSUPPLIERS)
  If cgrp.IsValid Then
    cgrp.GroupID = CLng(.Text)
    cgrp.Save False
  Else
    cgrp.GroupID = CLng(.Text)
    cgrp.Description = TellCtrlGrp(eGrpSUPPLIERS)
    cgrp.Save True
  End If
Else
    cgrp.Init TellCtrlGrp(eGrpSUPPLIERS)
    cgrp.Remove
End If
End With

With txtSL
If .Text <> "" Then
  cgrp.Init TellCtrlGrp(eGrpSALES)
  If cgrp.IsValid Then
    cgrp.GroupID = CLng(.Text)
    cgrp.Save False
  Else
    cgrp.GroupID = CLng(.Text)
    cgrp.Description = TellCtrlGrp(eGrpSALES)
    cgrp.Save True
  End If
Else
  cgrp.Init TellCtrlGrp(eGrpSALES)
  cgrp.Remove
End If
End With

With txtPur
If .Text <> "" Then
  cgrp.Init TellCtrlGrp(eGrpPURCHASES)
  If cgrp.IsValid Then
    cgrp.GroupID = CLng(.Text)
    cgrp.Save False
  Else
    cgrp.GroupID = CLng(.Text)
    cgrp.Description = TellCtrlGrp(eGrpPURCHASES)
    cgrp.Save True
  End If
Else
  cgrp.Init TellCtrlGrp(eGrpPURCHASES)
  cgrp.Remove
End If
End With

With txtExp
If .Text <> "" Then
  cgrp.Init TellCtrlGrp(eGrpEXPENSE)
  If cgrp.IsValid Then
    cgrp.GroupID = CLng(.Text)
    cgrp.Save False
  Else
    cgrp.GroupID = CLng(.Text)
    cgrp.Description = TellCtrlGrp(eGrpEXPENSE)
    cgrp.Save True
  End If
Else
  cgrp.Init TellCtrlGrp(eGrpEXPENSE)
  cgrp.Remove
End If
End With

With txtLcAccts
If .Text <> "" Then
  cgrp.Init TellCtrlGrp(eGrpLCACCTS)
  If cgrp.IsValid Then
    cgrp.GroupID = CLng(.Text)
    cgrp.Save False
  Else
    cgrp.GroupID = CLng(.Text)
    cgrp.Description = TellCtrlGrp(eGrpLCACCTS)
    cgrp.Save True
  End If
Else
  cgrp.Init TellCtrlGrp(eGrpLCACCTS)
  cgrp.Remove
End If
End With

With txtCash
If .Text <> "" Then
  cgrp.Init TellCtrlGrp(egrpCASH)
  If cgrp.IsValid Then
    cgrp.GroupID = CLng(.Text)
    cgrp.Save False
  Else
    cgrp.GroupID = CLng(.Text)
    cgrp.Description = TellCtrlGrp(egrpCASH)
    cgrp.Save True
  End If
Else
  cgrp.Init TellCtrlGrp(egrpCASH)
  cgrp.Remove
End If
End With

With txtCash
If .Text <> "" Then
  cgrp.Init TellCtrlGrp(egrpCASH)
  If cgrp.IsValid Then
    cgrp.GroupID = CLng(NumVal(.Text))
    cgrp.Save False
  Else
    cgrp.GroupID = CLng(.Text)
    cgrp.Description = TellCtrlGrp(egrpCASH)
    cgrp.Save True
  End If
Else
    cgrp.Init TellCtrlGrp(egrpCASH)
    If cgrp.IsValid Then cgrp.Remove
End If
End With
End Sub


Sub OpenAll()
Dim ccgrp As New CControlGroup
Dim cgrp As New CGroup

ccgrp.Init TellCtrlGrp(eGrpBANK)
If ccgrp.IsValid Then
  cgrp.Init ccgrp.GroupID
  If cgrp.IsValid Then
    txtCAB = cgrp.GroupID
    txtCABDesc = cgrp.Description
    txtCAB.DataChanged = False
  End If
End If

ccgrp.Init TellCtrlGrp(eGrpPARTIES)
If ccgrp.IsValid Then
  cgrp.Init ccgrp.GroupID
  If cgrp.IsValid Then
    txtPar = cgrp.GroupID
    txtParDesc = cgrp.Description
    txtPar.DataChanged = False
  End If
End If

ccgrp.Init TellCtrlGrp(eGrpCUSTOMERS)
If ccgrp.IsValid Then
  cgrp.Init ccgrp.GroupID
  If cgrp.IsValid Then
    txtCust = cgrp.GroupID
    txtCustDesc = cgrp.Description
    txtCust.DataChanged = False
  End If
End If

ccgrp.Init TellCtrlGrp(eGrpSUPPLIERS)
If ccgrp.IsValid Then
  cgrp.Init ccgrp.GroupID
  If cgrp.IsValid Then
    txtSUP = cgrp.GroupID
    txtSUPDesc = cgrp.Description
    txtSUP.DataChanged = False
  End If
End If

ccgrp.Init TellCtrlGrp(eGrpSALES)
If ccgrp.IsValid Then
  cgrp.Init ccgrp.GroupID
  If cgrp.IsValid Then
    txtSL = cgrp.GroupID
    txtSLDesc = cgrp.Description
    txtSL.DataChanged = False
  End If
End If

ccgrp.Init TellCtrlGrp(eGrpPURCHASES)
If ccgrp.IsValid Then
  cgrp.Init ccgrp.GroupID
  If cgrp.IsValid Then
    txtPur = cgrp.GroupID
    txtPurDesc = cgrp.Description
    txtPur.DataChanged = False
  End If
End If

ccgrp.Init TellCtrlGrp(eGrpEXPENSE)
If ccgrp.IsValid Then
  cgrp.Init ccgrp.GroupID
  If cgrp.IsValid Then
    txtExp = cgrp.GroupID
    txtExpDesc = cgrp.Description
    txtExp.DataChanged = False
  End If
End If

ccgrp.Init TellCtrlGrp(eGrpLCACCTS)
If ccgrp.IsValid Then
  cgrp.Init ccgrp.GroupID
  If cgrp.IsValid Then
    txtLcAccts = cgrp.GroupID
    txtLcAcctsDesc = cgrp.Description
    txtLcAccts.DataChanged = False
  End If
End If

ccgrp.Init TellCtrlGrp(egrpCASH)
If ccgrp.IsValid Then
  cgrp.Init ccgrp.GroupID
  If cgrp.IsValid Then
    txtCash = cgrp.GroupID
    txtCashDesc = cgrp.Description
    txtCash.DataChanged = False
  End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
  SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Load()
OpenAll
frmMain.SetTlbLayout -1
End Sub

Sub CheckIfValid(ByRef txtBx As TextBox, ByRef CancelFocusShift As Boolean, txtBxDesc As TextBox)
If txtBx.Text = "" Then Exit Sub
Dim cG As New CGroup
cG.Init CLng(NumVal(txtBx.Text))
If cG.IsValid Then
  txtBxDesc.Text = cG.Description
Else
  MsgBox "Group ID " & txtBx & " is not valid.", vbInformation
  txtBxDesc.Text = ""
  CancelFocusShift = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.SetTlbDefLayout
End Sub

Private Sub txtCAB_GotFocus()
HighlightText txtCAB
End Sub

Private Sub txtCAB_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtCAB_Validate(Cancel As Boolean)
CheckIfValid txtCAB, Cancel, txtCABDesc
End Sub

Private Sub txtCash_GotFocus()
HighlightText txtCash
End Sub

Private Sub txtCash_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtCash_Validate(Cancel As Boolean)
CheckIfValid txtCash, Cancel, txtCashDesc
End Sub

Private Sub txtCust_GotFocus()
HighlightText txtCust
End Sub

Private Sub txtCust_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtCust_Validate(Cancel As Boolean)
CheckIfValid txtCust, Cancel, txtCustDesc
End Sub

Private Sub txtExp_GotFocus()
HighlightText txtExp
End Sub

Private Sub txtExp_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtExp_Validate(Cancel As Boolean)
CheckIfValid txtExp, Cancel, txtExpDesc
End Sub

Private Sub txtPar_GotFocus()
HighlightText txtPar
End Sub

Private Sub txtPar_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtPar_Validate(Cancel As Boolean)
CheckIfValid txtPar, Cancel, txtParDesc
End Sub

Private Sub txtLcAccts_GotFocus()
HighlightText txtLcAccts
End Sub

Private Sub txtLcAccts_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtLcAccts_Validate(Cancel As Boolean)
CheckIfValid txtLcAccts, Cancel, txtLcAcctsDesc
End Sub

Private Sub txtPur_GotFocus()
HighlightText txtPur
End Sub

Private Sub txtPur_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtPur_Validate(Cancel As Boolean)
CheckIfValid txtPur, Cancel, txtPurDesc
End Sub

Private Sub txtSL_GotFocus()
HighlightText txtSL
End Sub

Private Sub txtSL_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtSL_Validate(Cancel As Boolean)
CheckIfValid txtSL, Cancel, txtSLDesc
End Sub

Private Sub txtSUP_GotFocus()
HighlightText txtSUP
End Sub

Private Sub txtSUP_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtSUP_Validate(Cancel As Boolean)
CheckIfValid txtSUP, Cancel, txtSUPDesc
End Sub
