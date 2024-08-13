VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCashBook 
   Caption         =   "Cash Book"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10350
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   10350
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport crstRpt 
      Left            =   1770
      Top             =   990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picUp2 
      Align           =   1  'Align Top
      HasDC           =   0   'False
      Height          =   450
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   10290
      TabIndex        =   12
      Top             =   450
      Width           =   10350
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Receipts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   30
         TabIndex        =   14
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Payments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   5235
         TabIndex        =   13
         Top             =   45
         Width           =   1185
      End
   End
   Begin VB.PictureBox picUp 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   10350
      TabIndex        =   11
      Top             =   0
      Width           =   10350
      Begin VB.TextBox txtOpening 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8415
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   45
         Width           =   1860
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   555
         TabIndex        =   0
         Top             =   60
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   47054849
         CurrentDate     =   37723
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   17
         Top             =   105
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   6900
         TabIndex        =   16
         Top             =   120
         Width           =   1470
      End
      Begin VB.Label lblCashBook 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "DAY BOOK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   0
         TabIndex        =   15
         Top             =   30
         Width           =   10245
      End
   End
   Begin VB.PictureBox picDown 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   10350
      TabIndex        =   10
      Top             =   4764
      Width           =   10350
      Begin VB.TextBox txtClosing 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   420
         Width           =   1860
      End
      Begin VB.TextBox txtReceipts 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   30
         Width           =   1860
      End
      Begin VB.TextBox txtPayments 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8385
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   30
         Width           =   1860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Closing Balance"
         Height          =   195
         Index           =   2
         Left            =   7125
         TabIndex        =   8
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Receipts"
         Height          =   195
         Index           =   3
         Left            =   2250
         TabIndex        =   4
         Top             =   90
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Payment"
         Height          =   195
         Index           =   4
         Left            =   7260
         TabIndex        =   6
         Top             =   90
         Width           =   1020
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgrdReceipts 
      Height          =   3255
      Left            =   60
      TabIndex        =   2
      Top             =   1335
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "TransNo|AccountID|Account Name|Descripiton|Amount"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid fgrdPayments 
      Height          =   3255
      Left            =   5235
      TabIndex        =   3
      Top             =   1335
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuNewRecipt 
         Caption         =   "New Recipt"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuNewPayment 
         Caption         =   "New Payment"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuSep_a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditReceipt 
         Caption         =   "Edit Receipt"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEditPayment 
         Caption         =   "Edit Payment"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSep_b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteReceipt 
         Caption         =   "Delete Receipt"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDeletePayment 
         Caption         =   "Delete Payment"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuDash_Q 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReport 
         Caption         =   "Report"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDash_P 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnuReceipts 
      Caption         =   "Receipts"
      Visible         =   0   'False
      Begin VB.Menu mnuContextNewReceipt 
         Caption         =   "New Receipt"
      End
      Begin VB.Menu mnuContextEditReceipt 
         Caption         =   "Edit Receipt"
      End
      Begin VB.Menu mnuContextDeleteReceipt 
         Caption         =   "Delete Receipt"
      End
   End
   Begin VB.Menu mnuPayments 
      Caption         =   "Payments"
      Visible         =   0   'False
      Begin VB.Menu mnuContextNewPayment 
         Caption         =   "New Payment"
      End
      Begin VB.Menu mnuContextEditPayment 
         Caption         =   "Edit Payment"
      End
      Begin VB.Menu mnuContextDeletePayment 
         Caption         =   "Delete Payment"
      End
   End
   Begin VB.Menu mnuOpeningBalance 
      Caption         =   "&OpeningBalance"
   End
   Begin VB.Menu mnuJournalVoucher 
      Caption         =   "&JournalVoucher"
   End
   Begin VB.Menu mnuCustomers 
      Caption         =   "A&ccounts"
   End
End
Attribute VB_Name = "frmCashBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_cTransDetCol2 As CTransDetCol2
Dim m_cTrans As CTrans
Dim m_cTransDets As CTransDets
Dim m_cTransDet As CTransDet
Dim m_cControlAccounts As ControlAccounts
Dim m_CIDs As New CIDs

Private Sub dtpDate_Change()
    Dim cAccts As New cAccounts
    txtOpening.Text = cAccts.Balance(m_cControlAccounts.AccountNo(CashInHand), dtpDate.value - 1)
    Set cAccts = Nothing
    PopGrids
End Sub

Private Sub fgrdPayments_KeyUp(KeyCode As Integer, Shift As Integer)
fgrdPayments.SetFocus
   If KeyCode = 93 And Shift = 0 Then
      PopupMenu mnuPayments, , , , mnuContextNewPayment
   End If
End Sub

Private Sub fgrdPayments_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
fgrdPayments.SetFocus
   If Button = vbRightButton And Shift = 0 Then
      PopupMenu mnuPayments, , , , mnuContextNewPayment
   End If
End Sub

Private Sub fgrdReceipts_KeyUp(KeyCode As Integer, Shift As Integer)
fgrdReceipts.SetFocus
   If KeyCode = 93 And Shift = 0 Then
      PopupMenu mnuReceipts, , , , mnuContextNewReceipt
   End If
End Sub

Private Sub fgrdReceipts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
fgrdReceipts.SetFocus
   If Button = vbRightButton And Shift = 0 Then
      PopupMenu mnuReceipts, , , , mnuContextNewReceipt
   End If
End Sub

Private Sub Form_Load()
Set m_cControlAccounts = New ControlAccounts
m_cControlAccounts.Initialize
Dim lCW As Long
With fgrdReceipts
   .TextMatrix(0, 0) = "Trans No"
   .TextMatrix(0, 1) = "Ac. No."
   .TextMatrix(0, 2) = "Account Name"
   .TextMatrix(0, 3) = "Description"
   .TextMatrix(0, 4) = "Amount"
   .ColWidth(0) = TextWidth(String(9, "8"))
   .ColWidth(1) = TextWidth(String(6, "A"))
   .ColWidth(2) = TextWidth(String(40, " "))
   .ColWidth(3) = TextWidth(String(39, " "))
   .ColWidth(4) = TextWidth(String(9, "8"))
   .ColWidth(5) = 0
   .ColAlignment(3) = vbLeftJustify
   .Rows = 1
End With
With fgrdPayments
   .TextMatrix(0, 0) = "Trans No"
   .TextMatrix(0, 1) = "Ac. No."
   .TextMatrix(0, 2) = "Account Name"
   .TextMatrix(0, 3) = "Description"
   .TextMatrix(0, 4) = "Amount"
   .ColWidth(0) = TextWidth(String(9, "8"))
   .ColWidth(1) = TextWidth(String(6, "A"))
   .ColWidth(2) = TextWidth(String(40, " "))
   .ColWidth(3) = TextWidth(String(38, " "))
   .ColWidth(4) = TextWidth(String(9, "8"))
   .ColWidth(5) = 0
   .ColAlignment(3) = vbLeftJustify
   .Rows = 1
End With
dtpDate.value = Date
dtpDate_Change
frmMain.SetTlbLayout 2
End Sub

Sub NewReceipt()
If Not HasRights(soCashBook, CanAdd) Then
  MsgBox "You do not have required privileges to proceed !", vbExclamation
  Exit Sub
End If
Dim AccountID As Long
Dim AccountName As String
Dim Description As String
Dim Amount As Double
AccountID = 0
AccountName = ""
Description = ""
Amount = 0
If frmCash.ShowForm(AccountID, AccountName, Description, Amount, "N", "New Cash Receipt Entry") Then
   Dim r As VbMsgBoxResult
   r = MsgBox("Do you want to save ? ", vbYesNo Or vbQuestion Or vbDefaultButton1, "Save")
   If r = vbYes Then
   
      With fgrdReceipts
         .Rows = .Rows + 1
         .Row = .Rows - 1

         m_CIDs.NewID ("CR")
         .TextMatrix(.Row, 0) = m_CIDs.ID & "-CR"
         .TextMatrix(.Row, 1) = AccountID
         .TextMatrix(.Row, 2) = AccountName
         .TextMatrix(.Row, 3) = Description
         .TextMatrix(.Row, 4) = Amount
        .TextMatrix(.Row, 5) = "CR"
      End With
      Set m_cTrans = New CTrans
      Set m_cTransDet = New CTransDet
      Set m_cTransDets = New CTransDets
      'Trans
      m_cTrans.ID = m_CIDs.ID
      m_cTrans.TransType = "CR"
      m_cTrans.TDate = dtpDate.value
      m_cTrans.Save True
      'TransDet
      'Credit Side Entry
      m_cTransDet.ID = m_CIDs.ID
      m_cTransDet.TransType = "CR"
      m_cTransDet.AccountID = AccountID
      m_cTransDet.Description = Description
      m_cTransDet.Credit = Amount
      m_cTransDets.AddC m_cTransDet, 1
      'Debit Side Entry
      Set m_cTransDet = New CTransDet
      m_cTransDet.ID = m_CIDs.ID
      m_cTransDet.TransType = "CR"
      m_cTransDet.AccountID = m_cControlAccounts.AccountNo(CashInHand)     'Cash In Hand Account
      m_cTransDet.Description = Description
      m_cTransDet.Debit = Amount
      m_cTransDets.AddC m_cTransDet, 2
      m_cTransDets.Save True, m_CIDs.ID
   End If
End If
TellTotals
End Sub
Sub EditReceipt()
If Not HasRights(soCashBook, CanEdit) Then
  MsgBox "You do not have required privileges to proceed !", vbExclamation
  Exit Sub
End If
Dim AccountID As Long
Dim RefAccountID As Long
Dim AccountName As String
Dim Description As String
Dim Amount As Double
With fgrdReceipts
    If .Rows <= 1 Then
      MsgBox "No receipt entry to edit.", vbInformation
      Exit Sub
    End If
    .RowSel = .Row
    .Col = .Cols - 1
    .ColSel = 0
    If StrComp(.TextMatrix(.Row, 5), "CR") <> 0 Then
        MsgBox "Only Cash Receipt can be edited.", vbExclamation
        Exit Sub
    End If
    AccountID = CLng(IIf(Len(.TextMatrix(.Row, 1)) = 0, 0, .TextMatrix(.Row, 1)))
    AccountName = .TextMatrix(.Row, 2)
    Description = .TextMatrix(.Row, 3)
    Amount = CDbl(IIf(Len(.TextMatrix(.Row, 4)) = 0, 0, .TextMatrix(.Row, 4)))
    RefAccountID = AccountID
End With

If frmCash.ShowForm(AccountID, AccountName, Description, Amount, "E", "Edit Cash Receipt Entry") Then
   With fgrdReceipts
      .TextMatrix(.Row, 1) = AccountID
      .TextMatrix(.Row, 2) = AccountName
      .TextMatrix(.Row, 3) = Description
      .TextMatrix(.Row, 4) = Amount
       Set m_cTransDets = New CTransDets
       m_cTransDets.Init Val(.TextMatrix(.Row, 0)), "CR"
       Dim a As Integer
       For a = 1 To m_cTransDets.Count
         If m_cTransDets.Item(a).AccountID <> m_cControlAccounts.AccountNo(CashInHand) Then
            m_cTransDets.Item(a).AccountID = AccountID
            m_cTransDets.Item(a).Description = Description
            m_cTransDets.Item(a).Credit = Amount
        Else
          m_cTransDets.Item(a).Debit = Amount
         End If
       Next
       m_cTransDets.Save False, Val(.TextMatrix(.Row, 0))
   End With
End If
TellTotals
End Sub

Sub NewPayment()
If Not HasRights(soCashBook, CanAdd) Then
  MsgBox "You do not have required privileges to proceed !", vbExclamation
  Exit Sub
End If
Dim AccountID As Long
Dim AccountName As String
Dim Description As String
Dim Amount As Double
AccountID = 0
AccountName = ""
Description = ""
Amount = 0
If frmCash.ShowForm(AccountID, AccountName, Description, Amount, "N", "New Cash Payment Entry") Then
   Dim r As VbMsgBoxResult
   r = MsgBox("Do you want to save ? ", vbYesNo Or vbQuestion Or vbDefaultButton1, "Save")
   If r = vbYes Then
      With fgrdPayments
         .Rows = .Rows + 1
         .Row = .Rows - 1

         m_CIDs.NewID ("CP")
         .TextMatrix(.Row, 0) = m_CIDs.ID & "-CP"
         .TextMatrix(.Row, 1) = AccountID
         .TextMatrix(.Row, 2) = AccountName
         .TextMatrix(.Row, 3) = Description
         .TextMatrix(.Row, 4) = Amount
        .TextMatrix(.Row, 5) = "CP"
      End With
      Set m_cTrans = New CTrans
      Set m_cTransDet = New CTransDet
      Set m_cTransDets = New CTransDets
      'Trans
      m_cTrans.ID = m_CIDs.ID
      m_cTrans.TransType = "CP"
      m_cTrans.TDate = dtpDate.value
      m_cTrans.Save True
      'TransDet
      'Credit Side Entry
      m_cTransDet.ID = m_CIDs.ID
      m_cTransDet.TransType = "CP"
      m_cTransDet.AccountID = AccountID
      m_cTransDet.Description = Description
      m_cTransDet.Debit = Amount
      m_cTransDets.AddC m_cTransDet, 1
      'Debit Side Entry
      Set m_cTransDet = New CTransDet
      m_cTransDet.ID = m_CIDs.ID
      m_cTransDet.TransType = "CP"
      m_cTransDet.AccountID = m_cControlAccounts.AccountNo(CashInHand)     'Cash In Hand Account
      m_cTransDet.Description = Description
      m_cTransDet.Credit = Amount
      m_cTransDets.AddC m_cTransDet, 2
      m_cTransDets.Save True, m_CIDs.ID
   End If
End If
TellTotals
End Sub

Sub EditPayment()
If Not HasRights(soCashBook, CanEdit) Then
  MsgBox "You do not have required privileges to proceed !", vbExclamation
  Exit Sub
End If
Dim AccountID As Long
Dim RefAccountID As Long
Dim AccountName As String
Dim Description As String
Dim Amount As Double
With fgrdPayments
    If .Rows <= 1 Then
      MsgBox "No payment entry to edit.", vbInformation
      Exit Sub
    End If
    .RowSel = .Row
    .Col = .Cols - 1
    .ColSel = 0
    If StrComp(.TextMatrix(.Row, 5), "CP") <> 0 Then
        MsgBox "Only Cash Payment can be edited.", vbExclamation
        Exit Sub
    End If
    AccountID = CLng(IIf(Len(.TextMatrix(.Row, 1)) = 0, 0, .TextMatrix(.Row, 1)))
    AccountName = .TextMatrix(.Row, 2)
    Description = .TextMatrix(.Row, 3)
    Amount = CDbl(IIf(Len(.TextMatrix(.Row, 4)) = 0, 0, .TextMatrix(.Row, 4)))
    RefAccountID = AccountID
End With

If frmCash.ShowForm(AccountID, AccountName, Description, Amount, "E", "Edit Cash Payment Entry") Then
   With fgrdPayments
      .TextMatrix(.Row, 1) = AccountID
      .TextMatrix(.Row, 2) = AccountName
      .TextMatrix(.Row, 3) = Description
      .TextMatrix(.Row, 4) = Amount
       Set m_cTransDets = New CTransDets
       m_cTransDets.Init Val(.TextMatrix(.Row, 0)), "CP"
       Dim a As Integer
       For a = 1 To m_cTransDets.Count
         If m_cTransDets.Item(a).AccountID <> m_cControlAccounts.AccountNo(CashInHand) Then
            m_cTransDets.Item(a).AccountID = AccountID
            m_cTransDets.Item(a).Description = Description
            m_cTransDets.Item(a).Debit = Amount
        Else
          m_cTransDets.Item(a).Credit = Amount
         End If
       Next
       m_cTransDets.Save False, Val(.TextMatrix(.Row, 0))
   End With
End If
TellTotals
End Sub

Private Sub Form_Resize()
SetFormLayout
End Sub

Private Sub mnuEdit_Click()
If Me.ActiveControl Is fgrdReceipts Then
   EditReceipt
ElseIf Me.ActiveControl Is fgrdPayments Then
   EditPayment
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.SetTlbDefLayout
End Sub

Private Sub mnuContextDeletePayment_Click()
DeletePayment
End Sub

Private Sub mnuContextDeleteReceipt_Click()
DeleteReceipt
End Sub

Private Sub mnuContextEditPayment_Click()
mnuEditPayment_Click
End Sub

Private Sub mnuContextEditReceipt_Click()
mnuEditReceipt_Click
End Sub

Private Sub mnuContextNewPayment_Click()
mnuNewPayment_Click
End Sub

Private Sub mnuContextNewReceipt_Click()
mnuNewRecipt_Click
End Sub

Private Sub mnuCustomers_Click()
If Not HasRights(soChartOfAccounts, CanView) Then
    MsgBox "You do not have required privileges to proceed !", vbExclamation
    Exit Sub
End If
    frmAccounts2.Show
    frmAccounts2.SetFocus
End Sub

Private Sub mnuDeletePayment_Click()
DeletePayment
End Sub

Private Sub mnuDeleteReceipt_Click()
DeleteReceipt
End Sub

Private Sub mnuEditPayment_Click()
EditPayment
End Sub

Private Sub mnuEditReceipt_Click()
EditReceipt
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuNew_Click()
If Me.ActiveControl Is fgrdReceipts Then
   NewReceipt
ElseIf Me.ActiveControl Is fgrdPayments Then
   NewPayment
End If
End Sub

Private Sub SetFormLayout()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
If Me.WindowState <> vbMaximized Then
   If Me.Width < Screen.Width * (2 / 3) Then
      Me.Width = Screen.Width * (2 / 3)
   End If
   If Me.Height < Screen.Height / 2 Then
      Me.Height = Screen.Height / 2
   End If
End If
fgrdReceipts.Top = picUp.Height + picUp2.Height
fgrdReceipts.Height = Me.Height - (picUp.Height + picDown.Height + 700 + picUp2.Height)
fgrdPayments.Top = fgrdReceipts.Top
fgrdPayments.Height = fgrdReceipts.Height
fgrdReceipts.Width = Me.Width / 2 - 100
fgrdPayments.Width = Me.Width / 2 - 150
fgrdReceipts.Left = Me.ScaleLeft + 50
fgrdPayments.Left = Me.ScaleLeft + Me.Width / 2
Label2(0).Left = fgrdReceipts.Left + 10
Label2(1).Left = fgrdPayments.Left + 10
txtOpening.Left = (fgrdPayments.Left + fgrdPayments.Width) - (txtOpening.Width + 50)
txtClosing.Left = txtOpening.Left
txtPayments.Left = txtOpening.Left
lblCashBook.Width = Me.ScaleWidth
lblCashBook.Left = (Me.Width - lblCashBook.Width) / 2

With Label1(1)
   .Left = txtOpening.Left - (.Width + 100)
End With
With Label1(2)
   .Left = txtOpening.Left - (.Width + 100)
End With
With Label1(4)
   .Left = txtOpening.Left - (.Width + 100)
End With

txtReceipts.Left = (fgrdReceipts.Left + fgrdReceipts.Width) - (txtReceipts.Width + 25)
With Label1(3)
   .Left = txtReceipts.Left - (.Width + 100)
End With
End Sub

Sub PopGrids()
Dim rscAccounts As New cAccounts
Set m_cTransDetCol2 = New CTransDetCol2
fgrdReceipts.Rows = 1
fgrdPayments.Rows = 1
With m_cTransDetCol2
   .Init dtpDate.value
   Dim a As Integer
   For a = 1 To .Count
     If m_cTransDetCol2.Item(a).AccountID <> m_cControlAccounts.AccountNo(CashInHand) Then
        If m_cTransDetCol2.Item(a).Credit > 0 Then 'StrComp(m_cTransDetCol2.Item(a).TransType, "CR", vbTextCompare) = 0 Then
         With fgrdReceipts
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Row, 0) = m_cTransDetCol2.Item(a).ID & "-" & m_cTransDetCol2.Item(a).TransType
            .TextMatrix(.Row, 1) = m_cTransDetCol2.Item(a).AccountID
            .TextMatrix(.Row, 2) = rscAccounts.GetAccountName(m_cTransDetCol2.Item(a).AccountID)   'm_cTransDetCol2.Item(a).Description 'Account Name
            .TextMatrix(.Row, 3) = m_cTransDetCol2.Item(a).Description
            .TextMatrix(.Row, 4) = m_cTransDetCol2.Item(a).Credit
            .TextMatrix(.Row, 5) = m_cTransDetCol2.Item(a).TransType
         End With
         ElseIf m_cTransDetCol2.Item(a).Debit > 0 Then 'StrComp(m_cTransDetCol2.Item(a).TransType, "CP", vbTextCompare) = 0 Then
         With fgrdPayments
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Row, 0) = m_cTransDetCol2.Item(a).ID & "-" & m_cTransDetCol2.Item(a).TransType
            .TextMatrix(.Row, 1) = m_cTransDetCol2.Item(a).AccountID
            .TextMatrix(.Row, 2) = rscAccounts.GetAccountName(m_cTransDetCol2.Item(a).AccountID)   'm_cTransDetCol2.Item(a).Description 'Account Name
            .TextMatrix(.Row, 3) = m_cTransDetCol2.Item(a).Description
            .TextMatrix(.Row, 4) = m_cTransDetCol2.Item(a).Debit
            .TextMatrix(.Row, 5) = m_cTransDetCol2.Item(a).TransType
         End With
         End If
      End If
   Next
End With
TellTotals
End Sub

Sub TellTotals()
Dim I As Integer, dV As Double
dV = 0
For I = 1 To fgrdReceipts.Rows
   dV = dV + Val(fgrdReceipts.TextMatrix(I - 1, 4))
Next
txtReceipts.Text = dV
dV = 0
For I = 1 To fgrdPayments.Rows
   dV = dV + Val(fgrdPayments.TextMatrix(I - 1, 4))
Next
txtPayments.Text = dV
txtClosing = Val(txtOpening) + Val(txtReceipts) - Val(txtPayments)
End Sub

Private Sub mnuJournalVoucher_Click()
Static JV As New frmJournalVoucher
Unload JV
DoEvents
If Not HasRights(soJournalVoucher, CanView) Then
    MsgBox "You do not have required privileges to proceed !", vbExclamation
    Exit Sub
End If
    Set JV.Icon = frmMain.imgShowForm.ListImages.Item(8).Picture
'    JV.Top = 1
    JV.Show
    DoEvents
End Sub

Private Sub mnuNewPayment_Click()
NewPayment
End Sub

Private Sub mnuNewRecipt_Click()
NewReceipt
End Sub

Sub DeleteReceipt()
If Not HasRights(soCashBook, CanDelete) Then
    MsgBox "You do not have required privileges to proceed !", vbExclamation
    Exit Sub
End If
frmDeleteTrans.ShowForm "CR", vbModal, frmMain
End Sub
Sub DeletePayment()
If Not HasRights(soCashBook, CanDelete) Then
    MsgBox "You do not have required privileges to proceed !", vbExclamation
    Exit Sub
End If
frmDeleteTrans.ShowForm "CP", vbModal, frmMain
End Sub

Sub ShowReport()
    Dim dOB As Double
    Dim dtDate As Date
    dtDate = dtpDate.value
    With crstRpt
        .Reset
        .Connect = m_objConnectDB.cnnMyshop.ConnectionString
        .ReportFileName = App.Path & "\reports\DayBook.rpt"
        If Dir(.ReportFileName) = "" Then
            MsgBox "The Report file is missing.", vbExclamation
            Exit Sub
        End If
        ' Get External Values
        Dim cCA As New ControlAccounts
        Dim cAccts As New cAccounts
        cCA.Initialize
        dOB = cAccts.Balance(cCA.AccountNo(CashInHand), dtDate - 1)
        cCA.Purge
        Set cAccts = Nothing
        ' Set Up Report
        .ParameterFields(0) = "pmOpeningBalance;" & dOB & ";true"
        .ParameterFields(1) = "pmCompany;" & psCompanyName & ";true"
        .ParameterFields(2) = "pmCashInHand;" & SelAcntIDs.CashInHand & ";true"
        .StoredProcParam(0) = Format(dtDate, "mm/dd/yyyy")
       '.ParameterFields(3) = "sDate;" & Format(dtpDate.value, "mm/dd/yyyy") & ";true"
        .WindowState = crptMaximized
        .ReportTitle = "Day Book"
        .WindowTitle = .ReportTitle
        .Destination = crptToWindow
        .Action = 1
    End With
End Sub

Private Sub mnuOpeningBalance_Click()
Static OB As New frmOpeningBalance
Unload OB
If Not HasRights(soOpeningBalance, CanView) Then
    MsgBox "You do not have required privileges to proceed !", vbExclamation
    Exit Sub
End If
    OB.Show
    DoEvents
End Sub

Private Sub mnuReport_Click()
ShowReport
End Sub
