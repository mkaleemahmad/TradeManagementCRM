VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCustomerVoucher 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Voucher"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6840
   ClipControls    =   0   'False
   Icon            =   "frmCustomerVoucher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTotalCredit 
      Height          =   315
      Left            =   5393
      TabIndex        =   16
      Text            =   "0"
      Top             =   5310
      Width           =   1335
   End
   Begin VB.TextBox txtTotalDebit 
      Height          =   315
      Left            =   3953
      TabIndex        =   15
      Text            =   "0"
      Top             =   5310
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transaction"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   113
      TabIndex        =   9
      Top             =   600
      Width           =   6615
      Begin MSMask.MaskEdBox txtCredit 
         Height          =   285
         Left            =   4560
         TabIndex        =   4
         Top             =   615
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##########"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDebit 
         Height          =   285
         Left            =   3240
         TabIndex        =   3
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##########"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDescription 
         Height          =   285
         Left            =   255
         TabIndex        =   2
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         MaxLength       =   50
         Mask            =   "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdDel 
         Height          =   375
         Left            =   5880
         Picture         =   "frmCustomerVoucher.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   375
         Left            =   5880
         Picture         =   "frmCustomerVoucher.frx":0996
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Debit"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   11
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   900
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgridCTrans 
      Height          =   3135
      Left            =   128
      TabIndex        =   6
      Top             =   1920
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   5
      SelectionMode   =   1
   End
   Begin VB.ComboBox cmbCustomer 
      Height          =   315
      Left            =   3128
      TabIndex        =   1
      Top             =   120
      Width           =   3600
   End
   Begin MSComCtl2.DTPicker dtpTransDate 
      Height          =   315
      Left            =   728
      TabIndex        =   0
      Top             =   120
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58589185
      CurrentDate     =   37414
   End
   Begin VB.Label lblCustomerBalance 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   5415
      Width           =   2535
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Account Balance"
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
      Left            =   480
      TabIndex        =   17
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Totals:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   5340
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Account"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2400
      TabIndex        =   8
      Top             =   156
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   233
      TabIndex        =   7
      Top             =   150
      Width           =   375
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit Customer Voucher"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmCustomerVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_sTransType As String
Dim m_iGridCount As Integer
Dim m_bIsNewRecord As Boolean
Dim m_lTransID As Long
Dim m_bDataChanged As Boolean
Dim m_objConnectDB As ConnectDB
Dim m_objTransDets As CTransDets
Dim m_objTrans As CTrans
Dim m_objTransDet As CTransDet
Dim m_objCIDs As CIDs
'

Private Sub SetNew()

    'txtDescription = ""
    txtDebit = 0
    txtCredit = 0
    txtTotalDebit = 0
    txtTotalCredit = 0
    cmbCustomer.ListIndex = -1
    fgridCTrans.Clear
    fgridCTrans.Rows = 2
    SetGridTitles
    m_bIsNewRecord = True
    m_iGridCount = 1
    dtpTransDate.Value = Date
End Sub

Sub SetGridTitles()
  ' Grid Titles
    fgridCTrans.Row = 0
  ' fgridTrans.ColWidth(0) = 1000
    fgridCTrans.ColWidth(0) = 200
    fgridCTrans.ColWidth(1) = 3500
    fgridCTrans.ColWidth(2) = 1400
    fgridCTrans.ColWidth(3) = 1400
    fgridCTrans.ColWidth(4) = 0
    fgridCTrans.Col = 1
    fgridCTrans.Text = "Description"
    fgridCTrans.Col = 2
    fgridCTrans.Text = "Debit"
    fgridCTrans.Col = 3
    fgridCTrans.Text = "Credit"
        
End Sub

'Private Sub cmbCustomer_Click()
'    If cmbCustomer.ListIndex > -1 Then
'        m_objTransDet.Customer.Init cmbCustomer.ItemData(cmbCustomer.ListIndex)
'        lblCustomerBalance.Caption = str(m_objTransDet.Customer.Balance) + " " + m_objTransDet.Customer.BalanceType
'    End If
'End Sub

Private Sub cmbCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode, False
End Sub

Private Sub cmdAdd_Click()
    If cmbCustomer.ListIndex = -1 Then
        MsgBox "You must select a Customer to make a transaction", vbCritical + vbOKOnly
    ElseIf (txtDebit = 0 And txtCredit = 0) Then
        MsgBox "Debit or Credit amount must be entered to make a transaction", vbCritical + vbOKOnly
    Else
            
        Set m_objTransDet = New CTransDet
'        cmbCustomer_Click
        m_objTransDet.Description = txtDescription
        m_objTransDet.Debit = txtDebit
        m_objTransDet.Credit = txtCredit
        m_objTransDet.TransType = m_sTransType
        
        UpdateGrid
        
        m_objTransDets.AddC m_objTransDet, m_iGridCount
        m_iGridCount = m_iGridCount + 1
        txtTotalDebit = m_objTransDets.TotalDebit
        txtTotalCredit = m_objTransDets.TotalCredit
    End If
    txtDebit = 0
    txtCredit = 0
    'txtDescription = "Cash Received"
    txtDescription.SetFocus
    m_bDataChanged = True
End Sub

Private Sub UpdateGrid()
    fgridCTrans.Rows = fgridCTrans.Rows + 1
    fgridCTrans.Row = fgridCTrans.Rows - 2
    fgridCTrans.Col = 1
    fgridCTrans.Text = txtDescription
    fgridCTrans.Col = 2
    fgridCTrans.Text = txtDebit
    fgridCTrans.Col = 3
    fgridCTrans.Text = txtCredit
    fgridCTrans.Col = 4
    fgridCTrans.Text = m_iGridCount
End Sub

Private Sub cmdDel_Click()
    Dim oldrow As Integer
    oldrow = fgridCTrans.Row
    If fgridCTrans.Rows > 2 Then
        fgridCTrans.Col = 4
        m_objTransDets.Remove fgridCTrans.Text
        fgridCTrans.RemoveItem (fgridCTrans.Row)
    ElseIf fgridCTrans.Rows = 2 Then
        fgridCTrans.Col = 1
        fgridCTrans.Text = ""
        fgridCTrans.Col = 2
        fgridCTrans.Text = 0
        fgridCTrans.Col = 3
        fgridCTrans.Text = 0
        fgridCTrans.Col = 4
        fgridCTrans.Text = 0
    End If
    fgridCTrans.Row = oldrow - 1
    txtTotalDebit = m_objTransDets.TotalDebit
    txtTotalCredit = m_objTransDets.TotalCredit
    m_bDataChanged = True
End Sub

Private Sub dtpTransDate_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode, False
End Sub

Private Sub dtpTransDate_LostFocus()
    If dtpTransDate.Value > Date Then
        MsgBox "The date must not be greater than " & Date, vbOKOnly, "Invalid Date"
        dtpTransDate.Value = Date
        dtpTransDate.SetFocus
    End If
End Sub

Private Sub fgridCTrans_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyA Or KeyAscii = (vbKeyA + 32) Then
        txtDescription.SetFocus
    End If
    
    If KeyAscii = vbKeyD Or KeyAscii = (vbKeyD + 32) Then
        fgridCTrans.Col = 4
        If Val(fgridCTrans.Text) <> 0 Then
            cmdDel_Click
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Left = (frmMain.ScaleWidth - Me.ScaleWidth) / 2
    Me.Top = (frmMain.ScaleHeight - Me.ScaleHeight) / 2
    Set m_objConnectDB = New ConnectDB
    Set m_objCIDs = New CIDs
    Set m_objTrans = New CTrans
    Set m_objTransDets = New CTransDets
    Set m_objTransDet = New CTransDet
    Set m_objCCustomer = New CCustomer
    'txtDescription = "Cash Received    "
    m_sTransType = "CV"
    txtDescription = "Cash Received"
    m_objCCustomer.UpdateFormCombo cmbCustomer
    m_objConnectDB.Connect 'Class-Module connect to connect to the database.
    SetNew
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim iUserResp As Integer
   If m_bDataChanged Then
   iUserResp = MsgBox("Save changes?", vbQuestion + vbYesNoCancel, Me.Caption)
      Select Case iUserResp
      Case vbYes
         Save iUserResp
      Case vbCancel
         Cancel = True
      End Select
    Else
        Unload Me
   End If
End Sub
Private Sub mnuDelete_Click()
Delete
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuSave_Click()
    Save
End Sub

Private Sub Delete()
  '   To Delete a transaction
  frmDeleteVoucher.Show
End Sub

Private Sub Save(Optional iResp1 As Integer = 0)
    ' To Save a transaction
    Dim iUserResp As Integer
    If iResp1 = 0 Then
        iUserResp = MsgBox("Do you want to Save the changes", vbQuestion + vbYesNoCancel, "Save Voucher")
    Else
        iUserResp = iResp1
    End If
    If iUserResp = vbYes Then
        If Val(txtTotalDebit) > 0 Or Val(txtTotalCredit) > 0 Then
            m_objCIDs.NewID (m_sTransType)
            Set m_objTrans = New CTrans
            m_objTrans.ID = m_objCIDs.ID
            m_objTrans.TransType = m_sTransType
            m_objTrans.TDate = dtpTransDate.Value
            m_objTrans.Save (m_bIsNewRecord)
            MsgBox "Voucher Number is " & m_objTrans.ID, vbInformation + vbOKOnly, "Customer Voucher Saved"
            m_objTransDets.Save (m_bIsNewRecord), (m_objTrans.ID)
            m_bDataChanged = False
            m_bDirty = False
            SetNew
        Else
            MsgBox "Nothing found to save", vbCritical + vbOKOnly, Me.Caption
        End If
    ElseIf iUserResp = vbNo Then
        Set m_objTrans = New CTrans
        Set m_objTransDets = New CTransDets
        m_bDataChanged = False
        m_bDirty = False
        SetNew
    End If
    
    dtpTransDate.SetFocus
End Sub

Private Sub txtCredit_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtDebit_Change()
    If Len(txtDebit) = 0 Then
        txtDebit = 0
    End If
    If txtDebit > 0 Then
        txtCredit = 0
        txtCredit.Enabled = False
    Else
        txtCredit.Enabled = True
    End If
End Sub

Private Sub txtDebit_GotFocus()
    txtDebit.SelStart = 0
    txtDebit.SelLength = Len(txtDebit)
End Sub

Private Sub txtCredit_Change()
    If Len(txtCredit) = 0 Then
        txtCredit = 0
    End If
End Sub

Private Sub txtCredit_GotFocus()
     txtCredit.SelStart = 0
    txtCredit.SelLength = Len(txtDebit)
End Sub

Private Sub txtDebit_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtDescription_GotFocus()
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription)
End Sub

Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub
