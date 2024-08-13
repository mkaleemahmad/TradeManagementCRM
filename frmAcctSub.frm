VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAcctSub 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5550
   ClientLeft      =   2550
   ClientTop       =   540
   ClientWidth     =   5415
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5415
   Begin VB.TextBox txtTradeDiscount 
      Height          =   315
      Left            =   1410
      TabIndex        =   11
      Top             =   4813
      Width           =   1035
   End
   Begin VB.TextBox txtCreditDays 
      Height          =   315
      Left            =   1410
      TabIndex        =   12
      Top             =   5160
      Width           =   1035
   End
   Begin VB.TextBox txtNTN 
      Height          =   315
      Left            =   1410
      TabIndex        =   9
      Top             =   4131
      Width           =   3555
   End
   Begin VB.TextBox txtSTaxRegNo 
      Height          =   315
      Left            =   1410
      TabIndex        =   10
      Top             =   4472
      Width           =   3555
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   4050
      Picture         =   "frmAcctSub.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5175
      Width           =   825
   End
   Begin VB.CommandButton cmdOK 
      Height          =   330
      Left            =   4050
      Picture         =   "frmAcctSub.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4830
      Width           =   825
   End
   Begin MSMask.MaskEdBox txtContactTitle 
      Height          =   315
      Left            =   1410
      TabIndex        =   1
      Top             =   1035
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtContactName 
      Height          =   315
      Left            =   1410
      TabIndex        =   0
      Top             =   690
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtAddress 
      Height          =   675
      Left            =   1410
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1380
      Width           =   3555
   End
   Begin MSMask.MaskEdBox txtEmail 
      Height          =   315
      Left            =   1410
      TabIndex        =   8
      Top             =   3790
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFax 
      Height          =   315
      Left            =   1410
      TabIndex        =   7
      Top             =   3449
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtMobileNumber 
      Height          =   315
      Left            =   1410
      TabIndex        =   6
      Top             =   3108
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtPhone2 
      Height          =   315
      Left            =   1410
      TabIndex        =   5
      Top             =   2767
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtPhone1 
      Height          =   315
      Left            =   1410
      TabIndex        =   4
      Top             =   2426
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtCity 
      Height          =   315
      Left            =   1410
      TabIndex        =   3
      Top             =   2085
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E7F9FA&
      Caption         =   "Account's Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   660
      Left            =   -15
      TabIndex        =   28
      Top             =   -15
      Width           =   5055
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Mobile Number"
      Height          =   195
      Index           =   14
      Left            =   120
      TabIndex        =   27
      Top             =   3150
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "National Tax No."
      Height          =   195
      Index           =   13
      Left            =   120
      TabIndex        =   26
      Top             =   4170
      Width           =   1200
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "S. Tax Reg. No."
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   25
      Top             =   4500
      Width           =   1155
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Credit Days"
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   24
      Top             =   5190
      Width           =   810
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Email Address"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   23
      Top             =   3825
      Width           =   990
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Disc./Comm. %"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   4845
      Width           =   1080
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Contact Title:"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   1065
      Width           =   945
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Contact Name:"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Fax Number"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   3495
      Width           =   855
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Phone Number-2"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   2805
      Width           =   1200
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Phone Number-1"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   2460
      Width           =   1200
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "City:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   2115
      Width           =   300
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   1380
      Width           =   615
   End
End
Attribute VB_Name = "frmAcctSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iMode As Integer
Private m_bDirty As Boolean
Private Const MODE_ADD = 1
Private m_objAccts As cAccounts2
Attribute m_objAccts.VB_VarHelpID = -1

Public Sub ShowMe(ByRef objAccts As cAccounts2)
   Set m_objAccts = objAccts
   RefreshData
   Me.Show vbModal, frmAccount2
End Sub

Private Sub chkActive_Click()
End Sub

Private Sub cmdOK_Click()
   Save
   Unload Me
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub txtAddress_gotfocus()
    txtAddress.SelLength = Len(txtAddress.Text)
End Sub

Private Sub txtCity_gotfocus()
txtCity.SelLength = Len(txtCity.Text)
End Sub

Private Sub txtCity_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtContactName_GotFocus()
txtContactName.SelLength = Len(txtContactName.Text)
End Sub

Private Sub txtContactName_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtContactTitle_GotFocus()
txtContactTitle.SelLength = Len(txtContactTitle.Text)
End Sub

Private Sub txtContactTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtCreditDays_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtCustomerID_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtEmail_GotFocus()
txtEmail.SelLength = Len(txtEmail.Text)
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtFax_GotFocus()
    txtFax.SelLength = Len(txtFax.Text)
End Sub

Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtMobileNumber_GotFocus()
   txtMobileNumber.SelLength = Len(txtMobileNumber.Text)
End Sub

Private Sub txtMobileNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtName_Change()
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtNTN_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtPhone1_GotFocus()
txtPhone1.SelLength = Len(txtPhone1.Text)
End Sub

Private Sub txtPhone1_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub RefreshData()
   txtAddress = m_objAccts.Address
   txtCity = m_objAccts.City
   txtPhone1 = m_objAccts.PhoneNumber1
   txtPhone2 = m_objAccts.PhoneNumber2
   txtMobileNumber = m_objAccts.MobileNumber
   txtFax = m_objAccts.FaxNumber
   txtEmail = m_objAccts.Email
   txtContactName = m_objAccts.ContactName
   txtContactTitle = m_objAccts.ContactTitle
   txtTradeDiscount = m_objAccts.TradeDiscount
   txtCreditDays = m_objAccts.CreditDays
   txtSTaxRegNo = m_objAccts.STaxRegNumber
   txtNTN = m_objAccts.NTN
   m_bDirty = False
End Sub

Private Sub txtPhone2_GotFocus()
txtPhone2.SelLength = Len(txtPhone2.Text)
End Sub

Private Sub txtPhone2_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtSTaxRegNo_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtTradeDiscount_GotFocus()
txtTradeDiscount.SelLength = Len(txtTradeDiscount.Text)
End Sub

Private Sub txtTradeDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Sub Save()
   m_objAccts.ContactName = txtContactName
    If txtCreditDays = "" Then
        m_objAccts.CreditDays = 0
    Else
        m_objAccts.CreditDays = txtCreditDays
    End If
    m_objAccts.Email = txtEmail
   m_objAccts.MobileNumber = txtMobileNumber
   m_objAccts.ContactTitle = txtContactTitle
   m_objAccts.Address = txtAddress
   m_objAccts.City = txtCity
    If txtTradeDiscount = "" Then
        m_objAccts.TradeDiscount = 0
    Else
        m_objAccts.TradeDiscount = txtTradeDiscount
    End If
   m_objAccts.NTN = txtNTN
   m_objAccts.PhoneNumber1 = txtPhone1
   m_objAccts.PhoneNumber2 = txtPhone2
   m_objAccts.FaxNumber = txtFax
   m_objAccts.STaxRegNumber = txtSTaxRegNo
End Sub
