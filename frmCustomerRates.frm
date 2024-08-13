VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCustomerRates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Rates"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   5565
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4411
            MinWidth        =   4411
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid fgridCRates 
      Height          =   3615
      Left            =   780
      TabIndex        =   5
      Top             =   1800
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   6376
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Product Information"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   90
      TabIndex        =   8
      Top             =   600
      Width           =   7335
      Begin MSMask.MaskEdBox txtRate 
         Height          =   315
         Left            =   5235
         TabIndex        =   3
         Top             =   630
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
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
      Begin VB.ComboBox cmbProductGroup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   3360
      End
      Begin VB.CommandButton cmdAdd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         Picture         =   "frmCustomerRates.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdDel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         Picture         =   "frmCustomerRates.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbProducts 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   630
         Width           =   3375
      End
      Begin VB.TextBox txtDecimal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Text            =   "2"
         Top             =   170
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComCtl2.UpDown updnDecimal 
         Height          =   375
         Left            =   6375
         TabIndex        =   11
         Top             =   165
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "Label6"
         BuddyDispid     =   196615
         OrigLeft        =   7800
         OrigTop         =   1800
         OrigRight       =   8040
         OrigBottom      =   2175
         Enabled         =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Product Group"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   14
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   13
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Decimal"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   12
         Top             =   225
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.ComboBox cmbCustomers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1871
      TabIndex        =   0
      Top             =   150
      Width           =   4155
   End
   Begin VB.CommandButton cmdCustomers 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6116
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Customers Info"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblCustSupp 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1038
      TabIndex        =   7
      Top             =   210
      Width           =   750
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmCustomerRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim m_objCustomer As CCustomer
Dim m_objCustRate As CCustomerRate
Dim m_objCustRates As CCustomerRates
Dim m_objProductGroup As CProductGroup
Dim m_objProduct As CProduct
Dim m_bDataChanged As Boolean
Dim iStatusCounter As Integer
'Dim m_iGridCount As Integer

Private Sub cmbCustomers_Click()
    Set m_objCustRates = New CCustomerRates
    UpdateGrid
End Sub

Private Sub cmbCustomers_Validate(Cancel As Boolean)
    If cmbCustomers.ListIndex = -1 And Len(cmbCustomers.Text) > 0 Then
        cmbCustomers.Text = ""
    End If
End Sub
Private Sub cmbProductGroup_click()
    If cmbProductGroup.ListIndex <> -1 Then
        cmbProducts.Clear
        m_objProduct.UpdateFormCombo cmbProducts, cmbProductGroup.ItemData(cmbProductGroup.ListIndex)
    End If

End Sub


Private Sub cmdAdd_Click()
    If cmbCustomers.ListIndex = -1 Then
        MsgBox "You must select a Customer to Add rates", vbCritical + vbOKOnly
    Else
        If cmbProducts.ListIndex = -1 Then
            MsgBox "You must select a product to Add", vbCritical + vbOKOnly
        ElseIf Val(txtRate) = 0 Then
            MsgBox "Rate can't be zero for a product", vbCritical + vbOKOnly
        ElseIf ProductExist = True Then
            MsgBox "Product Already Exists", vbCritical + vbOKOnly
            cmbProducts.SetFocus
        Else
            iStatusCounter = iStatusCounter + 1
            StatusBar1.Panels(1).Text = "Total Products Are " + Str(iStatusCounter)
            fgridCRates.Rows = fgridCRates.Rows + 1
            fgridCRates.Row = fgridCRates.Rows - 2
            fgridCRates.Col = 1
            fgridCRates.Text = cmbProducts.List(cmbProducts.ListIndex)
            fgridCRates.Col = 2
            fgridCRates.Text = txtRate
            Set m_objCustRate = New CCustomerRate
            m_objCustRate.Rate = txtRate
            m_objCustRate.ProductID = cmbProducts.ItemData(cmbProducts.ListIndex)
            fgridCRates.Col = 3
            fgridCRates.Text = cmbProducts.ItemData(cmbProducts.ListIndex)
            'm_iGridCount = m_iGridCount + 1
            m_objCustRates.AddC m_objCustRate, cmbProducts.ItemData(cmbProducts.ListIndex)
            'cmbProducts.ItemData(cmbProducts.ListIndex)
            fgridCRates.TopRow = fgridCRates.Rows - 2
        End If
        txtRate = 0
        cmbProducts.SetFocus
        m_bDataChanged = True
    End If
End Sub

Private Sub cmdCustomers_Click()
    frmListCustomers.Show
End Sub

Private Sub cmdDel_Click()
    Dim oldrow As Integer
    oldrow = fgridCRates.Row
    If fgridCRates.Rows > 2 Then
        fgridCRates.Col = 3
        m_objCustRates.Remove fgridCRates.Text
        fgridCRates.RemoveItem (fgridCRates.Row)
        iStatusCounter = iStatusCounter - 1
        StatusBar1.Panels(1).Text = "Total Products Are " + Str(iStatusCounter)
    ElseIf fgridCRates.Rows = 2 Then
        fgridCRates.Col = 1
        fgridCRates.Text = ""
        fgridCRates.Col = 2
        fgridCRates.Text = ""
        fgridCRates.Col = 3
        fgridCRates.Text = 0
    End If
    fgridCRates.Row = oldrow - 1
    m_bDataChanged = True
End Sub


Private Sub fgridCRates_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyA Or KeyAscii = (vbKeyA + 32) Then
        cmbCustomers.SetFocus
    End If
    
    If KeyAscii = vbKeyD Or KeyAscii = (vbKeyD + 32) Then
        cmdDel_Click
    End If
End Sub

Private Sub Form_Load()
    Me.Left = (frmMain.ScaleWidth - Me.ScaleWidth) / 2
    Me.Top = (frmMain.ScaleHeight - Me.ScaleHeight) / 2
    Set m_objCustRate = New CCustomerRate
    Set m_objCustRates = New CCustomerRates
    Set m_objCustomer = New CCustomer
    Set m_objProduct = New CProduct
    Set m_objProductGroup = New CProductGroup
    m_objCustomer.UpdateFormCombo cmbCustomers
    m_objProductGroup.UpdateFormCombo cmbProductGroup
    m_objProduct.UpdateFormCombo cmbProducts
    iStatusCounter = 0
    StatusBar1.Panels(1).Text = "Total Products Are " + Str(iStatusCounter)
    'm_iGridCount = 1
    txtRate = 0
End Sub

Sub SetNew()
    txtRate = 0
    cmbCustomers.ListIndex = 0
    cmbProducts.ListIndex = -1
    iStatusCounter = 0
    'SetGridTitles
End Sub

Sub SetGridTitles()
  ' Grid Titles
    fgridCRates.Cols = 4
    fgridCRates.Row = 0
    fgridCRates.ColWidth(0) = 200
    fgridCRates.ColWidth(1) = 3800
    fgridCRates.ColWidth(2) = 1200
    fgridCRates.ColWidth(3) = 0
    'fgridCRates.CellAlignment = 1
    fgridCRates.Col = 1
    fgridCRates.Text = "Product Name"
    fgridCRates.Col = 2
    fgridCRates.Text = "Rate"
    StatusBar1.Panels(1).Text = "Total Products Are " + Str(iStatusCounter)
    
End Sub

Private Sub UpdateGrid()
    fgridCRates.Clear
    SetGridTitles
    fgridCRates.Rows = 2
        m_objCustRates.Init cmbCustomers.ItemData(cmbCustomers.ListIndex)
         iStatusCounter = 0
        StatusBar1.Panels(1).Text = "Total Products Are " + Str(iStatusCounter)

   For a = 1 To m_objCustRates.Count
        iStatusCounter = iStatusCounter + 1
        fgridCRates.AddItem (vbTab + m_objCustRates.Item(a).ProductTotalDescription + vbTab + Str(m_objCustRates.Item(a).Rate) + vbTab + Str(m_objCustRates.Item(a).ProductID)), fgridCRates.Rows - 1
    Next
        StatusBar1.Panels(1).Text = "Total Products Are " + Str(iStatusCounter)
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
   End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuSave_Click()
    Save
End Sub

Private Sub Save(Optional iResp1 As Integer = 0)
    Dim iUserResp As Integer
    If iResp1 = 0 Then
        iUserResp = MsgBox("Do you want to Save the changes", vbQuestion + vbYesNoCancel, "Save Customer Rates")
    Else
        iUserResp = iResp1
    End If
    Select Case iUserResp
        Case vbYes
            If cmbCustomers.ListIndex = -1 Then
                MsgBox "You must select a Customer before u can Save", vbCritical + vbOKOnly
            Else
                m_objCustRates.Save (cmbCustomers.ItemData(cmbCustomers.ListIndex))
                m_bDataChanged = False
            End If
        Case vbNo
            If cmbCustomers.ListIndex <> -1 Then
                cmbCustomers_Click
            End If
            m_bDataChanged = False
    End Select
    cmbCustomers.SetFocus
End Sub

Private Sub txtRate_GotFocus()
    txtRate.SelStart = 0
    txtRate.SelLength = Len(txtRate)
End Sub
Private Function ProductExist() As Boolean
    Dim counter As Integer
    For counter = 1 To fgridCRates.Rows - 1
        fgridCRates.Row = counter
        fgridCRates.Col = 3
        If Val(fgridCRates.Text) = cmbProducts.ItemData(cmbProducts.ListIndex) Then
            ProductExist = True
            Exit Function
        End If
    Next counter
    ProductExist = False
End Function


