VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCashSales 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8685
   ClientLeft      =   15
   ClientTop       =   300
   ClientWidth     =   11790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCashSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11790
   Begin VB.PictureBox Picture1 
      Height          =   1545
      Left            =   0
      Picture         =   "frmCashSales.frx":000C
      ScaleHeight     =   1485
      ScaleWidth      =   11835
      TabIndex        =   8
      Top             =   0
      Width           =   11895
   End
   Begin VB.Timer Timer1 
      Left            =   2520
      Top             =   8040
   End
   Begin VB.Frame Frame2 
      Caption         =   "Summary"
      ForeColor       =   &H00C00000&
      Height          =   5145
      Left            =   8430
      TabIndex        =   5
      Top             =   2565
      Width           =   3012
      Begin VB.TextBox txtCashReceived 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   552
         Left            =   228
         TabIndex        =   18
         Text            =   "0"
         Top             =   3360
         Width           =   2532
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   552
         Left            =   228
         TabIndex        =   10
         Text            =   "0"
         Top             =   2388
         Width           =   2532
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Currency Received"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   228
         TabIndex        =   19
         Top             =   3120
         Width           =   2532
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "This Sale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   228
         TabIndex        =   17
         Top             =   1236
         Width           =   2532
      End
      Begin VB.Label lblThisSale 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   516
         Left            =   228
         TabIndex        =   16
         Top             =   1476
         Width           =   2532
      End
      Begin VB.Label lblTotalItems 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   516
         Left            =   228
         TabIndex        =   15
         Top             =   540
         Width           =   2532
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   228
         TabIndex        =   14
         Top             =   288
         Width           =   2532
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   228
         TabIndex        =   13
         Top             =   2148
         Width           =   2532
      End
      Begin VB.Label lblBalanceAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   516
         Left            =   228
         TabIndex        =   12
         Top             =   4416
         Width           =   2532
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Receivable Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   276
         Left            =   228
         TabIndex        =   11
         Top             =   4140
         Width           =   2532
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgridTrans 
      Height          =   5145
      Left            =   150
      TabIndex        =   2
      Top             =   2565
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   9075
      _Version        =   393216
      Cols            =   7
      ScrollTrack     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   72
      TabIndex        =   4
      Top             =   1560
      Width           =   11370
      Begin VB.TextBox txtStock 
         Enabled         =   0   'False
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
         Left            =   6720
         TabIndex        =   27
         Top             =   570
         Width           =   1335
      End
      Begin VB.TextBox txtProductDescription 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         TabIndex        =   9
         Top             =   570
         Width           =   4545
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
         Height          =   192
         Left            =   10116
         Picture         =   "frmCashSales.frx":FE84
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   252
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
         Height          =   192
         Left            =   9840
         Picture         =   "frmCashSales.frx":1080E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   252
      End
      Begin MSMask.MaskEdBox txtQuantity 
         Height          =   315
         Left            =   5175
         TabIndex        =   20
         Top             =   195
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#######0.000;(#######0.000)"
         Mask            =   "############"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRate 
         Height          =   315
         Left            =   7185
         TabIndex        =   21
         Top             =   195
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         Format          =   "##,###,##0.00;(##,###,##0.00)"
         Mask            =   "##########"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtProductCode 
         Height          =   300
         Left            =   3420
         TabIndex        =   0
         Top             =   204
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "&&&&&"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpPTDate 
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Top             =   195
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   22806529
         CurrentDate     =   37416
      End
      Begin VB.Label lblStockStatus 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Status"
         Height          =   252
         Left            =   8160
         TabIndex        =   28
         Top             =   600
         Width           =   2652
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   6240
         TabIndex        =   26
         Top             =   630
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
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
         Left            =   4470
         TabIndex        =   25
         Top             =   255
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Product Code"
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
         Left            =   2310
         TabIndex        =   24
         Top             =   255
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Rate"
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
         Left            =   6765
         TabIndex        =   23
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date"
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
         Left            =   630
         TabIndex        =   7
         Top             =   255
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Product Description"
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
         Left            =   90
         TabIndex        =   6
         Top             =   615
         Width           =   1545
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit Invoice"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmCashSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim m_bIsCash As Boolean
Dim m_sPTransType As String
Dim m_dtPTDate As Date
Dim m_iSenderReceiverID As Integer
Dim m_sShippingInfo As String
Dim m_dFreight As Double
Dim m_dDiscountPCent As Double
Dim m_dDiscountValue As Double
Dim m_sDescription As String
Dim m_vUserResp As Variant
Dim m_bIsNewRecord As Boolean
Dim m_iProductID As String
Dim m_iGridCount As Integer
Dim m_dTotAmount As Double
Dim m_objCPTrans As CPTrans
Dim m_objCPTransDet As CPTransDet
Dim m_objCPTransDets As CPTransDets
Dim m_objCPGroup As CProductGroup
Dim m_objCProduct As CProduct
Dim m_objCIDs As CIDs
Dim GetProductID As CProduct
Dim m_lTransID As Long
Dim m_objReports As frmReports
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER = &H2000

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Sub Save()
    m_vUserResp = MsgBox("Do You want to Save this Transaction", vbQuestion + vbYesNoCancel)
    If m_vUserResp = vbYes Then
        If Val(lblTotalItems.Caption) = 0 Then
            MsgBox "You must select a Product before you can Save an Invoice", vbOKOnly + vbCritical
        Else
         m_objCIDs.NewID (m_sPTransType)
         Set m_objCPTrans = New CPTrans
         m_objCPTrans.Init
         m_objCPTrans.ID = m_objCIDs.ID
         m_objCPTrans.PTransType = m_sPTransType
         m_objCPTrans.PTDate = dtpPTDate.Value
         m_objCPTrans.SenderReceiverID = m_iSenderReceiverID
         m_objCPTrans.ShippingInfo = m_sShippingInfo
         m_objCPTrans.Freight = m_dFreight
         m_objCPTrans.DiscountPCent = m_dDiscountPCent
         m_objCPTrans.DiscountValue = m_dDiscountValue
         m_objCPTrans.Description = m_sDescription
         m_objCPTrans.Save (m_bIsNewRecord)
         MsgBox "Invoice Number is " & m_objCPTrans.ID, vbInformation + vbOKOnly, "Invoice Saved"
         m_objCPTransDets.Save (m_bIsNewRecord), (m_objCPTrans.ID), m_sPTransType, (m_iSenderReceiverID), (m_dTotAmount - m_dDiscountValue)
         m_bDirty = False
         Set m_objCPTrans = New CPTrans
         Set m_objCPTransDets = New CPTransDets
         SetNew
         txtProductCode.SetFocus
        End If
     ElseIf m_vUserResp = vbNo Then
        Set m_objCPTrans = New CPTrans
        Set m_objCPTransDets = New CPTransDets
        m_bDirty = False
        SetNew
    End If
End Sub

Private Sub SetNew()
    txtStock = 0
    m_dTotAmount = 0
    chkCashCredit = 0
    dtpPTDate.Value = Date
    txtRate.Text = 0
    txtQuantity.Text = 0
    fgridTrans.Clear
    fgridTrans.Rows = 2
    SetGridTitles
    m_sPTransType = "SL"
    m_bIsNewRecord = True
    m_bIsCash = True
    m_iGridCount = 1
    lblBalanceAmount.Caption = 0
    lblThisSale = 0
    lblTotalItems = 0
    txtCashReceived = 0
    txtDiscount = 0
End Sub

Private Sub cmdDel_Click()
    If fgridTrans.Rows > 2 Then
        fgridTrans.Col = 6
        m_objCPTransDets.Remove fgridTrans.Text
        fgridTrans.RemoveItem (fgridTrans.Row)
    ElseIf fgridTrans.Rows = 2 Then
        fgridTrans.Col = 1
        fgridTrans.Text = ""
        fgridTrans.Col = 2
        fgridTrans.Text = ""
        fgridTrans.Col = 3
        fgridTrans.Text = ""
        fgridTrans.Col = 4
        fgridTrans.Text = ""
    End If
        m_dTotAmount = m_objCPTransDets.Total
        lblThisSale = m_objCPTransDets.Total
        lblTotalItems = m_objCPTransDets.TotalQuantity
        lblBalanceAmount = m_dTotAmount - txtCashReceived - txtDiscount
End Sub

Private Sub dtpPTDate_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode, False
End Sub

Private Sub dtpPTDate_LostFocus()
    If dtpPTDate.Value > Date Then
        MsgBox "The date must not be greater than " & Date, vbOKOnly, "Invalid Date"
        dtpPTDate.Value = Date
        dtpPTDate.SetFocus
    End If
End Sub

Private Sub fgridTrans_GotFocus()
    fgridTrans.Row = 1
End Sub

Private Sub fgridTrans_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyD Or KeyAscii = (vbKeyD + 32) Then
        fgridTrans.Col = 6
        If fgridTrans.Text <> "" Then
            cmdDel_Click
        End If
    End If
'    If KeyAscii = vbKeyE Or KeyAscii = (vbKeyE + 32) Then
'        fgridTrans.Col = 6
'        If fgridTrans.Text <> "" Then
'            For a = 0 To cmbProducts.ListCount - 1
'                If cmbProducts.ItemData(a) = fgridTrans.Text Then
'                    cmbProducts.ListIndex = a
'                    Exit For
'                End If
'            Next
'            fgridTrans.Col = 2
'            txtRate = fgridTrans.Text
'            fgridTrans.Col = 3
'            txtQuantity = fgridTrans.Text
'            cmdDel_Click
'        End If
'    End If
End Sub

Private Sub fgridTrans_RowColChange()
 
    '  If fgridTrans.Col = 0 Then
    '     cmbProducts.Left = fgridTrans.CellLeft + fgridTrans.Left
    '     cmbProducts.Top = fgridTrans.CellTop + fgridTrans.Top
    '  End If
 
End Sub

Private Sub Form_Load()
'   Label8.Caption = SysInfo1.OSPlatform
    Me.Top = 0
    Me.Left = 0
    Me.Height = frmMain.ScaleHeight
    Me.Width = frmMain.ScaleWidth
    'Set m_objConnectDB = New ConnectDB
    Set m_objCIDs = New CIDs
    Set m_objCPTrans = New CPTrans
    Set m_objCPTransDets = New CPTransDets
    Set m_objCPGroup = New CProductGroup
    Set m_objCProduct = New CProduct
'    m_objConnectDB.Connect 'Class-Module connect to connect to the database.
'    m_objCPGroup.UpdateFormCombo cmbProductGroupID 'Update Product groups
'    m_objCProduct.UpdateFormCombo cmbProducts, cmbProducts2, , True 'Update Products Combo
'    m_objCCustomer.UpdateFormCombo cmbCustSupp 'Update Customers Combo
    Set m_objReports = New frmReports
    m_iSenderReceiverID = 33
    SetNew
End Sub
Sub SetGridTitles()
  ' Grid Titles
    fgridTrans.Row = 0
  ' fgridTrans.ColWidth(0) = 1000
    fgridTrans.ColWidth(0) = 150
    fgridTrans.ColWidth(1) = 1000
    fgridTrans.ColWidth(2) = 3150
    fgridTrans.ColWidth(3) = 1200
    fgridTrans.ColWidth(4) = 1200
    fgridTrans.ColWidth(5) = 1300
    fgridTrans.ColWidth(6) = 0
  ' fgridTrans.Text = "Product ID"
  ' fgridTrans.Col = 1
    fgridTrans.ColAlignment(1) = flexAlignLeftCenter
    fgridTrans.ColAlignment(2) = flexAlignLeftCenter
    fgridTrans.ColAlignment(3) = flexAlignRightCenter
    fgridTrans.ColAlignment(4) = flexAlignRightCenter
    fgridTrans.ColAlignment(5) = flexAlignRightCenter
    fgridTrans.Col = 1
    fgridTrans.Text = "P.Code"
    fgridTrans.Col = 2
    fgridTrans.Text = "Product Name"
    fgridTrans.Col = 3
    fgridTrans.Text = "Rate"
    fgridTrans.Col = 4
    fgridTrans.Text = "Quantity"
    fgridTrans.Col = 5
    fgridTrans.Text = "Value"
    fgridTrans.Col = 1
End Sub

Private Sub cmdAdd_Click()
    If txtRate = 0 Then
        MsgBox "Rate can't be zero for a sale", vbCritical + vbOKOnly
        If txtRate.Enabled = True Then
            txtRate.SetFocus
        Else
            txtProductCode.SetFocus
        End If
    ElseIf txtQuantity = 0 Then
        MsgBox "Quantity can't be zero for a sale", vbCritical + vbOKOnly
        txtQuantity.SetFocus
    Else
        fgridTrans.Rows = fgridTrans.Rows + 1
        fgridTrans.Row = fgridTrans.Rows - 2
        fgridTrans.Col = 2
        fgridTrans.Text = txtProductDescription
        fgridTrans.Col = 3
        fgridTrans.Text = txtRate
        fgridTrans.Col = 4
        fgridTrans.Text = txtQuantity
        fgridTrans.Col = 5
        If Val(txtRate) > 0 And Val(txtQuantity) > 0 Then
            fgridTrans.Text = Val(txtRate) * Val(txtQuantity)
            Set m_objCPTransDet = New CPTransDet
            m_objCPTransDet.PTransType = m_sPTransType
            m_objCPTransDet.ProductID = m_objCProduct.ID   'm_iProductID
            m_objCPTransDet.Quantity = Val(txtQuantity)
            m_objCPTransDet.Rate = Val(txtRate)
            m_objCPTransDets.AddC m_objCPTransDet, m_iGridCount
        Else
            fgridTrans.Text = 0
        End If
            fgridTrans.Col = 6
            fgridTrans.Text = m_iGridCount
            m_iGridCount = m_iGridCount + 1
            fgridTrans.Col = 1
            fgridTrans.Text = txtProductCode 'm_iProductID
            txtProductCode = ""
            txtProductDescription = ""
            txtRate = 0
            txtQuantity = 0
            lblStockStatus = ""
            txtProductCode.SetFocus
    End If
        m_dTotAmount = m_objCPTransDets.Total
        lblThisSale = m_objCPTransDets.Total
        lblTotalItems = m_objCPTransDets.TotalQuantity
        lblBalanceAmount = m_dTotAmount - txtCashReceived - txtDiscount
End Sub

Private Sub mnuDelete_Click()
    frmDeleteRec.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuSave_Click()
    Save
End Sub

Private Sub txtCashReceived_Change()
    If txtCashReceived = "" Then
        txtCashReceived = 0
    End If
       lblBalanceAmount.Caption = m_dTotAmount - m_dDiscountValue - Val(txtCashReceived)
End Sub

Private Sub txtCashReceived_GotFocus()
    txtCashReceived.SelStart = 0
    txtCashReceived.SelLength = Len(txtCashReceived.Text)
End Sub

Private Sub txtCashReceived_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then   '45=Minus sign, 43=Plus sign
        txtProductCode.SetFocus
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        Save
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDiscount_Change()
    If txtDiscount = "" Then
        txtDiscount = 0
    End If
   'If Val(txtDiscount) > 0 Then
        m_dDiscountValue = Val(txtDiscount)
   'End If
   lblBalanceAmount.Caption = m_dTotAmount - m_dDiscountValue - Val(txtCashReceived)
End Sub

Private Sub txtDiscount_GotFocus()
    txtDiscount.SelStart = 0
    txtDiscount.SelLength = Len(txtDiscount.Text)
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then   '45=Minus sign, 43=Plus sign
        txtProductCode.SetFocus
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        txtCashReceived.SetFocus
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtProductCode_GotFocus()
    txtProductCode.SelStart = 0
    txtProductCode.SelLength = Len(txtProductCode)
End Sub

Private Sub txtProductCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then 'if enter key pressed
        'txtProductCode_LostFocus
        m_objCProduct.FindRec Trim(txtProductCode)
        If txtProductCode <> "" And m_objCProduct.IsFound Then
            txtRate = m_objCProduct.SalePrice
            txtProductDescription = m_objCProduct.Description
            txtStock = m_objCProduct.Balance
            lblStockStatus.Caption = Trim(m_objCProduct.StockStatus)
            txtQuantity.SetFocus
        Else
            MsgBox "This product code is not found", vbOKOnly + vbCritical, "Invalid Code"
            txtProductCode = ""
            txtRate = 0
            txtProductDescription = ""
            txtStock = 0
            txtProductCode.SetFocus
        End If
    End If
    If KeyAscii = 43 Then 'Plus sign
        'txtCashReceived.SetFocus
        txtDiscount.SetFocus
    End If
End Sub

Private Sub txtDescription_Change()
    m_sDescription = txtDescription
End Sub


Private Function NumbersOnly(keyasc As Integer) As Integer
    If (keyasc > vbKey9 Or keyasc < vbKey0) And keyasc <> vbKeyBack And keyasc <> vbKeyDecimal And keyasc <> vbKeyDelete Then
    '    If  Then 'And KeyAsc <> vbKeyBack
            NumbersOnly = 0
           ' Return
        Else
            NumbersOnly = keyasc
          '  Return
    '    End If
    End If
    'NumbersOnly = keyasc
End Function

Private Sub txtQuantity_GotFocus()
    txtQuantity.SelStart = 0
    txtQuantity.SelLength = Len(txtQuantity.Text)
End Sub

Private Sub txtQuantity_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KeyCodeConstants.vbKeyUp Then
        txtProductCode.SetFocus
    ElseIf KeyCode = KeyCodeConstants.vbKeyReturn Then
        cmdAdd_Click
    End If
End Sub

Private Sub txtRate_GotFocus()
If isProductExisting(Trim(txtProductCode.Text)) Then
    Dim strMessage As String
    strMessage = MsgBox("Product Already Entered in Current Invoice", vbOKOnly)
    txtProductCode.SetFocus
End If
    txtRate.SelStart = 0
    txtRate.SelLength = Len(txtRate.Text)
End Sub

Function isProductExisting(strProduct As String) As Boolean
    Dim icounter, iCode As Integer
    iCode = m_objCProduct.GetProductID(strProduct)
    If fgridTrans.Rows > 2 Then
        fgridTrans.Col = 6
        For icounter = 0 To fgridTrans.Rows - 2
            fgridTrans.Row = icounter
            If iCode = Val(fgridTrans.Text) Then
                isProductExisting = True
                Exit Function
            End If
        Next icounter
    End If
    isProductExisting = False
End Function
