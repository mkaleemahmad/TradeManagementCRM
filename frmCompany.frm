VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Basic Information"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7260
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker dtStartDate 
      Height          =   345
      Left            =   1725
      TabIndex        =   19
      Top             =   3825
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      _Version        =   393216
      Format          =   22675457
      CurrentDate     =   37795
   End
   Begin MSComCtl2.DTPicker dtEndDate 
      Height          =   345
      Left            =   1725
      TabIndex        =   21
      Top             =   4245
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      _Version        =   393216
      Format          =   22675457
      CurrentDate     =   37795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   3090
      TabIndex        =   22
      Top             =   4725
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   1725
      TabIndex        =   23
      Top             =   4725
      Width           =   1260
   End
   Begin VB.TextBox txtNTNo 
      Height          =   345
      Left            =   1725
      MaxLength       =   10
      TabIndex        =   17
      Top             =   3416
      Width           =   1230
   End
   Begin VB.TextBox txtSTRNo 
      Height          =   345
      Left            =   1725
      MaxLength       =   17
      TabIndex        =   15
      Top             =   3004
      Width           =   1935
   End
   Begin VB.TextBox txtETRate 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   1725
      MaxLength       =   5
      TabIndex        =   13
      Top             =   2592
      Width           =   660
   End
   Begin VB.TextBox txtSTRate 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   1725
      MaxLength       =   5
      TabIndex        =   11
      Top             =   2180
      Width           =   660
   End
   Begin VB.TextBox txtPhone 
      Height          =   345
      Left            =   1725
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1768
      Width           =   4320
   End
   Begin VB.TextBox txtCity 
      Height          =   345
      Left            =   1725
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1356
      Width           =   5385
   End
   Begin VB.TextBox txtAddress 
      Height          =   345
      Left            =   1725
      MaxLength       =   50
      TabIndex        =   5
      Top             =   944
      Width           =   5385
   End
   Begin VB.TextBox txtSlogan 
      Height          =   345
      Left            =   1725
      MaxLength       =   50
      TabIndex        =   3
      Top             =   532
      Width           =   3270
   End
   Begin VB.TextBox txtCompany 
      Height          =   345
      Left            =   1725
      MaxLength       =   50
      TabIndex        =   1
      Top             =   120
      Width           =   3270
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year End Date"
      Height          =   195
      Index           =   10
      Left            =   45
      TabIndex        =   20
      Top             =   4320
      Width           =   1050
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year Start Date"
      Height          =   195
      Index           =   9
      Left            =   45
      TabIndex        =   18
      Top             =   3900
      Width           =   1095
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "National Tax No."
      Height          =   195
      Index           =   8
      Left            =   45
      TabIndex        =   16
      Top             =   3495
      Width           =   1200
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax Reg. No."
      Height          =   195
      Index           =   7
      Left            =   45
      TabIndex        =   14
      Top             =   3075
      Width           =   1395
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ExSTax"
      Height          =   195
      Index           =   6
      Left            =   45
      TabIndex        =   12
      Top             =   2670
      Width           =   555
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax Rate"
      Height          =   195
      Index           =   5
      Left            =   45
      TabIndex        =   10
      Top             =   2250
      Width           =   1095
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone, Fax"
      Height          =   195
      Index           =   4
      Left            =   45
      TabIndex        =   8
      Top             =   1845
      Width           =   810
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Index           =   3
      Left            =   45
      TabIndex        =   6
      Top             =   1425
      Width           =   255
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Index           =   2
      Left            =   45
      TabIndex        =   4
      Top             =   1020
      Width           =   570
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Slogan"
      Height          =   195
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      Height          =   195
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   195
      Width           =   1125
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Company As New CCompany
Dim m_bDirty As Boolean

Sub UpdateData(bClass As Boolean)
If bClass Then
    With Company
        .Address = txtAddress
        .City = txtCity
        .CompanyName = txtCompany
        .ExSTax = Val(txtETRate)
        .Fax = txtPhone
        .NationalTaxNo = txtNTNo
        .Phone = txtPhone
        .SalesTaxRegistrationNo = txtSTRNo
        .SaleTaxRate = Val(txtSTRate)
        .Slogan = txtSlogan
        .StartDate = dtStartDate.value
        .EndDate = dtEndDate.value
    End With
Else
    With Company
        txtAddress = .Address
        txtCity = .City
        txtCompany = .CompanyName
        txtETRate = .ExSTax
        txtPhone = .Fax
        txtNTNo = Trim(.NationalTaxNo)
        txtPhone = .Phone
        txtSTRNo = Trim(.SalesTaxRegistrationNo)
        txtSTRate = Trim(.SaleTaxRate)
        txtSlogan = .Slogan
        dtStartDate.value = .StartDate
        dtEndDate.value = .EndDate
    End With
End If
End Sub

Private Sub cmdCancel_Click()
m_bDirty = False
Unload Me
End Sub

Private Sub cmdOK_Click()
UpdateData True
Company.Save
m_bDirty = False
MsgBox "Company Information Saved.", vbInformation
End Sub

Private Sub dtEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub dtStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Load()
Company.Init
UpdateData False
m_bDirty = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim rV As VbMsgBoxResult
If m_bDirty Then
    rV = MsgBox("The company information has changed. Do you want to save it", vbQuestion + vbYesNoCancel)
    If rV = vbYes Then
        cmdOK_Click
    ElseIf rV = vbCancel Then
        Cancel = 1
    End If
End If
End Sub

Private Sub txtAddress_Change()
m_bDirty = True
End Sub

Private Sub txtAddress_gotfocus()
HighlightText txtAddress
End Sub

Private Sub txtCity_Change()
m_bDirty = True
End Sub

Private Sub txtCity_gotfocus()
HighlightText txtCity
End Sub

Private Sub txtCompany_Change()
m_bDirty = True
End Sub

Private Sub txtCompany_gotfocus()
HighlightText txtCompany
End Sub

Private Sub txtETRate_Change()
m_bDirty = True
End Sub

Private Sub txtETRate_GotFocus()
HighlightText txtETRate
End Sub

Private Sub txtETRate_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtNTNo_Change()
m_bDirty = True
End Sub

Private Sub txtNTNo_GotFocus()
HighlightText txtNTNo
End Sub

Private Sub txtPhone_Change()
m_bDirty = True
End Sub

Private Sub txtPhone_gotfocus()
HighlightText txtPhone
End Sub

Private Sub txtSlogan_Change()
m_bDirty = True
End Sub

Private Sub txtSlogan_gotfocus()
HighlightText txtSlogan
End Sub

Private Sub txtSTRate_Change()
m_bDirty = True
End Sub

Private Sub txtSTRate_GotFocus()
HighlightText txtSTRate
End Sub

Private Sub txtSTRate_KeyPress(KeyAscii As Integer)
AcceptKeys KeyAscii, Integers
End Sub

Private Sub txtSTRNo_Change()
m_bDirty = True
End Sub

Private Sub txtSTRNo_GotFocus()
HighlightText txtSTRNo
End Sub
