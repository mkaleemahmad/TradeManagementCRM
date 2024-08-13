VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmEditProductCategory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Category Information"
   ClientHeight    =   1995
   ClientLeft      =   2550
   ClientTop       =   540
   ClientWidth     =   6645
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox txtDescription 
      Height          =   312
      Left            =   1068
      TabIndex        =   1
      Top             =   840
      Width           =   5508
      _ExtentX        =   9710
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtProductCategoryID 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1068
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   420
      Width           =   5508
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2101
      TabIndex        =   3
      Top             =   1392
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3321
      TabIndex        =   2
      Top             =   1392
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      Caption         =   "Category ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   10
      Left            =   36
      TabIndex        =   5
      Top             =   444
      Width           =   1092
   End
   Begin VB.Label lblLabel 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   36
      TabIndex        =   4
      Top             =   864
      Width           =   1092
   End
End
Attribute VB_Name = "frmEditProductCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_frmCalledBy As frmListProductCategory
Private m_iMode As Integer
Private m_bDirty As Boolean
Private Const MODE_ADD = 1
Private WithEvents m_objPCategory As CProductCategory
Attribute m_objPCategory.VB_VarHelpID = -1
Public Event DataSaved(objPCategory As CProductCategory)
Dim m_sDesc As String

Public Sub Add()
   Set m_objPCategory = New CProductCategory
   m_objPCategory.Init
   RefreshData
   Me.Show
End Sub

Public Sub Edit(objProductCategory As CProductCategory)
   Set m_objPCategory = objProductCategory
   RefreshData
   Me.Show
End Sub

Private Sub cmdOK_Click()

If txtDescription.Text = "" Then
MsgBox "Not Allowed Empty Description", vbCritical + vbOKOnly, "Error"
End If
     
  If m_objPCategory.IsValid = True Then
  If m_sDesc <> txtDescription.Text Then
  If m_objPCategory.CategoryExist(txtDescription.Text) = True Then
 MsgBox " Cant Not Duplicate Categories", vbCritical + vbOKOnly, "Error"
  Exit Sub
  End If
  Else
  DoEvents
  End If
   Save
   End If
   
   Unload Me
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub Save()
   m_objPCategory.Save
   RaiseEvent DataSaved(m_objPCategory)
'   m_frmCalledBy.UpdateListItem m_objPCategory
'   If m_iMode = MODE_ADD Then
'      m_frmCalledBy.AddListItem m_objPCategory
'   Else
'      m_frmCalledBy.EditListItem m_objPCategory
'   End If
   m_bDirty = False
End Sub

Private Sub Form_Load()
    Me.Top = (frmMain.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmMain.ScaleWidth - Me.ScaleWidth) / 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim iRet As Integer
   If m_bDirty Then
      iRet = MsgBox("Save changes?", vbQuestion + vbYesNoCancel)
      Select Case iRet
      Case vbYes
         Save
      Case vbCancel
         Cancel = True
      End Select
   End If
   
End Sub

Private Sub m_objPCategory_DataValidated(bValid As Boolean)
   cmdOK.Enabled = bValid
   m_bDirty = True
End Sub

Private Sub RefreshData()
   
   txtProductCategoryID = m_objPCategory.ID
   txtDescription = m_objPCategory.Description
   m_sDesc = txtDescription
   
   cmdOK.Enabled = False
   m_bDirty = False
End Sub

Private Sub txtDescription_Change()
    m_objPCategory.Description = txtDescription
End Sub

Private Sub txtDescription_GotFocus()
    txtDescription.SelLength = Len(txtDescription.Text)
End Sub

Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Private Sub txtProductCategoryID_Change()
   ' m_objPCategory.ID = txtProductCategoryID
End Sub

Private Sub txtProductCategoryID_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub
