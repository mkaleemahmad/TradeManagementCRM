VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmDeleteRec 
   Caption         =   "Invoice Deletion"
   ClientHeight    =   1965
   ClientLeft      =   3645
   ClientTop       =   3975
   ClientWidth     =   3615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmDeleteRec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3615
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   312
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   972
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   6
      Format          =   "0"
      Mask            =   "######"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   300
      TabIndex        =   3
      Top             =   750
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1575
      TabIndex        =   2
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Entry Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmDeleteRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_objCPTrans As CPTrans
Dim m_objCPTransDet As CPTransDet
Dim m_objTrans As CTrans
Dim m_objCTransDet As CTransDet

Dim m_sTransType As String
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode, False
End Sub

Private Sub Command1_Click()
            Dim strPTransType, strMessage As String
            Dim iComboData As Integer
            Dim lValue As Long
                    Select Case Combo1.ListIndex
                        Case -1
                            strPTransType = "SL"
                        Case 0
                            strPTransType = "SL"
                        Case 1
                            strPTransType = "SR"
                        Case 2
                            strPTransType = "PH"
                    End Select
                    lValue = Val(MaskEdBox1.Text)
                    If m_objCPTrans.FindRec(lValue, m_sTransType) = True Then
                        m_objCPTrans.DeleteRec Val(MaskEdBox1.Text), m_sTransType
                        'm_objCPTransDet.DeleteRec Val(MaskEdBox1.Text), strPTransType
                        m_objTrans.DeleteRec Val(MaskEdBox1.Text), m_sTransType
                        'm_objCTransDet.DeleteRec Val(MaskEdBox1.Text), strPTransType
                        
                        strMessage = MsgBox("Operation Completed", vbInformation, vbOKOnly)
                    Else
                        strMessage = MsgBox("No Invoice Found To Delete", vbOKOnly)
                    End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Combo1.AddItem "Sales Invoice"
    Combo1.AddItem "Sale Return Invoice"
    Combo1.AddItem "Purchase Invoice"

    Set m_objCPTrans = New CPTrans
    Set m_objCPTransDet = New CPTransDet
    Set m_objTrans = New CTrans
    Set m_objCTransDet = New CTransDet

End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode
End Sub

Public Sub ShowMe(sTransType As String, Optional Modal = vbModeless, Optional OwnerForm = Nothing)
Load Me
m_sTransType = sTransType
Select Case LCase(sTransType)
Case "os"
    Combo1.Text = "Opening Stock"
Case "ph"
    Combo1.Text = "Purchase Invoice"
Case "sl"
    Combo1.Text = "Sales Invoice"
Case "pr"
    Combo1.Text = "Purchase Return"
Case "sr"
    Combo1.Text = "Sales Return"
End Select
Me.Show Modal, OwnerForm
End Sub

