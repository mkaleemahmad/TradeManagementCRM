VERSION 5.00
Begin VB.Form frmDeleteVoucher 
   Caption         =   "Voucher Deletion"
   ClientHeight    =   1410
   ClientLeft      =   3330
   ClientTop       =   4140
   ClientWidth     =   4740
   HasDC           =   0   'False
   Icon            =   "frmDeleteVoucher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4740
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   2280
      TabIndex        =   1
      Text            =   "Sales Invoice"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Voucher Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
End
Attribute VB_Name = "frmDeleteVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_objTrans As CTrans
Dim m_objCTransDet As CTransDet
Dim m_lTransID As Long
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    EnterKey KeyCode, False
End Sub

Private Sub Command1_Click()
    
    Dim strMessage As String
    If Combo1.ListIndex = -1 Then
        MsgBox "No Item Selected.", vbInformation
        Exit Sub
    End If
    m_lTransID = Val(Combo1.ItemData(Combo1.ListIndex))
    
        If m_objTrans.FindRec(m_lTransID, "CV") = True Then
            m_objTrans.DeleteRec m_lTransID, "CV"
            m_objCTransDet.DeleteRec m_lTransID, "CV"
            strMessage = MsgBox("Operation Completed ", vbOKOnly)
        Else
            strMessage = MsgBox("No Voucher Found To Delete", vbInformation, vbOKOnly)
        End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Combo1.Clear
    Set m_objTrans = New CTrans
    Set m_objCTransDet = New CTransDet
    m_objTrans.UpdateFormCombo Combo1, "CV"
    Combo1.ListIndex = Combo1.ListCount - 1
End Sub


