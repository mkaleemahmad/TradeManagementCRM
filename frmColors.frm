VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmColors 
   Caption         =   "View Companies"
   ClientHeight    =   5925
   ClientLeft      =   2355
   ClientTop       =   1155
   ClientWidth     =   7260
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   7260
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   5610
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ListView ctlListView 
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1402
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   4586
      EndProperty
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuActionAdd 
         Caption         =   "&Add New Color"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuActionModify 
         Caption         =   "&Modify Color"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuActionDelete 
         Caption         =   "&Delete Color"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuActionDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit Colors"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_frmEdit As frmEditProductCategory
Attribute m_frmEdit.VB_VarHelpID = -1
Private Const ITEMPREFIX = "Item"
Private WithEvents m_objPCategory As CProductCategory
Attribute m_objPCategory.VB_VarHelpID = -1

Private Sub Add()
If Not HasRights(soProductCategories, CanAdd) Then
    MsgBox "You do not have required privileges to proceed !", vbExclamation
    Exit Sub
End If
   '
   ' Create a new instance of the ProductCategory
   ' editor window.
   '
   Set m_frmEdit = New frmEditProductCategory
   m_frmEdit.Add ' objDB:=m_DB
   
End Sub

Public Sub UpdateListItem(objPCategory As CProductCategory)
   Dim objCurrItem As ListItem
   '
   ' Since accessing an item that doesn't
   ' exist will cause an error, "disable"
   ' the error handler temporarily.
   '
   On Error Resume Next
   Set objCurrItem = ctlListView.ListItems(ITEMPREFIX & objPCategory.ID)
   If objCurrItem Is Nothing Then
      Set objCurrItem = _
         ctlListView.ListItems.Add(, _
         ITEMPREFIX & objPCategory.ID, _
         objPCategory.Description)
   End If
   
   With objCurrItem
      .Text = objPCategory.Description
   End With
   
End Sub

Public Sub AddListItem(objPCategory As CProductCategory)
   '
   ' This public method is called after a
   ' new record has been successfully added.
   ' It creates a new ListItem object and
   ' then updates the other columns of the
   ' ListItem.
   '
   ctlListView.ListItems.Add , _
      ITEMPREFIX & objPCategory.ID, _
      objPCategory.Description
   EditListItem objPCategory
End Sub

Private Sub Edit()
If Not HasRights(soProductCategories, CanEdit) Then
    MsgBox "You do not have required privileges to proceed !", vbExclamation
    Exit Sub
End If
   Dim objPCategory As New CProductCategory
   Set m_frmEdit = New frmEditProductCategory
   objPCategory.Init iID:=Mid$(ctlListView.SelectedItem.Key, _
      Len(ITEMPREFIX) + 1)
   m_frmEdit.Edit objProductCategory:=objPCategory
End Sub

Public Sub EditListItem(objPCategory As CProductCategory)
   '
   ' This public method is called by forms
   ' that are editing the contents of the
   ' form. This method is also called
   ' internally to fill the listitem with
   ' the other attributes of the object.
   '
   With ctlListView.ListItems(ITEMPREFIX & objPCategory.ID)
      .Text = objPCategory.Description
'      .SubItems(1) = objCust.City
'      .SubItems(2) = objCust.ContactName
'      .SubItems(3) = objCust.ContactTitle
   End With
End Sub

Private Sub ctlListView_ColumnClick _
   (ByVal ColumnHeader As ComctlLib.ColumnHeader)
   '
   ' This code causes the ListView control
   ' to be sorted by the column that was
   ' clicked by the user.
   '
   ctlListView.SortKey = ColumnHeader.Index - 1
   ctlListView.Sorted = True
End Sub

Private Sub ctlListView_DblClick()
   '
   ' A double-clicked row is the
   ' same as having Edit picked
   ' from the application's menu.
   '
   Edit
End Sub

Private Sub ctlListView_GotFocus()
    StatusBar1.Panels(1).Text = " Total Product Categories Are: " + Str(ctlListView.ListItems.Count)
End Sub

Private Sub ctlListView_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        mnuActionModify_Click
    End If
End Sub

Private Sub Form_Load()
   '
   ' Open the database and load the
   ' records into the window.
   '
   'Set m_DB = OpenDatabase("E:\DevStudio\VB\NWind.MDB")
    Me.Left = (frmMain.ScaleWidth - Me.ScaleWidth) / 2
    Me.Top = (frmMain.ScaleHeight - Me.ScaleHeight) / 2
    Set m_objPCategory = New CProductCategory
   RefreshData
   frmMain.SetTlbLayout 5
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   ctlListView.Height = _
      Me.ScaleHeight - (2 * ctlListView.Top) - StatusBar1.Height
   ctlListView.Width = _
      Me.ScaleWidth - (2 * ctlListView.Left)
   
End Sub

Public Sub RefreshData()
   Dim rsData As ADODB.Recordset
   Set rsData = New ADODB.Recordset
   rsData.Open "Select * from Colors", m_objConnectDB.cnnMyshop, adOpenForwardOnly
   ctlListView.ListItems.Clear
   'Set rsData = m_objCust.f_rsCust
   Do While Not rsData.EOF
      ctlListView.ListItems.Add , _
         ITEMPREFIX & rsData("ID"), _
         rsData("Description")
'
'      With ctlListView.ListItems(ITEMPREFIX _
'         & rsData("CustomerID"))
'         .SubItems(1) = rsData("City") & ""
'         .SubItems(2) = rsData("ContactName") & ""
'         .SubItems(3) = rsData("ContactTitle") & ""
'      End With
      rsData.MoveNext
    '  Exit Do
   Loop
   'rsData.Close
   'Set rsData = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_frmEdit = Nothing
   'Set the old form's toolbar on unload, -5 is for main toolbar
    If Not frmMain.m_lOldTBID = -5 Then
       frmMain.SetTlbLayout frmMain.m_lOldTBID
    Else
       frmMain.SetTlbDefLayout
    End If
End Sub

Private Sub m_frmEdit_DataSaved(objCust As CProductCategory)
   'UpdateListItem objCust
   ctlListView.ListItems.Clear
   RefreshData
End Sub

Private Sub mnuActionAdd_Click()
   Add
End Sub

Private Sub mnuActionDelete_Click()
Delete
End Sub

Private Sub mnuActionModify_Click()
   Edit
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub

Private Sub Delete()
If Not HasRights(soProductCategories, CanDelete) Then
    MsgBox "You do not have required privileges to proceed !", vbExclamation
    Exit Sub
End If
If ctlListView.ListItems.Count < 1 Then
    MsgBox "There is no Product Category to delete.", vbInformation
ElseIf ctlListView.SelectedItem Is Nothing Then
    MsgBox "There is no Product Category selected to delete", vbInformation
ElseIf MsgBox("Are you sure to delete the selected Product Category? ", vbQuestion + vbYesNo) = vbYes Then
   Dim objPCat As New CProductCategory
   If objPCat.CanDelete(Mid$(ctlListView.SelectedItem.Key, Len(ITEMPREFIX) + 1)) Then
        objPCat.Init iID:=Mid$(ctlListView.SelectedItem.Key, _
            Len(ITEMPREFIX) + 1)
        objPCat.Delete
        Set objPCat = Nothing
        MsgBox "Product Category Deleted.", vbInformation
        ctlListView.ListItems.Remove ctlListView.SelectedItem.Key
    Else
        MsgBox "This Product Category can not be deleted.", vbExclamation
    End If
End If
End Sub
