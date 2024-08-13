VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBalSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents lvwBalance As ListView
Attribute lvwBalance.VB_VarHelpID = -1
Private cn As New ADODB.Connection
Private liEdit As ListItem
Private sType As String

Public Sub AddItem(sName As String)
Dim sSQL As String
Dim li As ListItem
Dim iInd As Long
Dim iOrd As Long
With lvwBalance
    .SetFocus
    If .SelectedItem Is Nothing Then
        iInd = 1
        iOrd = 1
    Else
        iInd = .SelectedItem.Index
        iOrd = CLng(.SelectedItem.Tag)
    End If
    sSQL = "Update BalanceSheet Set Ordinal=Ordinal+1 Where Ordinal>=" & iOrd & " And Type LIKE '" & sType & "'"
    cn.Execute sSQL
    sSQL = "Insert Into BalanceSheet(Name,Type,Ordinal) Values('" & sName & "','" & sType & "'," & iOrd & ");"
    cn.Execute sSQL
    For Each li In .ListItems
        If CLng(li.Tag) >= iOrd Then li.Tag = CLng(li.Tag) + 1
    Next
    Set liEdit = .ListItems.Add(iInd, "ID=" & LastIdentity("BalanceSheet"), sName)
    liEdit.Tag = iOrd
    .SelectedItem = liEdit
    Edit
End With
End Sub

Public Function Delete() As Boolean
Dim sSQL As String
Dim li As ListItem
Dim k As Long
lvwBalance.SetFocus
Set li = lvwBalance.SelectedItem
If li Is Nothing Then Exit Function
' Removing From Database
sSQL = "Delete From BalanceSheet Where " & li.Key
cn.Execute sSQL
sSQL = "Update BalanceSheet Set Ordinal=Ordinal-1 Where Ordinal>" & CLng(li.Tag) & " And Type='" & sType & "';"
cn.Execute sSQL
' Removing Visually
k = CLng(li.Tag)
lvwBalance.ListItems.Remove li.Index
For Each li In lvwBalance.ListItems
    If CLng(li.Tag) > k Then li.Tag = CLng(li.Tag) - 1
Next
Delete = True
End Function

Public Function MoveDown() As Boolean
Dim li As ListItem
lvwBalance.SetFocus
Set li = lvwBalance.SelectedItem
If li Is Nothing Then Exit Function
MoveDown = ExchangeItems(li.Index, li.Index + 1)
End Function

Public Sub Edit()
lvwBalance.SetFocus
lvwBalance.StartLabelEdit
End Sub

Public Function MoveUp() As Boolean
Dim li As ListItem
lvwBalance.SetFocus
Set li = lvwBalance.SelectedItem
If li Is Nothing Then Exit Function
MoveUp = ExchangeItems(li.Index, li.Index - 1)
End Function

Private Sub Class_Terminate()
Set lvwBalance = Nothing
Set cn = Nothing
Set liEdit = Nothing
End Sub

Private Sub lvwBalance_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim sSQL As String
sSQL = "Update BalanceSheet Set Name='" & NewString & "' Where " & liEdit.Key
cn.Execute sSQL
End Sub

Private Sub lvwBalance_BeforeLabelEdit(Cancel As Integer)
Set liEdit = lvwBalance.SelectedItem
End Sub

Function ExchangeItems(ind1 As Long, ind2 As Long) As Boolean
Dim vK1 As String, vT1 As String, iOrd1 As Long, li1 As ListItem
Dim vK2 As String, vT2 As String, iOrd2 As Long, li2 As ListItem
Dim liC As Long
Dim sSQL As String
With lvwBalance
    ' Check Index Bounds
    If ind1 < 1 Or ind2 < 1 Then Exit Function
    liC = .ListItems.Count
    If ind1 > liC Or ind2 > liC Then Exit Function
    ' 1st Item
    Set li1 = .ListItems.Item(ind1)
    vK1 = li1.Key
    vT1 = li1.Text
    iOrd1 = CLng(li1.Tag)
    ' 2nd Item
    Set li2 = .ListItems.Item(ind2)
    vK2 = li2.Key
    vT2 = li2.Text
    iOrd2 = CLng(li2.Tag)
    ' Eschange Items in Database
    sSQL = "UPDATE BalanceSheet SET Ordinal=" & iOrd2 & " Where " & vK1
    sSQL = sSQL & ";" & vbCrLf & _
        "UPDATE BalanceSheet SET Ordinal=" & iOrd1 & " Where " & vK2
    cn.Execute sSQL
    ' Exchange Items Visually
    .ListItems.Item(ind2).Tag = iOrd1
    .ListItems.Remove ind1
    Set li1 = .ListItems.Add(ind2, vK1, vT1)
    li1.Tag = iOrd2
    ' Highlight
    Set .SelectedItem = li1
End With
ExchangeItems = True
End Function

Function LastIdentity(sTableName As String) As Long
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("select Ident_current('" & sTableName & "')")
    Dim v
    v = rs.Fields(0).Value
    LastIdentity = IIf(IsNull(v), 1, v)
    Set rs = Nothing
End Function

Public Sub PopulateLVW()
Dim rs As New ADODB.Recordset
Dim li As ListItem
rs.Open "SELECT * FROM BalanceSheet Where Type Like '" & sType & "' ORDER BY Ordinal ASC;", cn, adOpenForwardOnly, adLockReadOnly
lvwBalance.ListItems.Clear
Do Until rs.EOF
    With lvwBalance
        Set li = .ListItems.Add(, "ID=" & rs("ID").Value, rs("Name").Value)
        li.Tag = rs("Ordinal").Value
    End With
    rs.MoveNext
Loop
End Sub

Public Property Get ActiveConnection() As ADODB.Connection
Set ActiveConnection = cn
End Property

Public Property Set ActiveConnection(ByVal cnNew As ADODB.Connection)
Set cn = cnNew
End Property

Public Property Get BalType() As String
BalType = sType
End Property

Public Property Let BalType(ByVal sNewType As String)
sType = sNewType
End Property

Public Property Get ListView() As ListView
Set ListView = lvwBalance
End Property

Public Property Set ListView(ByVal lvwNew As ListView)
Set lvwBalance = lvwNew
End Property
