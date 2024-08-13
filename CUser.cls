VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_lUserID As Long
Dim m_sUserName As String
Dim m_sUserDesig As String
Dim m_sPassword As String

Dim m_rsUser As New ADODB.Recordset

Dim ED(1 To 50) As Byte ' Encrypt Decrypt

Public Property Get UserID() As Long
UserID = m_lUserID
End Property

Public Property Let UserID(ByVal NewUserID As Long)
m_lUserID = NewUserID
End Property

Public Property Get UserName() As String
UserName = m_sUserName
End Property

Public Property Let UserName(ByVal NewUserName As String)
m_sUserName = NewUserName
End Property

Public Property Get UserDesig() As String
UserDesig = m_sUserDesig
End Property

Public Property Let UserDesig(ByVal NewDesig As String)
m_sUserDesig = NewDesig
End Property

Public Property Get Password() As String
Password = m_sPassword
End Property

Public Property Let Password(ByVal NewPassword As String)
m_sPassword = NewPassword
End Property

Public Sub AddNew()
With m_rsUser
   If .State = adStateOpen Then .Close
   .Open "Users", m_objConnectDB.cnnMyshop, adOpenKeyset, adLockOptimistic
   .AddNew
End With
End Sub

Public Sub InitWithName(sUserName As String)
With m_rsUser
   If .State = adStateOpen Then .Close
   .Open "Select * From Users Where UserName='" & Encrypt(sUserName) & "';", m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
   If .EOF Then .Close: Exit Sub
   m_lUserID = !UserID
   m_sUserName = Decrypt(!UserName)
   m_sUserDesig = !UserDesig
   m_sPassword = Decrypt(!Password)
End With
End Sub
Public Sub Init(lUserID As Long)
With m_rsUser
   If .State = adStateOpen Then .Close
   .Open "Select * From Users Where UserID=" & lUserID & ";"
   If .EOF Then .Close: Exit Sub
   m_lUserID = !UserID
   m_sUserName = Decrypt(!UserName)
   m_sUserDesig = !UserDesig
   m_sPassword = Decrypt(!Password)
End With
End Sub

Public Sub Save()
With m_rsUser
   If .State = adStateClosed Then MsgBox "Can not save ", vbOKOnly Or vbCritical: Exit Sub
   !UserName = Encrypt(m_sUserName)
   !UserDesig = m_sUserDesig
   !Password = Encrypt(m_sPassword)
   .Update
   .Close
End With
End Sub

Public Sub Remove()
With m_rsUser
   If .State = adStateClosed Then MsgBox "Can not delete.", vbOKOnly Or vbCritical: Exit Sub
   .Delete
   .Close
End With
End Sub

Public Function IsValid() As Boolean
With m_rsUser
   If .State = adStateClosed Then
      IsValid = False
   Else
      IsValid = True
   End If
End With
End Function

Sub PopComboBox(refCmbBx As ComboBox)
Dim k As Long
With m_rsUser
   If .State = adStateOpen Then .Close
   .Open "Select * From Users Order By UserName", m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
   k = 0
   refCmbBx.Clear
   'refCmbBx.Text = ""
   Do Until .EOF
      refCmbBx.AddItem "" & Decrypt(!UserName), k
      refCmbBx.ItemData(k) = !UserID
      k = k + 1
      .MoveNext
   Loop
   If k > 0 Then refCmbBx.ListIndex = 0
End With
End Sub

Sub PopImageCombo(refImgCmb As ImageCombo, ImageIndex As Integer)
On Error GoTo EH
Dim k As Long
Dim s As String
s = "a="
With m_rsUser
   If .State = adStateOpen Then .Close
   .Open "Select * From Users", m_objConnectDB.cnnMyshop, adOpenStatic, adLockOptimistic
   k = 0
   refImgCmb.ComboItems.Clear
    refImgCmb.Text = ""
   Do Until .EOF
    refImgCmb.ComboItems.Add , s & !UserID, "" & Decrypt(!UserName), ImageIndex
      .MoveNext
   Loop
   refImgCmb.GetFirstVisible
End With
Exit Sub
EH:
s = "a" & k & "="
k = k + 1
Resume
End Sub


Function Encrypt(s As String) As String
Dim I As Byte
Dim sn As String
For I = 1 To Len(s)
  sn = sn & Chr(Asc(Mid(s, I, 1)) + Asc(ED(I)))
Next
Encrypt = sn
End Function

Function Decrypt(s As String) As String
Dim I As Byte
Dim sn As String
For I = 1 To Len(s)
  sn = sn & Chr(Asc(Mid(s, I, 1)) - Asc(ED(I)))
Next
Decrypt = sn
End Function

Private Sub Class_Initialize()
ED(1) = 166
ED(2) = 177
ED(3) = 167
ED(4) = 249
ED(5) = 253
ED(6) = 179
ED(7) = 164
ED(8) = 149
ED(9) = 149
ED(10) = 211
ED(11) = 180
ED(12) = 181
ED(13) = 219
ED(14) = 170
ED(15) = 209
ED(16) = 155
ED(17) = 152
ED(18) = 203
ED(19) = 138
ED(20) = 187
ED(21) = 244
ED(22) = 161
ED(23) = 229
ED(24) = 176
ED(25) = 165
ED(26) = 246
ED(27) = 209
ED(28) = 208
ED(29) = 183
ED(30) = 141
ED(31) = 200
ED(32) = 217
ED(33) = 245
ED(34) = 235
ED(35) = 131
ED(36) = 198
ED(37) = 245
ED(38) = 183
ED(39) = 215
ED(40) = 192
ED(41) = 194
ED(42) = 187
ED(43) = 173
ED(44) = 180
ED(45) = 163
ED(46) = 135
ED(47) = 159
ED(48) = 253
ED(49) = 136
ED(50) = 178
End Sub
