Attribute VB_Name = "Common"
Option Explicit
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerNameW Lib "kernel32" (lpBuffer As Any, nSize As Long) As Long
Private Declare Function GetSystemMenu Lib "User32" _
    (ByVal hWnd As Long, _
    ByVal bRevert As Long) As Long
Private Declare Function ModifyMenu Lib "User32" _
    Alias "ModifyMenuA" _
    (ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long, _
    ByVal wIDNewItem As Long, _
    ByVal lpString As Any) As Long

Private Declare Function GetMenuItemID Lib "User32" _
    (ByVal hMenu As Long, _
    ByVal nPos As Long) As Long

Const MF_BYPOSITION = &H400&
Const MF_GRAYED = &H1&
Const MF_BYCOMMAND = &H0&


Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158

Public Declare Function SendMessage Lib "User32" _
        Alias "SendMessageA" _
    (ByVal hWnd As Long, _
     ByVal wMsg As Integer, _
     ByVal wParam As Integer, _
     lParam As Any) As Long

Public m_objConnectDB As ConnectDB
Public cEvntLg As CTransLog
' Control Accounts
Public Const CashInHand As String = "CashInHand"
Public Const Purchase As String = "Purchase"
Public Const Sales As String = "Sales"
Public Const PurchaseReturn As String = "PurchaseReturn"
Public Const SalesReturn As String = "SalesReturn"
Public Const DiscountsReceived As String = "DiscountsReceived"
Public Const DiscountsOffered As String = "DiscountsOffered"
Public Const Customers As String = "Customers"
Public Const Suppliers As String = "Suppliers"

Public Type AcntIDs
    CashInHand As Long
    Purchase As Long
    Sales As Long
    PurchaseReturn As Long
    SalesReturn As Long
    DiscountsReceived As Long
    DiscountsOffered As Long
    Customers As Long
    Suppliers As Long
End Type
Public SelAcntIDs As AcntIDs
' Control Groups
Public Const grpBank  As String = "Bank"
Public Const grpPARTIES  As String = "Parties"
Public Const grpCUSTOMERs As String = "Customers"
Public Const grpSUPPLIERs As String = "Suppliers"
Public Const grpSALEs As String = "Sales"
Public Const grpPURCHASES As String = "Purchases"
Public Const grpEXPENSE  As String = "Expense"
Public Const grpLCACCTS  As String = "LCAccts"
Public Const grpCASH  As String = "Cash"
Public psCompanyName As String

Function TellCtrlGrp(eCG As enumControlGroup) As String
Select Case eCG
  Case eGrpBANK
    TellCtrlGrp = grpBank
  Case eGrpCUSTOMERS
    TellCtrlGrp = grpCUSTOMERs
  Case eGrpEXPENSE
    TellCtrlGrp = grpEXPENSE
  Case eGrpPARTIES
    TellCtrlGrp = grpPARTIES
  Case eGrpLCACCTS
    TellCtrlGrp = grpLCACCTS
  Case eGrpPURCHASES
    TellCtrlGrp = grpPURCHASES
  Case eGrpSALES
    TellCtrlGrp = grpSALEs
  Case eGrpSUPPLIERS
    TellCtrlGrp = grpSUPPLIERs
  Case egrpCASH
    TellCtrlGrp = grpCASH
End Select
End Function

Public Sub HighlightText(txtBox As TextBox)
If txtBox.Locked Then Exit Sub
txtBox.SelStart = 0
txtBox.SelLength = Len(txtBox)
End Sub

Public Sub HighlightMask(txtBox As MaskEdBox)
txtBox.SelStart = 0
txtBox.SelLength = Len(txtBox)
End Sub

Public Function SoftVal(v)
If IsNull(v) Then
   SoftVal = "" '& vbEmpty
Else
   SoftVal = v
End If
End Function

Sub Main()
'    If App.PrevInstance Then
'        MsgBox "The application is already running.", vbInformation
'        Exit Sub
   If Not ChkPeriod(Date) Then
        MsgBox "A file has been corrupted, contact Microzone to repair", vbCritical
        Exit Sub
    End If
    Set m_objConnectDB = New ConnectDB
    m_objConnectDB.Connect
    PopSelAcntIDs
    Set cEvntLg = New CTransLog
    frmLogin.Show
'    frmMain.Show
    'psCompanyName = "AL-MEHRAG BAKKERS"
    psCompanyName = "WOFASOFT - FAIZ CHINA SHOES"
End Sub

Sub AcceptKeys(ByRef KeyAscii As Integer, AcceptType As KeysGroup, Optional value As String)
    If (KeyAscii < 31 Or AcceptType = AnyKey) And AcceptType <> NoKey Then Exit Sub
    Select Case AcceptType
        Case NoKey
            KeyAscii = 0
       Case Alphabets
          Select Case KeyAscii
             Case 65 To 90, 97 To 122, 32
             Case Else
                KeyAscii = 0
          End Select
       Case AlphaNumeric
          Select Case KeyAscii
             Case 48 To 57, 65 To 90, 97 To 122, 46, 32
             Case Else
                KeyAscii = 0
          End Select
       Case Floats
          If IsMissing(value) Then
             Err.Raise vbObjectError + 1, "Function: AcceptKeys()", "Value parameter is missing"
          End If
          Select Case KeyAscii
             Case 48 To 57
             Case 46
                If InStr(1, value, ".", vbBinaryCompare) > 0 Then
                   KeyAscii = 0
                End If
             Case 45
                If InStr(1, value, "-", vbBinaryCompare) > 0 Then
                   KeyAscii = 0
                End If
             Case Else
                KeyAscii = 0
          End Select
       Case Integers
          Select Case KeyAscii
             Case 48 To 57
             Case Else
                KeyAscii = 0
          End Select
    End Select
End Sub

Sub DisableItem(hWnd As Long, sMenuCaption As String, _
               iMenuPos As Integer)
    'User-defined function to disable the Close button on the
    'MDI Child Form toolbar.
    Dim hMenu As Long
    Dim hItem As Long
    
    hMenu = GetSystemMenu(hWnd, 0)
    hItem = GetMenuItemID(hMenu, iMenuPos)
    Call ModifyMenu(hMenu, hItem, MF_BYCOMMAND Or MF_GRAYED, -9, sMenuCaption)
End Sub

Function HasRights(RightsObject As SecurityObjects, RightName As enumRights) As Boolean
    Dim cUR As New CUserRights
    cUR.Init frmLogin.UserID, RightsObject
    If cUR.ExistInDB Then
      If RightName = CanAdd Then
        HasRights = cUR.CanAdd
      ElseIf RightName = CanDelete Then
        HasRights = cUR.CanDelete
      ElseIf RightName = CanEdit Then
        HasRights = cUR.CanEdit
      ElseIf RightName = CanView Then
        HasRights = cUR.CanView
      Else
        HasRights = False
      End If
    Else
      HasRights = False
    End If
End Function

Function NumVal(sNumber As String) As String
    If Trim(sNumber) = "" Then
      sNumber = "0"
    End If
    NumVal = sNumber
End Function

Function Map2YN(bVal) As String
    If bVal = "" Then
      Map2YN = ""
    ElseIf CBool(bVal) Then
      Map2YN = "Yes"
    Else
      Map2YN = "No"
    End If
End Function

Function TellIndexInDataItem(refCmbBx As ComboBox, lItem As Long) As Long
Dim k As Long
If refCmbBx.ListCount < 1 Then
  TellIndexInDataItem = -1
  Exit Function
End If
For k = 0 To refCmbBx.ListCount - 1
  If refCmbBx.ItemData(k) = lItem Then Exit For
Next
If k > refCmbBx.ListCount - 1 Then
  TellIndexInDataItem = -1
Else
  TellIndexInDataItem = k
End If
End Function

Public Function Trim0(sName As String) As String
   ' Right trim string at first null.
   Dim X As Integer
   X = InStr(sName, vbNullChar)
   If X > 0 Then Trim0 = Left$(sName, X - 1) Else Trim0 = sName
End Function

Sub AddToEvntLg(sEvntNm As String, Optional ByVal sTransType As String, Optional ByVal lTransID As Long)
With cEvntLg
    .Init frmLogin.UserID
    .EventName = sEvntNm
    .TransID = lTransID
    .TransType = sTransType
    .Save True
End With
End Sub

Sub PopSelAcntIDs()
Dim cCA As New ControlAccounts
cCA.Initialize
With SelAcntIDs
    .CashInHand = cCA.AccountNo(CashInHand)
    .Customers = cCA.AccountNo(Customers)
    .DiscountsOffered = cCA.AccountNo(DiscountsOffered)
    .DiscountsReceived = cCA.AccountNo(DiscountsReceived)
    .Purchase = cCA.AccountNo(Purchase)
    .PurchaseReturn = cCA.AccountNo(PurchaseReturn)
    .Sales = cCA.AccountNo(Sales)
    .SalesReturn = cCA.AccountNo(SalesReturn)
    .Suppliers = cCA.AccountNo(Suppliers)
End With
End Sub

Sub EnterKey(keyasc As Variant, Optional upkey As Boolean = True) 'To interpret keys pressed by user
    If keyasc = vbKeyReturn Then
        SendKeys ("{Tab}")
    End If
    If upkey = True Then 'for comboboxes, no up key interpretation
        If keyasc = vbKeyUp Then
            SendKeys ("+{Tab}")
        End If
    End If
End Sub

Function IsNull2(InValue, Substitute)
    If IsNull(InValue) Then
        IsNull2 = Substitute
    Else
        IsNull2 = InValue
    End If
End Function

Function ListBoxFind(refList As ListBox, sText As String, Optional bExactFind As Boolean = True) As Long
    If bExactFind Then
    ListBoxFind = SendMessage(refList.hWnd, LB_FINDSTRINGEXACT, -1, _
                ByVal sText)
    Else
    ListBoxFind = SendMessage(refList.hWnd, LB_FINDSTRING, -1, _
                ByVal sText)
    End If
End Function

Function ComboBoxFind(refCombo As ComboBox, sText As String, Optional bExactFind As Boolean = True) As Long
    If bExactFind Then
    ComboBoxFind = SendMessage(refCombo.hWnd, CB_FINDSTRINGEXACT, -1, _
                ByVal sText)
    Else
    ComboBoxFind = SendMessage(refCombo.hWnd, CB_FINDSTRING, -1, _
                ByVal sText)
    End If
End Function

Public Function ChkPeriod(Optional dTransDate As Date) As Boolean
    If dTransDate < CDate("01/01/2007") Or dTransDate > CDate("31/07/2019") Then
        ChkPeriod = False
    Else
        ChkPeriod = True
    End If
End Function
Public Function BoolConvert(v) As Byte
If v = True Then
   BoolConvert = 1
Else
   BoolConvert = 0
End If
End Function
