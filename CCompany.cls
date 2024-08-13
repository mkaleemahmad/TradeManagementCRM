VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tpCompany
    sCompanyName As String
    sSlogan As String
    sAddress As String
    sCity As String
    sPhone As String
    sFax As String
    dSaleTaxRate As Double
    dExSTax As Double
    sSalesTaxRegistrationNo As String
    sNationalTaxNo As String
    dtStartDate As Date
    dtEndDate As Date
End Type

Dim inf As tpCompany

Dim rs As ADODB.Recordset

Sub UpdateDate(bRecordset As Boolean)
If bRecordset Then
    With rs
        !CompanyName = inf.sCompanyName
        !Slogan = inf.sSlogan
        !Address = inf.sAddress
        !City = inf.sCity
        !Phone = inf.sPhone
        !Fax = inf.sFax
        !SaleTaxRate = inf.dSaleTaxRate
        !ExSTax = inf.dExSTax
        !SalesTaxRegistrationNo = inf.sSalesTaxRegistrationNo
        !NationalTaxNo = inf.sNationalTaxNo
        !StartDate = inf.dtStartDate
        !EndDate = inf.dtEndDate
    End With
Else
    With rs
        If .BOF And .EOF Then
            inf.sCompanyName = ""
            inf.sSlogan = ""
            inf.sAddress = ""
            inf.sCity = ""
            inf.sPhone = ""
            inf.sFax = ""
            inf.dSaleTaxRate = 0
            inf.dExSTax = 0
            inf.sSalesTaxRegistrationNo = ""
            inf.sNationalTaxNo = ""
            inf.dtStartDate = Date
            inf.dtEndDate = Date
        Else
            inf.sCompanyName = !CompanyName
            inf.sSlogan = IsNull2(!Slogan, "")
            inf.sAddress = IsNull2(!Address, "")
            inf.sCity = IsNull2(!City, "")
            inf.sPhone = IsNull2(!Phone, "")
            inf.sFax = IsNull2(!Fax, "")
            inf.dSaleTaxRate = IsNull2(!SaleTaxRate, 0)
            inf.dExSTax = IsNull2(!ExSTax, 0)
            inf.sSalesTaxRegistrationNo = IsNull2(!SalesTaxRegistrationNo, "")
            inf.sNationalTaxNo = IsNull2(!NationalTaxNo, "")
            inf.dtStartDate = IsNull2(!StartDate, Date)
            inf.dtEndDate = IsNull2(!EndDate, Date)
        End If
    End With
End If
End Sub

Public Sub Init()
    Dim sSQL As String
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
            rs.Close
        End If
        Set rs = Nothing
    End If
    sSQL = "SELECT Top 1 * FROM Company"
    Set rs = New ADODB.Recordset
    rs.Open sSQL, m_objConnectDB.cnnMyshop, adOpenKeyset, adLockOptimistic
    UpdateDate False
End Sub

Public Sub Save()
    Dim sSQL As String
    sSQL = "DELETE FROM Company"
    m_objConnectDB.cnnMyshop.Execute (sSQL)
    If rs Is Nothing Then
        sSQL = "SELECT Top 1 * FROM  Company"
        Set rs = New ADODB.Recordset
        rs.Open sSQL, m_objConnectDB.cnnMyshop, adOpenKeyset, adLockOptimistic
    End If
    If rs.State = adStateClosed Then rs.Open
    rs.AddNew
    UpdateDate True
    rs.Update
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Public Property Get CompanyName() As String
    CompanyName = inf.sCompanyName
End Property

Public Property Let CompanyName(ByVal sNV As String)
inf.sCompanyName = sNV
End Property

Public Property Get Slogan() As String
Slogan = inf.sSlogan
End Property

Public Property Let Slogan(ByVal sNV As String)
inf.sSlogan = sNV
End Property

Public Property Get Address() As String
Address = inf.sAddress
End Property

Public Property Let Address(ByVal sNV As String)
inf.sAddress = sNV
End Property

Public Property Get City() As String
City = inf.sCity
End Property

Public Property Let City(ByVal sNV As String)
inf.sCity = sNV
End Property

Public Property Get Phone() As String
Phone = inf.sPhone
End Property

Public Property Let Phone(ByVal sNV As String)
inf.sPhone = sNV
End Property

Public Property Get Fax() As String
Fax = inf.sFax
End Property

Public Property Let Fax(ByVal sNV As String)
inf.sFax = sNV
End Property

Public Property Get SaleTaxRate() As Double
SaleTaxRate = inf.dSaleTaxRate
End Property

Public Property Let SaleTaxRate(ByVal dNV As Double)
inf.dSaleTaxRate = dNV
End Property

Public Property Get SalesTaxRegistrationNo() As String
SalesTaxRegistrationNo = inf.sSalesTaxRegistrationNo
End Property

Public Property Let SalesTaxRegistrationNo(ByVal sNV As String)
inf.sSalesTaxRegistrationNo = sNV
End Property

Public Property Get NationalTaxNo() As String
NationalTaxNo = inf.sNationalTaxNo
End Property

Public Property Let NationalTaxNo(ByVal sNV As String)
inf.sNationalTaxNo = sNV
End Property

Public Property Get ExSTax() As Double
ExSTax = inf.dExSTax
End Property

Public Property Let ExSTax(ByVal dNV As Double)
inf.dExSTax = dNV
End Property

Public Property Get StartDate() As Date
StartDate = inf.dtStartDate
End Property

Public Property Let StartDate(ByVal dtNV As Date)
inf.dtStartDate = dtNV
End Property

Public Property Get EndDate() As Date
EndDate = inf.dtEndDate
End Property

Public Property Let EndDate(ByVal dtNV As Date)
inf.dtEndDate = dtNV
End Property
