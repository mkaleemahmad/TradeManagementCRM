VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBackUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Backup and Restore"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5895
   Begin VB.CommandButton cmdClose 
      Caption         =   "Done"
      Height          =   360
      Left            =   5100
      TabIndex        =   11
      Top             =   390
      Width           =   750
   End
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   4155
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab Page 
      Height          =   1845
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   3254
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Back Up"
      TabPicture(0)   =   "frmBackUp.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtFile"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdBackUpNow"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdBrowse"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Restore"
      TabPicture(1)   =   "frmBackUp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "cmdNavigate"
      Tab(1).Control(2)=   "cmdRestore"
      Tab(1).Control(3)=   "txtBkUpSet"
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtBkUpSet 
         Height          =   315
         Left            =   -74880
         TabIndex        =   7
         Top             =   825
         Width           =   4350
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Restore ..."
         Height          =   360
         Left            =   -74880
         TabIndex        =   6
         Top             =   1215
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "?"
         Height          =   315
         Left            =   -70485
         TabIndex        =   5
         Top             =   825
         Width           =   315
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "?"
         Height          =   315
         Left            =   4515
         TabIndex        =   4
         Top             =   825
         Width           =   315
      End
      Begin VB.CommandButton cmdBackUpNow 
         Caption         =   "Backup Now !"
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   1215
         Width           =   1215
      End
      Begin VB.TextBox txtFile 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   825
         Width           =   4350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Specify File to Use for Retoring Database"
         Height          =   195
         Left            =   -74880
         TabIndex        =   8
         Top             =   570
         Width           =   2925
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Specify Database Backup File"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   570
         Width           =   2145
      End
   End
   Begin VB.DirListBox Folder 
      Height          =   315
      Left            =   60
      TabIndex        =   9
      Top             =   2025
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.FileListBox File 
      Height          =   285
      Left            =   810
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "frmBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim bkupCount As Integer

Private Sub cmdBackUpNow_Click()
On Error GoTo EH
Dim s As String
Dim t As String
Dim k As Long
s = Trim(txtFile.Text)
If s = "" Then
    MsgBox "No File is specified.", vbExclamation, "File Path Error"
    Exit Sub
End If
k = InStrRev(s, "\", , vbTextCompare)
If k = 0 Then
    MsgBox "Path is not valid", vbExclamation, "File Path Error"
    Exit Sub
End If
t = Strings.Left(s, k - 1)
If t = "" Then
    MsgBox "Path is not valid", vbExclamation, "File Path Error"
    Exit Sub
End If
If Dir(t, vbDirectory) = "" Then
    MsgBox "Path is not valid", vbExclamation, "File Path Error"
    Exit Sub
End If
Folder.Path = t
If Dir(s) <> "" Then
    If MsgBox("The specified file already exists, Do you want to overwrite it ?", vbYesNo + vbQuestion, "Overwrite File") = vbNo Then
        Exit Sub
    End If
End If
BackupDatabase s
Exit Sub
EH:
MsgBox Err.Description
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo EH
With cmnDlg
    .CancelError = True
    .DialogTitle = "Where to Place Back Up Files ?"
    .Filter = "SQL Backup|*.SQB|All Files|*.*"
    .InitDir = App.Path
    .FileName = "Backup" & bkupCount
    .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNNoChangeDir
    .ShowSave
    txtFile.Text = .FileName
End With
Exit Sub
EH:
End Sub

Sub BackupDatabase(sFile As String)
Dim sSQL As String
If cn Is Nothing Then
    MsgBox "There is no Connection to Database.", vbOKOnly, "Can Not Backup"
    Exit Sub
ElseIf cn.State = adStateClosed Then
    MsgBox "The Connection to Database is invalid.", vbOKOnly, "Can Not Backup"
    Exit Sub
End If
sSQL = "BACKUP DATABASE " & cn.DefaultDatabase & " TO DISK = '" & sFile & "'"
cn.Execute sSQL
MsgBox "Backup Process Completed Successfully.", vbInformation, "Backup Success"
bkupCount = bkupCount + 1
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNavigate_Click()
On Error GoTo EH
With cmnDlg
    .CancelError = True
    .DialogTitle = "Which File to Use for Restoring Database ?"
    .Filter = "SQL Backup|*.SQB|All Files|*.*"
    .InitDir = App.Path
    .FileName = ""
    .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNNoChangeDir + cdlOFNFileMustExist
    .ShowOpen
    txtBkUpSet.Text = .FileName
End With
Exit Sub
EH:
End Sub

Private Sub cmdRestore_Click()
On Error GoTo EH
Dim s As String
Dim t As String
Dim k As Long
s = Trim(txtBkUpSet.Text)
If s = "" Then
    MsgBox "The specified file not found.", vbExclamation
    Exit Sub
End If
File.FileName = s
RestoreDatabase s
Exit Sub
EH:
MsgBox Err.Description
End Sub

Private Sub Form_Initialize()
bkupCount = 1
End Sub

Sub RestoreDatabase(sFile As String)
On Error GoTo EH
Dim sSQL As String
Dim rs As ADODB.Recordset
If cn Is Nothing Then
    MsgBox "There is no Connection to Database.", vbOKOnly, "Can Not Backup"
    Exit Sub
ElseIf cn.State = adStateClosed Then
    MsgBox "The Connection to Database is invalid.", vbOKOnly, "Can Not Backup"
    Exit Sub
End If
If cn Is Nothing Then
    MsgBox "There is no Connection to Database.", vbOKOnly, "Can Not Backup"
    Exit Sub
ElseIf cn.State = adStateClosed Then
    MsgBox "The Connection to Database is invalid.", vbOKOnly, "Can Not Backup"
    Exit Sub
End If
Dim dbName As String
dbName = cn.DefaultDatabase
sSQL = "RESTORE FILELISTONLY FROM DISK = '" & sFile & "'"
Set rs = cn.Execute(sSQL)
cn.Close
DoEvents
Dim cnRestore As New ADODB.Connection
cnRestore.Open "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=master;Data Source=ACCTT"
DoEvents
sSQL = "RESTORE DATABASE " & dbName & " FROM DISK = '" & sFile & "'"
MsgBox "OK"
cnRestore.Execute sSQL
MsgBox "Restore Process Completed Successfully.", vbInformation, "Backup Success"
cnRestore.Close
cn.Open
Exit Sub
EH:
MsgBox Mid(Err.Description, InStrRev(Err.Description, "]") + 1), vbInformation, "Restore Fail"
cnRestore.Close
cn.Open
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Open m_objConnectDB.cnnMyshop.ConnectionString
Page.Tab = 0
Page.TabVisible(1) = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
End Sub
