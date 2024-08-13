VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class ConnectDB

Option Explicit
Public sConnect As String
Public cnnMyshop As New ADODB.Connection

Sub Connect()
      If cnnMyshop.State <> adStateOpen Then
      cnnMyshop.Open sConnect
        If cnnMyshop.State <> 1 Then
            MsgBox "The connection failed"
        End If
    End If
End Sub

Sub Disconnect()
    If cnnMyshop.State <> adStateClosed Then
        cnnMyshop.Close
    End If
End Sub

Private Sub Class_Initialize()
sConnect = "DSN=sqlMegaMall2007;UID=sa;PWD=mzsol855455noori"
End Sub

Public Sub ReConnect()
Disconnect
DoEvents
      If cnnMyshop.State <> adStateOpen Then
        cnnMyshop.Open sConnect
        DoEvents
        If cnnMyshop.State <> 1 Then
            MsgBox "The connection failed"
        End If
    End If
End Sub
