VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classOziAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub Destroy()
    oziCloseApi
End Sub

Public Function apiVersion() As Double
    Dim ret As String
    If (oziGetApiVersion(ret, 5) = -1) Then
        apiVersion = -1
    Else
        apiVersion = Val(ret)
    End If
End Function

Public Function exeVersion() As Double
    Dim ret As String
    If (oziGetVersion(ret, 5) = -1) Then
        exeVersion = -1
    Else
        exeVersion = Val(ret)
    End If
End Function

Public Function askForOziAPI(Optional sMessage As String = "OziExplorer API 1.08 or greater must be installed") As Boolean
    Dim apiVersion As Double
    apiVersion = Me.apiVersion()
    
     Do While (apiVersion < 1.08)
        answer = MsgBox(sMessage, vbCritical + vbRetryCancel, "OziExplorer API v1.08 or greater required")
        
        If answer = vbCancel Then
            askForOziAPI = False
            Exit Function
        End If
        
        apiVersion = Me.apiVersion
    Loop
    
    askForOziAPI = True
End Function

Public Function askForOzi(Optional sMessage As String = "OziExplorer must be running to use this class") As Boolean
    Dim isOziExplorerRunning As Boolean
    isOziExplorerRunning = oziFindOzi
    
     Do While (isOziExplorerRunning = -1)
        answer = MsgBox(sMessage, vbCritical + vbRetryCancel, "OziExplorer Not Running")
        
        If answer = vbCancel Then
            askForOzi = False
            Exit Function
        End If
        
        isOziExplorerRunning = oziFindOzi
    Loop
    
    askForOzi = True
    
End Function

