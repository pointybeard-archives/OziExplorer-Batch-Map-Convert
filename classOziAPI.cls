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
Public Function OziApiVersion() As String
    If (oziGetApiVersion(OziApiVersion, 5) = -1) Then OziApiVersion = "Unknown"
End Function

Public Function OziVersion() As String
     If (oziGetVersion(OziVersion, 5) = -1) Then OziVersion = "Unknown"
End Function

Public Function askForOzi(Optional sMessage As String = "OziExplorer must be running to use this class") As Boolean

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
