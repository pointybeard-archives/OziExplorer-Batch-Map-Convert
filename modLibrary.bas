Attribute VB_Name = "modLibrary"
Option Explicit

Public objOziAPI As New classOziAPI
Public sMapFileArray() As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long
   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function

Public Function OpenBrowser(ByVal URL As String) As Boolean
    Dim res As Long
    
    ' it is mandatory that the URL is prefixed with http:// or https://
    If InStr(1, URL, "http", vbTextCompare) <> 1 Then
        URL = "http://" & URL
    End If
    
    res = ShellExecute(0&, "open", URL, vbNullString, vbNullString, vbNormalFocus)
    OpenBrowser = (res > 32)
End Function

Public Function FindFilesByExtension(ByRef fso As FileSystemObject, _
                         ByRef fld As Folder, _
                         ByVal sFol As String, _
                         sFileExt As String, _
                         ByRef sOutput, _
                         Optional bRecurseIntoSubdirectories As Boolean = False) As Integer

    Dim tFld As Folder, tFil As file, filename As String

    Set fld = fso.GetFolder(sFol)

    sFileExt = LCase(sFileExt)
    
    For Each tFil In fld.Files
        If LCase(getFileExtension(tFil.Name)) = sFileExt Then
            sOutput(UBound(sOutput)) = fso.BuildPath(fld.Path, tFil.Name)
            ReDim Preserve sOutput(UBound(sOutput) + 1) As String
            FindFilesByExtension = FindFilesByExtension + 1
        End If
        DoEvents
    Next
    
    If (fld.SubFolders.Count > 0 And bRecurseIntoSubdirectories = True) Then
        For Each tFld In fld.SubFolders
            DoEvents
            FindFilesByExtension = FindFilesByExtension + FindFilesByExtension(fso, fld, tFld.Path, sFileExt, sOutput)
        Next
    End If
End Function

Public Function getFileExtension(sFile As String) As String
    getFileExtension = Mid(sFile, InStrRev(sFile, ".") + 1)
End Function

Public Function getFileNameFromPath(sFile As String) As String
    getFileNameFromPath = Mid(sFile, InStrRev(sFile, "\") + 1)
End Function

Public Function CombineArrays(ByVal Array1, ByVal Array2) As String()
    Dim newArray() As String
    ReDim newArray(0) As String
    
    For Each Item In Array1
        newArray(UBound(newArray)) = Item
        ReDim Preserve newArray(UBound(newArray) + 1) As String
    Next
    

    For Each Item In Array2
        newArray(UBound(newArray)) = Item
        ReDim Preserve newArray(UBound(newArray) + 1) As String
    Next
 
    CombineArrays = newArray
    
End Function
