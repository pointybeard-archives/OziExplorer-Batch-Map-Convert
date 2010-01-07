Attribute VB_Name = "modLibrary"
Public objOziAPI As New classOziAPI
Public sMapFileArray() As String

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
