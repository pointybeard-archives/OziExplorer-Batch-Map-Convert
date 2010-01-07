Attribute VB_Name = "modOziAPIDeclarations"

Declare Function oziGetVersion Lib "OziAPI" (ByRef Data As String, ByRef DataLength As Long) As Long
Declare Function oziGetApiVersion Lib "OziAPI" (ByRef Data As String, ByRef DataLength As Long) As Long
Declare Function oziGetExePath Lib "OziAPI" (ByRef Data As String, ByRef DataLength As Long) As Long

Declare Function oziLoadMap Lib "OziAPI" (ByRef MapName As String) As Long
Declare Function oziFindOzi Lib "OziAPI" () As Long
Declare Function oziSaveMapImage Lib "OziAPI" (ByVal MapName As String) As Integer

