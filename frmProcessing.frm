VERSION 5.00
Begin VB.Form frmProcessing 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Processing Maps"
   ClientHeight    =   900
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin OziExpBatchMapConv.ucProgressBar pbProcessing 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
   End
   Begin VB.Label lblProcessing 
      Caption         =   "processing ..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "frmProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim killProcessing As Boolean

Public Sub beginProcessing(outputType As String, _
                           Optional destinationDirectory As String = vbNullString, _
                           Optional regenerateCalibrationFiles As Boolean = False, _
                           Optional promptForMapName As Boolean = False)

    killProcessing = False

    Me.MousePointer = vbHourglass
    Me.pbProcessing.Value = 0
    
    Dim ii As Integer
    Dim file As String
    Dim retVal
    Dim destinationFile As String
    Dim progressbarStep As Double
    
    Dim errorLog() As String
    ReDim errorLog(0) As String
    
    ' Used later to check if a .MAP file has an accompanying .OZF file
    Dim oFile As New Scripting.FileSystemObject
   
    progressbarStep = 1
    Me.pbProcessing.Max = (UBound(sMapFileArray) * 2)
    
    For ii = 0 To UBound(modLibrary.sMapFileArray) - 1
    
        If killProcessing = True Then Exit For
    
        file = modLibrary.sMapFileArray(ii)
        Me.lblProcessing.Caption = file
        Me.pbProcessing.Value = Me.pbProcessing.Value + progressbarStep
        Me.Caption = "Processing Maps - " & (ii + 1) & " of " & UBound(modLibrary.sMapFileArray)
        
        DoEvents

        If (oFile.FileExists(Replace(file, getFileExtension(file), "ozf")) = False) Then
            ReDim Preserve errorLog(UBound(errorLog) + 1) As String
            errorLog(UBound(errorLog)) = "ERROR: No corresponding .OZF file could be located. File: " & file
            
        ElseIf (oziLoadMap(file) = -1) Then
            ReDim Preserve errorLog(UBound(errorLog) + 1) As String
            errorLog(UBound(errorLog)) = "ERROR: OziExplorer could not load map. File: " & file
        
        Else
            Me.pbProcessing.Value = Me.pbProcessing.Value + progressbarStep
            DoEvents
        
            If (destinationDirectory = vbNullString) Then
                destinationFile = Replace(file, getFileExtension(file), outputType)
            Else
                destinationFile = destinationDirectory & Replace(getFileNameFromPath(file), getFileExtension(file), outputType)
            End If

            retVal = oziSaveMapImage(destinationFile)
        
            If (retVal = -1 Or oFile.FileExists(destinationFile) = False) Then
                ReDim Preserve errorLog(UBound(errorLog) + 1) As String
                errorLog(UBound(errorLog)) = "ERROR: Destination file could not be saved. File: " _
                    & file & " Destination: " & destinationFile
            
            'Regenerate the MAP file
            ElseIf (regenerateCalibrationFiles = True) Then

                Dim sMapName As String, sNewMapFile As String
                
                sMapName = getFileNameFromPath(destinationFile)
                sNewMapFile = Replace(destinationFile, getFileExtension(destinationFile), "MAP")
                
                If (promptForMapName = True) Then
                    sMapName = InputBox("Please enter a name for the map " & sMapName, "Enter Map Name", sMapName)
                End If
                
                ' Read in the existing file
                Dim tmp As String, contents As String, currLine As Integer
                currLine = 0
                contents = vbNullString
                
                Open file For Input As #1
                While EOF(1) = 0
                    currLine = currLine + 1
                    Line Input #1, tmp
                    
                    'Replace the second line with the new map name
                    If (currLine = 2) Then
                        tmp = sMapName
                        
                    'Replace the 3rd line with new map filename
                    ElseIf (currLine = 3) Then
                        tmp = getFileNameFromPath(Replace(destinationFile, getFileExtension(destinationFile), "ozf"))
                        
                    End If
                    
                    contents = contents + tmp + vbNewLine
                Wend
                Close #1
                
                'Save the new map file
                Open sNewMapFile For Output As #1
                    Print #1, contents
                Close #1
                
            End If
        End If
        
    Next
    
    Dim sError As Variant, errorCount As Integer, errorLogPath As String
    errorLogPath = CurDir + "\error.log"
    errorCount = 0
    
    Open errorLogPath For Output As #1
        For Each sError In errorLog
            If (LenB(sError) > 0) Then
                errorCount = errorCount + 1
                Write #1, sError
            End If
        Next
    Close #1
    
    If (errorCount > 0) Then MsgBox errorCount & _
                                    " errors were encountered. See " & _
                                    errorLogPath & _
                                    " for details", _
                                vbOKOnly + vbCritical, "Errors During Processing"
    
    Me.MousePointer = vbDefault
    Me.Visible = False
    
    Unload Me

End Sub

Private Sub Form_Load()
    'Dim lR As Long
    'lR = SetTopMostWindow(Me.hwnd, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (killProcessing = False) Then
        killProcessing = True
        Exit Sub
    End If
End Sub

