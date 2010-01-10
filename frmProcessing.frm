VERSION 5.00
Begin VB.Form frmProcessing 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Processing Maps"
   ClientHeight    =   1080
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin OziExpBatchMapConv.ucProgressBar pbProcessing 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      _extentx        =   7011
      _extenty        =   661
   End
   Begin VB.Label lblProcessing 
      Caption         =   "processing ..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "frmProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim killProcessing As Boolean

Public Sub beginProcessing(outputType As String, Optional destinationDirectory As String = vbNullString)

    killProcessing = False

    Me.MousePointer = vbHourglass
    Me.pbProcessing.Value = 0
    
    Dim ii As Integer
    Dim file As String
    Dim retVal
    Dim destinationFile As String
    Dim progressbarStep As Double
    
    progressbarStep = 100 / (UBound(sMapFileArray) * 2)
    'progressbarStep = 1
    'Me.pbProcessing.Max = (UBound(sMapFileArray) * 2) + 1
    
    For ii = 0 To UBound(modLibrary.sMapFileArray) - 1
    
        If killProcessing = True Then Exit For
    
        file = modLibrary.sMapFileArray(ii)
        Me.lblProcessing.Caption = "Processing map " & file & " (" & (ii + 1) & " of " & UBound(modLibrary.sMapFileArray) & ")"
        Me.pbProcessing.Value = Me.pbProcessing.Value + progressbarStep '(100 / UBound(sMapFileArray)) * (ii + 1)
        
        DoEvents
        
        If (oziLoadMap(file) = -1) Then
            MsgBox "Error loading map " & file, vbOKOnly + vbCritical, "Error"
            
        Else
            Me.pbProcessing.Value = Me.pbProcessing.Value + progressbarStep
            DoEvents
            
            If (destinationDirectory = vbNullString) Then
                destinationFile = Replace(file, getFileExtension(file), outputType)
            Else
                destinationFile = destinationDirectory & Replace(getFileNameFromPath(file), getFileExtension(file), outputType)
            End If

            retVal = oziSaveMapImage(destinationFile) 'Replace(file, getFileExtension(file), outputType))
        End If
        
    Next

    Me.MousePointer = vbDefault
    Me.Visible = False
    
    Unload Me

End Sub

Private Sub Form_Load()
    Dim lR As Long
    lR = SetTopMostWindow(Me.hwnd, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (killProcessing = False) Then
        killProcessing = True
        Exit Sub
    End If
    frmMain.SetFocus
End Sub
