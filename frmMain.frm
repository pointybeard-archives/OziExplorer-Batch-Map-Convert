VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OziExplorer Batch Map Convert"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRecurse 
      Caption         =   "Recurse into sub-directories"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Value           =   1  'Checked
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   5160
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Start"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   5775
   End
   Begin VB.CommandButton cmdBrowseForTargetDirectory 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtTargetFolder 
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   5535
      Begin VB.OptionButton optPNG 
         Caption         =   "PNG"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optBMP 
         Caption         =   "BMP"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Output Format: "
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.TextBox txtMapFolder 
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   4335
   End
   Begin VB.CommandButton cmdBrowseForMaps 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Optional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblStep1 
      Caption         =   $"frmMain.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   5535
   End
   Begin VB.Label lblStep1 
      Caption         =   "Begin by selecting a folder that contains OziExplorer map files."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5535
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu menuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu menuSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Quit OziExplorer Toolkit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowseForMaps_Click()
    Dim x As Integer
    x = 3
    
    With cdMain
        .DialogTitle = "Choose Map Folder"
        .filename = "*.*"
        .Flags = cdlOFNHideReadOnly + cdlOFNNoChangeDir + cdlOFNExplorer + cdlOFNNoValidate
        .CancelError = True

        On Error Resume Next

        .Action = 1

        If LenB(.FileTitle$) <> 0 Then x = Len(.FileTitle)

        If Err = 0 Then
            ChDrive cdMain.filename
            
            Dim sPath As String
            sPath = Left(cdMain.filename, Len(cdMain.filename) - x)
            
            Dim fso As New FileSystemObject
            Dim fld As Folder
            
            ReDim sMapFileArray(0) As String
            
            Me.MousePointer = vbHourglass
            
            If FindFilesByExtension(fso, fld, sPath, "MAP", modLibrary.sMapFileArray, Me.chkRecurse.Value = vbChecked) = 0 Then
                MsgBox "No map files found at '" & sPath & "'", vbExclamation + vbOKOnly, "Error"
            
            Else
                Me.txtMapFolder.Text = sPath
                MsgBox "Found " & UBound(modLibrary.sMapFileArray) & " MAP files.", vbOKOnly + vbInformation
            End If
            
            Me.MousePointer = vbDefault

        End If
    End With
    
End Sub

Private Sub cmdBrowseForTargetDirectory_Click()
    Dim x As Integer
    x = 3
    
    With cdMain
        .DialogTitle = "Choose Map Folder"
        .filename = "*.*"
        .Flags = cdlOFNHideReadOnly + cdlOFNNoChangeDir + cdlOFNExplorer + cdlOFNNoValidate
        .CancelError = True

        On Error Resume Next

        .Action = 1

        If LenB(.FileTitle$) <> 0 Then x = Len(.FileTitle)

        If Err = 0 Then
            ChDrive cdMain.filename
            Dim fso As New FileSystemObject
            Dim fld As Folder

            Me.txtTargetFolder.Text = Left(cdMain.filename, Len(cdMain.filename) - x)

        End If
    End With

End Sub

Private Sub cmdProcess_Click()
    If objOziAPI.askForOzi("You must start OziExplorer before processing any maps.") = False Then
        Exit Sub
    End If
    
    frmMain.Enabled = False
    
    frmProcessing.Show
    
    Dim extension As String
    
    extension = "png"
    If (Me.optBMP.Value = True) Then
        extension = "bmp"
    End If
    frmProcessing.beginProcessing extension, Me.txtTargetFolder.Text
    
    frmMain.Enabled = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub menuAbout_Click()
    frmAbout.Show
End Sub

Private Sub menuExit_Click()
    Unload Me
    End
End Sub

Private Sub txtMapFolder_Change()
    If LenB(Me.txtMapFolder.Text) > 0 Then
        Me.cmdProcess.Enabled = True
    End If
End Sub

