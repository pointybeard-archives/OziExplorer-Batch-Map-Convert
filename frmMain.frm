VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OziExplorer Batch Map Convert"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   480
      Top             =   3600
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
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtMapFolder 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   4455
   End
   Begin VB.CommandButton cmdBrowseForMaps 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.CheckBox chkShowAdvanced 
      Caption         =   "Show Advanced Options"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   960
      Width           =   2295
   End
   Begin VB.Frame frmAdvanced 
      Caption         =   "Advanced Options"
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkAskForName 
         Caption         =   "Ask for map name"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3000
         TabIndex        =   15
         Top             =   720
         Width           =   2655
      End
      Begin VB.CheckBox chkUpdateMAPFile 
         Caption         =   "Regenerate MAP calibration file."
         Height          =   495
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Note, this will OVERWRITE the existing .MAP file"
         Top             =   720
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   5535
         Begin VB.OptionButton optBMP 
            Caption         =   "BMP"
            Height          =   255
            Left            =   1560
            TabIndex        =   10
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optPNG 
            Caption         =   "PNG"
            Height          =   255
            Left            =   2640
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Output Format: "
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.TextBox txtTargetFolder 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   4335
      End
      Begin VB.CommandButton cmdBrowseForTargetDirectory 
         Caption         =   "Browse"
         Height          =   375
         Left            =   4680
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblStep1 
         Caption         =   "Choose the target location for the newly converted map files. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   $"frmMain.frx":0000
         Top             =   1320
         Width           =   5535
      End
   End
   Begin VB.CheckBox chkRecurse 
      Caption         =   "Recurse into sub-directories"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Value           =   1  'Checked
      Width           =   2415
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
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkShowAdvanced_Click()
    Me.Height = 5220
    Me.chkShowAdvanced.Enabled = False
    Me.cmdProcess.Top = 264
    Me.frmAdvanced.Visible = True
End Sub


Private Sub chkUpdateMAPFile_Click()
    If (Me.chkUpdateMAPFile.Value = vbChecked) Then
        MsgBox "Warning, This will overwrite the existing .MAP file.", vbOKOnly + vbInformation
        Me.chkAskForName.Enabled = True
        Exit Sub
    End If
    
    Me.chkAskForName.Enabled = False
End Sub

Private Sub cmdBrowseForMaps_Click()
    Dim x As Integer
    x = 3
    
    With cdMain
        .DialogTitle = "Choose Map Folder"
        .filename = "*.*"
        .FLAGS = cdlOFNHideReadOnly + cdlOFNNoChangeDir + cdlOFNExplorer + cdlOFNNoValidate
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
        .FLAGS = cdlOFNHideReadOnly + cdlOFNNoChangeDir + cdlOFNExplorer + cdlOFNNoValidate
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
    frmProcessing.beginProcessing extension, Me.txtTargetFolder.Text, Me.chkUpdateMAPFile.Value, Me.chkAskForName.Value
    
    frmMain.Enabled = True
    
End Sub

Private Sub Form_Load()
    If objOziAPI.askForOzi("OziExplorer must be running") = False Then
        Unload Me
        End
    End If
    
    If objOziAPI.askForOziAPI() = False Then
        Unload Me
        End
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    objOziAPI.Destroy
    Set objOziAPI = Nothing
    End
End Sub

Private Sub menuAbout_Click()
    frmAbout.Show
End Sub

Private Sub menuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuHelp_Click()
    OpenBrowser ("http://github.com/pointybeard/OziExplorer-Batch-Map-Convert")
End Sub

Private Sub txtMapFolder_Change()
    If LenB(Me.txtMapFolder.Text) > 0 Then
        Me.cmdProcess.Enabled = True
    End If
End Sub
