VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About OziExplorer Batch Map Convert"
   ClientHeight    =   1410
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      Caption         =   "http://pointybeard.com"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblApiVersion 
      Caption         =   "ozi api version"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblOziVersion 
      Caption         =   "ozi version"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Written by Alistair Kearney"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim apiVersion As Double
    apiVersion = modLibrary.objOziAPI.apiVersion
    
    Dim exeVersion As Double
    exeVersion = modLibrary.objOziAPI.exeVersion
       
    
    Me.lblApiVersion.Caption = "OziExplorer API Version: "
    Me.lblOziVersion = "OziExplorer Version: "
    
    If (apiVersion = -1) Then
        Me.lblApiVersion.Caption = Me.lblApiVersion.Caption + "Unknown"
    Else
        Me.lblApiVersion.Caption = Me.lblApiVersion.Caption + CStr(apiVersion)
    End If
    
    If (exeVersion = -1) Then
        Me.lblOziVersion.Caption = Me.lblOziVersion.Caption + "Unknown"
    Else
        Me.lblOziVersion.Caption = Me.lblOziVersion.Caption + CStr(exeVersion)
    End If
End Sub

