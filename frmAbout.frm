VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About OziExplorer Batch Map Convert"
   ClientHeight    =   2640
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      Caption         =   "http://pointybeard.com"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright 2010"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label lblApiVersion 
      Caption         =   "ozi api version"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label lblOziVersion 
      Caption         =   "ozi version"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Written by Alistair Kearney"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Dim OziAPI As New classOziAPI
    
    Me.lblApiVersion.Caption = "OziExplorer API Version: " + modLibrary.objOziAPI.OziApiVersion
    Me.lblOziVersion = "OziExplorer Version: " + modLibrary.objOziAPI.OziVersion
    
End Sub
