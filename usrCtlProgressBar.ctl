VERSION 5.00
Begin VB.UserControl ucProgressBar 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   277
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.PictureBox picProgress 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   0
         TabIndex        =   1
         Top             =   0
         Width           =   0
      End
   End
End
Attribute VB_Name = "ucProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================
' CUSTOM PROGRESS BAR
'===========================
'
'Author: Alistair Kearney
'E-Mail: alistair@pointybeard.com
'Website: http://pointybeard.com
'Desc: A custom progress bar to replace the standard VB one
'Usage: Exactly like the normal progress bar
'
'Copyright Alistair Kearney  2002
'===========================


Public Enum Border
    [None] = 0
    [Fixed Single] = 1
End Enum
    
Public Enum Appearence
    [Flat] = 0
    [3D] = 1
End Enum

Public Enum Style
    [Standard] = 0
    [Graphical] = 1
End Enum

Private progressPercent As Integer
Private maxVal As Integer
'Default Property Values:
'Const m_def_Style = 0
Const m_def_Value = 1
Const m_def_Max = 100
Const m_def_Min = 1
'Property Variables:
'Dim m_Style As Integer
Dim m_Value As Integer
Dim m_Max As Integer
Dim m_Min As Integer

Public Sub Reset()
    picProgress.Width = 1
    picProgress.Left = Picture1.Left
    progressPercent = 1
    maxVal = 100
    lblPercentage.Caption = "0%"
End Sub

Private Sub UserControl_Resize()

    Picture1.Width = UserControl.ScaleWidth
    Picture1.Height = UserControl.ScaleHeight
    picProgress.Height = Picture1.Height - 1
    'UserControl.Height = 375
    If UserControl.Width < 1335 Then UserControl.Width = 1335
    
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,BorderStyle
Public Property Get BorderStyle() As Border
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Picture1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Border)
    Picture1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get Value() As Integer
    Value = m_Value
    
    picProgress.Width = Int(Picture1.ScaleWidth * (progressPercent / m_Max))
    picProgress.Left = Picture1.Left
    DoEvents
    
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    progressPercent = m_Value
    If progressPercent < m_Min Then progressPercent = m_Min
    If progressPercent > m_Max Then progressPercent = m_Max

    picProgress.Width = Int(Picture1.ScaleWidth * (progressPercent / m_Max))
    picProgress.Left = Picture1.Left
 
    DoEvents
    
    PropertyChanged "Value"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,100
Public Property Get Max() As Integer
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get Min() As Integer
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    m_Min = New_Min
    PropertyChanged "Min"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_Max = m_def_Max
    m_Min = m_def_Min
'    m_Style = m_def_Style
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Picture1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
'    Picture1.Appearance = PropBag.ReadProperty("AppearanceStyle", 1)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    Picture1.Appearance = PropBag.ReadProperty("Appearance", 1)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    picProgress.BackColor = PropBag.ReadProperty("FillColor", &HFF0000)
    Picture1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
'    m_Style = PropBag.ReadProperty("Style", m_def_Style)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderStyle", Picture1.BorderStyle, 1)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
'    Call PropBag.WriteProperty("AppearanceStyle", Picture1.Appearance, 1)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Appearance", Picture1.Appearance, 1)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("FillColor", picProgress.BackColor, &HFF0000)
    Call PropBag.WriteProperty("BackColor", Picture1.BackColor, &H8000000F)
'    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,Appearance
Public Property Get Appearance() As Appearence
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = Picture1.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Appearence)
    Picture1.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picProgress,picProgress,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = picProgress.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set picProgress.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picProgress,picProgress,-1,BackColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    FillColor = picProgress.BackColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    picProgress.BackColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Picture1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Picture1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=7,0,0,0
'Public Property Get Style() As Style
'    Style = m_Style
'End Property
'
'Public Property Let Style(ByVal New_Style As Style)
'    m_Style = New_Style
'
'    If m_Style = 0 Then Set picProgress.Picture = Nothing
'
'    PropertyChanged "Style"
'End Property
'
