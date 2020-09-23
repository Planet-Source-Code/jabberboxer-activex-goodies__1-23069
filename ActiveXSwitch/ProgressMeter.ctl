VERSION 5.00
Begin VB.UserControl ProgressMeter 
   AutoRedraw      =   -1  'True
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   46
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   128
   ToolboxBitmap   =   "ProgressMeter.ctx":0000
   Begin VB.PictureBox picDEST 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   50
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   50
      Width           =   960
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00FFC0FF&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   0
      Max             =   100
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picSRC0 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   2520
      Picture         =   "ProgressMeter.ctx":0312
      ScaleHeight     =   3135
      ScaleWidth      =   1935
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "ProgressMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum xBorderStyle 'Controls Border Styles
    None = 0
    Inset = 1
    Raised = 2
    FixedSingle = 3
    Flat1 = 4
    Flat2 = 5
End Enum

Dim BSBorderStyle As xBorderStyle 'Controls BorderStyle

' Coding credit for bitblt function code goes to "DosAscii"

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020     'Copies the source bitmap to destination bitmap.
Private Const SRCAND = &H8800C6      'Combines pixels of the destination with source bitmap
                                    'using the Boolean AND operator.
Private Const SRCINVERT = &H660046   'Combines pixels of the destination with source bitmap
                                    'using the Boolean XOR operator.
Private Const SRCPAINT = &HEE0086    'Combines pixels of the destination with source bitmap
                                    'using the Boolean OR operator.
Private Const SRCERASE = &H4400328   'Inverts the destination bitmap and then combines the
                                    'results with the source bitmap using the Boolean AND
                                    'operator.
Private Const WHITENESS = &HFF0062   'Turns all output white.
Private Const BLACKNESS = &H42       'Turn output black.
 'This foreces all varibles to be declared now.
Dim i As Integer

Public Function DrawBorder()
    'bit of credit to "Daniel Davies"
    Cls
    
    Select Case BSBorderStyle 'Draw The Border (If Any)
    
        Case 1 'Inset, We Need To Draw Several lines around the edge (8 to be exact)
            Line (0, 0)-(ScaleWidth, 0), vb3DDKShadow 'Darkest Shadow
            Line (1, 1)-(ScaleWidth - 1, 1), vb3DShadow 'Dark Shadow
            Line (0, 0)-(0, ScaleHeight), vb3DDKShadow 'Darkest Shadow
            Line (1, 1)-(1, ScaleHeight - 1), vb3DShadow 'Dark Shadow
            Line (ScaleWidth - 1, 1)-(ScaleWidth - 1, ScaleHeight), vb3DHighlight 'Lightest Shadow
            Line (ScaleWidth - 2, 2)-(ScaleWidth - 2, ScaleHeight - 2), vb3DLight 'Light Shadow
            Line (ScaleWidth - 1, ScaleHeight - 1)-(0, ScaleHeight - 1), vb3DHighlight 'Lightest Shadow
            Line (ScaleWidth - 2, ScaleHeight - 2)-(1, ScaleHeight - 2), vb3DLight 'Light Shadow
            Refresh
        
        Case 2 'Raised, Same As Inset (But Colors Are Inverted)
            Line (0, 0)-(ScaleWidth, 0), vb3DHighlight 'Lightest Shadow
            Line (1, 1)-(ScaleWidth - 1, 1), vb3DLight 'Light Shadow
            Line (0, 0)-(0, ScaleHeight), vb3DHighlight 'Lightest Shadow
            Line (1, 1)-(1, ScaleHeight - 1), vb3DLight 'Light Shadow
            Line (ScaleWidth - 1, 1)-(ScaleWidth - 1, ScaleHeight), vb3DDKShadow 'Darkest Shadow
            Line (ScaleWidth - 2, 2)-(ScaleWidth - 2, ScaleHeight - 2), vb3DShadow 'Dark Shadow
            Line (ScaleWidth - 1, ScaleHeight - 1)-(0, ScaleHeight - 1), vb3DDKShadow 'Darkest Shadow
            Line (ScaleWidth - 2, ScaleHeight - 2)-(1, ScaleHeight - 2), vb3DShadow 'Dark Shadow
            Refresh
        
        Case 3 'Fixed Single (Black 1 Pixel Width Border)
            Line (0, 0)-(ScaleWidth, 0), vbBlack
            Line (0, 0)-(0, ScaleHeight), vbBlack
            Line (ScaleWidth - 1, 1)-(ScaleWidth - 1, ScaleHeight), vbBlack
            Line (ScaleWidth - 1, ScaleHeight - 1)-(0, ScaleHeight - 1), vbBlack
            Refresh
            
        Case 4 'Flat1 (Raised Then Inset)
            Line (0, 0)-(ScaleWidth, 0), vb3DHighlight 'Lightest Shadow
            Line (2, 2)-(ScaleWidth - 2, 2), vb3DDKShadow 'Darkest Shadow
            Line (0, 0)-(0, ScaleHeight), vb3DHighlight 'Lightest Shadow
            Line (2, 2)-(2, ScaleHeight - 2), vb3DDKShadow 'Darkest Shadow
            Line (ScaleWidth - 1, 1)-(ScaleWidth - 1, ScaleHeight), vb3DDKShadow 'Darkest Shadow
            Line (ScaleWidth - 3, 3)-(ScaleWidth - 3, ScaleHeight - 3), vb3DHighlight 'Lightest Shadow
            Line (ScaleWidth - 1, ScaleHeight - 1)-(0, ScaleHeight - 1), vb3DDKShadow 'Darkest Shadow
            Line (ScaleWidth - 3, ScaleHeight - 3)-(1, ScaleHeight - 3), vb3DHighlight 'Lightest Shadow
            Refresh
        
        Case 5 'Flat2 (Inset Then Raised)
            Line (0, 0)-(ScaleWidth, 0), vb3DDKShadow 'Darkest Shadow
            Line (2, 2)-(ScaleWidth - 2, 2), vb3DHighlight 'Lightest Shadow
            Line (0, 0)-(0, ScaleHeight), vb3DDKShadow 'Darkest Shadow
            Line (2, 2)-(2, ScaleHeight - 2), vb3DHighlight 'Lightest Shadow
            Line (ScaleWidth - 1, 1)-(ScaleWidth - 1, ScaleHeight), vb3DHighlight 'Lightest Shadow
            Line (ScaleWidth - 3, 3)-(ScaleWidth - 3, ScaleHeight - 3), vb3DDKShadow 'Darkest Shadow
            Line (ScaleWidth - 1, ScaleHeight - 1)-(0, ScaleHeight - 1), vb3DHighlight 'Lightest Shadow
            Line (ScaleWidth - 3, ScaleHeight - 3)-(1, ScaleHeight - 3), vb3DDKShadow 'Darkest Shadow
            Refresh
    
    End Select

End Function

Private Sub HScroll1_Change()

    SwapImages

End Sub

Private Sub HScroll1_Scroll()

    SwapImages

End Sub

Private Sub UserControl_Initialize()
    
    Call BitBlt(picDEST.hDC, 0, 0, 63, 13, picSRC0.hDC, 1, 0, SRCAND)
'    picDEST.BorderStyle = 1

End Sub

Function SwapImages()

    Dim OneTwentyEighth As Long
    OneTwentyEighth = 100 / 28
    
    picDEST.Picture = Nothing
    i = HScroll1.Value / OneTwentyEighth
    
    Dim iY As Integer
        
    If i < 14 Then
        iY = 1
        
    Else
        iY = 66
        i = i - 14
    End If
    
    Call BitBlt(picDEST.hDC, 0, 0, 63, 13, picSRC0.hDC, iY, i * 15, SRCAND)
'    picDEST.BorderStyle = 1

    Label1.Caption = HScroll1.Value & "%"
    
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
'Public Property Get BorderStyle() As Integer
'    BorderStyle = UserControl.BorderStyle
'End Property
'
'Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
'    UserControl.BorderStyle() = New_BorderStyle
'    PropertyChanged "BorderStyle"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
 
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picSRC0,picSRC0,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = picSRC0.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set picSRC0.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=HScroll1,HScroll1,-1,Value
Public Property Get Value() As Integer
Attribute Value.VB_Description = "Returns/sets the value of an object."
    Value = HScroll1.Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    HScroll1.Value() = New_Value
    PropertyChanged "Value"

End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BorderStyle = PropBag.ReadProperty("BorderStyle", xBorderStyle.Flat1)

'    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &HFFC0FF)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    HScroll1.Value = PropBag.ReadProperty("Value", 0)
End Sub

Private Sub UserControl_Resize()
    
    With UserControl
        .Height = 305
        .Width = 1045
    End With
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", BSBorderStyle, xBorderStyle.Flat1)

'    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &HFFC0FF)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Value", HScroll1.Value, 0)
End Sub

Public Property Get BorderStyle() As EBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    
    BorderStyle = BSBorderStyle  'Change The Value

End Property

Public Property Let BorderStyle(ByVal NewStyle As EBorderStyle)
   
    BSBorderStyle = NewStyle 'Change The BorderStyle
    DrawBorder 'Redraw The Border
    SwapImages
End Property
