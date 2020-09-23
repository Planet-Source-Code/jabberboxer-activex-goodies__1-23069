VERSION 5.00
Begin VB.UserControl LED 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2340
   ScaleHeight     =   86
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   156
End
Attribute VB_Name = "LED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum iBorderStyle 'Controls Border Styles
    None = 0
    Inset = 1
    Raised = 2
    FixedSingle = 3
    Flat1 = 4
    Flat2 = 5
End Enum

Dim BSBorderStyle As iBorderStyle 'Controls BorderStyle

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000&)
    BorderStyle = PropBag.ReadProperty("BorderStyle", iBorderStyle.Flat1)
    
End Sub

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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    DrawBorder
End Property

Public Property Get BorderStyle() As EBorderStyle
    
    BorderStyle = BSBorderStyle  'Change The Value

End Property

Public Property Let BorderStyle(ByVal NewStyle As EBorderStyle)
   
    BSBorderStyle = NewStyle 'Change The BorderStyle
    DrawBorder 'Redraw The Border

End Property

Private Sub UserControl_Resize()
    
    DrawBorder
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000&)
    Call PropBag.WriteProperty("BorderStyle", BSBorderStyle, iBorderStyle.Flat1)
    
End Sub

