VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "Picclp32.ocx"
Begin VB.UserControl ToggleSwitch 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   ClipBehavior    =   0  'None
   ControlContainer=   -1  'True
   PaletteMode     =   4  'None
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   115
   ToolboxBitmap   =   "ToggleSwitch.ctx":0000
   Begin PicClip.PictureClip PictureClip1 
      Left            =   240
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1402
      _Version        =   393216
      Cols            =   2
      Picture         =   "ToggleSwitch.ctx":0312
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   795
      Left            =   60
      Stretch         =   -1  'True
      ToolTipText     =   "Switch"
      Top             =   60
      Width           =   300
   End
End
Attribute VB_Name = "ToggleSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum EBorderStyle 'Controls Border Styles
    None = 0
    Inset = 1
    Raised = 2
    FixedSingle = 3
    Flat1 = 4
    Flat2 = 5
End Enum

'Property Variables:
Dim m_Value As Boolean
Dim BSBorderStyle As EBorderStyle 'Controls BorderStyle

'Event Declarations:
Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const m_def_Value = 0

Private Sub UserControl_Initialize()

    ' Place picture in Image1
    Image1.Picture = PictureClip1.GraphicCell(0)

End Sub

Private Sub UserControl_Resize()
    
    With UserControl

        .Height = 920
        .Width = 420

    End With

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    BorderStyle = PropBag.ReadProperty("BorderStyle", EBorderStyle.Inset)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    Image1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderStyle", BSBorderStyle, EBorderStyle.Inset)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("ToolTipText", Image1.ToolTipText, "")
    
End Sub

Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Switch Value"
Attribute Value.VB_MemberFlags = "400"
    
    Value = m_Value
    
    If Image1.Picture() = PictureClip1.GraphicCell(1) Then Value = True Else Value = False

End Property

Public Property Let Value(ByVal New_Value As Boolean)
    
    If Ambient.UserMode = False Then Err.Raise 387
    
    m_Value = New_Value
    PropertyChanged "Value"
    
    If m_Value = True Then Image1.Picture() = PictureClip1.GraphicCell(1) Else Image1.Picture() = PictureClip1.GraphicCell(0)

End Property

Private Sub UserControl_InitProperties()
    
    m_Value = m_def_Value

End Sub

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    
    ToolTipText = Image1.ToolTipText

End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    
    Image1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
    
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseDown(Button, Shift, X, Y)

    If Y < 34 Then ' Set center of switch change picture accordingly

        Image1.Picture = PictureClip1.GraphicCell(1)
        m_Value = True

    Else

        Image1.Picture = PictureClip1.GraphicCell(0)
        m_Value = False

    End If

End Sub



Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Set MouseMove Properties
    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Set MouseUp Properties
    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Click()
    
    ' Set Click Properties
    RaiseEvent Click

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


Public Property Get BorderStyle() As EBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    
    BorderStyle = BSBorderStyle  'Change The Value

End Property

Public Property Let BorderStyle(ByVal NewStyle As EBorderStyle)
   
    BSBorderStyle = NewStyle 'Change The BorderStyle
    DrawBorder 'Redraw The Border

End Property
