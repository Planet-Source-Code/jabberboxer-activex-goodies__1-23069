VERSION 5.00
Begin VB.UserControl CheckMark 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1410
   ScaleHeight     =   300
   ScaleWidth      =   1410
   ToolboxBitmap   =   "CheckMark.ctx":0000
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Box"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   15
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   270
      Index           =   1
      Left            =   0
      Picture         =   "CheckMark.ctx":0312
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "CheckMark.ctx":06A9
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "CheckMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum iChecked
    NotMarked = 0
    Marked = 1
End Enum
    
Event Click()

Const m_def_Checked = True
Dim m_Checked As iChecked

Public Function CheckValue()
    
    Select Case m_Checked
        
        Case 0
            
            Image1(1).Visible = False
        
        Case 1
        
            Image1(1).Visible = True
    
    End Select
    
End Function

Private Sub Image1_Click(Index As Integer)

    RaiseEvent Click
    
    Select Case Index
    
        Case 0
            Checked = Marked
            
        Case 1
            Checked = NotMarked
            
    End Select
    
End Sub

Private Sub Label1_Click()
     RaiseEvent Click
'    Checked = Not Checked
    
    If Checked = Marked Then
        Checked = NotMarked
        Image1(1).Visible = False
        
    Else
        Checked = Marked
        Image1(1).Visible = True
        
    End If
     
End Sub

Private Sub UserControl_Resize()

    With UserControl
        .Height = 255
        Label1.Width = .Width - Label1.Left
    End With
    
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."

    Caption = Label1.Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)
    
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"

End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Checked = PropBag.ReadProperty("Checked", iChecked.NotMarked)
    Label1.Caption = PropBag.ReadProperty("Caption", "Check Box")
'    m_Checked = PropBag.ReadProperty("Checked", m_def_Checked)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Checked", m_Checked, iChecked.NotMarked)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Check Box")
'    Call PropBag.WriteProperty("Checked", m_Checked, m_def_Checked)
    
End Sub

Public Property Get Checked() As iChecked

    Checked = m_Checked

End Property

Public Property Let Checked(ByVal New_Value As iChecked)

    m_Checked = New_Value
    CheckValue
    
End Property



