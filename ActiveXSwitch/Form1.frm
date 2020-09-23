VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   2100
   ClientLeft      =   6435
   ClientTop       =   4185
   ClientWidth     =   1935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   1935
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1425
      Top             =   1275
   End
   Begin Project1.ProgressMeter ProgressMeter1 
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Top             =   1650
      Width           =   1050
      _extentx        =   1852
      _extenty        =   529
      borderstyle     =   5
      picture         =   "Form1.frx":030A
   End
   Begin Project1.LED LED2 
      Height          =   135
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   375
      _extentx        =   661
      _extenty        =   238
   End
   Begin Project1.LED LED1 
      Height          =   135
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   375
      _extentx        =   661
      _extenty        =   238
      borderstyle     =   5
   End
   Begin Project1.CheckMark CheckMark1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Check Box"
      Top             =   1320
      Width           =   1215
      _extentx        =   2143
      _extenty        =   450
      caption         =   "On/Off"
   End
   Begin Project1.ToggleLight ToggleLight1 
      Height          =   915
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   420
      _extentx        =   741
      _extenty        =   1614
      borderstyle     =   4
      Object.tooltiptext     =   "Switch"
   End
   Begin Project1.TrayControl TrayControl1 
      Left            =   1440
      Top             =   600
      _extentx        =   794
      _extenty        =   794
   End
   Begin Project1.ToggleSwitch ToggleSwitch1 
      Height          =   915
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Toggle Switch"
      Top             =   360
      Width           =   420
      _extentx        =   741
      _extenty        =   1614
      borderstyle     =   4
      Object.tooltiptext     =   "Switch"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tray"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuMenuRestore 
         Caption         =   "Restore"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnInTray As Boolean

Const WM_RBUTTONUP = &H205
Const WM_LBUTTONDBLCLK = &H203

Private Sub CheckMark1_Click()

    If CheckMark1.Checked Then
        
        Timer1.Enabled = True
    
    Else
    
        Timer1.Enabled = False
    
    End If

End Sub

Private Sub Command1_Click()
    
    TrayControl1.SendToTray

End Sub

Private Sub Form_Load()

'    ToggleSwitch1.Value = True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then PopupMenu mnuMenu

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Static Message As Long
    Message = X / Screen.TwipsPerPixelX
    
    Select Case Message
        
        Case WM_LBUTTONDBLCLK
            
            Call mnuRestore_Click
        
        Case WM_RBUTTONUP:
            
            Me.PopupMenu mnuMenu
    
    End Select

End Sub

Private Sub mnuMenuRestore_Click()
    
    TrayControl1.RestoreFromTray

End Sub

Private Sub mnuRestore_Click()

    blnInTray = False
    TrayControl1.RestoreFromTray
    
End Sub

Private Sub Timer1_Timer()
If ProgressMeter1.Value = 100 Then ProgressMeter1.Value = 0
ProgressMeter1.Value = ProgressMeter1.Value + 1
End Sub

Private Sub ToggleLight1_Click()
    If ToggleLight1.Value = True Then
    
        LED1.BackColor = &HFF00&
    
    Else
    
        LED1.BackColor = &H8000&
        
    End If
End Sub

Private Sub ToggleSwitch1_Click()
If ToggleSwitch1.Value = False Then LED2.BackColor = &H8000& Else LED2.BackColor = &HFF00&
End Sub

Private Sub ToggleSwitch1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ToggleSwitch1.Value = False Then LED2.BackColor = &H8000& Else LED2.BackColor = &HFF00&

End Sub
