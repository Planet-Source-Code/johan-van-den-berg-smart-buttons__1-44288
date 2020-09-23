VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Smart Buttons"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label cmdExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FEE2E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2610
      TabIndex        =   1
      Top             =   450
      Width           =   1245
   End
   Begin VB.Label cmdClick 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FEE2E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click Me"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   450
      Width           =   1245
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MouseButDown As Boolean 'This value holds whether a mouse button is being held down

Private Sub cmdExit_Click()

    End

End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'The MouseDown event triggers everytime you press a mouse button down while the mouse pointer is
'over the control that owns this event.

    MouseButDown = True
    cmdExit.BackColor = &HBCF84E
    cmdExit.ForeColor = vbBlack

End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'The MouseUp event triggers everytime you release a mouse button over the control that
'owns this event
'From this you can deduce that a click event triggers after a MouseUp directly follows
'a MouseDown event on the same control. MouseUp occurs BEFORE Click

    UpdateButton 1, cmdExit 'Call the UpdateButton sub in modButton
    MouseButDown = False

End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This event triggers everytime you move the mouse pointer over the control that owns this event

    If MouseButDown = False Then
    
        UpdateButton 0, Me
        UpdateButton 1, cmdExit
        
    End If

End Sub

Private Sub Form_Load()

    MouseButDown = False
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UpdateButton 0, Me
    MouseButDown = False

End Sub

Private Sub cmdClick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'The MouseDown event triggers everytime you press a mouse button down while the mouse pointer is
'over the control that owns this event.

    MouseButDown = True
    cmdClick.BackColor = &HBCF84E
    cmdClick.ForeColor = vbBlack

End Sub

Private Sub cmdClick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'The MouseUp event triggers everytime you release a mouse button over the control that
'owns this event
'From this you can deduce that a click event triggers after a MouseUp directly follows
'a MouseDown event on the same control.

    UpdateButton 1, cmdClick 'Call the UpdateButton sub in modButton
    MouseButDown = False

End Sub

Private Sub cmdClick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This event triggers everytime you move the mouse pointer over the control that owns this event

    If MouseButDown = False Then
    
        UpdateButton 0, Me
        UpdateButton 1, cmdClick
        
    End If

End Sub

