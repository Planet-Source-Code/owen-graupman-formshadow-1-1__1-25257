VERSION 5.00
Object = "{60301B94-3C1C-4A47-9916-2BD5E9491FFE}#1.0#0"; "Shadow.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.PictureBox shpTop 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   0
      Width           =   3795
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   2085
      End
   End
   Begin ctlShadow.goeShadow goeShadow1 
      Left            =   60
      Top             =   390
      _extentx        =   741
      _extenty        =   741
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private movenow As Boolean
Private TX As Single
Private TY As Single

Private Sub Form_Deactivate()
    Call goeShadow1.Refresh
End Sub

Private Sub Form_GotFocus()
    Call goeShadow1.Refresh
End Sub

Private Sub Form_LostFocus()
    Call goeShadow1.Refresh
End Sub

Private Sub Form_Resize()
    shpTop.Left = 0
    shpTop.Top = 0
    shpTop.Width = Me.ScaleWidth
    Label1.Width = Me.ScaleWidth - 120
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    movenow = True
    TX = X
    TY = Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If movenow Then
        Me.Top = Me.Top + Y - TY
        Me.Left = Me.Left + X - TX
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    movenow = False
End Sub

Private Sub shpTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    movenow = True
    TX = X
    TY = Y
End Sub

Private Sub shpTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If movenow Then
        Me.Top = Me.Top + Y - TY
        Me.Left = Me.Left + X - TX
    End If
End Sub

Private Sub shpTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    movenow = False
End Sub
