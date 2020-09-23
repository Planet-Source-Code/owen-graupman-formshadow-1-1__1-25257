VERSION 5.00
Object = "{65AA350C-1DAD-45E0-BAB6-71B5A1A2AD9F}#1.0#0"; "MacroTimer.ocx"
Begin VB.UserControl goeShadow 
   Appearance      =   0  'Flat
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1575
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   1575
   ToolboxBitmap   =   "goeShadow.ctx":0000
   Begin MacroTimer.NSTimer tmrShadow 
      Left            =   1080
      Top             =   0
      _ExtentX        =   953
      _ExtentY        =   953
      Interval        =   1
      Milliseconds    =   1
   End
   Begin VB.PictureBox picShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "goeShadow.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "goeShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Control:   Shadow.OCX
'Purpose:   Places a shadow at the bottom and right edges of the containing form
'Author:    Owen Graupman
'Created:   07/19/2001
'Usage:     Drop onto a form and it will draw a shadow automatically
'Notes:     Only works when contained on a form. Does not work when placed
'               in a container control (such as a picture box, etc.)
'           There are no properties to set. Just drop and go.
'           Uses alphablending.dll, and NSTimer.ocx, both available
'               on Planet Source Code.
'           Inspired by the dShadow control created by ^dark^, which was
'               in turn inspired (I believe) by the b2 shell replacement.
'           If you have a hard time seeing the shadow, move the form over a
'           light background (I.E., white).
'           This code requires Win98 or later (98,ME,NT,2000,XP) to function.

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function AlphaBlending Lib "Alphablending.dll" (ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal AlphaSource As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private rectRight As RECT   'Used to redraw areas that have been shaded over
Private rectBottom As RECT
Private sngX As Single      'Left and right coordinates used to prevent
Private sngY As Single      '   multiple redraws


Public Sub Refresh()
'Purpose:   Redraws the shadow
'Author:    Owen Graupman
'Created:   07/19/2001
'Inputs:    None
'Outputs:   None
'Returns:   None

    On Error Resume Next
    
    Dim lngCount As Long    'Loop counter
    Dim lngHDc As Long      'The main screen's hdc
    
    'First, erase the existing shadow
    RedrawWindow 0&, rectRight, 0&, 135
    RedrawWindow 0&, rectBottom, 0&, 135
    DoEvents    'Only need one DoEvents, which forces the screen to be cleaned up
                '   from the previous draw before drawing again (prevents
                '   screen artifacts).
    
    'Fetch the main window device context which we are going to draw on.
    lngHDc = GetDC(0)
    
    'Now draw the shadow
    For lngCount = 1 To 6   'Loop through 6 times, each time the shadow gets lighter
        'Bottom
        Call AlphaBlending(lngHDc, (UserControl.Parent.Left / Screen.TwipsPerPixelX) + (lngCount * 2), (UserControl.Parent.Top + UserControl.Parent.Height) / Screen.TwipsPerPixelY + lngCount - 1, (UserControl.Parent.Width / Screen.TwipsPerPixelX) - lngCount - 1 - (lngCount \ 6), 1, picShadow.hdc, 0, 0, 1, 1, 200 / lngCount - 25)
        'Right
        Call AlphaBlending(lngHDc, (UserControl.Parent.Left + UserControl.Parent.Width) / Screen.TwipsPerPixelX + lngCount - 1, (UserControl.Parent.Top / Screen.TwipsPerPixelY) + (lngCount * 2), 1, (UserControl.Parent.Height / Screen.TwipsPerPixelY) - lngCount - (lngCount \ 6), picShadow.hdc, 0, 0, 1, 1, 200 / lngCount - 25)
        'Note:  The funky math in the two AlphaBlending calls do several things:
        '   1). Make the leading shadow edges (top right and bottom left) sharper
        '       than a 45degree angle so that the shadow to the eye appears
        '       to blend smoother without doing some funny extra calls to
        '       AlphaBlending.
        '   2). The right side shadow is 1 pixel shorter than the bottom in
        '       order to prevent a small, dark line from appearing where the
        '       two shadows would normally overlap at the bottom right.
        '   3). The last couple of lines are actually a couple of pixels shorter,
        '       which gives the hint of a rounded edge in the bottom right, again
        '       without having to make any extra calls to AlphaBlending.
    Next lngCount

    'Move the boxes that repaint the underlying screen to the new shadow's position
    rectBottom.Left = (UserControl.Parent.Left / Screen.TwipsPerPixelX)
    rectBottom.Top = (UserControl.Parent.Top + UserControl.Parent.Height) / Screen.TwipsPerPixelY
    rectBottom.Right = (UserControl.Parent.Left + UserControl.Parent.Width) / Screen.TwipsPerPixelX + 6
    rectBottom.Bottom = rectBottom.Top + 6
    rectRight.Left = (UserControl.Parent.Left + UserControl.Parent.Width) / Screen.TwipsPerPixelX
    rectRight.Top = (UserControl.Parent.Top / Screen.TwipsPerPixelY)
    rectRight.Right = rectRight.Left + 6
    rectRight.Bottom = (UserControl.Parent.Top + UserControl.Parent.Height) / Screen.TwipsPerPixelY + 6

    'Free the device context
    ReleaseDC 0&, lngHDc
    
End Sub

Private Sub tmrShadow_Timer()
    
    On Error Resume Next

    'Used static since none of the variables need to be visible outside
    '   of this routine

    Static blnProcessing As Boolean    'Are we already processing a redraw?
    Static blnVisible As Boolean       'Is the parent currently visible
        
    'If we are not the current foreground window then
    If GetForegroundWindow <> UserControl.Parent.hwnd Then
        
        'Were we the foreground window the last time we checked?
        '   If so, hide our shadow, otherwise do nothing (no reason to
        '   process shadows while we are hidden).
        If blnVisible = True Then
            
            'We were in the foreground last we checked, but now we are not,
            '   so hide our shadow.
            blnVisible = False
            RedrawWindow 0&, rectRight, 0&, 135
            RedrawWindow 0&, rectBottom, 0&, 135
            DoEvents
        
        End If  'Are we visible
    
    Else    'We are currently in the foreground, so allow shadow drawing
        
        'If the last time we checked we weren't visible, then update
        If blnVisible = False Then
            blnVisible = True
            sngX = -1   'Force the shadow to redraw
        End If
    
    End If  'Parent is the foreground window
    
    'Only draw if we are visible
    If blnVisible = True Then
    
        'Only process a redraw if we are not currently drawing
        If blnProcessing = False Then
        
            'If the form has moved, execute the redraw (no need to keep redrawing it
            '   when we haven't moved.
            If sngX <> UserControl.Parent.Left Or sngY <> UserControl.Parent.Top Then
    
                'Set the flag to prevent further processing
                blnProcessing = True
                            
                'Store the current coordinates
                sngX = UserControl.Parent.Left
                sngY = UserControl.Parent.Top
                
                'Now redraw the shadow
                Call Refresh
                
                'Allow a new redraw to occur
                blnProcessing = False
                            
            End If  'Usercontrol moved
            
        End If  'blnProcessing = true
        
    End If  'blnVisible = true
    
End Sub

Private Sub UserControl_Initialize()
    tmrShadow.Enabled = False   'This statement, which in itself is useless,
                                '   causes the UserControl_Resize event to
                                '   occur. The resize event does not fire
                                '   becuase the usercontrol is set to be
                                '   invisible at run time.
                                '   If it does not exist, the event
                                '   will not fire. We cannot directly call
                                '   UserControl_Resize because the control
                                '   isn't actually initialized until this
                                '   procedure completes, and hence, resize
                                '   (which references the ambient properties)
                                '   would cause an error.
                                
    'Set the current screen coordinates to something other than 0 so that
    '   on the off chance the form is starting up at 0,0 the shadow will
    '   still be drawn. (see checks in Sub Refresh)
    sngX = -1
End Sub

Private Sub UserControl_Resize()

    'Enable the timer control...not the best location but only guaranteed to
    '   fire if placed here...wait...if this is the only place guaranteed to
    '   fire, doesn't that make it the best location?
    If UserControl.Ambient.UserMode = True Then
        tmrShadow.Enabled = True
    Else
        tmrShadow.Enabled = False
    End If

    'Force the control to a particular height
    UserControl.Width = 420
    UserControl.Height = 420
    
End Sub


