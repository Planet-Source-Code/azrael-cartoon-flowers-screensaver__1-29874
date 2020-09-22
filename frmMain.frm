VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDrawFlower 
      Interval        =   1
      Left            =   90
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long) As Long

'##################################################################################################################
'Events
'##################################################################################################################

Private Sub Form_KeyPress(KeyAscii As Integer)
    End
End Sub

Private Sub Form_Load()
    'Don't run multiple copies of program
    If App.PrevInstance Then End
    
    'If program wasn't started as a screensaver, then end
    If Trim(LCase(Command)) <> "/s" Then End
    
    'Seed the random number generator
    Randomize Timer
    
    'Draw the desktop onto the form and set other things up
    SetupScreen
    
    DrawSun
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static lMoveCount As Long
    
    'Mice can be very sensitive - it's best to count mousemoves a few times before
    'actually ending the program.
    
    lMoveCount = lMoveCount + 1
    
    If lMoveCount = 10 Then End
End Sub

Private Sub tmrDrawFlower_Timer()
    'Draw a new flower
    DrawFlower
End Sub

'##################################################################################################################
'Procs
'##################################################################################################################
Private Sub SetupScreen()
    'This sub copies a picture of the desktop onto the form, which is maximised and doesn't
    'have a border. You also need to have AutoRedraw set to true for the form.
    Dim lReturn As Long
    Dim DeskhWnd As Long
    Dim DeskDC As Long
    Dim FormDC As Long
    
    'Set up some 'constant' colours
    BLACK = QBColor(0)
    GREEN = QBColor(10)
    YELLOW = QBColor(14)
    
    'Force form to same position and size as desktop
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Me.Top = 0
    Me.Left = 0
    
    'Need to change scale mode to draw the desktop picture
    Me.ScaleMode = vbPixels
    
    'Get a handle to the desktop
    DeskhWnd = GetDesktopWindow
    
    'Get a device contet for that handle
    DeskDC = GetWindowDC(DeskhWnd)
    
    'This is the line that actually paints the desktop on to the form
    lReturn = BitBlt(Me.hDC, 0, 0, (Screen.Width \ Screen.TwipsPerPixelX), _
        (Screen.Height \ Screen.TwipsPerPixelY), DeskDC, 0, 0, vbSrcCopy)
        
    'Now change it back...
    Me.ScaleMode = vbTwips
    
    'Make sure form fills screen
    Me.WindowState = vbMaximized
    
    'This makes the circles 'solid', sets the width of the lines, etc.
    Me.FillStyle = 0
    Me.DrawStyle = 0
    Me.DrawWidth = 3
    Me.ForeColor = BLACK
End Sub

Private Sub DrawSun()
    'Draw Sun
    Me.FillColor = YELLOW
    Me.Circle (Screen.Width / 7, Screen.Width / 7), Screen.Width / 9, BLACK
    
    'Draw Smile
    Me.Circle (Screen.Width / 7, Screen.Width / 7), Screen.Width / 9 * SMILE_SCALE, _
        BLACK, SMILE_START, SMILE_END
    
    'Draw Eyes
    Me.FillColor = BLACK
    Me.Circle (Screen.Width / 7 + XChange(Screen.Width / 9 * EYE_SCALE, RIGHT_EYE_BEARING), _
        Screen.Width / 7 + YChange(Screen.Width / 9 * EYE_SCALE, RIGHT_EYE_BEARING)), _
        EYE_SIZE * 3, BLACK
    Me.Circle (Screen.Width / 7 + XChange(Screen.Width / 9 * EYE_SCALE, LEFT_EYE_BEARING), _
        Screen.Width / 7 + YChange(Screen.Width / 9 * EYE_SCALE, LEFT_EYE_BEARING)), _
        EYE_SIZE * 3, BLACK
End Sub

'Every time the timer goes off...
Private Sub DrawFlower()
    Dim lX As Long, lY As Long
    
    Dim lCentreSize As Long
    Dim lCentreColour As Long
    
    Dim lPetalSize As Long
    Dim lPetalCount As Long
    Dim lPetalLoopTMP As Long
    Dim lPetalSpacing As Long
    Dim lPetalOffset As Long
    Dim Petals() As Long
    
    'Position of centre of flower - note how Y pos is biased
    lX = Int(Rnd * Screen.Width)
    lY = Int(Rnd * Screen.Height) / 3 + Int(Rnd * Screen.Height) / 3 + _
        Screen.Height / 3
    
    lCentreSize = Int(Rnd * CENTRE_SIZE) + CENTRE_SIZE
    lPetalSize = lCentreSize + Int(Rnd * lCentreSize) / 2 - lCentreSize / 2
    lPetalOffset = lCentreSize
    
    'Number of petals
    lPetalCount = Int(Rnd * 3) + 4 '4-6
    ReDim Petals(lPetalCount, 3)
    
    'Spacing between petals around centre, in degrees
    lPetalSpacing = FULL_CIRCLE / lPetalCount
    
    'Work out the positions of the centres of the petals
    For lPetalLoopTMP = 1 To lPetalCount
        'Bearing of centre from centre (degrees)
        Petals(lPetalLoopTMP, 1) = (lPetalLoopTMP - 1) * lPetalSpacing
        
        'Distance = lPetalOffset
        
        'XPos
        Petals(lPetalLoopTMP, 2) = lX + XChange(lPetalOffset, Petals(lPetalLoopTMP, 1))
        
        'YPos
        Petals(lPetalLoopTMP, 3) = lY + YChange(lPetalOffset, Petals(lPetalLoopTMP, 1))
    Next lPetalLoopTMP
    
    
    'Right, let's actually draw something...
    
    'object.Line [Step] (x1, y1) [Step] - (x2, y2), [color], [B][F]
    'object.Circle [Step] (x, y), radius, [color, start, end, aspect]
    
    'DrawStalk
    Me.FillColor = GREEN
    Me.Line (lX - DEFAULT_STALK, lY)-(lX + DEFAULT_STALK, Screen.Height), BLACK, B
    
    'DrawPetals
    Me.FillColor = GetPetalColour
    For lPetalLoopTMP = 1 To lPetalCount
        Me.Circle (Petals(lPetalLoopTMP, 2), Petals(lPetalLoopTMP, 3)), lPetalSize, _
            BLACK
    Next lPetalLoopTMP
    
    'Draw Centre
    Me.FillColor = YELLOW
    Me.Circle (lX, lY), lCentreSize, BLACK
    
    'Draw Smile
    Me.Circle (lX, lY), lCentreSize * SMILE_SCALE, _
        BLACK, SMILE_START, SMILE_END
    
    'Draw Eyes
    Me.FillColor = BLACK
    Me.Circle (lX + XChange(lCentreSize * EYE_SCALE, RIGHT_EYE_BEARING), _
        lY + YChange(lCentreSize * EYE_SCALE, RIGHT_EYE_BEARING)), EYE_SIZE, BLACK
    Me.Circle (lX + XChange(lCentreSize * EYE_SCALE, LEFT_EYE_BEARING), _
        lY + YChange(lCentreSize * EYE_SCALE, LEFT_EYE_BEARING)), EYE_SIZE, BLACK
        
End Sub

'##################################################################################################################
'Functions
'##################################################################################################################

Private Function GetPetalColour() As Long
    'Return a random colour - I went for Bold and Bright and Simple, but you
    'could get fancy with the RGB function in here ... (avoid 'Grey' however)...
    Select Case Int(Rnd * 5) + 1
        
        Case 1 'Dark Blue
            GetPetalColour = QBColor(1)
            
        Case 2 'Light Blue
            GetPetalColour = QBColor(9)
            
        Case 3 'Red
            GetPetalColour = QBColor(12)
            
        Case 4 'Pink
            GetPetalColour = QBColor(13)
            
        Case 5 'White
            GetPetalColour = QBColor(15)
            
    End Select
End Function
