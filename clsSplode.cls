VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSplode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is the explosion class...also known to me as Splodes
'I guess Splodes actually means the shrapnal or particles in the explosion

Private mX As Double    'Initial X position of the splode
Private mY As Double    'Initial Y position of the ball
Private mDead As Boolean    'Are we dead yet?
Private mtForm As Form      'The form that we want to draw on
                            'this is good encapsulation

Private Splodes() As SplodeType 'Our array that will contain our splodes
Const PI = 3.14159  'We all know what this is...dont we?
Const cHEADING = 15 * PI / 8    'Used for our angle calculations
Const FALL_GRAVITY = 0.1    ''Gravity represented as a simple ole number :)
Private Const mSplodes = 50 'How many splodes per explosion

Private Type SplodeType
    x As Double   'X Coordinates
    y As Double   'Y Coordinates
    OldX As Double    'Old X Coordinates
    OldY As Double    'Old Y Coordinates
    SplodeSize As Long   'Size of splode
    Heading As Double   'Current heading of the splode
    Speed As Double     'General speed of the splode
    Dead As Boolean     'Is the splode dead (reached the bottom of the form,
                        'of piled up on some more snow)
End Type

Public Property Let x(ByVal InX As Double)
    mX = InX
End Property
Public Property Let y(ByVal InY As Double)
    mY = InY
End Property
Public Property Get x() As Double
    x = mX
End Property
Public Property Get y() As Double
    y = mY
End Property
Public Property Let Dead(ByVal InDead As Boolean)
    mDead = InDead
End Property
Public Property Get Dead() As Boolean
    Dead = mDead
End Property

Public Sub InitSplode(ByVal inForm As Form, _
                      ByVal InX As Double, _
                      ByVal InY As Double, _
                      Optional ByVal InFlakeSize As Long)
Dim i As Long

    'This routines sets up the splode, its inits its starting position
    'which is passed in
    
    Set mtForm = inForm
    
    'If there is no size passed in
    'set it to 1
    If InFlakeSize > 0 Then
        mFlakeSize = InFlakeSize
    Else
        mFlakeSize = 1
    End If
    
    'Resize the array to our chosen number of splodes
    ReDim Splodes(mSplodes)
    'Loop through them all and set them up
    For i = 1 To mSplodes
        Splodes(i).x = InX
        Splodes(i).y = InY
        Splodes(i).SplodeSize = mFlakeSize
        'This set the heading of our splode
        'This will set the heading to be
        'about between 10 o'clock and 2 o'clock
        'I.e: up and a little to the sides
        Splodes(i).Heading = ((2.5 * Rnd) + 5.1)
        'This gives our splode a random speed
        Splodes(i).Speed = (2 * Rnd) + 0.5
        'We havent died yet
        Splodes(i).Dead = False
    Next i
    'Set the object to be alive
    mDead = False
End Sub

Public Sub CheckSplode()
Dim i As Long
Dim oldDW As Long
Dim NewX As Single, NewY As Single
Dim allDead As Boolean

    'This sub draws the spodes, moves them and
    'checks for collisions (the bottom of the form, and piles of snow)

    allDead = True
    'Store the old forms draw width value
    oldDW = mtForm.DrawWidth

    'loop through all the splodes
    For i = 1 To mSplodes
        mtForm.DrawWidth = Splodes(i).SplodeSize
        'Remove the old splode
        mtForm.PSet (Splodes(i).OldX, Splodes(i).OldY), RGB(0, 0, 0)
        'Draw the new one
        mtForm.PSet (Splodes(i).x, Splodes(i).y), RGB(255, 255, 255)
        'Update the splodes oldx and oldy values
        Splodes(i).OldX = Splodes(i).x
        Splodes(i).OldY = Splodes(i).y
        'This call our routine to move the splode
        'in an explosion type of way
        SetPhysics i, NewX, NewY
        'Check for collisions (bottom of form and piles of snow)
        CheckCol i, NewX, NewY
        'This lets us keep a track on if all the splodes
        'have died
        If Splodes(i).Dead = False Then
            allDead = False
        End If
    Next i
    'Restore the original form draw width
    mtForm.DrawWidth = oldDW
    'If all the splodes are dead (stoped moving) then
    'set the objects dead property to true
    If allDead = True Then
        mDead = True
    End If
End Sub

Private Sub CheckCol(ByVal Index As Long, _
                     ByVal NewX As Single, _
                     ByVal NewY As Single)
            
    'This sub checks to see if splode can move or if it
    'has come to an obstacal
    If GetPixel(mtForm.hDC, NewX, NewY) = vbBlack Then
        Splodes(Index).y = NewY
        Splodes(Index).x = NewX
    Else
        'Look left and right to find a way through
        If GetPixel(mtForm.hDC, (Splodes(Index).x + 1), (Splodes(Index).y + 1)) = vbBlack Then
            Splodes(Index).x = Splodes(Index).x + 1
            Splodes(Index).y = Splodes(Index).y + 1
        ElseIf GetPixel(mtForm.hDC, (Splodes(Index).x - 1), (Splodes(Index).y + 1)) = vbBlack Then
            Splodes(Index).x = Splodes(Index).x - 1
            Splodes(Index).y = Splodes(Index).y + 1
        Else
            'We are stuck so reset the splode
            'and set it to be dead
            Splodes(Index).OldX = 0
            Splodes(Index).OldY = 0
            Splodes(Index).Dead = True
        End If
    End If
    
    If Splodes(Index).y >= frmMain.ScaleHeight Then
        'The splode is at the bottom of the form
        'so reset it, and set it to be dead
        Splodes(Index).OldX = 0
        Splodes(Index).OldY = 0
        Splodes(Index).Dead = True
     End If
End Sub


Private Sub SetPhysics(ByVal Index As Long, _
                       ByRef NewX As Single, _
                       ByRef NewY As Single)
Dim sX As Single
Dim sY As Single

    'This sub calculates the splodes trajectory
    'This simulates the effect of gravity
    sX = Sin(Splodes(Index).Heading) * Splodes(Index).Speed
    sY = Cos(Splodes(Index).Heading) * Splodes(Index).Speed - FALL_GRAVITY
    
    'Determine the speed
    Splodes(Index).Speed = Sqr(sX ^ 2 + sY ^ 2)
    
    'Calculate the heading and adjust for arc by adding Pi
    If sY > 0 Then Splodes(Index).Heading = Atn(sX / sY)
    If sY < 0 Then Splodes(Index).Heading = Atn(sX / sY) + PI
    
    'If the splode has reached the forms bottom then stop it!
    If Splodes(Index).y > mtForm.ScaleHeight Then Splodes(Index).Speed = 0
    
    'Return the splodes new x and new y values
    NewX = Splodes(Index).x + Splodes(Index).Speed * Sin(Splodes(Index).Heading)
    NewY = Splodes(Index).y - Splodes(Index).Speed * Cos(Splodes(Index).Heading)
    
End Sub

