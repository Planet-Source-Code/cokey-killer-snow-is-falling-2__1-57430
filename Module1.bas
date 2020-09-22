Attribute VB_Name = "Module1"
'These 2 api calls do the same job as 'Point' and 'PSet'
'But they do it a lot faster
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
'Only problem with setpixel, is you can only draw a single pixel, where
'with PSet you can set the draw width to be any size you wish.
'You can get around this by using SetPixel to draw 4 or more pixels in a tiny circle
'But this means you have to draw (4 * 'number of flakes') each cycle of the loop
'which would be slower than using PSet.  doh!
'This function isnt actually being used in the program, but I thought I would include
'it as it may be of use to someone.
Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'A simple snow structure
'Could you a class instead if you want, doesnt make a lot of difference.
Type Snow
    x As Long   'Current X Coordinates
    y As Long   'Current Y Coordinates
    oldX As Long    'Old X Coordinates
    oldY As Long    'Old Y Coordinates
    xSpeed As Long  'Horizontal Speed
    ySpeed As Long  'Vertical Speed
    FlakeSize As Long   'Size of flake
End Type

Public Flakes() As Snow 'Create an array of type 'Snow', this will be ReDimmed during 'Setup'

'Global Variables for our snow
Public mLeftWind As Long    'Stores the 'wind' factor
Public mRightWind As Long   '   '   '   '   '   '   '
Public mSnowColor As Long   'Colour of the snow flakes
Public mFlakeNum As Long    'How many snow flakes are we having
Public mfrmMainHDC As Long  'The hDC of frmMain. This is the form's handle
Public mtStopSnow As Boolean    'Used in the main loop
Public mSnowText As String      'Text drawn onto the screen
Public mtFlakeSize As Long      'Size of the snow flakes

Public Sub InitFlake(ByVal InIndex As Long)
    'Setup a new flake
    'Put the flakes X anywhere withing the forms limits
    Flakes(InIndex).x = CInt(Int(frmMain.ScaleWidth * Rnd))
    'Put the flake at the top of the screen
    Flakes(InIndex).y = 10
    'Reset the flakes oldY & oldY
    'This is important, because if we dont, the flakes will be
    'redrawn with black, and they will not appear to settle.
    Flakes(InIndex).oldX = 0
    Flakes(InIndex).oldY = 0
End Sub

Public Sub Setup()
Dim SnowText As String, SnowText2 As String
Dim i As Long
Dim NewX As Long, NewY As Long

    'This is the main setup routine
    
    'Clear the form of everything
    frmMain.Cls
    'Get the forms handle (hDC)
    'Its gonna be quicker to store this in a variable
    'instead of querying the form each time we need its handle
    mfrmMainHDC = frmMain.hDC
    
    'Resize the Snow() array to be the selected size (mFlakeNum)
    ReDim Flakes(mFlakeNum) As Snow
    
    'Loop through all the flakes
    For i = 0 To mFlakeNum - 1
        'Set the flakes x coordinates to be withing the forms limits
        Flakes(i).x = CInt(Int(frmMain.ScaleWidth * Rnd))
        'Set the flakes y coordinates to be random so that the snow
        'doesnt just all start from the top
        Flakes(i).y = CInt(Int(frmMain.ScaleHeight * Rnd))
        'Set the speed to 1
        Flakes(i).ySpeed = 1
        'Creates a random x speed
        NewX = Flakes(i).x + Int(2 * Rnd)
        NewX = NewX - Int(2 * Rnd)
        Flakes(i).xSpeed = NewX
        'Create a random flake size
        Flakes(i).FlakeSize = Int((mtFlakeSize * Rnd) + 1)
    Next i
    
    'Set up a couple of defaults
    mLeftWind = 2
    mRightWind = 2
    'Draw the text onto the main form
    DrawText
    frmMain.DrawWidth = mtFlakeSize
    frmMain.BackColor = vbBlack
    
    
    'Start the snow off
    frmMain.GoSnow
End Sub

Public Sub DrawText()
    'This sub simply draws the text on to the main form
    'It puts the text in the dead center
    frmMain.ForeColor = vbRed
    frmMain.CurrentX = (frmMain.ScaleWidth / 2 - frmMain.TextWidth(mSnowText) / 2)
    frmMain.CurrentY = (frmMain.ScaleHeight / 2 - frmMain.TextHeight(mSnowText) / 2)
    frmMain.Print mSnowText
    'Change the forms caption as well.
    frmMain.Caption = mSnowText
End Sub

Public Function GetDrawWidth()
    If frmMain.DrawWidth > 1 Then
        GetDrawWidth = (frmMain.DrawWidth - 1)
    Else
        GetDrawWidth = 1
    End If
End Function
