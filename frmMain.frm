VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   FillColor       =   &H00404000&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   72
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   572
   Begin VB.Timer Timer1 
      Left            =   3720
      Top             =   2400
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Colin Woor
'Email: colin@woor.co.uk
'Version: 2
'
'Hello, this is teh 2nd version of the snow program
'I got bored again! so i thought i would add some damaging big snowballs to
'make it a little bit more lively ;)

'If you have any comments etc. then feel free to email me.
'If you use any of this code, then please mention me in your code.
'
'Have fun :)

Option Explicit
Private mLoading As Boolean

Public Sub GoSnow()
Dim i  As Long
Dim oBiggie As clsBiggie
Dim NewX As Long
Dim NewY As Long

    Do
        For i = 0 To mFlakeNum - 1
            'Change the draw width to be the size of the flake
            frmMain.DrawWidth = Flakes(i).FlakeSize
            
            'Used PSet instead of SetPixel
            'As with PSet you can change the size of the pixels using DrawWidth
                        
            'Draw the snow flake again in its old position, but with black to remove it
            'SetPixel mfrmMainHDC, Flakes(i).oldX, Flakes(i).oldY, vbBlack
            frmMain.PSet (Flakes(i).OldX, Flakes(i).OldY), vbBlack
            
            'Now draw it again in its new position using the snow color
            'SetPixel mfrmMainHDC, Flakes(i).x, Flakes(i).y, GetSnowColor
            frmMain.PSet (Flakes(i).x, Flakes(i).y), GetSnowColor
            
        Next i
        
        For i = 0 To mFlakeNum - 1
            'Store the flakes old x & y
            Flakes(i).OldX = Flakes(i).x
            Flakes(i).OldY = Flakes(i).y
            
            'Get a new xspeed
            NewX = GetXSpeed(i)
                               
            'Keep the X coordinate withing the form
            If NewX < 0 Then NewX = 0
            If NewX >= frmMain.ScaleWidth Then NewX = frmMain.ScaleWidth - 1
            
            'Set the flakes new X coordinate
            Flakes(i).xSpeed = NewX + GetDrawWidth
            'Set the flakes Y cood, adding on the YSpeed and the DrawWidth so it moves
            'in proportion to its size
            NewY = Flakes(i).y + (Flakes(i).ySpeed + GetDrawWidth)
                                    
            'If the new target coordinates are black then we can
            'set the actual flakes x & y to the new x & y.
            If GetPixel(mfrmMainHDC, NewX, NewY) = vbBlack Then
                Flakes(i).y = NewY
                Flakes(i).x = NewX
            Else
                'This section attempts to let the snow move in a realistic kinda way
                'Basically it looks to see if there is anywhere for it to go left or right
                'This simulates snow sliding
                
                'We look left to see if there is any room for us to move
                If GetPixel(mfrmMainHDC, (Flakes(i).x + 1) + GetDrawWidth, (Flakes(i).y + 1) + GetDrawWidth) = vbBlack Then
                    Flakes(i).x = Flakes(i).x + 1
                    Flakes(i).y = Flakes(i).y + 1
                    
                'We look right to see if there is any room for us to move
                ElseIf GetPixel(mfrmMainHDC, (Flakes(i).x - 1) - GetDrawWidth, (Flakes(i).y + 1) + GetDrawWidth) = vbBlack Then
                    Flakes(i).x = Flakes(i).x - 1
                    Flakes(i).y = Flakes(i).y + 1
                Else
                    'Theres no where to go, so the flake in the array can be initiated
                    'Because the actual flake has been drawn onto the screen, it looks
                    'as if the flake has settled.
                    InitFlake i
                End If
            End If
            
            'Keep the flakes inside of the form
            If Flakes(i).y >= frmMain.ScaleHeight Then
                InitFlake i
            End If
        Next i
        
        'Draw the text. This is done each cycle of the do-loop
        'in case the form has been resized.
        DrawText
        
        'This sub draws, checks for collision and moves the explosions (splodes ;) )
        CheckSplode
        'This sub iterates through the collection, and removes any dead objects
        BuryTheDead
        
        'Now go through each biggie in the biggie collection
        For Each oBiggie In BiggieCol
            'Set the size of the biggie
            frmMain.DrawWidth = oBiggie.BiggieSize
            'Overwrite the old one with black
            frmMain.PSet (oBiggie.OldX, oBiggie.OldY), vbBlack
            'Redraw the biggie in its new x/y position
            frmMain.PSet (oBiggie.x, oBiggie.y), oBiggie.BiggieColor
            'update the biggies oldx/oldy positions
            oBiggie.OldX = oBiggie.x
            oBiggie.OldY = oBiggie.y
            'Set the biggies new Y position...current position + biggie speed
            NewY = oBiggie.y + oBiggie.ySpeed
            'Get a random x position
            NewX = oBiggie.x + (mRightWind * Rnd)
            NewX = NewX - (mLeftWind * Rnd)
            'Check if the biggie can move to its new position
            If GetPixel(mfrmMainHDC, NewX, NewY + oBiggie.BiggieSize) = vbBlack Then
                'Nothing but empty space, so set its x/y to the new x and y
                oBiggie.y = NewY
                oBiggie.x = NewX
            Else
                'Because something was in the way of its new position, try and look left and right
                'to see if we can move there...this simulates sliding, almost like real snow ;)
                If GetPixel(mfrmMainHDC, (oBiggie.x + 1) + oBiggie.BiggieSize, (oBiggie.y + 1) + oBiggie.BiggieSize) = vbBlack Then
                    'We can move to the right as there is nothing there
                    oBiggie.x = oBiggie.x + 1
                    oBiggie.y = oBiggie.y + 1
                ElseIf GetPixel(mfrmMainHDC, (oBiggie.x - 1) - oBiggie.BiggieSize, (oBiggie.y + 1) + oBiggie.BiggieSize) = vbBlack Then
                    'We can move to the left as there is nothing there
                    oBiggie.x = oBiggie.x - 1
                    oBiggie.y = oBiggie.y + 1
                Else
                    'We can't move anywhere
                    'Increase the drawwidth to simulate a crator
                    frmMain.DrawWidth = oBiggie.BiggieSize + 2
                    'create a crator
                    frmMain.PSet (NewX, NewY + 1), vbBlack
                    'put the drawwidth back to the original size
                    frmMain.DrawWidth = oBiggie.BiggieSize
                    'remove the biggie
                    frmMain.PSet (oBiggie.x, oBiggie.y), vbBlack
                    'Start the explosion
                    Splode CSng(NewX), CSng(NewY)
                    'Reset the biggie
                    oBiggie.y = 0
                    oBiggie.Dead = True
                End If
            End If
        Next
        
        DoEvents
        
        'Loop until we set the myStopNow variable, via the Property Let
        'This is done from the settings form, to reset the flakes variables etc.
    Loop Until mtStopSnow = True
End Sub

Private Function GetXSpeed(ByVal Index As Long) As Long
Dim NewX As Long

    'Add some wind to the XSpeed
    'This method allows us to get positive
    'and negative numbers
    NewX = Flakes(Index).x + (mRightWind * Rnd)
    NewX = NewX - (mLeftWind * Rnd)
    GetXSpeed = NewX
End Function

Private Sub Form_DblClick()
    'Show the settings
    frmSettings.Show
End Sub

Private Sub Form_Load()
    mLoading = True
    'Set up a couple of variables
    mSnowText = "Killer Snow" 'The text in the middle of the form
    mFlakeNum = 200 'How many flakes we gonna have?
    mtFlakeSize = 2 'This is the draw width
    
    'Position the form
    frmMain.Left = (Screen.Width / 2) - (frmMain.Width / 2)
    frmMain.Top = (Screen.Height / 2) - (frmMain.Height / 2)
    frmMain.Show
    mLoading = False
    'Give the timer a nice random interval
    Timer1.Interval = Int((500 * Rnd) + 100)
    Timer1.Enabled = True
    
    frmSettings.Show
    'Init the collection
    Set SplodeCol = New Collection
    Set BiggieCol = New Collection
    
    'Setup everything else
    Setup
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This create an explosion at the mouses position
    'when the mouse is clicked
    Splode x, y
End Sub

Private Sub Form_Resize()
    If mLoading = False Then
        mtStopSnow = True
        frmMain.Cls
        Setup
        mtStopSnow = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Function GetSnowColor() As Long
Dim tmpRnd As Long
Dim tNumOptions As Long
    
    'This function allows you to have flakes with lots of different colours
    'or different shades of grey/white

    'This alters out random number to be within the same
    'number as the number of case statements
    tNumOptions = 2
    
    tmpRnd = Int((tNumOptions * Rnd) + 1)
    Select Case tmpRnd
        Case 1
            GetSnowColor = 16777215
        Case 2
            GetSnowColor = 14737632
        'You can add as many colours as you like
        'but you need to set the tNumOptions variable to be the same
        'as the amount of case statements you have
    End Select
    
    'This is hard coded to always return white.
    'Just remove this line to allow the random
    'colours above to be used instead.
    GetSnowColor = vbWhite
End Function

Public Property Let StopSnow(ByVal InStop As Boolean)
    'Either stop the loop (if true)
    'or let it start again (if false)
    'This is set from the settings form.
    mtStopSnow = InStop
End Property

Private Sub Timer1_Timer()
Dim tmpFlake As Long
Dim oBiggie As clsBiggie

    'Set up the biggie
    Set oBiggie = New clsBiggie
    'Its X position
    oBiggie.x = CInt(Int(frmMain.ScaleWidth * Rnd))
    oBiggie.y = 0
    oBiggie.Dead = False
    oBiggie.ySpeed = 2
    oBiggie.BiggieColor = vbWhite
    'How big is it gonna be
    oBiggie.BiggieSize = 5
    'Add the biggie to the collection
    BiggieCol.Add oBiggie
    Set oBiggie = Nothing
    Timer1.Interval = Int((500 * Rnd) + 500)
End Sub
