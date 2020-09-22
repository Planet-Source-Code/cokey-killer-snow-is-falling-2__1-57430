Attribute VB_Name = "Module2"
Public SplodeCol As Collection  'The collection containing all our Splodes
Public BiggieCol As Collection  'The collection containing all our Biggies

Public Sub Splode(ByVal InX As Single, InY As Single)
Dim i As Long
Dim oSplode As clsSplode

    'A new explosion is needed
    'Create a new object of our splode type
    Set oSplode = New clsSplode
    'Call the init routine in the object
    oSplode.InitSplode frmMain, InX, InY
    'Add the splode to our collection
    SplodeCol.Add oSplode
    'Tidy up by setting the object to nothing
    'we dont need it anymore, cause its been added to our collection
    Set oSplode = Nothing
End Sub

Public Sub CheckSplode()
Dim i As Long
Dim oldDW As Long
Dim NewX As Single, NewY As Single
    
    'Loop through the complete collection and call
    'each splode objects CheckSplode sub,
    'This sub Checks for collision, move the splode
    'and draws it on the screen
    For i = 1 To SplodeCol.Count
        SplodeCol.Item(i).CheckSplode
    Next i
    
End Sub

Public Sub BuryTheDead()
Dim i As Long
    
    'This sub loops through both of the collections
    'and checks if the objects dead property has been set to true
    'if they are then the objects are removed from the collection
    
    'Remove Dead Splodes
    For i = 1 To SplodeCol.Count
        If SplodeCol.Item(i).Dead = True Then
            SplodeCol.Remove i
            Exit For
        End If
    Next i
    
    'Remove Dead Biggies
    For i = 1 To BiggieCol.Count
        If BiggieCol.Item(i).Dead = True Then
            BiggieCol.Remove i
            Exit For
        End If
    Next i
End Sub

