'It is suggested that the comments for RenderForm are read first; comments were written in that order and some comments here may be simpler with the assumption the other comments
'have already been read


Dim fps, framesProcessed, playProgress, rewindModifier, fastforwardModifier As Integer
Dim paused, atEnd, rewinding, fastforwarding, switch As Boolean

'Error checking for the FPS box; minimum value of 1 and non-numeric values that are not "" are not allowed
Private Sub FPSBox_Change()
    With FPSBox
        If (IsNumeric(.Value)) Then
            If (.Value >= 1) Then
                .Value = CInt(.Value)
                fps = CInt(.Value)
            Else
                fps = 1
            End If
        ElseIf (Not .Value = "") Then
            If (Val(.Value) = 0) Then
                .Value = ""
            Else
                .Value = Val(.Value)
            End If
        End If
    End With
End Sub

'Rewinds the video by some rate that is defined in the program (user has no control over it at runtime
Private Sub RewindButton_Change()
    'Switch variable is a boolean that is here so that this sub does not do anything if the values of the button were changed by another sub instead of a mouse click
    'Switch variable is set to True whenever some part of the program alters the rewind button's value
    If (Not switch) Then
        If (rewinding) Then                                     'rewinding
            rewinding = False
        Else
            If (paused And fastforwarding) Then                 'fastforwarding
                fastforwarding = False
                'Calls a function to set the switch boolean to true if the value needs to be changed; there will be a few calls to switcher
                'Functions must return a value, so X is created to hold that value but is not used otherwise
                X = switcher(FastForwardButton, False)
            ElseIf (paused And Not fastforwarding) Then         'paused
            'Does nothing; mostly wrote this situation in for completeness' sake, but can easily be removed with no repurcussion
            
            ElseIf (Not paused And Not fastforwarding) Then     'playing
                PlayButton.Caption = "Play"
                paused = True
            End If
            
            rewinding = True
            
            'Scroll backward while rewinding until no longer rewinding or the screen reaches the beginning of the video
            Do While (rewinding And playProgress > 1)
                ScrollPage ActiveWindow.VisibleRange.Rows.Count, fps, False, rewindModifier
                'Didn't do this, but could cache playProgress (save value to cell) to save it between sessions so that the video can be resumed
                playProgress = playProgress - rewindModifier
                
                'Allow other things to happen while playing the video
                DoEvents
            Loop
            
            'If loop ends and the screen is at the beginning, stop rewinding and deactivate the rewind button
            If (playProgress = 1) Then
                rewinding = False
                X = switcher(RewindButton, False)
            End If
        End If
    'Let switch deactivate one call of the button's sub and then set it to false so that it doesn't deactivate multiple calls
    Else
        switch = False
    End If
End Sub

'Identical to rewind (with different associated variables)
Private Sub FastForwardButton_Change()
    If (Not switch) Then
        If (fastforwarding) Then
            fastforwarding = False
        Else
            If (paused And rewinding) Then
                rewinding = False
                X = switcher(RewindButton, False)
            ElseIf (paused And Not rewinding) Then
                
            ElseIf (Not paused And Not rewinding) Then
                PlayButton.Caption = "Play"
                paused = True
            End If
                    
            fastforwarding = True
            
            Do While (fastforwarding And playProgress < framesProcessed)
                ScrollPage ActiveWindow.VisibleRange.Rows.Count, fps, True, fastforwardModifier
                playProgress = playProgress + fastforwardModifier
                DoEvents
            Loop
            
            If (playProgress = framesProcessed) Then
                fastforwarding = False
                X = switcher(FastForwardButton, False)
            End If
        End If
    Else
        switch = False
    End If
End Sub

'Initialize variables when the userform is created
Private Sub UserForm_Activate()
    framesProcessed = ActiveSheet.Range("A1").Value
    playProgress = 1
    fps = 1
    FPSBox.Value = fps
    paused = False
    atEnd = False
    rewinding = False
    fastforwarding = False
    
    'Speed settings for rewind and fast forward; cannot be adjusted through the userform
    rewindModifier = 2
    fastforwardModifier = 3
    
    'Put the userform near the bottom-center of the Excel window
    With PlaybackForm
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.9 * Application.Height) - (0.5 * .Height)
    End With
End Sub

'Returns to the top (i.e. start of video) and resets all other buttons and variables
Private Sub ToBeginningButton_Click()
    ActiveWindow.ScrollRow = 1
    fastforwarding = False
    rewinding = False
    X = switcher(FastForwardButton, False)
    X = switcher(RewindButton, False)
    paused = True
    PlayButton.Caption = "Play"
    switch = False
End Sub

'Plays the video; similar to the render button before, toggles between play/pause and changes the caption of the button
Private Sub PlayButton_Click()
    fastforwarding = False
    rewinding = False
    X = switcher(FastForwardButton, False)
    X = switcher(RewindButton, False)

    'If the video was at the end, return to the top
    If (atEnd) Then
        ActiveWindow.ScrollRow = 1
        playProgress = 1
        atEnd = False
    End If
        
    'If the video was paused, unpause and alter the caption
    If (paused) Then
        paused = False
        PlayButton.Caption = "Pause"
      
        'Play while there is still more to play and while not paused
        Do While (playProgress < framesProcessed And Not paused)
            ScrollPage ActiveWindow.VisibleRange.Rows.Count, fps, True, 1
            DoEvents
            playProgress = playProgress + 1
        Loop
        
        'If the video is at the end after the loop ends, pause, set atEnd to true and alter the button's caption
        If (playProgress = framesProcessed) Then
            atEnd = True
            PlayButton.Caption = "Restart"
            paused = True
        End If
    'Otherwise if it wasn't paused, pause the video
    Else
        PlayButton.Caption = "Play"
        paused = True
    End If
End Sub

'Scrolls the page in a given direction; has a delay dependent on fps and a frame skip modifier
Sub ScrollPage(Offset, fps, down As Boolean, ByVal modifier As Integer)
    If (down) Then
        ActiveWindow.SmallScroll down:=Offset * modifier
    Else
        ActiveWindow.SmallScroll up:=Offset * modifier
    End If
    delay (1 / fps * 1000)
End Sub

'Calculate a delay based on the number of milliseconds of delay desired
Private Function delay(NumberOfMs As Variant)
    On Error GoTo Error_GoTo

    Dim PauseTime As Variant
    Dim Start As Variant
    Dim Elapsed As Variant

    PauseTime = NumberOfMs / 1000
    Start = Timer
    Elapsed = 0
    Do While Timer < Start + PauseTime
        Elapsed = Elapsed + 1
        If Timer = 0 Then
            ' Crossing midnight
            PauseTime = PauseTime - Elapsed
            Start = 0
            Elapsed = 0
        End If
        DoEvents
    Loop

Exit_GoTo:
    On Error GoTo 0
    Exit Function
Error_GoTo:
    Debug.Print Err.Number, Err.Description, Erl
    GoTo Exit_GoTo
End Function

'Gives the current time in milliseconds
Private Function timeInMs() As String
    timeInMs = Strings.Format(Now, "dd-MMM-yyyy HH:nn:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2)
End Function

'If a button value changes, set the switch variable to true to disable the button_Change() sub call once
Private Function switcher(ByRef button As Object, newval As Boolean) As Boolean
    If (Not (button.Value = newval)) Then
        switch = True
        button.Value = newval
    End If
End Function

