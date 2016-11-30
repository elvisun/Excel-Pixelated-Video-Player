Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private hDC As Long 'userForm handler
Private pRight, pBot, pTop, pLeft As Long
Private paused As Boolean
Private pixelSize, rowCount, colCount, frameCount, frameskip, framesProcessed As Integer
Private targetSheet As Worksheet
Private startingCell As Range

'Error checking for the box that lets users define the frameskip during rendering
'Frameskip means that the rendering will completely ignore any frames that are "skipped"
Private Sub FrameskipBox_Change()
    With FrameskipBox
    'Checks if the box's contents are numeric
        If (IsNumeric(.Value)) Then
        'Checks if box's contents are at least 1
            If (.Value >= 1) Then
            'If so, set the variable to that value and allow the box to change to that value
                .Value = CInt(.Value)
                frameskip = CInt(.Value)
            'If numeric but less than 1, set the minimum to be 1 (less than 1 is undefined behavior)
            Else
                frameskip = 1
            End If
        'Allow a blank box as a special case of non-numeric contents
        ElseIf (Not .Value = "") Then
            'Treat 0 as a blank
            If (Val(.Value) = 0) Then
                .Value = ""
            Else
            'Otherwise if not blank, take out the non-numeric parts of the box contents
                .Value = Val(.Value)
            End If
        End If
    End With
    
    'Update the label for the total number of frames to be rendered based on frameskip
    ProgressLabel.Caption = "Progress: " & framesProcessed & " / " & CInt(frameCount / frameskip) & " frames rendered"
End Sub

Private Sub Image1_Click()

End Sub

'Error checking for pixel size; pixel size refers to the dimensions of each "pixel" on the worksheet
'In other words, pixel size is used to change column width and row height on the worksheet
Private Sub PixelSizeBox_Change()
    With PixelSizeBox
        If (IsNumeric(.Value)) Then
            If (.Value >= 1) Then
                .Value = CInt(.Value)
                pixelSize = CInt(.Value)
            Else
                pixelSize = 1
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

'Initialize variables when the userform is opened
Private Sub UserForm_Activate()
    Dim Path As String: Path = Application.ActiveWorkbook.Path

    'Initialize variables'''''''''''''''''''''''''''''''''''''''''''''''
    hDC = GetDC(FindWindow(vbNullString, Me.Caption))
        
    'Defines the range of pixels to get while rendering (bounds of the image)
    'Couldn't find a way to express these without using magic numbers
    pLeft = 12
    pRight = 640
    pTop = 148
    pBot = 540

    pixelSize = 6                                                   'Default  pixels size (width and height)
    frameskip = 1                                                   '1 is no frameskip
    paused = True                                                   'Rendering is paused at the beginning
    frameCount = Int(FilesInFolder(Path & "\frames") / frameskip)   'Counts the number of frames that will need to be rendered
    framesProcessed = ActiveSheet.Range("A1").Value                 'Cache rendering progress in cell A1
    
    
    
    'Initialize userform labels and textbox contents''''''''''''''''''''
    FrameskipBox.Value = frameskip
    PixelSizeBox.Value = pixelSize
    ProgressLabel.Caption = "Progress: " & framesProcessed & " / " & CInt(frameCount / frameskip) & " frames rendered"
    RenderButton.Caption = "Start Rendering"
    


    'Initialize grid''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set targetSheet = ActiveSheet
    Set startingCell = targetSheet.Range("A1")

    resizeGrid targetSheet, pixelSize               'Resize grid given pixel size
    result = countCells(targetSheet)                'Find the "resolution" of the picture
    rowCount = result(0)
    colCount = result(1)
End Sub

'Counts the number of files in a folder; used to find the number of images that need to be rendered
Private Function FilesInFolder(Path As String) As Long
    On Error GoTo Handler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FilesInFolder = fso.GetFolder(Path).Files.Count
    Exit Function
    
Handler:
    FilesInFolder = 0
End Function

'Counts the number of visible cells on-screen
Private Function countCells(targetSheet) As Variant
    With ActiveWindow.VisibleRange
        visRows = .Rows.Count
        visCols = .Columns.Count
    End With
    countCells = Array(visRows, visCols)
End Function

'Resizes the cells on screen to create a certain pixel size
Private Sub resizeGrid(targetSheet, pixelSize)
    With targetSheet
        Rows().RowHeight = (9.75 / 13) * pixelSize              ' 9.75 == 13px, 10px = 9.75/1.3
        Columns().ColumnWidth = (1 / 12) * pixelSize            ' 1 == 12 px, 10 px = 1/1.2
    End With
End Sub

'Downloads a video using the given URL and calls a Python script to extract the individual frames from the video
Private Sub DownloadButton_Click()
    'Calls a sub to download a video
    If (Not downloadVideo(URLBox.Value)) Then
        MsgBox ("Download failed! Check the URL again.")
        Exit Sub
    End If
    
    'Calls a function to find the file that was just downloaded
    Dim filename As String
    
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Do While (Not Right(filename, 3) = "mp4")
        filename = mostRecentDownload()
        DoEvents
    Loop
    
    'Creates a file that the Python script uses as input; gives the script the location of the downloaded file
    Open Application.ActiveWorkbook.Path & "\" & "input.txt" For Output As #2
    Print #2, filename;
    Close #2
    
    'Run the main() function of convert.py; this works using xlwings to smooth the process of connecting Excel and Python
    RunPython ("import convert; convert.main()")
    
    'Delete the input file that was just created; no longer needed
    Kill Application.ActiveWorkbook.Path & "\" & "input.txt"
End Sub

'Downloads a video from a given URL using a third-party site to extract the video from the video sharing site (e.g. Youtube)
Private Function downloadVideo(video As String) As Boolean
    Dim videoURL As String: videoURL = video
    Dim siteURL As String: siteURL = "http://www.videograbby.com/"
    
    On Error GoTo Handler
    
    'Creates an instance of Internet Explorer that navigates to the site
    Set IE = CreateObject("InternetExplorer.application")
        With IE
        .Visible = True
        .navigate siteURL
        
        'Wait until the site is fully loaded before continuing
        Do Until Not .Busy And .readyState = 4
            DoEvents
        Loop
    
        'Fills in the forms on the site and clicks the download button
        .document.getElementsByClassName("input-url ui-autocomplete-input")(0).Value = videoURL
        .document.getElementsByClassName("btn-download-video")(0).Click
        
        'Wait until the file is created by the site before continuing
        Do Until Not .Busy And .readyState = 4 And .document.getElementsByClassName("input-url ui-autocomplete-input")(0).Value = ""
            DoEvents
        Loop
        
        'Use Alt+S for the Internet Explorer download prompt; the Internet Explorer window must still be active for this to work
        Application.Wait (Now + TimeValue("0:00:02"))
        Application.SendKeys "%{S}"
    End With
    
    'If the function doesn't throw an error, return true
    downloadVideo = True
    Exit Function
    
    'Otherwise return false
Handler:
    downloadVideo = False
End Function

'Looks for the last modified file in the downloads folder; the site used to download the video creates a new file every time, so it should be the last modified download
Private Function mostRecentDownload() As String
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim f As File
    Dim Path As String: Path = "C:\Users\" & Environ$("UserName") & "\Downloads"
    Dim downloadsFolder: Set downloadsFolder = fso.GetFolder(Path)
    Dim dateModified As Date: dateModified = DateSerial(1900, 1, 1)

    For Each f In downloadsFolder.Files
        If (f.DateLastModified > dateModified) Then
            dateModified = f.DateLastModified
            mostRecentDownload = Path & "\" & f.Name
        End If
    Next f

    Set fso = Nothing
    Set downloadsFolder = Nothing
End Function

'Clears all cells in the current sheet of formatting; also calls a sub to drop now-unused styles and resets the counter for processed frames to 0 (clearing cell A1 as part of that)
Private Sub ClearButton_Click()
    ActiveSheet.Cells.Interior.color = -1
    DropUnusedStyles
    ActiveSheet.Range("A1").Value = 0
    framesProcessed = 0
End Sub

'This sub drops unused styles, preventing Excel from exceeding its hard limit of 64000 unique styles between different renders (shouldn't hit the limit anyway because of a different process)
Private Sub DropUnusedStyles()
    Dim styleObj As Style
    Dim rngCell As Range
    Dim wb As Workbook
    Dim Wsh As Worksheet
    Dim str As String
    Dim iStyleCount As Long
    Dim dict As New Scripting.Dictionary    ' <- from Tools / References... / "Microsoft Scripting Runtime"

    ' Choose a workbook to clear
    Set wb = ThisWorkbook ' The workbook that holds this sub
    'Set wb = ActiveWorkbook ' the active workbook

    Debug.Print "BEGINNING # of styles in workbook: " & wb.Styles.Count

    ' dict := list of styles
    For Each styleObj In wb.Styles
        str = styleObj.NameLocal
        iStyleCount = iStyleCount + 1
        Call dict.Add(str, 0)    ' First time:  adds keys
    Next styleObj
    Debug.Print "  dictionary now has " & dict.Count & " entries."
    ' Status, dictionary has styles (key) which are known to workbook

    ' Traverse each visible worksheet and increment count each style occurrence
    For Each Wsh In wb.Worksheets
        If Wsh.Visible Then
            For Each rngCell In Wsh.UsedRange.Cells
                str = rngCell.Style
                dict.Item(str) = dict.Item(str) + 1     ' This time:  counts occurrences
            Next rngCell
        End If
    Next Wsh
    ' Status, dictionary styles (key) has cell occurrence count (item)

    ' Try to delete unused styles
    Dim aKey As Variant
    On Error Resume Next    ' wb.Styles(aKey).Delete may throw error

    For Each aKey In dict.Keys

        ' display count & stylename
        '    e.g. "24   Normal"
        Debug.Print dict.Item(aKey) & vbTab & aKey

        If dict.Item(aKey) = 0 Then
            ' Occurrence count (Item) indicates this style is not used
            Call wb.Styles(aKey).Delete
            If Err.Number <> 0 Then
                Debug.Print vbTab & "^-- failed to delete"
                Err.Clear
            End If
            Call dict.Remove(aKey)
        End If

    Next aKey

    Debug.Print "ENDING # of style in workbook: " & wb.Styles.Count
End Sub

'Modifies each individual frame extracted by the Python script and transposes pixels onto the worksheet
Private Sub RenderButton_Click()
    'Pressing the button toggles pause status and changes the button text
    If (paused) Then
        paused = False
        RenderButton.Caption = "Pause Rendering"
        
        'Render while not paused and rendering is not complete
        Do While (framesProcessed < frameCount And Not paused)
            'Draws the next frame with an offset based on the number of frames already rendered; only renders frames that aren't skipped
            drawFrame startingCell.Offset(framesProcessed * rowCount, 0), colCount, rowCount, framesProcessed * frameskip
            
            'Allow other events to occur while rendering
            DoEvents
            
            'Update the number of frames rendered
            framesProcessed = framesProcessed + 1
            ActiveSheet.Range("A1").Value = framesProcessed
            ProgressLabel.Caption = "Progress: " & framesProcessed & " / " & CInt(frameCount / frameskip) & " frames rendered"
            
            'Automatically save every 100 frames
            If (framesProcessed Mod 100 = 0) Then
                ActiveWorkbook.Save
            End If
        Loop
        
        'If rendering is complete when loop exits, pause and save the workbook
        If (framesProcessed = frameCount) Then
            RenderButton.Caption = "Rendering complete!"
            paused = True
            ActiveWorkbook.Save
        End If
    'If rendering was already occuring, pause and change the button's caption
    Else
        paused = True
        RenderButton.Caption = "Continue Rendering"
        
        'Backtrack one frame because pausing tends to break the rendering for the current frame
        framesProcessed = framesProcessed - 1
        ActiveSheet.Range("A1").Value = framesProcessed
    End If
End Sub

'Draws one image on the userform onto the worksheet
Private Sub drawFrame(startingRange, X, Y, index)     'draw frame given picture
    Dim i, j, r As Integer
    Dim xOffset, yOffset, xStep, yStep As Integer
    Dim grey As Long: grey = 15790320                  'default colour of userform
    
    'Pixel offset for the image to keep track of which pixel to draw next
    xStep = (pRight - pLeft) / X                      'X is number of cells in x direction
    yStep = (pBot - pTop) / Y                         'Y is number of cells in y direction
    
    'Image offset to keep track of where to start looking for pixels
    xOffset = pLeft
    yOffset = pTop
    
    'Load the given picture (index) from the frames folder (created by the Python script)
    RenderForm.Image.Picture = LoadPicture(ActiveWorkbook.Path & "\frames\" & index & ".jpg")
    'Paint the image onto the userform
    Me.Repaint
    
    'Iterate through all the pixels in the image
    For i = 0 To Y
        For j = 0 To X
            'Get the color of the current pixel
            Dim color As Long: color = GetPixel(hDC, xOffset + j * xStep, yOffset + i * yStep)
            
            'If the pixel is presumably from the userform (i.e. is the same color) or there is no color, turn the pixel black (looks better than grey)
            'This step is optional; one downside is that it sometimes creates black dots on the actual video when the color is too close to userform grey
            If color = grey Or color = -1 Then
                color = 1
                startingRange.Offset(i, j).Interior.color = color
            'If not grey or no color, call a function to convert the color and transpose it on the worksheet at some offset
            Else
                Dim temp As Variant: temp = ColorConvert(color)
                startingRange.Offset(i, j).Interior.color = RGB(CInt(temp(0)), CInt(temp(1)), CInt(temp(2)))
            End If
        Next j
    Next i
End Sub

'Excel has a hard limit of 64000 different unique cell formats; must convert 2^24 colors of JPG files down to the 64000 maximum by rounding
Private Function ColorConvert(color As Long) As Variant
    Dim RGB As Variant: RGB = LongToRGB(color)
    
    'Divide by 8 and round the decimal so that only multiples of 8 exist for each value of R, G, and B; then multiply back up to 8 to match RGB format
    'Reduces the number of possible colors down to 2^15, or 32768; should never run out of unique cell formats this way since we're only formatting with color
    For i = 0 To 2
        RGB(i) = Int(RGB(i) / 8) * 8
    Next i
    
    ColorConvert = RGB
End Function

'Convert a color in long format to a three-cell array with R, G, and B values in the cells
Private Function LongToRGB(color As Long) As String()
    Dim red, green, blue
    
    'Convert Decimal Color Code to RGB
    red = (color Mod 256)
    green = (color \ 256) Mod 256
    blue = (color \ 65536) Mod 256
    
    'Return RGB Code
    LongToRGB = Split(red & "," & green & "," & blue, ",")
End Function


