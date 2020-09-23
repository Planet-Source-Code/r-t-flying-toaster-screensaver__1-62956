Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
Dim args As String

iSpeed = GetSetting(APP_NAME, "Settings", "Speed", 25)
iToasters = GetSetting(APP_NAME, "Settings", "Toasters", 4)
iToasts = GetSetting(APP_NAME, "Settings", "Toasts", 4)
iSprites = iToasters + iToasts
If iSprites <= 0 Or iSprites > 60 Then iToasters = 4: iToasts = 4

args = UCase$(Trim$(Command$))
    Select Case Mid$(args, 1, 2)
        Case "/C" ' Configuration
            frmConfig.Show
        Case "/S" ' RUN IT!
            ' exit if screen saver is running already
            CheckIfRunning
            
            Load frmSaver
            frmSaver.Show                      'Bring up the main display form
            frmSaver.WindowState = 2           'make sure it's maximized
            Initialize
             
        Case "/P" ' Preview
            Load frmSaver
            PreviewWnd = True
            frmSaver.MousePointer = 0
            Adopt frmSaver.hwnd, Val(Mid(Interaction.Command, 4)) 'Put it in the preview window
            frmSaver.Move 0, 0, 152 * 15, 112 * 15                'Re dimention it to the preview window size
            frmSaver.Image1.Visible = True
            frmSaver.Show
        Case "/A" ' Password Prompt
        Case Else
            frmConfig.Show
    End Select
End Sub

' ### Timer Sub
Sub Controller()
    Do
        DoAnimation ' redraw screen elements
        DoEvents ' allow system to not crawl to a snail's pace
        If iSpeed > 0 Then Sleep (iSpeed) 'delay to slow down the animation
    Loop
End Sub

Sub DoAnimation()
'Reset Background
    DDC = frmWork.WorkScr.hdc
    SDC = frmWork.CleanScreen.hdc

    For SpCnt = 0 To iSprites ' Erase all Sprites from screen
        With Sprites(SpCnt)
            Xo = .Xp
            Yo = .Yp
            Xs = .Width
            Ys = .Height
            res = BitBlt(DDC, Xo, Yo, Xs, Ys, SDC, Xo, Yo, SRCCOPY)
        End With
    Next

    For SpCnt = 0 To iSprites ' Move Screen elements, 1 per loop
        With Sprites(SpCnt)
            If SpCnt < iSprites / 2 Then
                .FrameDelay = .FrameDelay + 1
                If .FrameDelay > 5 Then
                    If .FrIndex < 1 Then
                        .FrameNext = 1
                    ElseIf .FrIndex > 2 Then
                        .FrameNext = -1
                    End If
                    
                    .FrIndex = .FrIndex + .FrameNext
                    .FrameDelay = 0
                End If
            Else
                .FrIndex = 4
            End If
            
        ' collision right
            If .Xp > RightX Or .Yp > BottomY Or .Xp < -92 Or .Yp < -92 Then
                If Rnd * 1 < 0.5 Then
                    .Yp = Rnd * (BottomY \ 3)
                    .Xp = -92
                Else
                    .Yp = -92
                    .Xp = Rnd * (RightX \ 3)
                End If
        
                .Xmove = (.Xmove * Rnd(4) + 2)
            End If
            
            .Xp = .Xp + .Xmove
            .Yp = .Yp + .Ymove
            
            ' AND mask to WorkScr
            Ind = .FrIndex
            Xo = .Xp
            Yo = .Yp
            Xs = .Width - 10
            Ys = .Height - 10
            XSrc = .ImageMaskX(Ind)
            YSrc = .ImageMaskY(Ind)
            DDC = frmWork.WorkScr.hdc
            SDC = frmWork.Master.hdc
            res = BitBlt(DDC, Xo, Yo, Xs, Ys, SDC, XSrc, YSrc, SRCAND)
            
            ' OR image to WorkScr
            XSrc = .ImageSrcX(Ind)
            YSrc = .ImageSrcY(Ind)
            res = BitBlt(DDC, Xo, Yo, Xs, Ys, SDC, XSrc, YSrc, SRCPAINT)
        End With
    Next
    
    ' Copy all sprites to screen
    For SpCnt = 0 To iSprites
        With Sprites(SpCnt)
            Xo = .Xp
            Yo = .Yp
            Xs = .Width
            Ys = .Height
            DDC = frmSaver.hdc
            SDC = frmWork.WorkScr.hdc
            res = BitBlt(DDC, Xo, Yo, Xs, Ys, SDC, Xo, Yo, SRCCOPY)
        End With
    Next
End Sub


Sub Initialize()
'** copying screen Variables
Dim DestDC As Long               'Destination DC
Dim XPixels As Long              'Transfer Picture Width
Dim YPixels As Long              'Trandfer Picture Height
Dim destX As Long                'Destination X Position
Dim destY As Long                'Destination Y Position
Dim srcDC As Long                'Source DC
Dim SrcX As Long                 'Source X Position
Dim SrcY As Long                 'Source Y Position
Dim RasterOp As Long             'Raster Operation to Perform (Copy, And, Or)

BottomY = frmSaver.ScaleHeight     'Set Bottom Screen Limit
RightX = frmSaver.ScaleWidth       'Set Right Screen Limit


'** Make CleanScreen and WorkSpace Screens are the same size as frmsaver Screen
frmWork.CleanScreen.Width = frmSaver.Width
frmWork.CleanScreen.Height = frmSaver.Height
frmWork.WorkScr.Width = frmSaver.Width
frmWork.WorkScr.Height = frmSaver.Height
frmSaver.Refresh  'Make sure frmsaver is current

DoEvents

SetWindowPos frmSaver.hwnd, -1, 0, 0, 0, 0, &H2 Or 1 ' -1 = HWND_TOPMOST, &H2 = SWP_NOMOVE, 1 = SWP_NOSIZE
Do Until ShowCursor(False) < -5
Loop

'clear these out so they don't take up any space anymore
DestDC = 0
XPixels = 0
YPixels = 0
srcDC = 0
SrcX = 0
SrcY = 0
RasterOp = 0

Randomize (Timer) 'initialize the random number generator

'the following code initializes my sprites
Dim z As Long ' generic to count through the sprites
Dim r As Integer 'Random number from 0 to 5.???

'Setup the initial positions for the sprites (randomly)
For z = 0 To iSprites - 1  'loop through my sprites and setup ramdom start positions
    With Sprites(z)
        .Xp = Rnd * RightX 'set sprites initial position (horizontal)
        .Yp = Rnd * BottomY 'Set sprites initial position (vertical)
    
        .Xmove = 3
        .Ymove = 1
        .Width = 128      'the width of the sprites used
        .Height = 128     'the height of the sprites used
        .Show = True     'enable the show flag
        .Frames = 1      'not really used in this demo but tells the program how many frames a sprite has
        
        .ImageSrcX(0) = 0   'frame1
        .ImageSrcY(0) = 0
        .ImageSrcX(1) = 128
        .ImageSrcY(1) = 0   'frame2
        .ImageSrcX(2) = 256
        .ImageSrcY(2) = 0   'frame3
        .ImageSrcX(3) = 384
        .ImageSrcY(3) = 0   'frame4
        
        .ImageMaskX(0) = 0   'frame1 mask
        .ImageMaskY(0) = 128
        .ImageMaskX(1) = 128 'frame2 mask
        .ImageMaskY(1) = 128
        .ImageMaskX(2) = 256 'frame3 mask
        .ImageMaskY(2) = 128
        .ImageMaskX(3) = 384 'frame4 mask
        .ImageMaskY(3) = 128
        
        .ImageSrcX(4) = 512 ' toast
        .ImageSrcY(4) = 0
        .ImageMaskX(4) = 512
        .ImageMaskY(4) = 128
    End With
Next

'end of sprite initialization

Controller  ' Startup the animation (controller routine below)

End Sub

Private Function GetHwndFromCommand(ByVal args As String) As Long
Dim argslen As Integer
Dim i As Integer
Dim ch As String

    'take the rightmost numeric characters.
    args = Trim$(args)
    argslen = Len(args)
    For i = argslen To 1 Step -1
        ch = Mid$(args, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
    Next i

    GetHwndFromCommand = CLng(Mid$(args, i + 1))
End Function
