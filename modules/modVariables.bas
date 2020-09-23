Attribute VB_Name = "modShared"
Option Explicit

' ### CONSTANTS
Public Const APP_NAME = "FlyingToasterSaver"

Public iSprites As Integer             ' Total Sprite count
Public iToasters As Integer            ' Number of Toasters
Public iToasts As Integer              ' Number of pieces of toast
Public iSpeed As Integer              'Animation Speed

Type myCritter                     'Define Sprite User Defined Object
    FrIndex As Integer                 'Current Frame Index
    Xp As Integer                      'X Position on Display Screen
    Yp As Integer                      'Y Postion on Display Screen
    Width As Integer                   'Width of Critter in pixels
    Height As Integer                  'Height of Critter in pixels
    Xmove As Integer                   'Amount to Move Horizontally
    Ymove As Integer                   'amount to move vertically
    Frames As Integer                  'Amount of Frames in Sprite Set
    Show As Boolean                    'Display or not to display (true=display)
    'Increase array element number for frames amount
    ImageSrcX(20) As Integer            'X Position in Source File Main Image
    ImageSrcY(20) As Integer            'Y Position in Source File Main Graphic
    ImageMaskX(20) As Integer           'X Position in Source File Main Image (Mask)
    ImageMaskY(20) As Integer           'Y Position in Source File Main Graphic (Mask)
    FrameDelay As Integer
    FrameNext As Integer
End Type
Public Sprites(60) As myCritter         'Change Number to amount of Sprites needed

' ### Image API
Public Declare Function StretchBlt& Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal XSrc&, ByVal YSrc&, ByVal nSrcWidth&, ByVal nSrcHeight&, ByVal dwRop&)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
    Public Const SRCCOPY = &HCC0020
    Public Const SRCPAINT = &HEE0086
    Public Const SRCAND = &H8800C6

' ### Screensaver Password API
Public Declare Function VerifyScreenSavePwd Lib "password.cpl" (ByVal hwnd&) As Boolean
Public Declare Function PwdChangePassword& Lib "mpr" Alias "PwdChangePasswordA" (ByVal lpcRegkeyname$, ByVal hwnd&, ByVal uiReserved1&, ByVal uiReserved2&)
Public Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
    Public Const SPI_SCREENSAVERRUNNING = 97&

' ### Previous Instance API
Public Declare Function FindWindow& Lib "User32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)

' ### Miscellaneous API
Public Declare Function ShowCursor& Lib "User32" (ByVal bShow&)
Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

' ### BitBlt Variables
Public Ind As Long
Public Xo As Long
Public Yo As Long
Public Xs As Long
Public Ys As Long
Public XSrc As Long
Public YSrc As Long
Public DDC As Long
Public SDC As Long
Public res As Long

' ### Miscellaneous Variables
Public SpCnt As Integer               ' Counter used for Cycling through Sprites
Public RightX As Integer              ' Screen Width
Public BottomY As Integer             ' Screen Height
Public PreviewWnd As Boolean

' ### Preview Window API
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Const GWL_STYLE As Long = -16
    Private Const WS_CHILD As Long = &H40000000
    Private Const GWL_HWNDPARENT As Long = -8

Public Sub Adopt(ChildWnd As Variant, ParentWnd As Variant)
    Dim Style As Variant
    Style = GetWindowLong(ChildWnd, GWL_STYLE)
    SetWindowLong ChildWnd, GWL_STYLE, Style Or WS_CHILD 'Make our window a child
    SetParent ChildWnd, ParentWnd 'Make the parent adopt it
    SetWindowLong ChildWnd, GWL_HWNDPARENT, ParentWnd 'let the kid know it's parent
End Sub

Public Function CheckIfRunning() As Boolean
    Dim bRet As Boolean
    bRet = False
    
    If Not App.PrevInstance Then bRet = True
    If FindWindow(vbNullString, APP_NAME) Then bRet = True
    
    CheckIfRunning = bRet
End Function
