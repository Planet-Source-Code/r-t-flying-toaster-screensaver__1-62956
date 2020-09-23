VERSION 5.00
Begin VB.Form frmSaver 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Animation Code"
   ClientHeight    =   1665
   ClientLeft      =   3660
   ClientTop       =   2115
   ClientWidth     =   2280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSaver.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   111
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   152
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   960
      Left            =   600
      Picture         =   "frmSaver.frx":000C
      Top             =   360
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bStarted As Boolean
Private MouseX As Integer
Private MouseY As Integer
Private PwdOn As Boolean
Private PWProtect%

' ### END SCREENSAVER CONDITION
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not PreviewWnd Then Unload Me
End Sub

' ### END SCREENSAVER CONDITION
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not PreviewWnd Then Unload Me
End Sub

' ### END SCREENSAVER CONDITION
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not PreviewWnd Then
        If PwdOn = True Then
            PwdOn = False
            MouseX = X
            MouseY = Y
            Exit Sub
        End If
    
        If bStarted Then
            If MouseX <> X Or MouseY <> Y Then
                Unload Me
            Else
                MouseX = X
                MouseY = Y
            End If
        Else
            MouseX = X
            MouseY = Y
            bStarted = True
        End If
    End If
End Sub

' ### Password check (if applicable)
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If PWProtect And Environ("OS") <> "Windows_NT" Then
        PwdOn = True
        Dim PassChck As Boolean
        
        Do
        Loop Until ShowCursor(True) > 5
        
        PassChck = VerifyScreenSavePwd(Me.hwnd)
        If PassChck = False Then
            Do
            Loop Until ShowCursor(False) < -5
            
            Cancel = True
        End If
    End If
End Sub

' ### Clean up time!
Private Sub Form_Unload(Cancel As Integer)
    Do
    Loop Until ShowCursor(True) > 5
    
    If PWProtect And Environ("OS") <> "Windows_NT" Then
       SystemParametersInfo SPI_SCREENSAVERRUNNING, 0&, 0&, 0&
    End If
    
    Unload frmConfig
    Unload frmSaver
    Unload frmWork
    
    Set frmSaver = Nothing
    Set frmWork = Nothing
    Set frmConfig = Nothing
    End
End Sub
