VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6150
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtToasts 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Text            =   "8"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtToasters 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Text            =   "8"
      Top             =   1900
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtSpeed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   0
      Text            =   "25"
      ToolTipText     =   "Less is more ;-)"
      Top             =   1515
      Width           =   615
   End
   Begin VB.Label lblToast 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Toast Count"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   2325
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Toaster Count"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   1940
      Width           =   1095
   End
   Begin VB.Label lblSpeed 
      BackColor       =   &H00000000&
      Caption         =   "Toaster Speed"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblWhy 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "for UNEASYsilence.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Flying Toaster"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblCredit 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Recreated by Locohozt"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   2895
   End
   Begin VB.Image img1 
      Height          =   960
      Left            =   1080
      Picture         =   "frmConfig.frx":0E42
      Top             =   240
      Width           =   1140
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ### Load Screensaver settings into Form elements
Private Sub Form_Load()
    txtSpeed = iSpeed
    txtToasters = iSprites - iToasts
    txtToasts = iSprites - iToasters
End Sub

' ### Save ScreenSaver Settings
Private Sub cmdOK_Click()
    Dim blnError As Boolean
    blnError = False
    
    If IsNumeric(txtSpeed) And IsNumeric(txtToasters) And IsNumeric(txtToasts) Then
        If 1 <= txtSpeed And txtSpeed <= 100 Then
            SaveSetting APP_NAME, "Settings", "Speed", txtSpeed
        Else
            blnError = True
            MsgBox "ERROR: For Speed we're looking for a number between 1 and 100 here, buddy.", vbCritical
        End If

        If 1 <= txtToasters And txtToasters <= 30 Then
            SaveSetting APP_NAME, "Settings", "Toasters", txtToasters
        Else
            blnError = True
            MsgBox "ERROR: For count we're looking for a number between 1 and 30 here, buddy.", vbCritical
        End If
        
        If 1 <= txtToasts And txtToasts <= 30 Then
            SaveSetting APP_NAME, "Settings", "Toasts", txtToasters
        Else
            blnError = True
            MsgBox "ERROR: The number of pieces of toast must be a number between 1 and 30 here, buddy.", vbCritical
        End If
    Else
        blnError = True
        MsgBox "ERROR: We need numbers in the textboxes here, buddy.", vbCritical
    End If
    
    If blnError = False Then End
End Sub

' ### Exit program with no settings changes
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub txtSpeed_GotFocus()
    txtSpeed.SelStart = 0
    txtSpeed.SelLength = Len(txtSpeed)
End Sub

Private Sub txtToasters_GotFocus()
    txtToasters.SelStart = 0
    txtToasters.SelLength = Len(txtToasters)
End Sub

Private Sub txtToasts_GotFocus()
    txtToasts.SelStart = 0
    txtToasts.SelLength = Len(txtToasts)
End Sub
