VERSION 5.00
Begin VB.Form frmWork 
   Appearance      =   0  'Flat
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Valentine"
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmWork.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   4740
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Master 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3885
      Left            =   0
      Picture         =   "frmWork.frx":000C
      ScaleHeight     =   259
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   633
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.PictureBox WorkScr 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   840
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3930
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox CleanScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   45
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3945
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
