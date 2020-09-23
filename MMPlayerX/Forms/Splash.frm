VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblSplash 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading... "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   45
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1935
      Width           =   4965
   End
   Begin VB.Image imgLogo 
      Height          =   2250
      Left            =   0
      MousePointer    =   99  'Custom
      Top             =   120
      Width           =   5070
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
bolSplashScreen = True
imgLogo.Picture = frmPopUp.picDefaultLogo.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
 bolSplashScreen = False
End Sub


