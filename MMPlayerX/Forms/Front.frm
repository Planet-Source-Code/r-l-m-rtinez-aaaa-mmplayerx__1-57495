VERSION 5.00
Begin VB.Form frmCaratula 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   " Cover Front"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7125
   FontTransparent =   0   'False
   Icon            =   "Front.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picfondo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3555
      Left            =   -15
      ScaleHeight     =   3555
      ScaleWidth      =   4800
      TabIndex        =   1
      Top             =   0
      Width           =   4800
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1845
      Left            =   3765
      ScaleHeight     =   1845
      ScaleWidth      =   2100
      TabIndex        =   0
      Top             =   3615
      Visible         =   0   'False
      Width           =   2100
   End
End
Attribute VB_Name = "frmCaratula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Load()
On Error Resume Next
   
  bolCaratulaShow = True
   
  Me.Caption = LineLanguage(41)
  Me.Icon = frmMain.Icon
 '// si el album tiene caratula mostrarla
 If Trim(strRutaCaratula) <> "" Then
   Picture1.Picture = LoadPicture(strRutaCaratula)
   Picture1.AutoSize = True: Me.Width = Picture1.Width: Me.Height = Picture1.Height
 Else
   '// si no tiene caratula el album mostrar el default logo
   Picture1.Picture = frmPopUp.picDefaultLogo.Picture
   Picture1.AutoSize = True:  Me.Width = Picture1.Width: Me.Height = Picture1.Height
 End If
    Me.Left = (Screen.Width - Me.Width) / 2 '// centrar form
    Me.Top = (Screen.Height - Me.Height) / 2

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Resize()
 Mover_Form
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Sub Mover_Form()
 '// ajustar la imagen al ancho alto del form
 picfondo.Width = Me.ScaleWidth + 100
 picfondo.Height = Me.ScaleHeight
 picfondo.PaintPicture Picture1.Picture, 0, 0, picfondo.ScaleWidth, picfondo.ScaleHeight, 0, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
 bolCaratulaShow = False
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub picfondo_DblClick()
 '// ajustar el formulario al ancho-alto original de la caratula
   Picture1.AutoSize = True
   Me.Width = Picture1.Width
   Me.Height = Picture1.Height

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
