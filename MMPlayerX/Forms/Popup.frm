VERSION 5.00
Begin VB.Form frmPopUp 
   Caption         =   "MMPlayerX v. 2.0"
   ClientHeight    =   30
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Popup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   Begin VB.ListBox lstLanguage 
      Height          =   645
      Left            =   105
      TabIndex        =   2
      Top             =   3300
      Width           =   7470
   End
   Begin VB.PictureBox picDefaultLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   1230
      Picture         =   "Popup.frx":000C
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   338
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
      Width           =   5070
   End
   Begin VB.FileListBox fileBmps 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   225
      Hidden          =   -1  'True
      Left            =   1935
      Pattern         =   "*.jpg;*.bmp"
      System          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuMenuPrincipal 
      Caption         =   "MenuPrincipal"
      Begin VB.Menu mnuNuevaBusqueda 
         Caption         =   "Nueva Busqueda ..."
      End
      Begin VB.Menu mnuA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCFront 
         Caption         =   "Cover Front"
         Begin VB.Menu mnuCambiarListaCaratula 
            Caption         =   "L   Cambiar Lista Rep/Caratula"
         End
         Begin VB.Menu mnuWallpapper 
            Caption         =   "    Colocar Caratula como Wallpaper"
         End
         Begin VB.Menu mnuMCaratula 
            Caption         =   "    Maximizar Caratula"
         End
      End
      Begin VB.Menu mnuB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBrowsers 
         Caption         =   "Browsers"
         Begin VB.Menu mnuExplorar 
            Caption         =   "    Explorar ..."
         End
         Begin VB.Menu mnuExpAlbum 
            Caption         =   "    Explorar Album(s)"
         End
         Begin VB.Menu mnuZ 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTagEditor 
            Caption         =   "   Tag Editor"
         End
         Begin VB.Menu mnuLyrics 
            Caption         =   "    Lyrics"
         End
      End
      Begin VB.Menu mnuSepator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListSpec 
         Caption         =   "Visualization"
         Begin VB.Menu mnuMaxSpec 
            Caption         =   "Show Visualization"
         End
         Begin VB.Menu mnuShowSpec 
            Caption         =   "Configure Visualization"
         End
      End
      Begin VB.Menu mnuC 
         Caption         =   "-"
      End
      Begin VB.Menu mnuControles 
         Caption         =   "Controles de Reproduccion"
         Begin VB.Menu mnuVolumen 
            Caption         =   "   Volumen"
            Begin VB.Menu mnuSubirVolumen 
               Caption         =   "+   Subir Volumen"
            End
            Begin VB.Menu mnuBajarVolumen 
               Caption         =   "-   Bajar Volumen"
            End
         End
         Begin VB.Menu mnuD 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTrackAnterior 
            Caption         =   "Z   Track Anterior"
         End
         Begin VB.Menu mnuReproducir 
            Caption         =   "X   Reproducir"
         End
         Begin VB.Menu mnuPausa 
            Caption         =   "C   Pausa"
         End
         Begin VB.Menu mnuDetener 
            Caption         =   "V   Detener"
         End
         Begin VB.Menu mnuSigTrack 
            Caption         =   "B   Siguiente Track"
         End
         Begin VB.Menu mnuE 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSigAlbum 
            Caption         =   ">   Siguiente Album/Folder"
         End
         Begin VB.Menu mnuAnteriorAlbum 
            Caption         =   "<   Anterior Album/Folder"
         End
         Begin VB.Menu mnuF 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIntro 
            Caption         =   "I   Intro 10 Segundos"
         End
         Begin VB.Menu mnuRepetir 
            Caption         =   "R   Repetir Track"
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "S   Silencio"
         End
         Begin VB.Menu mnuj 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOrdenAleatorio 
            Caption         =   "  Orden Aleatorio"
            Begin VB.Menu mnuAleatorioActAlbum 
               Caption         =   "Q   Actual Album/Folder"
            End
            Begin VB.Menu mnuAleatorioTodaColec 
               Caption         =   "W   Toda la Colección"
            End
         End
         Begin VB.Menu mnug 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAtras5Seg 
            Caption         =   "A   Atras 5 Segundos"
         End
         Begin VB.Menu mnuAdelante5Seg 
            Caption         =   "D   Adelante 5 Segundos"
         End
      End
      Begin VB.Menu mnuh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones 
         Caption         =   "Opciones ..."
      End
      Begin VB.Menu mnuSkins 
         Caption         =   "Skins"
         WindowList      =   -1  'True
         Begin VB.Menu mnuExpSkins 
            Caption         =   "<<  Explorador de Skins >>"
         End
         Begin VB.Menu mnuXXX 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSkinsAdd 
            Caption         =   ""
            Index           =   1
         End
      End
      Begin VB.Menu mnuWOpacity 
         Caption         =   "Window Opacity"
         Begin VB.Menu mnuAlpha 
            Caption         =   "100%"
            Index           =   0
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "90%"
            Index           =   1
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "80%"
            Index           =   2
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "70%"
            Index           =   3
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "60%"
            Index           =   4
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "50%"
            Index           =   5
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "40%"
            Index           =   6
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "30%"
            Index           =   7
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "20%"
            Index           =   8
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "10%"
            Index           =   9
         End
         Begin VB.Menu mnuAlphaPer 
            Caption         =   "Personalizar..."
         End
      End
      Begin VB.Menu mnui 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "Acerca de ..."
      End
      Begin VB.Menu mnuXX 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuMainSpec 
      Caption         =   "MainSpectrum"
      Begin VB.Menu mnuSpecNone 
         Caption         =   "None Visualisation"
      End
      Begin VB.Menu mnuSpecBars 
         Caption         =   "Spectrum Analyzer"
      End
      Begin VB.Menu mnuSpecOsc 
         Caption         =   "Oscilloscope"
      End
   End
   Begin VB.Menu mnuMainAlbum 
      Caption         =   "MainAlbum"
      Begin VB.Menu mnuAlbumTags 
         Caption         =   "Edit Album Tags"
      End
      Begin VB.Menu mnuAlbumBrowser 
         Caption         =   "Explore in Album Browser"
      End
      Begin VB.Menu mnuAlbumExp 
         Caption         =   "Explore in Explorer.exe"
      End
      Begin VB.Menu mnuAlbumPlay 
         Caption         =   "Play"
      End
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_GotFocus()
    ' When the user presses the taskbar button, this form gets
    ' the focus, so we shift the focus to the frmPad
    'frmMain.SetFocus
End Sub

Private Sub Form_Load()
    ' The wrapper window is practically invisible to the user.
    ' Older version of WindowBlinds have a bug that causes this
    ' window to appear.
    Me.Top = -10000
    Me.Icon = frmMain.Icon
End Sub

Private Sub Form_Resize()
    ' Change frmmain's state according to changes made to this
    ' form using the taskbar
    
 If bLoading = True Then Exit Sub
   
    If Me.WindowState = vbMinimized Then
           If bolCaratulaShow = True Then frmCaratula.Hide
           If bolDirectoriosShow = True Then frmDirectorios.Hide
           If bolOpcionesShow = True Then frmOpciones.Hide
           If bolAcercaShow = True Then frmAcerca.Hide
           If bolTagsShow = True Then frmTags.Hide
           If bolLyricsShow = True Then frmLyrics.Hide
           If bolSplashScreen = True Then frmSplash.Hide
           If bolVisShow = True Then frmSpectrum.Hide
           If bolSearchShow = True Then frmSearch.Hide
           frmMain.Hide
    Else
           If bolAcercaShow = True Then frmAcerca.Show
           If bolCaratulaShow = True Then frmCaratula.Show
           If bolDirectoriosShow = True Then frmDirectorios.Show
           If bolOpcionesShow = True Then frmOpciones.Show
           If bolLyricsShow = True Then frmLyrics.Show
           If bolTagsShow = True Then frmTags.Show
           If bolVisShow = True Then frmSpectrum.Show
           If bolSearchShow = True Then frmSearch.Show
           
        frmMain.WindowState = Me.WindowState
        frmMain.Visible = True
        If bolSplashScreen = True Then frmSplash.Show
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload frmMain
 End
End Sub

Private Sub mnuAcercaDe_Click()
 If bolAcercaShow = True Then
   frmAcerca.ZOrder 0
 Else
   frmAcerca.Show
 End If
End Sub


Private Sub mnuAdelante5Seg_Click()
 frmMain.Five_Seg_Forward
End Sub

Private Sub mnuAlbumBrowser_Click()
 On Error GoTo HELLBITCH

  If bolDirectoriosShow = False Then
    frmDirectorios.Show
    frmDirectorios.TreeAlbums.Nodes(CStr(iIdAlbumRC & " \ " & 0)).Selected = True
  Else
    frmDirectorios.Show
    frmDirectorios.ZOrder 0
    frmDirectorios.TreeAlbums.Nodes(CStr(iIdAlbumRC & " \ " & 0)).Selected = True
  End If

Exit Sub
HELLBITCH:

End Sub

Private Sub mnuAlbumExp_Click()
  On Error Resume Next
  
  Shell "explorer.exe " & frmMain.btnAlbum(iIdAlbumRC).ToolTipText, vbMaximizedFocus

End Sub

Private Sub mnuAlbumPlay_Click()
  On Error Resume Next
  If iIdAlbumRC = intActiveAlbum Then Exit Sub
  frmMain.Play_Album iIdAlbumRC
End Sub

Private Sub mnuAlbumTags_Click()
  On Error Resume Next
  
  If bolTagsShow = True Then
     frmTags.FileTags.Path = frmMain.btnAlbum(iIdAlbumRC).ToolTipText
     frmTags.Make_List_Ref
     frmTags.FileTags.Selected(0) = True
     frmTags.ZOrder 0
  Else
     frmTags.Show
     frmTags.FileTags.Path = frmMain.btnAlbum(iIdAlbumRC).ToolTipText
     frmTags.Make_List_Ref
     frmTags.FileTags.Selected(0) = True
     frmTags.ZOrder 0
  End If

End Sub

Private Sub mnuAleatorioActAlbum_Click()
   frmMain.Randomize_Click True, False
End Sub

Private Sub mnuAleatorioTodaColec_Click()
   frmMain.Randomize_Click False, False
End Sub

Private Sub mnuAlpha_Click(Index As Integer)
On Error GoTo Hell
 Dim tAlpha
 Dim i As Integer
   tAlpha = mnuAlpha(Index).Caption
   tAlpha = Left(tAlpha, Len(tAlpha) - 1)
  Call SetWindowLong(frmMain.hwnd, GWL_EXSTYLE, GetWindowLong(frmMain.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
  Call SetLayeredWindowAttributes(frmMain.hwnd, 0, (255 * tAlpha) / 100, LWA_ALPHA)
  mnuAlpha(Index).Checked = True
  OpcionesMusic.Alpha = tAlpha
    
    frmPopUp.mnuAlphaPer.Caption = LineLanguage(37)
    frmPopUp.mnuAlphaPer.Checked = False
  For i = 0 To 9
   If i <> Index Then mnuAlpha(i).Checked = False
  Next i
 Exit Sub
Hell:
End Sub

Private Sub mnuAlphaPer_Click()
   frmOpciones.Show
   frmOpciones.Select_Option 1
   frmOpciones.TSAppConfig.Tabs(3).Selected = True
End Sub

Private Sub mnuAnteriorAlbum_Click()
 frmMain.Previous_Album
End Sub

Private Sub mnuAtras5Seg_Click()
 frmMain.Five_Seg_Backward
End Sub

Private Sub mnuBajarVolumen_Click()
 '// bajar volumen
  frmMain.Form_KeyPress 45
End Sub

Private Sub mnuCambiarListaCaratula_Click()
frmMain.Front_Click
End Sub


Private Sub mnuDetener_Click()
 frmMain.Stop_Player
End Sub

Private Sub mnuExpAlbum_Click()
If bolDirectoriosShow = False Then
 frmDirectorios.Show
Else
 Load_Language_Directorios
 frmDirectorios.Show
' frmDirectorios.TreeAlbums.Nodes(CStr(intActiveAlbum & " \ " & frmMain.ListRep.ListIndex)).Selected = True
 frmDirectorios.ZOrder 0
 
End If
End Sub

Private Sub mnuExplorar_Click()
On Error Resume Next
Dim x As Long
Dim strPathExplore As String
 If TotalAlbumS = 0 Then
   strPathExplore = tAppConfig.AppPath
 Else
   strPathExplore = frmMain.btnAlbum(intActiveAlbum).ToolTipText
 End If
x = Shell("explorer.exe " & strPathExplore, vbMaximizedFocus)

End Sub

Private Sub mnuExpSkins_Click()
   frmOpciones.Select_Option 2
   frmOpciones.Show
End Sub

Private Sub mnuIntro_Click()
  frmPopUp.mnuIntro.Checked = Not frmPopUp.mnuIntro.Checked
  frmMain.Intro
End Sub

Private Sub mnuLyrics_Click()
 If bolLyricsShow = True Then
   frmLyrics.ZOrder 0
 Else
   frmLyrics.Load_config_KARAOKE
 End If
End Sub

Private Sub mnuMaxSpec_Click()
 If bolVisShow = False Then frmMain.Stop_Draw_Spectrum
  bolVisShow = True
  frmSpectrum.Show
End Sub

Private Sub mnuMCaratula_Click()
If bolCaratulaShow = False Then
 frmCaratula.Show
Else
 frmCaratula.ZOrder 0
End If
End Sub


Private Sub mnuNuevaBusqueda_Click()
  frmSearch.Show
End Sub

Private Sub mnuOpciones_Click()
   frmOpciones.Select_Option 1
   frmOpciones.Show
 
End Sub
Private Sub mnuPausa_Click()
 frmMain.Pause_Play
End Sub

Private Sub mnuRepetir_Click()
 frmPopUp.mnuRepetir.Checked = Not frmPopUp.mnuRepetir.Checked
 frmMain.Player_Repeat
End Sub

Private Sub mnuReproducir_Click()
 frmMain.Play
End Sub
Private Sub mnuSalir_Click()
 Unload frmMain
End Sub

Private Sub mnuShowSpec_Click()
   
   frmOpciones.Select_Option 8
   frmOpciones.Show
    
End Sub

Private Sub mnuSigAlbum_Click()
 frmMain.Next_Album
End Sub

Private Sub mnuSigTrack_Click()
 frmMain.Next_Track
End Sub

Private Sub mnuSilencio_Click()
  frmPopUp.mnuSilencio.Checked = Not frmPopUp.mnuSilencio.Checked
  frmMain.Player_Mute
End Sub

Private Sub mnuSkinsAdd_Click(Index As Integer)
 On Error Resume Next
 Dim Skins As String, MiRuta As String
 Dim i As Integer
 Skins = Trim(mnuSkinsAdd(Index).Caption)
 '// si es el mismo skin salir
 If Skins = "" Then Exit Sub
 
 '// chekar si existe la carpeta
 MiRuta = tAppConfig.AppConfig & "Skins\"
 If Dir(MiRuta & Skins, vbDirectory) = "" Then Exit Sub

 If LCase(Skins) = LCase(tAppConfig.Skin) Then Exit Sub
   '// seleccionar el skin
   For i = 1 To mnuSkinsAdd.count
      If i = Index Then
        mnuSkinsAdd(Index).Checked = True
      Else
        mnuSkinsAdd(i).Checked = False
      End If
   Next i
 

    frmMain.Visible = False

    '// Cambiar el skin
    Change_Skin Skins
    '// ajustar los bordes
    Form_Mini_Normal
    
    Change_Mask bMiniMask, False
    
    frmOpciones.lblSkin(2).Caption = Skins
    frmOpciones.ListaSkins.Selected(Index) = True
    frmOpciones.ListaSkins.ListIndex = Index - 1
    
    frmMain.Show_ScrollBar

    frmMain.Visible = True

End Sub

Public Sub mnuSpecBars_Click()
With frmPopUp
   .mnuSpecBars.Checked = True
   .mnuSpecNone.Checked = False
   .mnuSpecOsc.Checked = False
 End With
End Sub

Public Sub mnuSpecNone_Click()
 With frmPopUp
   .mnuSpecBars.Checked = False
   .mnuSpecNone.Checked = True
   .mnuSpecOsc.Checked = False
 End With
End Sub

Public Sub mnuSpecOsc_Click()
With frmPopUp
   .mnuSpecBars.Checked = False
   .mnuSpecNone.Checked = False
   .mnuSpecOsc.Checked = True
 End With
End Sub

Private Sub mnuSubirVolumen_Click()
 frmMain.Form_KeyPress 43 '// subir volumen

End Sub

Private Sub mnuTagEditor_Click()
 If bolTagsShow = True Then
    frmTags.Load_Album_Tags
    frmTags.ZOrder 0
 Else
   frmTags.Show
 End If
End Sub

Private Sub mnuTrackAnterior_Click()
 frmMain.Previous_Track
End Sub

Private Sub mnuWallpapper_Click()
 If frmMain.ListRep.ListCount = 0 Then Exit Sub
  mnuWallpapper.Checked = Not mnuWallpapper.Checked
   If mnuWallpapper.Checked = True Then
     ConfigurarWallpaper
   Else
     PoneRWallPapeROriginaL
   End If
  
End Sub


