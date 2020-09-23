Attribute VB_Name = "mLanguage"
Option Explicit

'Public arryLanguage() As String

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  IDIOMA                                                                               |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Load_Language_Spanish()
With frmPopUp
.lstLanguage.Clear
.lstLanguage.AddItem "Language"
' MENU
.lstLanguage.AddItem " Nueva Busqueda"  ' 1
.lstLanguage.AddItem " Caratula" ' 2
.lstLanguage.AddItem "  Cambiar ListaRep / Caratula" ' 3
.lstLanguage.AddItem "  Colocar caratula como Wallpaper" ' 4
.lstLanguage.AddItem "  Maximizar Caratula" '  5
.lstLanguage.AddItem " Exploradores" '  6
.lstLanguage.AddItem "  Explorar archivos" ' 7
.lstLanguage.AddItem "   Explorador de Albums" ' 8
.lstLanguage.AddItem "   Editar Track(s) Tag" '  9
.lstLanguage.AddItem "   Karaoke" ' 10
.lstLanguage.AddItem " Visualización Studio" ' 11
.lstLanguage.AddItem "   Configurar Visualización" ' 12
.lstLanguage.AddItem "   Mostrar Visualización" ' 13
.lstLanguage.AddItem " Controles de Reproducción" ' 14
.lstLanguage.AddItem "   Volumen" ' 15
.lstLanguage.AddItem "+     Subir Volumen" ' 16
.lstLanguage.AddItem "-     Bajar Volumen" ' 17
.lstLanguage.AddItem "Z   Track Anterior" ' 18
.lstLanguage.AddItem "X   Reproducir" ' 19
.lstLanguage.AddItem "C   Pausar" ' 20
.lstLanguage.AddItem "V   Detener" ' 21
.lstLanguage.AddItem "B   Siguiente Track" ' 22
.lstLanguage.AddItem "<   Anterior Album / Folder" ' 23
.lstLanguage.AddItem ">   Siguiente Album / Folder" ' 24
.lstLanguage.AddItem "I   Intro 10 seg." ' 25
.lstLanguage.AddItem "R   Repetir Track" ' 26
.lstLanguage.AddItem "S   Silencio" ' 27
.lstLanguage.AddItem "   Orden aleatorio" ' 28
.lstLanguage.AddItem "Q     Actual Album / Folder" ' 29
.lstLanguage.AddItem "W     Todos los Albums" ' 30
.lstLanguage.AddItem "A   Atras 5 seg." ' 31
.lstLanguage.AddItem "D   Adelante 5 Seg." ' 32
.lstLanguage.AddItem " Opciones" ' 33
.lstLanguage.AddItem " Skins" ' 34
.lstLanguage.AddItem "   << Explorador de Skins >>" ' 35
.lstLanguage.AddItem " Transparencia" ' 36
.lstLanguage.AddItem "   Personalizar" ' 37
.lstLanguage.AddItem " Acerca de" ' 38
.lstLanguage.AddItem " Salir" ' 39
' ACERCA
.lstLanguage.AddItem " Acerca de MMPlayerX" ' 40
' CARATULA
.lstLanguage.AddItem "Caratula actual"
' EXPLORADOR DE ALBUMS
.lstLanguage.AddItem "Explorador de Albums"
.lstLanguage.AddItem "  Explorar archivos"
.lstLanguage.AddItem "  Editar Tags"
.lstLanguage.AddItem "  Reproducir"
' KARAOKE
.lstLanguage.AddItem "Karaoke"
.lstLanguage.AddItem "  [ Letras no Encontradas ]"
' MAIN
.lstLanguage.AddItem "    Menu"
.lstLanguage.AddItem "    Minimizar"
.lstLanguage.AddItem "    Change Mode" ' 50
.lstLanguage.AddItem "    Salir"
.lstLanguage.AddItem "  Selecciona el directorio a buscar."
.lstLanguage.AddItem "  Sin Visualización"
.lstLanguage.AddItem "  Analizador de Espectro"
.lstLanguage.AddItem "  Osiloscopio"
.lstLanguage.AddItem "  Editar Album Tags"
.lstLanguage.AddItem "  Explorar Exp. de Albums"
.lstLanguage.AddItem "  Explorar Explorer.exe"
.lstLanguage.AddItem "  Reproducir"
' TAGS
.lstLanguage.AddItem "Editor de Tags + Información MPEG" ' 60
.lstLanguage.AddItem "  Multiples tracks estan seleccionados, Selecciona los checkboxs para aplicar los cambios  a TODOS los archivos seleccionados."
.lstLanguage.AddItem "  Seleccionar Todo"
.lstLanguage.AddItem "  Tags"
.lstLanguage.AddItem "  Karaoke"
.lstLanguage.AddItem "  Agregar"
.lstLanguage.AddItem "  Deshacer"
.lstLanguage.AddItem "  Aceptar"
.lstLanguage.AddItem "  Cancelar"
.lstLanguage.AddItem "  Aplicar"
' OPCIONES
.lstLanguage.AddItem "Opciones" ' 70
.lstLanguage.AddItem "  Aceptar" ' 71
.lstLanguage.AddItem "  Cancelar" ' 72
.lstLanguage.AddItem "  Aplicar" ' 73
.lstLanguage.AddItem "Aplicación" ' 74
.lstLanguage.AddItem "  Trayectoria Configuración" ' 75
.lstLanguage.AddItem "    Trayectoria de skins y configuración:" ' 76
.lstLanguage.AddItem "    Explorar..." ' 77
.lstLanguage.AddItem "    Nota: Algunas opciones requieren que se reinicie la aplicación."
.lstLanguage.AddItem "    Memoria Libre (Fisica):" ' 79
.lstLanguage.AddItem "  Aplicación" ' 80
.lstLanguage.AddItem "    Lenguaje:" ' 81
.lstLanguage.AddItem "    Siempre arriba." ' 82
.lstLanguage.AddItem "    Mostrar Splash Screen." ' 83
.lstLanguage.AddItem "    Permitir multiples instancias." ' 84
.lstLanguage.AddItem "    Habilitar menu en drives y directorios." ' 85
.lstLanguage.AddItem "    Mostrar MMPlayerX en:" ' 86
.lstLanguage.AddItem "    Barra de tareas." ' 87
.lstLanguage.AddItem "    Bandeja de sistema." ' 88
.lstLanguage.AddItem "  Transparencia" ' 89
.lstLanguage.AddItem "    Transparencia(Solo win 2000 o sup.)" ' 90
.lstLanguage.AddItem "Skins" ' 91
.lstLanguage.AddItem "  Skin actual:" ' 92
.lstLanguage.AddItem "  Información:" ' 93
.lstLanguage.AddItem "  Cargar región desde archivo." ' 94
.lstLanguage.AddItem "Wallpaper" ' 95
.lstLanguage.AddItem "  Opciones de fondo de escritorio." ' 96
.lstLanguage.AddItem "  No alterar." ' 97
.lstLanguage.AddItem "  Ajustar." ' 98
.lstLanguage.AddItem "  Centrar." ' 99
.lstLanguage.AddItem "  Mosaico." ' 100
.lstLanguage.AddItem "  Proporcional." ' 101
.lstLanguage.AddItem "Play List" ' 102
.lstLanguage.AddItem "  Formato Lista de Reproducción." ' 103
.lstLanguage.AddItem "  Formato de texto reproduciendo." ' 104
.lstLanguage.AddItem "  Tipo de Scroll:" ' 105
.lstLanguage.AddItem "    Rotar." ' 106
.lstLanguage.AddItem "    Zig Zag." ' 107
.lstLanguage.AddItem "  Velocidad del Scroll:" ' 108
.lstLanguage.AddItem "Reproductor" ' 109
.lstLanguage.AddItem "  Reproducir archivos:" ' 110
.lstLanguage.AddItem "  Mostrar icono en bandeja de sistema:" ' 111
.lstLanguage.AddItem "    Anterior Track." ' 112
.lstLanguage.AddItem "    Reproducir." ' 113
.lstLanguage.AddItem "    Pausar." ' 114
.lstLanguage.AddItem "    Detener." ' 115
.lstLanguage.AddItem "    Siguente Track." ' 116
.lstLanguage.AddItem "  Crossfade entre tracks (ms):" ' 117
.lstLanguage.AddItem "  Crossfade en Detener (ms):" ' 118
.lstLanguage.AddItem "  Reproducir al inicio." ' 119
.lstLanguage.AddItem "Efectos FX" ' 120
.lstLanguage.AddItem "  coro" ' 121
.lstLanguage.AddItem "    Habilitar coro." ' 122
.lstLanguage.AddItem "      Mezcla:" ' 123
.lstLanguage.AddItem "      Profundidad:" ' 124
.lstLanguage.AddItem "      Retroacción:" ' 125
.lstLanguage.AddItem "      Frecuencía:" ' 126
.lstLanguage.AddItem "      Forma de onda:" ' 127
.lstLanguage.AddItem "      Retrazo:" ' 128
.lstLanguage.AddItem "      Fase:" ' 129
.lstLanguage.AddItem "  compresor" ' 130
.lstLanguage.AddItem "    Habilitar compresor." ' 131
.lstLanguage.AddItem "      Incremento:" ' 132
.lstLanguage.AddItem "      Ataque:"
.lstLanguage.AddItem "      Edición:"
.lstLanguage.AddItem "      Umbral:"
.lstLanguage.AddItem "      Proporción:"
.lstLanguage.AddItem "      Preretrazo:"
.lstLanguage.AddItem "  Distorción"
.lstLanguage.AddItem "    Habilitar Distorción:"
.lstLanguage.AddItem "      Incremento:" ' 140
.lstLanguage.AddItem "      Bordes:"
.lstLanguage.AddItem "      frecuencia Central:"
.lstLanguage.AddItem "      Ancho frecuencia:"
.lstLanguage.AddItem "      Atenuación:"
.lstLanguage.AddItem "  eco" '145
.lstLanguage.AddItem "    Habilitar eco."
.lstLanguage.AddItem "      Mezcla:"
.lstLanguage.AddItem "      Retroaccion:"
.lstLanguage.AddItem "      Atraso izquierda:"
.lstLanguage.AddItem "      Atraso derecha:"
.lstLanguage.AddItem "      Atraso Central:"
.lstLanguage.AddItem "  Flanger" ' 152
.lstLanguage.AddItem "    Habilitar Flanger."
.lstLanguage.AddItem "      Mezcla:"
.lstLanguage.AddItem "      Profundidad:"
.lstLanguage.AddItem "      Retroaccion:"
.lstLanguage.AddItem "      Frecuencia:"
.lstLanguage.AddItem "      Forma de Onda:"
.lstLanguage.AddItem "      Retrazo:"
.lstLanguage.AddItem "      Fase:" '160
.lstLanguage.AddItem "  Gargarizar" ' 161
.lstLanguage.AddItem "    Habilitar gargarizar."
.lstLanguage.AddItem "      Hz:"
.lstLanguage.AddItem "      Forma de Onda:"
.lstLanguage.AddItem "  I3DL2 Reverberación" ' 165
.lstLanguage.AddItem "    Habilitar I3D nivel 2 Reverberación."
.lstLanguage.AddItem "      Cuarto:"
.lstLanguage.AddItem "      Cuarto HF:"
.lstLanguage.AddItem "      Factor giratorio:" '169
.lstLanguage.AddItem "      Tiempo decadencia:"
.lstLanguage.AddItem "      Prop. dec. HF:"
.lstLanguage.AddItem "      Reflecciones:"
.lstLanguage.AddItem "      Atraso Refleccción:"
.lstLanguage.AddItem "      Reverberación:"
.lstLanguage.AddItem "      Atraso de Rev.:"
.lstLanguage.AddItem "      Difusión:"
.lstLanguage.AddItem "      Densidad:"
.lstLanguage.AddItem "      HF Referencia:"
.lstLanguage.AddItem "  Reverberación" '179
.lstLanguage.AddItem "    Habilitar Reverberación de ondas."
.lstLanguage.AddItem "      Incremento:"
.lstLanguage.AddItem "      Mezcla Reverberación:"
.lstLanguage.AddItem "      Tiempo de Rev.:"
.lstLanguage.AddItem "      HF Proporción:"
.lstLanguage.AddItem "  Valores por default" ' 185
.lstLanguage.AddItem "  Desabilitar todos"
.lstLanguage.AddItem "Equalizador"
.lstLanguage.AddItem "  Habilitar EQ."
.lstLanguage.AddItem "  Presentes:" '190
.lstLanguage.AddItem "  Borrar EQ"
.lstLanguage.AddItem "  Guardar EQ"
.lstLanguage.AddItem "  Nombre del Equalizador:"
.lstLanguage.AddItem "  Borrar equalizador:"
.lstLanguage.AddItem "Visualización" '195
.lstLanguage.AddItem "  Visualizaciones:"
.lstLanguage.AddItem "  Presentes:"
.lstLanguage.AddItem "  Nuevos:"
.lstLanguage.AddItem "  Tipo Fondo:" '198
.lstLanguage.AddItem "  Peaks:" '200
.lstLanguage.AddItem "  Barras:"
.lstLanguage.AddItem "  Archivo Imagen:"
.lstLanguage.AddItem "  Escala:"
.lstLanguage.AddItem "  Color Barras:"
.lstLanguage.AddItem "  Num. Barras:" '205
.lstLanguage.AddItem "  Espacio:"
.lstLanguage.AddItem "  Reflejo:"
.lstLanguage.AddItem "  Color Peak:"
.lstLanguage.AddItem "  Alto Peak:"
.lstLanguage.AddItem "  Gravedad Peak:" '210
.lstLanguage.AddItem "  Gradiente:"
.lstLanguage.AddItem "  Color Fondo:"
.lstLanguage.AddItem "  Color Linea:"
.lstLanguage.AddItem "  Num. Lineas"
.lstLanguage.AddItem "  Alineacion:" '215
.lstLanguage.AddItem "  Guardar"
.lstLanguage.AddItem "  Guardar como"
.lstLanguage.AddItem "  Borrar"
.lstLanguage.AddItem "  Mostrar" '219
.lstLanguage.AddItem "  Borrar Visualización:"
.lstLanguage.AddItem "  Nombre de Visualización:"
.lstLanguage.AddItem "  Anterior Visualizacion"
.lstLanguage.AddItem "  Siguiente Visualizacion"
.lstLanguage.AddItem "  Configurar ..."
.lstLanguage.AddItem "  Salir"
.lstLanguage.AddItem "  Guardar Config."
.lstLanguage.AddItem "  Física:"
.lstLanguage.AddItem "  Virtual:"
.lstLanguage.AddItem "  Archivo:"
.lstLanguage.AddItem " Buscar Archivos de sonido."
.lstLanguage.AddItem " Buscar en:"
.lstLanguage.AddItem " Explorar..."
.lstLanguage.AddItem " Comenzar a Buscar"
.lstLanguage.AddItem " Detener Busqueda"
.lstLanguage.AddItem " Agregar Nuevos Archivos."

End With
End Sub
 

Public Sub Load_Language(strLang As String)
 On Error Resume Next
 Dim Linenr As Integer
 Dim InputData
 Dim strRuta As String, strTemp As String
 With frmPopUp
  
   strRuta = tAppConfig.AppConfig & "Language\" & strLang & ".lng"
   Load_Language_Spanish
   If Dir(strRuta) <> "" Then
    Open strRuta For Input As #2

     Linenr = 0
     Do While Not EOF(2)
       Line Input #2, InputData
        
        If Linenr > 234 Then
          Exit Do
        End If
        If Trim(InputData) <> "" And Linenr > 0 Then
        
          If Linenr > 15 And Linenr < 33 And Linenr <> 28 Then
            strTemp = Left(LineLanguage(Linenr), 1)
            strTemp = Trim(strTemp) & "" & InputData
            .lstLanguage.List(Linenr) = Trim(strTemp)
          Else
           .lstLanguage.List(Linenr) = Trim(InputData)
          End If
        End If
        
        Linenr = Linenr + 1
     Loop
    Close #2
   End If
   ' MENU
   .mnuNuevaBusqueda.Caption = LineLanguage(1)
   .mnuCFront.Caption = LineLanguage(2)
   .mnuCambiarListaCaratula.Caption = LineLanguage(3)
   frmMain.Button(10).ToolTipText = Trim(LineLanguage(3))
   .mnuWallpapper.Caption = LineLanguage(4)
   .mnuMCaratula.Caption = LineLanguage(5)
   .mnuBrowsers.Caption = LineLanguage(6)
   .mnuExplorar.Caption = LineLanguage(7)
   .mnuExpAlbum.Caption = LineLanguage(8)
   .mnuTagEditor.Caption = LineLanguage(9)
   .mnuLyrics.Caption = LineLanguage(10)
   .mnuListSpec.Caption = LineLanguage(11)
   .mnuShowSpec.Caption = LineLanguage(12)
   .mnuMaxSpec.Caption = LineLanguage(13)
   .mnuControles.Caption = LineLanguage(14)
   .mnuVolumen.Caption = LineLanguage(15)
   .mnuSubirVolumen.Caption = LineLanguage(16)
   .mnuBajarVolumen.Caption = LineLanguage(17)
   .mnuTrackAnterior.Caption = LineLanguage(18)
   frmMain.Button(0).ToolTipText = Trim(Right(LineLanguage(18), Len(LineLanguage(18)) - 1))
   frmMain.ButtonMini(0).ToolTipText = Trim(Right(LineLanguage(18), Len(LineLanguage(18)) - 1))
   .mnuReproducir.Caption = LineLanguage(19)
   frmMain.Button(1).ToolTipText = Trim(Right(LineLanguage(19), Len(LineLanguage(19)) - 1))
   frmMain.ButtonMini(1).ToolTipText = Trim(Right(LineLanguage(19), Len(LineLanguage(19)) - 1))
   .mnuPausa.Caption = LineLanguage(20)
   frmMain.Button(2).ToolTipText = Trim(Right(LineLanguage(20), Len(LineLanguage(20)) - 1))
   frmMain.ButtonMini(2).ToolTipText = Trim(Right(LineLanguage(20), Len(LineLanguage(20)) - 1))
   .mnuDetener.Caption = LineLanguage(21)
   frmMain.Button(3).ToolTipText = Trim(Right(LineLanguage(21), Len(LineLanguage(21)) - 1))
   frmMain.ButtonMini(3).ToolTipText = Trim(Right(LineLanguage(21), Len(LineLanguage(21)) - 1))
   .mnuSigTrack.Caption = LineLanguage(22)
   frmMain.Button(4).ToolTipText = Trim(Right(LineLanguage(22), Len(LineLanguage(22)) - 1))
   frmMain.ButtonMini(4).ToolTipText = Trim(Right(LineLanguage(22), Len(LineLanguage(22)) - 1))
   .mnuAnteriorAlbum.Caption = LineLanguage(23)
   frmMain.Button(9).ToolTipText = Trim(Right(LineLanguage(23), Len(LineLanguage(23)) - 1))
   .mnuSigAlbum.Caption = LineLanguage(24)
   frmMain.Button(11).ToolTipText = Trim(Right(LineLanguage(24), Len(LineLanguage(24)) - 1))
   .mnuIntro.Caption = LineLanguage(25)
   frmMain.Button(5).ToolTipText = Trim(Right(LineLanguage(25), Len(LineLanguage(25)) - 1))
   .mnuSilencio.Caption = LineLanguage(27)
   frmMain.Button(6).ToolTipText = Trim(Right(LineLanguage(27), Len(LineLanguage(27)) - 1))
   .mnuRepetir.Caption = LineLanguage(26)
   frmMain.Button(7).ToolTipText = Trim(Right(LineLanguage(26), Len(LineLanguage(26)) - 1))
   .mnuOrdenAleatorio.Caption = LineLanguage(28)
   frmMain.Button(8).ToolTipText = Trim(LineLanguage(28))
   .mnuAleatorioActAlbum.Caption = LineLanguage(29)
   .mnuAleatorioTodaColec.Caption = LineLanguage(30)
   .mnuAtras5Seg.Caption = LineLanguage(31)
   .mnuAdelante5Seg.Caption = LineLanguage(32)
   .mnuOpciones.Caption = LineLanguage(33)
   .mnuSkins.Caption = LineLanguage(34)
   .mnuExpSkins.Caption = LineLanguage(35)
   .mnuWOpacity.Caption = LineLanguage(36)
   .mnuAlphaPer.Caption = LineLanguage(37)
   .mnuAcercaDe.Caption = LineLanguage(38)
   .mnuSalir.Caption = LineLanguage(39)
   
   
   frmMain.Button(12).ToolTipText = LineLanguage(48)
   frmMain.ButtonMini(5).ToolTipText = LineLanguage(48)
   frmMain.Button(13).ToolTipText = LineLanguage(49)
   frmMain.ButtonMini(6).ToolTipText = LineLanguage(49)
   frmMain.Button(14).ToolTipText = LineLanguage(50)
   frmMain.ButtonMini(7).ToolTipText = LineLanguage(50)
   frmMain.Button(15).ToolTipText = LineLanguage(51)
   frmMain.ButtonMini(8).ToolTipText = LineLanguage(51)
   
   .mnuSpecNone.Caption = LineLanguage(53)
   .mnuSpecBars.Caption = LineLanguage(54)
   .mnuSpecOsc.Caption = LineLanguage(55)
   
   .mnuAlbumTags.Caption = LineLanguage(56)
   .mnuAlbumBrowser.Caption = LineLanguage(57)
   .mnuAlbumExp.Caption = LineLanguage(58)
   .mnuAlbumPlay.Caption = LineLanguage(59)
   
   Load_Language_Options
   
   If bolDirectoriosShow = True Then Load_Language_Directorios
   
   If bolCaratulaShow = True Then frmCaratula.Caption = LineLanguage(41)
   If bolAcercaShow = True Then frmAcerca.Caption = LineLanguage(40)
   If bolLyricsShow = True Then frmLyrics.Caption = LineLanguage(46): frmLyrics.lblNoLyrics.Caption = LineLanguage(47)
   If bolTagsShow = True Then Load_Language_Tags
   If bolSearchShow = True Then Load_Language_Search
   
   
   '//change language at systray icons
   If PlayerTrayIcon.Previous = True Then CambiarIcono frmMain.txtSTIcon(0).hwnd, frmMain.ImageList.ListImages(1).ExtractIcon.Handle, frmMain.Button(0).ToolTipText & " - MMPlayerX"
   If PlayerTrayIcon.Play = True Then CambiarIcono frmMain.txtSTIcon(1).hwnd, frmMain.ImageList.ListImages(2).ExtractIcon.Handle, frmMain.Button(1).ToolTipText & " - MMPlayerX"
   If PlayerTrayIcon.Pause = True Then CambiarIcono frmMain.txtSTIcon(2).hwnd, frmMain.ImageList.ListImages(3).ExtractIcon.Handle, frmMain.Button(2).ToolTipText & " - MMPlayerX"
   If PlayerTrayIcon.Stop = True Then CambiarIcono frmMain.txtSTIcon(3).hwnd, frmMain.ImageList.ListImages(4).ExtractIcon.Handle, frmMain.Button(3).ToolTipText & " - MMPlayerX"
   If PlayerTrayIcon.Next = True Then CambiarIcono frmMain.txtSTIcon(4).hwnd, frmMain.ImageList.ListImages(5).ExtractIcon.Handle, frmMain.Button(4).ToolTipText & " - MMPlayerX"
  
 End With
End Sub

Sub Load_Language_Search()
  With frmSearch
      .Caption = LineLanguage(229)
      .Label.Caption = LineLanguage(230)
      .cmdBrowse.Caption = LineLanguage(231)
      .chkAdd.Caption = LineLanguage(234)
      
    If bSearching = False Then
      .cmdSearch.Caption = LineLanguage(232)
    Else
      .cmdSearch.Caption = LineLanguage(233)
    End If
 End With
End Sub


Sub Load_Language_Directorios()
  With frmDirectorios
      .Caption = LineLanguage(42) & " [ " & TotalAlbumS & " Albums ]"
      .mnuExpArchivos.Caption = LineLanguage(43)
      .mnuEditTags.Caption = LineLanguage(44)
      .mnuPlay.Caption = LineLanguage(45)
 End With
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Load_Language_Options()
 On Error Resume Next
 Dim i As Integer
 With frmOpciones
   .Caption = LineLanguage(70)
   
   '// APLICACION
   .TreeOptions.Nodes("Application").Text = LineLanguage(74)
    .TSAppConfig.Tabs(1).Caption = LineLanguage(75)
     .lblApp(3).Caption = LineLanguage(76)
     .cmdAppConfig.Caption = LineLanguage(77)
     .lblApp(6).Caption = LineLanguage(78)
     .lblApp(7).Caption = LineLanguage(79)
    .TSAppConfig.Tabs(2).Caption = LineLanguage(80)
     .lblApp(0).Caption = LineLanguage(81)
     .chkWindowsState(2).Caption = LineLanguage(82)
     .chkWindowsState(3).Caption = LineLanguage(83)
     .chkWindowsState(4).Caption = LineLanguage(84)
     .chkDir.Caption = LineLanguage(85)
     .lblApp(1).Caption = LineLanguage(86)
     .chkWindowsState(0).Caption = LineLanguage(87)
     .chkWindowsState(1).Caption = LineLanguage(88)
    .TSAppConfig.Tabs(3).Caption = LineLanguage(89)
     .lblApp(2).Caption = LineLanguage(90)
     
  '// SKINS
    .TreeOptions.Nodes("Skins").Text = LineLanguage(91)
     .lblSkin(1).Caption = LineLanguage(92)
     .lblSkin(0).Caption = LineLanguage(93)
     .chkUseFile.Caption = LineLanguage(94)
     
  '// WALLPAPER
    .TreeOptions.Nodes("Wallpaper").Text = LineLanguage(95)
     .lblWallpaper.Caption = LineLanguage(96)
     .optWallpaper(0).Caption = LineLanguage(97)
     .optWallpaper(3).Caption = LineLanguage(98)
     .optWallpaper(2).Caption = LineLanguage(99)
     .optWallpaper(1).Caption = LineLanguage(100)
     .chkProporcional.Caption = LineLanguage(101)
     
  '// PLAY LIST FORMAT
    .TreeOptions.Nodes("ScrollText").Text = LineLanguage(102)
     .lblPL(0).Caption = LineLanguage(103)
     .lblPL(1).Caption = LineLanguage(104)
     .lblPL(2).Caption = LineLanguage(105)
     .optScrollType(0).Caption = LineLanguage(106)
     .optScrollType(1).Caption = LineLanguage(107)
     .lblPL(3).Caption = LineLanguage(108)
     
  '// REPRODUCTOR
    .TreeOptions.Nodes("Player").Text = LineLanguage(109)
     .lblPlayer(0).Caption = LineLanguage(110)
     .lblPlayer(1).Caption = LineLanguage(111)
     .chkPIcon(0).Caption = LineLanguage(112)
     .chkPIcon(1).Caption = LineLanguage(113)
     .chkPIcon(2).Caption = LineLanguage(114)
     .chkPIcon(3).Caption = LineLanguage(115)
     .chkPIcon(4).Caption = LineLanguage(116)
     .lblPlayer(2).Caption = LineLanguage(117)
     .lblPlayer(3).Caption = LineLanguage(118)
     .chkPlayStart.Caption = LineLanguage(119)
     
  '// EFECTOS DSP FX
    .TreeOptions.Nodes("Effects").Text = LineLanguage(120)
     .tsDSP.Tabs(1).Caption = LineLanguage(121)
     .chkDSP(0).Caption = LineLanguage(122)
      .lblChorus(0).Caption = LineLanguage(123)
      .lblChorus(1).Caption = LineLanguage(124)
      .lblChorus(2).Caption = LineLanguage(125)
      .lblChorus(3).Caption = LineLanguage(126)
      .lblChorus(4).Caption = LineLanguage(127)
      .lblChorus(5).Caption = LineLanguage(128)
      .lblChorus(6).Caption = LineLanguage(129)
      
     .tsDSP.Tabs(2).Caption = LineLanguage(130)
     .chkDSP(1).Caption = LineLanguage(131)
      .lblComp(0).Caption = LineLanguage(132)
      .lblComp(1).Caption = LineLanguage(133)
      .lblComp(2).Caption = LineLanguage(134)
      .lblComp(3).Caption = LineLanguage(135)
      .lblComp(4).Caption = LineLanguage(136)
      .lblComp(5).Caption = LineLanguage(137)
      
     .tsDSP.Tabs(3).Caption = LineLanguage(138)
     .chkDSP(2).Caption = LineLanguage(139)
      .lblDis(0).Caption = LineLanguage(140)
      .lblDis(1).Caption = LineLanguage(141)
      .lblDis(2).Caption = LineLanguage(142)
      .lblDis(3).Caption = LineLanguage(143)
      .lblDis(4).Caption = LineLanguage(144)
      
     .tsDSP.Tabs(4).Caption = LineLanguage(145)
     .chkDSP(3).Caption = LineLanguage(146)
      .lblEcho(0).Caption = LineLanguage(147)
      .lblEcho(1).Caption = LineLanguage(148)
      .lblEcho(2).Caption = LineLanguage(149)
      .lblEcho(3).Caption = LineLanguage(150)
      .lblEcho(4).Caption = LineLanguage(151)
      
     .tsDSP.Tabs(5).Caption = LineLanguage(152)
     .chkDSP(4).Caption = LineLanguage(153)
      .lblFlan(0).Caption = LineLanguage(154)
      .lblFlan(1).Caption = LineLanguage(155)
      .lblFlan(2).Caption = LineLanguage(156)
      .lblFlan(3).Caption = LineLanguage(157)
      .lblFlan(4).Caption = LineLanguage(158)
      .lblFlan(5).Caption = LineLanguage(159)
      .lblFlan(6).Caption = LineLanguage(160)
      
     .tsDSP.Tabs(6).Caption = LineLanguage(161)
     .chkDSP(5).Caption = LineLanguage(162)
      .lblGarg(0).Caption = LineLanguage(163)
      .lblGarg(1).Caption = LineLanguage(164)
      
     .tsDSP.Tabs(7).Caption = LineLanguage(165)
     .chkDSP(6).Caption = LineLanguage(166)
      .lblL2(0).Caption = LineLanguage(167)
      .lblL2(1).Caption = LineLanguage(168)
      .lblL2(2).Caption = LineLanguage(169)
      .lblL2(3).Caption = LineLanguage(170)
      .lblL2(4).Caption = LineLanguage(171)
      .lblL2(5).Caption = LineLanguage(172)
      .lblL2(6).Caption = LineLanguage(173)
      .lblL2(7).Caption = LineLanguage(174)
      .lblL2(8).Caption = LineLanguage(175)
      .lblL2(9).Caption = LineLanguage(176)
      .lblL2(10).Caption = LineLanguage(177)
      .lblL2(11).Caption = LineLanguage(178)
      
     .tsDSP.Tabs(8).Caption = LineLanguage(179)
     .chkDSP(8).Caption = LineLanguage(180)
      .lblWaves(0).Caption = LineLanguage(181)
      .lblWaves(1).Caption = LineLanguage(182)
      .lblWaves(2).Caption = LineLanguage(183)
      .lblWaves(3).Caption = LineLanguage(184)
     
     .cmdDSPReset.Caption = LineLanguage(185)
     .cmdDSPClear.Caption = LineLanguage(186)
     
  '// EQUALIZADOR
    .TreeOptions.Nodes("Equalizer").Text = LineLanguage(187)
     .chkDSP(7).Caption = LineLanguage(188)
     .lblEQ(10).Caption = LineLanguage(189)
     .cmdDeleteEQ.Caption = LineLanguage(190)
     .cmdSaveEQ.Caption = LineLanguage(191)
     
  '// VISUALISATION
    .TreeOptions.Nodes("Visualization").Text = LineLanguage(194)
     .lblCurrentVis(0) = LineLanguage(195)
     .lblCurrentVis(1) = LineLanguage(196)
     .lblCurrentVis(2) = LineLanguage(197)
     
     For i = 198 To 214
      .lblVis(i - 198).Caption = LineLanguage(i)
     Next i
     
     .cmdVisualizacion(0).Caption = LineLanguage(215)
     .cmdVisualizacion(4).Caption = LineLanguage(216)
     .cmdVisualizacion(1).Caption = LineLanguage(217)
     .cmdVisualizacion(2).Caption = LineLanguage(218)
     
     frmSpectrum.mnuPrevVis.Caption = LineLanguage(221)
     frmSpectrum.mnuNextVis.Caption = LineLanguage(222)
     frmSpectrum.mnuConfigVis.Caption = LineLanguage(223)
     frmSpectrum.mnuExit.Caption = LineLanguage(224)
     
     .lblApp(8).Caption = LineLanguage(226)
     .lblApp(9).Caption = LineLanguage(227)
     .lblApp(10).Caption = LineLanguage(228)
    
   '//buttons
   .cmdOk.Caption = LineLanguage(71)
   .cmdCancel.Caption = LineLanguage(72)
   .cmdApply.Caption = LineLanguage(73)
   .cmdSaveConfig.Caption = LineLanguage(225)
  End With
End Sub

Sub Load_Language_Tags()
 With frmTags
    .Caption = LineLanguage(60)
    .cmdOk.Caption = LineLanguage(67)
    .cmdCancel.Caption = LineLanguage(68)
    .cmdApply.Caption = LineLanguage(69)
    .cmdSelAll.Caption = LineLanguage(62)
    .TabStrip.Tabs(1).Caption = LineLanguage(63)
    .TabStrip.Tabs(2).Caption = LineLanguage(64)
    .cmdAdd.Caption = LineLanguage(65)
    .cmdUndo.Caption = LineLanguage(66)
 End With
End Sub

Public Function LineLanguage(Number As Integer) As String
 On Error Resume Next
  
   If frmPopUp.lstLanguage.ListCount = 0 Then Exit Function
   If Number > frmPopUp.lstLanguage.ListCount - 1 Then Exit Function
   LineLanguage = Trim(frmPopUp.lstLanguage.List(Number))
  
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

