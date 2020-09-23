Attribute VB_Name = "mConfig"
Option Explicit
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| INFORMACION DE PROCESADOR MEMORIA                                                     |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type


Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| OPERATING SYSTEM VERSIN INFORMATION                                                   |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Type OSVersionInfo
    OSVSize       As Long
    dwVerMajor    As Long
    dwVerMinor    As Long
    dwBuildNumber As Long
    PlatformID    As Long
    szCSDVersion  As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVersionInfo) As Long


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| EXECUTAR APLICACIONES CON LOS PARAMETROS DADOS                                        |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| ARRASTRE DEL FORMULARIO                                                               |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| APIS PARA PONER SIEMPRE ARRIBA EL FORMULARIO                                          |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOOWNERZORDER = &H200      '  No usar el orden Z del propietario
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| MOVER EL TEXTO POR LOS PICTURES
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Const DT_BOTTOM As Long = &H8
Public Const DT_CALCRECT As Long = &H400
Public Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Public Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Public Const DT_WORDBREAK As Long = &H10

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|    Declaraciones para Layered Windows (sólo Windows 2000 y superior)                  |
'|    APIS PARA PONER TRASPARENTE EL FORM                                                |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


Public Const WS_EX_LAYERED As Long = &H80000
Public Const LWA_ALPHA As Long = &H2
Public Const GWL_EXSTYLE = (-20)
Public Const RDW_INVALIDATE = &H1
Public Const RDW_ERASE = &H4
Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_FRAME = &H400

'
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function RedrawWindow2 Lib "user32" Alias "RedrawWindow" (ByVal hwnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  APIS PARA LEER LAS CONFIGURACIONES DE LOS ARCHIVOS .INI O DEMAS
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
    As String, lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpRetunedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
    As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lplFileName As String) As Long
    
 Public Enum Sel_Option
  PathExe = 0
  PathSkin = 1
End Enum

Public Function Read_INI(Section As String, Value As String, Default As Variant, Optional IsColor As Boolean = False, Optional ConfigurationMusic As Boolean = False, Optional FilePath As String) As Variant
 '// Funcion para leer configuraciones del INI
 '// Parametros
 '// [Section] -> Rama principal del .ini : ei:  [Configuration]
 '// [Value] -> Valor de la Seccion , ej: Intro = False
 '// [Default] -> Valor de retorno si no se encuantra el valor
 '// Valor de retorno el valor de la seccion si se encuantra
 
 Dim ColorArr As Variant
 Dim Str As String
    
  If ConfigurationMusic = True Then
    Str = String(255, Chr(0))
    Str = Left(Str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", Str, Len(Str), tAppConfig.AppPath & App.EXEName & ".ini"))
    If Str = "NO_TA" Then ' si no encuentra la clave
       Read_INI = Trim(Default)
    Else
       Read_INI = Trim(Str)
    End If
    Exit Function
  End If
      
  If Trim(FilePath) <> "" Then
    Str = String(255, Chr(0))
    Str = Left(Str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", Str, Len(Str), FilePath))
    If Str = "NO_TA" Then ' si no encuentra la clave
       Read_INI = Trim(Default)
    Else
       Read_INI = Trim(Str)
    End If
    Exit Function
  End If
  
  If IsColor = True Then ' is a color
    Str = String(255, Chr(0))
    Str = Left(Str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", Str, Len(Str), tAppConfig.AppConfig & "skins\" & tAppConfig.Skin & "\" & "Skin.ini"))
    
    If Str = "NO_TA" Then ' si no encuentra la clave
       Read_INI = Default
    Else
      ColorArr = Split(Str, ",")
       If UBound(ColorArr) <> 2 Then ' si esta mal la che clave
         Read_INI = Default
       Else
         Read_INI = RGB(ColorArr(0), ColorArr(1), ColorArr(2))
       End If
    End If
  Else
    Str = String(255, Chr(0))
    Str = Left(Str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", Str, Len(Str), tAppConfig.AppConfig & "skins\" & tAppConfig.Skin & "\" & "Skin.ini"))
    If Str = "NO_TA" Then ' si no encuentra la clave
       Read_INI = Trim(Default)
    Else
       Read_INI = Trim(Str)
    End If
  End If
End Function

Public Function Read_Config_Button(Objeto As Object, Section As String, Value As String, Default As Variant) As Boolean
  On Error Resume Next
  
  Dim Str As String
  Dim arry() As String
    Str = String(255, Chr(0))
    Str = Left(Str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", _
               Str, Len(Str), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin _
               & "\Skin.ini"))
    
    If Str = "NO_TA" Then Str = Default
    
    arry = Split(Str, ",")
      
      '// Slider Pos or Vol
      If UBound(arry) = 4 Then
        If UCase(arry(4)) = "V" Then '// Vertikal
           Objeto.Position = 0
        Else '// Horizontal
           Objeto.Position = 1
        End If
      End If
      
      '// Button normal
      If UBound(arry) = 4 Or UBound(arry) = 3 Then
         'Read_INI = Str
         Objeto.Left = arry(0)
         Objeto.Top = arry(1)
         Objeto.Width = arry(2)
         Objeto.Height = arry(3)
      End If
      
       
End Function


Public Function Write_INI(Section As String, KeyName As String, KeyValue As Variant, FilePath As String) As Boolean
Dim Ret As Long
    Ret = WritePrivateProfileString(Section, KeyName, CStr(KeyValue), FilePath)
    If Ret = 0 Then
        Write_INI = True
    Else
        Write_INI = False
    End If
End Function

Sub Load_Settings_INI(bolNormal As Boolean)
 Dim strRes As Variant, strAlbums As String
 Dim arryFormat() As String
 Dim i As Integer, intMp3 As Integer
 Dim strKeyQuery As Variant
 Dim lngRootKey As Long

 On Error Resume Next
  strKeyQuery = vbNullString
  lngRootKey = HKEY_CURRENT_USER
  
  tAppConfig.AppPath = App.Path
  If Right(tAppConfig.AppPath, 1) <> "\" Then tAppConfig.AppPath = tAppConfig.AppPath & "\"
  
  tAppConfig.AppConfig = Read_INI("Configuration", "AppConfiguration", tAppConfig.AppPath & "MMp3Player\", , True)
  
  If Right(tAppConfig.AppConfig, 1) <> "\" Then tAppConfig.AppConfig = tAppConfig.AppConfig & "\"
  If Dir(tAppConfig.AppConfig, vbDirectory) = "" Then tAppConfig.AppConfig = tAppConfig.AppPath & "MMp3Player\"
  
  '// Multiples instancias
  strRes = Read_INI("Configuration", "MulInstances", 0, , True)
  
  If CBool(strRes) = False Then       '// Si este en falso y hay otra
    If App.PrevInstance = True Then     '// Instancia terminar
        Beep
        'AppActivate "MMPlayerX"
      End
      Exit Sub
    End If
  End If
  If CBool(strRes) = True Then OpcionesMusic.Instancias = True
    
  '// Mostrar Splash Screen
  strRes = Read_INI("Configuration", "SplashScreen", 1, , True)
  If CBool(strRes) = True Then
     frmSplash.lblSplash(0).Caption = " Loading configuration..."
     frmSplash.Show
     OpcionesMusic.Splash = True
  End If
  
  '// cargar regiones desde archivo
  strRes = Read_INI("Configuration", "LoadRegionFile", 0, , True)
  If CBool(strRes) = True Then bLoadRegionFile = True

  '// Kargar Skin
  strRes = Read_INI("Configuration", "Skin", "", , True)
  
  If Trim(strRes) = "" Or Dir(tAppConfig.AppConfig & "Skins\" & strRes, vbDirectory) = "" Then
    Load_Skins_Menu LCase(tAppConfig.Skin)
    Change_Skin Trim(frmPopUp.mnuSkinsAdd(1).Caption)
    frmPopUp.mnuSkinsAdd(1).Checked = True
    Form_Mini_Normal
  Else
    Change_Skin Trim(strRes) '// cambiar skin, posicion de los controles
    Form_Mini_Normal '// si tiene zonas irregulares ajustar el form
    Load_Skins_Menu LCase(tAppConfig.Skin) '// kargar el menu de Skins y seleccionar el actual
  End If
  
  '// Estado de la maskara mini - normal
  strRes = Read_INI("Configuration", "Mini", 0, , True)
  If CBool(strRes) = True Then bMiniMask = True
  
  '// Mover los formularios
  strRes = Read_INI("Configuration", "MX", 0, , True)
  If IsNumeric(strRes) = False Then strRes = 0
     frmMain.Left = CInt(strRes)

   
  strRes = Read_INI("Configuration", "MY", 0, , True)
  If IsNumeric(strRes) = False Then strRes = 0
     frmMain.Top = CInt(strRes)
    
  '// si no esta seleccionado el splash screen mostrar los form ahora
  If bolSplashScreen = False Then
     If bMiniMask = True Then
       Change_Mask True, False
     Else
       Change_Mask False, False
     End If
     frmMain.Show
  End If
   
  '-----------------------------------------------------------------------
  'Guardar la ruta del Wallpaper al inicio que se tiene
  strKeyQuery = regQuery_A_Key(lngRootKey, "Control panel\Desktop", "Wallpaper")
  OriginalRutaWallpaper = strKeyQuery
  
  '-----------------------------------------------------------------------
  'Guardar el Estilo Wallpaper al inicio
  strKeyQuery = regQuery_A_Key(lngRootKey, "Control panel\Desktop", "WallpaperStyle")
  OriginalWallpaperStyle = strKeyQuery
  
  '-----------------------------------------------------------------------
  'Guardar el tileWallpaper al inicio
  strKeyQuery = regQuery_A_Key(lngRootKey, "Control panel\Desktop", "TileWallpaper")
  OriginalTileWallpaper = strKeyQuery
  
  
  '// Guardar los estilos de Walppaper al inicio
  strRes = Read_INI("Configuration", "Wallpaper", 0, , True)
  
  If CInt(strRes) < 0 Or CInt(strRes) > 3 Or IsNumeric(strRes) = False Then strRes = 0
  
  '//Poner valores correctos por si modifican el archivo
  If strRes = 0 Then
    OpcionesMusic.NoAlteraR = True
  ElseIf strRes = 1 Then
        OpcionesMusic.Mosaico = True
      ElseIf strRes = 2 Then
            OpcionesMusic.Centrar = True
          Else
            OpcionesMusic.Expander = True
          End If
  
  '// Visualizacion
  strRes = Read_INI("Configuration", "Visualization", 1, , True)
  
  If CInt(strRes) < 0 Or CInt(strRes) > 2 Or IsNumeric(strRes) = False Then strRes = 1
  
  If strRes = 0 Then
      frmPopUp.mnuSpecNone.Checked = True
    ElseIf strRes = 1 Then
          frmPopUp.mnuSpecBars.Checked = True
        ElseIf strRes = 2 Then
              frmPopUp.mnuSpecOsc.Checked = True
            End If
            
  '// format scroll
  sFormatPlayList = Trim(Read_INI("Configuration", "FormatPlayList", "%A - %S", , True))
    
  '// format scroll
  sFormatScroll = Trim(Read_INI("Configuration", "FormatScroll", "%S - %A (%T)", , True))
                   
  '// scroll caption type
  strRes = Read_INI("Configuration", "ScrollType", 0, , True)
  If CInt(strRes) < 0 Or CInt(strRes) > 1 Or IsNumeric(strRes) = False Then strRes = 0
   iScrollType = CInt(strRes)
   frmMain.ScrollText(1).ScrollType = iScrollType
   frmMain.ScrollText(5).ScrollType = iScrollType
                   
  '// scroll caption vel
  strRes = Read_INI("Configuration", "ScrollVel", 130, , True)
  If CInt(strRes) < 100 Or CInt(strRes) > 1000 Or IsNumeric(strRes) = False Then strRes = 130
   iScrollVel = CInt(strRes)
   frmMain.ScrollText(1).ScrollVelocity = iScrollVel
   frmMain.ScrollText(5).ScrollVelocity = iScrollVel
                 
  '// Crossfade entre Tracks
  strRes = Read_INI("Configuration", "CrossfadeTrack", 100, , True)
  If CInt(strRes) < 0 Or CInt(strRes) > 400 Or IsNumeric(strRes) = False Then strRes = 100
   iCrossfadeTrack = CInt(strRes)
   
  '// Crossfade para detener
  strRes = Read_INI("Configuration", "CrossfadeStop", 100, , True)
  If CInt(strRes) < 0 Or CInt(strRes) > 400 Or IsNumeric(strRes) = False Then strRes = 100
   iCrossfadeStop = CInt(strRes)
                    
   '// add files at search
  strRes = Read_INI("Configuration", "AddFiles", 0, , True)
  If CBool(strRes) = True Then bAddFiles = True
  
  '// play en el comienzo
  strRes = Read_INI("Configuration", "PlayStarting", 1, , True)
  If CBool(strRes) = True Then bPlayStarting = True
                 
  '// check proporcional
  strRes = Read_INI("Configuration", "Proportional", 0, , True)
  If CBool(strRes) = True Then OpcionesMusic.Proporcional = True
  
  '// check Directorio
  strRes = Read_INI("Configuration", "Directory", 0, , True)
  If CBool(strRes) = True Then OpcionesMusic.Directorio = True
 
  '// check show task bar
  strRes = Read_INI("Configuration", "TaskBar", 1, , True)
  If CBool(strRes) = True Then OpcionesMusic.TaskBar = True
   
   '// system tray icon
  strRes = Read_INI("Configuration", "SysTray", 0, , True)
  If CBool(strRes) = True Then
    OpcionesMusic.SysTray = True
    ColocarIcono frmMain.Text1.hwnd, frmMain.Icon.Handle, "MMPlayerX v 2.0 - by Raul Martinez"
  End If
  
  strRes = Read_INI("Configuration", "SysTrayNext", 0, , True)
  If CBool(strRes) = True Then
     PlayerTrayIcon.Next = True
     ColocarIcono frmMain.txtSTIcon(4).hwnd, frmMain.ImageList.ListImages(5).ExtractIcon.Handle, "Next Track - MMPlayerX"
  End If
  
  strRes = Read_INI("Configuration", "SysTrayStop", 0, , True)
  If CBool(strRes) = True Then
     PlayerTrayIcon.Stop = True
     ColocarIcono frmMain.txtSTIcon(3).hwnd, frmMain.ImageList.ListImages(4).ExtractIcon.Handle, "Stop - MMPlayerX"
  End If
  
  strRes = Read_INI("Configuration", "SysTrayPause", 0, , True)
  If CBool(strRes) = True Then
     PlayerTrayIcon.Pause = True
     ColocarIcono frmMain.txtSTIcon(2).hwnd, frmMain.ImageList.ListImages(3).ExtractIcon.Handle, "Pause  - MMPlayerX"
  End If
  
  strRes = Read_INI("Configuration", "SysTrayPlay", 0, , True)
  If CBool(strRes) = True Then
     PlayerTrayIcon.Play = True
     ColocarIcono frmMain.txtSTIcon(1).hwnd, frmMain.ImageList.ListImages(2).ExtractIcon.Handle, "Play - MMPlayerX"
  End If
    
    strRes = Read_INI("Configuration", "SysTrayPrevious", 0, , True)
  If CBool(strRes) = True Then
     PlayerTrayIcon.Previous = True
     ColocarIcono frmMain.txtSTIcon(0).hwnd, frmMain.ImageList.ListImages(1).ExtractIcon.Handle, "Previous - MMPlayerX"
  End If
  
   '//----------------------------------------------------------------------------------
 '// play files format
  strRes = Read_INI("Configuration", "FileType", "1;0;0;0", , True)
    
   Dim arryFiles(3) As String
      arryFiles(0) = "mp3"
      arryFiles(1) = "wma"
      arryFiles(2) = "wav"
      arryFiles(3) = "ogg"
      
     
  arryFormat = Split(strRes, ";", , vbTextCompare)
  
  For i = 0 To UBound(arryFormat)
     If CBool(arryFormat(i)) = True Then
      If i <= UBound(arryFiles) Then strPathern = strPathern & "*." & arryFiles(i) & ";"
     End If
    
  Next i
  sFileType = strRes
  If strPathern = "" Then strPathern = "*.mp3"
  If Right(strPathern, 1) = ";" Then strPathern = Left(strPathern, Len(strPathern) - 1)
  
  frmMain.ListRepRef.Pattern = strPathern
  frmMain.FileSearch.Pattern = strPathern
  frmMain.FileAleatorio.Pattern = strPathern
 '//----------------------------------------------------------------------------------

 
  '// Trasparencia del form
  strRes = Read_INI("Configuration", "Alpha", 100, , True)
  If strRes < 10 Or strRes > 100 Then strRes = 100
  OpcionesMusic.Alpha = strRes
  Make_Transparent frmMain.hwnd, OpcionesMusic.Alpha '// Poner Trasparente
      
      For i = 0 To 9
       If Left(frmPopUp.mnuAlpha(i).Caption, Len(frmPopUp.mnuAlpha(i).Caption) - 1) = OpcionesMusic.Alpha Then
         frmPopUp.mnuAlpha(i).Checked = True
            frmPopUp.mnuAlphaPer.Caption = Trim(LineLanguage(37))
            frmPopUp.mnuAlphaPer.Checked = False
         Exit For
       Else
         frmPopUp.mnuAlphaPer.Caption = Trim(LineLanguage(37)) & " [ " & OpcionesMusic.Alpha & "% ]"
         frmPopUp.mnuAlphaPer.Checked = True
       End If
     Next i
  
  '// Olways on top
  strRes = Read_INI("Configuration", "AlwaysTop", 0, , True)
  If CBool(strRes) = True Then OpcionesMusic.SiempreTop = True
    
  '// Ajustar Volumen
  strRes = Read_INI("Configuration", "Volume", 255, , True)
  If strRes < 0 Or strRes > 255 Then strRes = 255
   frmMain.Slider(1).Value = strRes
   frmMain.VolumeNActuaL = strRes
    
  '// load lenguaje y cambiarlo
  strRes = Read_INI("Configuration", "Language", "English", , True)
  OpcionesMusic.Language = strRes
  Load_Language OpcionesMusic.Language

'// -------------------------------------------------------------------------------
If bolNormal = True Then '// si es cargado normalmente
  
'---------------------------------------------------------------------------------------
'Hacer mientras se lea algo en el archivo .ini
 frmPopUp.fileBmps.Pattern = strPathern
 Do While strAlbums <> "\"
   i = i + 1
   '// Kargar los albums guardados
   strAlbums = Read_INI("albums", "Album_" & i, "\", , True)
   If strAlbums <> "\" Then '// Si existe el Album
     If Dir(strAlbums & "\") <> "" Then  '// Si existe el directorio
       frmPopUp.fileBmps.Path = strAlbums
       If frmPopUp.fileBmps.ListCount > 0 Then '// Si hay mp3's
         CopyMp3Totales = CopyMp3Totales + frmPopUp.fileBmps.ListCount
         intMp3 = intMp3 + 1
         '// Kargar el arreglos de btnAlbums para almacenar los albums paths
         If intMp3 = 1 Then
           frmMain.btnAlbum(intMp3).ToolTipText = strAlbums
         Else
           Load frmMain.btnAlbum(intMp3)
           frmMain.btnAlbum(intMp3).ToolTipText = strAlbums
         End If
       End If
     End If
   End If
   If strAlbums = "@-" Then Exit Do
 Loop
 
 
 CopyTotalAlbums = intMp3
If intMp3 > 0 Then frmMain.Process_Albums False
 
'// Album anterior reproducido
 strRes = Read_INI("Configuration", "AlbumPlaying", 1, , True)
 If CInt(strRes) > 0 And CInt(strRes) <= (intMp3 - 2) Then
   frmMain.Play_Album CInt(strRes)
 ElseIf intMp3 > 0 Then
        frmMain.Play_Album 1
     End If
     

 '// Numero de track anterior playing
  strRes = Read_INI("Configuration", "TrackNumber", 0, , True)
 If CInt(strRes) >= 0 Then
   frmMain.ListRep.ListIndex = -1
   frmMain.ListRep.ListIndex = CInt(strRes)
   frmMain.ListRep.Selected(CInt(strRes)) = True
 End If
    
 strRes = Read_INI("Configuration", "Intro", 0, , True)
  If CBool(strRes) = True Then frmMain.Intro
 
 strRes = Read_INI("Configuration", "Mute", 0, , True)
  If CBool(strRes) = True Then frmMain.Player_Mute
 
 strRes = Read_INI("Configuration", "Repeat", 0, , True)
  If CBool(strRes) = True Then frmMain.Player_Repeat
    
    
 '// Oreden Aleratorio en el album
 strRes = Read_INI("Configuration", "RandomizeAlbum", 0, , True)
  If CBool(strRes) = True And intMp3 > 0 Then
    frmMain.Randomize_Click True, False
  Else
     strRes = Read_INI("Configuration", "RandomizeCollection", 0, , True)
       If CBool(strRes) = True And intMp3 > 1 Then
         frmMain.Randomize_Click False, False
       End If
  End If
End If

'===============================================================================
' EQUALIZER
Load frmOpciones
With frmOpciones
 strRes = Read_INI("Equalizer", "Enabled", 0, , True)
   If CBool(strRes) = True Then .chkDSP(7).Value = 1
   
 strRes = Read_INI("Equalizer", "Present", -1, , True)
   If strRes >= 0 Or strRes <= .cboEQ.ListCount - 1 Then .cboEQ.ListIndex = CInt(strRes)
   
 For i = 0 To 9
    strRes = Read_INI("Equalizer", "EQ_" & i, 0, , True)
    .sldEQ(i).Value = CInt(strRes)
 Next i
 
 '===============================================================================
 ' SOUND EFFECTS
 strRes = Read_INI("Sound_Effect", "Chorus", 0, , True)
 If CBool(strRes) = True Then .chkDSP(0).Value = 1
      
 For i = 0 To .sldChorus.count - 1
     strRes = Read_INI("Sound_Effect", "Chorus_" & i, 0, , True)
    .sldChorus(i).Value = CInt(strRes)
 Next i
      
 strRes = Read_INI("Sound_Effect", "Compressor", 0, , True)
 If CBool(strRes) = True Then .chkDSP(1).Value = 1
 
 For i = 0 To .sldComp.count - 1
    strRes = Read_INI("Sound_Effect", "Compressor_" & i, 0, , True)
    .sldComp(i).Value = CInt(strRes)
 Next i
      
 strRes = Read_INI("Sound_Effect", "Distortion", 0, , True)
 If CBool(strRes) = True Then .chkDSP(2).Value = 1
   
 For i = 0 To .sldDis.count - 1
    strRes = Read_INI("Sound_Effect", "Distortion_" & i, 0, , True)
    .sldDis(i).Value = CInt(strRes)
 Next i
      
 strRes = Read_INI("Sound_Effect", "Echo", 0, , True)
 If CBool(strRes) = True Then .chkDSP(3).Value = 1
 
 For i = 0 To .sldEcho.count - 1
     strRes = Read_INI("Sound_Effect", "Echo_" & i, 0, , True)
    .sldEcho(i).Value = CInt(strRes)
 Next i
      
 strRes = Read_INI("Sound_Effect", "Flanger", 0, , True)
 If CBool(strRes) = True Then .chkDSP(4).Value = 1
 
 For i = 0 To .sldFlan.count - 1
     strRes = Read_INI("Sound_Effect", "Flanger_" & i, 0, , True)
    .sldFlan(i).Value = CInt(strRes)
 Next i
      
 strRes = Read_INI("Sound_Effect", "Gargle", 0, , True)
 If CBool(strRes) = True Then .chkDSP(5).Value = 1
 
 For i = 0 To .sldGarg.count - 1
     strRes = Read_INI("Sound_Effect", "Gargle_" & i, 0, , True)
    .sldGarg(i).Value = CInt(strRes)
 Next i
      
 strRes = Read_INI("Sound_Effect", "L2Reverb", 0, , True)
 If CBool(strRes) = True Then .chkDSP(6).Value = 1
  
 For i = 0 To .sldL2.count - 1
     strRes = Read_INI("Sound_Effect", "L2Reverb_" & i, 0, , True)
    .sldL2(i).Value = CInt(strRes)
 Next i
      
 strRes = Read_INI("Sound_Effect", "WReverb", 0, , True)
 If CBool(strRes) = True Then .chkDSP(8).Value = 1
      
 For i = 0 To .sldWaves.count - 1
     strRes = Read_INI("Sound_Effect", "WReverb_" & i, 0, , True)
    .sldWaves(i).Value = CInt(strRes)
 Next i
      
 '================================================================================
 '// Visualizacion
 strRes = Read_INI("Configuration", "IndexVis", "", , True)
 
 If strRes <> "" And strRes >= 0 And strRes < frmOpciones.cboVisualizacion.ListCount Then
     Load frmSpectrum
     frmOpciones.cboVisualizacion.ListIndex = CInt(strRes)
     frmOpciones.cmdVisualizacion(2).Value = True
'     frmSpectrum.Stop_Visualizacion
     bolVisShow = False
 Else
      With tConfigVis
       .BackColor = RGB(0, 0, 0)
       .BackColorBar = RGB(255, 255, 255)
       .BackColorPeak = RGB(255, 255, 255)
       .Bars = 30
       ReDim .arryPeaks(.Bars)
       ReDim .arryWaitPeak(.Bars)
       .DrawBars = True
       .DrawPeaks = True
       .DrawSource = 1
       .Exist = True
       .Gradient = "No hay.jpg"
       .GrandientIndex = 0
       .ImageFile = "[Cover Front]"
       .Mirrored = True
       .PeakGravity = 2
       .PeakHeight = 1
       .ScaleUp = 0
       .Spacio = 0
      End With
      Load frmSpectrum
      frmSpectrum.Setup_Visualizacion
'      frmSpectrum.Stop_Visualizacion
      bolVisShow = False
End If
 Unload frmOpciones
 'bolOpcionesShow = False
End With

End Sub

Sub Save_Settings_INI(Optional Normal As Boolean = False)
 Dim Fnum As Integer, i As Integer
 Dim ArchivoINI As String
 Dim intClave As Integer
 Dim INICheck As String
 On Error Resume Next
 
  '// delete systray icons
 If Normal = True Then
  If OpcionesMusic.SysTray = True Then QuitarIcono frmMain.Text1.hwnd
  If PlayerTrayIcon.Previous = True Then QuitarIcono frmMain.txtSTIcon(0).hwnd
  If PlayerTrayIcon.Play = True Then QuitarIcono frmMain.txtSTIcon(1).hwnd
  If PlayerTrayIcon.Pause = True Then QuitarIcono frmMain.txtSTIcon(2).hwnd
  If PlayerTrayIcon.Stop = True Then QuitarIcono frmMain.txtSTIcon(3).hwnd
  If PlayerTrayIcon.Next = True Then QuitarIcono frmMain.txtSTIcon(4).hwnd
 End If
   
 On Error GoTo BITCH
 
ArchivoINI = tAppConfig.AppPath & App.EXEName & ".ini"

'// Chekar los atributos
INICheck = Dir(ArchivoINI, vbNormal + vbSystem + vbHidden + vbReadOnly + vbArchive)

'// Si no se encuentra hacerlo...
If INICheck = "" Then
    Fnum = FreeFile  '// numeroaleatorio para asignar al archivo
    Open ArchivoINI For Output As Fnum
    Close
    'SetAttr ArchivoINI, vbHidden + vbSystem
End If
    
   Write_INI "Configuration", "AppConfiguration", tAppConfig.AppConfig, ArchivoINI
   Write_INI "Configuration", "SplashScreen", OpcionesMusic.Splash, ArchivoINI
   Write_INI "Configuration", "MulInstances", OpcionesMusic.Instancias, ArchivoINI
   Write_INI "Configuration", "Skin", tAppConfig.Skin, ArchivoINI
   Write_INI "Configuration", "LoadRegionFile", bLoadRegionFile, ArchivoINI
   Write_INI "Configuration", "MX", frmMain.Left, ArchivoINI
   Write_INI "Configuration", "MY", frmMain.Top, ArchivoINI
   Write_INI "Configuration", "Volume", frmMain.Slider(1).Value, ArchivoINI
   Write_INI "Configuration", "Mini", bMiniMask, ArchivoINI
   
   If OpcionesMusic.NoAlteraR = True Then
     intClave = 0
   ElseIf OpcionesMusic.Mosaico = True Then
         intClave = 1
       ElseIf OpcionesMusic.Centrar = True Then
             intClave = 2
           Else
             intClave = 3
           End If
    
    Write_INI "Configuration", "Wallpaper", intClave, ArchivoINI
    
    If frmPopUp.mnuSpecNone.Checked = True Then
      intClave = 0
    ElseIf frmPopUp.mnuSpecBars.Checked = True Then
          intClave = 1
        ElseIf frmPopUp.mnuSpecOsc.Checked = True Then
              intClave = 2
            End If
    Write_INI "Configuration", "Visualization", intClave, ArchivoINI
    Write_INI "Configuration", "Proportional", OpcionesMusic.Proporcional, ArchivoINI
    Write_INI "Configuration", "Directory", OpcionesMusic.Directorio, ArchivoINI
    Write_INI "Configuration", "Language", OpcionesMusic.Language, ArchivoINI
    Write_INI "Configuration", "FileType", sFileType, ArchivoINI
    Write_INI "Configuration", "FormatPlayList", sFormatPlayList, ArchivoINI
    Write_INI "Configuration", "FormatScroll", sFormatScroll, ArchivoINI
    Write_INI "Configuration", "ScrollType", iScrollType, ArchivoINI
    Write_INI "Configuration", "ScrollVel", iScrollVel, ArchivoINI
    Write_INI "Configuration", "CrossfadeTrack", iCrossfadeTrack, ArchivoINI
    Write_INI "Configuration", "CrossfadeStop", iCrossfadeStop, ArchivoINI
    Write_INI "Configuration", "PlayStarting", bPlayStarting, ArchivoINI
    Write_INI "Configuration", "Alpha", OpcionesMusic.Alpha, ArchivoINI
    Write_INI "Configuration", "AlwaysTop", OpcionesMusic.SiempreTop, ArchivoINI
    Write_INI "Configuration", "TaskBar", OpcionesMusic.TaskBar, ArchivoINI
    Write_INI "Configuration", "SysTray", OpcionesMusic.SysTray, ArchivoINI
    Write_INI "Configuration", "SysTrayPrevious", PlayerTrayIcon.Previous, ArchivoINI
    Write_INI "Configuration", "SysTrayPlay", PlayerTrayIcon.Play, ArchivoINI
    Write_INI "Configuration", "SysTrayPause", PlayerTrayIcon.Pause, ArchivoINI
    Write_INI "Configuration", "SysTrayStop", PlayerTrayIcon.Stop, ArchivoINI
    Write_INI "Configuration", "SysTrayNext", PlayerTrayIcon.Next, ArchivoINI
    Write_INI "Configuration", "Intro", frmPopUp.mnuIntro.Checked, ArchivoINI
    Write_INI "Configuration", "Mute", frmPopUp.mnuSilencio.Checked, ArchivoINI
    Write_INI "Configuration", "Repeat", frmPopUp.mnuRepetir.Checked, ArchivoINI
    Write_INI "Configuration", "RandomizeCollection", frmPopUp.mnuAleatorioTodaColec.Checked, ArchivoINI
    Write_INI "Configuration", "RandomizeAlbum", frmPopUp.mnuAleatorioActAlbum.Checked, ArchivoINI
    Write_INI "Configuration", "AlbumPlaying", intActiveAlbum, ArchivoINI
    Write_INI "Configuration", "TrackNumber", frmMain.ListRep.ListIndex, ArchivoINI
    Write_INI "Configuration", "IndexVis", IndexVisualization, ArchivoINI
    Write_INI "Configuration", "AddFiles", bAddFiles, ArchivoINI
       
    '===============================================================================
    ' EQUALIZER
    With frmOpciones
     Write_INI "Equalizer", "Enabled", CBool(.chkDSP(7).Value), ArchivoINI
     Write_INI "Equalizer", "Present", .cboEQ.ListIndex, ArchivoINI
      For i = 0 To 9
       Write_INI "Equalizer", "EQ_" & i, .sldEQ(i).Value, ArchivoINI
      Next i
    End With
    
    
    '===============================================================================
    ' SOUND EFFECTS
    With frmOpciones
      Write_INI "Sound_Effect", "Chorus", CBool(.chkDSP(0).Value), ArchivoINI
      For i = 0 To .sldChorus.count - 1
         Write_INI "Sound_Effect", "Chorus_" & i, .sldChorus(i).Value, ArchivoINI
      Next i
      
      Write_INI "Sound_Effect", "Compressor", CBool(.chkDSP(1).Value), ArchivoINI
      For i = 0 To .sldComp.count - 1
         Write_INI "Sound_Effect", "Compressor_" & i, .sldComp(i).Value, ArchivoINI
      Next i
      
      Write_INI "Sound_Effect", "Distortion", CBool(.chkDSP(2).Value), ArchivoINI
      For i = 0 To .sldDis.count - 1
         Write_INI "Sound_Effect", "Distortion_" & i, .sldDis(i).Value, ArchivoINI
      Next i
      
      Write_INI "Sound_Effect", "Echo", CBool(.chkDSP(3).Value), ArchivoINI
      For i = 0 To .sldEcho.count - 1
         Write_INI "Sound_Effect", "Echo_" & i, .sldEcho(i).Value, ArchivoINI
      Next i
      
      Write_INI "Sound_Effect", "Flanger", CBool(.chkDSP(4).Value), ArchivoINI
      For i = 0 To .sldFlan.count - 1
         Write_INI "Sound_Effect", "Flanger_" & i, .sldFlan(i).Value, ArchivoINI
      Next i
      
      Write_INI "Sound_Effect", "Gargle", CBool(.chkDSP(5).Value), ArchivoINI
      For i = 0 To .sldGarg.count - 1
         Write_INI "Sound_Effect", "Gargle_" & i, .sldGarg(i).Value, ArchivoINI
      Next i
      
      Write_INI "Sound_Effect", "L2Reverb", CBool(.chkDSP(6).Value), ArchivoINI
      For i = 0 To .sldL2.count - 1
         Write_INI "Sound_Effect", "L2Reverb_" & i, .sldL2(i).Value, ArchivoINI
      Next i
      
      Write_INI "Sound_Effect", "WReverb", CBool(.chkDSP(8).Value), ArchivoINI
      For i = 0 To .sldWaves.count - 1
         Write_INI "Sound_Effect", "WReverb_" & i, .sldWaves(i).Value, ArchivoINI
      Next i
    End With
    
    '===============================================================================
    ' ALBUMS
    '// Seccion para almecenar los albums actuales reproduciendo
    If TotalAlbumS = 0 Then Exit Sub
     For i = 1 To TotalAlbumS
       Write_INI "Albums", "Album_" & i, frmMain.btnAlbum(i).ToolTipText, ArchivoINI
     Next i
     
     Write_INI "Albums", "Album_" & TotalAlbumS + 1, "@-", ArchivoINI
         
Exit Sub
BITCH:

End Sub

Public Sub Always_on_Top()
 
 Const flag As Long = SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOSIZE
  If OpcionesMusic.SiempreTop = True Then
      SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
  Else
      SetWindowPos frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
  End If
  
End Sub

'+----------------------------------------------------------------------------------------+
'|             TRASPARENCIA                                                               |
'+----------------------------------------------------------------------------------------+

Sub Make_Transparent(lHwnd As Long, Porcentaje As Integer)
 On Error GoTo Hell
  '// procedimento para hacer transparente en porcentaje los formularios
  '// parametros
  '// [LHwnD] -> Manejador para a kual aplikar el efekto
  '// [Porcentaje] -> pus que va ser el ...che porcentaje
  
  '// only work with win 2000 and later
  
  Dim OSV As OSVersionInfo
    
  '/* Get OS compatability flag
  OSV.OSVSize = Len(OSV)
  If GetVersionEx(OSV) <> 1 Then Exit Sub
       
  If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then Exit Sub '/* Win 98/ME
  If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then '/* Win 2000/XP
    Call SetWindowLong(lHwnd, GWL_EXSTYLE, GetWindowLong(lHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(lHwnd, 0, (Porcentaje * 255) / 100, LWA_ALPHA)
  End If
Exit Sub
Hell:
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  PROCEDIMIENTO PARA ARRASTRAR EL FORMULARIO SOLO DEKLARARLO EN MOUSE DOWN             |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Sub FormDrag(TheForm As Form)
  ReleaseCapture
  Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

