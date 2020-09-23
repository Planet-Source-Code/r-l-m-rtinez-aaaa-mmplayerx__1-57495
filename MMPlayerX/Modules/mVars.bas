Attribute VB_Name = "mStart"
Option Explicit
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|   VARIABLES UTILIZADAS PARA TODO EL PROGRAMA                                          |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public strRutaCaratula As String
Public CopyMp3Totales As Integer
Public CopyTotalAlbums As Integer

Public bolCaratulaShow As Boolean, bolDirectoriosShow As Boolean
Public bolAcercaShow As Boolean, bolOpcionesShow As Boolean
Public bolLyricsShow As Boolean, bolTagsShow As Boolean
Public bolVisShow As Boolean, bolSearchShow As Boolean



Public OriginalWallpaperStyle As Integer
Public OriginalTileWallpaper As Integer
Public OriginalRutaWallpaper As String

Public bolCaratulaDefault As Boolean
Public bLoadRegionFile As Boolean
Public bolSplashScreen As Boolean
Public intActiveAlbum As Integer
Public TotalAlbumS As Integer
Public MP3totales As Integer
Public sTextScroll As String
Public sFileMainPlaying As String
Public PlayerState As String
Public bSearching As Boolean
Public bMinimize As Boolean
Public bLoading As Boolean
Public strPathern As String
Public sFileType As String
Public sFormatPlayList As String
Public sFormatScroll As String
Public iScrollType As Integer
Public iScrollVel As Integer
Public iIdAlbumRC As Integer
Public bAddFiles As Boolean
Public IndexVisualization As Integer

Public iCrossfadeTrack As Integer
Public iCrossfadeStop As Integer
Public bPlayStarting As Boolean

'=======================================================
' VISUALIZACION
Public Type ptVisSpect
  Exist As Boolean
  BackColor As Long
  Mirrored As Boolean
  DrawSource As Integer
  ScaleUp As Integer
  ImageFile As String
  DrawBars As Boolean
  Gradient As String
  GrandientIndex As Integer
  Bars As Integer
  Spacio As Integer
  BackColorBar As Long
  DrawPeaks As Boolean
  BackColorPeak As Long
  arryPeaks() As Single
  arryWaitPeak() As String
  PeakHeight As Integer
  PeakGravity As Integer
End Type

Public Type ptVisScope
  LinesScope As Integer
  BackColorScope As Long
  Align As Integer
End Type

Public tConfigVis As ptVisSpect
Public tConfigScope() As ptVisScope

'=======================================================
Public Type Entry
    NoAlteraR As Boolean
    Mosaico As Boolean
    Centrar As Boolean
    Proporcional As Boolean
    Expander As Boolean
    Directorio As Boolean
    Language As String
    Ingles As Boolean
    Alpha As Integer
    SiempreTop As Boolean
    Splash As Boolean
    Instancias As Boolean
    TaskBar As Boolean
    SysTray As Boolean
End Type

Public Type ptSpec
  bDrawBars As Boolean
  iBars As Integer
  iSpacio As Integer
  lBackColorBar As Long
  lLineColorBar As Long
  bDrawPeaks As Boolean
  lBackColorPeak As Long
  iPeakHeight As Integer
  iPeakGravity As Integer
  iLinesScope As Integer
  lBackColorScope As Long
End Type

Public Type ptSlider
  Width As Integer
  Height As Integer
End Type

Public Type ptApp
  AppPath As String
  AppConfig As String
  Skin As String
End Type

Public Enum peCrossfade
  CrossfadeNormal = 0
  FadeIn = 1
  FadeOut = 2
End Enum

Public tAppConfig As ptApp
Public tSpectrum As ptSpec

Public Type TrayIcon
    Previous As Boolean
    Play As Boolean
    Pause As Boolean
    Stop As Boolean
    Next As Boolean
End Type

Public bMiniMask As Boolean
Public OpcionesMusic As Entry
Public PlayerTrayIcon As TrayIcon
Public peCrossFadeType As peCrossfade

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  INICIO DE LA APLICATION                                                              |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Sub Main()
  On Error Resume Next
  Dim strPath As String, args As String
  
 '----------------------------------------
 '// Optional Load XP Theme need the component [Microsoft Windows Common Controls 5.0]
    XPStyle False
 '----------------------------------------
   
 '// running right click in explorer or other
 '// HKCR\Directory\shell\Search Music Mp3 Player X\command\predeterminado
bLoading = True
   args = Trim$(Command$)
 If args <> "" Then '// si lo executan desde el explorador solo buskar alli
    Load_Settings_INI False
    frmMain.Search_Files (args)
 Else
    Load_Settings_INI True
    
     If MP3totales = 0 Then
       strPath = tAppConfig.AppPath
       frmMain.Search_Files (strPath)
       If MP3totales = 0 Then frmSearch.Show
     End If
 End If
 Change_Mask bMiniMask, False
 
bLoading = False
 
 If bPlayStarting = True Then
   frmMain.Play
 Else
   frmMain.Button(3).Selected = True
  If TotalAlbumS > 0 Then
     frmMain.Load_File_Tags
  Else
     sTextScroll = "No Albums Loaded"
     frmMain.ScrollText(1).CaptionText = sTextScroll
     frmMain.ScrollText(2).CaptionText = "00"
     frmMain.ScrollText(3).CaptionText = "00"
  End If
   frmMain.Draw_Spectrum
 End If

 Unload frmSplash
 frmMain.Visible = True
 
 If OpcionesMusic.TaskBar = True Then frmPopUp.Show
 If OpcionesMusic.SiempreTop = True Then Always_on_Top
 
 
 
  

End Sub

