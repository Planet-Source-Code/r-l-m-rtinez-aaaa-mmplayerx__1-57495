Attribute VB_Name = "Skins"

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|   API PARA QUITAR UN DETERMINADO COLOR DE UNA IMAGEN                                  |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

'Public Declare Function TransparentBlt Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Integer, ByVal nYOriginDest As Integer, ByVal nWidthDest As Integer, ByVal nHeightDest As Integer, ByVal hDCSrc As Long, ByVal nXOriginSrc As Integer, ByVal nYOriginSrc As Integer, ByVal nWidthSrc As Integer, ByVal nHeightSrc As Integer, ByVal crTransparent As Long) As Boolean


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|   APIS PARA CREAR EL EFECTO DE PROPORCIONAL PARA EL WALLPAPER                         |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As PointAPI) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const STRETCH_HALFTONE  As Long = &H4&

Public Type PointAPI
    x  As Long
    Y  As Long
End Type

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| APIS PARA EFECTO DE CONTORNO DEL FORMULARIO                                           |
'| USADAS PARA TRATAMIENTO DE IMAGENES                                                   |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Declare Function GetRegionDataByte Lib "gdi32" Alias "GetRegionData" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Byte) As Long
Public Declare Function GetRegionDataLong Lib "gdi32" Alias "GetRegionData" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Long) As Long
Public Declare Function ExtCreateRegionByte Lib "gdi32" Alias "ExtCreateRegion" (lpXform As Long, ByVal nCount As Long, lpRgnData As Byte) As Long
Public Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal Y As Long) As Long

Type RegionDataType
    RegionData() As Byte
    DataLength As Long
End Type


Public Const RGN_OR = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Private EdgeRegions(1) As RegionDataType

Private tConfigSlider(4) As ptSlider
Private iAlbumsShow As Integer, iAlbumsCols As Integer, iAlbumsRows As Integer

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


Public Sub Read_Config_Skin()
'// procedimento para leer las configuraciones del skin de color y posicion de los
'// botones
  Dim i As Integer, arryS(4) As String, arry() As String
  Dim s As String
On Error Resume Next
 With frmMain

  '// Previous
  Read_Config_Button .Button(0), "NormalMode", "Previous", "12,137,23,18"
  '// Play
  Read_Config_Button .Button(1), "NormalMode", "Play", "37,137,23,18"
  '// Pause
  Read_Config_Button .Button(2), "NormalMode", "Pause", "62,137,23,18"
  '// Stop
  Read_Config_Button .Button(3), "NormalMode", "Stop", "87,137,23,18"
  '// Next
  Read_Config_Button .Button(4), "NormalMode", "Next", "112,137,23,18"
  '// Intro
  Read_Config_Button .Button(5), "NormalMode", "Intro", "21,62,15,13"
  '// Mute
  Read_Config_Button .Button(6), "NormalMode", "Mute", "49,62,15,13"
  '// Repeat
  Read_Config_Button .Button(7), "NormalMode", "Repeat", "77,62,15,13"
  '// Randomize
  Read_Config_Button .Button(8), "NormalMode", "Randomize", "105,62,15,13"
  '// Previous Album
  Read_Config_Button .Button(9), "NormalMode", "PreviousAlbum", "195,1,30,12"
  '// Front
  Read_Config_Button .Button(10), "NormalMode", "Front", "227,1,21,12"
  '// Next Album
  Read_Config_Button .Button(11), "NormalMode", "NextAlbum", "250,1,30,12"
 '// Menu button
  Read_Config_Button .Button(12), "NormalMode", "Menu", "1,1,10,10"
  '// Minimize
  Read_Config_Button .Button(13), "NormalMode", "Minimize", "305,1,10,10"
  '// Minimode
  Read_Config_Button .Button(14), "NormalMode", "MiniMode", "316,1,10,10"
  '// Close
  Read_Config_Button .Button(15), "NormalMode", "Close", "327,1,10,10"
  '// PosBar
  Read_Config_Button .Slider(0), "NormalMode", "PosBar", "1,120,144,10,V"
  '// volume bar
  Read_Config_Button .Slider(1), "NormalMode", "VolBar", "148,25,10,121,H"
  '// time
  Read_Config_Button .ScrollText(0), "NormalMode", "Time", "3,89,32,6"
  '// track title normal mode
  Read_Config_Button .ScrollText(1), "NormalMode", "TrackTitle", "1,110,144,6"
  .ScrollText(1).BackColor = Read_INI("NormalMode", "TTBackColor", RGB(0, 0, 0), True)
  
  '// Bit Rate
  Read_Config_Button .ScrollText(2), "NormalMode", "BitRate", "38,80,15,6"
  '// Frequencia
  Read_Config_Button .ScrollText(3), "NormalMode", "Freq", "38,90,10,6"
  
  '// albums
  Read_Config_Button .btnAlbum(1), "NormalMode", "Albums", "3,18,10,10"
  iAlbumsShow = CInt(Read_INI("NormalMode", "AlbumsShow", "36"))
  If iAlbumsShow <= 5 Then iAlbumsShow = 5
  iAlbumsCols = CInt(Read_INI("NormalMode", "AlbumsCols", "12"))
  If iAlbumsCols <= 0 Then iAlbumsCols = 1
  iAlbumsRows = CInt(Read_INI("NormalMode", "AlbumsRows", "3"))
  If iAlbumsRows <= 0 Then iAlbumsRows = 1
  
  '// list reproduction
  Read_Config_Button .picListRep, "NormalMode", "ListRep", "162,15,175,144"
  Read_Config_Button .ImgCaratula, "NormalMode", "ListRep", "162,15,175,144"
  
  .ListRep.Width = .picListRep.Width + 20
  .ListRep.Height = .picListRep.Height + 2
  .ListRep.BackColor = Read_INI("NormalMode", "ListRepBackColor", RGB(0, 0, 0), True)
  .ListRep.ForeColor = Read_INI("NormalMode", "ListRepForeColor", RGB(255, 255, 255), True)
  
  '// list bar
  .Slider(2).Width = Read_INI("NormalMode", "ListBarWidth", "10")
  .Slider(2).Height = .picListRep.Height
  .Slider(2).Left = .picListRep.Width - .Slider(2).Width
  .Slider(2).Top = 0
  
  '// spectrum
  Read_Config_Button .picSpectrum, "NormalMode", "Spectrum", "55,79,89,28"
  
  '// spectrum bars
  tSpectrum.bDrawBars = CBool(Read_INI("NormalMode", "DrawBars", True))
  
  tSpectrum.iBars = CInt(Read_INI("NormalMode", "Bars", 15))
  If tSpectrum.iBars > 50 Then tSpectrum.iBars = 50
  If tSpectrum.iBars <= 0 Then tSpectrum.iBars = 1
  
  tSpectrum.iSpacio = CInt(Read_INI("NormalMode", "SpaceBar", 2))
  If tSpectrum.iSpacio > 5 Then tSpectrum.iSpacio = 5
  If tSpectrum.iSpacio < 0 Then tSpectrum.iSpacio = 0
  
  tSpectrum.lBackColorBar = CLng(Read_INI("NormalMode", "BackColorBar", RGB(255, 255, 255), True))
  tSpectrum.lLineColorBar = CLng(Read_INI("NormalMode", "LineColorBar", RGB(255, 255, 255), True))
  
  '// spectrum peaks
  tSpectrum.bDrawPeaks = CBool(Read_INI("NormalMode", "DrawPeaks", True))
  
  tSpectrum.lBackColorPeak = CLng(Read_INI("NormalMode", "BackColorPeak", RGB(255, 255, 255), True))
  
  tSpectrum.iPeakHeight = CInt(Read_INI("NormalMode", "PeakHeight", 1))
  If tSpectrum.iPeakHeight > 3 Then tSpectrum.iPeakHeight = 3
  If tSpectrum.iPeakHeight <= 0 Then tSpectrum.iPeakHeight = 1
  
  tSpectrum.iPeakGravity = CInt(Read_INI("NormalMode", "PeakGravity", 1))
  If tSpectrum.iPeakGravity > 10 Then tSpectrum.iPeakGravity = 10
  If tSpectrum.iPeakGravity <= 0 Then tSpectrum.iPeakGravity = 1

  '// spectrum scope
  tSpectrum.iLinesScope = CInt(Read_INI("NormalMode", "LinesScope", 30))
  If tSpectrum.iLinesScope > 50 Then tSpectrum.iLinesScope = 50
  If tSpectrum.iLinesScope <= 0 Then tSpectrum.iLinesScope = 10
  
  tSpectrum.lBackColorScope = CLng(Read_INI("NormalMode", "BackColorScope", RGB(255, 255, 255), True))
    
  If tSpectrum.bDrawBars = False And tSpectrum.bDrawPeaks = False Then tSpectrum.bDrawBars = True
  
  arryS(0) = "PosSlider": arryS(1) = "VolSlider"
  arryS(2) = "ListSlider"
  arryS(3) = "PosSlider": arryS(4) = "VolSlider"
  
  '// sliders config normal Mode
  For i = 0 To 2
    s = Read_INI("NormalMode", arryS(i), "10,10")
    arry = Split(s, ",", , vbTextCompare)
    If UBound(arry) = 1 Then
       tConfigSlider(i).Width = arry(0)
       tConfigSlider(i).Height = arry(1)
    Else
       tConfigSlider(i).Width = 10
       tConfigSlider(i).Height = 10
    End If
  Next i


 '// sliders config minimode
  For i = 3 To 4
    s = Read_INI("MiniMode", arryS(i), "10,10")
    arry = Split(s, ",", , vbTextCompare)
    If UBound(arry) = 1 Then
       tConfigSlider(i).Width = arry(0)
       tConfigSlider(i).Height = arry(1)
    Else
       tConfigSlider(i).Width = 10
       tConfigSlider(i).Height = 10
    End If
  Next i

  
  '=====================================================================
  '  MINI MODE
  '=====================================================================
  
  '// Previous
  Read_Config_Button .ButtonMini(0), "MiniMode", "Previous", "172,1,10,10"
  '// Play
  Read_Config_Button .ButtonMini(1), "MiniMode", "Play", "183,1,10,10"
  '// Pause
  Read_Config_Button .ButtonMini(2), "MiniMode", "Pause", "194,1,10,10"
  '// Stop
  Read_Config_Button .ButtonMini(3), "MiniMode", "Stop", "205,1,10,10"
  '// Next
  Read_Config_Button .ButtonMini(4), "MiniMode", "Next", "216,1,10,10"
  '// Menu button
  Read_Config_Button .ButtonMini(5), "MiniMode", "Menu", "1,1,10,10"
  '// Minimize
  Read_Config_Button .ButtonMini(6), "MiniMode", "Minimize", "239,1,10,10"
  '// Minimode
  Read_Config_Button .ButtonMini(7), "MiniMode", "NormalMode", "250,1,10,10"
  '// Close
  Read_Config_Button .ButtonMini(8), "MiniMode", "Close", "261,1,10,10"
  '// time
  Read_Config_Button .ScrollText(4), "MiniMode", "Time", "13,3,25,6"
  '// track title normal mode
  Read_Config_Button .ScrollText(5), "MiniMode", "TrackTitle", "43,3,128,6"
  .ScrollText(5).BackColor = Read_INI("MiniMode", "TTBackColor", RGB(0, 0, 0), True)
  '// PosBar
  Read_Config_Button .Slider(3), "MiniMode", "PosBar", "41,13,97,6,V"
  '// volume bar
  Read_Config_Button .Slider(4), "MiniMode", "VolBar", "147,13,58,6,V"
  

        
           
  If bolLyricsShow = True Then
    frmLyrics.picLyrics.BackColor = .ListRep.BackColor
    frmLyrics.picBody.BackColor = .ListRep.BackColor
    frmLyrics.shpFocus.BorderColor = .ListRep.ForeColor
    frmLyrics.lblNoLyrics.ForeColor = .ListRep.ForeColor
          
    frmLyrics.Order_lblLyrics
    frmMain.LyricsIndex = 1
  End If
 End With
 
End Sub

Public Function MakeRegion(picSkin As PictureBox) As Long
 '// procedimento usado para hacer los bordes irregulares del formulario
 '// basado en un picture recorriendo pixel por pixel para buscar las areas
 '// que seran trasparentes o ireegulares
 
    Dim x As Long, Y As Long, StartLineX As Long
    Dim LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean
    Dim hdc As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hdc = picSkin.hdc
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    x = Y = StartLineX = 0
    '// Leer cual sera el color trasparente para el formulario
     TransparentColor = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
    
    For Y = 0 To PicHeight - 1
        For x = 0 To PicWidth - 1
            '// si el pixel es del color trasparente
            If GetPixel(hdc, x, Y) = TransparentColor Or x = PicWidth Then
                '// buscar los pixiles trasparentes
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, x, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        '// siempre borrar
                        DeleteObject LineRegion
                    End If
                End If
            Else
                '// buscar los pixeles de no transparente color
                If Not InLine Then
                    InLine = True
                    StartLineX = x
                End If
            End If
        Next
    Next
     MakeRegion = FullRegion
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Load_Buttons_Skin()
'// procedimiento para cargar todos los controles, ponerlos en su lugar

  Dim srcX As Integer, srcY As Integer, srcWidth As Integer, srcHeight As Integer
  Dim i As Integer, j As Integer, k As Integer
  Dim pTemp(11) As StdPicture, pImage As StdPicture
  Dim s As String
  Dim lColorTran As Long
  
  On Error Resume Next
With frmMain
  s = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\"
 
 .picNormalMode.Cls
 .picNormalMode.Picture = LoadPicture()
 .picNormalMode.Width = 5085
 .picNormalMode.Height = 2415
 .picNormalMode.Picture = LoadPicture(s & "main.bmp")
 .picNormalMode.AutoSize = True
 
 .picMiniMode.Cls
 .picMiniMode.Picture = LoadPicture()
 .picMiniMode.Width = 4110
 .picMiniMode.Height = 330
 .picMiniMode.Picture = LoadPicture(s & "minimode.bmp")
 .picNormalMode.AutoSize = True
 '// time font
 Set .ScrollText(0).PictureText = LoadPicture(s & "num_font.bmp")
 '// track title
 Set .ScrollText(1).PictureText = LoadPicture(s & "song_font.bmp")
 '// bitrate text
 Set .ScrollText(2).PictureText = LoadPicture(s & "songinfo_font.bmp")
 '// frecuencia text
 Set .ScrollText(3).PictureText = LoadPicture(s & "songinfo_font.bmp")
 
 '// time minimode
 Set .ScrollText(4).PictureText = LoadPicture(s & "num_minimode_font.bmp")
 '// track title minimode
 Set .ScrollText(5).PictureText = LoadPicture(s & "song_minimode_font.bmp")
 
 
 Set pTemp(0) = LoadPicture(s & "player_buttons.bmp")
 Set pTemp(1) = LoadPicture(s & "options_buttons.bmp")
 Set pTemp(2) = LoadPicture(s & "albums_buttons.bmp")
 Set pTemp(3) = LoadPicture(s & "titlebar_buttons.bmp")
 Set pTemp(4) = LoadPicture(s & "posbar_slider.bmp")
 Set pTemp(5) = LoadPicture(s & "volbar_slider.bmp")
 Set pTemp(6) = LoadPicture(s & "listbar_slider.bmp")
 
 '// minimode pictures
 Set pTemp(7) = LoadPicture(s & "posbar_minimode_slider.bmp")
 Set pTemp(8) = LoadPicture(s & "volbar_minimode_slider.bmp")
 
 
 Set pTemp(9) = LoadPicture(s & "albums_picture.bmp")
 
 '// minimode
 Set pTemp(10) = LoadPicture(s & "player_minimode_buttons.bmp")
 Set pTemp(11) = LoadPicture(s & "titlebar_minimode_buttons.bmp")
  
 lColorTran = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
 
 .picTemp.BackColor = &H808080
 
  For i = 0 To 15
    .Button(i).Reset
    .Button(i).MaskColor = lColorTran
    srcWidth = .Button(i).Width
    srcHeight = .Button(i).Height
        
    .picTemp.Width = srcWidth
    .picTemp.Height = srcHeight
    
    '// copy picture back
    .picTemp.Picture = LoadPicture()
    .picTemp.PaintPicture .picNormalMode.Image, 0, 0, srcWidth, srcHeight, .Button(i).Left, .Button(i).Top, srcWidth, srcHeight
    .picTemp.Picture = .picTemp.Image
    Set .Button(i).PictureBack = .picTemp.Picture
    
    If i = 0 Then '// play buttons
      srcX = 0
      Set pImage = pTemp(0)
    ElseIf i = 5 Then '// options buttons
           srcX = 0
           Set pImage = pTemp(1)
        ElseIf i = 9 Then '// albums buttons
               srcX = 0
               Set pImage = pTemp(2)
            ElseIf i = 12 Then '// titlebar buttons
                    srcX = 0
                    Set pImage = pTemp(3)
                 End If
                 
     For j = 0 To 3
       srcY = srcHeight * j
       .picTemp.Picture = LoadPicture()
       .picTemp.PaintPicture pImage, 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
       .picTemp.Picture = .picTemp.Image
       
       If j = 0 Then
         Set .Button(i).PictureNormal = .picTemp.Picture
       ElseIf j = 1 Then
             Set .Button(i).PictureOver = .picTemp.Picture
           ElseIf j = 2 Then
                 Set .Button(i).PictureDown = .picTemp.Picture
               Else
                 Set .Button(i).PictureDisabled = .picTemp.Picture
               End If
     Next j
    
   srcX = srcX + srcWidth
   
   DoEvents
  Next i
  
  '// Sliders pos - vol - list  ---- and minimode
  
  For i = 0 To 4
    .picTemp.BackColor = &H808080
    .Slider(i).ResetPictures
    srcX = 0
    srcWidth = .Slider(i).Width
    srcHeight = .Slider(i).Height
        
    .picTemp.Width = srcWidth
    .picTemp.Height = srcHeight
        
    srcY = 0
    '// picture back
    .picTemp.Picture = LoadPicture()
    .picTemp.PaintPicture pTemp(i + 4), 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
    .picTemp.Picture = .picTemp.Image
    Set .Slider(i).PictureBack = .picTemp.Picture
     
    srcY = srcHeight
    '// picture progress
    .picTemp.Picture = LoadPicture()
    .picTemp.PaintPicture pTemp(i + 4), 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
    .picTemp.Picture = .picTemp.Image
    Set .Slider(i).PictureProgress = .picTemp.Picture
     
    
    .picTemp.BackColor = &HC0&
    
     '// .Sliders
    srcX = srcWidth
    srcWidth = tConfigSlider(i).Width
    srcHeight = tConfigSlider(i).Height
    
    .picTemp.Width = srcWidth
    .picTemp.Height = srcHeight

    For j = 0 To 2
       srcY = srcHeight * j
       .picTemp.Picture = LoadPicture()
       .picTemp.PaintPicture pTemp(i + 4), 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
       .picTemp.Picture = .picTemp.Image
       
       If j = 0 Then
          Set .Slider(i).Bar = .picTemp.Picture
       ElseIf j = 1 Then
              Set .Slider(i).BarOver = .picTemp.Picture
           Else
              Set .Slider(i).BarDown = .picTemp.Picture
           End If
    Next j
       
   DoEvents
  Next i



'//==============================================================================
'// Botones de  minimode
'//==============================================================================

.picTemp.BackColor = &H808080

  For i = 0 To 8
    .ButtonMini(i).Reset
    .ButtonMini(i).MaskColor = lColorTran
    srcWidth = .ButtonMini(i).Width
    srcHeight = .ButtonMini(i).Height
        
    .picTemp.Width = srcWidth
    .picTemp.Height = srcHeight
    
    '// copy picture back
    .picTemp.Picture = LoadPicture()
    .picTemp.PaintPicture .picMiniMode.Image, 0, 0, srcWidth, srcHeight, .ButtonMini(i).Left, .ButtonMini(i).Top, srcWidth, srcHeight
    .picTemp.Picture = .picTemp.Image
    Set .ButtonMini(i).PictureBack = .picTemp.Picture
    
    If i = 0 Then '// play buttons
      srcX = 0
      Set pImage = pTemp(10)
    ElseIf i = 5 Then '// options buttons
           srcX = 0
           Set pImage = pTemp(11)
         End If
                 
     For j = 0 To 3
       srcY = srcHeight * j
       .picTemp.Picture = LoadPicture()
       .picTemp.PaintPicture pImage, 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
       .picTemp.Picture = .picTemp.Image
       
       If j = 0 Then
         Set .ButtonMini(i).PictureNormal = .picTemp.Picture
       ElseIf j = 1 Then
             Set .ButtonMini(i).PictureOver = .picTemp.Picture
           ElseIf j = 2 Then
                 Set .ButtonMini(i).PictureDown = .picTemp.Picture
               Else
                 Set .ButtonMini(i).PictureDisabled = .picTemp.Picture
               End If
     Next j
    
   srcX = srcX + srcWidth
   
   DoEvents
  Next i
  
  
  .picTemp.BackColor = &H8000&
  
 '// albums
    srcX = 0
    srcWidth = .btnAlbum(1).Width
    srcHeight = .btnAlbum(1).Height
    
    .picTemp.Width = srcWidth
    .picTemp.Height = srcHeight
    
    .btnAlbum(1).MaskColor = lColorTran
    
    '// copiar fondo al boton
    .picTemp.Picture = LoadPicture()
    .picTemp.PaintPicture .picNormalMode.Image, 0, 0, srcWidth, srcHeight, .btnAlbum(1).Left, .btnAlbum(1).Top, srcWidth, srcHeight
    .picTemp.Picture = .picTemp.Image
    
   
    Set .btnAlbum(1).PictureBack = .picTemp.Picture
     For j = 0 To 3
       srcY = srcHeight * j
       .picTemp.Picture = LoadPicture()
       .picTemp.PaintPicture pTemp(9), 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
       .picTemp.Picture = .picTemp.Image
       
       If j = 0 Then
         Set .btnAlbum(1).PictureNormal = .picTemp.Picture
       ElseIf j = 1 Then
             Set .btnAlbum(1).PictureOver = .picTemp.Picture
           ElseIf j = 2 Then
                 Set .btnAlbum(1).PictureDown = .picTemp.Picture
               Else
                 Set .btnAlbum(1).PictureDisabled = .picTemp.Picture
               End If
     Next j

 
 Set pTemp(0) = LoadPicture()
 Set pTemp(1) = LoadPicture()
 Set pTemp(2) = LoadPicture()
 Set pTemp(3) = LoadPicture()
 Set pTemp(4) = LoadPicture()
 Set pTemp(5) = LoadPicture()
 Set pTemp(6) = LoadPicture()
 Set pTemp(7) = LoadPicture()
 Set pTemp(8) = LoadPicture()
 Set pTemp(9) = LoadPicture()
 Set pTemp(10) = LoadPicture()
 Set pTemp(11) = LoadPicture()
 Set pImage = LoadPicture()
 
  .picTemp = LoadPicture()
  
 End With
End Sub

Public Sub Load_Albums()
   
 Dim iAlbum As Integer, iTop As Integer, ileft As Integer
 Dim srcWidth As Integer, srcHeight As Integer
 On Error Resume Next
 iAlbum = 2
With frmMain
 iTop = .btnAlbum(1).Top
 
 If TotalAlbumS >= 1 Then
   frmMain.btnAlbum(1).Visible = True
   frmMain.btnAlbum(1).Selected = False
 End If
 
 For i = iAlbumsRows To TotalAlbumS
   .btnAlbum(i).Visible = False
 Next i
 
 For i = 1 To iAlbumsRows
     For j = 1 To iAlbumsCols - 1
       If iAlbum > iAlbumsShow Or iAlbum > TotalAlbumS Then Exit Sub
         .btnAlbum(iAlbum).Top = iTop
         .btnAlbum(iAlbum).Left = .btnAlbum(iAlbum - 1).Left + .btnAlbum(iAlbum - 1).Width + 2
         
         .btnAlbum(iAlbum).Reset
         .btnAlbum(iAlbum).Width = .btnAlbum(1).Width
         .btnAlbum(iAlbum).Height = .btnAlbum(1).Height
         srcWidth = .btnAlbum(iAlbum).Width
         srcHeight = .btnAlbum(iAlbum).Height

        '// copiar fondo al boton
         .picTemp.Picture = LoadPicture()
         .picTemp.PaintPicture .picNormalMode.Image, 0, 0, srcWidth, srcHeight, .btnAlbum(iAlbum).Left, .btnAlbum(iAlbum).Top, srcWidth, srcHeight
         .picTemp.Picture = .picTemp.Image
         
         .btnAlbum(iAlbum).MaskColor = .btnAlbum(1).MaskColor
        
        Set .btnAlbum(iAlbum).PictureBack = .picTemp.Picture
     
        Set .btnAlbum(iAlbum).PictureNormal = .btnAlbum(1).PictureNormal
        Set .btnAlbum(iAlbum).PictureOver = .btnAlbum(1).PictureOver
        Set .btnAlbum(iAlbum).PictureDown = .btnAlbum(1).PictureDown
        Set .btnAlbum(iAlbum).PictureDisabled = .btnAlbum(1).PictureDisabled
        Set .btnAlbum(iAlbum).MouseIcon = .btnAlbum(1).MouseIcon
            .btnAlbum(iAlbum).Selected = False
            .btnAlbum(iAlbum).Visible = True
        
         iAlbum = iAlbum + 1

     Next j
     
     If iAlbum > iAlbumsShow Or iAlbum > TotalAlbumS Then Exit Sub
    
     iTop = iTop + (.btnAlbum(1).Height + 2)
     .btnAlbum(iAlbum).Top = iTop
     .btnAlbum(iAlbum).Left = .btnAlbum(1).Left

    .btnAlbum(iAlbum).Reset
    .btnAlbum(iAlbum).Width = .btnAlbum(1).Width
    .btnAlbum(iAlbum).Height = .btnAlbum(1).Height
    srcWidth = .btnAlbum(iAlbum).Width
    srcHeight = .btnAlbum(iAlbum).Height

     '// copiar fondo al boton
     .picTemp.Picture = LoadPicture()
     .picTemp.PaintPicture .picNormalMode.Image, 0, 0, srcWidth, srcHeight, .btnAlbum(iAlbum).Left, .btnAlbum(iAlbum).Top, srcWidth, srcHeight
     .picTemp.Picture = .picTemp.Image
     
     .btnAlbum(iAlbum).MaskColor = .btnAlbum(1).MaskColor
     
     Set .btnAlbum(iAlbum).PictureBack = .picTemp.Picture
     
     Set .btnAlbum(iAlbum).PictureNormal = .btnAlbum(1).PictureNormal
     Set .btnAlbum(iAlbum).PictureOver = .btnAlbum(1).PictureOver
     Set .btnAlbum(iAlbum).PictureDown = .btnAlbum(1).PictureDown
     Set .btnAlbum(iAlbum).PictureDisabled = .btnAlbum(1).PictureDisabled
     Set .btnAlbum(iAlbum).MouseIcon = .btnAlbum(1).MouseIcon
         .btnAlbum(iAlbum).Selected = False
         .btnAlbum(iAlbum).Visible = True
     iAlbum = iAlbum + 1
    
 Next i
 
End With
End Sub

Public Sub Load_Cursors()
  On Error Resume Next
  Dim sPath As String, sCursor As String
  Dim i As Integer
   
  sPath = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\"
  
 With frmMain
   '=======================================================================
   ' NORMAL MODE
   '=======================================================================
    '// cursor principal
    sCursor = ""
    If Dir(sPath & "curMain.cur") <> "" Then sCursor = sPath & "curMain.cur"
      .picNormalMode.MouseIcon = LoadPicture(sCursor)
        
    '// cursor para los botones de normalmode
    sCursor = ""
    If Dir(sPath & "curButtons.cur") <> "" Then sCursor = sPath & "curButtons.cur"
 
    For i = 0 To 15
       '//mascara normal
       Set .Button(i).MouseIcon = LoadPicture(sCursor)
    Next i
    
     '// cursor para los albums
     sCursor = ""
     If Dir(sPath & "curAlbums.cur") <> "" Then sCursor = sPath & "curAlbums.cur"
     Set .btnAlbum(1).MouseIcon = LoadPicture(sCursor)
     
     '// cursor posbar
    sCursor = ""
    If Dir(sPath & "curposbar.cur") <> "" Then sCursor = sPath & "curposbar.cur"
       Set .Slider(0).MouseIcon = LoadPicture(sCursor)
     
    '// cursor vol bar
    sCursor = ""
    If Dir(sPath & "curvolbar.cur") <> "" Then sCursor = sPath & "curvolbar.cur"
       Set .Slider(1).MouseIcon = LoadPicture(sCursor)
     
     '// Cursor para la lista de reproduccion
     sCursor = ""
     If Dir(sPath & "curListRep.cur") <> "" Then sCursor = sPath & "curlistrep.cur"
       .ListRep.MouseIcon = LoadPicture(sCursor)
       
    '// Cursor para la listbar
     sCursor = ""
     If Dir(sPath & "curListbar.cur") <> "" Then sCursor = sPath & "curlistbar.cur"
        Set .Slider(2).MouseIcon = LoadPicture(sCursor)
       
'=============================================================================
' MINI MODE
'=============================================================================
    '// cursor minimode
    sCursor = ""
    If Dir(sPath & "curMiniMode.cur") <> "" Then sCursor = sPath & "curMiniMode.cur"
      .picMiniMode.MouseIcon = LoadPicture(sCursor)
      
    '// cursor para los botones de minimode
    sCursor = ""
    If Dir(sPath & "curButtons_minimode.cur") <> "" Then sCursor = sPath & "curButtons_minimode.cur"
    
    For i = 0 To 8
       '// minimascara
       Set .ButtonMini(i).MouseIcon = LoadPicture(sCursor)
    Next i
    
    '// cursor posbar
    sCursor = ""
    If Dir(sPath & "curposbar_minimode.cur") <> "" Then sCursor = sPath & "curposbar_minimode.cur"
       Set .Slider(3).MouseIcon = LoadPicture(sCursor)
     
    '// cursor vol bar
    sCursor = ""
    If Dir(sPath & "curvolbar_minimode.cur") <> "" Then sCursor = sPath & "curvolbar_minimode.cur"
       Set .Slider(4).MouseIcon = LoadPicture(sCursor)
      
 End With
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Change_Skin(SkinName As String)
 On Error Resume Next
 
  tAppConfig.Skin = SkinName

'---------------------------------------------------------------------------------------
'// leer la configuracion del skin de las posiciones de los botones
   Read_Config_Skin

'----------------------------------------------------------------------------------------
'// colocar los botones si tienen partes que pueden ser transparentes
   Load_Buttons_Skin

'----------------------------------------------------------------------------------------
'// kargar todos los cursores
   Load_Cursors
  
'---------------------------------------------------------------------------------------
'// cargar los albums
   Load_Albums
 
 With frmMain
   .picSpectrum.PaintPicture .picNormalMode.Image, 0, 0, .picSpectrum.ScaleWidth, .picSpectrum.ScaleHeight, .picSpectrum.Left, .picSpectrum.Top, .picSpectrum.ScaleWidth, .picSpectrum.ScaleHeight
   .picSpectrum.Picture = .picSpectrum.Image
    If intActiveAlbum > 0 Then .btnAlbum(intActiveAlbum).Selected = True

 End With
 
End Sub

Public Sub Change_Mask(MiniMask As Boolean, bNormal As Boolean)
 On Error Resume Next
 Dim FormLeft As Long, FormTop As Long
 Dim NewRegion As Long
 If MiniMask = True Then
   bMiniMask = True
   
   frmMain.ScrollText(5).CaptionText = sTextScroll
   frmMain.ScrollText(5).ToolTipText = sTextScroll
   '// posbar
   frmMain.Slider(3).Max = frmMain.Slider(0).Max
   frmMain.Slider(3).Value = frmMain.Slider(0).Value
   '// volbar
   frmMain.Slider(4).Value = frmMain.VolumeNActuaL

   
    frmMain.picNormalMode.Visible = False
    frmMain.picMiniMode.Visible = True
    frmMain.Width = frmMain.picMiniMode.Width
    frmMain.Height = frmMain.picMiniMode.Height
    
    ' The API call requires the address of the region data,
    ' so we pass the first cell in the array. VB passes arrays
    ' ByRef, so here's our address.
       
    NewRegion = ExtCreateRegionByte(ByVal 0&, EdgeRegions(1).DataLength, EdgeRegions(1).RegionData(0))
    SetWindowRgn frmMain.hwnd, NewRegion, True
    DeleteObject NewRegion
   
  If bNormal = True Then
    FormLeft = frmMain.Left + (frmMain.Button(14).Left * Screen.TwipsPerPixelX)
    FormLeft = FormLeft - (frmMain.ButtonMini(7).Left * Screen.TwipsPerPixelX) + (frmMain.ButtonMini(7).Width * Screen.TwipsPerPixelX)
    frmMain.Left = FormLeft
  
    FormTop = frmMain.Top + (frmMain.Button(14).Top * Screen.TwipsPerPixelY)
    FormTop = FormTop - (frmMain.ButtonMini(7).Top * Screen.TwipsPerPixelY)
  
    frmMain.Top = FormTop
  End If
    
 Else
    bMiniMask = False
    frmMain.ScrollText(1).CaptionText = sTextScroll
    frmMain.ScrollText(1).ToolTipText = sTextScroll
   '// posbar
   frmMain.Slider(0).Max = frmMain.Slider(3).Max
   frmMain.Slider(0).Value = frmMain.Slider(3).Value
   '// volbar
   frmMain.Slider(1).Value = frmMain.VolumeNActuaL
   
    frmMain.picMiniMode.Visible = False
    frmMain.picNormalMode.Visible = True
    frmMain.Width = frmMain.picNormalMode.Width
    frmMain.Height = frmMain.picNormalMode.Height
            
    NewRegion = ExtCreateRegionByte(ByVal 0&, EdgeRegions(0).DataLength, EdgeRegions(0).RegionData(0))
    SetWindowRgn frmMain.hwnd, NewRegion, True
    DeleteObject NewRegion
    
  If bNormal = True Then
    FormLeft = frmMain.Left + (frmMain.ButtonMini(7).Left * Screen.TwipsPerPixelX)
    FormLeft = FormLeft - (frmMain.Button(14).Left * Screen.TwipsPerPixelX) - (frmMain.Button(14).Width * Screen.TwipsPerPixelX)
    frmMain.Left = FormLeft
   
    FormTop = frmMain.Top + (frmMain.ButtonMini(7).Top * Screen.TwipsPerPixelY)
    FormTop = FormTop - (frmMain.Button(14).Top * Screen.TwipsPerPixelY)

    frmMain.Top = FormTop
  End If
 End If

frmMain.Show_ToolTipText

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

'// procedimento para hacer calkular la maskara normal y la mini
Sub Form_Mini_Normal()
 Dim WinRegion As Long
 Dim Ret As Long
 
    '//-----------------------------------------------------------------
    '//  MASKARA NORMAL
    '//-----------------------------------------------------------------
    
    frmMain.picMiniMode.Move 0, 0
    frmMain.picNormalMode.Move 0, 0
 
   '// cargadas desde archivo
    If bLoadRegionFile = True Then
       If LoadRegions(EdgeRegions()) = True Then
          Exit Sub
       End If
    End If
 
    '// First create the region for the bitmap
    WinRegion = MakeRegion(frmMain.picNormalMode)
    '// Get the size needed for the region data buffer
    EdgeRegions(0).DataLength = GetRegionDataLong(WinRegion, 0&, ByVal 0&)

    If EdgeRegions(0).DataLength <> 0 Then
        ' Actually get the data into the buffer - a byte array
        ' of the proper size.
        ' You need 32 bytes more, because the API call attaches
        ' a 32-byte structure called RGNDATAHEADER before the
        ' data itself
        ReDim EdgeRegions(0).RegionData(EdgeRegions(0).DataLength + 32)
        
        Ret = GetRegionDataByte(WinRegion, EdgeRegions(0).DataLength, EdgeRegions(0).RegionData(0))
        
    End If
    
    '//-----------------------------------------------------------------
    '//  MINI MASCARA
    '//-----------------------------------------------------------------
    
    WinRegion = MakeRegion(frmMain.picMiniMode)
    EdgeRegions(1).DataLength = GetRegionDataLong(WinRegion, 0&, ByVal 0&)

    If EdgeRegions(1).DataLength <> 0 Then
        ReDim EdgeRegions(1).RegionData(EdgeRegions(1).DataLength + 32)
        Ret = GetRegionDataByte(WinRegion, EdgeRegions(1).DataLength, EdgeRegions(1).RegionData(0))
    End If
    
    SaveRegions EdgeRegions()
    DeleteObject WinRegion
     
End Sub

'=================================================================================
Public Sub SaveRegions(EdgeRegions() As RegionDataType)
 On Error GoTo HELL
 Dim i As Long
 Dim filename As String
   filename = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\regions.dat"
    Open filename For Binary As #1

    For i = 0 To 1
        Put 1, , EdgeRegions(i).DataLength
        Put 1, , EdgeRegions(i).RegionData
    Next

    Close
Exit Sub
HELL:
End Sub

'=================================================================================
Public Function LoadRegions(EdgeRegions() As RegionDataType) As Boolean
 On Error GoTo HELL
 Dim i As Long
 Dim filename As String
   
   filename = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\regions.dat"
   
    If Dir(filename) = "" Then Exit Function
    
    Open filename For Binary As #1
    
    For i = 0 To 1
        Get 1, , EdgeRegions(i).DataLength
        ReDim EdgeRegions(i).RegionData(EdgeRegions(i).DataLength + 32)
        Get 1, , EdgeRegions(i).RegionData
    Next
    
    Close
    
    LoadRegions = True
Exit Function
HELL:
 Close
 MsgBox err.Description
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Load_Skins_Menu(SelMenu As String)
'// Procedimiento para cargar los skins disponibles que son todos las carpetas
'// en la ruta del EXE mas \MMp3Player\Skins\  y los carga en el menu de frmpopup
'// parametros
'// [SelMenu] -> Menu el cual va ha estar seleccionado

Dim miNombre As String
Dim i As Integer
On Error Resume Next

MiRuta = tAppConfig.AppConfig & "skins\"
i = 0
miNombre = Dir(MiRuta, vbDirectory)   ' Recupera la primera entrada.

If miNombre = "" Then
 For i = 1 To frmPopUp.mnuSkinsAdd.count
   frmPopUp.mnuSkinsAdd(i).Caption = ""
   frmPopUp.mnuSkinsAdd(i).Visible = False
 Next i
 tAppConfig.Skin = "\No skin selected\"
 Exit Sub
End If

'/* para ver si hay imagnes en el directorio
frmPopUp.fileBmps.Pattern = "*.bmp"

Do While miNombre <> ""
   If miNombre <> "." And miNombre <> ".." Then
      ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
      If (GetAttr(MiRuta & miNombre) And vbDirectory) = vbDirectory Then
       frmPopUp.fileBmps.Path = MiRuta & miNombre
        If frmPopUp.fileBmps.ListCount > 0 Then
             i = i + 1
             
             If i <> 1 And i > frmPopUp.mnuSkinsAdd.count Then Load frmPopUp.mnuSkinsAdd(i)  '// cargar los menus dinamikamente
             
             frmPopUp.mnuSkinsAdd(i).Caption = " " & miNombre
             frmPopUp.mnuSkinsAdd(i).Checked = False
             frmPopUp.mnuSkinsAdd(i).Visible = True
             If LCase(miNombre) = LCase(SelMenu) Then frmPopUp.mnuSkinsAdd(i).Checked = True
          
        End If
      End If
   End If
  miNombre = Dir
Loop
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

'+--------------------------------------------------------------------------------------+
'|    CREAR LA IMAGEN DE WALLPAPER SEGUN LAS OPCIONES ESPECIFICADAS                     |
'+--------------------------------------------------------------------------------------+

Public Sub CreatePic(picSource As PictureBox, picDestination As PictureBox)
'// Procedimiento para krear el strech con la mas alta calidad posible
Dim hBrush          As Long
Dim hDummyBrush     As Long
Dim lOrigMode       As Long
Dim uBrushOrigPt    As PointAPI
Dim lWidth As Long
Dim lHeight As Long
Dim lLeft As Integer
Dim lTop As Integer
    picDestination.AutoRedraw = True
    picDestination.Cls
    lWidth = picDestination.Width
    lHeight = picDestination.Height
    lLeft = 0
    lTop = 0
    'Set picEdit's stretch mode to halftone (this may cause misalignment of the brush)
    lOrigMode = SetStretchBltMode(picDestination.hdc, STRETCH_HALFTONE)

    'Realign the brush...
    'Get picEdit's brush by selecting a dummy brush into the DC
    hDummyBrush = CreateSolidBrush(0)
    hBrush = SelectObject(picDestination.hdc, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    UnrealizeObject hBrush
    'Set picEdit's brush alignment coordinates to the left-top of the bitmap
    SetBrushOrgEx picDestination.hdc, lLeft, lTop, uBrushOrigPt
    'Now put the original brush back into the DC at the new alignment
    hDummyBrush = SelectObject(picDestination.hdc, hBrush)
    
    'Stretch the bitmap
    StretchBlt picDestination.hdc, lLeft, lTop, lWidth, lHeight, _
            picSource.hdc, 0, 0, picSource.Width, picSource.Height, vbSrcCopy
    
    'Set the stretch mode back to it's original mode
    SetStretchBltMode picDestination.hdc, lOrigMode
    
    'Reset the original alignment of the brush...
    'Get picEdit's brush by selecting the dummy brush back into the DC
    hBrush = SelectObject(picDestination.hdc, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    UnrealizeObject hBrush
    'Set the brush alignment back to the original coordinates
    SetBrushOrgEx picDestination.hdc, uBrushOrigPt.x, uBrushOrigPt.Y, uBrushOrigPt
    'Now put the original brush back into picEdit's DC at the original alignment
    hDummyBrush = SelectObject(picDestination.hdc, hBrush)
    'Get rid of the dummy brush
    DeleteObject hDummyBrush
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


'+--------------------------------------------------------------------------------------+
'|    CREAR LA IMAGEN DE WALLPAPER Y PONER EN EL ESCRITORIO                             |
'+--------------------------------------------------------------------------------------+

Public Sub ConfigurarWallpaper()
'// procedimiento para krear la imagen y ponerla en el escritorio como wallpaper
  On Error GoTo Bitch
    If frmPopUp.mnuWallpapper.Checked = False Then Exit Sub
       frmMain.picWallOriginal.Picture = Nothing
       frmMain.picWallOriginal.Width = 1
       frmMain.picWallOriginal.Height = 1
       
        If OpcionesMusic.NoAlteraR = True Then Exit Sub
         If Trim(strRutaCaratula) = "" Then '// no tiene caratula poner el default
           If bolCaratulaDefault = True Then Exit Sub '// ponerla solo una vez
           frmMain.picWallOriginal.Picture = frmPopUp.picDefaultLogo.Picture
           SavePicture frmMain.picWallOriginal.Image, DirectoriOWindowS & "MusicMp3.bmp"
           'PoneRWallPapeR "Mosaico"
           bolCaratulaDefault = True
           'GoTo Bitch
         Else  'si tiene caratula ponerla
           frmMain.picWallOriginal.Picture = LoadPicture(strRutaCaratula)
           bolCaratulaDefault = False
         End If
          
         '// Wallpaper estilo Expandido
         If OpcionesMusic.Expander Then
           SavePicture frmMain.picWallOriginal.Image, DirectoriOWindowS & "MusicMp3.bmp"
           PoneRWallPapeR "Expandido"
           Exit Sub
         End If
         
         '// Wallpaper Stylo proporcional
         If OpcionesMusic.Proporcional = True Then
            '----ajustar la ..che imagen para que quede chida-----------------------
            frmMain.picWallProp.Picture = Nothing
             If frmMain.picWallOriginal.Width > frmMain.picWallOriginal.Height Then
               frmMain.picWallProp.Width = Screen.Width
               frmMain.picWallProp.Height = frmMain.picWallOriginal.Height * Screen.Width / frmMain.picWallOriginal.Width
             Else
               frmMain.picWallProp.Height = Screen.Height
               frmMain.picWallProp.Width = frmMain.picWallOriginal.Width * Screen.Height / frmMain.picWallOriginal.Height
             End If
               CreatePic frmMain.picWallOriginal, frmMain.picWallProp
            '----------------------------------------------------------------------
            SavePicture frmMain.picWallProp.Image, DirectoriOWindowS & "MusicMp3.bmp"
              '// Wallpaper estilo Centrado
               If OpcionesMusic.Centrar = True Then
                 PoneRWallPapeR "Centro"
                 GoTo Bitch
               End If
              '// Wallpaper Estilo Mosaiko
               If OpcionesMusic.Mosaico = True Then
                 PoneRWallPapeR "Mosaico"
                 GoTo Bitch
               End If
         Else
            '// si no es proporcional
            SavePicture frmMain.picWallOriginal.Image, DirectoriOWindowS & "MusicMp3.bmp"
               If OpcionesMusic.Centrar = True Then
                 PoneRWallPapeR "Centro"
                 GoTo Bitch
               End If
               If OpcionesMusic.Mosaico = True Then
                 PoneRWallPapeR "Mosaico"
                 GoTo Bitch
               End If
         End If
Exit Sub
Bitch:
    frmMain.picWallOriginal.Picture = Nothing
    frmMain.picWallProp.Picture = Nothing
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+



