VERSION 5.00
Begin VB.Form frmSpectrum 
   BorderStyle     =   0  'None
   ClientHeight    =   1980
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   132
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   272
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer_Resize 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4920
      Top             =   1890
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4935
      Top             =   1380
   End
   Begin VB.PictureBox picFront 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   3720
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   1
      Top             =   3180
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox picSpectrum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   0
      ScaleHeight     =   117
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   266
      TabIndex        =   0
      Top             =   0
      Width           =   3990
      Begin VB.Label Label 
         BackColor       =   &H00000000&
         Caption         =   "Visualizacion"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   4245
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuPrevVis 
         Caption         =   "Previous Visualization"
      End
      Begin VB.Menu mnuNextVis 
         Caption         =   "Next Visualization"
      End
      Begin VB.Menu mnuConfigVis 
         Caption         =   "Configure Visualization"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSpectrum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'// Variables para la minimascara
Dim InFormDrag            As Boolean
Dim StartDragX            As Single
Dim StartDragY            As Single
Dim rWorkArea             As RECT
Dim mAttachedToRight      As Boolean
Dim mAttachedToLeft       As Boolean
Dim mAttachedToTop        As Boolean
Dim mAttachedToBottom     As Boolean
Dim mSnapDistance         As Long

' Used when resizing the window -
' X/Y distance of the mouse pointer from the form's edge
Dim XDistance As Long
Dim YDistance As Long

' Boolean flags - the current state of the form
Dim InXDrag As Boolean ' In horizontal resize
Dim InYDrag As Boolean ' In vertical resize

Dim NoRedraw As Boolean

 ' Current number of horizontal/vertical segments
Dim NumXSlices As Long
Dim NumYSlices As Long


Dim bResize As Boolean
Dim bLoadingVis As Boolean

Sub Load_Visualizacion(sFileVis As String)
 On Error Resume Next
 Dim s As String, i As Integer
 Dim bExistConfigScope As Boolean
 
 bLoadingVis = True
 
 sFileVis = tAppConfig.AppConfig & "Settings\" & sFileVis & ".vis"
 
 If Dir(sFileVis) = "" Then Exit Sub
 
 i = 0
 ' OSCILLOSCOPE
 ReDim tConfigScope(-1)
 Do
    s = Read_INI("Oscilloscope_" & i, "Number", "", , , sFileVis)
    If s <> "" Then
      ReDim Preserve tConfigScope(i)
       tConfigScope(i).Align = Read_INI("Oscilloscope_" & i, "Align", 1, , , sFileVis)
       If tConfigScope(i).Align < 0 Or tConfigScope(i).Align > 2 Then tConfigScope(i).Align = 1
       tConfigScope(i).BackColorScope = Read_INI("Oscilloscope_" & i, "BackColorScope", RGB(0, 255, 0), , , sFileVis)
       tConfigScope(i).LinesScope = Read_INI("Oscilloscope_" & i, "LinesScope", 50, , , sFileVis)
       If tConfigScope(i).LinesScope < 6 Or tConfigScope(i).LinesScope > 200 Then tConfigScope(i).LinesScope = 50
       bExistConfigScope = True
    End If
    i = i + 1
  Loop While s <> ""
   
 ' SPECTRUM
 With tConfigVis
  .BackColor = Read_INI("Spectrum", "BackColor", RGB(0, 0, 0), , , sFileVis)
  .BackColorBar = Read_INI("Spectrum", "BackColorBar", RGB(255, 255, 255), , , sFileVis)
  .BackColorPeak = Read_INI("Spectrum", "BackColorPeak", RGB(255, 255, 255), , , sFileVis)
  .Bars = Read_INI("Spectrum", "Bars", 50, , , sFileVis)
  If .Bars < 6 Or .Bars > 200 Then .Bars = 50
  .DrawBars = CBool(Read_INI("Spectrum", "DrawBars", 1, , , sFileVis))
  .DrawPeaks = CBool(Read_INI("Spectrum", "DrawPeaks", 1, , , sFileVis))
  .DrawSource = Read_INI("Spectrum", "DrawSource", 1, , , sFileVis)
  .Exist = CBool(Read_INI("Spectrum", "Exist", 1, , , sFileVis))
  .Gradient = Read_INI("Spectrum", "Gradient", "No Hay.jpg", , , sFileVis)
  .GrandientIndex = Read_INI("Spectrum", "GradientIndex", 0, , , sFileVis)
  .ImageFile = Read_INI("Spectrum", "ImageFile", "[Cover Front]", , , sFileVis)
  .Mirrored = CBool(Read_INI("Spectrum", "Mirrored", 1, , , sFileVis))
  .PeakGravity = Read_INI("Spectrum", "PeakGravity", 2, , , sFileVis)
  If .PeakGravity < 0 Or .PeakGravity > 4 Then .PeakGravity = 3
  .PeakHeight = Read_INI("Spectrum", "PeakHeight", 1, , , sFileVis)
  If .PeakHeight < 0 Or .PeakHeight > 4 Then .PeakHeight = 2
  .ScaleUp = Read_INI("Spectrum", "ScaleUp", 0, , , sFileVis)
  .Spacio = Read_INI("Spectrum", "Space", 0, , , sFileVis)
  If .Spacio > 10 Then .Spacio = 10
  
  
  If .DrawBars = False And .DrawPeaks = False Then .DrawBars = True

  
  ReDim tConfigVis.arryPeaks(tConfigVis.Bars)
  ReDim tConfigVis.arryWaitPeak(tConfigVis.Bars)
  Setup_Visualizacion

  bLoadingVis = False
End With


End Sub

Private Sub Form_Load()
  Me.Left = (Screen.Width - Me.Width) / 2   '// centrar formulario
  Me.Top = (Screen.Height - Me.Height) / 2
  NumXSlices = Me.ScaleWidth
  NumYSlices = Me.ScaleHeight
  '/* Distancia para anklar a los bordes del escritorio
  mSnapDistance = 10 * Screen.TwipsPerPixelX
End Sub

Sub Setup_Visualizacion()
 On Error Resume Next
   picSpectrum.Picture = LoadPicture()
   picFront.Picture = LoadPicture()
     
 If tConfigVis.Exist = True Then
   picSpectrum.BackColor = tConfigVis.BackColorBar
     
   If tConfigVis.DrawBars = True Then
     '// gradient
      If tConfigVis.DrawSource = 0 Then
          picFront.Picture = LoadPicture(tAppConfig.AppConfig & "Settings\" & tConfigVis.Gradient)
          '// Image
       ElseIf tConfigVis.DrawSource = 1 Then
              If Dir(tConfigVis.ImageFile) <> "" Then
                 picFront.Picture = LoadPicture(tConfigVis.ImageFile)
              Else
                 If Trim(strRutaCaratula) <> "" Then
                    picFront.Picture = LoadPicture(strRutaCaratula)
                 Else '// si no tiene caratula el album mostrar el default logo
                    picFront.Picture = frmPopUp.picDefaultLogo.Picture
                 End If
              End If
           End If
    Else
      picSpectrum.BackColor = tConfigVis.BackColor
    End If
   picFront.AutoSize = True
   Form_Resize
 Else
   picSpectrum.BackColor = 0
 End If
  
  bolVisShow = True

End Sub

Public Sub Form_Resize()
 On Error Resume Next
 bResize = True
 Timer_Resize.Interval = 200
 Timer_Resize.Enabled = True
 
 picSpectrum.Width = Me.ScaleWidth
 Label.Width = Me.ScaleWidth
 picSpectrum.Height = Me.ScaleHeight
 
 If tConfigVis.Exist = False Then Exit Sub
 
 picSpectrum.Cls
 picSpectrum.PaintPicture picFront.Picture, 0, 0, picSpectrum.ScaleWidth, picSpectrum.ScaleHeight, 0, 0
 picSpectrum.Picture = picSpectrum.Image

 DoEvents

End Sub


Public Sub Stop_Visualizacion()
Dim X1 As Single, Y1 As Single
Dim X2 As Single, Y2 As Single
Dim i As Integer, iSleep As Integer, j As Integer
Dim iSpacio As Integer, iPeak As Single, RaiseBars As Single, RaiseBars2 As Single
Dim Max&

'On Error Resume Next
On Error GoTo HELL
   
If bLoadingVis = True Then Exit Sub

picSpectrum.Cls
  
'// SPECTRUM ANALYZER
If tConfigVis.Exist = True Then
  For i = 0 To tConfigVis.Bars
      X1 = i * (picSpectrum.ScaleWidth / tConfigVis.Bars)
      X2 = X1 + (picSpectrum.ScaleWidth / tConfigVis.Bars)
      '---------------------------------------------------------------------
      '// full window
      If tConfigVis.Mirrored = True Then
         Y1 = picSpectrum.ScaleHeight / 2
      Else
         Y1 = picSpectrum.ScaleHeight
      End If
            
      '---------------------------------------------------------------------
      '// Raise bars
      If tConfigVis.ScaleUp = 0 Then 'Normal
         RaiseBars = Y1
         RaiseBars2 = Y1
      ElseIf tConfigVis.ScaleUp = 1 Then
             RaiseBars = (picSpectrum.ScaleHeight / 3)
             RaiseBars2 = Y1 + (picSpectrum.ScaleHeight / 5)
          ElseIf tConfigVis.ScaleUp = 2 Then
                 RaiseBars = (picSpectrum.ScaleHeight / 6)
                 RaiseBars2 = (Y1 + (picSpectrum.ScaleHeight / 6) * 2)
              ElseIf tConfigVis.ScaleUp = 3 Then
                     RaiseBars = (picSpectrum.ScaleHeight / 10)
                     RaiseBars2 = (Y1 + (picSpectrum.ScaleHeight / 10) * 4)
                  End If
      '---------------------------------------------------------------------
      Max = (0 * RaiseBars)
                        
                        
      If Max >= Y1 And tConfigVis.DrawPeaks = True Then Max = Y1 - tConfigVis.PeakHeight
                     
     '====================================================================
     '// bars
     If tConfigVis.DrawBars = True Then

        Y2 = RaiseBars - Max
        picSpectrum.Line (X1, Y2)-(X2, 0), tConfigVis.BackColor, BF
        
        If tConfigVis.Spacio >= 0 Then
          picSpectrum.Line (X2, Y1)-(X2 + tConfigVis.Spacio, 0), tConfigVis.BackColor, BF
        End If
       
       '// espejo
        If tConfigVis.Mirrored = True Then
           Y2 = RaiseBars2 + Max
           picSpectrum.Line (X1, Y2)-(X2, Y1 * 2), tConfigVis.BackColor, BF
           
           If tConfigVis.Spacio >= 0 Then
             picSpectrum.Line (X2, Y1)-(X2 + tConfigVis.Spacio, Y1 * 2), tConfigVis.BackColor, BF
           End If
        Else
           picSpectrum.Line (X1, Y1)-(X2, Y1 * 2), tConfigVis.BackColor, BF
        End If
     End If
          
     If tConfigVis.Spacio >= 0 Then
        X2 = X2 - 1
        X1 = X1 + 1 + tConfigVis.Spacio
     End If
     
     '====================================================================
     '// Peaks
     
     If tConfigVis.DrawPeaks = True Then
       tConfigVis.arryPeaks(i) = 0
       iPeak = RaiseBars - tConfigVis.arryPeaks(i)
       picSpectrum.Line (X1, iPeak - 1)-(X2, iPeak - tConfigVis.PeakHeight), tConfigVis.BackColorPeak, BF
       '// peaks de espejo
       If tConfigVis.Mirrored = True Then
          iPeak = RaiseBars2 + tConfigVis.arryPeaks(i) - tConfigVis.PeakHeight
          picSpectrum.Line (X1, iPeak + 1)-(X2, iPeak + tConfigVis.PeakHeight), tConfigVis.BackColorPeak, BF
       End If
    End If

  Next i
End If

'================================================================================
'// OSCILLOSCOPE
For j = 0 To UBound(tConfigScope)
  For i = 0 To tConfigScope(j).LinesScope
     X1 = i * (picSpectrum.ScaleWidth / tConfigScope(j).LinesScope)
     X2 = X1 + (picSpectrum.ScaleWidth / tConfigScope(j).LinesScope)
     Y1 = picSpectrum.ScaleHeight
      
     '// full window
     If tConfigScope(j).Align = 1 Then Y1 = picSpectrum.ScaleHeight / 2
      
     '// top y bottom bars
     If tConfigScope(j).Align = 0 Or tConfigScope(j).Align = 2 Then Y1 = picSpectrum.ScaleHeight / 4
      
     Y2 = 0
    
     '// bottom align
     If tConfigScope(j).Align = 0 Then Y1 = Y1 * 3
        

     picSpectrum.Line (X1, Y1)-(X1 + ((X2 - X1) / 3), Y1 - Y2), tConfigScope(j).BackColorScope
     picSpectrum.Line (X1 + ((X2 - X1) / 3), Y1 - Y2)-(X1 + (((X2 - X1) / 3) * 2), Y1 + Y2), tConfigScope(j).BackColorScope
     picSpectrum.Line (X1 + (((X2 - X1) / 3) * 2), Y1 + Y2)-(X2, Y1), tConfigScope(j).BackColorScope
  Next i
Next j
 
Exit Sub
HELL:
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  bolVisShow = False
  Cancel = 1
End Sub

Private Sub mnuConfigVis_Click()
 On Error Resume Next
   frmOpciones.cboVisualizacion.ListIndex = IndexVisualization
   frmOpciones.Select_Option 8
   frmOpciones.Show
End Sub

Private Sub mnuExit_Click()
 Unload Me
End Sub

Private Sub mnuNextVis_Click()
  On Error Resume Next
  If frmOpciones.cboVisualizacion.ListCount <= 1 Then Exit Sub
  
  IndexVisualization = IndexVisualization + 1
  
  If IndexVisualization >= frmOpciones.cboVisualizacion.ListCount Then IndexVisualization = 0
  
  Load_Visualizacion frmOpciones.cboVisualizacion.List(IndexVisualization)
  Label.Visible = True
  Label.Caption = frmOpciones.cboVisualizacion.List(IndexVisualization)
  Timer.Interval = 2000
  Timer.Enabled = True
End Sub

Private Sub mnuPrevVis_Click()
  On Error Resume Next
  If frmOpciones.cboVisualizacion.ListCount <= 1 Then Exit Sub
  
  IndexVisualization = IndexVisualization - 1
  
  If IndexVisualization < 0 Then IndexVisualization = frmOpciones.cboVisualizacion.ListCount - 1
  
  Load_Visualizacion frmOpciones.cboVisualizacion.List(IndexVisualization)
  Label.Visible = True
  Label.Caption = frmOpciones.cboVisualizacion.List(IndexVisualization)
  Timer.Interval = 2000
  Timer.Enabled = True
End Sub

Private Sub picSpectrum_DblClick()
Unload Me
End Sub

Public Sub Update_Visualizacion(arryValues() As Single)
Dim X1 As Single, Y1 As Single
Dim X2 As Single, Y2 As Single
Dim i As Integer, iSleep As Integer, j As Integer
Dim iSpacio As Integer, iPeak As Single, RaiseBars As Single, RaiseBars2 As Single
Dim Max&

'On Error Resume Next
On Error GoTo HELL
If bLoadingVis = True Or bResize = True Then Exit Sub
   
picSpectrum.Cls
  
'// SPECTRUM ANALYZER
If tConfigVis.Exist = True Then
  For i = 0 To tConfigVis.Bars
      X1 = i * (picSpectrum.ScaleWidth / tConfigVis.Bars)
      X2 = X1 + (picSpectrum.ScaleWidth / tConfigVis.Bars)
      '---------------------------------------------------------------------
      '// full window
      If tConfigVis.Mirrored = True Then
         Y1 = picSpectrum.ScaleHeight / 2
      Else
         Y1 = picSpectrum.ScaleHeight
      End If
            
      '---------------------------------------------------------------------
      '// Raise bars
      If tConfigVis.ScaleUp = 0 Then 'Normal
         RaiseBars = Y1
         RaiseBars2 = Y1
      ElseIf tConfigVis.ScaleUp = 1 Then
             RaiseBars = (picSpectrum.ScaleHeight / 3)
             RaiseBars2 = Y1 + (picSpectrum.ScaleHeight / 5)
          ElseIf tConfigVis.ScaleUp = 2 Then
                 RaiseBars = (picSpectrum.ScaleHeight / 6)
                 RaiseBars2 = (Y1 + (picSpectrum.ScaleHeight / 6) * 2)
              ElseIf tConfigVis.ScaleUp = 3 Then
                     RaiseBars = (picSpectrum.ScaleHeight / 10)
                     RaiseBars2 = (Y1 + (picSpectrum.ScaleHeight / 10) * 4)
                  End If
      '---------------------------------------------------------------------
      Max = (Format(arryValues(i), ".00") * RaiseBars)
                        
      'Max = Max * (tConfigVis.ScaleUp+1)
                        
      If Max >= Y1 And tConfigVis.DrawPeaks = True Then Max = Y1 - tConfigVis.PeakHeight
                     
     '====================================================================
     '// bars
     If tConfigVis.DrawBars = True Then

        Y2 = RaiseBars - Max
        picSpectrum.Line (X1, Y2)-(X2, 0), tConfigVis.BackColor, BF
        
        If tConfigVis.Spacio >= 0 Then
          picSpectrum.Line (X2, Y1)-(X2 + tConfigVis.Spacio, 0), tConfigVis.BackColor, BF
        End If
       
       '// espejo
        If tConfigVis.Mirrored = True Then
           Y2 = RaiseBars2 + Max
           picSpectrum.Line (X1, Y2)-(X2, Y1 * 2), tConfigVis.BackColor, BF
           
           If tConfigVis.Spacio >= 0 Then
             picSpectrum.Line (X2, Y1)-(X2 + tConfigVis.Spacio, Y1 * 2), tConfigVis.BackColor, BF
           End If
        Else
           picSpectrum.Line (X1, Y1)-(X2, Y1 * 2), tConfigVis.BackColor, BF
        End If
     End If
          
     If tConfigVis.Spacio >= 0 Then
        X2 = X2 - 1
        X1 = X1 + 1 + tConfigVis.Spacio
     End If
     
     '====================================================================
     '// Peaks
     
     If tConfigVis.DrawPeaks = True Then
       If tConfigVis.arryPeaks(i) < Max Then
          tConfigVis.arryPeaks(i) = Max
          tConfigVis.arryWaitPeak(i) = Time
       End If

       If tConfigVis.arryPeaks(i) < 0 Then tConfigVis.arryPeaks(i) = 0
        
       iPeak = RaiseBars - tConfigVis.arryPeaks(i)
     
       If iPeak <= tConfigVis.PeakHeight Then iPeak = tConfigVis.PeakHeight
            
       picSpectrum.Line (X1, iPeak - 1)-(X2, iPeak - tConfigVis.PeakHeight), tConfigVis.BackColorPeak, BF
     
       '// peaks de espejo
       If tConfigVis.Mirrored = True Then
          iPeak = RaiseBars2 + tConfigVis.arryPeaks(i) - tConfigVis.PeakHeight
          If iPeak >= picSpectrum.ScaleHeight Then iPeak = picSpectrum.ScaleHeight - tConfigVis.PeakHeight - 1
          
          picSpectrum.Line (X1, iPeak + 1)-(X2, iPeak + tConfigVis.PeakHeight), tConfigVis.BackColorPeak, BF
       End If
       
         If tConfigVis.arryWaitPeak(i) <> "" Then iSleep = DateDiff("s", tConfigVis.arryWaitPeak(i), Time)
         If (iSleep >= 1) Then tConfigVis.arryPeaks(i) = tConfigVis.arryPeaks(i) - tConfigVis.PeakGravity
     End If

  Next i
End If

'================================================================================
'// OSCILLOSCOPE
For j = 0 To UBound(tConfigScope)
  For i = 0 To tConfigScope(j).LinesScope
     X1 = i * (picSpectrum.ScaleWidth / tConfigScope(j).LinesScope)
     X2 = X1 + (picSpectrum.ScaleWidth / tConfigScope(j).LinesScope)
     Y1 = picSpectrum.ScaleHeight
      
     '// full window
     If tConfigScope(j).Align = 1 Then Y1 = picSpectrum.ScaleHeight / 2
      
     '// top y bottom bars
     If tConfigScope(j).Align = 0 Or tConfigScope(j).Align = 2 Then Y1 = picSpectrum.ScaleHeight / 4
      
     Y2 = (Format(arryValues(i), ".00") * Y1)
    
     '// bottom align
     If tConfigScope(j).Align = 0 Then Y1 = Y1 * 3
        

     picSpectrum.Line (X1, Y1)-(X1 + ((X2 - X1) / 3), Y1 - Y2), tConfigScope(j).BackColorScope
     picSpectrum.Line (X1 + ((X2 - X1) / 3), Y1 - Y2)-(X1 + (((X2 - X1) / 3) * 2), Y1 + Y2), tConfigScope(j).BackColorScope
     picSpectrum.Line (X1 + (((X2 - X1) / 3) * 2), Y1 + Y2)-(X2, Y1), tConfigScope(j).BackColorScope
  Next i

Next j

Exit Sub
HELL:

End Sub

Private Sub Timer_Resize_Timer()
  bResize = False
  Timer_Resize.Enabled = False
End Sub

Private Sub Timer_Timer()
  Label.Visible = False
  Timer.Enabled = False
End Sub

Private Sub picSpectrum_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
 On Error Resume Next
    If Button = vbRightButton Then PopupMenu Me.mnuMain

    If Button = vbLeftButton Then
    
        YDistance = Y - Me.ScaleHeight
        XDistance = x - Me.ScaleWidth
        
        ' If the mouse pointer is on the the bottom edge,
        ' flag Y (vertical) drag
        If Abs(YDistance) < 5 Then InYDrag = True
        
        ' If the mouse pointer is on the the right edge,
        ' flag X drag. Don't start drag if wer'e in the window
        ' title area
        If Abs(XDistance) < 5 And Y > 5 Then InXDrag = True
                   
        If InXDrag = False And InYDrag = False Then
           FormDrag_Down x * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
        End If
    
    End If

End Sub

Private Sub picSpectrum_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    InXDrag = False
    InYDrag = False
    InFormDrag = False
End Sub


Private Sub picSpectrum_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim NewYSlices As Single
Dim NewXSlices As Single

Dim ShowXResizeCursor As Boolean
Dim ShowYResizeCursor As Boolean
Dim ResizingNeeded As Boolean

   
  If InFormDrag = True Then
     FormDrag_Move x * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
     Exit Sub
  End If
 
   
   
    ' Determine what kind of cursor should be shown
    
    If Abs(Y - Me.ScaleHeight) < 5 Or InYDrag Then
        ShowYResizeCursor = True
    End If
    
    If (Abs(x - Me.ScaleWidth) < 5) Or InXDrag Then
        
        ShowXResizeCursor = True
    End If
    
    If ShowXResizeCursor And ShowYResizeCursor Then
        Me.MousePointer = vbSizeNWSE
        
    ElseIf ShowXResizeCursor Then
        Me.MousePointer = vbSizeWE
    
    ElseIf ShowYResizeCursor Then
        Me.MousePointer = vbSizeNS
    
    Else
        Me.MousePointer = vbDefault
    End If
    
    If InXDrag Then
        ' Compute new number of horizontal segments
        NewXSlices = (x - XDistance)
        If NewXSlices < 150 Then NewXSlices = 150
        
        ' Check if we should actually do the resize. Not every
        ' slightest mouse drag should cause a resize
        If Abs(NewXSlices - NumXSlices) > 10 Then
            NumXSlices = NewXSlices
            ResizingNeeded = True
        End If
    End If

    ' Same handling for vertical resize-drag
    If InYDrag Then
        
        NewYSlices = (Y - YDistance)
        If NewYSlices < 100 Then NewYSlices = 100
        
        If Abs(NewYSlices - NumYSlices) > 10 Then
            NumYSlices = NewYSlices
            ResizingNeeded = True
        End If
    End If

    If ResizingNeeded Then
       ' Compute width/height of form accodring to the number of
       ' x/y slices
       Me.Width = (NumXSlices * Screen.TwipsPerPixelX)
       Me.Height = (NumYSlices * Screen.TwipsPerPixelY)
  
    End If

End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub FormDrag_Move(x As Single, Y As Single)
 On Error Resume Next
  Dim DiffX As Long, DiffY As Long
  Dim NewX As Long, NewY As Long
  Dim ToLeftDistance As Long
  Dim ToRightDistance As Long
  Dim ToTopDistance As Long
  Dim ToBottomDistance As Long

 '// si estamos arrastrando
 If InFormDrag = True Then
    '// resta para mantener la posicion
    '// del cursor en la posicion inicial del objeto
    DiffX = x - StartDragX
    DiffY = Y - StartDragY
  
   If DiffX = 0 And DiffY = 0 Then Exit Sub
     '// obtener las coordenadas corectas
     NewX = Me.Left + DiffX
     NewY = Me.Top + DiffY

    '// Enkontrar los bordes del escritorio
    
    
    ToRightDistance = rWorkArea.Right - (NewX + Me.Width)
    ToLeftDistance = NewX - rWorkArea.Left
    ToBottomDistance = rWorkArea.Bottom - (NewY + Me.Height)
    ToTopDistance = NewY - rWorkArea.Top
    
    '// si no esta anklado
    If Not mAttachedToBottom Then
        '// si esta en el area minima para arrastrarse para abajo
        If Abs(ToBottomDistance) <= mSnapDistance Then
            '// anklar al borde de abajo
            NewY = rWorkArea.Bottom - Me.Height
            mAttachedToBottom = True
        End If
    Else
        
        If Abs(ToBottomDistance) > mSnapDistance Then
            '// Romper el anklado
            mAttachedToBottom = False
        Else
            '// mantener la actual posicion
            NewY = Me.Top
        End If
    End If

    If Not mAttachedToTop Then
        If Abs(ToTopDistance) <= mSnapDistance Then
            NewY = rWorkArea.Top
            mAttachedToTop = True
        End If
    Else
        If Abs(ToTopDistance) > mSnapDistance Then
            mAttachedToTop = False
        Else
            NewY = Me.Top
        End If
    End If

    If Not mAttachedToRight Then
        If Abs(ToRightDistance) <= mSnapDistance Then
            NewX = rWorkArea.Right - Me.Width
            mAttachedToRight = True
        End If
    Else
        If Abs(ToRightDistance) > mSnapDistance Then
            mAttachedToRight = False
        Else
            NewX = Me.Left
        End If
    End If

    If Not mAttachedToLeft Then
        If Abs(ToLeftDistance) <= mSnapDistance Then
            NewX = rWorkArea.Left
            mAttachedToLeft = True
        End If
    Else
        If Abs(ToLeftDistance) > mSnapDistance Then
            mAttachedToLeft = False
        Else
            NewX = Me.Left
        End If
    End If
   
   '// mover a la actual posicion
   Me.Move NewX, NewY
  
End If

End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub FormDrag_Down(x As Single, Y As Single)
 On Error Resume Next
    '// Obtener el Area de trabajo en rWorkArea
    '// del escritorio sin kontar la taskbar
    
    SystemGetWorkArea SPI_GETWORKAREA, 0&, rWorkArea, 0&
    
    '// Convretirlos de pixeles a twips
    rWorkArea.Top = rWorkArea.Top * Screen.TwipsPerPixelY
    rWorkArea.Left = rWorkArea.Left * Screen.TwipsPerPixelX
    rWorkArea.Bottom = rWorkArea.Bottom * Screen.TwipsPerPixelY
    rWorkArea.Right = rWorkArea.Right * Screen.TwipsPerPixelX
    
    '// almacenar las coordenadas iniciales
    StartDragX = x
    StartDragY = Y
    '// variable para empezar a arrastrar
    InFormDrag = True
     
End Sub
