VERSION 5.00
Begin VB.UserControl ScrollText 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2730
   ScaleHeight     =   115
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   182
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1065
      Top             =   1740
   End
   Begin VB.PictureBox picCaptionText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   90
      Left            =   0
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   3
      Top             =   0
      Width           =   2355
   End
   Begin VB.PictureBox picTextScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   360
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   2
      Top             =   1155
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.PictureBox picDefault 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   375
      Picture         =   "ScrollText.ctx":0000
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   1
      Top             =   1395
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   420
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2325
   End
End
Attribute VB_Name = "ScrollText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private iSpeedScroll As Integer
Private iorgXScroll As Integer, iDesHeight As Integer, iDesWidth As Integer
Private bZigZagScroll As Boolean
Private stcIzq As Boolean
Private sScrollText As String
Private bScrolling As Boolean
Private bScroll As Boolean
Private bAutosize As Boolean

Public Event Click()
Public Event DBLClick()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

Private eAlignText As AlignmentConstants

Public Enum peScrollType
  Rolling = 0
  ZigZag = 1
End Enum

Private peST As peScrollType

'//--------------------------------------------------------------------------
Public Property Get hwnd() As Variant
    hwnd = UserControl.hwnd
End Property


'//--------------------------------------------------------------------------
Public Property Get AutoSize() As Boolean
    AutoSize = bAutosize
End Property

Public Property Let AutoSize(ByVal bValue As Boolean)
    bAutosize = bValue
    If bAutosize = True Then
       UserControl.Width = picTextScroll.ScaleWidth * 15
    End If
    PropertyChanged "AutoSize"
End Property

'//--------------------------------------------------------------------------
Public Property Get Scroll() As Boolean
    Scroll = bScroll
End Property

Public Property Let Scroll(ByVal bValue As Boolean)
    bScroll = bValue
    Timer1.Enabled = False
    If bScroll = True Then Call ScrollText
    PropertyChanged "Scroll"
End Property

'//--------------------------------------------------------------------------
Public Property Get PictureText() As Picture
    Set PictureText = picText.Picture
End Property

Public Property Set PictureText(ByVal New_Picture As Picture)
    Set picText.Picture = New_Picture
    picText.AutoSize = True
    picCaptionText.Height = picText.ScaleHeight / 3
    picCaptionText.Width = UserControl.ScaleWidth
    UserControl.Height = (picText.ScaleHeight / 3) * 15

    BuildText sScrollText
    PropertyChanged "PictureText"
End Property

'//--------------------------------------------------------------------------
Public Property Get AlignText() As AlignmentConstants
  AlignText = eAlignText
End Property

Public Property Let AlignText(ByVal vNewValue As AlignmentConstants)
  eAlignText = vNewValue
  UpdateAlign
  PropertyChanged "AlignText"
End Property

'//--------------------------------------------------------------------------
Public Property Get CaptionText() As String
  CaptionText = sScrollText
End Property

Public Property Let CaptionText(ByVal vNewValue As String)
  sScrollText = vNewValue
  If sScrollText = "" Then sScrollText = " "
  BuildText sScrollText
  If bScroll = True Then Call ScrollText
  PropertyChanged "CaptionText"
End Property

'//--------------------------------------------------------------------------
Public Property Get ScrollType() As peScrollType
  ScrollType = peST
End Property

Public Property Let ScrollType(ByVal vNewValue As peScrollType)
  peST = vNewValue
  bZigZagScroll = peST
  If bScrolling = True Then Call ScrollText
  PropertyChanged "ScrollType"
End Property

'//--------------------------------------------------------------------------
Public Property Get ScrollVelocity() As Integer
  ScrollVelocity = Timer1.Interval
End Property

Public Property Let ScrollVelocity(ByVal vNewValue As Integer)
  Timer1.Interval = vNewValue
  PropertyChanged "ScrollVelocity"
End Property

'//--------------------------------------------------------------------------
Public Property Get ScrollingNow() As Boolean
  ScrollingNow = Timer1.Enabled
End Property

'//--------------------------------------------------------------------------
Public Sub StopScroll(ByVal bValue As Boolean)
  If bScrolling = True Then
     Timer1.Enabled = Not bValue
  End If
End Sub

'//--------------------------------------------------------------------------
Public Property Get BackColor() As OLE_COLOR
  BackColor = picCaptionText.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
  picCaptionText.BackColor = vNewValue
  picTextScroll.BackColor = picCaptionText.BackColor
  BuildText sScrollText
  PropertyChanged "BackColor"
End Property




Private Sub picCaptionText_Click()
 RaiseEvent Click
End Sub

Private Sub picCaptionText_DblClick()
 RaiseEvent DBLClick
End Sub

Private Sub picCaptionText_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
 RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub picCaptionText_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub picCaptionText_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub Timer1_Timer()
 Dim i As Integer
 Static stcWait As Boolean
 Static stcPause As Integer
  
 If bZigZagScroll = False Then
   
   For i = 0 To iSpeedScroll
      '// recorrer la imagen 1 pexel izquierda
    BitBlt picCaptionText.hdc, 0, 0, iDesWidth, iDesHeight, picCaptionText.hdc, 1, 0, &HCC0020
    '// copiar el pixel recorrido a la derecha de la imagen
    BitBlt picCaptionText.hdc, iDesWidth, 0, 1, iDesHeight, picTextScroll.hdc, iorgXScroll, 0, &HCC0020

    iorgXScroll = iorgXScroll + 1
    If iorgXScroll = picTextScroll.ScaleWidth Then iorgXScroll = 0
  Next i
  
  picCaptionText.Refresh
  Exit Sub
End If
  
  If stcWait = True And stcPause < 30 Then
    stcPause = stcPause + 1
    Exit Sub
  End If
  
  For i = 0 To iSpeedScroll
    
  If stcIzq = False Then
    '// recorrer la imagen 1 pexel izquierda
    BitBlt picCaptionText.hdc, 0, 0, iDesWidth, iDesHeight, picCaptionText.hdc, 1, 0, &HCC0020
    '// copiar el pixel recorrido a la derecha de la imagen
    BitBlt picCaptionText.hdc, iDesWidth, 0, 1, iDesHeight, picTextScroll.hdc, iorgXScroll, 0, &HCC0020
    
    iorgXScroll = iorgXScroll + 1
    If iorgXScroll > picTextScroll.ScaleWidth Then
      stcIzq = True
      iorgXScroll = Abs(picTextScroll.ScaleWidth - picCaptionText.ScaleWidth)
      stcWait = True
      stcPause = 0
    End If
  Else
    BitBlt picCaptionText.hdc, 1, 0, iDesWidth, iDesHeight, picCaptionText.hdc, 0, 0, &HCC0020
    BitBlt picCaptionText.hdc, 0, 0, 1, iDesHeight, picTextScroll.hdc, iorgXScroll, 0, &HCC0020
    
    iorgXScroll = iorgXScroll - 1
    If iorgXScroll < 0 Then
       stcIzq = False
       iorgXScroll = picCaptionText.ScaleWidth
       stcWait = True
       stcPause = 0
    End If
  End If
        
  Next i
  
    picCaptionText.Refresh
End Sub



Private Sub UserControl_Initialize()
 picText.Picture = picDefault.Picture
 sScrollText = "Scroll Text Control"
 BuildText sScrollText
End Sub

Private Sub UserControl_Resize()
  picCaptionText.Width = UserControl.ScaleWidth
  UserControl.Height = (picText.ScaleHeight / 3) * 15
  If bScrolling = True Then
    Call ScrollText
  Else
  'picCaptionText.Height = UserControl.ScaleHeight
  UpdateAlign
  End If
End Sub

'//--------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("PictureText", picText.Picture, picDefault.Picture)
    Call PropBag.WriteProperty("BackColor", picCaptionText.BackColor, &H0)
    Call PropBag.WriteProperty("CaptionText", sScrollText, "Scroll Text")
    Call PropBag.WriteProperty("AlignText", eAlignText, 0)
    Call PropBag.WriteProperty("ScrollType", peST, 0)
    Call PropBag.WriteProperty("ScrollVelocity", Timer1.Interval, 200)
    Call PropBag.WriteProperty("Scroll", bScroll, False)
    Call PropBag.WriteProperty("AutoSize", bAutosize, False)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 On Error Resume Next
    picText.Picture = PropBag.ReadProperty("PictureText", picDefault.Picture)
    If picText.Picture = 0 Then picText.Picture = picDefault.Picture
    picCaptionText.BackColor = PropBag.ReadProperty("BackColor", &H0)
    picTextScroll.BackColor = picCaptionText.BackColor
    eAlignText = PropBag.ReadProperty("AlignText", 0)
    Timer1.Interval = PropBag.ReadProperty("ScrollVelocity", 200)
    sScrollText = PropBag.ReadProperty("CaptionText", "Scroll Text")
    bScroll = PropBag.ReadProperty("Scroll", False)
    bAutosize = PropBag.ReadProperty("AutoSize", False)
    peST = PropBag.ReadProperty("ScrollType", 0)
    bZigZagScroll = peST
    BuildText sScrollText
    'If bScroll = True Then Call ScrollText
End Sub

Private Sub UpdateAlign()
  picCaptionText.Picture = LoadPicture()
  Select Case eAlignText
      Case 0 '// left
        BitBlt picCaptionText.hdc, 0, 0, picTextScroll.ScaleWidth, picTextScroll.ScaleHeight, picTextScroll.hdc, 0, 0, &HCC0020
      Case 2 '// center
        BitBlt picCaptionText.hdc, (picCaptionText.ScaleWidth / 2) - (picTextScroll.ScaleWidth / 2), 0, picTextScroll.ScaleWidth, picTextScroll.ScaleHeight, picTextScroll.hdc, 0, 0, &HCC0020
      Case 1 '//right
        BitBlt picCaptionText.hdc, (picCaptionText.ScaleWidth) - (picTextScroll.ScaleWidth), 0, picTextScroll.ScaleWidth, picTextScroll.ScaleHeight, picTextScroll.hdc, 0, 0, &HCC0020
      Case Else
        BitBlt picCaptionText.hdc, 0, 0, picTextScroll.ScaleWidth, picTextScroll.ScaleHeight, picTextScroll.hdc, 0, 0, &HCC0020
   End Select
End Sub

Private Sub ScrollText()
  Timer1.Enabled = False
 
 If picTextScroll.ScaleWidth > picCaptionText.ScaleWidth Then
   
   '// cool effect
   If peST = Rolling Then BuildText "* " & sScrollText & " *"
   
   picCaptionText.Picture = LoadPicture()
   picCaptionText.Picture = picTextScroll.Picture
   iDesHeight = picCaptionText.ScaleHeight
   iDesWidth = picCaptionText.ScaleWidth - 1
   iorgXScroll = picCaptionText.ScaleWidth
   iSpeedScroll = 1
   stcIzq = False
   bScrolling = True
   Timer1.Enabled = True
 Else
   bScrolling = False
 End If

End Sub

Private Sub BuildText(sText As String)
 Dim i As Integer, iCell As Integer, iCellX As Integer
 Dim s As String
 picTextScroll.Picture = LoadPicture()
 picTextScroll.Width = (Len(sText)) * (picText.ScaleWidth / 31)
 picTextScroll.Height = (picText.ScaleHeight / 3) * 15
 picCaptionText.Height = picTextScroll.Height

 If bAutosize = True Then UserControl.Width = picTextScroll.ScaleWidth * 15
 
 For i = 1 To Len(sText)
   s = Mid(sText, i, 1)
   iCell = IndexWord(UCase(s))
   iCellX = (picText.ScaleWidth / 31) * (i - 1)
   CopyCell iCell, iCellX
 Next i
 
  picTextScroll.Picture = picTextScroll.Image
  UpdateAlign
End Sub

Private Function IndexWord(sWord As String) As Integer
 Dim iWord As Integer
 Select Case sWord
     Case "A", "Á", "À": iWord = 0
     Case "B": iWord = 1
     Case "C": iWord = 2
     Case "D": iWord = 3
     Case "E", "É": iWord = 4
     Case "F": iWord = 5
     Case "G": iWord = 6
     Case "H": iWord = 7
     Case "I", "Í", "Ì": iWord = 8
     Case "J": iWord = 9
     Case "K": iWord = 10
     Case "L": iWord = 11
     Case "M": iWord = 12
     Case "N", "Ñ": iWord = 13
     Case "O", "Ó", "Ò": iWord = 14
     Case "P": iWord = 15
     Case "Q": iWord = 16
     Case "R": iWord = 17
     Case "S": iWord = 18
     Case "T": iWord = 19
     Case "U", "Ú", "Ù", "Ü", "Û": iWord = 20
     Case "V": iWord = 21
     Case "W": iWord = 22
     Case "X": iWord = 23
     Case "Y": iWord = 24
     Case "Z": iWord = 25
     Case """": iWord = 26
     Case "@", "®": iWord = 27
     Case " ": iWord = 29
     Case " ": iWord = 29
     Case " ": iWord = 29
     
     Case "0": iWord = 31
     Case "1": iWord = 32
     Case "2": iWord = 33
     Case "3": iWord = 34
     Case "4": iWord = 35
     Case "5": iWord = 36
     Case "6": iWord = 37
     Case "7": iWord = 38
     Case "8": iWord = 39
     Case "9": iWord = 40
     Case "_": iWord = 41
     Case ".": iWord = 42
     Case ":", ";": iWord = 43
     Case "(", "<": iWord = 44
     Case ")", ">": iWord = 45
     Case "-", "~", "°": iWord = 46
     Case "'", "`", "´": iWord = 47
     Case "!", "¡": iWord = 48
     Case "_": iWord = 49
     Case "+": iWord = 50
     Case "\", "|": iWord = 51
     Case "/": iWord = 52
     Case "[", "{": iWord = 53
     Case "]", "}": iWord = 54
     Case "^": iWord = 55
     Case "&": iWord = 56
     Case "%": iWord = 57
     Case ",": iWord = 58
     Case "=": iWord = 59
     Case "$": iWord = 60
     Case "#": iWord = 61
     
     Case "Ã": iWord = 62
     Case "ö", "õ", "ô": iWord = 63
     Case "Ä": iWord = 64
     Case "?", "¿": iWord = 65
     Case "*": iWord = 66
          
     Case Else
           iWord = 29

  End Select
 IndexWord = iWord
End Function

Private Sub CopyCell(iIndex As Integer, orgX As Integer)
  Dim iorgXScroll As Integer
  Dim srcY As Integer
  Dim srcWidth As Integer
  Dim srcHeight As Integer
    
    srcWidth = picText.ScaleWidth / 30
    srcHeight = picText.ScaleHeight / 3
   
    If iIndex <= 30 Then srcY = srcHeight * 0
    If iIndex > 30 And iIndex <= 61 Then srcY = srcHeight * 1: iIndex = iIndex - 31
    If iIndex > 61 Then srcY = srcHeight * 2: iIndex = iIndex - 62
    
    iorgXScroll = srcWidth * iIndex
    Call BitBlt(picTextScroll.hdc, orgX, 0, srcWidth, srcHeight, picText.hdc, iorgXScroll, srcY, vbSrcCopy)
End Sub


