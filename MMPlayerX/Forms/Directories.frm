VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmDirectorios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Albums Browser"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   Icon            =   "Directories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView TreeAlbums 
      Height          =   1650
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   2910
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList2"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2940
      Top             =   3705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directories.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directories.frx":220A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directories.frx":445D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directories.frx":66CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directories.frx":892F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.FileListBox FileSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDDEB5&
      Height          =   615
      Hidden          =   -1  'True
      Left            =   2745
      Pattern         =   "*.mp3;*.wav;*.wma"
      System          =   -1  'True
      TabIndex        =   0
      Top             =   2670
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Menu mnuPaths 
      Caption         =   "Paths"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuExpArchivos 
         Caption         =   "Explorar Archivos"
      End
      Begin VB.Menu mnuEditTags 
         Caption         =   "Editar Tags"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Reproducir"
      End
   End
End
Attribute VB_Name = "frmDirectorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Load()
On Error Resume Next

  bolDirectoriosShow = True
  
  Me.Icon = frmMain.Icon
  
  Load_Language_Directorios
  
  '// Centrar form
  frmDirectorios.Left = (Screen.Width - frmDirectorios.Width) / 2
  frmDirectorios.Top = (Screen.Height - frmDirectorios.Height) / 2

  Load_Albums
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Load_Albums()
 Dim iAlbum As Integer, iTrack As Integer
 Dim sKey As String, s As String, sLastNode As String
 Dim sPath() As String, i As Integer
 On Error Resume Next

FileSearch.Pattern = frmMain.ListRepRef.Pattern
TreeAlbums.Nodes.Clear

 '// add albums folders
 For iAlbum = 1 To TotalAlbumS
    s = frmMain.btnAlbum(iAlbum).ToolTipText
    If Right(Trim(frmMain.btnAlbum(iAlbum).ToolTipText), 1) = "\" Then s = Left(Trim(frmMain.btnAlbum(iAlbum).ToolTipText), Len(Trim(frmMain.btnAlbum(iAlbum).ToolTipText)) - 1)
     
    sPath = Split(s, "\", , vbTextCompare)
    
    If TreeAlbums.Nodes.count = 0 Or sLastNode <> sPath(0) Then
       TreeAlbums.Nodes.Add , , CStr(sPath(0) & "\"), sPath(0), 1
       sLastNode = sPath(0)
    End If
    
    sKey = sPath(0) & "\"
    
    For i = 1 To UBound(sPath)
       'If TreeAlbums.Nodes(sKey).Children = 0 Then
          TreeAlbums.Nodes.Add sKey, tvwChild, sKey & sPath(i) & "\", sPath(i), 2
        'End If
      sKey = sKey & sPath(i) & "\"
    Next i
                
   '// add files in album
     FileSearch.Path = sKey
     For iTrack = 0 To FileSearch.ListCount - 1
        If LCase(Right(FileSearch.List(iTrack), 3)) = "mp3" Then
            TreeAlbums.Nodes.Add sKey, tvwChild, CStr(iAlbum & " \ " & iTrack), FileSearch.List(iTrack), 3
        ElseIf LCase(Right(FileSearch.List(iTrack), 3)) = "wma" Then
                TreeAlbums.Nodes.Add sKey, tvwChild, CStr(iAlbum & " \ " & iTrack), FileSearch.List(iTrack), 4
            Else
                TreeAlbums.Nodes.Add sKey, tvwChild, CStr(iAlbum & " \ " & iTrack), FileSearch.List(iTrack), 5
            End If
'            TreeAlbums.Nodes(CStr(iAlbum & " \ " & iTrack)).ForeColor = &H0&
'
'            If iTrack Mod 2 <> 0 Then
'               TreeAlbums.Nodes(CStr(iAlbum & " \ " & iTrack)).BackColor = &HE0E0E0
'            Else
'               TreeAlbums.Nodes(CStr(iAlbum & " \ " & iTrack)).BackColor = &HFDDEB5
'            End If
     Next iTrack
     sKey = ""
 Next iAlbum
  '// seleccionar el album reproduciendo
  TreeAlbums.Nodes(CStr(intActiveAlbum & " \ " & frmMain.ListRep.ListIndex)).Selected = True
  Exit Sub
Hell:
  MsgBox err.Description
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Resize()
  TreeAlbums.Width = Me.ScaleWidth
  TreeAlbums.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
bolDirectoriosShow = False
Me.Hide
Cancel = 1
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


Private Sub mnuEditTags_Click()
  Dim arryTemp() As String
  
  On Error Resume Next
  
  If TreeAlbums.SelectedItem.Children = 0 Then
  arryTemp = Split(TreeAlbums.SelectedItem.Key, "\")
  
  If UBound(arryTemp) = 1 Then '// is a file click
     If bolTagsShow = True Then
       frmTags.FileTags.Path = frmMain.btnAlbum(CInt(arryTemp(0))).ToolTipText
       frmTags.Make_List_Ref
       If Trim(arryTemp(1)) <> "" Then frmTags.FileTags.Selected(CInt(arryTemp(1))) = True
       frmTags.ZOrder 0
     Else
       frmTags.Show
       frmTags.FileTags.Path = frmMain.btnAlbum(CInt(arryTemp(0))).ToolTipText
       frmTags.Make_List_Ref
       If Trim(arryTemp(1)) <> "" Then frmTags.FileTags.Selected(CInt(arryTemp(1))) = True
     End If
  End If
  
  End If
 
 
End Sub

Private Sub mnuExpArchivos_Click()
  Dim arryTemp() As String
  Dim x As Long
  
  If TreeAlbums.Nodes.count = 0 Then Exit Sub
   
  On Error Resume Next
  
    '// click in node archive
  If TreeAlbums.SelectedItem.Children = 0 Then
    arryTemp = Split(TreeAlbums.SelectedItem.Key, "\")
  
    If UBound(arryTemp) = 1 Then
       x = Shell("explorer.exe " & frmMain.btnAlbum(CInt(arryTemp(0))).ToolTipText, vbMaximizedFocus)
    End If
  End If
  

End Sub

Private Sub mnuPlay_Click()
  Dim arryTemp() As String
  
  On Error Resume Next
  
  If TreeAlbums.Nodes.count = 0 Then Exit Sub

  '// click in node archive
  If TreeAlbums.SelectedItem.Children = 0 Then
    arryTemp = Split(TreeAlbums.SelectedItem.Key, "\")
  
    If UBound(arryTemp) = 1 Then '// is a file click
      If intActiveAlbum = CInt(arryTemp(0)) Then
        If Trim(arryTemp(1)) <> "" Then frmMain.ListRep.Selected(CInt(arryTemp(1))) = True
      Else
        frmMain.Play_Album CInt(arryTemp(0))
        If Trim(arryTemp(1)) <> "" Then frmMain.ListRep.Selected(CInt(arryTemp(1))) = True
      End If
    End If
  End If

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub TreeAlbums_DblClick()
  Dim arryTemp() As String
  
  On Error Resume Next
  
  If TreeAlbums.Nodes.count = 0 Then Exit Sub

  '// click in node archive
  If TreeAlbums.SelectedItem.Children = 0 Then
    arryTemp = Split(TreeAlbums.SelectedItem.Key, "\")
  
    If UBound(arryTemp) = 1 Then '// is a file click
      If intActiveAlbum = CInt(arryTemp(0)) Then
        If Trim(arryTemp(1)) <> "" Then frmMain.ListRep.Selected(CInt(arryTemp(1))) = True
      Else
        frmMain.Play_Album CInt(arryTemp(0))
        If Trim(arryTemp(1)) <> "" Then frmMain.ListRep.Selected(CInt(arryTemp(1))) = True
      End If
    End If
  End If
End Sub

Private Sub TreeAlbums_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  
  If TreeAlbums.Nodes.count = 0 Then Exit Sub
  
  '// click in node archive
    If TreeAlbums.SelectedItem.Children = 0 Then
      frmDirectorios.mnuEditTags.Enabled = True
      frmDirectorios.mnuExpArchivos.Enabled = True
      frmDirectorios.mnuPlay.Enabled = True
    Else
      frmDirectorios.mnuEditTags.Enabled = False
      frmDirectorios.mnuExpArchivos.Enabled = False
      frmDirectorios.mnuPlay.Enabled = False
    End If
  
 If Button = vbRightButton Then PopupMenu Me.mnuPaths
End Sub


