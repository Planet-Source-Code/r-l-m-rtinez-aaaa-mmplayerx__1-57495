VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTags 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " MPEG File Info Box + ID3 Tag Editor"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Tags.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ListView listRef 
      Height          =   2550
      Left            =   180
      TabIndex        =   39
      Top             =   5310
      Visible         =   0   'False
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   4498
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "FILE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "TITLE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ARTIST"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ALBUM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "YEAR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "GENRE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "LYRICS"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Album >>"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1830
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   780
      UseMaskColor    =   -1  'True
      Width           =   1065
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<< Album"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   330
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   780
      UseMaskColor    =   -1  'True
      Width           =   1065
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   660
      TabIndex        =   3
      Top             =   4680
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   30
      TabIndex        =   36
      Top             =   -30
      Width           =   7890
      Begin ComctlLib.ProgressBar pbProgress 
         Height          =   270
         Left            =   60
         TabIndex        =   38
         Top             =   330
         Visible         =   0   'False
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblFile 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   60
         TabIndex        =   37
         Top             =   135
         Width           =   7785
      End
   End
   Begin VB.Frame Frame 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4425
      Index           =   2
      Left            =   30
      TabIndex        =   35
      Top             =   615
      Width           =   3270
      Begin VB.FileListBox FileTags 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3555
         Hidden          =   -1  'True
         Left            =   60
         MultiSelect     =   2  'Extended
         Pattern         =   "*.mp3;*.wma;*.wav"
         System          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6540
      TabIndex        =   24
      Top             =   4680
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5040
      TabIndex        =   23
      Top             =   4680
      Width           =   1305
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3570
      TabIndex        =   22
      Top             =   4680
      Width           =   1305
   End
   Begin VB.PictureBox pictab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3540
      Index           =   1
      Left            =   3390
      ScaleHeight     =   3540
      ScaleWidth      =   4440
      TabIndex        =   25
      Top             =   1035
      Width           =   4440
      Begin VB.Frame Frame 
         Caption         =   "MPEG Info"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Index           =   0
         Left            =   45
         TabIndex        =   32
         Top             =   1710
         Width           =   4380
         Begin VB.Label lblMPEGInfo 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1470
            Left            =   135
            TabIndex        =   33
            Top             =   225
            Width           =   4170
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "ID3v1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Index           =   1
         Left            =   45
         TabIndex        =   26
         Top             =   30
         Width           =   4380
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   2130
            TabIndex        =   12
            Top             =   1215
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   690
            TabIndex        =   10
            Top             =   1200
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   690
            TabIndex        =   8
            Top             =   900
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   705
            TabIndex        =   6
            Top             =   600
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.ComboBox cboGenre 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2415
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1185
            Width           =   1815
         End
         Begin VB.TextBox txtYear 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            MaxLength       =   4
            TabIndex        =   11
            Top             =   1185
            Width           =   540
         End
         Begin VB.TextBox txtAlbum 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            MaxLength       =   30
            TabIndex        =   9
            Top             =   885
            Width           =   3270
         End
         Begin VB.TextBox txtArtist 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            MaxLength       =   30
            TabIndex        =   7
            Top             =   570
            Width           =   3270
         End
         Begin VB.TextBox txtTitle 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            MaxLength       =   30
            TabIndex        =   5
            Top             =   240
            Width           =   3270
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Genre:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   1545
            TabIndex        =   31
            Top             =   1245
            Width           =   600
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Title:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   30
            Top             =   270
            Width           =   435
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Album:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   29
            Top             =   915
            Width           =   615
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Artist:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   28
            Top             =   615
            Width           =   525
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Year:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   27
            Top             =   1215
            Width           =   465
         End
      End
   End
   Begin VB.PictureBox pictab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3540
      Index           =   2
      Left            =   3390
      ScaleHeight     =   3540
      ScaleWidth      =   4440
      TabIndex        =   34
      Top             =   1035
      Width           =   4440
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Index           =   0
         Left            =   1170
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   15
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Index           =   1
         Left            =   1556
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   15
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Index           =   4
         Left            =   2715
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   15
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "||"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Index           =   2
         Left            =   1942
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   15
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Index           =   3
         Left            =   2328
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   15
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   690
         TabIndex        =   20
         Top             =   3195
         Width           =   1425
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "Deshacer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2385
         TabIndex        =   21
         Top             =   3195
         Width           =   1425
      End
      Begin VB.TextBox txtLyrics 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   19
         Top             =   255
         Width           =   4380
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   3975
      Left            =   3330
      TabIndex        =   4
      Top             =   675
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7011
      TabWidthStyle   =   2
      TabFixedWidth   =   3175
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Tags         "
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Lyrics        "
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
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
End
Attribute VB_Name = "frmTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FilesSelected As Integer
 

'// vars functions undo in lyrics
Private Arr() As Long
Private Const cChunk = 10
Private Last As Long, Cur As Long
Dim Pos As Long

Dim FilePlaying As String
Dim LastPosition As Long
Dim LastState As String
Dim iCurrentAlbum As Integer
Dim bRestartPlayer As Boolean
Dim sLastGenre As String

Private Sub cboGenre_Click()
 If sLastGenre = "" Then Exit Sub
 If sLastGenre <> cboGenre.Text Then
    Update_Tags_Ref
    If cmdApply.Enabled = False Then cmdApply.Enabled = True
 End If
 sLastGenre = ""
 
End Sub

Private Sub cboGenre_DropDown()
 sLastGenre = cboGenre.Text
End Sub

Private Sub chkTags_Click(Index As Integer)
  Dim bolEnabled As Boolean
  
  If chkTags(Index).Value = vbChecked Then
    bolEnabled = True
  End If
  
  Select Case Index
    Case 0 '// Artist
       txtArtist.Enabled = bolEnabled
    Case 1 '// Album
       txtAlbum.Enabled = bolEnabled
    Case 2 '// Year
       txtYear.Enabled = bolEnabled
    Case 3 '// genre
       cboGenre.Enabled = bolEnabled
  End Select
  
  If chkTags(0).Value = vbUnchecked And chkTags(1).Value = vbUnchecked And chkTags(2).Value = vbUnchecked _
    And chkTags(3).Value = vbUnchecked Then
      cmdApply.Enabled = False
  Else
    cmdApply.Enabled = True
  End If
End Sub

Private Sub cmdAdd_Click()
   'add a timestamp at the beginning of the current line (Lyrics)
   
   Dim OldMin As Long 'the minutes of old timestamp
   Dim OldSec As Long 'the seconds of old timestamp
   Dim oldHou As Long 'the hours of old timestamp
   Dim NewMin As Long 'the minutes of new timestamp
   Dim NewSec As Long 'the seconds of new timestamp
   Dim NewHou As Long 'the hours of new timestamp
   Dim LineLength As Long 'length of a line
   Dim CurrentLine As Long 'the current line number
   Dim TotalLines As Long 'how many lines there are
   Dim sCurrentTime As String 'the current time in string format
   Dim CharPos As Long 'character position
   
   Dim arryOldTime() As String
   Dim arryNewTime() As String
   Dim s As String, strTemp As String
   Dim j As Integer, fin As Integer
   
    'error handler
   On Error GoTo HELL
     '================================================================
     '  This is simple lyrics function
     '  how it work? good question :)
     '   - First load a file in tag editor
     '   - Write the lyrics in the text
     '   - Play the song with the over buttons
     '   - Use add button in the just time
     '                            is all, ¿Facil no?
     '================================================================
   
   If FileTags.ListCount = 0 Or PlayerState = "false" Then Exit Sub
   
   'check to make sure it contains a time
   sCurrentTime = Convert_Time(Stream_GetPosition(1))
   arryNewTime = Split(sCurrentTime, ":")
   
   'if has hours
   If UBound(arryNewTime) > 1 Then
     'convert the Time into integers
     NewHou = Val(arryNewTime(0))
     NewMin = Val(arryNewTime(1))
     NewSec = Val(arryNewTime(2))
   Else
     NewHou = 0
     NewMin = Val(arryNewTime(0))
     NewSec = Val(arryNewTime(1))
   End If
   
   'add the brackets to the time
   s = "[" & sCurrentTime & "]"
   
   'set the insert point to the beginning of the line, add 1 to it to make sure
   'we don't get a 0 length string compare.
   CurrentLine = SendMessage(txtLyrics.hwnd, EM_LINEFROMCHAR, txtLyrics.SelStart, ZERO)
   CharPos = SendMessage(txtLyrics.hwnd, EM_LINEINDEX, CurrentLine, ZERO)
   'get the length of the line
   LineLength = SendMessage(txtLyrics.hwnd, EM_LINELENGTH, CharPos, ZERO)
   LineLength = CharPos + LineLength
   Pos = CharPos + 1
   
   '// note: the [Do..Loop Until] is optional for look only
   '// you can delete and work lyrics function :P
   
   'check to make sure there is no timestamp already there, if so
   'then compare the new time to the old timestamp so the new one
   'is inserted at the correct point in end of old timestamp.

      'there is a timestamp here, get the time
       Do
         j = InStr(Pos, txtLyrics.Text, "[")
         If j > 0 And j <= LineLength Then
            fin = InStr(Pos, txtLyrics.Text, "]")
            '// solo agregar letras hasta el formato 00:00:00
            If ((fin - 1) - j) < 9 Then
              strTemp = Mid$(txtLyrics.Text, j + 1, fin - j - 1)
            End If
         Else
           Exit Do
         End If
         
         arryOldTime = Split(strTemp, ":")
                
            'if has hours
          If UBound(arryOldTime) > 1 Then
             'convert the Time into integers
            oldHou = Val(arryOldTime(0))
            OldMin = Val(arryOldTime(1))
            OldSec = Val(arryOldTime(2))
          Else
            oldHou = 0
            OldMin = Val(arryOldTime(0))
            OldSec = Val(arryOldTime(1))
          End If
      
          'check to see if new timestamp is newer that old timestamp
          If (NewHou > oldHou) Or (NewHou = oldHou And NewMin > OldMin) Or (NewHou = oldHou And NewMin = OldMin And NewSec > OldSec) Then
             'yes, it is, so skip this one
             Pos = fin + 1
          Else
             Exit Do
          End If
       Loop Until j = 0
   LineLength = 0
    
   'subtract one from the insert point and insert the stamp
   Pos = Pos - 1
   txtLyrics.SelStart = Pos
   txtLyrics.SelText = s
   'and push this position onto the undo stack
    Undo_Push Pos
   'enable the undo button
   cmdUndo.Enabled = True
   
   'now drop them to the next non blank line, or back to the beginning
   'how many lines?
   TotalLines = SendMessage(txtLyrics.hwnd, EM_GETLINECOUNT, ZERO, ZERO)
   'safety check... should always be true
   If TotalLines > CurrentLine Then
      Do
         'increment current line
         CurrentLine = CurrentLine + 1
         'Get the position of the beginning of the line
         CharPos = SendMessage(txtLyrics.hwnd, EM_LINEINDEX, CurrentLine, ZERO)
         'get the length of the line
         LineLength = SendMessage(txtLyrics.hwnd, EM_LINELENGTH, CharPos, ZERO)
      'and keep looping until we get a non blank line or we get to the end
      Loop Until LineLength > 0 Or CurrentLine = TotalLines
      'if charpos = -1 then we are at the end.  Send them back to beginning
      If CharPos = -1 Then CharPos = 0
      'place cursor
      txtLyrics.SelStart = CharPos
   End If
   
   '/* update tags
   If Trim(txtLyrics.Text) <> "" Then Update_Tags_Ref
   
   'and set the focus back to the text box
   txtLyrics.SetFocus
   Exit Sub
HELL:

End Sub

Private Function Format_PlayList(tID3v1 As ptID3, iIndex As Integer) As String
  On Error Resume Next
 
 Dim sFullPath As String, sFileName As String, sFileEx As String
 Dim sFormat As String, sNewString As String, SplitField() As String, CleanStr As String
 Dim i As Integer
 Dim iSpaces As Integer
  
   sFullPath = Trim(FileTags.Path & "\" & FileTags.List(iIndex))
   sFileName = Left(FileTags.List(iIndex), Len(FileTags.List(iIndex)) - 4)
   sFileEx = Right(FileTags.List(iIndex), 3)
        
   sFormat = ""
   If Trim(tID3v1.Title) = "" Then tID3v1.Title = sFileName
   sFormat = Replace(sFormatPlayList, "%S", Trim(tID3v1.Title))
   sFormat = Replace(sFormat, "%A", Trim(tID3v1.Artist))
   sFormat = Replace(sFormat, "%B", Trim(tID3v1.Album))
   sFormat = Replace(sFormat, "%Y", Trim(tID3v1.Year))
   sFormat = Replace(sFormat, "%G", Trim(tID3v1.GenreName))
   ' sFormat = Replace(sFormat, "%T", Trim(tID3v1.lenght))
   sFormat = Replace(sFormat, "%T", "")
   sFormat = Replace(sFormat, "%N", sFileName)
   sFormat = Replace(sFormat, "%P", sFullPath)
   sFormat = Replace(sFormat, "%F", sFileEx)
   
   If sFormat = sFormatPlayList Then sFormat = sFileName
      
   '------------------------------------------------------------------------------
    CleanStr = Trim$(sFormat)
    
    'Upper case and / or lower case the string correctly.
    SplitField = Split(CleanStr, " ", , vbTextCompare)
    CleanStr = ""
    For iSpaces = 0 To UBound(SplitField)
        If Not iSpaces = 0 Or Not IsNumeric(SplitField(iSpaces)) Then
          sNewString = UCase$(Left$(SplitField(iSpaces), 1))
          sNewString = sNewString & LCase$(Right$(SplitField(iSpaces), Len(SplitField(iSpaces)) - 1))
          CleanStr = CleanStr & sNewString & " "
        End If
    Next iSpaces
    Format_PlayList = Trim$(CleanStr)
  '------------------------------------------------------------------------------
  
End Function


Private Sub Save_Tags()
    
    Dim strFileName As String
    Dim OldTag As ptID3
    Dim NewTag As ptID3
    Dim i As Integer
    Dim iCount As Integer
    Dim iFUpdated As Integer
    
    On Error Resume Next
     '// if no checked all checkbox
    If FileTags.ListCount = 0 Then Exit Sub
    
    If FilesSelected > 1 Then
        If chkTags(0).Value = vbUnchecked And chkTags(1).Value = vbUnchecked And chkTags(2).Value = vbUnchecked _
         And chkTags(3).Value = vbUnchecked Then
         Exit Sub
        End If
    End If
    
    '// reset values for progress bar
    pbProgress.min = 0
    pbProgress.Max = FileTags.ListCount
    pbProgress.Value = 0
    
    pbProgress.Visible = True
  
    For i = 0 To FileTags.ListCount - 1
        strFileName = FileTags.Path & "\" & FileTags.List(i)
          
        DoEvents
          
       '// more than one files selected
       If FilesSelected > 1 Then
       
          If FileTags.Selected(i) = True Then
            lblFile.Caption = "Updating file: " & FileTags.List(i)
            
            '// make new tag
            NewTag.Title = Trim(txtTitle.Text)
            NewTag.Artist = Trim(txtArtist.Text)
            NewTag.Album = Trim(txtAlbum.Text)
            NewTag.Year = Trim(txtYear.Text)
            NewTag.Genre = Val(cboGenre.ListIndex)
            NewTag.GenreName = CStr(cboGenre.Text)
            
            '// load old tags
            OldTag = ReadFile_Tags(strFileName)
          
            NewTag.Title = OldTag.Title
            '// Artist unchecked change at old
            If chkTags(0).Value = vbUnchecked Then NewTag.Artist = OldTag.Artist
          
            '// Album unchecked change at old
            If chkTags(1).Value = vbUnchecked Then NewTag.Album = OldTag.Album
          
            '// year unchecked change at old
            If chkTags(2).Value = vbUnchecked Then NewTag.Year = OldTag.Year
          
            '// Genre unchecked change at old
            If chkTags(3).Value = vbUnchecked Then NewTag.Genre = OldTag.Genre
          
            NewTag.Comment = OldTag.Comment
            NewTag.Lyrics = OldTag.Lyrics
            
            '// write the tags
            WriteTag strFileName, NewTag
          
            '// modificar la lista de rep
            If iCurrentAlbum = intActiveAlbum Then
               frmMain.ListRep.List(i) = Format_PlayList(NewTag, i)
            End If
            
            If LCase(strFileName) = LCase(sFileMainPlaying) Then
               frmMain.Load_File_Tags
               If bolLyricsShow = True Then frmMain.Start_Lyrics
            End If
                      
            iFUpdated = iFUpdated + 1
            
         End If
         
       ElseIf Trim(listRef.ListItems.Item(FileTags.List(i)).SubItems(1)) <> "" Then
     
         '// make new tag
          NewTag.Title = Trim(listRef.ListItems.Item(FileTags.List(i)).SubItems(1))
          NewTag.Artist = Trim(listRef.ListItems.Item(FileTags.List(i)).SubItems(2))
          NewTag.Album = Trim(listRef.ListItems.Item(FileTags.List(i)).SubItems(3))
          NewTag.Comment = ""
          NewTag.Year = Trim(listRef.ListItems.Item(FileTags.List(i)).SubItems(4))
          NewTag.Genre = Val(listRef.ListItems.Item(FileTags.List(i)).SubItems(5))
          NewTag.GenreName = CStr(cboGenre.Text)
          NewTag.Lyrics = Trim(listRef.ListItems.Item(FileTags.List(i)).SubItems(6))
          
          '// write the tags
          WriteTag strFileName, NewTag
                              
          '// modificar la lista de rep
          If iCurrentAlbum = intActiveAlbum Then
            frmMain.ListRep.List(i) = Format_PlayList(NewTag, i)
          End If
                                        
          If LCase(strFileName) = LCase(sFileMainPlaying) Then
             frmMain.Load_File_Tags
             If bolLyricsShow = True Then frmMain.Start_Lyrics
          End If
          
          listRef.ListItems.Item(FileTags.List(i)).SubItems(1) = ""
          listRef.ListItems.Item(FileTags.List(i)).SubItems(2) = ""
          listRef.ListItems.Item(FileTags.List(i)).SubItems(3) = ""
          listRef.ListItems.Item(FileTags.List(i)).SubItems(4) = ""
          listRef.ListItems.Item(FileTags.List(i)).SubItems(5) = ""
          listRef.ListItems.Item(FileTags.List(i)).SubItems(6) = ""
          
          
          iFUpdated = iFUpdated + 1
          
       End If
          
          iCount = iCount + 1
          pbProgress.Value = iCount


  Next i
     pbProgress.Visible = False
     lblFile.Caption = " Listooooo! Updated [ " & iFUpdated & " ] files"
End Sub


Private Sub cmdApply_Click()
 On Error Resume Next
  cmdApply.Enabled = False
   Save_Tags
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Public Sub Make_List_Ref()
 Dim i As Integer
 If FileTags.ListCount = 0 Then Exit Sub
   listRef.ListItems.Clear
   For i = 0 To FileTags.ListCount - 1
        listRef.ListItems.Add , FileTags.List(i), FileTags.List(i)
   Next i
End Sub

Private Sub cmdNext_Click()
If TotalAlbumS = 0 Then Exit Sub
 
 iCurrentAlbum = iCurrentAlbum + 1
 
 If iCurrentAlbum > TotalAlbumS Then
   iCurrentAlbum = 1
 End If
 
 FileTags.Path = frmMain.btnAlbum(iCurrentAlbum).ToolTipText
 Make_List_Ref
 If FileTags.ListCount > 0 Then FileTags.Selected(0) = True
End Sub

Private Sub cmdOk_Click()
 cmdOk.Enabled = False
 If cmdApply.Enabled = True Then Save_Tags
 Unload Me
End Sub

Private Sub cmdPlayer_Click(Index As Integer)
  
  If FileTags.ListCount = 0 Then Exit Sub
  
  Select Case Index
    Case 0 '// skip backward
       Five_Seg_Backward
    Case 1 '// play
       If frmMain.PlayerIsPlaying = "true" Then frmMain.Stop_Player
       If FileTags.ListIndex = -1 Then FileTags.ListIndex = 0
       
       Player_Play FileTags.Path & "\" & FileTags.List(FileTags.ListIndex)
    Case 2 '// pause
       Pause_Play
    Case 3 '// stop
       Stop_Player
       FilePlaying = ""
    Case 4 '// skip forward
       Five_Seg_Forward
  End Select
     txtLyrics.SetFocus

End Sub

Private Sub cmdPrev_Click()
 
 If TotalAlbumS = 0 Then Exit Sub
 
 iCurrentAlbum = iCurrentAlbum - 1
 
 If iCurrentAlbum = 0 Then
   iCurrentAlbum = TotalAlbumS
 End If
 
 FileTags.Path = frmMain.btnAlbum(iCurrentAlbum).ToolTipText
 Make_List_Ref

 If FileTags.ListCount > 0 Then FileTags.Selected(0) = True
End Sub

Private Sub cmdSelAll_Click()
Dim i As Integer
For i = 0 To FileTags.ListCount - 1
 FileTags.Selected(i) = True
Next i
End Sub

Private Sub cmdUndo_Click()
 Dim fin As Integer, j As Integer, Start As Integer
  On Error GoTo HELL
  
    With txtLyrics
      Start = Undo_Pop
      If Start = 0 Then Start = 1
      'select the timestamp
       j = InStr(Start, txtLyrics.Text, "[")
         If j > 0 Then
            fin = InStr(Start + 1, txtLyrics.Text, "]")
            '// solo agregar letras hasta el formato 00:00:00
            If ((fin - 1) - j) > 9 Then
              fin = 0
            End If
         End If
      'get the postion of the last timestamp from the stack
      If Start = 1 Then Start = 0
      .SelStart = Start
      .SelLength = (fin - Start)
      'and delete it
      .SelText = ""
      .SetFocus
   End With
   'If there is nothing in the stack, undo should not be enabled
   If Cur = 0 Then cmdUndo.Enabled = False
Exit Sub
HELL:

End Sub


Private Sub Texts_Enableds(bolEnabled As Boolean)
   lblFile.Caption = ""
   chkTags(0).Value = vbUnchecked
   chkTags(1).Value = vbUnchecked
   chkTags(2).Value = vbUnchecked
   chkTags(3).Value = vbUnchecked
   
   
   chkTags(0).Enabled = Not bolEnabled
   chkTags(1).Enabled = Not bolEnabled
   chkTags(2).Enabled = Not bolEnabled
   chkTags(3).Enabled = Not bolEnabled
   
   chkTags(0).Visible = Not bolEnabled
   chkTags(1).Visible = Not bolEnabled
   chkTags(2).Visible = Not bolEnabled
   chkTags(3).Visible = Not bolEnabled
   txtTitle.Enabled = bolEnabled
   txtArtist.Enabled = bolEnabled
   txtAlbum.Enabled = bolEnabled
   txtYear.Enabled = bolEnabled
   cboGenre.Enabled = bolEnabled
   
   pictab(2).Enabled = bolEnabled
   txtLyrics.Text = ""
   lblMPEGInfo.Caption = ""
   

End Sub



Private Sub FileTags_Click()
 On Error Resume Next
 Dim tID3 As ptID3
 Dim tMPEGInfo As ptMPEG
 Dim i As Integer
 FilesSelected = 0
 
 If FileTags.ListCount = 0 Then Exit Sub
 
 For i = 0 To FileTags.ListCount - 1
   If FileTags.Selected(i) = True Then
     FilesSelected = FilesSelected + 1
   End If
 Next i
 
 If PlayerState <> "false" Then Stop_Player
 
 
 '// pop for stack in lytics function
   Last = 10
   Cur = 0
   ReDim Arr(1 To Last) As Long
   cmdUndo.Enabled = False
 
 If FilesSelected > 1 Then
   Texts_Enableds False
   lblFile.Caption = LineLanguage(61)
   cmdApply.Enabled = False
   Exit Sub
 Else
   Texts_Enableds True
 End If
 
 lblFile.Caption = FileTags.Path & "\" & FileTags.List(FileTags.ListIndex)
 lblFile.ToolTipText = FileTags.Path & "\" & FileTags.List(FileTags.ListIndex)
 
 tID3 = ReadFile_Tags(FileTags.Path & "\" & FileTags.List(FileTags.ListIndex))
 
 txtTitle.Text = tID3.Title
 txtAlbum.Text = tID3.Album
 txtArtist.Text = tID3.Artist
 txtYear.Text = tID3.Year
 cboGenre.ListIndex = 0
 If IsNumeric(tID3.Genre) Then
   cboGenre.ListIndex = CInt(tID3.Genre)
 Else
   cboGenre.Text = tID3.Genre
 End If
 
 txtLyrics.Text = tID3.Lyrics
 
 tMPEGInfo = MPEGInfo(FileTags.Path & "\" & FileTags.List(FileTags.ListIndex))
 
 lblMPEGInfo.Caption = "<> Size: " & tMPEGInfo.FileSize & "  <> Length: " & tMPEGInfo.Duration & vbCrLf & _
                       "<> Header found at: " & tMPEGInfo.Header & " Bytes" & vbCrLf & _
                       "<> MPEG " & tMPEGInfo.Version & " Layer " & tMPEGInfo.Layer & vbCrLf & _
                       "<> " & tMPEGInfo.bitrate & " Kbps, " & tMPEGInfo.Frames & " Frames" & vbCrLf & _
                       "<> " & tMPEGInfo.Frequency & " Hz  " & tMPEGInfo.Mode & vbCrLf & _
                       "<> CRCs: " & tMPEGInfo.CRCs & "  <> Copyrighted: " & tMPEGInfo.Copyrighted & vbCrLf & _
                       "<> Original: " & tMPEGInfo.Original & "  <> Emphasis: " & tMPEGInfo.Emphasis
                       

End Sub

Private Sub Form_Load()
 On Error Resume Next
 '// initialize values for undo functions
   Last = 10
   Cur = 0
   ReDim Arr(1 To Last) As Long
    
  bolTagsShow = True
 
  Load_Language_Tags
  
  Me.Icon = frmMain.Icon
  frmTags.Left = (Screen.Width - frmTags.Width) / 2
  frmTags.Top = (Screen.Height - frmTags.Height) / 2

 PlayerState = "false"
 '// make genres
 PopulateGenres
 
 Load_Album_Tags
 
End Sub

Sub Load_Album_Tags()
 FileTags.Pattern = frmMain.ListRepRef.Pattern
 FileTags.Path = frmMain.ListRepRef.Path
 Make_List_Ref
 FileTags.Selected(frmMain.ListRep.ListIndex) = True
 iCurrentAlbum = intActiveAlbum
End Sub
Private Sub Player_Play(FilePlay As String)

On Error GoTo error
  If PlayerState = "pause" Then Pause_Play: Exit Sub
   
   Stream_Open FilePlay, FSOUND_NORMAL, 1, True, frmMain.VolumeNActuaL
   
   '// volume in main form
   'Stream_SetVolume 1, frmMain.VolumeNActuaL
   PlayerState = "true"
   FilePlaying = FilePlay
Exit Sub
error:
PlayerState = "false"
FilePlaying = ""
Stop_Player
End Sub

Sub Stop_Player()
 On Error Resume Next
  
 If FileTags.ListCount = 0 Then Exit Sub
 
 Stream_Stop 1
 PlayerState = "false"
End Sub

Private Sub Pause_Play()
 Dim CurState As Long
 Dim x
 
 On Error Resume Next
 
 If FileTags.ListCount = 0 Then Exit Sub
 
  If PlayerState = "false" Then Exit Sub
     CurState = Stream_GetState(1)
 '------'Esta Reproduciendo, pausar-------------------------------------------
     If CurState = 2 Then
       Stream_Pause 1
       PlayerState = "pause"
     Else
'------'Si esta pausado, reproducir---------------------------------------------
       Stream_Pause 1
       PlayerState = "true"
     End If
End Sub

Sub Five_Seg_Forward()
 On Error GoTo HELL
 Dim CurPos As Long
 
  If FileTags.ListCount = 0 Or PlayerState = "false" Then Exit Sub
  If PlayerState = "pause" Then Pause_Play
  
  CurPos = Stream_GetPosition(1)
  CurPos = CurPos + 5
  If CurPos > Stream_GetDuration(1) Then CurPos = Stream_GetDuration(1)
  Stream_SetPosition 1, CurPos
Exit Sub
HELL:

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Five_Seg_Backward()
 On Error GoTo HELL
 Dim CurPos As Long
  If FileTags.ListCount = 0 Or PlayerState = "false" Then Exit Sub
  If PlayerState = "pause" Then Pause_Play
  CurPos = Stream_GetPosition(1)
  CurPos = CurPos - 5
  If CurPos < 0 Then CurPos = 0
  Stream_SetPosition 1, CurPos
Exit Sub
HELL:
End Sub


Private Function Convert_Time(ByVal LSec As Long) As String
 Dim HH As Long, MM As Long, SS As Long
 Dim tmp As String
 
 HH = LSec \ 3600  '// calkular horas
 MM = LSec \ 60 Mod 60 '// Calkular minutos
 SS = LSec Mod 60  '// calkular segundos
 
 If HH > 0 Then tmp = Format$(HH, "00:")
 Convert_Time = tmp & Format$(MM, "00:") & Format$(SS, "00")
End Function


Private Sub PopulateGenres()

  With cboGenre
    
        .AddItem "Blues"
        .ItemData(cboGenre.NewIndex) = 1
        
        .AddItem "Classic Rock"
        .ItemData(cboGenre.NewIndex) = 2
        
        .AddItem "Country"
        .ItemData(cboGenre.NewIndex) = 3
        
        .AddItem "Dance"
        .ItemData(cboGenre.NewIndex) = 4
        
        .AddItem "Disco"
        .ItemData(cboGenre.NewIndex) = 5
        
        .AddItem "Funk"
        .ItemData(cboGenre.NewIndex) = 6
        
        .AddItem "Grunge"
        .ItemData(cboGenre.NewIndex) = 7
        
        .AddItem "Hip-Hop"
        .ItemData(cboGenre.NewIndex) = 8
        
        .AddItem "Jazz"
        .ItemData(cboGenre.NewIndex) = 9
        
        .AddItem "Metal"
        .ItemData(cboGenre.NewIndex) = 10
        
        .AddItem "New Age"
        .ItemData(cboGenre.NewIndex) = 11
        
        .AddItem "Oldies"
        .ItemData(cboGenre.NewIndex) = 12
        
        .AddItem "Other"
        .ItemData(cboGenre.NewIndex) = 13
        
        .AddItem "Pop"
        .ItemData(cboGenre.NewIndex) = 14
        
        .AddItem "R&B"
        .ItemData(cboGenre.NewIndex) = 15
        
        .AddItem "Rap"
        .ItemData(cboGenre.NewIndex) = 16
        
        .AddItem "Reggae"
        .ItemData(cboGenre.NewIndex) = 17
        
        .AddItem "Rock"
        .ItemData(cboGenre.NewIndex) = 18
        
        .AddItem "Techno"
        .ItemData(cboGenre.NewIndex) = 19
        
        .AddItem "Industrial"
        .ItemData(cboGenre.NewIndex) = 20
        
        .AddItem "Alternative"
        .ItemData(cboGenre.NewIndex) = 21
        
        .AddItem "Ska"
        .ItemData(cboGenre.NewIndex) = 22
        
        .AddItem "Death Metal"
        .ItemData(cboGenre.NewIndex) = 23
        
        .AddItem "Pranks"
        .ItemData(cboGenre.NewIndex) = 24
        
        .AddItem "Soundtrack"
        .ItemData(cboGenre.NewIndex) = 25
        
        .AddItem "Euro-Techno"
        .ItemData(cboGenre.NewIndex) = 26
        
        .AddItem "Ambient"
        .ItemData(cboGenre.NewIndex) = 27
        
        .AddItem "Trip-Hop"
        .ItemData(cboGenre.NewIndex) = 28
        
        .AddItem "Vocal"
        .ItemData(cboGenre.NewIndex) = 29
        
        .AddItem "Jazz+Funk"
        .ItemData(cboGenre.NewIndex) = 30
        
        .AddItem "Fusion"
        .ItemData(cboGenre.NewIndex) = 31
        
        .AddItem "Trance"
        .ItemData(cboGenre.NewIndex) = 32
        
        .AddItem "Classical"
        .ItemData(cboGenre.NewIndex) = 33
        
        .AddItem "Instrumental"
        .ItemData(cboGenre.NewIndex) = 34
        
        .AddItem "Acid"
        .ItemData(cboGenre.NewIndex) = 35
        
        .AddItem "House"
        .ItemData(cboGenre.NewIndex) = 36
        
        .AddItem "Game"
        .ItemData(cboGenre.NewIndex) = 37
        
        .AddItem "Sound Clip"
        .ItemData(cboGenre.NewIndex) = 38
        
        .AddItem "Gospel"
        .ItemData(cboGenre.NewIndex) = 39
        
        .AddItem "Noise"
        .ItemData(cboGenre.NewIndex) = 40
        
        .AddItem "AlternRock"
        .ItemData(cboGenre.NewIndex) = 41
        
        .AddItem "Bass"
        .ItemData(cboGenre.NewIndex) = 42
        
        .AddItem "Soul"
        .ItemData(cboGenre.NewIndex) = 43
        
        .AddItem "Punk"
        .ItemData(cboGenre.NewIndex) = 44
        
        .AddItem "Space"
        .ItemData(cboGenre.NewIndex) = 45
        
        .AddItem "Meditative"
        .ItemData(cboGenre.NewIndex) = 46
        
        .AddItem "Instrumental Pop"
        .ItemData(cboGenre.NewIndex) = 47
        
        .AddItem "Instrumental Rock"
        .ItemData(cboGenre.NewIndex) = 48
        
        .AddItem "Ethnic"
        .ItemData(cboGenre.NewIndex) = 49
        
        .AddItem "Gothic"
        .ItemData(cboGenre.NewIndex) = 50
        
        .AddItem "Darkwave"
        .ItemData(cboGenre.NewIndex) = 51
        
        .AddItem "Techno-Industrial"
        .ItemData(cboGenre.NewIndex) = 52
        
        .AddItem "Electronic"
        .ItemData(cboGenre.NewIndex) = 53
        
        .AddItem "Pop-Folk"
        .ItemData(cboGenre.NewIndex) = 54
        
        .AddItem "Eurodance"
        .ItemData(cboGenre.NewIndex) = 55
        
        .AddItem "Dream"
        .ItemData(cboGenre.NewIndex) = 56
        
        .AddItem "Southern Rock"
        .ItemData(cboGenre.NewIndex) = 57
        
        .AddItem "Comedy"
        .ItemData(cboGenre.NewIndex) = 58
        
        .AddItem "Cult"
        .ItemData(cboGenre.NewIndex) = 59
        
        .AddItem "Gangsta"
        .ItemData(cboGenre.NewIndex) = 60
        
        .AddItem "Top 40"
        .ItemData(cboGenre.NewIndex) = 61
        
        .AddItem "Christian Rap"
        .ItemData(cboGenre.NewIndex) = 62
        
        .AddItem "Pop/Funk"
        .ItemData(cboGenre.NewIndex) = 63
        
        .AddItem "Jungle"
        .ItemData(cboGenre.NewIndex) = 64
        
        .AddItem "Native US"
        .ItemData(cboGenre.NewIndex) = 65
        
        .AddItem "Cabaret"
        .ItemData(cboGenre.NewIndex) = 66
        
        .AddItem "New Wave"
        .ItemData(cboGenre.NewIndex) = 67
        
        .AddItem "Psychadelic"
        .ItemData(cboGenre.NewIndex) = 68
        
        .AddItem "Rave"
        .ItemData(cboGenre.NewIndex) = 69
        
        .AddItem "Showtunes"
        .ItemData(cboGenre.NewIndex) = 70
        
        .AddItem "Trailer"
        .ItemData(cboGenre.NewIndex) = 71
        
        .AddItem "Lo-Fi"
        .ItemData(cboGenre.NewIndex) = 72
        
        .AddItem "Tribal"
        .ItemData(cboGenre.NewIndex) = 73
        
        .AddItem "Acid Punk"
        .ItemData(cboGenre.NewIndex) = 74
        
        .AddItem "Acid Jazz"
        .ItemData(cboGenre.NewIndex) = 75
        
        .AddItem "Polka"
        .ItemData(cboGenre.NewIndex) = 76
        
        .AddItem "Retro"
        .ItemData(cboGenre.NewIndex) = 77
        
        .AddItem "Musical"
        .ItemData(cboGenre.NewIndex) = 78
        
        .AddItem "Rock & Roll"
        .ItemData(cboGenre.NewIndex) = 79
        
        .AddItem "Hard Rock"
        .ItemData(cboGenre.NewIndex) = 80
        
        .AddItem "Folk"
        .ItemData(cboGenre.NewIndex) = 81
        
        .AddItem "Folk-Rock"
        .ItemData(cboGenre.NewIndex) = 82
        
        .AddItem "National Folk"
        .ItemData(cboGenre.NewIndex) = 83
        
        .AddItem "Swing"
        .ItemData(cboGenre.NewIndex) = 84
        
        .AddItem "Fast Fusion"
        .ItemData(cboGenre.NewIndex) = 85
        
        .AddItem "Bebop"
        .ItemData(cboGenre.NewIndex) = 86
        
        .AddItem "Latin"
        .ItemData(cboGenre.NewIndex) = 87
        
        .AddItem "Revival"
        .ItemData(cboGenre.NewIndex) = 88
        
        .AddItem "Celtic"
        .ItemData(cboGenre.NewIndex) = 89
        
        .AddItem "Bluegrass"
        .ItemData(cboGenre.NewIndex) = 90
        
        .AddItem "Avantgarde"
        .ItemData(cboGenre.NewIndex) = 91
        
        .AddItem "Gothic Rock"
        .ItemData(cboGenre.NewIndex) = 92
        
        .AddItem "Progressive Rock"
        .ItemData(cboGenre.NewIndex) = 93
        
        .AddItem "Psychedlic Rock"
        .ItemData(cboGenre.NewIndex) = 94
        
        .AddItem "Symphonic Rock"
        .ItemData(cboGenre.NewIndex) = 95
        
        .AddItem "Slow Rock"
        .ItemData(cboGenre.NewIndex) = 96
    
        .AddItem "Big Band"
        .ItemData(cboGenre.NewIndex) = 97
        
        .AddItem "Chorus"
        .ItemData(cboGenre.NewIndex) = 98
        
        .AddItem "Easy Listening"
        .ItemData(cboGenre.NewIndex) = 99
        
        .AddItem "Acoustic"
        .ItemData(cboGenre.NewIndex) = 100
        
        .AddItem "Humour"
        .ItemData(cboGenre.NewIndex) = 101
        
        .AddItem "Speech"
        .ItemData(cboGenre.NewIndex) = 102
        
        .AddItem "Chanson"
        .ItemData(cboGenre.NewIndex) = 103
        
        .AddItem "Opera"
        .ItemData(cboGenre.NewIndex) = 104
        
        .AddItem "Chamber Music"
        .ItemData(cboGenre.NewIndex) = 105
        
        .AddItem "Sonota"
        .ItemData(cboGenre.NewIndex) = 106
        
        .AddItem "Symphony"
        .ItemData(cboGenre.NewIndex) = 107
        
        .AddItem "Booty Bass"
        .ItemData(cboGenre.NewIndex) = 108
        
        .AddItem "Primus"
        .ItemData(cboGenre.NewIndex) = 109
        
        .AddItem "Porn Groove"
        .ItemData(cboGenre.NewIndex) = 110
        
        .AddItem "Satire"
        .ItemData(cboGenre.NewIndex) = 111
        
        .AddItem "Slow Jam"
        .ItemData(cboGenre.NewIndex) = 112
        
        .AddItem "Club"
        .ItemData(cboGenre.NewIndex) = 113
        
        .AddItem "Tango"
        .ItemData(cboGenre.NewIndex) = 114
        
        .AddItem "Samba"
        .ItemData(cboGenre.NewIndex) = 115
        
        .AddItem "Folklore"
        .ItemData(cboGenre.NewIndex) = 116
        
        .AddItem "Ballad"
        .ItemData(cboGenre.NewIndex) = 117
        
        .AddItem "Power Ballad"
        .ItemData(cboGenre.NewIndex) = 118
        
        .AddItem "Rhythmic Soul"
        .ItemData(cboGenre.NewIndex) = 119
        
        .AddItem "Freestyle"
        .ItemData(cboGenre.NewIndex) = 120
    
        .AddItem "Duet"
        .ItemData(cboGenre.NewIndex) = 121
        
        .AddItem "Punk Rock"
        .ItemData(cboGenre.NewIndex) = 122
        
        .AddItem "Drum Solo"
        .ItemData(cboGenre.NewIndex) = 123
        
        .AddItem "A Capella"
        .ItemData(cboGenre.NewIndex) = 124
        
        .AddItem "Eurohouse"
        .ItemData(cboGenre.NewIndex) = 125
        
        .AddItem "Dance Hall"
        .ItemData(cboGenre.NewIndex) = 126
        
        .AddItem "Goa"
        .ItemData(cboGenre.NewIndex) = 127
        
        .AddItem "Drum & Bass"
        .ItemData(cboGenre.NewIndex) = 128
        
       .AddItem "Club-House"
        .ItemData(cboGenre.NewIndex) = 129
       
       .AddItem "Hardcore"
        .ItemData(cboGenre.NewIndex) = 130
        
       .AddItem "Terror"
        .ItemData(cboGenre.NewIndex) = 131
        
       .AddItem "Indie"
        .ItemData(cboGenre.NewIndex) = 132
        
      .AddItem "BritPop"
        .ItemData(cboGenre.NewIndex) = 133
        
      .AddItem "Negerpunk"
        .ItemData(cboGenre.NewIndex) = 134
        
      .AddItem "Polsk Punk"
        .ItemData(cboGenre.NewIndex) = 135
        
      .AddItem "Beat"
        .ItemData(cboGenre.NewIndex) = 136
        
      .AddItem "Christian Gangsta Rap"
        .ItemData(cboGenre.NewIndex) = 137
        
      .AddItem "Heavy Metal"
        .ItemData(cboGenre.NewIndex) = 138
        
      .AddItem "Black Metal"
        .ItemData(cboGenre.NewIndex) = 139
        
      .AddItem "Crossover"
        .ItemData(cboGenre.NewIndex) = 140
        
      .AddItem "Contemporary Christian"
        .ItemData(cboGenre.NewIndex) = 141
   
      .AddItem "Christian Rock"
        .ItemData(cboGenre.NewIndex) = 142
   
      .AddItem "Merengue"
        .ItemData(cboGenre.NewIndex) = 143
        
      .AddItem "Salsa"
        .ItemData(cboGenre.NewIndex) = 144
        
      .AddItem "Trash Metal"
        .ItemData(cboGenre.NewIndex) = 145
        
      .AddItem "Anime"
        .ItemData(cboGenre.NewIndex) = 146
        
      .AddItem "Jpop"
        .ItemData(cboGenre.NewIndex) = 147
        
      .AddItem "Synthpop"
        .ItemData(cboGenre.NewIndex) = 148
            
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
  If PlayerState <> "false" Then Stop_Player
  bolTagsShow = False
End Sub

Private Sub TabStrip_Click()
  pictab(TabStrip.SelectedItem.Index).ZOrder vbBringToFront

End Sub

'//------------------------------------------------------------------------------//
'// functions for undo function in lyrics
Private Sub Undo_Push(Arg As Long)
    Cur = Cur + 1
    On Error GoTo FailPush
        Arr(Cur) = Arg
    Exit Sub
FailPush:
    Last = Last + cChunk  ' Grow
    ReDim Preserve Arr(1 To Last) As Long
    Resume                  ' Try again
End Sub

Private Function Undo_Pop() As Long
    If Cur Then
        Undo_Pop = Arr(Cur)
        Cur = Cur - 1
        If Cur < (Last - cChunk) Then
            Last = Last - cChunk      ' Shrink
            ReDim Preserve Arr(1 To Last) As Long
        End If
    End If
End Function


Private Sub Update_Tags_Ref()
If FileTags.ListCount = 0 Then Exit Sub
If FilesSelected > 1 Then Exit Sub
  With listRef.ListItems.Item(FileTags.List(FileTags.ListIndex))
    .SubItems(1) = txtTitle.Text
    .SubItems(2) = txtArtist.Text
    .SubItems(3) = txtAlbum.Text
    .SubItems(4) = txtYear.Text
    .SubItems(5) = cboGenre.ListIndex
    .SubItems(6) = txtLyrics.Text
  End With
End Sub



Private Sub txtAlbum_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True

End Sub

Private Sub txtArtist_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True
End Sub

Private Sub txtLyrics_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
     KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
 If cmdApply.Enabled = False Then cmdApply.Enabled = True

End Sub

Private Sub txtTitle_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True
End Sub

Private Sub txtYear_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True

End Sub



