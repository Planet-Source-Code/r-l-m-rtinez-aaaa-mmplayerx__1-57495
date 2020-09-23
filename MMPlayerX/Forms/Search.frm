VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search and add tracks"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Search.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAdd 
      Caption         =   "Add tracks."
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   465
      Width           =   5310
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Start Now"
      Height          =   330
      Left            =   1290
      TabIndex        =   3
      Top             =   1995
      Width           =   2970
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   330
      Left            =   3945
      TabIndex        =   1
      Top             =   75
      Width           =   1605
   End
   Begin VB.ComboBox cboDrives 
      Height          =   315
      ItemData        =   "Search.frx":000C
      Left            =   1065
      List            =   "Search.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   75
      Width           =   2850
   End
   Begin VB.Label lblProgress 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   15
      TabIndex        =   5
      Top             =   780
      Width           =   1995
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Look in:"
      Height          =   195
      Left            =   315
      TabIndex        =   4
      Top             =   120
      Width           =   690
   End
   Begin VB.Label lblProgress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   885
      Index           =   0
      Left            =   15
      TabIndex        =   2
      Top             =   1035
      Width           =   5535
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer, j As Integer
Dim bCancel As Boolean
Dim sLastPath As String

Private Sub cboDrives_Click()
   lblProgress(0).Caption = "All folders"
   sLastPath = "All folders"
End Sub

Private Sub ChkAdd_Click()
  bAddFiles = chkAdd.Value
End Sub

Private Sub cmdBrowse_Click()
 On Error GoTo Hell
 Dim sPath As String
  sPath = Explorador_Para_Directorios(Me.hwnd, LineLanguage(52))

  If sPath = "" Then Exit Sub
  For i = 2 To cboDrives.ListCount - 1
    If LCase(Left(sPath, 1)) = LCase(Left(cboDrives.List(i), 1)) Then
      cboDrives.ListIndex = i
      Exit For
    End If
  Next i
  lblProgress(0).Caption = sPath
  sLastPath = sPath
Exit Sub
Hell:

End Sub



Private Sub cmdSearch_Click()
    Dim FS As New FileSystemObject
    Dim dDrive As drive
    Dim dDrives As Drives
    Dim imp3 As Integer
  On Error Resume Next
   
  If cboDrives.ListIndex < 0 Then Exit Sub
   
  If bSearching = True Then
     If bSearching = True Then bSearching = False
     If cboDrives.ListIndex = 1 Or cboDrives.ListIndex = 0 Then frmMain.Process_Albums True
     GoTo BITCH
  End If
   
  
  cmdSearch.Caption = LineLanguage(233)
  cboDrives.Enabled = False
  cmdBrowse.Enabled = False
  
  '/* Search in All Hard Drives
  If cboDrives.ListIndex = 0 Then
        
     Set dDrives = FS.Drives
     bSearching = True
     For Each dDrive In dDrives
       If dDrive.IsReady = True Then
          If dDrive.DriveType = Fixed Or dDrive.DriveType = CDRom Then
             If imp3 = 0 Then frmMain.Search_Files dDrive.Path & "\", False
             imp3 = imp3 + 1
             frmMain.Start_Search dDrive.Path & "\"
          End If
       End If
    Next
    bSearching = False
    frmMain.Process_Albums True
    GoTo BITCH
  End If
  
  '/* Search in All Drives
  If cboDrives.ListIndex = 1 Then
        
     Set dDrives = FS.Drives
     bSearching = True
     For Each dDrive In dDrives
       If dDrive.IsReady = True Then
          If dDrive.DriveType = Fixed Or dDrive.DriveType = CDRom Then
             If imp3 = 0 Then frmMain.Search_Files dDrive.Path & "\", False
             imp3 = imp3 + 1
             frmMain.Start_Search dDrive.Path & "\"
          End If
       End If
    Next
    bSearching = False
    frmMain.Process_Albums True
    GoTo BITCH
  End If
  
  '/* search in other hard disk
  If cboDrives.ListIndex > 1 And lblProgress(0).Caption = "All folders" Then
    frmMain.Search_Files Left(cboDrives.List(cboDrives.ListIndex), 1) & ":\"
    GoTo BITCH
  End If
  
  '/* search in folder
  If lblProgress(0).Caption <> "All folders" Then
    If Dir(lblProgress(0).Caption, vbDirectory) <> "" Then
       frmMain.Search_Files lblProgress(0).Caption
    End If
  End If

BITCH:
     lblProgress(0).Caption = sLastPath
     cboDrives.Enabled = True
     cmdBrowse.Enabled = True
     cmdSearch.Caption = LineLanguage(232)

End Sub


Private Sub Form_Load()
  On Error Resume Next
    Dim FS As New FileSystemObject
    Dim dDrive As drive
    Dim dDrives As Drives
    
    
    Set dDrives = FS.Drives
      
    cboDrives.AddItem "Local hard drives"
    cboDrives.AddItem "All Drives"
    
    For Each dDrive In dDrives
       If dDrive.IsReady = True Then
          Select Case dDrive.DriveType
              
             Case 0 '/* Desconocido
             Case 1 '/* Separable
             Case 2 '/* Fijo
                cboDrives.AddItem dDrive.DriveLetter & ": [" & dDrive.VolumeName & "]"
             Case 3 '/* Red
             Case 4 '/* CDROM
               cboDrives.AddItem dDrive.DriveLetter & ": [" & dDrive.VolumeName & "]"
             Case 5 '/* Disco RAM
          End Select
       End If
    Next
  Load_Language_Search
  If bAddFiles = True Then chkAdd.Value = vbChecked
  bolSearchShow = True
  Me.Icon = frmMain.Icon
  Me.Left = (Screen.Width - Me.Width) / 2 '// centrar form
  Me.Top = (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
  If bSearching = True Then
     If bSearching = True Then bSearching = False
     If cboDrives.ListIndex = 1 Or cboDrives.ListIndex = 0 Then frmMain.Process_Albums True
  End If

  bolSearchShow = False
End Sub
