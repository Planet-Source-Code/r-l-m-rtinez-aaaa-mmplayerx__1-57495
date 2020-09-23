VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOpciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Options"
   ClientHeight    =   4785
   ClientLeft      =   2550
   ClientTop       =   4380
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picContenedor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3750
      Index           =   5
      Left            =   1965
      ScaleHeight     =   3750
      ScaleWidth      =   6285
      TabIndex        =   12
      Top             =   480
      Width           =   6285
      Begin VB.CommandButton cmdDSPClear 
         Caption         =   "Clear All FX"
         Height          =   315
         Left            =   2625
         TabIndex        =   129
         ToolTipText     =   "Clear all effects"
         Top             =   3420
         Width           =   2520
      End
      Begin VB.CommandButton cmdDSPReset 
         Caption         =   "Default FX Parameters"
         Height          =   315
         Left            =   90
         TabIndex        =   128
         ToolTipText     =   "Reset FX parameters"
         Top             =   3420
         Width           =   2370
      End
      Begin VB.Frame frmDSP 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   0
         Left            =   120
         TabIndex        =   75
         Top             =   630
         Width           =   6060
         Begin VB.CheckBox chkDSP 
            Caption         =   "Chorus"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   76
            Top             =   0
            Width           =   5340
         End
         Begin ComctlLib.Slider sldChorus 
            Height          =   315
            Index           =   6
            Left            =   4650
            TabIndex        =   77
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   4
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldChorus 
            Height          =   315
            Index           =   5
            Left            =   4650
            TabIndex        =   78
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   20
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldChorus 
            Height          =   315
            Index           =   4
            Left            =   1695
            TabIndex        =   79
            Top             =   1920
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   1
            SelStart        =   1
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   1
         End
         Begin ComctlLib.Slider sldChorus 
            Height          =   315
            Index           =   3
            Left            =   1695
            TabIndex        =   80
            Top             =   1560
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldChorus 
            Height          =   315
            Index           =   2
            Left            =   1695
            TabIndex        =   81
            Top             =   1200
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   -99
            Max             =   99
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldChorus 
            Height          =   315
            Index           =   1
            Left            =   1695
            TabIndex        =   82
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   100
            SelStart        =   25
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   25
         End
         Begin ComctlLib.Slider sldChorus 
            Height          =   315
            Index           =   0
            Left            =   1695
            TabIndex        =   83
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   100
            SelStart        =   50
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   50
         End
         Begin VB.Label lblChorus 
            AutoSize        =   -1  'True
            Caption         =   "Phase:"
            Height          =   195
            Index           =   6
            Left            =   3120
            TabIndex        =   90
            Top             =   840
            Width           =   585
         End
         Begin VB.Label lblChorus 
            AutoSize        =   -1  'True
            Caption         =   "Delay:"
            Height          =   195
            Index           =   5
            Left            =   3120
            TabIndex        =   89
            Top             =   480
            Width           =   570
         End
         Begin VB.Label lblChorus 
            AutoSize        =   -1  'True
            Caption         =   "Waveform:"
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   88
            Top             =   1920
            Width           =   960
         End
         Begin VB.Label lblChorus 
            AutoSize        =   -1  'True
            Caption         =   "Frequency:"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   87
            Top             =   1560
            Width           =   960
         End
         Begin VB.Label lblChorus 
            AutoSize        =   -1  'True
            Caption         =   "Feed back:"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   86
            Top             =   1200
            Width           =   945
         End
         Begin VB.Label lblChorus 
            AutoSize        =   -1  'True
            Caption         =   "Depth:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   85
            Top             =   840
            Width           =   585
         End
         Begin VB.Label lblChorus 
            AutoSize        =   -1  'True
            Caption         =   "Wet dry mix:"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   84
            Top             =   480
            Width           =   1125
         End
      End
      Begin VB.Frame frmDSP 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   2
         Left            =   120
         TabIndex        =   63
         Top             =   630
         Width           =   6060
         Begin VB.CheckBox chkDSP 
            Caption         =   "Distortion"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   64
            Top             =   0
            Width           =   5730
         End
         Begin ComctlLib.Slider sldDis 
            Height          =   315
            Index           =   4
            Left            =   1935
            TabIndex        =   65
            Top             =   1920
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            Min             =   100
            Max             =   8000
            SelStart        =   4000
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   4000
         End
         Begin ComctlLib.Slider sldDis 
            Height          =   315
            Index           =   3
            Left            =   1935
            TabIndex        =   66
            Top             =   1560
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            Min             =   100
            Max             =   8000
            SelStart        =   4000
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   4000
         End
         Begin ComctlLib.Slider sldDis 
            Height          =   315
            Index           =   2
            Left            =   1935
            TabIndex        =   67
            Top             =   1200
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            Min             =   100
            Max             =   8000
            SelStart        =   4000
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   4000
         End
         Begin ComctlLib.Slider sldDis 
            Height          =   315
            Index           =   1
            Left            =   1935
            TabIndex        =   68
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            Max             =   100
            SelStart        =   50
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   50
         End
         Begin ComctlLib.Slider sldDis 
            Height          =   315
            Index           =   0
            Left            =   1935
            TabIndex        =   69
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            Min             =   -60
            Max             =   0
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin VB.Label lblDis 
            AutoSize        =   -1  'True
            Caption         =   "Lowpass cutoff:"
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   74
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblDis 
            AutoSize        =   -1  'True
            Caption         =   "Eq Bandwidth:"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   73
            Top             =   1560
            Width           =   1230
         End
         Begin VB.Label lblDis 
            AutoSize        =   -1  'True
            Caption         =   "Eq center freq:"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   72
            Top             =   1200
            Width           =   1290
         End
         Begin VB.Label lblDis 
            AutoSize        =   -1  'True
            Caption         =   "Edge:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   71
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblDis 
            AutoSize        =   -1  'True
            Caption         =   "Gain:"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   70
            Top             =   480
            Width           =   465
         End
      End
      Begin VB.Frame frmDSP 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   4
         Left            =   120
         TabIndex        =   47
         Top             =   630
         Width           =   6060
         Begin VB.CheckBox chkDSP 
            Caption         =   "Flanger"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Width           =   5640
         End
         Begin ComctlLib.Slider sldFlan 
            Height          =   315
            Index           =   6
            Left            =   4530
            TabIndex        =   49
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   4
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldFlan 
            Height          =   315
            Index           =   5
            Left            =   4530
            TabIndex        =   50
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   4
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldFlan 
            Height          =   315
            Index           =   4
            Left            =   1815
            TabIndex        =   51
            Top             =   1920
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   1
            SelStart        =   1
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   1
         End
         Begin ComctlLib.Slider sldFlan 
            Height          =   315
            Index           =   3
            Left            =   1815
            TabIndex        =   52
            Top             =   1560
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldFlan 
            Height          =   315
            Index           =   2
            Left            =   1815
            TabIndex        =   53
            Top             =   1200
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   -99
            Max             =   99
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldFlan 
            Height          =   315
            Index           =   1
            Left            =   1815
            TabIndex        =   54
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   100
            SelStart        =   25
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   25
         End
         Begin ComctlLib.Slider sldFlan 
            Height          =   315
            Index           =   0
            Left            =   1815
            TabIndex        =   55
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   100
            SelStart        =   50
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   50
         End
         Begin VB.Label lblFlan 
            AutoSize        =   -1  'True
            Caption         =   "Wet dry mix:"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   62
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label lblFlan 
            AutoSize        =   -1  'True
            Caption         =   "Depth:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   61
            Top             =   840
            Width           =   585
         End
         Begin VB.Label lblFlan 
            AutoSize        =   -1  'True
            Caption         =   "Feed back:"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   60
            Top             =   1200
            Width           =   945
         End
         Begin VB.Label lblFlan 
            AutoSize        =   -1  'True
            Caption         =   "Frequency:"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   59
            Top             =   1560
            Width           =   960
         End
         Begin VB.Label lblFlan 
            AutoSize        =   -1  'True
            Caption         =   "Waveform:"
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   58
            Top             =   1920
            Width           =   960
         End
         Begin VB.Label lblFlan 
            AutoSize        =   -1  'True
            Caption         =   "Delay:"
            Height          =   195
            Index           =   5
            Left            =   3120
            TabIndex        =   57
            Top             =   480
            Width           =   570
         End
         Begin VB.Label lblFlan 
            AutoSize        =   -1  'True
            Caption         =   "Phase:"
            Height          =   195
            Index           =   6
            Left            =   3120
            TabIndex        =   56
            Top             =   840
            Width           =   585
         End
      End
      Begin VB.Frame frmDSP 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   5
         Left            =   120
         TabIndex        =   41
         Top             =   630
         Width           =   6060
         Begin VB.CheckBox chkDSP 
            Caption         =   "Gargle"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   5595
         End
         Begin ComctlLib.Slider sldGarg 
            Height          =   315
            Index           =   1
            Left            =   1860
            TabIndex        =   43
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   1
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldGarg 
            Height          =   315
            Index           =   0
            Left            =   1860
            TabIndex        =   44
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   10
            Min             =   1
            Max             =   1000
            SelStart        =   500
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   500
         End
         Begin VB.Label lblGarg 
            AutoSize        =   -1  'True
            Caption         =   "Wave shape:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   46
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label lblGarg 
            AutoSize        =   -1  'True
            Caption         =   "Hz:"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   45
            Top             =   480
            Width           =   285
         End
      End
      Begin VB.Frame frmDSP 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   630
         Width           =   6060
         Begin VB.CheckBox chkDSP 
            Caption         =   "I3d Level 2 Reverb"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   16
            Top             =   -30
            Width           =   5760
         End
         Begin ComctlLib.Slider sldL2 
            Height          =   315
            Index           =   5
            Left            =   1695
            TabIndex        =   17
            Top             =   2280
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   -10000
            Max             =   1000
            SelStart        =   -2602
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   -2602
         End
         Begin ComctlLib.Slider sldL2 
            Height          =   315
            Index           =   4
            Left            =   1695
            TabIndex        =   18
            Top             =   1920
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   2
            SelStart        =   1
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   1
         End
         Begin ComctlLib.Slider sldL2 
            Height          =   315
            Index           =   3
            Left            =   1695
            TabIndex        =   19
            Top             =   1560
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   20
            SelStart        =   2
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   2
         End
         Begin ComctlLib.Slider sldL2 
            Height          =   315
            Index           =   2
            Left            =   1695
            TabIndex        =   20
            Top             =   1200
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldL2 
            Height          =   315
            Index           =   1
            Left            =   1695
            TabIndex        =   21
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   -10000
            Max             =   0
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldL2 
            Height          =   315
            Index           =   0
            Left            =   1695
            TabIndex        =   22
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   -10000
            Max             =   0
            SelStart        =   -1000
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   -1000
         End
         Begin ComctlLib.Slider sldL2 
            Height          =   315
            Index           =   11
            Left            =   4695
            TabIndex        =   23
            Top             =   2280
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   20
            Max             =   20000
            SelStart        =   5000
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   5000
         End
         Begin ComctlLib.Slider sldL2 
            Height          =   315
            Index           =   10
            Left            =   4695
            TabIndex        =   24
            Top             =   1920
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   100
            SelStart        =   100
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   100
         End
         Begin ComctlLib.Slider sldL2 
            Height          =   315
            Index           =   9
            Left            =   4695
            TabIndex        =   25
            Top             =   1560
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   100
            SelStart        =   100
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   100
         End
         Begin ComctlLib.Slider sldL2 
            Height          =   315
            Index           =   8
            Left            =   4695
            TabIndex        =   26
            Top             =   1200
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   1
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldL2 
            Height          =   315
            Index           =   7
            Left            =   4710
            TabIndex        =   27
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   -10000
            Max             =   2000
            SelStart        =   200
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   200
         End
         Begin ComctlLib.Slider sldL2 
            Height          =   315
            Index           =   6
            Left            =   4695
            TabIndex        =   28
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   1
            SelStart        =   1
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   1
         End
         Begin VB.Label lblL2 
            AutoSize        =   -1  'True
            Caption         =   "Room:"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   40
            Top             =   480
            Width           =   570
         End
         Begin VB.Label lblL2 
            AutoSize        =   -1  'True
            Caption         =   "Room hf:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   39
            Top             =   840
            Width           =   795
         End
         Begin VB.Label lblL2 
            AutoSize        =   -1  'True
            Caption         =   "Roll off factor:"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   38
            Top             =   1200
            Width           =   1230
         End
         Begin VB.Label lblL2 
            AutoSize        =   -1  'True
            Caption         =   "Decay time:"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   37
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label lblL2 
            AutoSize        =   -1  'True
            Caption         =   "Decay hf ratio:"
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   36
            Top             =   1920
            Width           =   1290
         End
         Begin VB.Label lblL2 
            AutoSize        =   -1  'True
            Caption         =   "Reflections:"
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   35
            Top             =   2280
            Width           =   1005
         End
         Begin VB.Label lblL2 
            AutoSize        =   -1  'True
            Caption         =   "Reflec. delay:"
            Height          =   195
            Index           =   6
            Left            =   3000
            TabIndex        =   34
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label lblL2 
            AutoSize        =   -1  'True
            Caption         =   "Reverb:"
            Height          =   195
            Index           =   7
            Left            =   3000
            TabIndex        =   33
            Top             =   840
            Width           =   690
         End
         Begin VB.Label lblL2 
            AutoSize        =   -1  'True
            Caption         =   "Reverb delay:"
            Height          =   195
            Index           =   8
            Left            =   3000
            TabIndex        =   32
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblL2 
            AutoSize        =   -1  'True
            Caption         =   "Diffusion:"
            Height          =   195
            Index           =   9
            Left            =   3000
            TabIndex        =   31
            Top             =   1560
            Width           =   825
         End
         Begin VB.Label lblL2 
            AutoSize        =   -1  'True
            Caption         =   "Density:"
            Height          =   195
            Index           =   10
            Left            =   3000
            TabIndex        =   30
            Top             =   1920
            Width           =   720
         End
         Begin VB.Label lblL2 
            AutoSize        =   -1  'True
            Caption         =   "Hf reference:"
            Height          =   195
            Index           =   11
            Left            =   3000
            TabIndex        =   29
            Top             =   2280
            Width           =   1140
         End
      End
      Begin VB.Frame frmDSP 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   7
         Left            =   120
         TabIndex        =   117
         Top             =   630
         Width           =   6060
         Begin ComctlLib.Slider sldWaves 
            Height          =   315
            Index           =   0
            Left            =   2040
            TabIndex        =   119
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   -96
            Max             =   0
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldWaves 
            Height          =   315
            Index           =   1
            Left            =   2040
            TabIndex        =   120
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   -96
            Max             =   0
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldWaves 
            Height          =   315
            Index           =   2
            Left            =   2040
            TabIndex        =   121
            Top             =   1200
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   3000
            SelStart        =   1000
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   1000
         End
         Begin ComctlLib.Slider sldWaves 
            Height          =   315
            Index           =   3
            Left            =   2040
            TabIndex        =   122
            Top             =   1560
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   1
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin VB.CheckBox chkDSP 
            Caption         =   "Waves Reverb"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   118
            Top             =   0
            Width           =   6015
         End
         Begin VB.Label lblWaves 
            AutoSize        =   -1  'True
            Caption         =   "In gain:"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   126
            Top             =   480
            Width           =   675
         End
         Begin VB.Label lblWaves 
            AutoSize        =   -1  'True
            Caption         =   "Reverb mix:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   125
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label lblWaves 
            AutoSize        =   -1  'True
            Caption         =   "Reverb time:"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   124
            Top             =   1200
            Width           =   1125
         End
         Begin VB.Label lblWaves 
            AutoSize        =   -1  'True
            Caption         =   "High-freq Ratio:"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   123
            Top             =   1560
            Width           =   1365
         End
      End
      Begin VB.Frame frmDSP 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   3
         Left            =   120
         TabIndex        =   105
         Top             =   645
         Width           =   6060
         Begin VB.CheckBox chkDSP 
            Caption         =   "Echo"
            Height          =   255
            Index           =   3
            Left            =   15
            TabIndex        =   106
            Top             =   -15
            Width           =   5865
         End
         Begin ComctlLib.Slider sldEcho 
            Height          =   315
            Index           =   3
            Left            =   1950
            TabIndex        =   107
            Top             =   1560
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   1
            Max             =   2000
            SelStart        =   333
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   333
         End
         Begin ComctlLib.Slider sldEcho 
            Height          =   315
            Index           =   2
            Left            =   1950
            TabIndex        =   108
            Top             =   1200
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   1
            Max             =   2000
            SelStart        =   333
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   333
         End
         Begin ComctlLib.Slider sldEcho 
            Height          =   315
            Index           =   1
            Left            =   1950
            TabIndex        =   109
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   100
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldEcho 
            Height          =   315
            Index           =   0
            Left            =   1950
            TabIndex        =   110
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   100
            SelStart        =   50
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   50
         End
         Begin ComctlLib.Slider sldEcho 
            Height          =   315
            Index           =   4
            Left            =   1950
            TabIndex        =   111
            Top             =   1920
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   1
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin VB.Label lblEcho 
            AutoSize        =   -1  'True
            Caption         =   "Wet dry mix:"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   116
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label lblEcho 
            AutoSize        =   -1  'True
            Caption         =   "Feedback:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   115
            Top             =   840
            Width           =   885
         End
         Begin VB.Label lblEcho 
            AutoSize        =   -1  'True
            Caption         =   "Left delay:"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   114
            Top             =   1200
            Width           =   915
         End
         Begin VB.Label lblEcho 
            AutoSize        =   -1  'True
            Caption         =   "Right delay:"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   113
            Top             =   1560
            Width           =   1035
         End
         Begin VB.Label lblEcho 
            AutoSize        =   -1  'True
            Caption         =   "Pan delay:"
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   112
            Top             =   1920
            Width           =   915
         End
      End
      Begin VB.Frame frmDSP 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   1
         Left            =   120
         TabIndex        =   91
         Top             =   630
         Width           =   6060
         Begin VB.CheckBox chkDSP 
            Caption         =   "Compressor"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   92
            Top             =   0
            Width           =   6000
         End
         Begin ComctlLib.Slider sldComp 
            Height          =   315
            Index           =   5
            Left            =   4620
            TabIndex        =   93
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            Max             =   4
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldComp 
            Height          =   315
            Index           =   4
            Left            =   1815
            TabIndex        =   94
            Top             =   1920
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            Min             =   1
            Max             =   100
            SelStart        =   10
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   10
         End
         Begin ComctlLib.Slider sldComp 
            Height          =   315
            Index           =   3
            Left            =   1815
            TabIndex        =   95
            Top             =   1560
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            Min             =   -60
            Max             =   0
            SelStart        =   -10
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   -10
         End
         Begin ComctlLib.Slider sldComp 
            Height          =   315
            Index           =   2
            Left            =   1815
            TabIndex        =   96
            Top             =   1200
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            Min             =   50
            Max             =   3000
            SelStart        =   50
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   50
         End
         Begin ComctlLib.Slider sldComp 
            Height          =   315
            Index           =   1
            Left            =   1815
            TabIndex        =   97
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            Max             =   500
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin ComctlLib.Slider sldComp 
            Height          =   315
            Index           =   0
            Left            =   1815
            TabIndex        =   98
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   327682
            Min             =   -60
            Max             =   60
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin VB.Label lblComp 
            AutoSize        =   -1  'True
            Caption         =   "Pre delay"
            Height          =   195
            Index           =   5
            Left            =   3120
            TabIndex        =   104
            Top             =   480
            Width           =   810
         End
         Begin VB.Label lblComp 
            AutoSize        =   -1  'True
            Caption         =   "Ratio:"
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   103
            Top             =   1920
            Width           =   510
         End
         Begin VB.Label lblComp 
            AutoSize        =   -1  'True
            Caption         =   "Threshold:"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   102
            Top             =   1560
            Width           =   915
         End
         Begin VB.Label lblComp 
            AutoSize        =   -1  'True
            Caption         =   "Release:"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   101
            Top             =   1200
            Width           =   750
         End
         Begin VB.Label lblComp 
            AutoSize        =   -1  'True
            Caption         =   "Attack:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   100
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblComp 
            AutoSize        =   -1  'True
            Caption         =   "Gain:"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   99
            Top             =   480
            Width           =   465
         End
      End
      Begin ComctlLib.TabStrip tsDSP 
         Height          =   3390
         Left            =   60
         TabIndex        =   127
         Top             =   0
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   5980
         TabWidthStyle   =   1
         MultiRow        =   -1  'True
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   8
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Chorus"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Compressor"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Distortion"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Echo "
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Flanger"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Gargle"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "I3DL2 Reverb"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Waves Reverbs"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSaveConfig 
      Caption         =   "Save Config."
      Height          =   315
      Left            =   1965
      TabIndex        =   273
      Top             =   4425
      Width           =   2085
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   3825
      Top             =   5025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer_CPU 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6660
      Top             =   6165
   End
   Begin MSComctlLib.ListView lvEQ 
      Height          =   1320
      Left            =   375
      TabIndex        =   169
      Top             =   6030
      Visible         =   0   'False
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   2328
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "eq0"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "eq1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "eq2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "eq3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "eq4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "eq5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "eq6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "eq7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "eq8"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "eq9"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1215
      Top             =   5265
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483635
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Options.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Options.frx":267F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Options.frx":4AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Options.frx":6FA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Options.frx":9611
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Options.frx":BAE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Options.frx":E024
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Options.frx":1056B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ok"
      Height          =   315
      Left            =   4150
      TabIndex        =   1
      Top             =   4425
      Width           =   1305
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   315
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   4425
      UseMaskColor    =   -1  'True
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5555
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   4425
      Width           =   1305
   End
   Begin VB.FileListBox fileBmps 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Hidden          =   -1  'True
      Left            =   2580
      Pattern         =   "*.bmp"
      System          =   -1  'True
      TabIndex        =   3
      Top             =   5085
      Visible         =   0   'False
      Width           =   780
   End
   Begin MSComctlLib.TreeView TreeOptions 
      Height          =   4680
      Left            =   15
      TabIndex        =   4
      Top             =   75
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   8255
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      Style           =   1
      FullRowSelect   =   -1  'True
      Scroll          =   0   'False
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picContenedor 
      BorderStyle     =   0  'None
      Height          =   3750
      Index           =   7
      Left            =   1965
      ScaleHeight     =   3750
      ScaleWidth      =   6285
      TabIndex        =   136
      Top             =   480
      Width           =   6285
      Begin VB.CommandButton cmdVisualizacion 
         Caption         =   "Guardar como"
         Height          =   300
         Index           =   4
         Left            =   1605
         TabIndex        =   272
         Top             =   3435
         Width           =   1530
      End
      Begin VB.ComboBox cboVisualizacion 
         Height          =   315
         Left            =   2580
         Style           =   2  'Dropdown List
         TabIndex        =   271
         Top             =   0
         Width           =   3705
      End
      Begin VB.CommandButton cmdVisualizacion 
         Caption         =   "X"
         Height          =   285
         Index           =   3
         Left            =   1395
         TabIndex        =   261
         Top             =   2895
         Width           =   330
      End
      Begin VB.ListBox lstVis 
         Height          =   450
         ItemData        =   "Options.frx":12BD4
         Left            =   60
         List            =   "Options.frx":12BDE
         TabIndex        =   259
         Top             =   675
         Width           =   1665
      End
      Begin VB.ListBox lstNewVis 
         Height          =   1230
         ItemData        =   "Options.frx":12BFA
         Left            =   60
         List            =   "Options.frx":12BFC
         TabIndex        =   258
         Top             =   1590
         Width           =   1665
      End
      Begin VB.CommandButton cmdVisualizacion 
         Caption         =   "Mostrar"
         Height          =   300
         Index           =   2
         Left            =   4755
         TabIndex        =   257
         Top             =   3435
         Width           =   1530
      End
      Begin VB.CommandButton cmdVisualizacion 
         Caption         =   "Borrar"
         Height          =   300
         Index           =   1
         Left            =   3180
         TabIndex        =   256
         Top             =   3435
         Width           =   1530
      End
      Begin VB.CommandButton cmdVisualizacion 
         Caption         =   "Guardar"
         Height          =   300
         Index           =   0
         Left            =   30
         TabIndex        =   255
         Top             =   3435
         Width           =   1530
      End
      Begin VB.PictureBox picVis 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2745
         Index           =   0
         Left            =   1770
         ScaleHeight     =   183
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   301
         TabIndex        =   223
         Top             =   495
         Visible         =   0   'False
         Width           =   4515
         Begin VB.VScrollBar VSVis 
            Height          =   2685
            Left            =   4200
            Max             =   6
            TabIndex        =   254
            Top             =   30
            Width           =   240
         End
         Begin VB.PictureBox picVisSpect 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4845
            Left            =   0
            ScaleHeight     =   323
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   276
            TabIndex        =   224
            Top             =   0
            Width           =   4140
            Begin VB.CommandButton cmdVis 
               Caption         =   "..."
               Height          =   255
               Left            =   3720
               TabIndex        =   236
               Top             =   1065
               Width           =   405
            End
            Begin VB.ComboBox cboVis 
               Height          =   315
               Index           =   6
               ItemData        =   "Options.frx":12BFE
               Left            =   1725
               List            =   "Options.frx":12C11
               Style           =   2  'Dropdown List
               TabIndex        =   235
               Top             =   3795
               Width           =   1950
            End
            Begin VB.ComboBox cboVis 
               Height          =   315
               Index           =   5
               ItemData        =   "Options.frx":12C24
               Left            =   1725
               List            =   "Options.frx":12C37
               Style           =   2  'Dropdown List
               TabIndex        =   234
               Top             =   3450
               Width           =   1950
            End
            Begin VB.ComboBox cboVis 
               Height          =   315
               Index           =   7
               ItemData        =   "Options.frx":12C4A
               Left            =   1725
               List            =   "Options.frx":12C5A
               Style           =   2  'Dropdown List
               TabIndex        =   233
               Top             =   4140
               Width           =   1950
            End
            Begin VB.ComboBox cboVis 
               Height          =   315
               Index           =   4
               ItemData        =   "Options.frx":12C7B
               Left            =   1725
               List            =   "Options.frx":12C85
               Style           =   2  'Dropdown List
               TabIndex        =   232
               Top             =   2745
               Width           =   1950
            End
            Begin VB.TextBox txtVis 
               Height          =   285
               Index           =   2
               Left            =   1725
               MaxLength       =   2
               TabIndex        =   231
               Text            =   "1"
               Top             =   2415
               Width           =   1950
            End
            Begin VB.TextBox txtVis 
               Height          =   285
               Index           =   1
               Left            =   1725
               MaxLength       =   3
               TabIndex        =   230
               Text            =   "20"
               Top             =   2085
               Width           =   1950
            End
            Begin VB.ComboBox cboVis 
               Height          =   315
               Index           =   3
               ItemData        =   "Options.frx":12C92
               Left            =   1725
               List            =   "Options.frx":12CA2
               Style           =   2  'Dropdown List
               TabIndex        =   229
               Top             =   1380
               Width           =   1950
            End
            Begin VB.ComboBox cboVis 
               Height          =   315
               Index           =   2
               ItemData        =   "Options.frx":12CE8
               Left            =   1725
               List            =   "Options.frx":12CF2
               Style           =   2  'Dropdown List
               TabIndex        =   228
               Top             =   690
               Width           =   1950
            End
            Begin VB.ComboBox cboVis 
               Height          =   315
               Index           =   1
               ItemData        =   "Options.frx":12CFF
               Left            =   1725
               List            =   "Options.frx":12D09
               Style           =   2  'Dropdown List
               TabIndex        =   227
               Top             =   345
               Width           =   1950
            End
            Begin VB.TextBox txtVis 
               Height          =   285
               Index           =   0
               Left            =   1725
               TabIndex        =   226
               Text            =   "[Current Cover Art]"
               Top             =   1050
               Width           =   1950
            End
            Begin VB.ComboBox cboVis 
               Height          =   315
               Index           =   0
               ItemData        =   "Options.frx":12D16
               Left            =   1725
               List            =   "Options.frx":12D23
               Style           =   2  'Dropdown List
               TabIndex        =   225
               Top             =   0
               Width           =   1950
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Back Color:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   13
               Left            =   15
               TabIndex        =   253
               Top             =   4500
               Width           =   1650
            End
            Begin VB.Label lblVisColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   2
               Left            =   1725
               TabIndex        =   252
               Top             =   4500
               Width           =   1950
            End
            Begin VB.Label lblVisColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   1
               Left            =   1725
               TabIndex        =   251
               Top             =   3120
               Width           =   1950
            End
            Begin VB.Label lblVisColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   0
               Left            =   1725
               TabIndex        =   250
               Top             =   1740
               Width           =   1950
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Peak Gravity:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   11
               Left            =   15
               TabIndex        =   249
               Top             =   3810
               Width           =   1650
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Peak Height:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   10
               Left            =   15
               TabIndex        =   248
               Top             =   3465
               Width           =   1650
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Color Bars:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   5
               Left            =   15
               TabIndex        =   247
               Top             =   1740
               Width           =   1650
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Gradient:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   12
               Left            =   15
               TabIndex        =   246
               Top             =   4155
               Width           =   1650
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Peak Color:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   9
               Left            =   15
               TabIndex        =   245
               Top             =   3120
               Width           =   1650
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Mirrored:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   8
               Left            =   15
               TabIndex        =   244
               Top             =   2775
               Width           =   1650
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Bars Spacing:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   7
               Left            =   15
               TabIndex        =   243
               Top             =   2430
               Width           =   1650
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Bars:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   6
               Left            =   15
               TabIndex        =   242
               Top             =   2085
               Width           =   1650
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Scale:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   4
               Left            =   15
               TabIndex        =   241
               Top             =   1395
               Width           =   1650
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Image File:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   3
               Left            =   15
               TabIndex        =   240
               Top             =   1050
               Width           =   1650
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Draw Bars:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   2
               Left            =   15
               TabIndex        =   239
               Top             =   705
               Width           =   1650
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Draw Peaks:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   1
               Left            =   15
               TabIndex        =   238
               Top             =   360
               Width           =   1650
            End
            Begin VB.Label lblVis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00400000&
               Caption         =   "Draw Source:"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   0
               Left            =   15
               TabIndex        =   237
               Top             =   15
               Width           =   1650
            End
         End
      End
      Begin VB.PictureBox picVis 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2745
         Index           =   1
         Left            =   1770
         ScaleHeight     =   183
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   301
         TabIndex        =   264
         Top             =   495
         Visible         =   0   'False
         Width           =   4515
         Begin VB.ComboBox cboVis 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            ItemData        =   "Options.frx":12D44
            Left            =   1470
            List            =   "Options.frx":12D51
            Style           =   2  'Dropdown List
            TabIndex        =   266
            Top             =   750
            Width           =   1950
         End
         Begin VB.TextBox txtVis 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1470
            TabIndex        =   265
            Text            =   "20"
            Top             =   405
            Width           =   1950
         End
         Begin VB.Label lblVisColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   3
            Left            =   1485
            TabIndex        =   270
            Top             =   75
            Width           =   1950
         End
         Begin VB.Label lblVis 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00400000&
            Caption         =   "Color Line:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   14
            Left            =   30
            TabIndex        =   269
            Top             =   75
            Width           =   1410
         End
         Begin VB.Label lblVis 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00400000&
            Caption         =   "Align:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   16
            Left            =   30
            TabIndex        =   268
            Top             =   765
            Width           =   1410
         End
         Begin VB.Label lblVis 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00400000&
            Caption         =   "Lines:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   15
            Left            =   30
            TabIndex        =   267
            Top             =   420
            Width           =   1410
         End
      End
      Begin VB.Label lblCurrentVis 
         AutoSize        =   -1  'True
         Caption         =   "New:"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   263
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label lblCurrentVis 
         AutoSize        =   -1  'True
         Caption         =   "Presents:"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   262
         Top             =   450
         Width           =   810
      End
      Begin VB.Label lblCurrentVis 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Presents Visualizations:"
         Height          =   195
         Index           =   0
         Left            =   465
         TabIndex        =   260
         Top             =   45
         Width           =   2025
      End
   End
   Begin VB.PictureBox picContenedor 
      BorderStyle     =   0  'None
      Height          =   3750
      Index           =   6
      Left            =   1965
      ScaleHeight     =   3750
      ScaleWidth      =   6285
      TabIndex        =   135
      Top             =   480
      Width           =   6285
      Begin VB.PictureBox picEQLines 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1545
         Index           =   1
         Left            =   5550
         ScaleHeight     =   103
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   208
         Top             =   1230
         Width           =   510
         Begin ComctlLib.Slider SldLine 
            Height          =   1605
            Index           =   1
            Left            =   -330
            TabIndex        =   209
            Top             =   0
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   2831
            _Version        =   327682
            Orientation     =   1
            Min             =   -10
         End
         Begin VB.Label lblEQ 
            AutoSize        =   -1  'True
            Caption         =   "+10db"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   14
            Left            =   105
            TabIndex        =   211
            Top             =   120
            Width           =   405
         End
         Begin VB.Label lblEQ 
            AutoSize        =   -1  'True
            Caption         =   "-10db"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   105
            TabIndex        =   210
            Top             =   1320
            Width           =   345
         End
      End
      Begin VB.PictureBox picEQLines 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1545
         Index           =   0
         Left            =   165
         ScaleHeight     =   103
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   204
         Top             =   1230
         Width           =   510
         Begin ComctlLib.Slider SldLine 
            Height          =   1605
            Index           =   0
            Left            =   360
            TabIndex        =   205
            Top             =   0
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   2831
            _Version        =   327682
            Orientation     =   1
            Min             =   -10
            TickStyle       =   1
         End
         Begin VB.Label lblEQ 
            AutoSize        =   -1  'True
            Caption         =   "-10db"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   30
            TabIndex        =   207
            Top             =   1305
            Width           =   345
         End
         Begin VB.Label lblEQ 
            AutoSize        =   -1  'True
            Caption         =   "+10db"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   11
            Left            =   -30
            TabIndex        =   206
            Top             =   105
            Width           =   405
         End
      End
      Begin VB.PictureBox picEQWave 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         DrawWidth       =   2
         FillColor       =   &H80000008&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   1515
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   201
         TabIndex        =   199
         Top             =   615
         Width           =   3015
      End
      Begin VB.CommandButton cmdDeleteEQ 
         Caption         =   "Delete EQ"
         Height          =   315
         Left            =   2325
         TabIndex        =   170
         Top             =   3330
         Width           =   1800
      End
      Begin VB.CommandButton cmdSaveEQ 
         Caption         =   "Save EQ"
         Height          =   315
         Left            =   4365
         TabIndex        =   168
         Top             =   3330
         Width           =   1800
      End
      Begin VB.ComboBox cboEQ 
         Height          =   315
         Left            =   3375
         Style           =   2  'Dropdown List
         TabIndex        =   167
         Top             =   90
         Width           =   2865
      End
      Begin ComctlLib.Slider sldEQ 
         Height          =   1605
         Index           =   9
         Left            =   5055
         TabIndex        =   138
         Top             =   1230
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2831
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1
         Min             =   -10
         TickStyle       =   3
         TickFrequency   =   0
      End
      Begin ComctlLib.Slider sldEQ 
         Height          =   1605
         Index           =   8
         Left            =   4575
         TabIndex        =   139
         Top             =   1230
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2831
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1
         Min             =   -10
         TickStyle       =   3
         TickFrequency   =   0
      End
      Begin ComctlLib.Slider sldEQ 
         Height          =   1605
         Index           =   7
         Left            =   4095
         TabIndex        =   140
         Top             =   1230
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2831
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1
         Min             =   -10
         TickStyle       =   3
         TickFrequency   =   0
      End
      Begin ComctlLib.Slider sldEQ 
         Height          =   1605
         Index           =   6
         Left            =   3615
         TabIndex        =   141
         Top             =   1230
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2831
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1
         Min             =   -10
         TickStyle       =   3
         TickFrequency   =   0
      End
      Begin ComctlLib.Slider sldEQ 
         Height          =   1605
         Index           =   5
         Left            =   3135
         TabIndex        =   142
         Top             =   1230
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2831
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1
         Min             =   -10
         TickStyle       =   3
         TickFrequency   =   0
      End
      Begin ComctlLib.Slider sldEQ 
         Height          =   1605
         Index           =   4
         Left            =   2655
         TabIndex        =   143
         Top             =   1230
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2831
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1
         Min             =   -10
         TickStyle       =   3
         TickFrequency   =   0
      End
      Begin ComctlLib.Slider sldEQ 
         Height          =   1605
         Index           =   3
         Left            =   2175
         TabIndex        =   144
         Top             =   1230
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2831
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1
         Min             =   -10
         TickStyle       =   3
         TickFrequency   =   0
      End
      Begin ComctlLib.Slider sldEQ 
         Height          =   1605
         Index           =   2
         Left            =   1695
         TabIndex        =   145
         Top             =   1230
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2831
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1
         Min             =   -10
         TickStyle       =   3
         TickFrequency   =   0
      End
      Begin ComctlLib.Slider sldEQ 
         Height          =   1605
         Index           =   1
         Left            =   1215
         TabIndex        =   146
         Top             =   1230
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2831
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1
         Min             =   -10
         TickStyle       =   3
         TickFrequency   =   0
      End
      Begin ComctlLib.Slider sldEQ 
         Height          =   1605
         Index           =   0
         Left            =   735
         TabIndex        =   147
         Top             =   1230
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2831
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1
         Min             =   -10
         TickStyle       =   3
         TickFrequency   =   0
      End
      Begin VB.CheckBox chkDSP 
         Caption         =   "Equalizer"
         Height          =   255
         Index           =   7
         Left            =   150
         TabIndex        =   137
         Top             =   105
         Width           =   2040
      End
      Begin VB.Label lblEQ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Presents:"
         Height          =   195
         Index           =   10
         Left            =   2460
         TabIndex        =   166
         Top             =   135
         Width           =   810
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "16K"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   9
         Left            =   5055
         TabIndex        =   165
         Top             =   2865
         Width           =   435
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "14K"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   4575
         TabIndex        =   164
         Top             =   2865
         Width           =   435
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "12K"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   7
         Left            =   4095
         TabIndex        =   163
         Top             =   2865
         Width           =   435
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "6K"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   6
         Left            =   3615
         TabIndex        =   162
         Top             =   2865
         Width           =   435
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "3K"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   3135
         TabIndex        =   161
         Top             =   2865
         Width           =   435
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "1K"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   2655
         TabIndex        =   160
         Top             =   2865
         Width           =   435
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "600"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   2175
         TabIndex        =   159
         Top             =   2865
         Width           =   435
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "310"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   1695
         TabIndex        =   158
         Top             =   2865
         Width           =   435
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "170"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   1215
         TabIndex        =   157
         Top             =   2865
         Width           =   435
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "60"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   735
         TabIndex        =   156
         Top             =   2865
         Width           =   435
      End
   End
   Begin VB.PictureBox picContenedor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3750
      Index           =   2
      Left            =   1965
      ScaleHeight     =   3750
      ScaleWidth      =   6285
      TabIndex        =   5
      Top             =   480
      Width           =   6285
      Begin VB.CheckBox chkProporcional 
         Caption         =   "Proportional"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3120
         TabIndex        =   171
         Top             =   1470
         Width           =   3045
      End
      Begin VB.OptionButton optWallpaper 
         Caption         =   "Strech."
         Height          =   210
         Index           =   3
         Left            =   435
         TabIndex        =   9
         Top             =   1005
         Width           =   2940
      End
      Begin VB.OptionButton optWallpaper 
         Caption         =   "Center."
         Height          =   210
         Index           =   2
         Left            =   435
         TabIndex        =   8
         Top             =   1320
         Width           =   2670
      End
      Begin VB.OptionButton optWallpaper 
         Caption         =   "Tile."
         Height          =   210
         Index           =   1
         Left            =   435
         TabIndex        =   7
         Top             =   1620
         Width           =   2670
      End
      Begin VB.OptionButton optWallpaper 
         Caption         =   "No alter."
         Height          =   195
         Index           =   0
         Left            =   435
         TabIndex        =   6
         Top             =   585
         Width           =   2940
      End
      Begin VB.Label lblWallpaper 
         AutoSize        =   -1  'True
         Caption         =   "Options Wallpaper:"
         Height          =   195
         Left            =   255
         TabIndex        =   172
         Top             =   255
         Width           =   1635
      End
   End
   Begin VB.PictureBox picContenedor 
      BorderStyle     =   0  'None
      Height          =   3750
      Index           =   3
      Left            =   1965
      ScaleHeight     =   3750
      ScaleWidth      =   6285
      TabIndex        =   134
      Top             =   480
      Width           =   6285
      Begin VB.TextBox txtDisplay 
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   201
         Text            =   "%S - %A (%T)"
         Top             =   870
         Width           =   6015
      End
      Begin VB.TextBox txtDisplay 
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   200
         Text            =   "%A - %S"
         Top             =   270
         Width           =   6015
      End
      Begin ComctlLib.Slider sldScrollVel 
         Height          =   420
         Left            =   3330
         TabIndex        =   178
         Top             =   3255
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   741
         _Version        =   327682
         Min             =   100
         Max             =   1000
         SelStart        =   100
         TickFrequency   =   100
         Value           =   100
      End
      Begin VB.OptionButton optScrollType 
         Caption         =   "Zig Zag"
         Height          =   270
         Index           =   1
         Left            =   210
         TabIndex        =   175
         Top             =   3495
         Width           =   3000
      End
      Begin VB.TextBox txtFormat 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   173
         Text            =   "Options.frx":12D79
         Top             =   1395
         Width           =   6015
      End
      Begin VB.OptionButton optScrollType 
         Caption         =   "Rolling"
         Height          =   270
         Index           =   0
         Left            =   210
         TabIndex        =   174
         Top             =   3225
         Value           =   -1  'True
         Width           =   3000
      End
      Begin VB.Label lblPL 
         AutoSize        =   -1  'True
         Caption         =   "scroll text Format:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   203
         Top             =   645
         Width           =   1575
      End
      Begin VB.Label lblPL 
         AutoSize        =   -1  'True
         Caption         =   "Play list Format"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   202
         Top             =   30
         Width           =   1320
      End
      Begin VB.Label lblPL 
         AutoSize        =   -1  'True
         Caption         =   "Scroll Velocity:"
         Height          =   195
         Index           =   3
         Left            =   3330
         TabIndex        =   177
         Top             =   3000
         Width           =   1290
      End
      Begin VB.Label lblPL 
         AutoSize        =   -1  'True
         Caption         =   "Scroll Type:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   176
         Top             =   3000
         Width           =   1035
      End
   End
   Begin VB.PictureBox picContenedor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3750
      Index           =   4
      Left            =   1965
      ScaleHeight     =   3750
      ScaleWidth      =   6285
      TabIndex        =   11
      Top             =   480
      Width           =   6285
      Begin VB.CheckBox chkPlayStart 
         Caption         =   "Automatically play on startup"
         Height          =   285
         Left            =   2475
         TabIndex        =   216
         Top             =   3420
         Width           =   3765
      End
      Begin ComctlLib.Slider sldCrossfade 
         Height          =   510
         Index           =   1
         Left            =   2475
         TabIndex        =   214
         Top             =   2745
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   900
         _Version        =   327682
         LargeChange     =   50
         SmallChange     =   50
         Max             =   400
         SelStart        =   100
         TickFrequency   =   50
         Value           =   100
      End
      Begin ComctlLib.Slider sldCrossfade 
         Height          =   510
         Index           =   0
         Left            =   2475
         TabIndex        =   212
         Top             =   1995
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   900
         _Version        =   327682
         LargeChange     =   50
         SmallChange     =   50
         Max             =   400
         SelStart        =   100
         TickFrequency   =   50
         Value           =   100
      End
      Begin VB.CheckBox chkPIcon 
         Caption         =   "Next Icon."
         Height          =   285
         Index           =   4
         Left            =   2475
         TabIndex        =   154
         Top             =   1395
         Width           =   3525
      End
      Begin VB.CheckBox chkPIcon 
         Caption         =   "Stop Icon."
         Height          =   285
         Index           =   3
         Left            =   2475
         TabIndex        =   153
         Top             =   1125
         Width           =   3525
      End
      Begin VB.CheckBox chkPIcon 
         Caption         =   "Pause Icon."
         Height          =   285
         Index           =   2
         Left            =   2475
         TabIndex        =   152
         Top             =   870
         Width           =   3525
      End
      Begin VB.CheckBox chkPIcon 
         Caption         =   "Play Icon."
         Height          =   285
         Index           =   1
         Left            =   2475
         TabIndex        =   151
         Top             =   585
         Width           =   3525
      End
      Begin VB.CheckBox chkPIcon 
         Caption         =   "Previous Icon."
         Height          =   285
         Index           =   0
         Left            =   2475
         TabIndex        =   150
         Top             =   315
         Width           =   3525
      End
      Begin VB.ListBox lstTypes 
         Height          =   3435
         Left            =   15
         Style           =   1  'Checkbox
         TabIndex        =   133
         Top             =   285
         Width           =   2070
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         Caption         =   "Crossfade in stop:"
         Height          =   195
         Index           =   3
         Left            =   2475
         TabIndex        =   215
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         Caption         =   "Crossfade between Tracks:"
         Height          =   195
         Index           =   2
         Left            =   2475
         TabIndex        =   213
         Top             =   1755
         Width           =   2355
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         Caption         =   "Show system tray icon:"
         Height          =   195
         Index           =   1
         Left            =   2430
         TabIndex        =   149
         Top             =   15
         Width           =   2025
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         Caption         =   "File Types:"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   148
         Top             =   15
         Width           =   930
      End
   End
   Begin VB.PictureBox picContenedor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3750
      Index           =   1
      Left            =   1965
      ScaleHeight     =   3750
      ScaleWidth      =   6285
      TabIndex        =   13
      Top             =   480
      Width           =   6285
      Begin VB.CheckBox chkUseFile 
         Caption         =   "Load region data from file"
         Height          =   255
         Left            =   2700
         TabIndex        =   220
         Top             =   2355
         Width           =   3540
      End
      Begin VB.TextBox txtInfo 
         Height          =   1035
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   131
         Top             =   2670
         Width           =   6195
      End
      Begin VB.ListBox ListaSkins 
         Height          =   2010
         Left            =   60
         TabIndex        =   14
         Top             =   255
         Width           =   6195
      End
      Begin VB.Label lblSkin 
         AutoSize        =   -1  'True
         Caption         =   "Skin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1965
         TabIndex        =   222
         Top             =   -15
         Width           =   420
      End
      Begin VB.Label lblSkin 
         AutoSize        =   -1  'True
         Caption         =   "Skin Info:"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   132
         Top             =   2415
         Width           =   855
      End
      Begin VB.Label lblSkin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Current Skin:"
         Height          =   195
         Index           =   1
         Left            =   750
         TabIndex        =   197
         Top             =   -15
         Width           =   1170
      End
   End
   Begin VB.Frame frmFondo 
      Height          =   4335
      Left            =   1905
      TabIndex        =   130
      Top             =   0
      Width           =   6480
      Begin VB.PictureBox picHead 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00642909&
         BorderStyle     =   0  'None
         FillColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   45
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   425
         TabIndex        =   155
         Top             =   150
         Width           =   6375
      End
   End
   Begin VB.PictureBox picContenedor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3750
      Index           =   0
      Left            =   1965
      ScaleHeight     =   3750
      ScaleWidth      =   6285
      TabIndex        =   10
      Top             =   480
      Width           =   6285
      Begin VB.CommandButton cmdAppConfig 
         Caption         =   "Browse..."
         Height          =   345
         Left            =   3720
         TabIndex        =   219
         Top             =   1260
         Width           =   2355
      End
      Begin VB.Frame fraApp 
         BorderStyle     =   0  'None
         Caption         =   "fraApp"
         Height          =   3210
         Index           =   0
         Left            =   120
         TabIndex        =   179
         Top             =   390
         Width           =   6030
         Begin VB.TextBox txtAppConfig 
            Height          =   285
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   181
            Top             =   510
            Width           =   5865
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File:"
            Height          =   195
            Index           =   10
            Left            =   1155
            TabIndex        =   279
            Top             =   2925
            Width           =   360
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Virtual:"
            Height          =   195
            Index           =   9
            Left            =   885
            TabIndex        =   278
            Top             =   2655
            Width           =   630
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phys:"
            Height          =   195
            Index           =   8
            Left            =   1035
            TabIndex        =   277
            Top             =   2385
            Width           =   480
         End
         Begin VB.Shape shpFrame 
            Height          =   255
            Index           =   3
            Left            =   1560
            Top             =   2910
            Width           =   3135
         End
         Begin VB.Shape shpBar 
            BackStyle       =   1  'Opaque
            DrawMode        =   7  'Invert
            Height          =   255
            Index           =   3
            Left            =   1560
            Top             =   2910
            Width           =   1695
         End
         Begin VB.Shape shpFrame 
            Height          =   255
            Index           =   1
            Left            =   1560
            Top             =   2370
            Width           =   3135
         End
         Begin VB.Shape shpBar 
            BackStyle       =   1  'Opaque
            DrawMode        =   7  'Invert
            Height          =   255
            Index           =   1
            Left            =   1560
            Top             =   2370
            Width           =   1695
         End
         Begin VB.Shape shpFrame 
            Height          =   255
            Index           =   2
            Left            =   1560
            Top             =   2640
            Width           =   3135
         End
         Begin VB.Shape shpBar 
            BackStyle       =   1  'Opaque
            DrawMode        =   7  'Invert
            Height          =   255
            Index           =   2
            Left            =   1560
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Memoria Fisica Libre:"
            Height          =   195
            Index           =   7
            Left            =   1035
            TabIndex        =   221
            Top             =   2100
            Width           =   1815
         End
         Begin VB.Label lblCPU 
            AutoSize        =   -1  'True
            Caption         =   " 100.00"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   2880
            TabIndex        =   218
            Top             =   2100
            Width           =   645
         End
         Begin VB.Label lblApp 
            BackStyle       =   0  'Transparent
            Caption         =   "Nota: Algunas opciones requieren que se reinicie la aplicacion."
            Height          =   585
            Index           =   6
            Left            =   120
            TabIndex        =   217
            Top             =   1335
            Width           =   5760
         End
         Begin VB.Label lblApp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ruta de configuracion."
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   180
            Top             =   240
            Width           =   1875
         End
         Begin VB.Label lblCPU 
            Alignment       =   2  'Center
            Caption         =   "física"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   274
            Top             =   2385
            Width           =   3135
         End
         Begin VB.Label lblCPU 
            Alignment       =   2  'Center
            Caption         =   "usuario"
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   275
            Top             =   2640
            Width           =   3135
         End
         Begin VB.Label lblCPU 
            Alignment       =   2  'Center
            Caption         =   "archivo"
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   276
            Top             =   2925
            Width           =   3135
         End
      End
      Begin VB.Frame fraApp 
         BorderStyle     =   0  'None
         Caption         =   "fraApp"
         Height          =   3210
         Index           =   2
         Left            =   135
         TabIndex        =   192
         Top             =   405
         Width           =   6030
         Begin ComctlLib.Slider Slider1 
            Height          =   675
            Left            =   150
            TabIndex        =   193
            Top             =   735
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   1058
            _Version        =   327682
            LargeChange     =   10
            Min             =   10
            Max             =   100
            SelStart        =   100
            TickStyle       =   2
            TickFrequency   =   10
            Value           =   100
         End
         Begin VB.Label lblApp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100%"
            Height          =   195
            Index           =   5
            Left            =   5220
            TabIndex        =   196
            Top             =   1530
            Width           =   495
         End
         Begin VB.Label lblApp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "10%"
            Height          =   195
            Index           =   4
            Left            =   210
            TabIndex        =   195
            Top             =   1530
            Width           =   390
         End
         Begin VB.Label lblApp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alpha (Only win 2000 or later)"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   194
            Top             =   360
            Width           =   2595
         End
      End
      Begin VB.Frame fraApp 
         BorderStyle     =   0  'None
         Caption         =   "fraApp"
         Height          =   3210
         Index           =   1
         Left            =   135
         TabIndex        =   182
         Top             =   390
         Width           =   6030
         Begin VB.CheckBox chkDir 
            Caption         =   "Enable right click menu in drives and directories"
            Height          =   240
            Left            =   0
            TabIndex        =   189
            Top             =   1380
            Width           =   5955
         End
         Begin VB.ComboBox cboLanguage 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   188
            Top             =   195
            Width           =   4080
         End
         Begin VB.CheckBox chkWindowsState 
            Caption         =   "Barra de Tareas"
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   187
            Top             =   2085
            Width           =   5955
         End
         Begin VB.CheckBox chkWindowsState 
            Caption         =   "System Tray"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   186
            Top             =   2310
            Width           =   5955
         End
         Begin VB.CheckBox chkWindowsState 
            Caption         =   "Multiple Instances."
            Height          =   240
            Index           =   4
            Left            =   0
            TabIndex        =   185
            Top             =   1140
            Width           =   5955
         End
         Begin VB.CheckBox chkWindowsState 
            Caption         =   "Show Splash Screen."
            Height          =   240
            Index           =   3
            Left            =   0
            TabIndex        =   184
            Top             =   900
            Width           =   5955
         End
         Begin VB.CheckBox chkWindowsState 
            Caption         =   "Always on top."
            Height          =   240
            Index           =   2
            Left            =   0
            TabIndex        =   183
            Top             =   660
            Width           =   5955
         End
         Begin VB.Label lblApp 
            AutoSize        =   -1  'True
            Caption         =   "Show MMPlayerX in:"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   191
            Top             =   1830
            Width           =   1740
         End
         Begin VB.Label lblApp 
            AutoSize        =   -1  'True
            Caption         =   "Language:"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   190
            Top             =   255
            Width           =   900
         End
      End
      Begin ComctlLib.TabStrip TSAppConfig 
         Height          =   3645
         Left            =   60
         TabIndex        =   198
         Top             =   30
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   6429
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Path Settings"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "App Config"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Alpha"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bFormLoading As Boolean
Dim i As Integer, IndexScope As Integer
Dim tSpec As ptVisSpect
Dim tScope() As ptVisScope
Dim bLoadingVis As Boolean

Private Sub cboEQ_Click()
 On Error Resume Next
 If lvEQ.ListItems.count = 0 Then Exit Sub
 For i = 1 To 10
    sldEQ(i - 1).Value = CLng(lvEQ.ListItems(cboEQ.ListIndex + 1).SubItems(i))
 Next i
 SetFX 7
 Draw_EQWave
End Sub



Private Sub cboVis_Click(Index As Integer)
 If bLoadingVis = True Then Exit Sub
  If picVis(0).Visible = True Then
    tSpec.DrawBars = cboVis(2).ListIndex
    tSpec.DrawPeaks = cboVis(1).ListIndex
    tSpec.ScaleUp = cboVis(3).ListIndex
    tSpec.Mirrored = cboVis(4).ListIndex
    tSpec.PeakHeight = CInt(cboVis(5).Text)
    tSpec.PeakGravity = CInt(cboVis(6).Text)
    tSpec.Gradient = cboVis(7).Text
    tSpec.GrandientIndex = cboVis(7).ListIndex
    tSpec.DrawSource = cboVis(0).ListIndex
    
    If tSpec.DrawBars = False And tSpec.DrawPeaks = False Then tSpec.DrawBars = True
  End If
  
  If picVis(1).Visible = True Then
     tScope(lstNewVis.ItemData(lstNewVis.ListIndex)).Align = cboVis(8).ListIndex
  End If
  
End Sub

Private Sub cboVisualizacion_Click()
 On Error Resume Next
 Dim sFileVis As String
 Dim s As String, i As Integer
 Dim bExistScope As Boolean
 If bFormLoading = True Then Exit Sub
 
 If cboVisualizacion.ListCount = 0 Then Exit Sub
 bLoadingVis = True
 sFileVis = tAppConfig.AppConfig & "Settings\" & cboVisualizacion.Text & ".vis"
 
 If Dir(sFileVis) = "" Then Exit Sub
 
 i = 0
 ' OSCILLOSCOPE
 Do
   '// leer nombre del equalizador
    s = Read_INI("Oscilloscope_" & i, "Number", "", , , sFileVis)
    If s <> "" Then
      ReDim Preserve tScope(i)
       '// leer valores
       tScope(i).Align = Read_INI("Oscilloscope_" & i, "Align", 1, , , sFileVis)
       If tScope(i).Align < 0 Or tScope(i).Align > 2 Then tScope(i).Align = 1
       tScope(i).BackColorScope = Read_INI("Oscilloscope_" & i, "BackColorScope", RGB(0, 255, 0), , , sFileVis)
       tScope(i).LinesScope = Read_INI("Oscilloscope_" & i, "LinesScope", 50, , , sFileVis)
       If tScope(i).LinesScope < 6 Or tScope(i).LinesScope > 200 Then tScope(i).LinesScope = 50
       bExistScope = True
    End If
    i = i + 1
  Loop While s <> ""
  
  lstNewVis.Clear

  If bExistScope = True Then
    IndexScope = 0
   For i = 0 To UBound(tScope)
     lstNewVis.AddItem "Oscilloscope_" & IndexScope, IndexScope
     lstNewVis.ItemData(IndexScope) = IndexScope
     IndexScope = IndexScope + 1
   Next i
     IndexScope = IndexScope - 1
  Else
    IndexScope = -1
  End If
 ' SPECTRUM
 With tSpec
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
  If .Spacio < 0 Or .Spacio > 11 Then .Spacio = 1
    
  If .DrawBars = False And .DrawPeaks = False Then .DrawBars = True
    
  cboVis(0).ListIndex = .DrawSource
  If .DrawBars = True Then cboVis(1).ListIndex = 1 Else cboVis(1).ListIndex = 0
  If .DrawPeaks = True Then cboVis(2).ListIndex = 1 Else cboVis(2).ListIndex = 0
  txtVis(0).Text = .ImageFile
  cboVis(3).ListIndex = .ScaleUp
  lblVisColor(0).BackColor = .BackColorBar
  txtVis(1).Text = .Bars
  txtVis(2).Text = .Spacio + 1
  If .Mirrored = True Then cboVis(4).ListIndex = 1 Else cboVis(4).ListIndex = 0
  lblVisColor(1).BackColor = .BackColorPeak
  cboVis(5).ListIndex = .PeakHeight - 1
  cboVis(6).ListIndex = .PeakGravity - 1
  lblVisColor(2).BackColor = .BackColor
  cboVis(7).ListIndex = .GrandientIndex
  
  If .Exist = True Then lstNewVis.AddItem "Spectrum"
    
  picVis(0).Visible = False
  picVis(1).Visible = False
  bLoadingVis = False
End With
End Sub

Private Sub chkDSP_Click(Index As Integer)
    Dim i As Integer
  '=================================================================
  'Nota: Habilitar Muchos Efectos causara que el rendimiento
  '      del CPU aumente porque se reproducen a niverl de hardware
  '      En especial el equalizador (en ocaciones)
  '=================================================================
  
    FX_Disable   'Disable all FX
    'Enable FX that are checked
    For i = 0 To chkDSP.count - 1
        If (chkDSP(i).Value = vbChecked) Then
            FX_Enable CLng(i)
            SetFX i
        End If
        
    Next i
End Sub

Public Sub SetFX(intFX As Integer)
    Dim lngX As Long
    Select Case intFX
        Case 0
            FX_SetChorus sldChorus(0).Value, sldChorus(1).Value, sldChorus(2).Value, sldChorus(3).Value, sldChorus(4).Value, sldChorus(5).Value, sldChorus(6).Value
        Case 1
            FX_SetCompressor sldComp(0).Value, sldComp(1).Value, sldComp(2).Value, sldComp(3).Value, sldComp(4).Value, sldComp(5).Value
        Case 2
            FX_SetDistortion sldDis(0).Value, sldDis(1).Value, sldDis(2).Value, sldDis(3).Value, sldDis(4).Value
        Case 3
            FX_SetEcho sldEcho(0).Value, sldEcho(1).Value, sldEcho(2).Value, sldEcho(3).Value, sldEcho(4).Value
        Case 4
            FX_SetFlanger sldFlan(0).Value, sldFlan(1).Value, sldFlan(2).Value, sldFlan(3).Value, sldFlan(4).Value, sldFlan(5).Value, sldFlan(6).Value
        Case 5
            FX_SetGargle sldGarg(0).Value, sldGarg(1).Value
        Case 6
            FX_SetI3DL2Reverb sldL2(0).Value, sldL2(1).Value, sldL2(2).Value, sldL2(3).Value, sldL2(4).Value, sldL2(5).Value, sldL2(6).Value, sldL2(7).Value, sldL2(8).Value, sldL2(9).Value, sldL2(10).Value, sldL2(11).Value
        Case 7
            'Set up equalizer
            For lngX = 0 To sldEQ.count - 1
                FX_SetEQ lngX, -sldEQ(lngX).Value
            Next lngX
        Case 8
            FX_SetWavesReverb sldWaves(0).Value, sldWaves(1).Value, sldWaves(2).Value, sldWaves(3).Value
    End Select
End Sub



Private Sub chkPlayStart_Click()
 bPlayStarting = chkPlayStart.Value
End Sub

Private Sub chkUseFile_Click()
  bLoadRegionFile = chkUseFile.Value
End Sub

Private Sub cmdAppConfig_Click()
 On Error Resume Next
 Dim strNuevaPath As String

 strNuevaPath = Explorador_Para_Directorios(Me.hwnd, LineLanguage(76))
 If Trim(strNuevaPath) = "" Then Exit Sub
 If Right(strNuevaPath, 1) <> "\" Then strNuevaPath = strNuevaPath & "\"
 tAppConfig.AppConfig = strNuevaPath
 txtAppConfig.Text = strNuevaPath
 Load_Equalizer_Values
 Search_Skins_Languages
 Load_Skins_Menu ""
End Sub


Private Sub cmdDeleteEQ_Click()
 On Error Resume Next
 Dim iIndex As Integer, resp As Integer
 
 If lvEQ.ListItems.count = 0 Or lvEQ.ListItems.count = 1 Then Exit Sub
 resp = MsgBox("   " & LineLanguage(193) & "   " & vbCrLf & _
        "   " & cboEQ.List(cboEQ.ListIndex), vbYesNo + vbInformation, "Delete")
           
 If resp = vbNo Then Exit Sub

 iIndex = cboEQ.ListIndex
 cboEQ.RemoveItem iIndex
 lvEQ.ListItems.Remove iIndex + 1
 If cboEQ.ListCount <> 0 Then cboEQ.ListIndex = 0
 Save_Equalizers
End Sub

Private Sub cmdDSPClear_Click()
    'Uncheck all check boxes
    For i = 0 To chkDSP.count - 1
        chkDSP(i).Value = vbUnchecked
    Next i
End Sub

Private Sub cmdDSPReset_Click()
  Select Case tsDSP.SelectedItem.Index
   Case 1 '// chorus
     sldChorus(0).Value = 50: sldChorus(1).Value = 25
     sldChorus(2).Value = 0: sldChorus(3).Value = 0
     sldChorus(4).Value = 1: sldChorus(5).Value = 0
     sldChorus(0).Value = 0
   Case 2 '// compressor
     sldComp(0).Value = 0: sldComp(1).Value = 0
     sldComp(2).Value = 50: sldComp(3).Value = -10
     sldComp(4).Value = 10: sldComp(5).Value = 0
   Case 3 '// distortion
     sldDis(0).Value = 0: sldDis(1).Value = 50
     sldDis(2).Value = 4000: sldDis(3).Value = 4000
     sldDis(4).Value = 4000
   Case 4 '// echo
     sldEcho(0).Value = 50: sldEcho(1).Value = 0
     sldEcho(2).Value = 333: sldEcho(3).Value = 333
     sldEcho(4).Value = 0
   Case 5 '// flanger
     sldFlan(0).Value = 50: sldFlan(1).Value = 25
     sldFlan(2).Value = 0: sldFlan(3).Value = 0
     sldFlan(4).Value = 1: sldFlan(5).Value = 0
     sldFlan(6).Value = 0
   Case 6 '// gargle
     sldGarg(0).Value = 500: sldGarg(1).Value = 0
   Case 7 '// ID3L2 Rev
     sldL2(0).Value = -1000: sldL2(1).Value = 0
     sldL2(2).Value = 0: sldL2(3).Value = 2
     sldL2(4).Value = 1: sldL2(5).Value = -2602
     sldL2(6).Value = 1: sldL2(7).Value = 200
     sldL2(8).Value = 0: sldL2(9).Value = 100
     sldL2(10).Value = 100: sldL2(11).Value = 5000
   Case 8 '// waves rev
     sldWaves(0).Value = 0: sldWaves(1).Value = 0
     sldWaves(2).Value = 1000: sldWaves(3).Value = 0
  End Select
   
End Sub


Sub Load_Equalizer_Values()
 Dim j As Integer, sValue As String
 Dim s As String
 On Error Resume Next
  cboEQ.Clear
  
  i = 0
  Do
    '// leer nombre del equalizador
    s = Read_INI("equalizer_" & i, "name", "", , , tAppConfig.AppConfig & "Settings\Equalizer.eql")
    If s <> "" Then
       cboEQ.AddItem s
       lvEQ.ListItems.Add i + 1, , s
       
       '// leer valores
       For j = 0 To 9
         sValue = Read_INI("equalizer_" & i, "eq" & j, 0, , , tAppConfig.AppConfig & "Settings\Equalizer.eql")
         If IsNumeric(sValue) Then lvEQ.ListItems.Item(i + 1).SubItems(j + 1) = sValue
       Next j
    End If
    i = i + 1
  Loop While s <> ""
  
End Sub

Private Sub cmdSaveConfig_Click()
 On Error Resume Next
   Save_Settings_INI
End Sub

Private Sub cmdSaveEQ_Click()
 Dim s As String
  
 If Dir(tAppConfig.AppConfig & "Settings\", vbDirectory) = "" Then Exit Sub
 
 s = InputBox(LineLanguage(192), "Equalizer", "User " & lvEQ.ListItems.count, Me.Left + (Me.Width / 2) - 2000, Me.Top + 2000)
 If Trim(s) = "" Then Exit Sub
 
 For i = 1 To lvEQ.ListItems.count
   If LCase(Trim(lvEQ.ListItems(i).Text)) = LCase(Trim(s)) Then
    s = "user " & lvEQ.ListItems.count
   End If
 Next i
  
  lvEQ.ListItems.Add lvEQ.ListItems.count + 1, , s
  cboEQ.AddItem s
  cboEQ.ListIndex = cboEQ.ListCount - 1
 For i = 0 To 9
   lvEQ.ListItems.Item(lvEQ.ListItems.count).SubItems(i + 1) = sldEQ(i).Value
 Next i
  Save_Equalizers
End Sub


Sub Save_Equalizers()
 Dim Fnum As Integer, j As Integer
 Dim ArchivoINI As String
 Dim intClave As Integer

 On Error GoTo Bitch
ArchivoINI = tAppConfig.AppConfig & "Settings\Equalizer.eql"

If Dir(ArchivoINI) <> "" Then '// si existe el archivo borrarlo
 SetAttr ArchivoINI, vbNormal
 Kill ArchivoINI
End If
    Fnum = FreeFile  '// numeroaleatorio para asignar al archivo
    Open ArchivoINI For Output As Fnum
    Print #Fnum, "+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+"
    Print #Fnum, "   eQUALIZER VALUES FOR mUSIC mP3 pLAYER X       "
    Print #Fnum, "+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+"
    Print #Fnum, ""
         
    For i = 1 To lvEQ.ListItems.count
        Print #Fnum, "[equalizer_" & i - 1 & "]"
        Print #Fnum, "name=" & lvEQ.ListItems(i).Text
        For j = 1 To 10
          Print #Fnum, "eq" & j - 1 & "=" & lvEQ.ListItems(i).SubItems(j)
        Next j
    Next i
    Close Fnum
Exit Sub
Bitch:
MsgBox err.Description
End Sub

Private Sub Save_Visualizacion(SaveAs As Boolean)
 Dim Fnum As Integer, j As Integer
 Dim ArchivoINI As String
 Dim intClave As Integer
 Dim s As String
 On Error GoTo Bitch
 
 If Dir(tAppConfig.AppConfig & "Settings\", vbDirectory) = "" Then Exit Sub
 
 If SaveAs = True Then
    s = InputBox(LineLanguage(220), "Visualization", "Visualization_" & cboVisualizacion.ListCount, Me.Left + (Me.Width / 2) - 2000, Me.Top + 2000)
    If Trim(s) = "" Then Exit Sub
    ArchivoINI = tAppConfig.AppConfig & "Settings\" & s & ".vis"
 Else
    ArchivoINI = tAppConfig.AppConfig & "Settings\" & cboVisualizacion.Text & ".vis"
 End If

 

If Dir(ArchivoINI) <> "" Then '// si existe el archivo borrarlo
 SetAttr ArchivoINI, vbNormal
 Kill ArchivoINI
End If
    
    Fnum = FreeFile  '// numeroaleatorio para asignar al archivo
    Open ArchivoINI For Output As Fnum
    Print #Fnum, "+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+"
    Print #Fnum, "   vISUALIZATION fILE fOR mUSIC mP3 pLAYER X     "
    Print #Fnum, "+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+"
    Print #Fnum, ""
         
    If IndexScope >= 0 Then
      For i = 0 To UBound(tScope)
          Print #Fnum, "[Oscilloscope_" & i & "]"
          Print #Fnum, "Number=" & i
          Print #Fnum, "Align=" & tScope(i).Align
          Print #Fnum, "LinesScope=" & tScope(i).LinesScope
          Print #Fnum, "BackColorScope=" & tScope(i).BackColorScope
          Print #Fnum, ""
      Next i
    End If
    
    Print #Fnum, ""
    Print #Fnum, "[Spectrum]"
    Print #Fnum, "BackColor=" & tSpec.BackColor
    Print #Fnum, "BackColorBar=" & tSpec.BackColorBar
    Print #Fnum, "BackColorPeak=" & tSpec.BackColorPeak
    Print #Fnum, "Bars=" & tSpec.Bars
    Print #Fnum, "DrawBars=" & tSpec.DrawBars
    Print #Fnum, "DrawPeaks=" & tSpec.DrawPeaks
    Print #Fnum, "DrawSource=" & tSpec.DrawSource
    Print #Fnum, "Exist=" & tSpec.Exist
    Print #Fnum, "Gradient=" & tSpec.Gradient
    Print #Fnum, "GradientIndex=" & tSpec.GrandientIndex
    Print #Fnum, "ImageFile=" & tSpec.ImageFile
    Print #Fnum, "Mirrored=" & tSpec.Mirrored
    Print #Fnum, "PeakGravity=" & tSpec.PeakGravity
    Print #Fnum, "PeakHeight=" & tSpec.PeakHeight
    Print #Fnum, "ScaleUp=" & tSpec.ScaleUp
    Print #Fnum, "Space=" & tSpec.Spacio
        
    Close Fnum
    If SaveAs = True Then
      cboVisualizacion.AddItem s
      bFormLoading = True
      cboVisualizacion.ListIndex = cboVisualizacion.ListCount - 1
      bFormLoading = False
    End If
Exit Sub
Bitch:
Close Fnum
MsgBox err.Description
End Sub

Private Sub cmdVis_Click()
 Dialogo.Filter = "Image Files (*jpg, *.bmp)|*.jpg;*.bmp"
 Dialogo.filename = txtVis(0).Text
 Dialogo.ShowOpen

 If Dir(Dialogo.filename) = "" Or Dialogo.filename = "" Then Exit Sub
 txtVis(0).Text = Dialogo.filename
End Sub

Private Sub cmdVisualizacion_Click(Index As Integer)
 On Error Resume Next
 Dim i As Integer
 Dim ArchivoINI As String
 Dim resp As Integer
 
 Select Case Index
  Case 0 ' save
    If cboVisualizacion.ListCount <= 1 Or Trim(cboVisualizacion.Text) = "" Or lstNewVis.ListCount = 0 Then Exit Sub

    Save_Visualizacion False
  
  Case 1 ' delete
    If cboVisualizacion.ListCount <= 1 Or Trim(cboVisualizacion.Text) = "" Then Exit Sub
    
    resp = MsgBox("   " & LineLanguage(219) & "   " & vbCrLf & _
           "   " & cboVisualizacion.Text, vbYesNo + vbInformation, "Delete")
           
    If resp = vbNo Then Exit Sub
    ArchivoINI = tAppConfig.AppConfig & "Settings\" & cboVisualizacion.Text & ".vis"
    If Dir(ArchivoINI) <> "" Then '// si existe el archivo borrarlo
       SetAttr ArchivoINI, vbNormal
       Kill ArchivoINI
    End If
    cboVisualizacion.RemoveItem cboVisualizacion.ListIndex
    cboVisualizacion.ListIndex = 0
    
  Case 2 ' mostrar
    If cboVisualizacion.ListCount <= 1 Or lstNewVis.ListCount = 0 Then Exit Sub
    tConfigVis = tSpec
    ReDim tConfigVis.arryPeaks(tConfigVis.Bars)
    ReDim tConfigVis.arryWaitPeak(tConfigVis.Bars)
    frmSpectrum.Setup_Visualizacion
        
    ReDim tConfigScope(IndexScope)
    
    For i = 0 To IndexScope
       tConfigScope(i).BackColorScope = tScope(i).BackColorScope
       tConfigScope(i).LinesScope = tScope(i).LinesScope
       tConfigScope(i).Align = tScope(i).Align
    Next i
   
    IndexVisualization = cboVisualizacion.ListIndex
    
  Case 3 ' delete new visualizacion
    If lstNewVis.ListCount <= 1 Or lstNewVis.ListIndex = -1 Then Exit Sub
    
    If lstNewVis.List(lstNewVis.ListIndex) = "Spectrum" Then
      tSpec.Exist = False
      lstNewVis.RemoveItem lstNewVis.ListIndex
    End If
 
    'scope
    If Left(lstNewVis.List(lstNewVis.ListIndex), 1) = "O" Then
      lstNewVis.RemoveItem lstNewVis.ListIndex
      IndexScope = 0
       For i = 0 To lstNewVis.ListCount - 1
          'scope
         If Left(lstNewVis.List(i), 1) = "O" Then
            lstNewVis.ItemData(IndexScope) = IndexScope
            IndexScope = IndexScope + 1
         End If
       Next i
        IndexScope = IndexScope - 1
        If IndexScope < 0 Then
          ReDim tScope(IndexScope)
        Else
          ReDim Preserve tScope(IndexScope)
        End If
    End If
    
    lstNewVis.ListIndex = -1
    picVis(0).Visible = False
    picVis(1).Visible = False
  
   Case 4 ' save as
     If lstNewVis.ListCount = 0 Then Exit Sub
     Save_Visualizacion True
 End Select
 
End Sub


Private Sub lblVisColor_Click(Index As Integer)
 If bLoadingVis = True Then Exit Sub
  Dialogo.ShowColor
  lblVisColor(Index).BackColor = Dialogo.Color
  
  If picVis(0).Visible = True Then
    tSpec.BackColorBar = lblVisColor(0).BackColor
    tSpec.BackColorPeak = lblVisColor(1).BackColor
    tSpec.BackColor = lblVisColor(2).BackColor
  End If
  
  If picVis(1).Visible = True Then
     tScope(lstNewVis.ItemData(lstNewVis.ListIndex)).BackColorScope = lblVisColor(3).BackColor
  End If
End Sub

Private Sub ListaSkins_Click()
  On Error Resume Next
  Dim strInfo As String, strSkinTemp As String
   
   strSkinTemp = tAppConfig.AppConfig & "Skins\" & ListaSkins.List(ListaSkins.ListIndex) & "\skin.ini"
   strInfo = Read_INI("Info", "AuthorName", "", , , strSkinTemp)
   txtInfo.Text = "Author: " & strInfo
   strInfo = Read_INI("Info", "Email", "", , , strSkinTemp)
   txtInfo.Text = txtInfo.Text & vbCrLf & "E-mail: " & strInfo
   strInfo = Read_INI("Info", "Comments", "", , , strSkinTemp)
   txtInfo.Text = txtInfo.Text & vbCrLf & "Comments: " & strInfo
  
End Sub


Private Sub ListaSkins_DblClick()
   Apply_Skin
End Sub

Private Sub lstNewVis_Click()
 On Error Resume Next
If lstNewVis.ListCount = 0 Or lstNewVis.ListIndex = -1 Then Exit Sub
 If lstNewVis.List(lstNewVis.ListIndex) = "Spectrum" Then
    picVis(0).Visible = True
    picVis(1).Visible = False
 End If
 
 'scope
 If Left(lstNewVis.List(lstNewVis.ListIndex), 1) = "O" Then
    lblVisColor(3).BackColor = tScope(lstNewVis.ItemData(lstNewVis.ListIndex)).BackColorScope
    txtVis(3).Text = tScope(lstNewVis.ItemData(lstNewVis.ListIndex)).LinesScope
    cboVis(8).ListIndex = tScope(lstNewVis.ItemData(lstNewVis.ListIndex)).Align
    picVis(0).Visible = False
    picVis(1).Visible = True
 End If
End Sub

Private Sub lstTypes_ItemCheck(Item As Integer)
 If bLoading = True Then Exit Sub
 strPathern = ""
 sFileType = ""
 For i = 0 To lstTypes.ListCount - 1
   If lstTypes.Selected(i) = True Then
       strPathern = strPathern & "*." & lstTypes.List(i) & ";"
       sFileType = sFileType & 1 & ";"
   Else
       sFileType = sFileType & 0 & ";"
   End If
 Next i
  strPathern = Left(strPathern, Len(strPathern) - 1)
  sFileType = Left(sFileType, Len(sFileType) - 1)
End Sub


Private Sub lstVis_DblClick()
 Dim i As Integer
 Dim Eureka As Boolean
 IndexScope = 0
 For i = 0 To lstNewVis.ListCount - 1
   If lstNewVis.List(i) = "Spectrum" Then
     Eureka = True
   End If
   
   'scope
   If Left(lstNewVis.List(i), 1) = "O" Then
     IndexScope = IndexScope + 1
   End If
 Next i
 
 'spectrum
 If lstVis.ListIndex = 0 And Eureka = False Then
    lstNewVis.AddItem lstVis.List(lstVis.ListIndex)
    tSpec.Exist = True
 End If
      
 If IndexScope > 2 Then Exit Sub
 'scope
 If lstVis.ListIndex = 1 Then
   lstNewVis.AddItem "Oscilloscope_" & IndexScope, IndexScope
   lstNewVis.ItemData(IndexScope) = IndexScope
   
   ReDim Preserve tScope(IndexScope)
   tScope(IndexScope).BackColorScope = RGB(0, 255, 0)
   tScope(IndexScope).Align = 1
   tScope(IndexScope).LinesScope = 50
 End If
  
End Sub

Private Sub optScrollType_Click(Index As Integer)
 If bFormLoading = True Then Exit Sub
  Select Case Index
    Case 0  'Rolling
      frmMain.ScrollText(1).ScrollType = Rolling
      frmMain.ScrollText(5).ScrollType = Rolling
      iScrollType = 0
    Case 1  'Zig ZAg
      frmMain.ScrollText(1).ScrollType = ZigZag
      frmMain.ScrollText(5).ScrollType = ZigZag
      iScrollType = 1
  End Select
End Sub

Private Sub sldChorus_Scroll(Index As Integer)
    SetFX 0
End Sub

Private Sub sldComp_Scroll(Index As Integer)
    SetFX 1
End Sub


Private Sub sldCrossfade_Scroll(Index As Integer)
  Select Case Index
    Case 0 '// tracks
       iCrossfadeTrack = sldCrossfade(Index).Value
    Case 1 '// stop
       iCrossfadeStop = sldCrossfade(Index).Value
  End Select
      
    sldCrossfade(Index).ToolTipText = sldCrossfade(Index).Value & " ms"
End Sub

Private Sub sldDis_Scroll(Index As Integer)
    SetFX 2
End Sub

Private Sub sldEcho_Scroll(Index As Integer)
    SetFX 3
End Sub

Private Sub sldFlan_Scroll(Index As Integer)
    SetFX 4
End Sub

Private Sub sldGarg_Scroll(Index As Integer)
    SetFX 5
End Sub

Private Sub sldL2_Scroll(Index As Integer)
    SetFX 6
End Sub

Private Sub sldEQ_Scroll(Index As Integer)
    SetFX 7
    Draw_EQWave
    '// update custom field in listview
    If lvEQ.ListItems.count = 0 Or cboEQ.ListIndex = -1 Then Exit Sub
    lvEQ.ListItems(cboEQ.ListIndex + 1).SubItems(Index + 1) = sldEQ(Index).Value
End Sub
 
Sub Draw_EQWave()
  Dim X1 As Single, Y1 As Single
  Dim X2 As Single, Y2 As Single
  Dim i As Integer, iValue As Integer

  On Error Resume Next
  picEQWave.Cls
  picEQWave.CurrentY = picEQWave.ScaleHeight / 2
  picEQWave.CurrentX = 0
  For i = 0 To 9
     X1 = i * (picEQWave.ScaleWidth / 10)
     X2 = X1 + (picEQWave.ScaleWidth / 10)
     Y1 = picEQWave.ScaleHeight
     iValue = (sldEQ(i).Value + 10)
     Y2 = (iValue * Y1) / 20
     picEQWave.Line Step(0, 0)-(X2, Y2)
  Next i

End Sub


Private Sub sldScrollVel_Scroll()
    frmMain.ScrollText(1).ScrollVelocity = sldScrollVel.Value
    frmMain.ScrollText(5).ScrollVelocity = sldScrollVel.Value
    iScrollVel = sldScrollVel.Value
End Sub

Private Sub sldWaves_Scroll(Index As Integer)
    SetFX 8
End Sub

Private Sub chkPIcon_Click(Index As Integer)
  On Error Resume Next
  '// check if call of form_load
  If bFormLoading = True Then Exit Sub
  
  PlayerTrayIcon.Previous = chkPIcon(0).Value
  PlayerTrayIcon.Play = chkPIcon(1).Value
  PlayerTrayIcon.Pause = chkPIcon(2).Value
  PlayerTrayIcon.Stop = chkPIcon(3).Value
  PlayerTrayIcon.Next = chkPIcon(4).Value
  
  If chkPIcon(Index).Value = vbChecked Then
     ColocarIcono frmMain.txtSTIcon(Index).hwnd, frmMain.ImageList.ListImages(Index + 1).ExtractIcon.Handle, frmMain.Button(Index).ToolTipText & " - MMPlayerX"
  Else
     QuitarIcono frmMain.txtSTIcon(Index).hwnd
  End If
End Sub

Private Sub chkDir_Click()
 On Error Resume Next
  Dim lngRootKey As Long
  Dim RutaExe As String
  lngRootKey = HKEY_CLASSES_ROOT
  
  If bLoading = True Then Exit Sub
  
  '+--------------------------------------------------------------------------------+
  '|procedimento para poner un acceso en el registro para kuando demos click           |
  '|derecho en un folder o driver aparezka el texto 'Search Music Mp3 Player X'
  '|y se ejecute la aplicacion con los parametros enviados en este caso donde dimos
  '| click derecho
  '|las claves son:                                                                    |
  '| --> HKEY_CLASSES_ROOT\Directory\Shell\ 'Texto del Menu'                           |
  '| --> HKEY_CLASSES_ROOT\Directory\Shell\ 'Texto del Menu' \command                  |
  '|                                  con una clave con la ruta de la aplicacion y     |
  '|                                  comandos                                         |
  '+--------------------------------------------------------------------------------+
  
  If chkDir.Value = vbChecked Then
    OpcionesMusic.Directorio = True
    '// obtener la string correcta para ponerla en el registro
    RutaExe = tAppConfig.AppPath & App.EXEName & ".exe %1"
     'Verifikar si existe la clave
    If Not regDoes_Key_Exist(lngRootKey, "Directory\shell\Search Music Mp3 Player X") Then
      regCreate_A_Key lngRootKey, "Directory\shell\Search Music Mp3 Player X"
      regCreate_A_Key lngRootKey, "Directory\shell\Search Music Mp3 Player X\command"
      regCreate_Key_Value lngRootKey, "Directory\shell\Search Music Mp3 Player X\command", "", RutaExe
    End If
    If Not regDoes_Key_Exist(lngRootKey, "Drive\shell\Search Music Mp3 Player X") Then
      regCreate_A_Key lngRootKey, "Drive\shell\Search Music Mp3 Player X"
      regCreate_A_Key lngRootKey, "Drive\shell\Search Music Mp3 Player X\command"
      regCreate_Key_Value lngRootKey, "Drive\shell\Search Music Mp3 Player X\command", "", RutaExe
    End If
  Else
     OpcionesMusic.Directorio = False
     regDelete_A_Key lngRootKey, "Directory\shell\Search Music Mp3 Player X", "command"
     regDelete_A_Key lngRootKey, "Directory\shell", "Search Music Mp3 Player X"
     regDelete_A_Key lngRootKey, "Drive\shell\Search Music Mp3 Player X", "command"
     regDelete_A_Key lngRootKey, "Drive\shell", "Search Music Mp3 Player X"
  End If
End Sub


Private Sub chkProporcional_Click()
  If chkProporcional.Value = vbChecked Then
    OpcionesMusic.Proporcional = True
  Else
    OpcionesMusic.Proporcional = False
  End If
End Sub

Private Sub Apply_Skin()
 On Error Resume Next
 Dim Skins As String

If ListaSkins.ListIndex < 0 Then Exit Sub

cmdApply.Enabled = False
Skins = Trim(ListaSkins.Text)

'// si es el mismo skin no kambiarlo
If LCase(Skins) = LCase(tAppConfig.Skin) Then: cmdApply.Enabled = True: Exit Sub

'// chekar si existe la carpeta
If Dir(tAppConfig.AppConfig & "Skins\" & Skins, vbDirectory) <> "" Then
    frmMain.Visible = False

    '// seleccionar el menu correcto del skin
    For i = 1 To frmPopUp.mnuSkinsAdd.count
      If LCase(Trim(frmPopUp.mnuSkinsAdd(i).Caption)) = LCase(Skins) Then
         frmPopUp.mnuSkinsAdd(i).Checked = True
      Else
         frmPopUp.mnuSkinsAdd(i).Checked = False
      End If
    Next i
    
    '// Cambiar el skin
    Change_Skin Skins
    '// ajustar los bordes
    Form_Mini_Normal
    
    Change_Mask bMiniMask, False
    
    lblSkin(2).Caption = Skins
    
    frmMain.Show_ScrollBar
    
   frmMain.Visible = True
End If

cmdApply.Enabled = True
frmOpciones.ZOrder 0
End Sub




Private Sub chkWindowsState_Click(Index As Integer)
 On Error Resume Next
 
 If bFormLoading = True Then Exit Sub
 
  Select Case Index
    Case 0  '// Show in task bar
       OpcionesMusic.TaskBar = chkWindowsState(0).Value
       frmPopUp.Visible = chkWindowsState(0).Value
      
    Case 1 '// show in sysTray icons
      OpcionesMusic.SysTray = chkWindowsState(1).Value
        If chkWindowsState(1).Value = vbChecked Then
           ColocarIcono frmMain.Text1.hwnd, frmMain.Icon.Handle, frmMain.sSysTrayText
        Else
           QuitarIcono frmMain.Text1.hwnd
        End If
    Case 2 '// Olways on top
       OpcionesMusic.SiempreTop = chkWindowsState(2).Value
       Always_on_Top
    Case 3 '// Show splash screen
       OpcionesMusic.Splash = chkWindowsState(3).Value
    Case 4 '// Multiples instances
       OpcionesMusic.Instancias = chkWindowsState(4).Value
  End Select
End Sub


Private Sub cmdApply_Click()
  On Error Resume Next
      
   Apply_Skin
  
   If frmPopUp.mnuWallpapper.Checked = True Then
     bolCaratulaDefault = False
     ConfigurarWallpaper
   End If
       
   If LCase(OpcionesMusic.Language) <> LCase(cboLanguage.Text) Then
       If cboLanguage.ListIndex = 0 Then
         OpcionesMusic.Language = "\Spanish"
       Else
         OpcionesMusic.Language = Trim(cboLanguage.Text)
       End If
       
       Load_Language OpcionesMusic.Language
   End If
  
   sFormatScroll = UCase(Trim(txtDisplay(1).Text))
   If sFormatScroll = "" Then sFormatScroll = "%S - %A (%T)"
  
   If LCase(Trim(txtDisplay(0).Text)) <> LCase(sFormatPlayList) Then
      sFormatPlayList = UCase(Trim(txtDisplay(0).Text))
      If sFormatPlayList = "" Then sFormatPlayList = "%A - %S"
      frmMain.Load_Format_PlayList
   End If

End Sub


Private Sub cmdCancel_Click()
  Unload Me
End Sub


Private Sub Load_Last_State()

On Error Resume Next
bFormLoading = True
 'configuration options wallpaper
 optWallpaper(0).Value = OpcionesMusic.NoAlteraR
 optWallpaper(1).Value = OpcionesMusic.Mosaico
 optWallpaper(2).Value = OpcionesMusic.Centrar
 optWallpaper(3).Value = OpcionesMusic.Expander

 If OpcionesMusic.Proporcional = True Then chkProporcional.Value = vbChecked
 If OpcionesMusic.Splash = True Then chkWindowsState(3).Value = vbChecked
 If OpcionesMusic.Instancias = True Then chkWindowsState(4).Value = vbChecked
 If OpcionesMusic.Directorio = True Then chkDir.Value = vbChecked
 If OpcionesMusic.SiempreTop = True Then chkWindowsState(2).Value = vbChecked
 If OpcionesMusic.TaskBar = True Then chkWindowsState(0).Value = vbChecked
 If OpcionesMusic.SysTray = True Then chkWindowsState(1).Value = vbChecked
 
 
'// alpha slider
 Slider1.Value = OpcionesMusic.Alpha
 
 '// Player icons
 If PlayerTrayIcon.Previous = True Then chkPIcon(0).Value = vbChecked
 If PlayerTrayIcon.Play = True Then chkPIcon(1).Value = vbChecked
 If PlayerTrayIcon.Pause = True Then chkPIcon(2).Value = vbChecked
 If PlayerTrayIcon.Stop = True Then chkPIcon(3).Value = vbChecked
 If PlayerTrayIcon.Next = True Then chkPIcon(4).Value = vbChecked
 
 '// scroll caption
 txtDisplay(1).Text = sFormatScroll
 txtDisplay(0).Text = sFormatPlayList
 
 sldScrollVel.Value = iScrollVel
 optScrollType(iScrollType).Value = True
 
 txtAppConfig.Text = tAppConfig.AppConfig
 
 lblSkin(2).Caption = tAppConfig.Skin
 If bLoadRegionFile = True Then chkUseFile.Value = vbChecked
 
 If bPlayStarting = True Then chkPlayStart.Value = vbChecked
 
 sldCrossfade(0).Value = iCrossfadeTrack
 sldCrossfade(1).Value = iCrossfadeStop
 
 bFormLoading = False
 
End Sub



Private Sub cmdOk_Click()
 On Error Resume Next
   Me.Hide
   cmdApply.Value = True
   Unload Me
End Sub


Private Sub Form_Load()
 Dim arryFormat() As String
  On Error Resume Next
  
  bolOpcionesShow = True
  
  TreeOptions.Nodes.Add , , "Application", "Application", 1
  TreeOptions.Nodes.Add , , "Skins", "Skins", 2
  TreeOptions.Nodes.Add , , "Wallpaper", "Wallpaper", 3
  TreeOptions.Nodes.Add , , "ScrollText", "Format PL", 4
  TreeOptions.Nodes.Add , , "Player", "Player", 5
  TreeOptions.Nodes.Add , , "Effects", "DSP FX", 6
  TreeOptions.Nodes.Add , , "Equalizer", "Equalizer", 7
  TreeOptions.Nodes.Add , , "Visualization", "Visualization", 8
  
    lstTypes.AddItem "mp3"
    lstTypes.AddItem "wma"
    lstTypes.AddItem "wav"
    lstTypes.AddItem "ogg"
    
    arryFormat = Split(sFileType, ";", , vbTextCompare)
   
   For i = 0 To UBound(arryFormat)
     If CBool(arryFormat(i)) = True Then
      If i <= lstTypes.ListCount - 1 Then lstTypes.Selected(i) = True
     End If
   Next i
   
  Draw_EQWave
   
  Load_Equalizer_Values
  
  Load_Language_Options '// cargar lenguaje siempre
  
  Me.Icon = frmMain.Icon
  
  Load_Last_State
  
 'center form
 Me.Left = (Screen.Width - Me.Width) / 2: Me.Top = (Screen.Height - Me.Height) / 2
 
 Search_Skins_Languages
 
 TreeOptions.Nodes.Item(1).Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
  bolOpcionesShow = False
  Timer_CPU.Enabled = False
  Cancel = -1
  Me.Hide
End Sub

Sub Search_Skins_Languages()
 Dim miNombre As String, strInfo As String
 Dim sPathskin As String

 On Error Resume Next
'search skins in musicmp3/skins only directories
 Timer_CPU.Enabled = True
 fileBmps.Pattern = "*.bmp"
 ListaSkins.Clear
 miNombre = Dir(tAppConfig.AppConfig & "Skins\", vbDirectory) '// recuperar la primera entrada en la ruta
 sPathskin = tAppConfig.AppConfig & "Skins\"
 i = 0
 Do While miNombre <> ""
   If miNombre <> "." And miNombre <> ".." Then
      ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
      If (GetAttr(sPathskin & miNombre) And vbDirectory) = vbDirectory Then
        fileBmps.Path = tAppConfig.AppConfig & "Skins\" & miNombre
        '// chekar si hay archivos jpg o bmps pára ponerlos como posible skin
        If fileBmps.ListCount > 0 Then
          ListaSkins.AddItem miNombre
          '// Seleccionar el skin actual si esta
          If LCase(Trim(miNombre)) = LCase(Trim(tAppConfig.Skin)) Then ListaSkins.Selected(i) = True
          i = i + 1
        End If
      End If
   End If
   miNombre = Dir
 Loop

'-----------------------------------------------------------------------------------
'// buskar los archivos de lenguaje y agragarlos
miNombre = Dir(tAppConfig.AppConfig & "Language\*.lng")
 cboLanguage.Clear
 cboLanguage.AddItem "Spanish"
 cboLanguage.ListIndex = 0
 
i = 1

Do While miNombre <> ""
   If miNombre <> "." And miNombre <> ".." Then
      ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
        If Right(LCase(miNombre), 3) = "lng" Then  '// verifikar la extencion del archivo
           strInfo = Left(Trim(miNombre), Len(Trim(miNombre)) - 4)
           cboLanguage.AddItem strInfo
         '// Seleccionar el lenguaje que se esta utilizando
         If LCase(Trim(strInfo)) = LCase(Trim(OpcionesMusic.Language)) Then
            cboLanguage.ListIndex = i
         End If
         i = i + 1
        End If
   End If
   miNombre = Dir
Loop


'-----------------------------------------------------------------------------------
'// buskar los archivos de visualizacion y agragarlos
fileBmps.Pattern = "*.vis"

cboVisualizacion.Clear
lstNewVis.Clear

If Dir(tAppConfig.AppConfig & "Settings\", vbDirectory) <> "" Then
  fileBmps.Path = tAppConfig.AppConfig & "Settings\"
  
  For i = 0 To fileBmps.ListCount - 1
     cboVisualizacion.AddItem Left(fileBmps.List(i), Len(fileBmps.List(i)) - 4)
  Next i
End If

'-----------------------------------------------------------------------------------
'// buskar los archivos de gradient y agragarlos
fileBmps.Pattern = "*.bmp;*.jpg"
cboVis(7).Clear
If Dir(tAppConfig.AppConfig & "Settings\", vbDirectory) <> "" Then
   fileBmps.Path = tAppConfig.AppConfig & "Settings\"
     
   For i = 0 To fileBmps.ListCount - 1
      cboVis(7).AddItem fileBmps.List(i)
   Next i
End If

End Sub


Private Sub optWallpaper_Click(Index As Integer)
  
  If optWallpaper(0).Value = True Or optWallpaper(3).Value = True Then
    chkProporcional.Value = vbUnchecked
    chkProporcional.Enabled = False
    OpcionesMusic.NoAlteraR = optWallpaper(0).Value
    OpcionesMusic.Expander = optWallpaper(3).Value
    OpcionesMusic.Mosaico = False
    OpcionesMusic.Centrar = False
  Else
    OpcionesMusic.Mosaico = optWallpaper(1).Value
    OpcionesMusic.Centrar = optWallpaper(2).Value
    OpcionesMusic.Expander = False
    OpcionesMusic.NoAlteraR = False
    chkProporcional.Enabled = True
  End If
End Sub


Private Sub Slider1_Scroll()
 On Error GoTo HELL
     '// Ajustar a porcentaje
      Slider1.ToolTipText = (Slider1.Value * 100) / 100 & "%"
      Make_Transparent frmMain.hwnd, Slider1.Value
      OpcionesMusic.Alpha = Slider1.Value
      
      For i = 0 To 9 '// deseleccionar los menus de porcentaje
        frmPopUp.mnuAlpha(i).Checked = False
      Next i
        '// seleccionar el menu de personalizado y  poner porcentaje
        frmPopUp.mnuAlphaPer.Caption = Trim(LineLanguage(34)) & " [ " & Slider1.Value & "% ]"
        frmPopUp.mnuAlphaPer.Checked = True
 Exit Sub
HELL:

End Sub

Sub Select_Option(Index As Integer)
 On Error Resume Next
 picHead.Cls
 TreeOptions.Nodes(Index).Selected = True
 picHead.Print "  " & TreeOptions.SelectedItem.Text
 picContenedor(Index - 1).ZOrder vbBringToFront
 bolOpcionesShow = True
 If Index = 1 Then
   Timer_CPU.Enabled = True
 Else
   Timer_CPU.Enabled = False
 End If
 
End Sub


Private Sub Timer_CPU_Timer()
 On Error Resume Next
  Dim YourMemory As MEMORYSTATUS
  Dim lWidth As Integer
  Dim iValue As Integer

  YourMemory.dwLength = Len(YourMemory)
  GlobalMemoryStatus YourMemory

  With YourMemory
  lblCPU(0).Caption = .dwAvailPhys / 1024 & " KB"
   
   iValue = CInt((CDbl(.dwAvailPhys) * 100) / CDbl(.dwTotalPhys))
   lblCPU(1).Caption = iValue & " %"
        
   lWidth = shpFrame(1).Width * (.dwAvailPhys / .dwTotalPhys)
   If lWidth <> shpBar(1).Width Then
       shpBar(1).Width = lWidth
   End If
        
   iValue = CInt((CDbl(.dwAvailVirtual) * 100) / CDbl(.dwTotalVirtual))
   lblCPU(2).Caption = iValue & " %"
        
   lWidth = shpFrame(2).Width * (.dwAvailVirtual / .dwTotalVirtual)
   If lWidth <> shpBar(2).Width Then
       shpBar(2).Width = lWidth
   End If
        
   iValue = CInt((CDbl(.dwAvailPageFile) * 100) / CDbl(.dwTotalPageFile))
   lblCPU(3).Caption = iValue & " %"
   
   lWidth = shpFrame(3).Width * (.dwAvailPageFile / .dwTotalPageFile)
   If lWidth <> shpBar(3).Width Then
       shpBar(3).Width = lWidth
   End If

  End With
End Sub

Private Sub TreeOptions_Click()
 Select_Option TreeOptions.SelectedItem.Index
End Sub

Private Sub TSAppConfig_Click()
  fraApp(TSAppConfig.SelectedItem.Index - 1).ZOrder 0
  If TSAppConfig.SelectedItem.Index = 1 Then
     cmdAppConfig.ZOrder 0
     cmdAppConfig.Visible = True
  Else
     cmdAppConfig.Visible = False
  End If
End Sub

Private Sub tsDSP_Click()
  frmDSP(tsDSP.SelectedItem.Index - 1).ZOrder 0
End Sub


Private Sub txtVis_Change(Index As Integer)
If bLoadingVis = True Then Exit Sub
 If Index = 1 Or Index = 2 Then
  If IsNumeric(txtVis(Index).Text) = False Then Exit Sub
 End If
 
 
 If picVis(0).Visible = True Then
    tSpec.Bars = txtVis(1).Text
    tSpec.Spacio = txtVis(2).Text - 1
    tSpec.ImageFile = txtVis(0).Text
    If tSpec.Bars < 6 Or tSpec.Bars > 200 Then tSpec.Bars = 200
    If tSpec.Spacio > 10 Then tSpec.Spacio = 10
 End If
 
 If picVis(1).Visible = True Then
    tScope(lstNewVis.ItemData(lstNewVis.ListIndex)).LinesScope = txtVis(3).Text
    If txtVis(3).Text < 6 Or txtVis(3).Text > 200 Then tScope(lstNewVis.ItemData(lstNewVis.ListIndex)).LinesScope = 50
 End If
End Sub

Private Sub VSVis_Scroll()
  picVisSpect.Top = -VSVis.Value * 23
End Sub
