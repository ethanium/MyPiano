VERSION 5.00
Begin VB.Form frmPiano 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   1815
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5640
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   Icon            =   "frmPiano.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrBeat 
      Interval        =   150
      Left            =   4718
      Top             =   1249
   End
   Begin VB.Label lblBaseKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " C"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   4088
      TabIndex        =   31
      Top             =   191
      Width           =   240
   End
   Begin VB.Label lblBank 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Index           =   3
      Left            =   1013
      TabIndex        =   30
      ToolTipText     =   "Intro"
      Top             =   431
      Width           =   210
   End
   Begin VB.Label lblBank 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Index           =   2
      Left            =   773
      TabIndex        =   29
      ToolTipText     =   "Intro"
      Top             =   431
      Width           =   210
   End
   Begin VB.Label lblBank 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Index           =   1
      Left            =   533
      TabIndex        =   28
      ToolTipText     =   "Intro"
      Top             =   431
      Width           =   210
   End
   Begin VB.Label lblBank 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Index           =   0
      Left            =   293
      TabIndex        =   27
      ToolTipText     =   "Intro"
      Top             =   431
      Width           =   210
   End
   Begin VB.Label lblDevice 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "þ"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   4853
      TabIndex        =   26
      Top             =   146
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   30
      Left            =   4485
      Top             =   769
      Width           =   255
   End
   Begin VB.Label lblRate 
      BackStyle       =   0  'Transparent
      Height          =   450
      Left            =   3960
      TabIndex        =   25
      Top             =   195
      Width           =   75
   End
   Begin VB.Label lblRatePtr 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   3923
      TabIndex        =   24
      Top             =   341
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   3953
      TabIndex        =   23
      Top             =   191
      Width           =   60
   End
   Begin VB.Label lblBeatCtrl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   6.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   4
      Left            =   3683
      TabIndex        =   22
      ToolTipText     =   "Stop"
      Top             =   431
      Width           =   210
   End
   Begin VB.Label lblBeatCtrl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   6.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   2
      Left            =   3203
      TabIndex        =   21
      ToolTipText     =   "Pause"
      Top             =   431
      Width           =   210
   End
   Begin VB.Label lblBeatCtrl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   6.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   1
      Left            =   2963
      TabIndex        =   20
      ToolTipText     =   "Fill"
      Top             =   431
      Width           =   210
   End
   Begin VB.Label lblBeatCtrl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   6.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   3
      Left            =   3443
      TabIndex        =   19
      ToolTipText     =   "Ending"
      Top             =   431
      Width           =   210
   End
   Begin VB.Label lblBeatCtrl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   6.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   0
      Left            =   2723
      TabIndex        =   18
      ToolTipText     =   "Intro"
      Top             =   431
      Width           =   210
   End
   Begin VB.Label lblBeat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   2723
      TabIndex        =   17
      Top             =   191
      Width           =   1170
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   15
      Left            =   2325
      Top             =   769
      Width           =   255
   End
   Begin VB.Label lblVolume 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   1328
      TabIndex        =   16
      Top             =   461
      Width           =   855
   End
   Begin VB.Label lblVolumePtr 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   2048
      TabIndex        =   14
      Top             =   461
      Width           =   135
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ý"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   5108
      TabIndex        =   13
      Top             =   146
      Width           =   255
   End
   Begin VB.Label lblDrums 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   11
      Left            =   2513
      TabIndex        =   12
      Top             =   521
      Width           =   135
   End
   Begin VB.Label lblDrums 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   10
      Left            =   2393
      TabIndex        =   11
      Top             =   521
      Width           =   135
   End
   Begin VB.Label lblDrums 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   9
      Left            =   2273
      TabIndex        =   10
      Top             =   521
      Width           =   135
   End
   Begin VB.Label lblDrums 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   8
      Left            =   2513
      TabIndex        =   9
      Top             =   401
      Width           =   135
   End
   Begin VB.Label lblDrums 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   7
      Left            =   2393
      TabIndex        =   8
      Top             =   401
      Width           =   135
   End
   Begin VB.Label lblDrums 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   6
      Left            =   2273
      TabIndex        =   7
      Top             =   401
      Width           =   135
   End
   Begin VB.Label lblDrums 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   5
      Left            =   2513
      TabIndex        =   6
      Top             =   281
      Width           =   135
   End
   Begin VB.Label lblDrums 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   4
      Left            =   2393
      TabIndex        =   5
      Top             =   281
      Width           =   135
   End
   Begin VB.Label lblDrums 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   3
      Left            =   2273
      TabIndex        =   4
      Top             =   281
      Width           =   135
   End
   Begin VB.Label lblDrums 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   2
      Left            =   2513
      TabIndex        =   3
      Top             =   161
      Width           =   135
   End
   Begin VB.Label lblDrums 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   1
      Left            =   2393
      TabIndex        =   2
      Top             =   161
      Width           =   135
   End
   Begin VB.Label lblDrums 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   0
      Left            =   2273
      TabIndex        =   1
      Top             =   161
      Width           =   135
   End
   Begin VB.Label lblInstrument 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   293
      TabIndex        =   0
      Top             =   191
      Width           =   1890
   End
   Begin VB.Shape shpDrums 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   11
      Left            =   2513
      Shape           =   3  'Circle
      Top             =   521
      Width           =   135
   End
   Begin VB.Shape shpDrums 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   10
      Left            =   2393
      Shape           =   3  'Circle
      Top             =   521
      Width           =   135
   End
   Begin VB.Shape shpDrums 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   9
      Left            =   2273
      Shape           =   3  'Circle
      Top             =   521
      Width           =   135
   End
   Begin VB.Shape shpDrums 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   8
      Left            =   2513
      Shape           =   3  'Circle
      Top             =   401
      Width           =   135
   End
   Begin VB.Shape shpDrums 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   7
      Left            =   2393
      Shape           =   3  'Circle
      Top             =   401
      Width           =   135
   End
   Begin VB.Shape shpDrums 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   6
      Left            =   2273
      Shape           =   3  'Circle
      Top             =   401
      Width           =   135
   End
   Begin VB.Shape shpDrums 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   5
      Left            =   2513
      Shape           =   3  'Circle
      Top             =   281
      Width           =   135
   End
   Begin VB.Shape shpDrums 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   4
      Left            =   2393
      Shape           =   3  'Circle
      Top             =   281
      Width           =   135
   End
   Begin VB.Shape shpDrums 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   3
      Left            =   2273
      Shape           =   3  'Circle
      Top             =   281
      Width           =   135
   End
   Begin VB.Shape shpDrums 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   2513
      Shape           =   3  'Circle
      Top             =   161
      Width           =   135
   End
   Begin VB.Shape shpDrums 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   1
      Left            =   2393
      Shape           =   3  'Circle
      Top             =   161
      Width           =   135
   End
   Begin VB.Shape shpDrums 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   2273
      Shape           =   3  'Circle
      Top             =   161
      Width           =   135
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   34
      Left            =   4965
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   32
      Left            =   4725
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   29
      Left            =   4365
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   27
      Left            =   4005
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   25
      Left            =   3765
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   24
      Left            =   3645
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   22
      Left            =   3285
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   20
      Left            =   3045
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   18
      Left            =   2805
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   17
      Left            =   2685
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   6
      Left            =   1125
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   13
      Left            =   2085
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   10
      Left            =   1605
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   8
      Left            =   1365
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   3
      Left            =   645
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   1
      Left            =   405
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   16
      Left            =   2445
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   14
      Left            =   2205
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   12
      Left            =   1965
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   11
      Left            =   1725
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   9
      Left            =   1485
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   7
      Left            =   1245
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   5
      Left            =   1005
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   4
      Left            =   765
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   2
      Left            =   525
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   0
      Left            =   285
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   21
      Left            =   3165
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   19
      Left            =   2925
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   23
      Left            =   3405
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   26
      Left            =   3885
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   28
      Left            =   4125
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   31
      Left            =   4605
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   33
      Left            =   4845
      Top             =   769
      Width           =   255
   End
   Begin VB.Shape shpKeys 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   855
      Index           =   35
      Left            =   5085
      Top             =   769
      Width           =   255
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   1328
      TabIndex        =   15
      Top             =   506
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   0
      Picture         =   "frmPiano.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5640
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Piano"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuPiano 
         Caption         =   "Piano"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Drum"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuDrum 
         Caption         =   "Drum"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Beat"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnuBeat 
         Caption         =   "Beat"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Device"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnuDevice 
         Caption         =   "Device"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Key"
      Index           =   4
      Visible         =   0   'False
      Begin VB.Menu mnuBaseKey 
         Caption         =   "C"
         Index           =   0
      End
      Begin VB.Menu mnuBaseKey 
         Caption         =   "C#"
         Index           =   1
      End
      Begin VB.Menu mnuBaseKey 
         Caption         =   "D"
         Index           =   2
      End
      Begin VB.Menu mnuBaseKey 
         Caption         =   "D#"
         Index           =   3
      End
      Begin VB.Menu mnuBaseKey 
         Caption         =   "E"
         Index           =   4
      End
      Begin VB.Menu mnuBaseKey 
         Caption         =   "F"
         Index           =   5
      End
      Begin VB.Menu mnuBaseKey 
         Caption         =   "F#"
         Index           =   6
      End
      Begin VB.Menu mnuBaseKey 
         Caption         =   "G"
         Index           =   7
      End
      Begin VB.Menu mnuBaseKey 
         Caption         =   "G#"
         Index           =   8
      End
      Begin VB.Menu mnuBaseKey 
         Caption         =   "A"
         Index           =   9
      End
      Begin VB.Menu mnuBaseKey 
         Caption         =   "A#"
         Index           =   10
      End
      Begin VB.Menu mnuBaseKey 
         Caption         =   "B"
         Index           =   11
      End
   End
End
Attribute VB_Name = "frmPiano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long

Dim KeyState(0 To 255) As Byte
Dim KeyState1(0 To 255) As Byte
Dim LastKey As Integer
Dim Keys
Dim DrumKeys
Dim Drums

Dim Banks(0 To 4) As Integer
Dim BaseKey As Integer

Dim InstrSelect As Integer
Dim DrumSelect As Integer
Dim Volume As Integer
Dim Record As Boolean

Dim BeatSelect As Integer
Dim BeatPtr As Integer
Dim BeatMode As Integer
Dim BeatNext As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If LastKey <> KeyCode Then
        LastKey = KeyCode
        If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF4 Then
            lblBank_MouseDown KeyCode - vbKeyF1, vbLeftButton, 0, 0, 0
            lblBank(KeyCode - vbKeyF1).BackColor = &H808080
        End If
        
        If KeyCode >= vbKeyF5 And KeyCode <= vbKeyF9 Then
            lblBeatCtrl_Click KeyCode - vbKeyF5
        End If
        
        KeyState1(KeyCode) = 1
        CheckKeys
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    LastKey = 0
    KeyState1(KeyCode) = 0
    CheckKeys
End Sub

Private Sub Form_Load()
    Dim hrgn As Long
    hrgn = CreateRoundRectRgn(1, 1, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 50, 30)
    SetWindowRgn Me.hwnd, hrgn, True
    
    'prepare keyboard keys
    Keys = Array(90, 83, 88, 68, 67, 86, 71, 66, 72, 78, 74, 77, 188, 76, 190, 186, 191 _
        , 81, 50, 87, 51, 69, 52, 82, 84, 54, 89, 55, 85, 73, 57, 79, 48, 80, 189, 219)
    'Keys = Array(44, 31, 45, 32, 46, 47, 34, 48, 35, 49, 36, 50, 51, 38, 52, 39, 53 _
        , 16, 3, 17, 4, 18, 5, 19, 20, 7, 21, 8, 22, 23, 9, 24, 10, 25, 11, 26)
    
    DrumKeys = Array(111, 106, 109, 103, 104, 105, 100, 101, 102, 97, 98, 99)
    'DrumKeys = Array(53, 55, 74, 71, 72, 73, 75, 76, 77, 79, 80, 81)
    Drums = Array(49, 57, 51, 45, 47, 48, 46, 42, 44, 37, 36, 40)
    
    'load config
    On Error GoTo SkipLoadConfig
    Open App.Path & "\config.dat" For Input As #1
    Input #1, DeviceID
    Input #1, Banks(0), Banks(1), Banks(2), Banks(3)
    Input #1, Drums(0), Drums(1), Drums(2), Drums(3), Drums(4), Drums(5), Drums(6), Drums(7), Drums(8), Drums(9), Drums(10), Drums(11)

SkipLoadConfig:
    Close #1

    Dim MidiCaps As MIDIOUTCAPS
    MidiOpen
    
    'prepare device
    N = VBMIDI.midiOutGetNumDevs
    For i = 0 To N
        If i > 0 Then Load mnuDevice(i)
        midiOutGetDevCaps i, MidiCaps, Len(MidiCaps)
        mnuDevice(i).Caption = MidiCaps.szPname
    Next
    
    'prepare beats
    LoadBeat
    For i = 0 To UBound(Beats)
        If i > 0 Then Load mnuBeat(i)
        mnuBeat(i).Caption = Beats(i).Name
    Next
    mnuBeat_Click 0
    BeatMode = -1
    BeatNext = -2
    
    'load instrument
    Open App.Path & "\instrument.txt" For Input As #1
    For i = 0 To 127
        Line Input #1, S
        t = Split(S, vbTab)
        If i > 0 Then Load mnuPiano(i)
        mnuPiano(i).Caption = t(1)
    Next
    Line Input #1, S
    For i = 35 To 66
        Line Input #1, S
        t = Split(S, vbTab)
        If i > 35 Then Load mnuDrum(i - 35)
        mnuDrum(i - 35).Caption = t(1)
    Next
    Close #1
    Volume = 127
    mnuPiano_Click Banks(0)
End Sub

Private Sub Form_Resize()
    Height = Image1.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MidiClose
    
    'save config
    On Error GoTo SkipSaveConfig
    Open App.Path & "\config.dat" For Output As #1
    Print #1, DeviceID
    Print #1, Banks(0), Banks(1), Banks(2), Banks(3)
    Print #1, Drums(0), Drums(1), Drums(2), Drums(3), Drums(4), Drums(5), Drums(6), Drums(7), Drums(8), Drums(9), Drums(10), Drums(11)
    Close #1

SkipSaveConfig:
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub lblBank_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblBank(Index).BackColor = vbYellow
    If Button = vbRightButton Then
        mnuMain(0).Tag = CStr(Index)
        PopupMenu mnuMain(0)
        mnuMain(0).Tag = ""
        lblBank(Index).BackColor = &H808080
    Else
        mnuPiano_Click Banks(Index)
    End If
End Sub

Private Sub lblBank_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblBank(Index).BackColor = &H808080
End Sub

Private Sub lblBaseKey_Click()
    PopupMenu mnuMain(4)
End Sub

Private Sub lblBeat_Click()
    lblBeat.BackColor = vbYellow
    PopupMenu mnuMain(2)
    lblBeat.BackColor = &H808080
End Sub

Private Sub lblBeatCtrl_Click(Index As Integer)
    Select Case Index
        Case 0: BeatNext = 0
        Case 1: BeatNext = 2
        Case 2: BeatNext = 3
        Case 3: BeatNext = 4
        Case 4: BeatNext = -1
    End Select
    If Index < 4 And BeatMode < 0 Then BeatMode = BeatNext: BeatNext = -2
End Sub

Private Sub lblBeatCtrl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblBeatCtrl(Index).BackColor = vbYellow
End Sub

Private Sub lblBeatCtrl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblBeatCtrl(Index).BackColor = &H808080
End Sub

Private Sub lblDevice_Click()
    lblDevice.ForeColor = vbYellow
    PopupMenu mnuMain(3)
    lblDevice.ForeColor = &HC0C0C0
End Sub

Private Sub lblDrums_Click(Index As Integer)
    DrumSelect = Index
    PopupMenu mnuMain(1)
End Sub

Private Sub lblExit_Click()
    Unload Me
End Sub

Private Sub lblInstrument_Click()
    lblInstrument.BackColor = vbYellow
    PopupMenu mnuMain(0)
    lblInstrument.BackColor = &H808080
End Sub

Private Sub lblRate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblRate.Tag = "1"
    lblRate_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblRate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblRate.Tag = "1" Then
        Y = Y - lblRatePtr.Height / 2
        If Y < 0 Then Y = 0
        If Y > lblRate.Height - lblRatePtr.Height Then Y = lblRate.Height - lblRatePtr.Height
        lblRatePtr.Top = lblRate.Top + Y
        tmrBeat.Interval = 100 * (Y / (lblRate.Height - lblRatePtr.Height)) + Val(lblRatePtr.Tag)
    End If
End Sub

Private Sub lblRate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblRate.Tag = "0"
End Sub

Private Sub lblVolume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblVolume.Tag = "1" Then
        X = X - lblVolumePtr.Width / 2
        If X < 0 Then X = 0
        If X > lblVolume.Width - lblVolumePtr.Width Then X = lblVolume.Width - lblVolumePtr.Width
        lblVolumePtr.Left = lblVolume.Left + X
        Volume = 127 * (X / (lblVolume.Width - lblVolumePtr.Width))
    End If
End Sub

Private Sub lblVolume_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblVolume.Tag = "1"
    lblVolume_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblVolume_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblVolume.Tag = "0"
End Sub

Private Sub mnuBaseKey_Click(Index As Integer)
    BaseKey = Index
    lblBaseKey = mnuBaseKey(Index).Caption
    MidiClose
    MidiOpen
End Sub

Private Sub mnuBeat_Click(Index As Integer)
    BeatSelect = Index
    BeatPtr = 0
    BeatMode = 0
    lblBeat.Caption = mnuBeat(Index).Caption
    lblRatePtr.Tag = CStr(Beats(Index).Rate)
    tmrBeat.Interval = Beats(Index).Rate
End Sub

Private Sub mnuDevice_Click(Index As Integer)
    MidiClose
    DeviceID = Index
    MidiOpen
    mnuPiano_Click InstrSelect
End Sub

Private Sub mnuDrum_Click(Index As Integer)
    Drums(DrumSelect) = Index + 35
End Sub

Private Sub mnuPiano_Click(Index As Integer)
    InstrSelect = Index
    If mnuMain(0).Tag <> "" Then Banks(Val(mnuMain(0).Tag)) = Index
    
    ChangeInstrument 0, Index
    lblInstrument.Caption = mnuPiano(Index).Caption
End Sub

Public Sub CheckKeys()
    For i = 0 To UBound(Keys)
        N = Keys(i)
        If KeyState(N) <> KeyState1(N) Then
            KeyState(N) = KeyState1(N)
            Debug.Print N, KeyState(N)
            If KeyState(N) = 1 Then
                shpKeys(i).BackColor = IIf(shpKeys(i).Height < 500, RGB(100, 100, 100), RGB(200, 200, 200))
                StartNote 0, i + BaseKey + 48, Volume
            Else
                If shpKeys(i).Height < 500 Then
                    shpKeys(i).BackColor = vbBlack
                Else
                    shpKeys(i).BackColor = vbWhite
                End If
                StopNote 0, i + BaseKey + 48
            End If
        End If
    Next
    
    For i = 0 To UBound(DrumKeys)
        N = DrumKeys(i)
        If KeyState(N) <> KeyState1(N) Then
            KeyState(N) = KeyState1(N)
            If KeyState(N) = 1 Then
                shpDrums(i).BackColor = &H808080
                StartNote 9, Drums(i), 127
            Else
                shpDrums(i).BackColor = &HC0C0C0
                StopNote 9, Drums(i)
            End If
        End If
    Next
End Sub

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Private Sub tmrBeat_Timer()
    If BeatMode >= 0 Then
        t = Beats(BeatSelect).Data
        t = Split(t, ";")(BeatMode)
        t = Split(t, "|")
        N = UBound(t)
        t = t(BeatPtr)
        t = Split(t, ",")
        
        For i = 0 To UBound(t)
            StartNote 9, t(i), Volume
        Next
        BeatPtr = BeatPtr + 1
        If BeatPtr > N Then
            If BeatNext < -1 Then
                Select Case BeatMode
                    Case 0, 2, 3: BeatMode = 1
                    Case 4: BeatMode = -1
                End Select
            Else
                BeatMode = BeatNext
                BeatNext = -2
            End If
            BeatPtr = 0
        End If
    End If
End Sub
