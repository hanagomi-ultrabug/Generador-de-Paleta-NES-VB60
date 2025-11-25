VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor paleta avanzado NES/FAMICON - Tesoro del saber retro Dic 2017 - beta 1"
   ClientHeight    =   8856
   ClientLeft      =   1788
   ClientTop       =   984
   ClientWidth     =   15900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8856
   ScaleWidth      =   15900
   Begin VB.Frame FrmRGBEditor 
      Caption         =   "Editor RGB"
      Height          =   1572
      Left            =   8760
      TabIndex        =   93
      Top             =   3720
      Width           =   3492
      Begin VB.HScrollBar hsBeditor 
         Height          =   252
         Left            =   600
         Max             =   255
         TabIndex        =   100
         Top             =   1200
         Width           =   1812
      End
      Begin VB.HScrollBar hsGeditor 
         Height          =   252
         Left            =   600
         Max             =   255
         TabIndex        =   97
         Top             =   840
         Width           =   1812
      End
      Begin VB.HScrollBar hsReditor 
         Height          =   252
         Left            =   600
         Max             =   255
         TabIndex        =   94
         Top             =   480
         Width           =   1812
      End
      Begin VB.Label lb7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B"
         Height          =   252
         Left            =   120
         TabIndex        =   102
         Top             =   1200
         Width           =   492
      End
      Begin VB.Label lbBeditor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   252
         Left            =   2400
         TabIndex        =   101
         Top             =   1200
         Width           =   612
      End
      Begin VB.Label LB6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         Height          =   252
         Left            =   120
         TabIndex        =   99
         Top             =   840
         Width           =   492
      End
      Begin VB.Label lbGeditor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   252
         Left            =   2400
         TabIndex        =   98
         Top             =   840
         Width           =   612
      End
      Begin VB.Label lb5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         Height          =   252
         Left            =   120
         TabIndex        =   96
         Top             =   480
         Width           =   492
      End
      Begin VB.Label lbReditor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   252
         Left            =   2400
         TabIndex        =   95
         Top             =   480
         Width           =   612
      End
   End
   Begin VB.HScrollBar hsScalara 
      Height          =   252
      LargeChange     =   10
      Left            =   2400
      Max             =   300
      Min             =   90
      TabIndex        =   90
      Top             =   8400
      Value           =   100
      Width           =   2892
   End
   Begin VB.CommandButton cmGuardar 
      Caption         =   "Guardar"
      Height          =   612
      Left            =   8760
      TabIndex        =   87
      Top             =   120
      Width           =   1092
   End
   Begin VB.Frame fmIdealPalete 
      Caption         =   "Paleta ideal"
      Height          =   2892
      Left            =   6360
      TabIndex        =   77
      Top             =   5400
      Width           =   3012
      Begin VB.CheckBox chSaturar 
         Caption         =   "Saturar"
         Height          =   372
         Left            =   360
         TabIndex        =   88
         Top             =   2280
         Width           =   852
      End
      Begin VB.CommandButton cmPaletaIdeal 
         Caption         =   "Generar"
         Height          =   372
         Left            =   1440
         TabIndex        =   86
         Top             =   2280
         Width           =   1212
      End
      Begin VB.TextBox txItetaMin 
         Height          =   288
         Left            =   1680
         TabIndex        =   82
         Text            =   "-99"
         Top             =   1800
         Width           =   732
      End
      Begin VB.TextBox txILumin 
         Height          =   288
         Left            =   1680
         TabIndex        =   81
         Text            =   "0,207307"
         Top             =   1440
         Width           =   732
      End
      Begin VB.TextBox txACroma 
         Height          =   288
         Left            =   1680
         TabIndex        =   80
         Text            =   "0,10454"
         Top             =   720
         Width           =   732
      End
      Begin VB.TextBox txILumax 
         Height          =   288
         Left            =   1680
         TabIndex        =   79
         Text            =   "0,775"
         Top             =   1080
         Width           =   732
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Angulo Inicial"
         Height          =   252
         Left            =   360
         TabIndex        =   85
         Top             =   1800
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Luma Min"
         Height          =   252
         Left            =   360
         TabIndex        =   84
         Top             =   1440
         Width           =   1332
      End
      Begin VB.Label lb1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitud Croma"
         Height          =   252
         Left            =   360
         TabIndex        =   83
         Top             =   720
         Width           =   1332
      End
      Begin VB.Label lb2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Luma Max"
         Height          =   252
         Left            =   360
         TabIndex        =   78
         Top             =   1080
         Width           =   1332
      End
   End
   Begin VB.CheckBox chkVerCoodenadas 
      Caption         =   "Mostrar valores"
      Height          =   252
      Left            =   11400
      TabIndex        =   12
      Top             =   5640
      Width           =   2052
   End
   Begin VB.Frame frmVer 
      Caption         =   "Sistema"
      Enabled         =   0   'False
      Height          =   1812
      Left            =   11400
      TabIndex        =   8
      Top             =   6000
      Width           =   2172
      Begin VB.CheckBox chVector 
         Caption         =   "Vectores YIQ/YUV"
         Height          =   192
         Left            =   240
         TabIndex        =   89
         Top             =   240
         Width           =   1812
      End
      Begin VB.OptionButton opCOLOR 
         Caption         =   "YUV"
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1092
      End
      Begin VB.OptionButton opCOLOR 
         Caption         =   "YIQ"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1092
      End
      Begin VB.OptionButton opCOLOR 
         Caption         =   "RGB"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1092
      End
   End
   Begin VB.CommandButton cmYIQPORTA 
      Caption         =   "YIQ porta papeles"
      Height          =   492
      Left            =   13080
      TabIndex        =   7
      Top             =   240
      Width           =   1572
   End
   Begin VB.CommandButton cmRGBClipboard 
      Caption         =   "RGB a Porta papeles"
      Height          =   492
      Left            =   11040
      TabIndex        =   6
      Top             =   240
      Width           =   1572
   End
   Begin VB.Frame frSystem 
      Caption         =   "TV Sistema"
      Height          =   1332
      Left            =   6240
      TabIndex        =   3
      Top             =   3840
      Width           =   2052
      Begin VB.OptionButton opTVSystem 
         Caption         =   "PAL"
         Height          =   492
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   972
      End
      Begin VB.OptionButton opTVSystem 
         Caption         =   "NTSC"
         Height          =   492
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   972
      End
   End
   Begin VB.CommandButton cmDibujar 
      Caption         =   "Dibujar"
      Height          =   612
      Left            =   9960
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin VB.PictureBox picVector2 
      BackColor       =   &H00000000&
      Height          =   4332
      Left            =   1080
      ScaleHeight     =   4284
      ScaleWidth      =   4764
      TabIndex        =   1
      Top             =   3960
      Width           =   4812
   End
   Begin VB.CommandButton cmCargar 
      Caption         =   "Cargar"
      Height          =   612
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label lb4 
      Caption         =   "Garfico de Vectores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1200
      TabIndex        =   92
      Top             =   3600
      Width           =   3012
   End
   Begin VB.Label lbEscala 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Escala: 100%"
      Height          =   252
      Left            =   1200
      TabIndex        =   91
      Top             =   8400
      Width           =   1092
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   63
      Left            =   14640
      TabIndex        =   76
      Top             =   2760
      Width           =   948
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   62
      Left            =   13680
      TabIndex        =   75
      Top             =   2760
      Width           =   948
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   61
      Left            =   12720
      TabIndex        =   74
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   60
      Left            =   11760
      TabIndex        =   73
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   59
      Left            =   10800
      TabIndex        =   72
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   58
      Left            =   9840
      TabIndex        =   71
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   57
      Left            =   8880
      TabIndex        =   70
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   56
      Left            =   7920
      TabIndex        =   69
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   55
      Left            =   6960
      TabIndex        =   68
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   54
      Left            =   6000
      TabIndex        =   67
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   53
      Left            =   5040
      TabIndex        =   66
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   52
      Left            =   4080
      TabIndex        =   65
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   51
      Left            =   3120
      TabIndex        =   64
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   50
      Left            =   2160
      TabIndex        =   63
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   49
      Left            =   1200
      TabIndex        =   62
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   48
      Left            =   240
      TabIndex        =   61
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   47
      Left            =   14640
      TabIndex        =   60
      Top             =   2160
      Width           =   948
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   46
      Left            =   13680
      TabIndex        =   59
      Top             =   2160
      Width           =   948
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   45
      Left            =   12720
      TabIndex        =   58
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   44
      Left            =   11760
      TabIndex        =   57
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   43
      Left            =   10800
      TabIndex        =   56
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   42
      Left            =   9840
      TabIndex        =   55
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   41
      Left            =   8880
      TabIndex        =   54
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   40
      Left            =   7920
      TabIndex        =   53
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   39
      Left            =   6960
      TabIndex        =   52
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   38
      Left            =   6000
      TabIndex        =   51
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   37
      Left            =   5040
      TabIndex        =   50
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   36
      Left            =   4080
      TabIndex        =   49
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   35
      Left            =   3120
      TabIndex        =   48
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   34
      Left            =   2160
      TabIndex        =   47
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   33
      Left            =   1200
      TabIndex        =   46
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   32
      Left            =   240
      TabIndex        =   45
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   31
      Left            =   14640
      TabIndex        =   44
      Top             =   1560
      Width           =   948
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   30
      Left            =   13680
      TabIndex        =   43
      Top             =   1560
      Width           =   948
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   29
      Left            =   12720
      TabIndex        =   42
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   28
      Left            =   11760
      TabIndex        =   41
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   27
      Left            =   10800
      TabIndex        =   40
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   26
      Left            =   9840
      TabIndex        =   39
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   25
      Left            =   8880
      TabIndex        =   38
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   24
      Left            =   7920
      TabIndex        =   37
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   23
      Left            =   6960
      TabIndex        =   36
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   22
      Left            =   6000
      TabIndex        =   35
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   21
      Left            =   5040
      TabIndex        =   34
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   20
      Left            =   4080
      TabIndex        =   33
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   19
      Left            =   3120
      TabIndex        =   32
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   18
      Left            =   2160
      TabIndex        =   31
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   17
      Left            =   1200
      TabIndex        =   30
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   16
      Left            =   240
      TabIndex        =   29
      Top             =   1560
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   15
      Left            =   14640
      TabIndex        =   28
      Top             =   960
      Width           =   948
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   14
      Left            =   13680
      TabIndex        =   27
      Top             =   960
      Width           =   948
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   13
      Left            =   12720
      TabIndex        =   26
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   12
      Left            =   11760
      TabIndex        =   25
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   11
      Left            =   10800
      TabIndex        =   24
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   10
      Left            =   9840
      TabIndex        =   23
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   9
      Left            =   8880
      TabIndex        =   22
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   8
      Left            =   7920
      TabIndex        =   21
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   7
      Left            =   6960
      TabIndex        =   20
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   6
      Left            =   6000
      TabIndex        =   19
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   5
      Left            =   5040
      TabIndex        =   18
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   4
      Left            =   4080
      TabIndex        =   17
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   3
      Left            =   3120
      TabIndex        =   16
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   2
      Left            =   2160
      TabIndex        =   15
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   1
      Left            =   1200
      TabIndex        =   14
      Top             =   960
      Width           =   950
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   950
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type PaletType
   R As Byte
   G As Byte
   B As Byte
End Type

Private Type YIQType
   y As Single
   I As Single
   Q As Single
   C As Single
   Teta As Double
End Type

Private Type YUVType
   y As Single
   U As Single
   V As Single
   C As Single
   Teta As Double
End Type

Dim Paleta_Load As Boolean

Private Color_actual As Integer

Private PaletaCargada(63) As PaletType
Private PaletaCargadaYIQ(63) As YIQType
Private PaletaCargadaYUV(63) As YUVType

Private Sub Printar(escala As Single, x As Single, y As Single, text As String)

Dim dx, dy As Integer

dx = picVector2.Width / 2
dy = picVector2.Height / 2

With picVector2

  .ForeColor = RGB(255, 255, 255)
  .CurrentX = x * escala * dx + dx   ' + 100
  .CurrentY = -y * escala * dy + dy  '+ 100


End With

  picVector2.Print text

End Sub


Private Sub RelletarMatrices()



Dim Ya, IYa, QVa As Single
Dim Ra, Ga, Ba As Single
Dim x, y As Integer
Dim Ca, Tetaa As Single

Dim color As Integer




For color = 1 To 63
       
   Ra = PaletaCargada(color).R / 256
   Ga = PaletaCargada(color).G / 256
   Ba = PaletaCargada(color).B / 256
 
    Ya = 0.299 * Ra + 0.587 * Ga + 0.114 * Ba
    IYa = 0.596 * Ra - 0.274 * Ga - 0.322 * Ba
    QVa = 0.211 * Ra - 0.523 * Ga + 0.312 * Ba
    
    Ca = Sqr(IYa * IYa + QVa * QVa)
    If QVa = 0 Then QVa = 0.000001
    If IYa = 0 Then IYa = 0.000001
    Tetaa = Atn(QVa / IYa) * (180 / 3.14159265358979)

  PaletaCargadaYIQ(color).y = Ya
  PaletaCargadaYIQ(color).I = IYa
  PaletaCargadaYIQ(color).Q = QVa
  PaletaCargadaYIQ(color).C = QVa
  PaletaCargadaYIQ(color).Teta = Tetaa

    Ya = 0.2126 * Ra + 0.7152 * Ga + 0.0772 * Ba
    IYa = -0.09991 * Ra - 0.33609 * Ga + 0.436 * Ba
    QVa = 0.615 * Ra - 0.555861 * Ga - 0.05639 * Ba

    Ca = Sqr(IYa * IYa + QVa * QVa)
    If QVa = 0 Then QVa = 0.000001
    If IYa = 0 Then IYa = 0.000001
    Tetaa = Atn(QVa / IYa) * (180 / 3.14159265358979)

  PaletaCargadaYUV(color).y = Ya
  PaletaCargadaYUV(color).U = IYa
  PaletaCargadaYUV(color).V = QVa
  PaletaCargadaYUV(color).C = QVa
  PaletaCargadaYUV(color).Teta = Tetaa


Next



End Sub


Private Sub chkVerCoodenadas_Click()

frmVer.Enabled = chkVerCoodenadas.Value

Dim color As Integer


If chkVerCoodenadas.Value = False Then
 For color = 0 To 63
   lbColor(color).Caption = ""
 Next color
End If


End Sub

Private Sub cmCargar_Click()


    Dim f As Long
    Dim register As Integer
    
    
    f = FreeFile()

    Open App.Path + "\" + "ntscpalette.pal" For Binary As #f
     Get #f, , PaletaCargada
    Close #f

 For register = 0 To 63
    lbColor(register).BackColor = RGB(PaletaCargada(register).R, PaletaCargada(register).G, PaletaCargada(register).B)
 Next
 RelletarMatrices

 Paleta_Load = True

End Sub

Private Sub cmDibujar_Click()


If Paleta_Load = False Then Exit Sub

Set picVector2.Picture = Nothing
Dim wx, wy As Integer
Dim ex, ey As Single
Dim dx, dy As Integer

Dim Ya, IYa, QVa As Single
Dim Ra, Ga, Ba As Single
Dim x, y As Integer
Dim Ca, Tetaa As Single
Dim escala As Single

Dim color As Integer

escala = hsScalara.Value / 100

wx = picVector2.Width
wy = picVector2.Height
dx = wx / 2
dy = wy / 2

picVector2.AutoRedraw = True



    
    picVector2.DrawWidth = 1
    picVector2.Line (0, picVector2.Height / 2)-(picVector2.Width, picVector2.Height / 2), RGB(256, 256, 256)
    picVector2.Line (picVector2.Width / 2, 0)-(picVector2.Width / 2, picVector2.Height), RGB(256, 256, 256)
        
    Printar escala, 0.01, 0.01, "0"
    Printar escala, -0.998, 0.01, "-1"
    Printar escala, 0.95, 0.01, "1"
    Printar escala, 0.01, 0.998, "1"
    Printar escala, 0.01, -0.9, "-1"
        
Printar escala, -0.501, 0.01, "-0.5"
    Printar escala, 0.501, 0.01, "0.5"
    Printar escala, 0.01, 0.5, "0.5"
    Printar escala, 0.01, -0.5, "-0.5"
    
    
  Printar escala, -0.251, 0.01, "-0.25"
    Printar escala, 0.251, 0.01, "0.25"
    Printar escala, 0.01, 0.251, "0.25"
    Printar escala, 0.01, -0.251, "-0.25"


 For color = 0 To 63
 
 If opTVSystem(0).Value Then
    Ya = PaletaCargadaYIQ(color).y
    IYa = PaletaCargadaYIQ(color).I
    QVa = PaletaCargadaYIQ(color).Q
    Ca = PaletaCargadaYIQ(color).C
  Else
    Ya = PaletaCargadaYUV(color).y
    IYa = PaletaCargadaYUV(color).U
    QVa = PaletaCargadaYUV(color).V
    Ca = PaletaCargadaYIQ(color).C
 End If
 
  x = IYa * escala * dx + dx
  y = -QVa * escala * dy + dy
  
  picVector2.DrawWidth = 15 * (Ya) + 1
  picVector2.PSet (x, y), RGB(PaletaCargada(color).R, PaletaCargada(color).G, PaletaCargada(color).B)
  
 Next

Clipboard.SetData picVector2.Image, vbCFBitmap
picVector2.AutoRedraw = False


End Sub

Private Sub cmGuardar_Click()
Dim f As Long
    Dim register As Integer
    
    
    f = FreeFile()

    Open App.Path + "\" + "ideal.pal" For Binary As #f
     Put #f, , PaletaCargada
    Close #f

 For register = 0 To 63
    lbColor(register).BackColor = RGB(PaletaCargada(register).R, PaletaCargada(register).G, PaletaCargada(register).B)
 Next
End Sub

Private Sub cmPaletaIdeal_Click()

Dim Ya, Ia, Qa As Single

Dim Ra, Ga, Ba As Single

Dim Teta As Double
Dim TetaMin As Double
Dim deltaTETA As Double

Dim Luma(3) As Single
Dim Lstep As Single

Dim CromaBase, Croma As Single

Dim x, y As Integer
Dim color As Integer

Luma(0) = CSng(txILumin.text)
Luma(3) = CSng(txILumax.text)
Lstep = (Luma(3) - Luma(0)) / 3
Luma(1) = Luma(0) + Lstep
Luma(2) = Luma(1) + Lstep

CromaBase = CSng(txACroma)


TetaMin = CDbl(txItetaMin) * 3.14159265358979 / 180
deltaTETA = 30 * 3.14159265358979 / 180

For y = 0 To 3
  
  For x = 0 To 15
      
         Select Case x
           Case 0
             Ya = Luma(y)
             Croma = 0
             Teta = 3.14159265358979 / 4
           Case 13
              Croma = 0
              Teta = 0
           Case 14, 15
               Croma = 0
               Teta = 0
               Ya = Lstep
           Case 1
             Ya = Luma(y)
             Croma = CromaBase
             Teta = TetaMin
           Case Else
               Ya = Luma(y)
               Croma = CromaBase
               Teta = Teta + deltaTETA
           End Select
          
         
        Ia = Croma * Sin(Teta)
        Qa = Croma * Cos(Teta)
        
        Ra = 1.017294 * Ya + 0.9514548 * Ia + 0.6102466 * Qa
        Ga = 1.017294 * Ya - 0.2774045 * Ia - 0.6579992 * Qa
        Ba = 1.017294 * Ya - 1.10846 * Ia + 1.6894371 * Qa
         
         
     If chSaturar Then
       If Ra < 0 Then Ra = 0
       If Ga < 0 Then Ga = 0
       If Ba < 0 Then Ba = 0
       If Ra > 1 Then Ra = 1
       If Ga > 1 Then Ga = 1
       If Ba > 1 Then Ba = 1
     
     Else
        If Ra < 0 Or Ba < 0 Or Ga < 0 Then Ra = 0: Ba = 0: Ga = 0
        If Ra > 1 Or Ba > 1 Or Ga > 1 Then Ra = 0: Ba = 0: Ga = 0
     End If
        
        PaletaCargadaYIQ(color).y = Ya
        PaletaCargadaYIQ(color).I = Ia
        PaletaCargadaYIQ(color).Q = Qa
        PaletaCargadaYIQ(color).C = Croma
        PaletaCargadaYIQ(color).Teta = Teta
        
        PaletaCargada(color).R = Ra * 255
        PaletaCargada(color).G = Ga * 255
        PaletaCargada(color).B = Ba * 255
        
        lbColor(color).BackColor = RGB(Ra * 255, Ga * 255, Ba * 255)
        
        color = color + 1
 Next x
Next y
  
End Sub

Private Sub cmRGBClipboard_Click()
Dim x, y As Integer
Dim color As Integer
Dim Cadena As String

 For y = 0 To 3
    For x = 0 To 15
      Cadena = Cadena + CStr(PaletaCargada(color).R) + "," + CStr(PaletaCargada(color).G) + "," + CStr(PaletaCargada(color).B) + ","
     color = color + 1
    Next x
 Cadena = Cadena + Chr(10)

 Next y

Clipboard.Clear
Clipboard.SetText (Cadena)

End Sub

Private Sub cmYIQPORTA_Click()


Dim x, y As Integer
Dim color As Integer
Dim Cadena As String

 For y = 0 To 3
    For x = 0 To 15
      Cadena = Cadena + CStr(PaletaCargadaYIQ(color).y) + ";" + CStr(PaletaCargadaYIQ(color).I) + ";" + CStr(PaletaCargadaYIQ(color).Q) + ";"
     color = color + 1
   Cadena = Cadena + Chr(10)
    Next x
 Next y

Clipboard.Clear
Clipboard.SetText (Cadena)


End Sub


Private Sub Form_Load()
 Paleta_Load = False
End Sub

Private Sub hsBeditor_Change()
 lbBeditor.Caption = CStr(hsBeditor.Value)
 PaletaCargada(Color_actual).B = hsBeditor.Value

lbColor(Color_actual).BackColor = RGB(PaletaCargada(Color_actual).R, PaletaCargada(Color_actual).G, PaletaCargada(Color_actual).B)

End Sub

Private Sub hsGeditor_Change()
 lbGeditor.Caption = CStr(hsGeditor.Value)
 PaletaCargada(Color_actual).G = hsGeditor.Value
 lbColor(Color_actual).BackColor = RGB(PaletaCargada(Color_actual).R, PaletaCargada(Color_actual).G, PaletaCargada(Color_actual).B)

End Sub

Private Sub hsReditor_Change()
 lbReditor.Caption = CStr(hsReditor.Value)
 PaletaCargada(Color_actual).R = hsReditor.Value
lbColor(Color_actual).BackColor = RGB(PaletaCargada(Color_actual).R, PaletaCargada(Color_actual).G, PaletaCargada(Color_actual).B)

End Sub

Private Sub hsScalara_Change()
 
If Paleta_Load Then
 lbEscala.Caption = "Escala: " + CStr(hsScalara.Value / 100) + "%"
 cmDibujar_Click
End If
End Sub

Private Sub lbColor_Click(Index As Integer)
 
Static Anterior As Integer

  Color_actual = Index
  lbColor(Anterior).BorderStyle = 1
  lbColor(Index).BorderStyle = 0
  
  hsReditor.Value = PaletaCargada(Index).R
  hsGeditor.Value = PaletaCargada(Index).G
  hsBeditor.Value = PaletaCargada(Index).B
  
  Anterior = Index

  
'If Anterior = Color_actual Then

End Sub

Private Sub opCOLOR_Click(Index As Integer)

Dim color As Integer


 For color = 0 To 63
 
  Select Case Index
    Case 0
       lbColor(color).Caption = CStr(PaletaCargada(color).R) + " " + CStr(PaletaCargada(color).G) + " " + CStr(PaletaCargada(color).B)
    Case 1
       If chVector.Value = False Then
         lbColor(color).Caption = Format(PaletaCargadaYIQ(color).y, "#0.000") + " " + Format(PaletaCargadaYIQ(color).I, "#0.000") + " " + Format(PaletaCargadaYIQ(color).Q, "#0.000")
        Else
         lbColor(color).Caption = Format(PaletaCargadaYIQ(color).y, "#0.000") + " " + Format(PaletaCargadaYIQ(color).C, "#0.000") + " " + Format(PaletaCargadaYIQ(color).Teta, "#0.000")
        End If
    Case 2
      If chVector.Value = False Then
        lbColor(color).Caption = Format(PaletaCargadaYUV(color).y, "#0.000") + " " + Format(PaletaCargadaYUV(color).U, "#0.000") + " " + Format(PaletaCargadaYUV(color).V, "#0.000")
       Else
        lbColor(color).Caption = Format(PaletaCargadaYUV(color).y, "#0.000") + " " + Format(PaletaCargadaYUV(color).C, "#0.000") + " " + Format(PaletaCargadaYUV(color).Teta, "#0.000")
       End If
    
    End Select
   
 If PaletaCargadaYIQ(color).y < 0.4 Then lbColor(color).ForeColor = RGB(255, 255, 255)
 
 Next color

End Sub

