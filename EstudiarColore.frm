VERSION 5.00
Begin VB.Form fmESTUDIO 
   Caption         =   "Form1"
   ClientHeight    =   9360
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   14988
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   14988
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frSystem 
      Caption         =   "TV Sistema"
      Height          =   1332
      Left            =   120
      TabIndex        =   21
      Top             =   4080
      Width           =   2052
      Begin VB.OptionButton opTVSystem 
         Caption         =   "NTSC"
         Height          =   492
         Index           =   0
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   972
      End
      Begin VB.OptionButton opTVSystem 
         Caption         =   "PAL"
         Height          =   492
         Index           =   1
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Width           =   972
      End
   End
   Begin VB.CheckBox chLineas 
      Caption         =   "Dibujar lienas"
      Height          =   492
      Left            =   720
      TabIndex        =   19
      Top             =   3120
      Width           =   1572
   End
   Begin VB.PictureBox picVECTOR 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   9012
      Left            =   2520
      MousePointer    =   2  'Cross
      ScaleHeight     =   8964
      ScaleWidth      =   11484
      TabIndex        =   0
      Top             =   120
      Width           =   11532
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X,Y:"
      Height          =   192
      Left            =   720
      TabIndex        =   20
      Top             =   2520
      Width           =   276
   End
   Begin VB.Label lbColor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   612
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   732
   End
   Begin VB.Label lbXY 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,0"
      Height          =   252
      Left            =   1200
      TabIndex        =   17
      Top             =   2520
      Width           =   972
   End
   Begin VB.Label TTTT 
      Alignment       =   2  'Center
      Caption         =   "Angulo"
      Height          =   252
      Left            =   1680
      TabIndex        =   16
      Top             =   1560
      Width           =   492
   End
   Begin VB.Label lbTETA 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   252
      Left            =   1680
      TabIndex        =   15
      Top             =   1800
      Width           =   492
   End
   Begin VB.Label CCC 
      Alignment       =   2  'Center
      Caption         =   "Croma"
      Height          =   252
      Left            =   1080
      TabIndex        =   14
      Top             =   1560
      Width           =   492
   End
   Begin VB.Label lbC 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   252
      Left            =   1080
      TabIndex        =   13
      Top             =   1800
      Width           =   492
   End
   Begin VB.Label QQQ 
      Alignment       =   2  'Center
      Caption         =   "Q"
      Height          =   252
      Left            =   1680
      TabIndex        =   12
      Top             =   840
      Width           =   492
   End
   Begin VB.Label lbQ 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   252
      Left            =   1680
      TabIndex        =   11
      Top             =   1080
      Width           =   492
   End
   Begin VB.Label III 
      Alignment       =   2  'Center
      Caption         =   "I"
      Height          =   252
      Left            =   1080
      TabIndex        =   10
      Top             =   840
      Width           =   492
   End
   Begin VB.Label lbI 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   252
      Left            =   1080
      TabIndex        =   9
      Top             =   1080
      Width           =   492
   End
   Begin VB.Label LLL 
      Alignment       =   2  'Center
      Caption         =   "Luma"
      Height          =   252
      Left            =   480
      TabIndex        =   8
      Top             =   840
      Width           =   492
   End
   Begin VB.Label lbLuma 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   252
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   492
   End
   Begin VB.Label BBB 
      Alignment       =   2  'Center
      Caption         =   "Azul"
      Height          =   252
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   492
   End
   Begin VB.Label lbB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   252
      Left            =   1680
      TabIndex        =   5
      Top             =   360
      Width           =   492
   End
   Begin VB.Label GGG 
      Alignment       =   2  'Center
      Caption         =   "Verde"
      Height          =   252
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   492
   End
   Begin VB.Label lbG 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   252
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   492
   End
   Begin VB.Label RRR 
      Alignment       =   2  'Center
      Caption         =   "Rojo"
      Height          =   252
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   492
   End
   Begin VB.Label lbR 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   252
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   492
   End
End
Attribute VB_Name = "fmESTUDIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long


Private Sub chLineas_Click()

If chLineas.Value Then
  
  
  picVECTOR.DrawWidth = 2

  picVECTOR.Line (picVECTOR.Width / 2, 0)-(picVECTOR.Width / 2, picVECTOR.Height)
  picVECTOR.Line (0, picVECTOR.Height / 2)-(picVECTOR.Width, picVECTOR.Height / 2)

  picVECTOR.CurrentX = picVECTOR.Width / 2
  picVECTOR.CurrentY = picVECTOR.Height / 2
  picVECTOR.Print "0"

Else
 picVECTOR.Cls
 
End If
  

End Sub


Private Sub Form_Load()
 
 picVECTOR.Picture = LoadPicture(App.Path + "\" + "NTSC_POLAR29.bmp")
 picVECTOR.AutoRedraw = False
 
End Sub

Private Sub opTVSystem_Click(Index As Integer)

If Index = 0 Then
 

Else
  III.Caption = "Y"
  QQQ.Caption = "Q"
End If
  III.Caption = "Y"
  QQQ.Caption = "Q"
End Sub

Private Sub picVECTOR_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)


Dim R As Integer
Dim G As Integer
Dim B As Integer
Dim point As Long
Dim qX, qY As Integer

Dim Ya, IYa, QVa As Single
Dim Ra, Ga, Ba As Single
Dim Ca, Tetaa As Single

 
qX = x / Screen.TwipsPerPixelX
qY = y / Screen.TwipsPerPixelY
 
 point = GetPixel(picVECTOR.hdc, qX, qY)
 
 R = &HFF& And point
 G = (&HFF00& And point) \ 256
 B = (&HFF0000 And point) \ 65536
 
 Ra = R / 255
 Ga = G / 255
 Ba = B / 255
 
 lbColor.BackColor = RGB(R, G, B)
 
If opTVSystem(0).Value Then
 Ya = 0.299 * Ra + 0.587 * Ga + 0.114 * Ba
 IYa = 0.596 * Ra - 0.274 * Ga - 0.322 * Ba
 QVa = 0.211 * Ra - 0.523 * Ga + 0.312 * Ba
Else
 Ya = 0.2126 * Ra + 0.7152 * Ga + 0.0772 * Ba
 IYa = -0.09991 * Ra - 0.33609 * Ga + 0.436 * Ba
 QVa = 0.615 * Ra - 0.555861 * Ga - 0.05639 * Ba

End If
 
 Ca = Sqr(IYa * IYa + QVa * QVa)

If QVa = 0 Then QVa = 0.000001
If IYa = 0 Then IYa = 0.000001
Tetaa = Atn(QVa / IYa) * (180 / 3.14159265358979)
 
 lbR.Caption = CStr(R)
 lbG.Caption = CStr(G)
 lbB.Caption = CStr(B)
 lbLuma.Caption = Format(Ya, "##0.000")
 lbQ.Caption = Format(QVa, "##0.000")
 lbI.Caption = Format(IYa, "##0.000")
 lbC.Caption = Format(Ca, "##0.000")
 
 lbTETA.Caption = Format(Tetaa, "##0,00")
 
 
 

 lbXY.Caption = CStr(qX) + "," + CStr(qY)

End Sub
