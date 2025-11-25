VERSION 5.00
Begin VB.Form PalVecto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TV CRT Generador de Color"
   ClientHeight    =   10125
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmPOLAR 
      Caption         =   "YIQ-YUV Polar"
      Height          =   492
      Left            =   720
      TabIndex        =   15
      Top             =   5400
      Width           =   1332
   End
   Begin VB.Frame fmBase 
      Caption         =   "Ajuste base"
      Height          =   5052
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   2652
      Begin VB.CheckBox chEJES 
         Caption         =   "Dibujar ejes guia"
         Height          =   372
         Left            =   480
         TabIndex        =   16
         Top             =   2880
         Width           =   1500
      End
      Begin VB.CommandButton cmGenerar 
         Caption         =   "YIQ-YUV Cartesiano"
         Height          =   612
         Left            =   480
         TabIndex        =   9
         Top             =   3240
         Width           =   1332
      End
      Begin VB.CommandButton cmStop 
         Caption         =   "STOP!!"
         Height          =   372
         Left            =   480
         TabIndex        =   8
         Top             =   4560
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.HScrollBar hsLuma 
         Height          =   216
         LargeChange     =   5
         Left            =   360
         Max             =   100
         TabIndex        =   7
         Top             =   1080
         Value           =   50
         Width           =   1452
      End
      Begin VB.HScrollBar hsScala 
         Height          =   252
         Left            =   360
         Max             =   3
         Min             =   1
         TabIndex        =   6
         Top             =   2160
         Value           =   1
         Width           =   1452
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Luma base (0% a 100%)"
         Height          =   492
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1572
      End
      Begin VB.Label lbLuma 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50%"
         Height          =   252
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   1452
      End
      Begin VB.Label lbXY 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,0"
         Height          =   252
         Left            =   480
         TabIndex        =   13
         Top             =   2520
         Width           =   852
      End
      Begin VB.Label lbColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   1320
         TabIndex        =   12
         Top             =   2520
         Width           =   252
      End
      Begin VB.Label lbScala 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   252
         Left            =   360
         TabIndex        =   11
         Top             =   1920
         Width           =   1452
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tamaño imagen 1 a 3 (x 255)"
         Height          =   492
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   1572
      End
   End
   Begin VB.Frame frSystem 
      Caption         =   "TV Sistema"
      Height          =   1332
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   2412
      Begin VB.OptionButton opTVSystem 
         Caption         =   "PAL"
         Height          =   492
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   972
      End
      Begin VB.OptionButton opTVSystem 
         Caption         =   "NTSC"
         Height          =   492
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   972
      End
   End
   Begin VB.PictureBox picVector 
      BackColor       =   &H00FFFFFF&
      Height          =   5652
      Left            =   3120
      ScaleHeight     =   5595
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   120
      Width           =   5892
   End
   Begin VB.Label lbY 
      BorderStyle     =   1  'Fixed Single
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Top             =   6600
      Width           =   2652
   End
End
Attribute VB_Name = "PalVecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PARAR  As Boolean
Public Xmin, Xmax, Ymin, Ymax As Single

Private Sub Printar(escala As Single, x As Single, y As Single, text As String)

Dim dx, dy As Integer

dx = picVector.Width / 2
dy = picVector.Height / 2

With picVector

  .ForeColor = RGB(255, 255, 255)
  .CurrentX = x * escala * dx + dx   ' + 100
  .CurrentY = -y * escala * dy + dy  '+ 100


End With

  picVector.Print text

End Sub





Private Sub cmGenerar_Click()


  hsLuma.Enabled = False
  frSystem.Enabled = False
  cmGenerar.Enabled = False
  hsScala.Enabled = False
  cmStop.Visible = True
  PARAR = False

Dim wX As Integer
Dim wY As Integer
Dim dx As Integer
Dim dy As Integer
  Dim color As Long

Dim x, y  As Integer
Dim fX, fy As Single


Dim Ya, Ia, Qa As Single

Dim Yb, Ib, Qb As Single

Dim Ra, Ga, Ba As Single

Xmin = 0: Ymin = 0: Xmax = 0: Xmin = 0

Dim escala As Single

escala = 1

wX = picVector.Width / Screen.TwipsPerPixelX
wY = picVector.Height / Screen.TwipsPerPixelY
dx = wX / 2
dy = wY / 2

fX = 1 / dx
fy = 1 / dy

Ya = hsLuma.Value / 100

 For y = 0 To wY
   For x = 0 To wX

  Qa = (-1) * (y - dy) * fy
  Ia = (x - dx) * fX
  
  lbXY.Caption = Format(Ia, "#0.00") + "," + Format(Qa, "#0.00")
  
  
If opTVSystem(0).Value Then
  'NTSC
 Ra = 1.017294 * Ya + 0.9514548 * Ia + 0.6102466 * Qa
 Ga = 1.017294 * Ya - 0.2774045 * Ia - 0.6579992 * Qa
 Ba = 1.017294 * Ya - 1.10846 * Ia + 1.6894371 * Qa
  
  
 Yb = 0.299 * Ra + 0.587 * Ga + 0.114 * Ba
 Ib = 0.596 * Ra - 0.274 * Ga - 0.322 * Ba
 Qb = 0.211 * Ra - 0.523 * Ga + 0.312 * Ba
 
 'If Yb < 0.45 Then PARAR = True
 
 lbY.Caption = "Y:" + Format(Yb, "0,00") + "  I:" + Format(Ib, "0,00") + "  Q:" + Format(Qb, "0,00")
  
  Else
 Ra = 1 * Ya - 0.0000395 * Ia + 1.139828 * Qa
 Ga = 1 * Ya - 0.3946102 * Ia - 0.5805003 * Qa
 Ba = 1 * Ya + 2.0319997 * Ia - 0.0004814 * Qa


End If

     'If Ra < 0 Then Ra = 0: PARAR = True
     'If Ba < 0 Then Ba = 0: PARAR = True
     'If Ga < 0 Then Ga = 0: PARAR = True
    
    ' If Ra > 1 Then Ra = 1: PARAR = True
    ' If Ba > 1 Then Ba = 1: PARAR = True
    ' If Ga > 1 Then Ga = 1: PARAR = True
    
    If Ra < 0 Or Ba < 0 Or Ga < 0 Then Ra = 0: Ba = 0: Ga = 0
    If Ra > 1 Or Ba > 1 Or Ga > 1 Then Ra = 0: Ba = 0: Ga = 0
    
    
    picVector.PSet (x * Screen.TwipsPerPixelX * escala, y * Screen.TwipsPerPixelX * escala), RGB(Ra * 256, Ga * 256, Ba * 256)
     lbColor.BackColor = RGB(Ra * 256, Ga * 256, Ba * 256)
  DoEvents
   
    If PARAR Then
       hsLuma.Enabled = True
        frSystem.Enabled = True
       cmStop.Visible = False
       cmGenerar.Enabled = True
       hsScala.Enabled = True
        Exit Sub
   End If
   
   Next x
 Next y


If chEJES.Value Then
 
    picVector.DrawWidth = 1
    picVector.Line (0, picVector.Height / 2)-(picVector.Width, picVector.Height / 2), RGB(256, 256, 256)
    picVector.Line (picVector.Width / 2, 0)-(picVector.Width / 2, picVector.Height), RGB(256, 256, 256)
        
        
    Printar escala, -0.98, 0.98, "Luma:" + Str(y)
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
End If

Dim Sistem As String
 If opTVSystem(0).Value Then Sistem = "NTSC" Else Sistem = "PAL"

SavePicture picVector.Image, App.Path + "\" + Sistem + CStr(hsLuma.Value) + ".bmp"

 MsgBox "Imagen generada", vbOKOnly

  hsLuma.Enabled = True
  frSystem.Enabled = True
  cmStop.Visible = False
  cmGenerar.Enabled = True
 hsScala.Enabled = True

End Sub

Private Sub cmPOLAR_Click()

  
  hsLuma.Enabled = False
  frSystem.Enabled = False
  cmPOLAR.Enabled = False
  hsScala.Enabled = False
  cmStop.Visible = True
  PARAR = False

Dim wX As Integer
Dim wY As Integer
Dim dx As Integer
Dim dy As Integer
  Dim color As Long

Dim x, y  As Integer
Dim fX, fy As Single
Dim teta As Double

Dim Ya, Ia, Qa As Single

Dim Yb, Ib, Qb As Single

Dim Ra, Ga, Ba As Single

Xmin = 0: Ymin = 0: Xmax = 0: Xmin = 0



wX = picVector.Width / Screen.TwipsPerPixelX
wY = picVector.Height / Screen.TwipsPerPixelY
dx = wX / 2
dy = wY / 2

'unidad 2pi
fX = (1) / (3.14159265) / (15 * 1)
'unidad altura
fy = (1 / wY) / (1.5 * 1)
Ya = hsLuma.Value / 100


'x es Angulo
'y es Croma
 For y = 0 To wY
   For x = 0 To wX


 teta = (x - dx) * fX
' teta = Round(teta * 3 * 3, 0) / 3 * 3
 
  Qa = (fy * y) * Cos(teta)
  Ia = (fy * y) * Sin(teta)
  
  lbXY.Caption = Format(Ia, "#0.00") + "," + Format(Qa, "#0.00")
  
  
If opTVSystem(0).Value Then
  'NTSC
 Ra = 1.017294 * Ya + 0.9514548 * Ia + 0.6102466 * Qa
 Ga = 1.017294 * Ya - 0.2774045 * Ia - 0.6579992 * Qa
 Ba = 1.017294 * Ya - 1.10846 * Ia + 1.6894371 * Qa
  
  
 Yb = 0.299 * Ra + 0.587 * Ga + 0.114 * Ba
 Ib = 0.596 * Ra - 0.274 * Ga - 0.322 * Ba
 Qb = 0.211 * Ra - 0.523 * Ga + 0.312 * Ba
 
 'If Yb < 0.45 Then PARAR = True
 
 lbY.Caption = "Y:" + Format(Yb, "0,00") + "  I:" + Format(Ib, "0,00") + "  Q:" + Format(Qb, "0,00")
  
  Else
 Ra = 1 * Ya - 0.0000395 * Ia + 1.139828 * Qa
 Ga = 1 * Ya - 0.3946102 * Ia - 0.5805003 * Qa
 Ba = 1 * Ya + 2.0319997 * Ia - 0.0004814 * Qa


End If

     'If Ra < 0 Then Ra = 0: PARAR = True
     'If Ba < 0 Then Ba = 0: PARAR = True
     'If Ga < 0 Then Ga = 0: PARAR = True
    
    ' If Ra > 1 Then Ra = 1: PARAR = True
    ' If Ba > 1 Then Ba = 1: PARAR = True
    ' If Ga > 1 Then Ga = 1: PARAR = True
    
    If Ra < 0 Or Ba < 0 Or Ga < 0 Then Ra = 0: Ba = 0: Ga = 0
    If Ra > 1 Or Ba > 1 Or Ga > 1 Then Ra = 0: Ba = 0: Ga = 0
    
    

    
    picVector.PSet (x * Screen.TwipsPerPixelX, (wY - y) * Screen.TwipsPerPixelX), RGB(Ra * 256, Ga * 256, Ba * 256)
     lbColor.BackColor = RGB(Ra * 256, Ga * 256, Ba * 256)
  DoEvents
   
    If PARAR Then
       hsLuma.Enabled = True
        frSystem.Enabled = True
       cmStop.Visible = False
       cmPOLAR.Enabled = True
       hsScala.Enabled = True
        Exit Sub
   End If
   
   Next x
 Next y
 
 
 

Dim Sistem As String
 If opTVSystem(0).Value Then Sistem = "NTSC_POLAR" Else Sistem = "PAL_POLAR"

SavePicture picVector.Image, App.Path + "\" + Sistem + CStr(hsLuma.Value) + ".bmp"

 MsgBox "Imagen generada", vbOKOnly

  hsLuma.Enabled = True
  frSystem.Enabled = True
  cmStop.Visible = False
  cmPOLAR.Enabled = True
 hsScala.Enabled = True

End Sub


Private Sub cmStop_Click()
  PARAR = True
End Sub
Private Sub Form_Load()
 
 With picVector
   .Width = 256 * Screen.TwipsPerPixelX
   .Height = 256 * Screen.TwipsPerPixelY
   .AutoRedraw = True
 End With
PARAR = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub hsLuma_Change()
  lbLuma.Caption = Str(hsLuma.Value) + "%"
End Sub

Private Sub hsScala_Change()
 
lbScala.Caption = CStr(hsScala.Value)
 
 With picVector
   .Width = 256 * Screen.TwipsPerPixelX * hsScala.Value
   .Height = 256 * Screen.TwipsPerPixelY * hsScala.Value

 End With
End Sub

Private Sub txXmin_Change()

End Sub
