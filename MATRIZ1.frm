VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6612
   ClientLeft      =   3420
   ClientTop       =   3036
   ClientWidth     =   14880
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   6612
   ScaleWidth      =   14880
   Begin VB.CommandButton BtnNTSPhase 
      Caption         =   "P"
      Height          =   252
      Index           =   0
      Left            =   3120
      TabIndex        =   13
      Top             =   1200
      Width           =   252
   End
   Begin VB.HScrollBar hBLUE 
      Height          =   252
      Left            =   10920
      TabIndex        =   12
      Top             =   840
      Width           =   1452
   End
   Begin VB.HScrollBar hGREN 
      Height          =   252
      Left            =   10920
      TabIndex        =   11
      Top             =   600
      Width           =   1452
   End
   Begin VB.CommandButton YIQ 
      Caption         =   "Y"
      Height          =   252
      Index           =   0
      Left            =   3120
      TabIndex        =   10
      Top             =   840
      Width           =   252
   End
   Begin VB.HScrollBar hRED 
      Height          =   252
      Left            =   10920
      TabIndex        =   9
      Top             =   360
      Width           =   1452
   End
   Begin VB.TextBox Tetan 
      Alignment       =   2  'Center
      Height          =   288
      Index           =   0
      Left            =   2520
      TabIndex        =   8
      Text            =   "0"
      Top             =   1200
      Width           =   500
   End
   Begin VB.TextBox Cn 
      Alignment       =   2  'Center
      Height          =   288
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Text            =   "0"
      Top             =   1200
      Width           =   500
   End
   Begin VB.TextBox Q 
      Alignment       =   2  'Center
      Height          =   288
      Index           =   0
      Left            =   2520
      TabIndex        =   6
      Text            =   "0"
      Top             =   840
      Width           =   500
   End
   Begin VB.TextBox I 
      Alignment       =   2  'Center
      Height          =   288
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Text            =   "0"
      Top             =   840
      Width           =   500
   End
   Begin VB.TextBox Yn 
      Alignment       =   2  'Center
      Height          =   288
      Index           =   0
      Left            =   1320
      TabIndex        =   4
      Text            =   "0"
      Top             =   960
      Width           =   500
   End
   Begin VB.PictureBox COLOR1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   492
      Index           =   0
      Left            =   1320
      ScaleHeight     =   468
      ScaleWidth      =   1668
      TabIndex        =   3
      Top             =   1920
      Width           =   1692
   End
   Begin VB.TextBox B 
      Alignment       =   2  'Center
      Height          =   288
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Text            =   "255"
      Top             =   1560
      Width           =   500
   End
   Begin VB.TextBox G 
      Alignment       =   2  'Center
      Height          =   288
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Text            =   "255"
      Top             =   1560
      Width           =   500
   End
   Begin VB.TextBox R 
      Alignment       =   2  'Center
      Height          =   288
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Text            =   "255"
      Top             =   1560
      Width           =   500
   End
   Begin VB.Menu mnArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnAbrir 
         Caption         =   "&Abrir"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RGB_P As Boolean


Private Sub RGBtoYIQ(Index As Integer)

RGB_P = True

Dim Ya, Ia, Qa As Single
Dim Ra, Ga, Ba As Single
Dim Ca, Tetaa As Single


If Len(R(Index).Text) = 0 Or Len(G(Index).Text) = 0 Or Len(G(Index).Text) = 0 Then Exit Sub

Ra = CSng(R(Index).Text) / 255
Ga = CSng(G(Index).Text) / 255
Ba = CSng(B(Index).Text) / 255

Ya = 0.299 * Ra + 0.587 * Ga + 0.114 * Ba
Ia = 0.596 * Ra - 0.274 * Ga - 0.322 * Ba
Qa = 0.211 * Ra - 0.523 * Ga + 0.312 * Ba
Ca = Sqr(Ia * Ia + Qa * Qa)

If Qa = 0 Then Qa = 0.000001
Tetaa = Atn(Ia / Qa) * (180 / 3.14159265358979)



Yn(Index).Text = Format(Ya, "##0.000")
Q(Index).Text = Format(Qa, "##0.000")
I(Index).Text = Format(Ia, "##0.000")
Cn(Index).Text = Format(Ca, "##0.000")
Tetan(Index).Text = Format(Tetaa, "##0,00")

RGB_P = False

End Sub


Private Sub YIQtoRGB(Index As Integer)

RGB_P = True

Dim Ya, Ia, Qa As Single
Dim Ra, Ga, Ba As Single
Dim Ca, Tetaa As Single


If Len(Yn(Index).Text) = 0 Or Len(I(Index).Text) = 0 Or Len(Q(Index).Text) = 0 Then Exit Sub

Ya = CSng(Yn(Index).Text)
Ia = CSng(I(Index).Text)
Qa = CSng(Q(Index).Text)

Ra = 1 * Ya + 0.956 * Ia + 0.621 * Qa
Ga = 1 * Ya - 0.272 * Ia - 0.647 * Qa
Ba = 1 * Ya - 1.106 * Ia + 1.703 * Qa

Ca = Sqr(Ia * Ia + Qa * Qa)

If Qa = 0 Then Qa = 0.000001
Tetaa = Atn(Ia / Qa) * (180 / 3.14159265358979)


R(Index).Text = Str$(Round(Ra * 255, 0))
G(Index).Text = Str$(Round(Ga * 255, 0))
B(Index).Text = Str$(Round(Ba * 255, 0))
Cn(Index).Text = Format(Ca, "##0.000")
Tetan(Index).Text = Format(Tetaa, "##0,00")

RGB_P = False

End Sub


Private Sub NTSCPhase_RGB(Index As Integer)

RGB_P = True

Dim Ya, Ia, Qa As Single
Dim Ra, Ga, Ba As Single
Dim Ca, Tetaa As Single


If Len(Yn(Index).Text) = 0 Or Len(Cn(Index).Text) = 0 Or Len(Tetan(Index).Text) = 0 Then Exit Sub

Ca = CSng(Cn(Index).Text)
Tetaa = CSng(Tetan(Index).Text)
Ya = CSng(Yn(Index).Text)

Tetaa = Tetaa * 3.14159265358979 / 180


Ia = Ca * Math.Sin(Tetaa)
Qa = Ca * Math.Cos(Tetaa)



Ra = 1 * Ya + 0.956 * Ia + 0.621 * Qa
Ga = 1 * Ya - 0.272 * Ia - 0.647 * Qa
Ba = 1 * Ya - 1.106 * Ia + 1.703 * Qa

R(Index).Text = Str$(Round(Ra * 255, 0))
G(Index).Text = Str$(Round(Ga * 255, 0))
B(Index).Text = Str$(Round(Ba * 255, 0))
Q(Index).Text = Format(Qa, "##0.000")
I(Index).Text = Format(Ia, "##0.000")

RGB_P = False

End Sub

Private Sub B_Change(Index As Integer)
If RGB_P Then Exit Sub
 If Len(B(Index).Text) > 0 Then
     If Int(B(Index).Text) > 255 Then
       B(Index).Text = 255
     End If
 End If
COLOR1_Click (Index)
End Sub

Private Sub B_KeyPress(Index As Integer, KeyAscii As Integer)

 If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
 End If

 

End Sub




Private Sub BtnNTSPhase_Click(Index As Integer)
Dim Rx, Gx, Bx As Integer


If Len(Cn(Index).Text) = 0 Or Len(Tetan(Index).Text) = 0 Then Exit Sub
If Cn(Index).Text = "0" And Tetan(Index).Text = 0 Then Exit Sub

NTSCPhase_RGB (Index)

Rx = Int(R(Index).Text)
Gx = Int(G(Index).Text)
Bx = Int(B(Index).Text)

COLOR1(Index).BackColor = RGB(Rx, Gx, Bx)

End Sub

Private Sub COLOR1_Click(Index As Integer)

Dim Rx, Gx, Bx As Integer

On erro GoTo fin

If Len(R(Index).Text) = 0 Or Len(G(Index).Text) = 0 Or Len(B(Index).Text) = 0 Then Exit Sub

Rx = Int(R(Index).Text)
Gx = Int(G(Index).Text)
Bx = Int(B(Index).Text)

COLOR1(Index).BackColor = RGB(Rx, Gx, Bx)
RGBtoYIQ (Index)

fin:

End Sub

Private Sub G_Change(Index As Integer)
If RGB_P Then Exit Sub
 If Len(G(Index).Text) > 0 Then
     If Int(G(Index).Text) > 255 Then
       G(Index).Text = 255
     End If
 End If
COLOR1_Click (Index)
End Sub

Private Sub G_KeyPress(Index As Integer, KeyAscii As Integer)

 If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
 End If

End Sub



Private Sub I_Change(Index As Integer)
 
 Rem Ia = 0.596 * Ra - 0.274 * Ga - 0.322 * Ba
 If RGB_P Then Exit Sub
  
 If InStr(I(Index).Text, ",,") > 1 Then
     I(Index).Text = "1"
     Exit Sub
 End If
 
   
 If InStr(I(Index).Text, "--") > 1 Then
     I(Index).Text = "1"
     Exit Sub
 End If
 
 If I(Index).Text = "-" Then
    I(Index).Text = "-0,"
 End If
 
 If I(Index).Text = "," Then
    I(Index).Text = "0,596"
 End If
 
 If Len(I(Index).Text) > 0 Then
     If CDbl(I(Index).Text) > 0.596 Then
       I(Index).Text = "0,596"
     End If
          If CDbl(I(Index).Text) < -0.322 Then
       I(Index).Text = "-0,322"
     End If
 End If
End Sub

Private Sub I_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
   If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
    If KeyAscii <> Asc(",") Then
     If KeyAscii <> Asc("-") Then
        KeyAscii = 0
     End If
    End If
  End If
End Sub


Private Sub mnAbrir_Click()
Rem http://www.recursosvisualbasic.com.ar/htm/tutoriales/control-commondialog.htm
Rem http://www.dreamincode.net/forums/topic/56171-file-handling-in-visual-basic-6-part-2-binary-file-handling/
'Titulo del CommonDialog

CommonDialog1.DialogTitle = "Seleccione el archivo gif"

'Extensión del CommonDialog
CommonDialog1.Filter = "Archivos gráficos gif|*.gif"

'Abrimos el CommonDialog
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then
   'No se ha seleccionado ningún archivo
   MsgBox "No se ha seleccionado ningún archivo", vbInformation
Else
  'Mostramos la ruta archivo seleccionado
  Image1.Picture = LoadPicture(CommonDialog1.FileName)
End If



End Sub

Private Sub Q_Change(Index As Integer)
 
Rem Qa = 0.211 * Ra - 0.523 * Ga + 0.312 * Ba
 If RGB_P Then Exit Sub
  
 If InStr(Q(Index).Text, ",,") > 1 Then
     Q(Index).Text = "1"
     Exit Sub
 End If
 
   
 If InStr(Q(Index).Text, "--") > 1 Then
     Q(Index).Text = "1"
     Exit Sub
 End If
 
 If Q(Index).Text = "-" Then
    Q(Index).Text = "-0,"
 End If
 
 If Q(Index).Text = "," Then
    Q(Index).Text = "0,211"
 End If
 
 If Len(Q(Index).Text) > 0 Then
     If CDbl(Q(Index).Text) > 0.596 Then
       Q(Index).Text = "0,211"
     End If
          If CDbl(Q(Index).Text) < -0.322 Then
       Q(Index).Text = "-0,523"
     End If
 End If
End Sub

Private Sub Q_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
   If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
    If KeyAscii <> Asc(",") Then
     If KeyAscii <> Asc("-") Then
        KeyAscii = 0
     End If
    End If
  End If
End Sub


Private Sub R_Change(Index As Integer)
 
If RGB_P Then Exit Sub
 
 If Len(R(Index).Text) > 0 Then
     If Int(R(Index).Text) > 255 Then
       R(Index).Text = 255
     End If
 End If
COLOR1_Click (Index)
End Sub

Private Sub R_KeyPress(Index As Integer, KeyAscii As Integer)

 If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
 End If

 

End Sub



Private Sub YIQ_Click(Index As Integer)

Dim Rx, Gx, Bx As Integer


If Len(Yn(Index).Text) = 0 Or Len(I(Index).Text) = 0 Or Len(Q(Index).Text) = 0 Then Exit Sub
If Yn(Index).Text = "0" And I(Index).Text = 0 And Yn(Index).Text = 0 Then Exit Sub

YIQtoRGB (Index)

Rx = Int(R(Index).Text)
Gx = Int(G(Index).Text)
Bx = Int(B(Index).Text)

COLOR1(Index).BackColor = RGB(Rx, Gx, Bx)

End Sub


Private Sub Yn_Change(Index As Integer)
 
If RGB_P Then Exit Sub
 
 Dim I As Integer
 I = InStr(Yn(Index).Text, ",,")
  
 If I > 1 Then
     Yn(Index).Text = "1"
     Exit Sub
End If

 If Yn(Index).Text = "," Then
    Yn(Index).Text = "1"
 End If

 If Len(Yn(Index).Text) > 0 Then
     If CDbl(Yn(Index).Text) > 1 Then
       Yn(Index).Text = 1
     End If
 End If
End Sub

Private Sub Yn_KeyPress(Index As Integer, KeyAscii As Integer)


Exit Sub
  
  If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
   If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
    If KeyAscii <> Asc(",") Then
        KeyAscii = 0
    End If
  End If
End Sub


