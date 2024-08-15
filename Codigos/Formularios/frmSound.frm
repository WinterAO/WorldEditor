VERSION 5.00
Begin VB.Form frmSound 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sonido"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraOpcionesDe 
      BackColor       =   &H00535353&
      Caption         =   "Opciones de Sonido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1785
      Left            =   2970
      TabIndex        =   7
      Top             =   60
      Width           =   3075
      Begin VB.HScrollBar Slider1 
         Height          =   405
         Index           =   1
         LargeChange     =   15
         Left            =   60
         Max             =   100
         SmallChange     =   2
         TabIndex        =   10
         Top             =   1260
         Width           =   2895
      End
      Begin VB.HScrollBar Slider1 
         Height          =   405
         Index           =   0
         LargeChange     =   15
         Left            =   60
         Max             =   100
         SmallChange     =   2
         TabIndex        =   8
         Top             =   450
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de sonidos ambientales:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Top             =   1050
         Width           =   2370
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de música:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   240
         Width           =   1410
      End
   End
   Begin WinterMapEditor.lvButtons_H LvBReproducir 
      Height          =   465
      Left            =   1230
      TabIndex        =   5
      Top             =   1410
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   820
      Caption         =   "Reproducir"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   0
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   65280
   End
   Begin VB.Frame FraSonido 
      BackColor       =   &H00535353&
      Caption         =   "Musica y Ambient"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2865
      Begin VB.OptionButton OptSound 
         BackColor       =   &H00535353&
         Caption         =   "Ambient"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1530
         TabIndex        =   4
         Top             =   810
         Width           =   1095
      End
      Begin VB.OptionButton OptSound 
         BackColor       =   &H00535353&
         Caption         =   "Musica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   810
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.TextBox txtPista 
         Appearance      =   0  'Flat
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
         Left            =   1620
         TabIndex        =   2
         Text            =   "1"
         Top             =   330
         Width           =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   2610
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblNumeroDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de la pista:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1395
      End
   End
   Begin WinterMapEditor.lvButtons_H LvBDetener 
      Height          =   465
      Left            =   60
      TabIndex        =   6
      Top             =   1410
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   820
      Caption         =   "Detener"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   0
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   255
   End
End
Attribute VB_Name = "frmSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    If modEngine_Audio.MusicEnabled Then
        Slider1(0).Enabled = True
        Slider1(0).value = modEngine_Audio.MusicVolume
    Else
        Slider1(0).Enabled = False
    End If
    
    If modEngine_Audio.AmbientEnabled Then
        Slider1(1).Enabled = True
        Slider1(1).value = modEngine_Audio.AmbientVolume
    Else
        Slider1(1).Enabled = False
    End If
    
End Sub

Private Sub LvBDetener_Click()
'***********************************
'Autor: Lorwik
'Fecha: 01/05/2021
'***********************************

    Call modEngine_Audio.StopMusic
    Call modEngine_Audio.StopAmbient

End Sub

Private Sub LvBReproducir_Click()
'***********************************
'Autor: Lorwik
'Fecha: 01/05/2021
'***********************************

    If Not IsNumeric(txtPista.Text) Then
        MsgBox "Solo valores numericos."
        Exit Sub

    End If
    
    If Val(txtPista.Text) < 1 Then
        MsgBox "Valor invalido."
        Exit Sub
    End If
    
    If OptSound(0).value Then 'Musica
        Call modEngine_Audio.PlayMusic(Val(txtPista.Text) & ".mp3")
    
    Else 'Ambient
        Call modEngine_Audio.PlayAmbient(Val(txtPista.Text) & ".wav")
    
    End If
    
End Sub

Private Sub OptSound_Click(Index As Integer)
'***********************************
'Autor: Lorwik
'Fecha: 01/05/2021
'***********************************

    Call LvBDetener_Click
End Sub

Private Sub Slider1_Change(Index As Integer)
    Select Case Index
        Case 0
            modEngine_Audio.MusicVolume = Slider1(0).value
        Case 1
            modEngine_Audio.EffectVolume = Slider1(1).value
    End Select
End Sub

Private Sub Slider1_Scroll(Index As Integer)
    Select Case Index
        Case 0
            modEngine_Audio.MusicVolume = Slider1(0).value
        Case 1
            modEngine_Audio.EffectVolume = Slider1(1).value
    End Select
End Sub

