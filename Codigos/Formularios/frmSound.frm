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
      Begin VB.HScrollBar scrAmbient 
         Height          =   225
         LargeChange     =   15
         Left            =   60
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   12
         Top             =   1470
         Width           =   2895
      End
      Begin VB.HScrollBar scrMusic 
         Height          =   225
         LargeChange     =   15
         Left            =   60
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   10
         Top             =   1020
         Width           =   2895
      End
      Begin VB.CheckBox chkop 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Sonidos Ambientales habilitado"
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
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   540
         Width           =   2805
      End
      Begin VB.CheckBox chkop 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Música habilitada"
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
         Left            =   90
         TabIndex        =   8
         Top             =   240
         Width           =   2805
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
         TabIndex        =   13
         Top             =   1260
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
         TabIndex        =   11
         Top             =   810
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

Private Sub chkop_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Lorwik
'Fecha: 01/05/2021
'Descripcion: Activa/Desactiva los sonidos
'***************************************************

    Select Case Index
    
        Case 0 'Musica
                    
            If chkop(Index).value = vbUnchecked Then
                Sound.Music_Stop
                ClientSetup.bMusic = CONST_DESHABILITADA
                scrMusic.Enabled = False
            Else
                ClientSetup.bMusic = CONST_MP3
                scrMusic.Enabled = True
            End If
        
'        Case 1 'Sonido
'
'            If chkop(Index).value = vbUnchecked Then
'                'scrAmbient.Enabled = False
'                scrVolume.Enabled = False
'                ClientSetup.bSound = 0
'            Else
'                ClientSetup.bSound = 1
'                scrVolume.Enabled = True
'            End If
            
        Case 2 'Ambiente
            
            If chkop(Index).value = vbUnchecked Then
                ClientSetup.bAmbient = 0
                Call Sound.Sound_Stop_All
            Else
                ClientSetup.bAmbient = 1
                scrAmbient.Enabled = True
                Call Sound.Ambient_Load(Sound.AmbienteActual, ClientSetup.AmbientVol)
                Call Sound.Ambient_Play
            End If
    End Select

End Sub

Private Sub Form_Load()
    'Musica
    If ClientSetup.bMusic = CONST_DESHABILITADA Then
        chkop(0).value = 0
        scrMusic.value = ClientSetup.MusicVolume
    Else
        chkop(0).value = 1
        scrMusic.value = ClientSetup.MusicVolume
    End If
    
    'Sonidos
'    If ClientSetup.bSound = 1 Then
'        chkop(1).value = vbChecked
'        chkop(3).value = IIf(ClientSetup.Invertido = True, 1, 0)
'        scrVolume.value = ClientSetup.SoundVolume
'    Else
'        'chkop(1).value = vbUnchecked
'        chkop(3).value = IIf(ClientSetup.Invertido = True, 1, 0)
'        chkop(3).Enabled = False
'        scrVolume.value = ClientSetup.SoundVolume
'        scrVolume.Enabled = False
'    End If
    
    'Ambiente
    If ClientSetup.bAmbient = 1 Then
        chkop(2).value = vbChecked
        scrAmbient.value = ClientSetup.AmbientVol
    Else
        chkop(2).value = vbUnchecked
        scrAmbient.value = ClientSetup.AmbientVol
    End If
End Sub

Private Sub LvBDetener_Click()
'***********************************
'Autor: Lorwik
'Fecha: 01/05/2021
'***********************************

    Sound.Music_Stop
    Sound.Ambient_Stop
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
    
        'Reproducimos la música del mapa
        If ClientSetup.bMusic <> CONST_DESHABILITADA Then
            If ClientSetup.bMusic <> CONST_DESHABILITADA Then
                Sound.NextMusic = Val(txtPista.Text)
                Sound.Fading = 250
            End If

        End If
    
    Else 'Ambient
        MapInfo.ambient = Val(txtPista.Text)
    
    End If
    
End Sub

Private Sub OptSound_Click(Index As Integer)
'***********************************
'Autor: Lorwik
'Fecha: 01/05/2021
'***********************************

    Call LvBDetener_Click
End Sub
