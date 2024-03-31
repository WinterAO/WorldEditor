VERSION 5.00
Begin VB.Form frmAcercade 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acerca del WorldEditor"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3750
      Left            =   0
      ScaleHeight     =   3750
      ScaleWidth      =   7500
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   7500
   End
   Begin WinterMapEditor.lvButtons_H LvBRepositorioOficial 
      Height          =   480
      Index           =   0
      Left            =   2280
      TabIndex        =   5
      ToolTipText     =   "Cerrar"
      Top             =   7980
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   847
      Caption         =   "Repositorio oficial"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      ImgSize         =   32
      cBack           =   16777215
   End
   Begin VB.Label lblHowellNeoSefirot 
      BackStyle       =   0  'Transparent
      Caption         =   "- Howell, NeoSefirot y Sensui por sus consejos, criticas y tests. Ayudaron a hacer una herramienta mas completa y comoda."
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   750
      TabIndex        =   4
      Top             =   5610
      Width           =   5445
   End
   Begin VB.Label lblAgradecimientosEspeciales 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agradecimientos Especiales:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   1710
      TabIndex        =   3
      Top             =   4470
      Width           =   3480
   End
   Begin VB.Label lblAgradecimientoEspecial 
      BackStyle       =   0  'Transparent
      Caption         =   "- Hide (^[GS]^) por su colaboracion, y codigos reutilizados de su WorldEditor de Argentum Online"
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   780
      TabIndex        =   2
      Top             =   5040
      Width           =   5445
   End
   Begin VB.Label lblLorwik 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desarrollado por Lorwik para WinterAO y Argentum Online - 2020 / 2024"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   7065
   End
   Begin VB.Label lblDescripcion 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAcercade.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   390
      TabIndex        =   0
      Top             =   6210
      Width           =   6855
   End
End
Attribute VB_Name = "frmAcercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    picLogo.Picture = LoadPicture(IniPath & INITDIR & "WorldEditor.jpg")
End Sub

Private Sub LvBRepositorioOficial_Click(Index As Integer)
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'***************************************************

    Call ShellExecute(0, "Open", "https://github.com/WinterAO/WorldEditor", "", App.Path, SW_SHOWNORMAL)
End Sub
