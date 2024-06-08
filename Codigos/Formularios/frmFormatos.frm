VERSION 5.00
Begin VB.Form frmFormatos 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sobre los formatos de mapas soportados"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   ScaleHeight     =   4095
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WinterMapEditor.lvButtons_H LvBVolver 
      Height          =   435
      Left            =   1680
      TabIndex        =   5
      Top             =   3510
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   767
      Caption         =   "Volver"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      cBack           =   49152
   End
   Begin VB.Label lblActualmenteNo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFormatos.frx":0000
      ForeColor       =   &H008080FF&
      Height          =   825
      Left            =   150
      TabIndex        =   4
      Top             =   2640
      Width           =   5580
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFormatos.frx":00E9
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   180
      TabIndex        =   3
      Top             =   1890
      Width           =   5580
   End
   Begin VB.Label lblAlIniciar 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFormatos.frx":0193
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   210
      TabIndex        =   2
      Top             =   1230
      Width           =   5730
   End
   Begin VB.Label lblEsteEs 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFormatos.frx":026A
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5850
   End
   Begin VB.Label lblFormatosSoportados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Formatos Soportados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1410
      TabIndex        =   0
      Top             =   210
      Width           =   3075
   End
End
Attribute VB_Name = "frmFormatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LvBVolver_Click()
    Unload Me
End Sub
