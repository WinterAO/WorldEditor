VERSION 5.00
Begin VB.Form frmPropiedades 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Propiedades del Mapa"
   ClientHeight    =   2715
   ClientLeft      =   16365
   ClientTop       =   9750
   ClientWidth     =   3870
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
   ScaleHeight     =   2715
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraTamañoDel 
      BackColor       =   &H00535353&
      Caption         =   "Tamaño del Mapa"
      ForeColor       =   &H00FFFFFF&
      Height          =   1995
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   3585
      Begin WinterMapEditor.lvButtons_H OptX 
         Height          =   405
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   450
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   714
         Caption         =   "1000 x 1000"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H OptX 
         Height          =   405
         Index           =   1
         Left            =   1620
         TabIndex        =   4
         Top             =   450
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   714
         Caption         =   "100 x 100"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lblEstaFunción 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmPropiedades.frx":0000
         ForeColor       =   &H8000000E&
         Height          =   795
         Left            =   180
         TabIndex        =   1
         Top             =   1050
         Width           =   3330
      End
   End
   Begin WinterMapEditor.lvButtons_H LvBGuardarY 
      Height          =   405
      Left            =   900
      TabIndex        =   2
      Top             =   2190
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      Caption         =   "Guardar y Salir"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
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
      cBack           =   255
   End
End
Attribute VB_Name = "frmPropiedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public Sub OptX_Click(Index As Integer)
''*************************************************
''Author: Lorwik
''Last modified: 25/04/2020
''*************************************************
''Nota: Hay que cambiar muchas cosas, el engine cuando inicia hace calculos con el tamaï¿½o de los mapas
''ademas hay mas funciones que manejan estos datos, no basta con cambiar el XMax & YMax.
'
'    'Seteamos el nuevo tamaño del mapa
'    Select Case Index
'
'        Case 0
'            Call setMapSize(1000, 1000)
'
'        Case 1
'            Call setMapSize(100, 100)
'
'    End Select
'End Sub

Private Sub LvBGuardarY_Click()

    Call modMapas.NuevoMapa
    
    Unload Me
    
End Sub
