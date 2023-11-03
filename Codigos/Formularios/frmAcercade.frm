VERSION 5.00
Begin VB.Form frmAcercade 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acerca del WorldEditor"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6045
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
   ScaleHeight     =   191
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   403
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblAgradecimientoEspecial 
      BackStyle       =   0  'Transparent
      Caption         =   "Agradecimiento especial para Hide (^[GS]^) por su colaboracion, y codigos reutilizados del WorldEditor Original de Argentum Online"
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   300
      TabIndex        =   2
      Top             =   2250
      Width           =   5445
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   62
      X2              =   318
      Y1              =   140
      Y2              =   140
   End
   Begin VB.Label lblLorwik 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Lorwik"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   4170
      TabIndex        =   1
      Top             =   1500
      Width           =   1425
   End
   Begin VB.Label Label1 
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
      Height          =   1605
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   6165
   End
End
Attribute VB_Name = "frmAcercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Picture = LoadPicture(IniPath & INITDIR & "WorldEditor.jpg")
End Sub
