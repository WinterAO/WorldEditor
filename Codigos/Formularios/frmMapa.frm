VERSION 5.00
Begin VB.Form frmMapa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mapa - Click Izq. Renderiza / Click Der. Guarda render"
   ClientHeight    =   16485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16500
   ClipControls    =   0   'False
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
   ScaleHeight     =   1099
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMapa 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H8000000B&
      Height          =   16500
      Left            =   0
      ScaleHeight     =   1100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1100
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16500
      Begin VB.Shape ApuntadorRadar 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   600
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub picMapa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        Call DibujarMinimapa
    
    ElseIf Button = 2 Then
        Call AddtoRichTextBox(frmConsola.StatTxt, "Guardando Minimapa...", 255, 255, 255)
        'Call frmRenderer.Capturar_Imagen(frmMapa.picMapa, frmMapa.picMapa)
        'Call SavePicture(frmMapa.picMapa, App.Path & "\Render\Minimapa\" & NumMap_Save & ".bmp")
        Call AddtoRichTextBox(frmConsola.StatTxt, "Minimapa guardado.", 0, 255, 0)
    End If
End Sub
