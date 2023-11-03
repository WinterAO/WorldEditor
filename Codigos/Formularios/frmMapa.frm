VERSION 5.00
Begin VB.Form frmMapa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mapa - [Click Izq. + Shift = Renderiza / Click Der. = Guarda render]"
   ClientHeight    =   14985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15090
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
   ScaleHeight     =   999
   ScaleMode       =   0  'User
   ScaleWidth      =   999.007
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMapa 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H8000000B&
      Height          =   15000
      Left            =   0
      ScaleHeight     =   1000
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1007
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15105
      Begin VB.Shape ApuntadorRadar 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7500
         Top             =   7500
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Función BitBlt para copiar la imagen del control en un picturebox
Private Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long
      
' Recupera la imagen del área del control
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Public Sub Capturar_Imagen(Control As Control, Destino As Object)
'***********************************************
'Autor: ????
'Fecha: ????
'Descripcion: copia la imagen del control en un picturebox
'***********************************************
          
    Dim hDC             As Long
    Dim Escala_Anterior As Integer
    Dim Ancho           As Long
    Dim Alto            As Long
          
    ' Para que se mantenga la imagen por si se repinta la ventana
    Destino.AutoRedraw = True
          
    On Error Resume Next

    ' Si da error es por que el control está dentro de un Frame _
      ya que  los Frame no tiene  dicha propiedad
    Escala_Anterior = Control.Container.ScaleMode
          
    If Err.Number = 438 Then
        ' Si el control está en un Frame, convierte la escala
        Ancho = ScaleX(Control.Width, vbTwips, vbPixels)
        Alto = ScaleY(Control.Height, vbTwips, vbPixels)
    Else
        ' Si no cambia la escala del  contenedor a pixeles
        Control.Container.ScaleMode = vbPixels
        Ancho = Control.Width
        Alto = Control.Height

    End If
          
    ' limpia el error
    On Error GoTo 0

    ' Captura el área de pantalla correspondiente al control
    hDC = GetWindowDC(Control.hwnd)
    
    ' Copia esa área al picturebox
    Call BitBlt(Destino.hDC, 0, 0, 3000, 3000, hDC, 0, 0, vbSrcCopy)
    
    ' Convierte la imagen anterior en un Mapa de bits
    Destino.Picture = Destino.Image
    
    ' Borra la imagen ya que ahora usa el Picture
    Call Destino.Cls
          
    On Error Resume Next

    If Err.Number = 0 Then
        ' Si el control no está en un  Frame, restaura la escala del contenedor
        Control.Container.ScaleMode = Escala_Anterior

    End If
          
End Sub

Private Sub picMapa_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

    If Button = 1 And Shift = vbShiftMask Then
        Call DibujarMinimapa
        
    ElseIf Button = 1 Then
    
        UserPos.X = X
        UserPos.y = y
        frmMapa.ApuntadorRadar.Left = (UserPos.X) - HalfWindowTileWidth
        frmMapa.ApuntadorRadar.Top = (UserPos.y) - HalfWindowTileHeight
    
    ElseIf Button = 2 Then
        Call AddtoRichTextBox(frmConsola.StatTxt, "Guardando Minimapa...", 255, 255, 255)
        Call Capturar_Imagen(frmMapa.picMapa, frmMapa.picMapa)
        Call SavePicture(frmMapa.picMapa, App.Path & "\Render\Minimapa\" & NumMap_Save & ".bmp")
        Call AddtoRichTextBox(frmConsola.StatTxt, "Minimapa guardado.", 0, 255, 0)
    End If
End Sub
