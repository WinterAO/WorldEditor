VERSION 5.00
Begin VB.Form frmOptimizar 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Optimizar el Mapa"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   264
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   251
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WinterMapEditor.lvButtons_H cOptimizar 
      Height          =   525
      Left            =   1620
      TabIndex        =   7
      Top             =   3360
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   926
      Caption         =   "Optimizar"
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
   Begin VB.Frame FraFuncionesDe 
      BackColor       =   &H00535353&
      Caption         =   "Funciones de optimización"
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
      Height          =   3195
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3615
      Begin VB.CheckBox chkLimpiarGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Limpiar Grh nulos"
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
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   3375
      End
      Begin VB.CheckBox chkQuitarTrans 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Quitar Translados Bloqueados"
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
         Left            =   150
         TabIndex        =   6
         Top             =   330
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.CheckBox chkQuitarTrigBloq 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Quitar Trigger's Bloqueados"
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
         Left            =   150
         TabIndex        =   5
         Top             =   690
         Width           =   3375
      End
      Begin VB.CheckBox chkQuitarTrigTrans 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Quitar Trigger's en Translados"
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
         Left            =   150
         TabIndex        =   4
         Top             =   1050
         Width           =   3375
      End
      Begin VB.CheckBox chkQuitarTodoBordes 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Quitar NPCs, Objetos y Translados en los Bordes Exteriores"
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
         Height          =   375
         Left            =   150
         TabIndex        =   3
         Top             =   1410
         Width           =   3375
      End
      Begin VB.CheckBox chkMapearArbolesEtc 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Mapear Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa"
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
         Height          =   375
         Left            =   150
         TabIndex        =   2
         Top             =   1890
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.CheckBox chkBloquearArbolesEtc 
         Appearance      =   0  'Flat
         BackColor       =   &H00535353&
         Caption         =   "Bloquear Arboles, Carteles, Foros y Yacimientos"
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
         Height          =   375
         Left            =   150
         TabIndex        =   1
         Top             =   2370
         Value           =   1  'Checked
         Width           =   3375
      End
   End
   Begin WinterMapEditor.lvButtons_H cCancelar 
      Height          =   525
      Left            =   60
      TabIndex        =   8
      Top             =   3360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   926
      Caption         =   "Cancelar"
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
Attribute VB_Name = "frmOptimizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private optimizacionMasiva As Boolean

Private Sub cCancelar_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************
    Unload Me
    
End Sub

Private Sub cOptimizar_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************
    
    Dim RutadelMapa As String
    Dim MinMapa As Integer
    Dim MaxMapa As Integer
    Dim i As Integer
    
    Call Optimizar
        
End Sub

Public Sub Optimizar()
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************
    
    Dim y As Integer
    Dim x As Integer
    Dim i As Byte
    
    If Not MapaCargado Then
        Exit Sub

    End If
    
    ' Quita Traslados Bloqueados
    ' Quita Trigger's Bloqueados
    ' Quita Trigger's en Traslados
    ' Quita NPCs, Objetos y Traslados en los Bordes Exteriores
    ' Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
    ' Eliminar Grh inexistentes
    
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
    
            With MapData(x, y)
    
                ' ** Quitar NPCs, Objetos y Traslados en los Bordes Exteriores
                If (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) And chkQuitarTodoBordes.value = 1 Then
    
                    'Quitar NPCs
                    If MapData(x, y).NPCIndex > 0 Then
                        Char_Erase MapData(x, y).CharIndex
                        MapData(x, y).NPCIndex = 0
    
                    End If
    
                    ' Quitar Objetos
                    MapData(x, y).OBJInfo.ObjIndex = 0
                    MapData(x, y).OBJInfo.Amount = 0
                    MapData(x, y).ObjGrh.GrhIndex = 0
                    ' Quitar Traslados
                    MapData(x, y).TileExit.Map = 0
                    MapData(x, y).TileExit.x = 0
                    MapData(x, y).TileExit.y = 0
                    ' Quitar Triggers
                    MapData(x, y).Trigger = 0
    
                End If
    
                ' ** Quitar Traslados y Triggers en Bloqueo
                If MapData(x, y).bLocked = 1 Then
                    If MapData(x, y).TileExit.Map > 0 And chkQuitarTrans.value = 1 Then ' Quita Translado Bloqueado
                        MapData(x, y).TileExit.Map = 0
                        MapData(x, y).TileExit.y = 0
                        MapData(x, y).TileExit.x = 0
                    ElseIf MapData(x, y).Trigger > 0 And chkQuitarTrigBloq.value = 1 Then ' Quita Trigger Bloqueado
                        MapData(x, y).Trigger = 0
    
                    End If
    
                End If
    
                ' ** Quitar Triggers en Translado
                If MapData(x, y).TileExit.Map > 0 And chkQuitarTrigTrans.value = 1 Then
                    If MapData(x, y).Trigger > 0 Then ' Quita Trigger en Translado
                        MapData(x, y).Trigger = 0
    
                    End If
    
                End If
    
                ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
                If MapData(x, y).OBJInfo.ObjIndex > 0 And (chkMapearArbolesEtc.value = 1 Or chkBloquearArbolesEtc.value = 1) Then
    
                    Select Case ObjData(MapData(x, y).OBJInfo.ObjIndex).OBJType
    
                        Case 8, 10 ' Carteles, Foros
    
                            If MapData(x, y).Graphic(3).GrhIndex <> MapData(x, y).ObjGrh.GrhIndex And chkMapearArbolesEtc.value = 1 Then MapData(x, y).Graphic(3) = MapData(x, y).ObjGrh
                            If chkBloquearArbolesEtc.value = 1 And MapData(x, y).bLocked = 0 Then MapData(x, y).bLocked = 1
    
                    End Select
    
                End If
                
                ' ** Eliminar GRH inexistentes
                If chkLimpiarGrh.value = 1 Then
                
                    ' Capa 1
                    If .Graphic(1).GrhIndex = 0 Then
                        .Graphic(1).GrhIndex = 1
                        Call AddtoRichTextBox(frmConsola.StatTxt, "Grh inexistente en la capa 1, en las coordenadas " & x & " - " & y, 255, 201, 14)
                
                    ElseIf GrhData(.Graphic(1).GrhIndex).NumFrames < 2 And GrhData(.Graphic(1).GrhIndex).FileNum = 0 Then
                        .Graphic(1).GrhIndex = 1
                        Call AddtoRichTextBox(frmConsola.StatTxt, "Se ha eliminado el Grh: " & .Graphic(1).GrhIndex & " en la capa 1, en las coordenadas " & x & " - " & y, 255, 201, 14)

                    End If
                        
                    ' Capa 2, 3 y 4
                    For i = 2 To 4
                        If .Graphic(i).GrhIndex > 0 Then
                            If GrhData(.Graphic(i).GrhIndex).NumFrames < 2 And GrhData(.Graphic(i).GrhIndex).FileNum = 0 Then
                                .Graphic(i).GrhIndex = 0
                                Call AddtoRichTextBox(frmConsola.StatTxt, "Se ha eliminado el Grh: " & .Graphic(i).GrhIndex & ", en la capa " & i & " en las coordenadas " & x & " - " & y, 255, 201, 14)
    
                            End If
                        End If
                    Next i
                
                End If
                
            End With

        Next x
    Next y

    'Set changed flag
    MapInfo.Changed = 1

End Sub

Private Sub Form_Load()
    'Si tenemos ventana grande es que queremos optimizar muchos mapas
    If Me.Height > 4215 Then
        optimizacionMasiva = True
        
    Else 'Ventana chica, normal...
        optimizacionMasiva = False
        
    End If
End Sub
