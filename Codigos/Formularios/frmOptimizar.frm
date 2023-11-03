VERSION 5.00
Begin VB.Form frmOptimizar 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Optimizar el Mapa"
   ClientHeight    =   3660
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
   ScaleHeight     =   244
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   251
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WinterMapEditor.lvButtons_H cOptimizar 
      Height          =   525
      Left            =   1650
      TabIndex        =   7
      Top             =   3120
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
      Height          =   2985
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3615
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
      Left            =   90
      TabIndex        =   8
      Top             =   3120
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

    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub

    End If
    
    ' Quita Translados Bloqueados
    ' Quita Trigger's Bloqueados
    ' Quita Trigger's en Translados
    ' Quita NPCs, Objetos y Translados en los Bordes Exteriores
    ' Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            ' ** Quitar NPCs, Objetos y Translados en los Bordes Exteriores
            If (X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) And chkQuitarTodoBordes.value = 1 Then

                'Quitar NPCs
                If MapData(X, y).NPCIndex > 0 Then
                    EraseChar MapData(X, y).CharIndex
                    MapData(X, y).NPCIndex = 0

                End If

                ' Quitar Objetos
                MapData(X, y).OBJInfo.ObjIndex = 0
                MapData(X, y).OBJInfo.Amount = 0
                MapData(X, y).ObjGrh.GrhIndex = 0
                ' Quitar Translados
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0
                ' Quitar Triggers
                MapData(X, y).Trigger = 0

            End If

            ' ** Quitar Translados y Triggers en Bloqueo
            If MapData(X, y).bLocked = 1 Then
                If MapData(X, y).TileExit.Map > 0 And chkQuitarTrans.value = 1 Then ' Quita Translado Bloqueado
                    MapData(X, y).TileExit.Map = 0
                    MapData(X, y).TileExit.y = 0
                    MapData(X, y).TileExit.X = 0
                ElseIf MapData(X, y).Trigger > 0 And chkQuitarTrigBloq.value = 1 Then ' Quita Trigger Bloqueado
                    MapData(X, y).Trigger = 0

                End If

            End If

            ' ** Quitar Triggers en Translado
            If MapData(X, y).TileExit.Map > 0 And chkQuitarTrigTrans.value = 1 Then
                If MapData(X, y).Trigger > 0 Then ' Quita Trigger en Translado
                    MapData(X, y).Trigger = 0

                End If

            End If

            ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
            If MapData(X, y).OBJInfo.ObjIndex > 0 And (chkMapearArbolesEtc.value = 1 Or chkBloquearArbolesEtc.value = 1) Then

                Select Case ObjData(MapData(X, y).OBJInfo.ObjIndex).ObjType

                    Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos

                        If MapData(X, y).Graphic(3).GrhIndex <> MapData(X, y).ObjGrh.GrhIndex And chkMapearArbolesEtc.value = 1 Then MapData(X, y).Graphic(3) = MapData(X, y).ObjGrh
                        If chkBloquearArbolesEtc.value = 1 And MapData(X, y).bLocked = 0 Then MapData(X, y).bLocked = 1

                End Select

            End If

            ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
        Next X
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
