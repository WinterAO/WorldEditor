VERSION 5.00
Begin VB.Form frmInformes 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informes"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6495
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
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WinterMapEditor.lvButtons_H LvBInformar 
      Height          =   525
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   4020
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   926
      Caption         =   "&Objetos"
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
      cBack           =   65535
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
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
      Height          =   3855
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   90
      Width           =   6375
   End
   Begin WinterMapEditor.lvButtons_H LvBInformar 
      Height          =   525
      Index           =   1
      Left            =   1950
      TabIndex        =   2
      Top             =   4020
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   926
      Caption         =   "&Traslados"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
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
   Begin WinterMapEditor.lvButtons_H LvBInformar 
      Height          =   525
      Index           =   2
      Left            =   4140
      TabIndex        =   3
      Top             =   4020
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   926
      Caption         =   "&NPC's"
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
      cBack           =   33023
   End
End
Attribute VB_Name = "frmInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LvBInformar_Click(Index As Integer)

    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************
    On Error GoTo LvBInformar_Click_Err
    
    Select Case Index
    
        Case 0 ' Objetos
            Call ActalizarObjetos
            
        Case 1 ' Traslados
            Call ActalizarTraslados
        
        Case 2 ' NPC's
            Call ActalizarNPCs
    
    End Select
    
    Exit Sub
    
LvBInformar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.ObtenerNombreArchivo", Erl)

    Resume Next

End Sub

Private Sub ActalizarObjetos()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    'Genera el informe de Objetos
    '*************************************************
    On Error Resume Next

    Dim y As Integer

    Dim x As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    txtInfo.Text = "Informe de Objetos (X,Y)"

    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize

            If MapData(x, y).OBJInfo.ObjIndex > 0 Then
                txtInfo.Text = txtInfo.Text & vbCrLf & x & "," & y & " tiene " & MapData(x, y).OBJInfo.Amount & " del Objeto " & MapData(x, y).OBJInfo.ObjIndex & " - " & ObjData(MapData(x, y).OBJInfo.ObjIndex).name

            End If

        Next x
    Next y

End Sub

Private Sub ActalizarNPCs()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    'Genera el informe de NPCs
    '*************************************************
    On Error Resume Next

    Dim y As Integer

    Dim x As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    txtInfo.Text = "Informe de NPCs/Hostiles (X,Y)"

    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize

            If MapData(x, y).NPCIndex > 0 Then
                If MapData(x, y).NPCIndex >= 500 Then
                    txtInfo.Text = txtInfo.Text & vbCrLf & x & "," & y & " tiene " & NpcData(MapData(x, y).NPCIndex).name & " (Hostil)"
                    
                Else
                    txtInfo.Text = txtInfo.Text & vbCrLf & x & "," & y & " tiene " & NpcData(MapData(x, y).NPCIndex).name

                End If

            End If

        Next x
    Next y

End Sub

Private Sub ActalizarTraslados()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    'Genera el informe de Traslados
    '*************************************************
    On Error Resume Next

    Dim y As Integer

    Dim x As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    txtInfo.Text = "Informe de Traslados (X,Y)"

    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize

            If MapData(x, y).TileExit.Map > 0 Then
                txtInfo.Text = txtInfo.Text & vbCrLf & x & "," & y & " nos traslada a la posición " & MapData(x, y).TileExit.x & "," & MapData(x, y).TileExit.y & " del Mapa " & MapData(x, y).TileExit.Map

                If ((x < 20 And MapData(x, y).TileExit.x < 20) Or (x > 80 And MapData(x, y).TileExit.x > 80)) And (x <> MapData(x, y).TileExit.x) Then
                    txtInfo.Text = txtInfo.Text & " (X sospechoso)"

                End If

                If ((y < 20 And MapData(x, y).TileExit.y < 20) Or (y > 80 And MapData(x, y).TileExit.y > 80)) And (y <> MapData(x, y).TileExit.y) Then
                    txtInfo.Text = txtInfo.Text & " (Y sospechoso)"

                End If

            End If

        Next x
    Next y

End Sub
