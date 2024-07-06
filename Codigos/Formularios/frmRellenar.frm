VERSION 5.00
Begin VB.Form frmRellenar 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rellenar en area"
   ClientHeight    =   5160
   ClientLeft      =   16905
   ClientTop       =   9480
   ClientWidth     =   4185
   ControlBox      =   0   'False
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
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   279
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraParticulas 
      BackColor       =   &H00535353&
      Caption         =   "Particulas"
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
      Height          =   765
      Left            =   60
      TabIndex        =   22
      Top             =   3960
      Width           =   4035
      Begin WinterMapEditor.lvButtons_H LvBAreas 
         Height          =   405
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   714
         Caption         =   "Eliminar"
         CapAlign        =   2
         BackStyle       =   2
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
   Begin VB.CheckBox chkPedirConfirmación 
      BackColor       =   &H00424242&
      Caption         =   "Pedir confirmación"
      ForeColor       =   &H8000000B&
      Height          =   225
      Left            =   90
      TabIndex        =   21
      Top             =   4830
      Value           =   1  'Checked
      Width           =   3975
   End
   Begin VB.Frame FraZonas 
      BackColor       =   &H00535353&
      Caption         =   "Zonas"
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
      Height          =   765
      Left            =   60
      TabIndex        =   18
      Top             =   3180
      Width           =   4035
      Begin WinterMapEditor.lvButtons_H LvBAreas 
         Height          =   405
         Index           =   6
         Left            =   1830
         TabIndex        =   19
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   714
         Caption         =   "Rellenar"
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
      Begin WinterMapEditor.lvButtons_H LvBAreas 
         Height          =   405
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   714
         Caption         =   "Eliminar"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00535353&
      Caption         =   "Triggers"
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
      Height          =   765
      Left            =   60
      TabIndex        =   15
      Top             =   2400
      Width           =   4035
      Begin WinterMapEditor.lvButtons_H LvBAreas 
         Height          =   405
         Index           =   4
         Left            =   1830
         TabIndex        =   16
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   714
         Caption         =   "Rellenar"
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
      Begin WinterMapEditor.lvButtons_H LvBAreas 
         Height          =   405
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   714
         Caption         =   "Eliminar"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00535353&
      Caption         =   "Superficies"
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
      Height          =   765
      Left            =   60
      TabIndex        =   12
      Top             =   1620
      Width           =   4035
      Begin WinterMapEditor.lvButtons_H LvBAreas 
         Height          =   405
         Index           =   2
         Left            =   1830
         TabIndex        =   13
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   714
         Caption         =   "Rellenar"
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
      Begin WinterMapEditor.lvButtons_H LvBAreas 
         Height          =   405
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   714
         Caption         =   "Eliminar"
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
   Begin VB.Frame FraSuperficie 
      BackColor       =   &H00535353&
      Caption         =   "Bloqueos"
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
      Height          =   765
      Left            =   60
      TabIndex        =   9
      Top             =   780
      Width           =   4035
      Begin WinterMapEditor.lvButtons_H LvBAreas 
         Height          =   405
         Index           =   1
         Left            =   1830
         TabIndex        =   10
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   714
         Caption         =   "Rellenar"
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
      Begin WinterMapEditor.lvButtons_H LvBAreas 
         Height          =   405
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   714
         Caption         =   "Eliminar"
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
   Begin VB.Frame FraCoordenadas 
      BackColor       =   &H00535353&
      Caption         =   "Coordenadas"
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
      Height          =   645
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4035
      Begin VB.TextBox DY2 
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
         Left            =   3390
         TabIndex        =   4
         Text            =   "5"
         Top             =   210
         Width           =   495
      End
      Begin VB.TextBox DY1 
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
         Left            =   2430
         TabIndex        =   3
         Text            =   "1"
         Top             =   210
         Width           =   495
      End
      Begin VB.TextBox DX2 
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
         Left            =   1470
         TabIndex        =   2
         Text            =   "5"
         Top             =   210
         Width           =   495
      End
      Begin VB.TextBox DX1 
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
         Left            =   510
         TabIndex        =   1
         Text            =   "1"
         Top             =   210
         Width           =   495
      End
      Begin VB.Label lblX2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X1:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblX2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X2:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1185
         TabIndex        =   7
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblY1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y1:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2070
         TabIndex        =   6
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblY2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y2:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3105
         TabIndex        =   5
         Top             =   225
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmRellenar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LvBAreas_Click(Index As Integer)
    If IsNumeric(DX1.Text) = False Or _
       IsNumeric(DX2.Text) = False Or _
       IsNumeric(DY1.Text) = False Or _
       IsNumeric(DY2.Text) = False Then
    
        Call MsgBox("Debes introducir valores nï¿½mericos. Estos pueden tener un mï¿½nimo de 1 y un mï¿½ximo de " & (YMinMapSize + XMinMapSize) / 2 & ".")
    
       Exit Sub
    End If
    
    Select Case Index
        Case 0
            Call Bloqueos_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
            
        Case 1
            Call Bloqueos_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
            
        Case 2
            Call Superficie_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
            
        Case 3
            Call Superficie_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
            
        Case 4
            Call Triggers_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
            
        Case 5
            Call Triggers_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
            
        Case 6
            Call Zonas_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
            
        Case 7
            Call Zonas_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
            
        Case 8
            Call Particulas_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
            
    End Select
End Sub

Public Sub Superficie_Area(ByVal x1 As Long, ByVal x2 As Long, ByVal y1 As Long, ByVal y2 As Long, ByVal Poner As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 07/12/2018
'*************************************************

    If chkPedirConfirmación.value Then
        If EditWarning Then Exit Sub
    End If
    
    Dim y As Long
    Dim x As Long
    
    If Not MapaCargado Then Exit Sub

    For y = y1 To y2
        For x = x1 To x2
            If Poner = True Then
                If frmConfigSup.MOSAICO.value = vbChecked Then
                    Dim aux As Long
                    aux = Val(frmSuperficies.cGrh.Text) + _
                    ((y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (x Mod frmConfigSup.mAncho)
                     MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                    'Setup GRH
                    InitGrh MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)), aux
                Else
                    'Else Place graphic
                    MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = Val(frmSuperficies.cGrh.Text)
                    'Setup GRH
                    InitGrh MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)), Val(frmSuperficies.cGrh.Text)
                End If
            Else
                
                If Val(frmSuperficies.cCapas.Text) = 1 Then
                    MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = 1
                    
                Else
                    MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = 0
                    
                End If
            End If
        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

Public Sub Bloqueos_Area(ByVal x1 As Long, ByVal x2 As Long, ByVal y1 As Long, ByVal y2 As Long, ByVal Inserta As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 07/12/2018
'*************************************************

    If chkPedirConfirmación.value Then
        If EditWarning Then Exit Sub
    End If
    
    Dim y As Long
    Dim x As Long
    
    If Not MapaCargado Then Exit Sub

    For y = y1 To y2
        For x = x1 To x2
    
            If y > YMinMapSize Or y < YMaxMapSize Then
                If x > XMinMapSize Or x < XMaxMapSize Then
    
                    If Inserta = True Then
                        MapData(x, y).bLocked = 1
                    Else
                        MapData(x, y).bLocked = 0
                    End If
                    
                End If
            End If
            
    
        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

Public Sub Triggers_Area(ByVal x1 As Long, _
                         ByVal x2 As Long, _
                         ByVal y1 As Long, _
                         ByVal y2 As Long, _
                         ByVal Poner As Boolean)
    '*************************************************
    'Author: Lorwik
    'Last modified: 25/03/2021
    '*************************************************

    If chkPedirConfirmación.value Then
        If EditWarning Then Exit Sub

    End If
    
    Dim y As Long

    Dim x As Long
    
    If Not MapaCargado Then Exit Sub

    For y = y1 To y2
        For x = x1 To x2

            If Poner = True Then
                MapInfo.Changed = 1 'Set changed flag
                MapData(x, y).Trigger = frmTriggers.LynxTriggers.CellText(, 0)

            Else
                MapInfo.Changed = 1
                MapData(x, y).Trigger = 0
                
            End If

        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

Public Sub Zonas_Area(ByVal x1 As Long, _
                      ByVal x2 As Long, _
                      ByVal y1 As Long, _
                      ByVal y2 As Long, _
                      ByVal Poner As Boolean)
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/04/2021
    '*************************************************

    If chkPedirConfirmación.value Then
        If EditWarning Then Exit Sub

    End If
    
    Dim y As Long

    Dim x As Long
    
    If Not MapaCargado Then Exit Sub

    For y = y1 To y2
        For x = x1 To x2

            If Poner = True Then
                MapInfo.Changed = 1 'Set changed flag
                MapData(x, y).ZonaIndex = frmZonas.LstZona.ListIndex + 1
            Else
                MapInfo.Changed = 1
                MapData(x, y).ZonaIndex = 0
                
            End If

        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

Public Sub Particulas_Area(ByVal x1 As Long, _
                           ByVal x2 As Long, _
                           ByVal y1 As Long, _
                           ByVal y2 As Long, _
                           ByVal Poner As Boolean)
    '*************************************************
    'Author: Lorwik
    'Last modified: 13/06/2024
    '*************************************************

    If chkPedirConfirmación.value Then
        If EditWarning Then Exit Sub

    End If
    
    Dim y As Long

    Dim x As Long
    
    If Not MapaCargado Then Exit Sub

    For y = y1 To y2
        For x = x1 To x2

            If Poner Then
                General_Particle_Create CLng(frmParticulas.LynxParticulas.CellText(, 0)), x, y, CLng(-1)
                MapData(x, y).Particle_Index = CLng(frmParticulas.LynxParticulas.CellText(, 0))
            Else
                MapInfo.Changed = 1
                MapData(x, y).Particle_Index = 0
                MapData(x, y).Particle_Group_Index = 0
            End If
                

        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub
