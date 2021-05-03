VERSION 5.00
Begin VB.Form frmRellenar 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rellenar en area"
   ClientHeight    =   4020
   ClientLeft      =   16905
   ClientTop       =   9480
   ClientWidth     =   4185
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
   ScaleHeight     =   268
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   279
   ShowInTaskbar   =   0   'False
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
    
        Call MsgBox("Debes introducir valores n�mericos. Estos pueden tener un m�nimo de 1 y un m�ximo de " & (YMinMapSize + XMinMapSize) / 2 & ".")
    
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
            
    End Select
End Sub

Public Sub Superficie_Area(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer, ByVal Poner As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 07/12/2018
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If

    For Y = y1 To y2
        For X = x1 To x2
            If Poner = True Then
                If frmConfigSup.MOSAICO.value = vbChecked Then
                    Dim aux As Integer
                    aux = Val(frmSuperficies.cGrh.Text) + _
                    ((Y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (X Mod frmConfigSup.mAncho)
                     MapData(X, Y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                    'Setup GRH
                    InitGrh MapData(X, Y).Graphic(Val(frmSuperficies.cCapas.Text)), aux
                Else
                    'Else Place graphic
                    MapData(X, Y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = Val(frmSuperficies.cGrh.Text)
                    'Setup GRH
                    InitGrh MapData(X, Y).Graphic(Val(frmSuperficies.cCapas.Text)), Val(frmSuperficies.cGrh.Text)
                End If
            Else
                MapData(X, Y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = 0
            End If
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

Public Sub Bloqueos_Area(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer, ByVal Inserta As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 07/12/2018
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If

    For Y = y1 To y2
        For X = x1 To x2
    
            If Inserta = True Then
                MapData(X, Y).bLocked = 1
            Else
                MapData(X, Y).bLocked = 0
            End If
    
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

Public Sub Triggers_Area(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer, ByVal Poner As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 25/03/2021
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If

    For Y = y1 To y2
        For X = x1 To x2
            If Poner = True Then
                If frmConfigSup.MOSAICO.value = vbChecked Then
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(X, Y).Trigger = frmTriggers.LynxTriggers.CellText(, 0)
                Else
                    MapInfo.Changed = 1
                    'Else Place trigger
                    MapData(X, Y).Trigger = 0

                End If
            Else
                MapInfo.Changed = 1
                MapData(X, Y).Trigger = 0
                
            End If
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

Public Sub Zonas_Area(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer, ByVal Poner As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 01/04/2021
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If

    For Y = y1 To y2
        For X = x1 To x2
            If Poner = True Then
                If frmConfigSup.MOSAICO.value = vbChecked Then
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(X, Y).ZonaIndex = frmZonas.LstZona.ListIndex + 1
                    Debug.Print "Inserto"
                Else
                    MapInfo.Changed = 1
                    'Else Place Zona
                    MapData(X, Y).ZonaIndex = 0
                    Debug.Print "Quito"
                End If
            Else
                MapInfo.Changed = 1
                MapData(X, Y).ZonaIndex = 0
                Debug.Print "Quito 2"
                
            End If
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub


