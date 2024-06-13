VERSION 5.00
Begin VB.Form frmZonas 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración de Zonas"
   ClientHeight    =   5520
   ClientLeft      =   10215
   ClientTop       =   6885
   ClientWidth     =   3825
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
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraAreas 
      BackColor       =   &H00535353&
      Caption         =   "Zonas"
      ForeColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3705
      Begin WinterMapEditor.lvButtons_H LvBQuitar 
         Height          =   405
         Left            =   120
         TabIndex        =   5
         Top             =   4920
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   714
         Caption         =   "Quitar"
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
         Mode            =   1
         Value           =   0   'False
         cBack           =   255
      End
      Begin WinterMapEditor.lvButtons_H LvBPintar 
         Height          =   405
         Left            =   1650
         TabIndex        =   4
         Top             =   4890
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   714
         Caption         =   "Insertar"
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
         Mode            =   1
         Value           =   0   'False
         cBack           =   65280
      End
      Begin WinterMapEditor.lvButtons_H LvBZona 
         Height          =   405
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   3720
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
         Caption         =   "Eliminar Zona"
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
      Begin WinterMapEditor.lvButtons_H LvBZona 
         Height          =   405
         Index           =   0
         Left            =   1650
         TabIndex        =   2
         Top             =   3720
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   714
         Caption         =   "Nueva Zona"
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
      Begin VB.ListBox LstZona 
         Appearance      =   0  'Flat
         Height          =   3345
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin WinterMapEditor.lvButtons_H LvBEdit 
         Height          =   480
         Index           =   8
         Left            =   2430
         TabIndex        =   6
         ToolTipText     =   "Información de la zona"
         Top             =   4260
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "frmZonas.frx":0000
         ImgSize         =   32
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBRecolorear 
         Height          =   480
         Left            =   1560
         TabIndex        =   7
         ToolTipText     =   "Información de la zona"
         Top             =   4260
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "frmZonas.frx":02D8
         ImgSize         =   32
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H lvBResetear 
         Height          =   480
         Left            =   720
         TabIndex        =   8
         ToolTipText     =   "Información de la zona"
         Top             =   4260
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "frmZonas.frx":09C8
         ImgSize         =   32
         cBack           =   -2147483633
      End
   End
End
Attribute VB_Name = "frmZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub LstZona_Click()
'********************************
'Autor: Lorwik
'Fecha: ???
'********************************

    Call MapZona_Actualizar(LstZona.ListIndex + 1)
    
End Sub

Private Sub LvBEdit_Click(Index As Integer)
'********************************
'Autor: Lorwik
'Fecha: ???
'********************************

    frmMapInfo.Show , frmMain
End Sub

Private Sub LvBPintar_Click()
'********************************
'Autor: Lorwik
'Fecha: ???
'********************************

    If LvBPintar.value = True Then
        LvBQuitar.Enabled = False
        
    Else
        LvBQuitar.Enabled = True
        
    End If
End Sub

Private Sub LvBQuitar_Click()
'********************************
'Autor: Lorwik
'Fecha: ???
'********************************

    If LvBQuitar.value = True Then
        LvBPintar.Enabled = False
        
    Else
        LvBPintar.Enabled = True
        
    End If
End Sub

Private Sub LvBRecolorear_Click()
'********************************
'Autor: Lorwik
'Fecha: 13/06/2024
'********************************

    Call coloresZona
    
End Sub

Private Sub LvBResetear_Click()
'********************************
'Autor: Lorwik
'Fecha: ???
'********************************

    If MsgBox("¿¡Estas seguro que deseas resetear las propiedades de la zona!?", vbExclamation + vbYesNo) = vbYes Then
        Call ResetearZona(LstZona.ListIndex + 1)
        Call ActualizarZonaList(LstZona.ListIndex)
    End If
End Sub

Private Sub LvBZona_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'********************************
'Autor: Lorwik
'Fecha: ???
'********************************

    Select Case Index
    
        Case 0
            Call NuevaZona(Index + 1)
            
        Case 1
        
            If EditWarning Then Exit Sub
        
            If LstZona.ListIndex + 1 <> CantZonas Then
                MsgBox "Solo puedes eliminar la ultima zona de la lista. Si no vas a utilizar mas esa zona, reseteala para reutilizarla en el futuro."
                Exit Sub
            End If
            
            Call EliminarZona
            
    End Select
    
End Sub
