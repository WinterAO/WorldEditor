VERSION 5.00
Begin VB.Form frmZonas 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración de Zonas"
   ClientHeight    =   5220
   ClientLeft      =   10215
   ClientTop       =   6885
   ClientWidth     =   3465
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
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   231
   ShowInTaskbar   =   0   'False
   Begin WinterMapEditor.lvButtons_H LvBCerrar 
      Height          =   405
      Left            =   750
      TabIndex        =   7
      Top             =   4740
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   714
      Caption         =   "Cerrar"
      CapAlign        =   2
      BackStyle       =   2
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
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Frame FraAreas 
      BackColor       =   &H00535353&
      Caption         =   "Zonas"
      ForeColor       =   &H00FFFFFF&
      Height          =   4635
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3345
      Begin WinterMapEditor.lvButtons_H LvBQuitar 
         Height          =   315
         Left            =   1770
         TabIndex        =   6
         Top             =   4200
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         Caption         =   "Quitar"
         CapAlign        =   2
         BackStyle       =   2
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
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBPintar 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   4200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "Insertar"
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
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBResetear 
         Height          =   345
         Left            =   2280
         TabIndex        =   4
         Top             =   3810
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "Resetear"
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
      Begin WinterMapEditor.lvButtons_H LvBZona 
         Height          =   345
         Index           =   1
         Left            =   1020
         TabIndex        =   3
         Top             =   3810
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         Caption         =   "Eliminar"
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
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBZona 
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   3810
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         Caption         =   "Nueva"
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
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ListBox LstZona 
         Appearance      =   0  'Flat
         Height          =   3345
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
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
 
    Call MapZona_Actualizar(LstZona.ListIndex + 1)
    
End Sub

Private Sub LvBCerrar_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Unload Me
End Sub

Private Sub LvBPintar_Click()
    If LvBPintar.value = True Then
        LvBQuitar.Enabled = False
        
    Else
        LvBQuitar.Enabled = True
        
    End If
End Sub

Private Sub LvBQuitar_Click()
    If LvBQuitar.value = True Then
        LvBPintar.Enabled = False
        
    Else
        LvBPintar.Enabled = True
        
    End If
End Sub

Private Sub LvBResetear_Click()
    If MsgBox("¿¡Estas seguro que deseas resetear las propiedades de la zona!?", vbExclamation + vbYesNo) = vbYes Then
        Call ResetearZona(LstZona.ListIndex + 1)
        Call ActualizarZonaList(LstZona.ListIndex)
    End If
End Sub

Private Sub LvBZona_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    Select Case Index
    
        Case 0
            Call NuevaZona(Index + 1)
            
        Case 1
            If LstZona.ListIndex + 1 <> CantZonas Then
                MsgBox "Solo puedes eliminar la ultima zona de la lista. Si no vas a utilizar mas esa zona, reseteala para reutilizarla en el futuro."
                Exit Sub
            End If
            
            Call EliminarZona
            
    End Select
End Sub
