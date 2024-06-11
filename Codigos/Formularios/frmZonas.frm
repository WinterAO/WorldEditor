VERSION 5.00
Begin VB.Form frmZonas 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración de Zonas"
   ClientHeight    =   4830
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
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraAreas 
      BackColor       =   &H00535353&
      Caption         =   "Zonas"
      ForeColor       =   &H00FFFFFF&
      Height          =   4785
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3705
      Begin WinterMapEditor.lvButtons_H LvBQuitar 
         Height          =   405
         Left            =   1230
         TabIndex        =   6
         Top             =   4260
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   2190
         TabIndex        =   5
         Top             =   4260
         Width           =   1365
         _ExtentX        =   2408
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
      Begin WinterMapEditor.lvButtons_H LvBResetear 
         Height          =   405
         Left            =   180
         TabIndex        =   4
         Top             =   4260
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
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
         cBack           =   16711935
      End
      Begin WinterMapEditor.lvButtons_H LvBZona 
         Height          =   405
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   3720
         Width           =   1515
         _ExtentX        =   2672
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
      Begin WinterMapEditor.lvButtons_H LvBZona 
         Height          =   405
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   3720
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   714
         Caption         =   "Nueva"
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
         Height          =   405
         Index           =   8
         Left            =   3090
         TabIndex        =   7
         ToolTipText     =   "Información de la zona"
         Top             =   3720
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   714
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

Private Sub LvBEdit_Click(Index As Integer)
    frmMapInfo.Show , frmMain
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

Private Sub LvBZona_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
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
