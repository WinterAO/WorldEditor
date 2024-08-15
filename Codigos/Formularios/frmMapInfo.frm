VERSION 5.00
Begin VB.Form frmMapInfo 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información de la Zona"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
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
   Icon            =   "frmMapInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraInformacion 
      BackColor       =   &H00535353&
      Caption         =   "Informacion"
      ForeColor       =   &H00FFFFFF&
      Height          =   6105
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4695
      Begin VB.TextBox TxtAmbient 
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
         Left            =   1680
         TabIndex        =   31
         Text            =   "0"
         Top             =   1050
         Width           =   2655
      End
      Begin VB.TextBox txtNivelMaximo 
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
         Left            =   1680
         TabIndex        =   29
         Text            =   "0"
         Top             =   2490
         Width           =   2655
      End
      Begin WinterMapEditor.lvButtons_H cmdCerrar 
         Height          =   525
         Left            =   180
         TabIndex        =   28
         Top             =   5400
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   714
         Caption         =   "Cerrar"
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
      Begin WinterMapEditor.lvButtons_H LvBGuardar 
         Height          =   525
         Left            =   2400
         TabIndex        =   27
         Top             =   5400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         Caption         =   "Guardar"
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
         cBack           =   12583104
      End
      Begin WinterMapEditor.lvButtons_H cmdMusica 
         Height          =   345
         Left            =   3630
         TabIndex        =   25
         Top             =   690
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   609
         Caption         =   "&Más"
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
      Begin VB.ComboBox txtMapZona 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmMapInfo.frx":628A
         Left            =   1680
         List            =   "frmMapInfo.frx":6297
         TabIndex        =   24
         Text            =   "txtMapZona"
         Top             =   1380
         Width           =   2655
      End
      Begin VB.Frame FraLuzBase 
         BackColor       =   &H00535353&
         Caption         =   "Luz base"
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
         Height          =   975
         Left            =   150
         TabIndex        =   20
         Top             =   4350
         Width           =   2175
         Begin WinterMapEditor.lvButtons_H LvBActualizarLuces 
            Height          =   375
            Left            =   1440
            TabIndex        =   26
            Top             =   150
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   661
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
            Image           =   "frmMapInfo.frx":62B3
            cBack           =   -2147483633
         End
         Begin VB.TextBox LuzMapa 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   23
            Top             =   580
            Width           =   1335
         End
         Begin VB.PictureBox PicColorMap 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   480
            Width           =   375
         End
         Begin VB.CheckBox chkLuzClimatica 
            BackColor       =   &H00535353&
            Caption         =   "Luz climatica"
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
            Height          =   195
            Left            =   120
            MaskColor       =   &H00404040&
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkMapMagiaSinEfecto 
         BackColor       =   &H00535353&
         Caption         =   "Magia Sin Efecto"
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
         TabIndex        =   13
         Top             =   3510
         Width           =   1575
      End
      Begin VB.CheckBox chkMapBackup 
         BackColor       =   &H00535353&
         Caption         =   "Backup"
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
         Left            =   2430
         TabIndex        =   12
         Top             =   3510
         Value           =   2  'Grayed
         Width           =   1575
      End
      Begin VB.TextBox txtMapNombre 
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
         Left            =   1680
         TabIndex        =   11
         Text            =   "Mapa Desconocido"
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtMapMusica 
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
         Left            =   1680
         TabIndex        =   10
         Text            =   "0"
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox txtMapTerreno 
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
         Height          =   315
         ItemData        =   "frmMapInfo.frx":1CBF5
         Left            =   1680
         List            =   "frmMapInfo.frx":1CC02
         TabIndex        =   9
         Text            =   "txtMapTerreno"
         Top             =   1770
         Width           =   2655
      End
      Begin VB.CheckBox chkMapPK 
         BackColor       =   &H00535353&
         Caption         =   "PK (inseguro)"
         BeginProperty DataFormat 
            Type            =   4
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   8
         EndProperty
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
         TabIndex        =   8
         Top             =   3750
         Width           =   1575
      End
      Begin VB.ComboBox txtMapRestringir 
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
         Height          =   315
         ItemData        =   "frmMapInfo.frx":1CC1F
         Left            =   1680
         List            =   "frmMapInfo.frx":1CC3B
         TabIndex        =   7
         Text            =   "txtMapRestringir"
         Top             =   2130
         Width           =   2655
      End
      Begin VB.CheckBox chkMapInviSinEfecto 
         BackColor       =   &H00535353&
         Caption         =   "InviSinEfecto"
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
         Top             =   3270
         Width           =   2055
      End
      Begin VB.CheckBox chkMapResuSinEfecto 
         BackColor       =   &H00535353&
         Caption         =   "ResuSinEfecto"
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
         Left            =   2430
         TabIndex        =   5
         Top             =   3270
         Width           =   1815
      End
      Begin VB.CheckBox ChkMapNpc 
         BackColor       =   &H00535353&
         Caption         =   "Robo de NPC Permitido"
         BeginProperty DataFormat 
            Type            =   4
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   8
         EndProperty
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
         Top             =   3990
         Width           =   2055
      End
      Begin VB.TextBox TxtlvlMinimo 
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
         Left            =   1680
         TabIndex        =   3
         Text            =   "0"
         Top             =   2850
         Width           =   2655
      End
      Begin VB.CheckBox chkInvocarSin 
         BackColor       =   &H00535353&
         Caption         =   "Invocar sin efecto"
         BeginProperty DataFormat 
            Type            =   4
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   8
         EndProperty
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
         Left            =   2430
         TabIndex        =   2
         Top             =   3750
         Width           =   1935
      End
      Begin VB.CheckBox chkOcultarSin 
         BackColor       =   &H00535353&
         Caption         =   "Ocultar sin Efecto"
         BeginProperty DataFormat 
            Type            =   4
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   8
         EndProperty
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
         Left            =   2430
         TabIndex        =   1
         Top             =   3990
         Width           =   1935
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sonido Ambiental:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1050
         Width           =   1290
      End
      Begin VB.Label lblNivelMaximo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Maximo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   2490
         Width           =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   180
         X2              =   4620
         Y1              =   3990
         Y2              =   3990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de la Zona:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Musica:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zona:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1380
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terreno:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1770
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restringir:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2130
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Minimo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2850
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmMapInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************
Option Explicit

Private Sub chkInvocarSin_LostFocus()
    '*************************************************
    'Author: Hardoz
    'Last modified: 28/08/2010
    '*************************************************
    MapInfo.InvocarSinEfecto = ChkMapNpc.value
    MapInfo.Changed = 1
 
End Sub

Private Sub chkLuzClimatica_Click()

    If chkLuzClimatica.value = Unchecked Then
        PicColorMap.BackColor = 0
            
        MapZonas(frmZonas.LstZona.ListIndex + 1).LuzBase = 0
        
        Call Actualizar_Estado
    End If
    
    frmMapInfo.chkLuzClimatica.value = chkLuzClimatica.value
    
End Sub

Private Sub chkMapBackup_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    MapInfo.BackUp = chkMapBackup.value
    MapInfo.Changed = 1
End Sub

Private Sub chkMapMagiaSinEfecto_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    MapInfo.MagiaSinEfecto = chkMapMagiaSinEfecto.value
    MapInfo.Changed = 1
    
End Sub

Private Sub chkMapInviSinEfecto_LostFocus()
    '*************************************************
    'Author:
    'Last modified:
    '*************************************************
    MapInfo.InviSinEfecto = chkMapInviSinEfecto.value
    MapInfo.Changed = 1

End Sub

Private Sub chkMapnpc_LostFocus()
    '*************************************************
    'Author: Hardoz
    'Last modified: 28/08/2010
    '*************************************************
    MapInfo.RoboNpcsPermitido = ChkMapNpc.value
    MapInfo.Changed = 1
 
End Sub

Private Sub chkMapResuSinEfecto_LostFocus()
    '*************************************************
    'Author:
    'Last modified:
    '*************************************************
    MapInfo.ResuSinEfecto = chkMapResuSinEfecto.value
    MapInfo.Changed = 1

End Sub

Private Sub chkMapPK_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    MapInfo.PK = chkMapPK.value
    MapInfo.Changed = 1
    
End Sub

Private Sub chkOcultarSin_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2020
    '*************************************************
    MapInfo.OcultarSinEfecto = ChkMapNpc.value
    MapInfo.Changed = 1
    
End Sub

Private Sub cmdCerrar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Me.Hide
    
End Sub

Private Sub cmdMusica_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************
    frmSound.Show , frmMain
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
    
End Sub

Private Sub LvBActualizarLuces_Click()
    Call Actualizar_Estado
End Sub

Private Sub LvBGuardar_Click()
    Call guardarInfoZona(frmZonas.LstZona.ListIndex + 1)
End Sub

Public Sub guardarInfoZona(ByVal id As Integer)
    Dim i As Integer
    
    With MapZonas(id)
        .name = txtMapNombre.Text
        .Music = Val(txtMapMusica.Text)
        .Ambient = Val(TxtAmbient.Text)
        .PK = chkMapPK.value
        .MagiaSinEfecto = chkMapMagiaSinEfecto.value
        .InviSinEfecto = chkMapInviSinEfecto.value
        .ResuSinEfecto = chkMapResuSinEfecto.value
        .RoboNpcsPermitido = ChkMapNpc.value
        .InvocarSinEfecto = chkInvocarSin.value
        .OcultarSinEfecto = chkOcultarSin.value
        .Terreno = txtMapTerreno.Text
        .Zona = txtMapZona.Text
        .Restringir = txtMapRestringir.Text
        .LuzBase = .LuzBase
        .lvlMinimo = Val(TxtlvlMinimo.Text)
        .lvlMaximo = Val(txtNivelMaximo.Text)
        
    End With
    
    Call ActualizarZonaList(id - 1)

End Sub

Private Sub PicColorMap_Click()

    If chkLuzClimatica.value = False Then Exit Sub
    
    frmColorPicker.Show
End Sub

Private Sub txtAmbient_Change()
    '*************************************************
    'Author: Lorwik
    'Last modified: 10/08/14
    '*************************************************
    MapInfo.Ambient = TxtAmbient.Text
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapMusica_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    MapInfo.Music = txtMapMusica.Text
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapNombre_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    MapInfo.name = txtMapNombre.Text
    MapInfo.Changed = 1
    Call AddtoRichTextBox(frmConsola.StatTxt, "Nombre de mapa cambiado a:  " & MapInfo.name, 255, 255, 255, False, True, True, , True)
    
End Sub

Private Sub txtlvlminimo_LostFocus()
    '*************************************************
    'Author: Lorwik
    'Last modified: 13/09/11
    '*************************************************
    MapInfo.lvlMinimo = TxtlvlMinimo.Text
    MapInfo.Changed = 1
    
End Sub

Private Sub txtnivelmaximo_LostFocus()
    '*************************************************
    'Author: Lorwik
    'Last modified: 13/09/11
    '*************************************************
    MapInfo.lvlMaximo = txtNivelMaximo.Text
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapRestringir_KeyPress(KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    KeyAscii = 0
    
End Sub

Private Sub txtMapRestringir_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    MapInfo.Restringir = txtMapRestringir.Text
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapTerreno_KeyPress(KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    KeyAscii = 0
    
End Sub

Private Sub txtMapTerreno_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    MapInfo.Terreno = txtMapTerreno.Text
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapZona_KeyPress(KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    KeyAscii = 0
    
End Sub

Private Sub txtMapZona_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    MapInfo.Zona = txtMapZona.Text
    MapInfo.Changed = 1
    
End Sub

Public Sub CambiarColorMap()
    On Error GoTo PicColorMap_Err
    
    PicColorMap.BackColor = MapZonas(frmZonas.LstZona.ListIndex + 1).LuzBase
    
    frmMapInfo.PicColorMap.BackColor = PicColorMap.BackColor
    
    MapInfo.Changed = 1
    
    Exit Sub

PicColorMap_Err:
    Call RegistrarError(Err.Number, Err.Description, " FrmMapInfo.CambiarColorMap_Click", Erl)
    Resume Next
End Sub
