VERSION 5.00
Begin VB.Form frmTraslados 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Traslados"
   ClientHeight    =   2760
   ClientLeft      =   24495
   ClientTop       =   6465
   ClientWidth     =   4020
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
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   184
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   ShowInTaskbar   =   0   'False
   Begin WinterMapEditor.lvButtons_H cUnionManual 
      Height          =   405
      Left            =   300
      TabIndex        =   7
      Top             =   1770
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   714
      Caption         =   "Union con Mapa Adyacente (Manual)"
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
   Begin WinterMapEditor.lvButtons_H cInsertarTransOBJ 
      Height          =   450
      Left            =   1950
      TabIndex        =   5
      Top             =   660
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   794
      Caption         =   "Colorcar Obj. Automatico"
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
   Begin WinterMapEditor.lvButtons_H cQuitarTrans 
      Height          =   450
      Left            =   1950
      TabIndex        =   6
      Top             =   1170
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   794
      Caption         =   "Quitar Traslado"
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
   Begin WinterMapEditor.lvButtons_H cInsertarTrans 
      Height          =   450
      Left            =   1950
      TabIndex        =   4
      Top             =   150
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   794
      Caption         =   "Insertar Traslado"
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
   Begin VB.Frame FraDestino 
      BackColor       =   &H00535353&
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1485
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   1695
      Begin VB.TextBox tTMapa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   660
         TabIndex        =   0
         Text            =   "1"
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox tTX 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   660
         TabIndex        =   1
         Text            =   "1"
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox tTY 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   660
         TabIndex        =   3
         Text            =   "1"
         Top             =   990
         Width           =   915
      End
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   450
         TabIndex        =   11
         Top             =   1020
         Width           =   180
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   450
         TabIndex        =   10
         Top             =   630
         Width           =   180
      End
      Begin VB.Label lblMapa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   270
         Width           =   450
      End
   End
   Begin WinterMapEditor.lvButtons_H cUnionAuto 
      Height          =   405
      Left            =   300
      TabIndex        =   8
      Top             =   2220
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   714
      Caption         =   "Union con Mapa Adyacente (Auto)"
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
End
Attribute VB_Name = "frmTraslados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cUnionAuto_Click()
    frmUnionAdyacente.Show , frmMain
End Sub

Private Sub cUnionManual_Click()
    cInsertarTrans.value = (cUnionManual.value = True)
    'Call cInsertarTrans_Click
End Sub
