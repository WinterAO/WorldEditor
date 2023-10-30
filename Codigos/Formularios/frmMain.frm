VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00424242&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinterMapEditor"
   ClientHeight    =   10800
   ClientLeft      =   150
   ClientTop       =   690
   ClientWidth     =   19200
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
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   720
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMapConver 
      Caption         =   "Command1"
      Height          =   270
      Left            =   18750
      TabIndex        =   22
      Top             =   240
      Width           =   300
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   120
      Top             =   750
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox MainViewPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10110
      Left            =   30
      ScaleHeight     =   672
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1278
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   675
      Width           =   19200
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   0
      Left            =   30
      TabIndex        =   12
      Top             =   75
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
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
      ImgAlign        =   4
      Image           =   "frmMain.frx":10CA
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   1
      Left            =   540
      TabIndex        =   13
      Top             =   75
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
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
      Image           =   "frmMain.frx":1D1C
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   2
      Left            =   1050
      TabIndex        =   14
      Top             =   75
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
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
      ImgAlign        =   4
      Image           =   "frmMain.frx":296E
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   3
      Left            =   1560
      TabIndex        =   15
      Top             =   75
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
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
      ImgAlign        =   4
      Image           =   "frmMain.frx":35C0
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   4
      Left            =   2070
      TabIndex        =   16
      Top             =   75
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
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
      ImgAlign        =   4
      Image           =   "frmMain.frx":4212
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   5
      Left            =   2580
      TabIndex        =   17
      Top             =   75
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
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":4E64
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   6
      Left            =   3090
      TabIndex        =   18
      Top             =   75
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
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":5AB6
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   7
      Left            =   3600
      TabIndex        =   19
      Top             =   75
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
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":6138
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   8
      Left            =   4110
      TabIndex        =   20
      Top             =   75
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
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":65DA
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   9
      Left            =   4620
      TabIndex        =   21
      Top             =   75
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
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":68B2
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   10
      Left            =   5130
      TabIndex        =   23
      Top             =   75
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
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":C6904
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   11
      Left            =   5640
      TabIndex        =   24
      Top             =   75
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
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":F6956
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   12
      Left            =   6150
      TabIndex        =   25
      Top             =   90
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
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":1B69A8
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   13
      Left            =   6660
      TabIndex        =   26
      Top             =   90
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
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":1DBC1E
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   17910
      TabIndex        =   10
      Top             =   225
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   10260
      TabIndex        =   0
      Top             =   225
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   17145
      TabIndex        =   9
      Top             =   225
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   16380
      TabIndex        =   8
      Top             =   225
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   15615
      TabIndex        =   7
      Top             =   225
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   14850
      TabIndex        =   6
      Top             =   225
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   14085
      TabIndex        =   5
      Top             =   225
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   13320
      TabIndex        =   4
      Top             =   225
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   12555
      TabIndex        =   3
      Top             =   225
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   11790
      TabIndex        =   2
      Top             =   225
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   11025
      TabIndex        =   1
      Top             =   225
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNuevoMapa 
         Caption         =   "&Nuevo Mapa"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuArchivoLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuArchivoLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReAbrirMapa 
         Caption         =   "&Re-Abrir Mapa"
      End
      Begin VB.Menu mnuArchivoLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarMapa 
         Caption         =   "&Guardar Mapa"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGuardarMapaComo 
         Caption         =   "Guardar Mapa &como..."
      End
      Begin VB.Menu mnuArchivoLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edición"
      Begin VB.Menu mnuLineEdicion0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSeleccionado 
         Caption         =   "C&ortar Selección"
         Index           =   0
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuSeleccionado 
         Caption         =   "&Copiar Selección"
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSeleccionado 
         Caption         =   "&Pegar Selección"
         Index           =   2
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSeleccionado 
         Caption         =   "&Realizar Operación en Selección"
         Index           =   3
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSeleccionado 
         Caption         =   "Deshacer P&egado de Selección"
         Index           =   4
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuCopiarZonas 
         Caption         =   "Copiar zonas en CRTL + C"
      End
      Begin VB.Menu mnuLineEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCapturarSuperficie 
         Caption         =   "Auto-C&apturar información de la Superficie"
      End
      Begin VB.Menu mnuAutoCapturarTranslados 
         Caption         =   "Auto-&Capturar información de los Translados"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoGuardarMapas 
         Caption         =   "Configuración de Auto-&Guardar Mapas"
      End
      Begin VB.Menu mnuConfigAvanzadaSup 
         Caption         =   "Configuración Avanzada de Superficies"
         Index           =   3
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuLineEdicion2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInserDel 
         Caption         =   "Insertar/Eliminar"
         Begin VB.Menu mnuInsertarSuperficieEnTodo 
            Caption         =   "Inser. Superficie en todo el mapa"
         End
         Begin VB.Menu mnuInsertarSuperficieEnBordes 
            Caption         =   "Inser. Superficie en bordes del mapa"
         End
         Begin VB.Menu mnuBloquearBordes 
            Caption         =   "Inser. Bloqueos en bordes"
         End
         Begin VB.Menu mnuBloquearMapa 
            Caption         =   "Inser. Bloqueos en todo el mapa"
         End
         Begin VB.Menu mnuInsertarZonasEnBordes 
            Caption         =   "Inser. Zonas en bordes del mapa"
         End
         Begin VB.Menu mnuLineInserDel0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuitarSuperficieDeCapa 
            Caption         =   "Elim. Superficie capa seleccionada"
         End
         Begin VB.Menu QuitarSuperficieBordes 
            Caption         =   "Elim. Superficie en bordes del mapa"
         End
         Begin VB.Menu mnuDesbloquearBordes 
            Caption         =   "Elim. Bloqueos en bordes"
         End
         Begin VB.Menu mnuDesbloquearMapa 
            Caption         =   "Elim. Bloqueos en todo el mapa"
         End
         Begin VB.Menu mnuEliminarZona 
            Caption         =   "Elim. una zona del mapa"
         End
         Begin VB.Menu mnuEliminarNPCs 
            Caption         =   "Elim. NPCs de todo el mapa"
            Index           =   0
         End
         Begin VB.Menu mnuEliminarNPCs 
            Caption         =   "Elim. NPC's Hostiles de todo el mapa"
            Index           =   1
         End
         Begin VB.Menu mnuEliminarNPCs 
            Caption         =   "Elim. NPC's de la zona"
            Index           =   2
         End
         Begin VB.Menu mnuEliminarNPCs 
            Caption         =   "Elim. NPC's Hostiles de la zona"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuver 
      Caption         =   "&Ver"
      Begin VB.Menu mnuCapas 
         Caption         =   "&Capas"
         Begin VB.Menu mnuVerCapa1 
            Caption         =   "Capa &1 (Piso)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa2 
            Caption         =   "Capa &2 (costas, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa3 
            Caption         =   "Capa &3 (arboles, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa4 
            Caption         =   "Capa &4 (techos, etc)"
         End
      End
      Begin VB.Menu mnuVerTranslados 
         Caption         =   "...&Translados"
      End
      Begin VB.Menu mnuVerBloqueos 
         Caption         =   "...&Bloqueos"
      End
      Begin VB.Menu mnuVerNPCs 
         Caption         =   "...&NPC's"
      End
      Begin VB.Menu mnuVerObjetos 
         Caption         =   "...&Objetos"
      End
      Begin VB.Menu mnuVerTriggers 
         Caption         =   "...Tri&gger's"
      End
      Begin VB.Menu mnuVerGrilla 
         Caption         =   "...Gri&lla"
      End
      Begin VB.Menu mnuVerParticulas 
         Caption         =   "...Parti&culas"
      End
      Begin VB.Menu mnuvZonas 
         Caption         =   "Zonas"
         Begin VB.Menu mnuverZonas 
            Caption         =   "Zona Actual"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuverZonas 
            Caption         =   "Todas las Zonas"
            Index           =   1
         End
         Begin VB.Menu mnuverZonas 
            Caption         =   "Ocultar Zonas"
            Index           =   2
         End
      End
      Begin VB.Menu mnuLinMostrar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerAutomatico 
         Caption         =   "Control &Automaticamente"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuMinimapa 
      Caption         =   "Minimapa"
      Begin VB.Menu MinimapMenu 
         Caption         =   "-"
      End
      Begin VB.Menu Minimap 
         Caption         =   "Capa 1"
         Index           =   0
      End
      Begin VB.Menu Minimap 
         Caption         =   "Capa 2"
         Index           =   1
      End
      Begin VB.Menu Minimap 
         Caption         =   "Capa 3"
         Index           =   2
      End
      Begin VB.Menu Minimap 
         Caption         =   "Capa 4"
         Index           =   3
      End
      Begin VB.Menu Minimap 
         Caption         =   "NPC's"
         Index           =   4
      End
      Begin VB.Menu Minimap 
         Caption         =   "Objetos"
         Index           =   5
      End
      Begin VB.Menu Minimap 
         Caption         =   "Bloqueos"
         Index           =   6
      End
      Begin VB.Menu Minimap 
         Caption         =   "Particulas"
         Index           =   7
      End
      Begin VB.Menu Minimap 
         Caption         =   "Nº de mapa"
         Index           =   8
      End
      Begin VB.Menu Minimap 
         Caption         =   "Cuadrantes"
         Index           =   9
      End
      Begin VB.Menu Minimap 
         Caption         =   "Zonas"
         Index           =   10
      End
      Begin VB.Menu Dibujarmini 
         Caption         =   "Dibujar"
      End
      Begin VB.Menu mnuLineMinimap 
         Caption         =   "-"
      End
      Begin VB.Menu minimapaSave 
         Caption         =   "Guardar Minimapa"
      End
      Begin VB.Menu saveAllMinimap 
         Caption         =   "Guardar todos los Minimapas"
      End
   End
   Begin VB.Menu mnuFunciones 
      Caption         =   "Funciones"
      Begin VB.Menu mnuLineFunciones0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptimizar 
         Caption         =   "Optimizar"
      End
      Begin VB.Menu mnuInformes 
         Caption         =   "Informes"
      End
      Begin VB.Menu mnuModoCaminata 
         Caption         =   "Modo Caminata"
      End
      Begin VB.Menu mnuzonanula 
         Caption         =   "Buscar zonas nulas"
      End
   End
   Begin VB.Menu mnusobre 
      Caption         =   "Sobre..."
      Begin VB.Menu mnuFormatos 
         Caption         =   "Formatos de Mapa"
      End
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMapConver_Click()
    frmMain.Dialog.CancelError = True
    
    DeseaGuardarMapa frmMain.Dialog.filename
    
    frmMain.ObtenerNombreArchivo False
    
    If Len(frmMain.Dialog.filename) < 3 Then Exit Sub
    
    If WalkMode = True Then Call modGeneral.ToggleWalkMode
        
    Call modMapas.NuevoMapa
        
    Call ModMapConver.Cargar_ConverCSM(frmMain.Dialog.filename)
        
    DoEvents
    frmMain.mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
    Exit Sub
End Sub

Private Sub LvBEdit_Click(Index As Integer)
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2020
    '*************************************************

    Select Case Index
    
        Case 0 ' Superficies

            If LvBEdit(Index).value Then
                frmSuperficies.Show , frmMain
            
            Else
                frmSuperficies.Visible = False
                
            End If
            
        Case 1 ' Traslados

            If LvBEdit(Index).value Then
                frmTraslados.Show , frmMain
            
            Else
                frmTraslados.Visible = False
                
            End If
        
        Case 2 ' Bloqueos

            If LvBEdit(Index).value Then
                frmBloqueos.Show , frmMain
            
            Else
                frmBloqueos.Visible = False
                
            End If
        
        Case 3 ' NPCs

            If LvBEdit(Index).value Then
                frmNPCs.Show , frmMain
            
            Else
                frmNPCs.Visible = False
                
            End If
        
        Case 4 ' Objetos

            If LvBEdit(Index).value Then
                frmOBJs.Show , frmMain
            
            Else
                frmOBJs.Visible = False
                
            End If
        
        Case 5 ' Triggers

            If LvBEdit(Index).value Then
                frmTriggers.Show , frmMain
            
            Else
                frmTriggers.Visible = False
                
            End If
        
        Case 6 ' Particulas

            If LvBEdit(Index).value Then
                frmParticulas.Show , frmMain
            
            Else
                frmParticulas.Visible = False
                
            End If
        
        Case 7 ' Luces

            If LvBEdit(Index).value Then
                frmLuces.Show , frmMain
            
            Else
                frmLuces.Visible = False
                
            End If
        
        Case 8 ' Informacion

            If LvBEdit(Index).value Then
                frmMapInfo.Show , frmMain
            
            Else
                frmMapInfo.Visible = False
                
            End If
            
        Case 9 ' Preview

            If LvBEdit(Index).value Then
                frmPreview.Show , frmMain
                
            Else
                frmPreview.Visible = False
            End If
            
        Case 10 ' Zonas
            If LvBEdit(Index).value Then
                frmZonas.Show , frmMain
                
            Else
                frmZonas.Visible = False
            End If
            
        Case 11 ' Consola
            If LvBEdit(Index).value Then
                frmConsola.Show , frmMain
                
            Else
                frmConsola.Visible = False
            End If
            
        Case 12 ' Rellenar Area
            If LvBEdit(Index).value Then
                frmRellenar.Show , frmMain
                
            Else
                frmRellenar.Visible = False
            End If
            
        Case 13 ' Mapa
            If LvBEdit(Index).value Then
                frmMapa.Show , frmMain
                
            Else
                frmMapa.Visible = True
            End If
    
    End Select

End Sub

Private Sub MapPest_Click(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/03/2021
    'Lorwik> Ahora distingue entre csm y map
    '*************************************************
    Dim Formato As String
    
    Formato = ".csm"
            
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then Call modMapas.GuardarMapa(Dialog.filename)

    End If
        
    If (Index + NumMap_Save - 4) <> NumMap_Save Then
        Dialog.CancelError = True

        On Error GoTo errhandler

        Dialog.filename = PATH_Save & NameMap_Save & (Index + NumMap_Save - 4) & Formato
        
        Call modMapas.NuevoMapa
        
        DoEvents
        
        Call abrirCargarMapa(frmMain.Dialog.filename)
        
        EngineRun = True
        
    End If
    
    Exit Sub
    
errhandler:
    Call MsgBox(Err.Description)
    
End Sub

Private Sub Minimap_Click(Index As Integer)
    '*************************************************
    'Author: Lorwik
    'Last modified: 29/09/2023
    '*************************************************

    On Error GoTo Minimap_Err
    
    Select Case Index
    
        Case 0 'Capa 1
            MMiniMap_capa1 = Not MMiniMap_capa1
            
        Case 1 'Casa 2
            MMiniMap_capa2 = Not MMiniMap_capa2
            
        Case 2 'Capa 3
            MMiniMap_capa3 = Not MMiniMap_capa3
        
        Case 3 'Capa 4
            MMiniMap_capa4 = Not MMiniMap_capa4
            
        Case 4 'NPC's
            MMiniMap_Npcs = Not MMiniMap_Npcs
            
        Case 5 'Objetos
            MMiniMap_objetos = Not MMiniMap_objetos
            
        Case 6 'Bloqueos
            MMiniMap_Bloqueos = Not MMiniMap_Bloqueos
        
        Case 7 'Particulas
            MMiniMap_particulas = Not MMiniMap_particulas
            
        Case 8 'Nº Mapa
            MMiniMap_Nombre = Not MMiniMap_Nombre
        
        Case 9 'Cuadrantes
            MMiniMap_cuadrantes = Not MMiniMap_cuadrantes
            
        Case 10 'Zonas
            MMiniMap_Zonas = Not MMiniMap_Zonas
            
    End Select
    
    Minimap(Index).Checked = (Minimap(Index).Checked = False)
    Call DibujarMinimapa
    Call guardarPerfilMinimap
    
    Exit Sub
    
Minimap_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Minimap_Click", Erl)
    Resume Next
End Sub

Private Sub mnuAbrirMapa_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 25/04/2020
    '*************************************************

    Call AbrirMapa
    
End Sub

Private Sub mnuAcerca_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************

    frmAcercade.Show , frmMain
End Sub

Private Sub mnuBloquearBordes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Bloquear_Bordes
End Sub

Private Sub mnuBloquearMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Bloqueo_Todo(1)
End Sub

Private Sub mnuConfigAvanzadaSup_Click(Index As Integer)
    '*************************************************
    'Author: Lorwik
    'Last modified: 30/10/2023
    '*************************************************
    
    frmConfigSup.Show , frmMain
End Sub

Private Sub mnuCopiarZonas_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 30/10/2023
    '*************************************************
    
    mnuCopiarZonas.Checked = Not mnuCopiarZonas.Checked
End Sub

Private Sub mnuDesbloquearMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/2022
    '*************************************************
    Call modEdicion.Bloqueo_Todo(0)
End Sub

Private Sub mnuDesbloquearBordes_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 20/05/2022
    '*************************************************
    Call modEdicion.Desbloquear_Bordes
End Sub

Private Sub mnuEliminarNPCs_Click(Index As Integer)
    '*************************************************
    'Author: Lorwik
    'Last modified: 14/10/2023
    '*************************************************
    On Error GoTo mnuQuitarNPCs_Click_Err
    
    Select Case Index
    
        Case 0 'Eliminar todos los NPC's del mapa
            Call modEdicion.Quitar_NPCs(False, False)
        
        Case 1 'Eliminar todos los NPC's hostiles del mapa
            Call modEdicion.Quitar_NPCs(True, False)
            
        Case 2 'Eliminar todos los NPC's de la zona
            Call modEdicion.Quitar_NPCs(False, True)
            
        Case 3 'Eliminar todos los NPC's Hostiles de la zona
            Call modEdicion.Quitar_NPCs(False, True)
    
    End Select
    
    Exit Sub

mnuQuitarNPCs_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarNPCs_Click", Erl)
    Resume Next
End Sub

Private Sub mnuEliminarZona_Click()
    Dim zonaDel As Integer
    Dim X, y As Integer
    zonaDel = InputBox("Por favor, ingrese el número de la zona a borrar:")
    
    For X = XMinMapSize To XMaxMapSize
    
        For y = YMinMapSize To YMaxMapSize
        
            If MapData(X, y).ZonaIndex = zonaDel Then MapData(X, y).ZonaIndex = 0
        
        Next y
        
    Next X

End Sub

Private Sub mnuFormatos_Click()
    frmFormatos.Show , frmMain
End Sub

Private Sub mnuInformes_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************

    frmInformes.Show , frmMain

End Sub

Private Sub mnuInsertarSuperficieEnBordes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Superficie_Bordes
End Sub

Private Sub mnuQuitarSuperficieBordes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Quitar_Bordes
End Sub

Private Sub mnuInsertarSuperficieEnTodo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Superficie_Todo
End Sub

Private Sub mnuInsertarZonasEnBordes_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 25/05/2022
    '*************************************************
    Call modEdicion.Zonas_Bordes
End Sub

Private Sub mnuMapTam_Click(Index As Integer)

End Sub

Private Sub mnuQuitarSuperficieDeCapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Quitar_Capa(frmSuperficies.cCapas.Text)
End Sub

Private Sub mnuModoCaminata_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************

    Call ToggleWalkMode
    
End Sub

Private Sub mnuNuevoMapa_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 29/04/2021
    '*************************************************
    On Error Resume Next
    Dim LoopC As Integer
    
    DeseaGuardarMapa Dialog.filename
    
    For LoopC = 0 To frmMain.MapPest.Count - 1
        frmMain.MapPest(LoopC).Visible = False
    Next
    
    frmMain.Dialog.filename = Empty
    
    If WalkMode Then Call modGeneral.ToggleWalkMode
    
    Call modMapas.NuevoMapa
    
    Call LvBEdit_Click(9)

End Sub

Public Sub mnuGuardarMapa_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 29/04/2021
    '*************************************************

    On Error GoTo mnuGuardarMapa_Click_Err

    modMapas.GuardarMapa Dialog.filename
    
    Exit Sub

mnuGuardarMapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGuardarMapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGuardarMapaComo_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 29/04/2021
    '*************************************************
    
    On Error GoTo mnuGuardarMapaComo_Click_Err
    
    modMapas.GuardarMapa
    
    Exit Sub

mnuGuardarMapaComo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGuardarMapaComo_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuOptimizar_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************

    frmOptimizar.Show , frmMain
    
End Sub

Private Sub mnuReAbrirMapa_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************
    On Error GoTo mnuReAbrirMapa_Click_Err
 
    If FileExist(Dialog.filename, vbArchive) = False Then Exit Sub
    
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapas.GuardarMapa Dialog.filename
        End If
    End If
    
    Call modMapas.NuevoMapa
    
    Call abrirCargarMapa(frmMain.Dialog.filename)
    
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
    Exit Sub

mnuReAbrirMapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuReAbrirMapa_Click", Erl)
    Resume Next
End Sub

Private Sub mnuSalir_Click()
    Call CloseMapEditor
    
End Sub

Private Sub mnuSeleccionado_Click(Index As Integer)
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************

    On Error GoTo mnuSeleccionado_Click_Err

    Select Case Index
    
        Case 0 ' Cortar
            Call CortarSeleccion
        
        Case 1 ' Copiar
            Call CopiarSeleccion
        
        Case 2 ' Pegar
            Call PegarSeleccion
        
        Case 3 ' Realizar
            Call AccionSeleccion
        
        Case 4 ' Deshacer
            Call DePegar
            
    End Select
    
    Exit Sub

mnuSeleccionado_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuSeleccionado_Click", Erl)
    Resume Next
    
End Sub

Private Sub Dibujarmini_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2021
    '*************************************************

    Call DibujarMinimapa
End Sub

Public Sub ObtenerNombreArchivo(ByVal Guardar As Boolean)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    On Error GoTo ObtenerNombreArchivo_Err
    
    With Dialog
        .Filter = "Mapas del nuevo formato (*.csm)|*.csm|Mapas clasicos de Argentum Online (*.map)|*.map"

        If Guardar Then
            .DialogTitle = "Guardar"
            .DefaultExt = ".txt"
            .filename = vbNullString
            .flags = cdlOFNPathMustExist
            .ShowSave
        Else
            .DialogTitle = "Cargar"
            .filename = vbNullString
            .flags = cdlOFNFileMustExist
            .ShowOpen
            
        End If
    End With
    
    Exit Sub
    
ObtenerNombreArchivo_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.ObtenerNombreArchivo", Erl)
    Resume Next
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  y As Single)
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2021
    '*************************************************

    Call Form_MouseMove(Button, Shift, X, y)
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  y As Single)
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2021
    '*************************************************

    Call Form_MouseDown(Button, Shift, X, y)
End Sub

Private Sub MainViewPic_DblClick()
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2021
    '*************************************************

    If Not MapaCargado Then Exit Sub
    
    If SobreX > 0 And SobreY > 0 Then
        DobleClick Val(SobreX), Val(SobreY)
        
    End If
End Sub

Private Sub Form_Load()

    Dim i As Byte

    Me.Caption = Form_Caption
    
    mnuVerBloqueos.Checked = VerBlockeados
    mnuVerTranslados.Checked = VerTranslados
    mnuVerNPCs.Checked = VerNpcs
    mnuVerObjetos.Checked = VerObjetos
    mnuVerTriggers.Checked = VerTriggers
    mnuVerParticulas.Checked = VerParticulas
    mnuVerCapa1.Checked = VerCapa1
    mnuVerCapa2.Checked = VerCapa2
    mnuVerCapa3.Checked = VerCapa3
    mnuVerCapa4.Checked = VerCapa4
    mnuVerGrilla.Checked = VerGrilla

End Sub

Private Sub Form_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2021
    '*************************************************

    Me.SetFocus

End Sub

Private Sub Form_Resize()
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2021
    '*************************************************

    On Error Resume Next

    With MainViewPic
    
        .Height = Me.ScaleHeight - 50
        .Width = Me.ScaleWidth
    
    End With
    
    With MainScreenRect
        .Bottom = frmMain.MainViewPic.ScaleHeight
        .Right = frmMain.MainViewPic.ScaleWidth
    End With
    
    Call ChangeView
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2021
    '*************************************************

    Call CloseMapEditor
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2021
    '*************************************************
    ' HotKeys
    If HotKeysAllow = False Then Exit Sub
    
    Select Case UCase(Chr(KeyAscii))

        Case "S" 'Superficies
            Call LvBEdit_Click(0)
            
    End Select
    
End Sub

Private Sub Form_DblClick()
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2021
    '*************************************************
    Dim tX As Integer
    Dim tY As Integer
    
    If Not MapaCargado Then Exit Sub
    
    If SobreX > 0 And SobreY > 0 Then
        DobleClick Val(SobreX), Val(SobreY)
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2021
    '*************************************************

    Dim tX As Integer
    Dim tY As Integer
    
    If Not MapaCargado Then Exit Sub
    
    Call ConvertCPtoTP(X, y, tX, tY)
    
    If EstadoSelect > 0 And Button = 2 Then
        EstadoSelect = 0
    End If
    
    If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionIX = tX '+ UserPos.X
        SeleccionIY = tY '+ UserPos.Y
        frmRellenar.DX1.Text = tX
        frmRellenar.DY1.Text = tY
        
    Else
        Call modEdit.ClickEdit(Button, tX, tY)
        
    End If
    
    Debug.Print Button

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2021
    '*************************************************
    Dim tX As Integer
    Dim tY As Integer
    
    'Make sure map is loaded
    If Not MapaCargado Then Exit Sub
    HotKeysAllow = True

    Call ConvertCPtoTP(X, y, tX, tY)
    
    MousePos = "X: " & tX & " - Y: " & tY
    
    If EstadoSelect > 0 And Button = 2 Then
        EstadoSelect = 0
        CopyX = tX
        CopyY = tY
    End If
    
    If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionFX = tX '+ TileX
        SeleccionFY = tY '+ TileY
        frmRellenar.DX2.Text = tX
        frmRellenar.DY2.Text = tY
        
        ContadorTiles = (frmRellenar.DX2.Text - frmRellenar.DX1.Text + 1) * (frmRellenar.DY2.Text - frmRellenar.DY1.Text + 1)
        
    Else
        Call modEdit.ClickEdit(Button, tX, tY)
        
    End If
    
End Sub

Private Sub mnuVerBloqueos_Click()
    VerBlockeados = Not VerBlockeados
End Sub

Private Sub mnuVerCapa1_Click()
    VerCapa1 = Not VerCapa1
    mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)
End Sub

Private Sub mnuVerCapa2_Click()
    VerCapa2 = Not VerCapa2
    mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)
End Sub

Private Sub mnuVerCapa3_Click()
    VerCapa3 = Not VerCapa3
    mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)
End Sub

Private Sub mnuVerCapa4_Click()
    VerCapa4 = Not VerCapa4
    mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)
End Sub

Private Sub mnuVerGrilla_Click()
    VerGrilla = Not VerGrilla
    mnuVerGrilla.Checked = VerGrilla
End Sub

Private Sub mnuVerNPCs_Click()
    VerNpcs = Not VerNpcs
    mnuVerNPCs.Checked = (mnuVerNPCs.Checked = False)
End Sub

Private Sub mnuVerObjetos_Click()
    VerObjetos = Not VerObjetos
    mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)
End Sub

Private Sub mnuVerParticulas_Click()
    VerParticulas = Not VerParticulas
    mnuVerParticulas.Checked = (mnuVerParticulas.Checked = False)
End Sub

Private Sub mnuVerTranslados_Click()
    VerTranslados = Not VerTranslados
    mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)
End Sub

Private Sub mnuVerTriggers_Click()
    VerTriggers = Not VerTriggers
    mnuVerTriggers.Checked = (mnuVerTriggers.Checked = False)
End Sub

Private Sub mnuVerZonas_Click(Index As Integer)
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/04/2021
    '*************************************************

    mnuverZonas(Index).Checked = True

    Select Case Index
    
        Case 0
            mnuverZonas(1).Checked = False
            mnuverZonas(2).Checked = False
        
        Case 1
            mnuverZonas(0).Checked = False
            mnuverZonas(2).Checked = False
        
        Case 2
            mnuverZonas(0).Checked = False
            mnuverZonas(1).Checked = False
    
    End Select

End Sub

Private Sub mnuzonanula_Click()
    Dim X As Integer
    Dim y As Integer
    
    For X = XMinMapSize To XMaxMapSize
    
        For y = YMinMapSize To YMaxMapSize
        
            If MapData(X, y).ZonaIndex = 0 Then
                MsgBox "Se ha encontrado una zona nula en la posicion X: " & X & " Y: " & y
                Exit Sub
            End If
        
        Next y
    
    Next X
    
End Sub

Private Sub saveAllMinimap_Click()
    Call RenderizarCuadrantes
End Sub
