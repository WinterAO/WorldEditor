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
   Begin VB.Timer TimerMinuto 
      Interval        =   60000
      Left            =   690
      Top             =   780
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
      Top             =   660
      Width           =   19200
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   0
      Left            =   30
      TabIndex        =   12
      ToolTipText     =   "Superficies"
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
      ToolTipText     =   "Traslados"
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
      ToolTipText     =   "Bloqueos"
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
      ToolTipText     =   "NPC's"
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
      ToolTipText     =   "Objetos"
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
      ToolTipText     =   "Trigger's"
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
      ToolTipText     =   "Particulas"
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
      ToolTipText     =   "Luces"
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
      Left            =   4620
      TabIndex        =   20
      ToolTipText     =   "Preview"
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
      Left            =   4110
      TabIndex        =   21
      ToolTipText     =   "Editor de zonas"
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
      Image           =   "frmMain.frx":C662C
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   10
      Left            =   5130
      TabIndex        =   22
      ToolTipText     =   "Consola"
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
      Image           =   "frmMain.frx":F667E
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   11
      Left            =   5640
      TabIndex        =   23
      ToolTipText     =   "Rellenar Area"
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
      Image           =   "frmMain.frx":1B66D0
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   12
      Left            =   6150
      TabIndex        =   24
      ToolTipText     =   "Mapa"
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
      Image           =   "frmMain.frx":1DB946
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   13
      Left            =   6660
      TabIndex        =   25
      ToolTipText     =   "Velocidad"
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
      Image           =   "frmMain.frx":2D7658
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H LvBEdit 
      Height          =   480
      Index           =   14
      Left            =   7170
      TabIndex        =   26
      ToolTipText     =   "Inventario"
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
      Image           =   "frmMain.frx":2D795E
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   702
      X2              =   702
      Y1              =   4
      Y2              =   42
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000004&
      X1              =   698
      X2              =   698
      Y1              =   4
      Y2              =   42
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
      Left            =   18330
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
      Left            =   10680
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
      Left            =   17565
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
      Left            =   16800
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
      Left            =   16035
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
      Left            =   15270
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
      Left            =   14505
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
      Left            =   13740
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
      Left            =   12975
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
      Left            =   12210
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
      Left            =   11445
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
      Begin VB.Menu mnuAbrirMapaCien 
         Caption         =   "Abrir Mapa 100 x 100"
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
      Begin VB.Menu mnuArchivoLine6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExportar 
         Caption         =   "Importar / Exportar"
         Begin VB.Menu mnuImportarZonas 
            Caption         =   "Importar zonas"
         End
         Begin VB.Menu mnuExportar1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExportarZonas 
            Caption         =   "Exportar zonas"
         End
      End
      Begin VB.Menu mnuArchivoLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "Configuración"
      End
      Begin VB.Menu mnuArchivoLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   +{DEL}
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
         Begin VB.Menu mnuZonasxCuadrantes 
            Caption         =   "Inser. Zonas por cuadrantes"
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
            Caption         =   "Elim. Una zona del mapa"
         End
         Begin VB.Menu mnuEliminarZonas 
            Caption         =   "Elim. Zonas en todo el mapa"
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
         Begin VB.Menu mnuLineInserDe1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEliminarLuces 
            Caption         =   "Elim. Luces de la Selección"
         End
      End
   End
   Begin VB.Menu mnuMapa 
      Caption         =   "Mapa"
      Begin VB.Menu mnuOptimizar 
         Caption         =   "Optimizar"
      End
      Begin VB.Menu mnuInformes 
         Caption         =   "Informes"
      End
      Begin VB.Menu mnuzonanula 
         Caption         =   "Buscar zonas nulas"
      End
      Begin VB.Menu mnuzonasinuso 
         Caption         =   "Buscar zonas sin uso"
      End
      Begin VB.Menu mnuMapSize 
         Caption         =   "Tamaño de Mapa.."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuver 
      Caption         =   "&Ver"
      Begin VB.Menu mnuCapas 
         Caption         =   "&Capas"
         Begin VB.Menu mnuVerCapa1 
            Caption         =   "Capa &1 (Piso)"
            Checked         =   -1  'True
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnuVerCapa2 
            Caption         =   "Capa &2 (costas, etc)"
            Checked         =   -1  'True
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuVerCapa3 
            Caption         =   "Capa &3 (arboles, etc)"
            Checked         =   -1  'True
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnuVerCapa4 
            Caption         =   "Capa &4 (techos, etc)"
            Shortcut        =   ^{F4}
         End
      End
      Begin VB.Menu mnuVerTraslados 
         Caption         =   "...&Traslados"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuVerBloqueos 
         Caption         =   "...&Bloqueos"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuVerNPCs 
         Caption         =   "...&NPC's"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu mnuVerObjetos 
         Caption         =   "...&Objetos"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnuVerTriggers 
         Caption         =   "...Tri&gger's"
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu mnuVerGrilla 
         Caption         =   "...Gri&lla"
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu mnuVerParticulas 
         Caption         =   "...Parti&culas"
         Shortcut        =   ^{F12}
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
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLinMostrar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimapa 
         Caption         =   "Minimapa"
         Begin VB.Menu MinimapMenu 
            Caption         =   "-"
         End
         Begin VB.Menu Minimap 
            Caption         =   "Ver capa 1"
            Index           =   0
         End
         Begin VB.Menu Minimap 
            Caption         =   "Ver capa 2"
            Index           =   1
         End
         Begin VB.Menu Minimap 
            Caption         =   "Ver capa 3"
            Index           =   2
         End
         Begin VB.Menu Minimap 
            Caption         =   "Ver capa 4"
            Index           =   3
         End
         Begin VB.Menu Minimap 
            Caption         =   "Ver NPC's"
            Index           =   4
         End
         Begin VB.Menu Minimap 
            Caption         =   "Ver Objetos"
            Index           =   5
         End
         Begin VB.Menu Minimap 
            Caption         =   "Ver Bloqueos"
            Index           =   6
         End
         Begin VB.Menu Minimap 
            Caption         =   "Ver Particulas"
            Index           =   7
         End
         Begin VB.Menu Minimap 
            Caption         =   "Ver Nº de mapa"
            Enabled         =   0   'False
            Index           =   8
         End
         Begin VB.Menu Minimap 
            Caption         =   "Ver Cuadrantes"
            Index           =   9
         End
         Begin VB.Menu Minimap 
            Caption         =   "Ver Zonas"
            Index           =   10
         End
         Begin VB.Menu mnuLineMinimap 
            Caption         =   "-"
         End
         Begin VB.Menu Dibujarmini 
            Caption         =   "Dibujar"
         End
         Begin VB.Menu minimapaSave 
            Caption         =   "Guardar Minimapa"
         End
         Begin VB.Menu saveAllMinimap 
            Caption         =   "Guardar todos los Minimapas"
         End
      End
   End
   Begin VB.Menu mnuFunciones 
      Caption         =   "Funciones"
      Begin VB.Menu mnuLineFunciones0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScreenshot 
         Caption         =   "Captura de pantalla"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuRender 
         Caption         =   "Renderizar"
      End
      Begin VB.Menu mnuModoCaminata 
         Caption         =   "Modo Caminata"
      End
      Begin VB.Menu mnuGrhList 
         Caption         =   "Lista de todos los Grh's"
      End
      Begin VB.Menu mnuline0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGrhtoPNG 
         Caption         =   "Grh -> PNG"
      End
      Begin VB.Menu mnuPNGtoGrh 
         Caption         =   "PNG -> Grh"
      End
   End
   Begin VB.Menu mnuRecargar 
      Caption         =   "Recargar"
      Begin VB.Menu mnuRecargarGraficos 
         Caption         =   "Indice de Graficos"
      End
      Begin VB.Menu mnuRecargarCuerpos 
         Caption         =   "Indice de Cuerpos"
      End
      Begin VB.Menu mnuRecargarCabezas 
         Caption         =   "Indice de Cabezas"
      End
      Begin VB.Menu mnuRecargarCascos 
         Caption         =   "Indice de Cascos"
      End
      Begin VB.Menu mnuRecargarEscudos 
         Caption         =   "Indice de Escudos"
      End
      Begin VB.Menu mnuRecargarParticulas 
         Caption         =   "Indice de Particulas"
      End
      Begin VB.Menu mnuRecargarObj 
         Caption         =   "Obj.dat"
      End
      Begin VB.Menu mnuRecargarNPCS 
         Caption         =   "Npcs.dat"
      End
      Begin VB.Menu mnuRecargarIndices 
         Caption         =   "Indices de Superficies"
      End
      Begin VB.Menu mnuRecargarTriggers 
         Caption         =   "Indices de Triggers"
      End
   End
   Begin VB.Menu mnusobre 
      Caption         =   "Sobre..."
      Begin VB.Menu mnuFormatos 
         Caption         =   "Formatos de Mapa"
         Enabled         =   0   'False
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
            
        Case 8 ' Preview

            If LvBEdit(Index).value Then
                frmPreview.Show , frmMain
                
            Else
                frmPreview.Visible = False
            End If
            
        Case 9 ' Zonas
            If LvBEdit(Index).value Then
                frmZonas.Show , frmMain
                
            Else
                frmZonas.Visible = False
            End If
            
        Case 10 ' Consola
            If LvBEdit(Index).value Then
                frmConsola.Show , frmMain
                
            Else
                frmConsola.Visible = False
            End If
            
        Case 11 ' Rellenar Area
            If LvBEdit(Index).value Then
                frmRellenar.Show , frmMain
                
            Else
                frmRellenar.Visible = False
            End If
            
        Case 12 ' Mapa
            If LvBEdit(Index).value Then
                frmMapa.Show , frmMain
                
            Else
                frmMapa.Visible = False
            End If
            
        Case 13 ' Velocidad
            If LvBEdit(Index).value Then
                frmWalkerSpeed.Show , frmMain
                
            Else
                frmWalkerSpeed.Visible = False
            End If
            
        Case 14 'Quick Superficies
            If LvBEdit(Index).value Then
                frmQuick.Show , frmMain
                
            Else
                frmQuick.Visible = False
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
    
    Minimap(Index).Checked = Not Minimap(Index).Checked
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

    Call AbrirMapa(True)
    
End Sub

Private Sub mnuAbrirMapaCien_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 25/04/2020
    '*************************************************

    Call AbrirMapa(False)
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

Private Sub mnuConfig_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 31/03/2024
    '*************************************************
    
    frmConfiguracion.Show , frmMain
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

Private Sub mnuEliminarLuces_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 03/11/2023
    '*************************************************
    On Error GoTo mnuEliminarLuces_Click_Err
    
    Dim x As Integer
    Dim y As Integer
    
    For y = SeleccionIY To SeleccionFY
        For x = SeleccionIX To SeleccionFX
        
            With MapData(x, y)
            
                Call Long_2_RGBAList(MapData(x, y).Light_Value(), -1)
                
                .Light.active = False
                .Light.range = 0
                .Light.map_x = 0
                .Light.map_y = 0
                .Light.RGBCOLOR.A = 0
                .Light.RGBCOLOR.R = 0
                .Light.RGBCOLOR.G = 0
                .Light.RGBCOLOR.B = 0
                
            End With
        
        Next x
        
    Next y
    
    Exit Sub
    
mnuEliminarLuces_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuEliminarLuces_Click", Erl)
    Resume Next
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
    '*************************************************
    'Author: Lorwik
    'Last modified: ???
    '*************************************************
    On Error GoTo mnuEliminarZona_Click_Err
    
    Call EliminarUnaZona

    Exit Sub

mnuEliminarZona_Click_Err:

    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuEliminarZona_Click", Erl)
    Resume Next
End Sub

Private Sub mnuEliminarZonas_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 08/06/2024
    '*************************************************
    On Error GoTo mnuEliminarZonas_Click_Err
    
    Call EliminarZonas
    
    Exit Sub

mnuEliminarZonas_Click_Err:

    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuEliminarZonas_Click", Erl)
    Resume Next
End Sub

Private Sub mnuImportarZonas_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 08/06/2024
    '*************************************************
    On Error GoTo mnuImportarZonas_Click_Err
    
    Dim path As String
    
    frmMain.Dialog.CancelError = True
    
    ObtenerNombreArchivo False, True
    
    path = frmMain.Dialog.filename

    If LenB(path) = 0 Then Exit Sub
    
    If modMapasWAO.Importar_Zonas(path) Then Call ShowMessageScreen("Zonas importadas.")
    
    If MMiniMap_Zonas Then Call DibujarMinimapa

    Exit Sub

mnuImportarZonas_Click_Err:

    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuImportarZonas_Click", Erl)
    Resume Next
End Sub

Private Sub mnuExportarZonas_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 17/05/2024
    '*************************************************
    On Error GoTo mnuExportarZonas_Click_Err
    
    Dim path As String
    
    frmMain.Dialog.CancelError = True
    
    ObtenerNombreArchivo True, True
    
    path = frmMain.Dialog.filename

    If LenB(path) = 0 Then Exit Sub
    
    If modMapasWAO.Exportar_Zonas(path) Then Call ShowMessageScreen("Zonas exportadas.")
    
    Exit Sub

mnuExportarZonas_Click_Err:

    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuExportarZonas_Click", Erl)

    Resume Next

End Sub

Private Sub mnuFormatos_Click()
    frmFormatos.Show , frmMain
End Sub

Private Sub mnuGrhList_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 08/11/2023
    '*************************************************
    On Error GoTo mnuGrhList_Click_Err
    
        frmGrh.Show , frmMain
    
    Exit Sub

mnuGrhList_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuGrhList_Click", Erl)
    Resume Next
End Sub

Private Sub mnuGrhtoPNG_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 08/11/2023
    '*************************************************
    
    On Error GoTo mnuGrhtoPNG_Click_Err
    
    Dim GrhIndex As Long
    
    GrhIndex = InputBox("Numero de Grh")
    
    If IsNumeric(grhCount) = True Then
        If GrhIndex > grhCount Then Exit Sub
        If GrhIndex < 1 Then Exit Sub
        MsgBox GrhData(GrhIndex).FileNum, vbInformation

    End If

    Exit Sub

mnuGrhtoPNG_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuGrhtoPNG_Click", Erl)
    Resume Next
    
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

Private Sub mnuMapSize_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 24/04/2024
    '*************************************************
    
    frmMapSize.Show , frmMain
    
End Sub

Private Sub mnuPNGtoGrh_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 08/11/2023
    '*************************************************
    
    On Error GoTo mnuPNGtoGrh_Click_Err
    
    Dim PNGIndex As Long
    Dim GrhIndex As Long
    Dim Count As Long
    
    PNGIndex = InputBox("Numero de PNG")
    Count = 1
    
    If IsNumeric(grhCount) = True Then
    
        Do While Count < grhCount
        
            If GrhData(Count).FileNum = PNGIndex Then Exit Do
        
            Count = Count + 1
        
        Loop
    
        MsgBox Count, vbInformation

    End If

    Exit Sub

mnuPNGtoGrh_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuPNGtoGrh_Click", Erl)
    Resume Next
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
    Dim loopc As Integer
    
    DeseaGuardarMapa Dialog.filename
    
    For loopc = 0 To frmMain.MapPest.Count - 1
        frmMain.MapPest(loopc).Visible = False
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

Private Sub mnuRecargarGraficos_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 05/07/2024
    '*************************************************
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Recargando Indices.ini...", 0, 0, 255, , , , , True)
    
    Call CargarIndicesSuperficie
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Indices.ini recargado!", 0, 255, 0, , , , , True)
End Sub

Private Sub mnuRecargarCuerpos_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 05/07/2024
    '*************************************************
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Recargando Personajes.ind...", 0, 0, 255, , , , , True)
    
    Call CargarCuerpos
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Personajes.ind recargado!", 0, 255, 0, , , , , True)
End Sub

Private Sub mnuRecargarCabezas_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 05/07/2024
    '*************************************************
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Recargando Head.ind...", 0, 0, 255, , , , , True)
    
    Call CargarCabezas
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Head.ind recargado!", 0, 255, 0, , , , , True)
End Sub

Private Sub mnuRecargarCascos_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 05/07/2024
    '*************************************************
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Recargando Helmet.ind...", 0, 0, 255, , , , , True)
    
    Call CargarCascos
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Helmet.ind recargado!", 0, 255, 0, , , , , True)
End Sub

Private Sub mnuRecargarEscudos_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 05/07/2024
    '*************************************************
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Recargando Shield.ind...", 0, 0, 255, , , , , True)
    
    Call CargarAnimEscudos
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Helmet.ind recargado!", 0, 255, 0, , , , , True)
End Sub

Private Sub mnuRecargarArmas_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 05/07/2024
    '*************************************************
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Recargando Weapon.ind...", 0, 0, 255, , , , , True)
    
    Call CargarAnimArmas
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Weapon.ind recargado!", 0, 255, 0, , , , , True)
End Sub

Private Sub mnuRecargarParticulas_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 05/07/2024
    '*************************************************
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Recargando Particulas.ind...", 0, 0, 255, , , , , True)
    
    Call CargarParticulas
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Particulas.ind recargado!", 0, 255, 0, , , , , True)
End Sub

Private Sub mnuRecargarTriggers_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 05/07/2024
    '*************************************************
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Recargando Triggers.ini...", 0, 0, 255, , , , , True)
    
    Call CargarIndicesTriggers
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Triggers.ini recargado!", 0, 255, 0, , , , , True)
End Sub

Private Sub mnuRecargarIndices_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 05/07/2024
    '*************************************************
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Recargando Graficos.ind...", 0, 0, 255, , , , , True)
    
    Call LoadGrhData
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Graficos.ind recargado!", 0, 255, 0, , , , , True)
    
End Sub

Private Sub mnuRecargarNPCS_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 05/07/2024
    '*************************************************
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Recargando NPCs.dat...", 0, 0, 255, , , , , True)
    
    Call CargarIndicesNPC
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "NPCs.dat recargado!", 0, 255, 0, , , , , True)
    
End Sub

Private Sub mnuRecargarObj_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 05/07/2024
    '*************************************************
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Recargando Obj.dat...", 0, 0, 255, , , , , True)
    
    Call CargarIndicesOBJ
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Obj.dat recargado!", 0, 255, 0, , , , , True)
End Sub

Private Sub mnuRender_Click()
    On Error GoTo mnuRender_Click_Err
    
    Dim tmpPic As StdPicture
    Dim picNr As Long
    Dim sFileName As String
    Dim maxCx As Long, maxCy As Long
    Dim picWidth As Long, picHeight As Long
    
   
    frmRenderer.Show
   
    frmRenderer.PicGrande.ScaleMode = vbPixels
    frmRenderer.PicGrande.AutoRedraw = True
    frmRenderer.PicGrande.BorderStyle = 0&
    Dim x2 As Integer
    Dim y2 As Integer
    
    
     maxCy = YMaxMapSize * 4
     maxCx = XMaxMapSize * 4
     
     
    For y2 = 0 To ((YMaxMapSize - 1) \ 100)
       For x2 = 0 To ((XMaxMapSize - 1) \ 100)
       
         Call MapCapture(False, True, x2 * 100, y2 * 100)
         
      Next x2
    Next y2
    
    frmRenderer.PicGrande.AutoRedraw = False
    
    Unload frmRenderer
    
    Shell (App.path & "\UnirMinimapa.exe " & UserMap)
    DoEvents
    Sleep 1000
    DoEvents
    'CargarMinimap

    Exit Sub

mnuRender_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain. mnuRender_Click", Erl)
    Resume Next
End Sub

Private Sub mnuSalir_Click()
    Call CloseMapEditor
    
End Sub

Private Sub mnuScreenshot_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 31/03/2024
    '*************************************************
    
    If modGeneral.Client_Screenshot(MainViewPic.hDC, MainViewPic.ScaleWidth, MainViewPic.ScaleHeight) Then _
        Call ShowMessageScreen("Captura realizada")
        
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

Public Sub ObtenerNombreArchivo(ByVal Guardar As Boolean, _
                                Optional ByVal Zonas As Boolean)

    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    On Error GoTo ObtenerNombreArchivo_Err
    
    With Dialog
    
        If Not Zonas Then
            .Filter = "Mapas del nuevo formato (*.csm)|*.csm|Mapas clasicos de Argentum Online (*.map)|*.map"
        Else
            .Filter = "Mapa de zonas (*.zon)|*.zon"

        End If

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
    
    With ClientSetup
        mnuVerBloqueos.Checked = .VerBlockeados
        mnuVerTraslados.Checked = .VerTraslados
        mnuVerNPCs.Checked = .VerNpcs
        mnuVerObjetos.Checked = .VerObjetos
        mnuVerTriggers.Checked = .VerTriggers
        mnuVerParticulas.Checked = .VerParticulas
        mnuVerCapa1.Checked = .VerCapa1
        mnuVerCapa2.Checked = .VerCapa2
        mnuVerCapa3.Checked = .VerCapa3
        mnuVerCapa4.Checked = .VerCapa4
        mnuVerGrilla.Checked = .VerGrilla
    End With

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

Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2021
    '*************************************************

    Dim tX As Integer
    Dim tY As Integer
    
    If Not MapaCargado Then Exit Sub
    
    Call ConvertCPtoTP(x, y, tX, tY)
    
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
        Call modEdicion.ClickEdit(Button, tX, tY)
        
    End If
    
    Debug.Print Button

End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2021
    '*************************************************
    Dim tX As Integer
    Dim tY As Integer
    
    'Make sure map is loaded
    If Not MapaCargado Then Exit Sub
    HotKeysAllow = True

    Call ConvertCPtoTP(x, y, tX, tY)
    
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
        Call modEdicion.ClickEdit(Button, tX, tY)
        
    End If
    
End Sub

Private Sub mnuVerBloqueos_Click()
    ClientSetup.VerBlockeados = Not ClientSetup.VerBlockeados
    
    If ClientSetup.VerBlockeados Then
        Call ShowMessageScreen("Bloqueos visibles")
    Else
        Call ShowMessageScreen("Bloqueos ocultos")
    End If
    
End Sub

Private Sub mnuVerCapa1_Click()

    ClientSetup.VerCapa1 = Not ClientSetup.VerCapa1
    mnuVerCapa1.Checked = Not mnuVerCapa1.Checked
    
    If ClientSetup.VerCapa1 Then
        Call ShowMessageScreen("Capa 1 visible")
    Else
        Call ShowMessageScreen("Capa 1 oculto")
    End If
End Sub

Private Sub mnuVerCapa2_Click()

    ClientSetup.VerCapa2 = Not ClientSetup.VerCapa2
    mnuVerCapa2.Checked = Not mnuVerCapa2.Checked
    
    If ClientSetup.VerCapa2 Then
        Call ShowMessageScreen("Capa 2 visible")
    Else
        Call ShowMessageScreen("Capa 2 oculto")
    End If
End Sub

Private Sub mnuVerCapa3_Click()

    ClientSetup.VerCapa3 = Not ClientSetup.VerCapa3
    mnuVerCapa3.Checked = Not mnuVerCapa3.Checked
    
    If ClientSetup.VerCapa3 Then
        Call ShowMessageScreen("Capa 3 visible")
    Else
        Call ShowMessageScreen("Capa 3 oculto")
    End If
End Sub

Private Sub mnuVerCapa4_Click()

    ClientSetup.VerCapa4 = Not ClientSetup.VerCapa4
    mnuVerCapa4.Checked = Not mnuVerCapa4.Checked
    
    If ClientSetup.VerCapa1 Then
        Call ShowMessageScreen("Capa 4 visible")
    Else
        Call ShowMessageScreen("Capa 4 oculto")
    End If
End Sub

Private Sub mnuVerGrilla_Click()

    ClientSetup.VerGrilla = Not ClientSetup.VerGrilla
    mnuVerGrilla.Checked = ClientSetup.VerGrilla
    
    If ClientSetup.VerGrilla Then
        Call ShowMessageScreen("Grilla visible")
    Else
        Call ShowMessageScreen("Grilla oculto")
    End If
End Sub

Private Sub mnuVerNPCs_Click()

    ClientSetup.VerNpcs = Not ClientSetup.VerNpcs
    mnuVerNPCs.Checked = Not mnuVerNPCs.Checked
    
    If ClientSetup.VerNpcs Then
        Call ShowMessageScreen("NPC's visible")
    Else
        Call ShowMessageScreen("NPC's ocultos")
    End If
End Sub

Private Sub mnuVerObjetos_Click()

    ClientSetup.VerObjetos = Not ClientSetup.VerObjetos
    mnuVerObjetos.Checked = Not mnuVerObjetos.Checked
    
    If ClientSetup.VerObjetos Then
        Call ShowMessageScreen("Objetos visible")
    Else
        Call ShowMessageScreen("Objetos ocultos")
    End If
End Sub

Private Sub mnuVerParticulas_Click()

    ClientSetup.VerParticulas = Not ClientSetup.VerParticulas
    mnuVerParticulas.Checked = Not mnuVerParticulas.Checked
    
    If ClientSetup.VerParticulas Then
        Call ShowMessageScreen("Particulas visible")
    Else
        Call ShowMessageScreen("Particulas ocultas")
    End If
End Sub

Private Sub mnuVerTraslados_Click()

    ClientSetup.VerTraslados = Not ClientSetup.VerTraslados
    mnuVerTraslados.Checked = Not mnuVerTraslados.Checked
    
    If ClientSetup.VerTraslados Then
        Call ShowMessageScreen("Traslados visibles")
    Else
        Call ShowMessageScreen("Traslados ocultos")
    End If
    
End Sub

Private Sub mnuVerTriggers_Click()

    ClientSetup.VerTriggers = Not ClientSetup.VerTriggers
    mnuVerTriggers.Checked = Not mnuVerTriggers.Checked
    
    If ClientSetup.VerTriggers Then
        Call ShowMessageScreen("Triggers visibles")
    Else
        Call ShowMessageScreen("Triggers ocultos")
    End If
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
    Dim x As Integer
    Dim y As Integer
    
    For x = XMinMapSize To XMaxMapSize
    
        For y = YMinMapSize To YMaxMapSize
        
            If MapData(x, y).ZonaIndex = 0 Then
                MsgBox "Se ha encontrado una zona nula en la posicion X: " & x & " Y: " & y
                Exit Sub
            End If
        
        Next y
    
    Next x
    
End Sub

Private Sub mnuzonasinuso_Click()
    Dim x As Integer
    Dim y As Integer
    Dim i As Integer
    Dim found As Boolean
    Dim Count As Integer
    
    Count = 0
    
    If Not frmConsola.Visible Then _
        frmConsola.Show , frmMain
    
    For i = 1 To frmZonas.LstZona.ListCount
    
        For x = XMinMapSize To XMaxMapSize
    
            For y = YMinMapSize To YMaxMapSize
        
                If MapData(x, y).ZonaIndex = i Then
                    found = True
                    Exit For
                End If
        
            Next y
            
            If found = True Then Exit For
        
        Next x
        
        If found = False Then
            Call AddtoRichTextBox(frmConsola.StatTxt, "La zona " & i & " no se esta usando.", 255, 0, 0)
            Count = Count + 1
        End If
        
        found = False
    Next i
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Busqueda de zonas sin uso completada. Se encontraron " & Count & " zonas sin usar.", 255, 0, 0)

End Sub

Private Sub mnuZonasxCuadrantes_Click()
'***************************************************
'Author: Lorwik
'Fecha: 16/05/2024
'***************************************************

    On Error GoTo mnuZonasxCuadrantes_Click_Err

    Call modEdicion.Insertar_ZonasxCuadrantes
    
    Exit Sub

mnuZonasxCuadrantes_Click_Err:

    Call AddtoRichTextBox(frmConsola.StatTxt, "Erro al establecer zonas por cuadrantes establecidas.", 255, 0, 0)
    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuZonasxCuadrantes_Click", Erl)
    Resume Next

End Sub

Private Sub saveAllMinimap_Click()
    Call RenderizarCuadrantes
End Sub

Private Sub TimerMinuto_Timer()
'***************************************************
'Author: Lorwik
'Fecha: 24/04/2024
'***************************************************
    On Error GoTo TimerMinuto_Timer_Err

    Static SaveInterval As Byte

    '¿El guardado automatico esta activado? ¿El mapa se ha modificado?
    If ClientSetup.GuardadoAuto And MapInfo.Changed = 1 Then
    
        '¿El intervalo llego al configurado por el usuario?
        If ClientSetup.IntervaloGuardado = SaveInterval Then
            Call AddtoRichTextBox(frmConsola.StatTxt, "Comenzando guardado automatico.", 0, 255, 0)
            NoSobreescribir = True
            Call modMapas.GuardarMapa(Dialog.filename)
            
            Call ShowMessageScreen("Mapa guardado automaticamente.")
            Call AddtoRichTextBox(frmConsola.StatTxt, "Mapa guardado correctamente", 0, 255, 0)
        
        Else 'Si no llego sumamos 1
            SaveInterval = SaveInterval + 1
            
        End If
    
    Else 'Si no esta activado el autosave reseteamos el intervalo.
    
        SaveInterval = 0
    
    End If
    
    Exit Sub

TimerMinuto_Timer_Err:

    Call RegistrarError(Err.Number, Err.Description, "frmMain.TimerMinuto_Timer", Erl)
    Resume Next
End Sub
