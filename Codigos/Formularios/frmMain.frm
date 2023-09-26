VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00424242&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinterMapEditor"
   ClientHeight    =   10800
   ClientLeft      =   150
   ClientTop       =   495
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
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraEditar 
      BackColor       =   &H00535353&
      Caption         =   "Editar"
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
      Height          =   660
      Left            =   4890
      TabIndex        =   13
      Top             =   30
      Width           =   4065
      Begin WinterMapEditor.lvButtons_H LvBEdit 
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   200
         Width           =   375
         _ExtentX        =   661
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
         Mode            =   1
         Value           =   0   'False
         Image           =   "frmMain.frx":10CA
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBEdit 
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   15
         Top             =   200
         Width           =   375
         _ExtentX        =   661
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
         Mode            =   1
         Value           =   0   'False
         Image           =   "frmMain.frx":1D1C
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBEdit 
         Height          =   375
         Index           =   2
         Left            =   870
         TabIndex        =   16
         Top             =   200
         Width           =   375
         _ExtentX        =   661
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
         Mode            =   1
         Value           =   0   'False
         Image           =   "frmMain.frx":296E
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBEdit 
         Height          =   375
         Index           =   3
         Left            =   1260
         TabIndex        =   17
         Top             =   200
         Width           =   375
         _ExtentX        =   661
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
         Mode            =   1
         Value           =   0   'False
         Image           =   "frmMain.frx":35C0
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBEdit 
         Height          =   375
         Index           =   4
         Left            =   1650
         TabIndex        =   18
         Top             =   200
         Width           =   375
         _ExtentX        =   661
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
         Mode            =   1
         Value           =   0   'False
         Image           =   "frmMain.frx":4212
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBEdit 
         Height          =   375
         Index           =   5
         Left            =   2040
         TabIndex        =   19
         Top             =   200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "frmMain.frx":4E64
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBEdit 
         Height          =   375
         Index           =   8
         Left            =   3210
         TabIndex        =   22
         Top             =   200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "frmMain.frx":5AB6
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBEdit 
         Height          =   375
         Index           =   6
         Left            =   2430
         TabIndex        =   20
         Top             =   200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "frmMain.frx":6708
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBEdit 
         Height          =   375
         Index           =   7
         Left            =   2820
         TabIndex        =   21
         Top             =   200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "frmMain.frx":6D8A
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBEdit 
         Height          =   375
         Index           =   9
         Left            =   3600
         TabIndex        =   23
         Top             =   200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "frmMain.frx":722C
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame FraVer 
      BackColor       =   &H00535353&
      Caption         =   "Ver"
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
      Height          =   660
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4815
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   200
         Width           =   375
         _ExtentX        =   661
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
         Mode            =   1
         Value           =   0   'False
         Image           =   "frmMain.frx":7504
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   1
         Top             =   195
         Width           =   375
         _ExtentX        =   661
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
         Mode            =   1
         Value           =   0   'False
         Image           =   "frmMain.frx":8156
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   3
         Left            =   1260
         TabIndex        =   3
         Top             =   195
         Width           =   375
         _ExtentX        =   661
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
         Mode            =   1
         Value           =   0   'False
         Image           =   "frmMain.frx":8DA8
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   4
         Left            =   1650
         TabIndex        =   4
         Top             =   195
         Width           =   375
         _ExtentX        =   661
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
         Mode            =   1
         Value           =   0   'False
         Image           =   "frmMain.frx":99FA
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   5
         Left            =   2040
         TabIndex        =   5
         Top             =   195
         Width           =   375
         _ExtentX        =   661
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
         Mode            =   1
         Value           =   0   'False
         Image           =   "frmMain.frx":A64C
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   6
         Left            =   2430
         TabIndex        =   6
         Top             =   195
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "1"
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
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   7
         Left            =   2820
         TabIndex        =   7
         Top             =   195
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "2"
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
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   8
         Left            =   3210
         TabIndex        =   8
         Top             =   195
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "3"
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
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   9
         Left            =   3600
         TabIndex        =   9
         Top             =   195
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "4"
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
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   10
         Left            =   3990
         TabIndex        =   10
         Top             =   195
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "G"
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
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   11
         Left            =   4360
         TabIndex        =   11
         Top             =   195
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "Tr"
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
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   2
         Left            =   870
         TabIndex        =   2
         Top             =   195
         Width           =   375
         _ExtentX        =   661
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
         Mode            =   1
         Value           =   0   'False
         Image           =   "frmMain.frx":ACCE
         cBack           =   -2147483633
      End
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
      Height          =   10050
      Left            =   0
      ScaleHeight     =   668
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1278
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   750
      Width           =   19200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1276
      X2              =   1276
      Y1              =   6
      Y2              =   48
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   1274
      X2              =   1274
      Y1              =   6
      Y2              =   48
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
      Left            =   16710
      TabIndex        =   34
      Top             =   270
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
      Left            =   9060
      TabIndex        =   24
      Top             =   270
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
      Left            =   15945
      TabIndex        =   33
      Top             =   270
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
      Left            =   15180
      TabIndex        =   32
      Top             =   270
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
      Left            =   14415
      TabIndex        =   31
      Top             =   270
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
      Left            =   13650
      TabIndex        =   30
      Top             =   270
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
      Left            =   12885
      TabIndex        =   29
      Top             =   270
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
      Left            =   12120
      TabIndex        =   28
      Top             =   270
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
      Left            =   11355
      TabIndex        =   27
      Top             =   270
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
      Left            =   10590
      TabIndex        =   26
      Top             =   270
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
      Left            =   9825
      TabIndex        =   25
      Top             =   270
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
      Index           =   11
      Left            =   17475
      TabIndex        =   35
      Top             =   270
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
      Index           =   12
      Left            =   18210
      TabIndex        =   36
      Top             =   270
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
      Begin VB.Menu mnuOtrosAbrirMapa 
         Caption         =   "&Abrir otros Mapas"
         Begin VB.Menu abrirOtroMapa 
            Caption         =   "IAO 1.3"
            Index           =   0
         End
         Begin VB.Menu abrirOtroMapa 
            Caption         =   "IAO 1.4"
            Index           =   1
         End
         Begin VB.Menu abrirOtroMapa 
            Caption         =   "&Abrir Mapa AO [Int]"
            Index           =   2
         End
         Begin VB.Menu abrirOtroMapa 
            Caption         =   "Abrir Mapa AO [Long]"
            Index           =   3
         End
         Begin VB.Menu abrirOtroMapa 
            Caption         =   "Winter Old"
            Index           =   4
         End
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
   Begin VB.Menu mnuVentanas 
      Caption         =   "Ventanas"
      Begin VB.Menu VentMenu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVent 
         Caption         =   "Zonas"
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuVent 
         Caption         =   "Consola"
         Index           =   1
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuVent 
         Caption         =   "Mapa"
         Index           =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuVent 
         Caption         =   "Preview"
         Index           =   3
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuVent 
         Caption         =   "Configuración Avanzada de Superficies"
         Index           =   4
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuRellenar 
         Caption         =   "Rellenar en area"
      End
   End
   Begin VB.Menu mnuMinimapa 
      Caption         =   "Minimapa"
      Begin VB.Menu MinimapMenu 
         Caption         =   "-"
      End
      Begin VB.Menu Minimap_capa1 
         Caption         =   "Capa 1"
      End
      Begin VB.Menu Minimap_capa2 
         Caption         =   "Capa 2"
      End
      Begin VB.Menu Minimap_capa3 
         Caption         =   "Capa 3"
      End
      Begin VB.Menu Minimap_capa4 
         Caption         =   "Capa 4"
      End
      Begin VB.Menu Minimap_npcs 
         Caption         =   "NPC's"
      End
      Begin VB.Menu Minimap_objetos 
         Caption         =   "Objetos"
      End
      Begin VB.Menu Minimap_bloqueos 
         Caption         =   "Bloqueos"
      End
      Begin VB.Menu Minimap_particulas 
         Caption         =   "Particulas"
      End
      Begin VB.Menu Minimap_ndemapa 
         Caption         =   "Nº de mapa"
      End
      Begin VB.Menu Minimap_cuadrantes 
         Caption         =   "Cuadrantes"
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
      Begin VB.Menu mnuLineEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCompletarSuperficies 
         Caption         =   "Auto-Completar &Superficies"
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
      End
   End
   Begin VB.Menu mnuZonas 
      Caption         =   "Ver Zonas"
      Begin VB.Menu mnuLineZonas1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerZonas 
         Caption         =   "Zona Actual"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuVerZonas 
         Caption         =   "Todas las Zonas"
         Index           =   1
      End
      Begin VB.Menu mnuVerZonas 
         Caption         =   "Ocultar Zonas"
         Index           =   2
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
      Begin VB.Menu mnuOptimizarMasiv 
         Caption         =   "Optimización masiva"
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
   Begin VB.Menu mnuMapSize 
      Caption         =   "[Mapa ? x ?]"
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

Public Sub RefreshMapSize()
    
    mnuMapSize.Caption = "[Mapa " & XMaxMapSize & " x " & YMaxMapSize & "]"

End Sub

Private Sub LvBOpcion_Click(Index As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2020
'*************************************************

    Select Case Index
        Case 0
            VerBlockeados = Not VerBlockeados
            
        Case 1
            VerTranslados = Not VerTranslados
            
        Case 2
            VerNpcs = Not VerNpcs
            
        Case 3
            VerObjetos = Not VerObjetos
            
        Case 4
            VerTriggers = Not VerTriggers
            
        Case 5
            VerParticulas = Not VerParticulas
            
        Case 6
            VerCapa1 = Not VerCapa1
            
        Case 7
            VerCapa2 = Not VerCapa2
            
        Case 8
            VerCapa3 = Not VerCapa3
            
        Case 9
            VerCapa4 = Not VerCapa4
            
        Case 10
            VerGrilla = Not VerGrilla
            
        Case 11
            'AlphaTecho = Not AlphaTecho
            
    End Select
    
    Call guardarPerfil
    
End Sub

Private Sub LvBEdit_Click(Index As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2020
'*************************************************

    Select Case Index
    
        Case 0 ' Superficies
            If LvBEdit(0).value Then
                frmSuperficies.Show , frmMain
            
            Else
                frmSuperficies.Visible = False
                
            End If
            
        Case 1 ' Traslados
            If LvBEdit(1).value Then
                frmTraslados.Show , frmMain
            
            Else
                frmTraslados.Visible = False
                
            End If
        
        Case 2 ' Bloqueos
            If LvBEdit(2).value Then
                frmBloqueos.Show , frmMain
            
            Else
                frmBloqueos.Visible = False
                
            End If
        
        Case 3 ' NPCs
            If LvBEdit(3).value Then
                frmNPCs.Show , frmMain
            
            Else
                frmNPCs.Visible = False
                
            End If
        
        Case 4 ' Objetos
            If LvBEdit(4).value Then
                frmOBJs.Show , frmMain
            
            Else
                frmOBJs.Visible = False
                
            End If
        
        Case 5 ' Triggers
            If LvBEdit(5).value Then
                frmTriggers.Show , frmMain
            
            Else
                frmTriggers.Visible = False
                
            End If
        
        Case 6 ' Particulas
            If LvBEdit(6).value Then
                frmParticulas.Show , frmMain
            
            Else
                frmParticulas.Visible = False
                
            End If
        
        Case 7 ' Luces
            If LvBEdit(7).value Then
                frmLuces.Show , frmMain
            
            Else
                frmLuces.Visible = False
                
            End If
        
        Case 8 ' Copiar
        
        Case 9 ' Informacion
            If LvBEdit(9).value Then
                frmMapInfo.Show , frmMain
            
            Else
                frmMapInfo.Visible = False
                
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
    
    Select Case frmMain.Dialog.FilterIndex
    
        Case 1
            Formato = ".csm"
            
        Case 2
            Formato = ".map"
            
    End Select
    
    
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then _
            Call modMapas.GuardarMapa(Dialog.filename)

    End If
        
    If (Index + NumMap_Save - 4) <> NumMap_Save Then
        Dialog.CancelError = True

        On Error GoTo errhandler

        Dialog.filename = PATH_Save & NameMap_Save & (Index + NumMap_Save - 4) & Formato
        
        Call modMapas.NuevoMapa
        
        DoEvents
        
        Call abrirCargarMapa(frmMain.Dialog.filename, TipoMapaActual)
        
        EngineRun = True
        
    End If
    
        Exit Sub
    
errhandler:
    Call MsgBox(Err.Description)
    
End Sub

Private Sub Minimap_cuadrantes_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo Minimap_cuadrante_Click_Err
    
    Minimap_cuadrantes.Checked = (Minimap_cuadrantes.Checked = False)
    MMiniMap_cuadrantes = Not MMiniMap_cuadrantes
    Call DibujarMinimapa
    
    Exit Sub

Minimap_cuadrante_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Minimap_cuadrante_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAbrirMapa_Click()
'*************************************************
'Author: Lorwik
'Last modified: 25/04/2020
'*************************************************

    Select Case ClientSetup.MeMode
    
        Case eMeMode.WinterAO
            Call AbrirMapa(eTipoMapa.tWinter)
    
        Case eMeMode.ImperiumClasico
            Call AbrirMapa(eTipoMapa.tIAOClasico)
            
        Case eMeMode.WinterUltimate
            Call AbrirMapa(eTipoMapa.tIAOClasico)
            
        Case eMeMode.ArgentumUnited
            Call AbrirMapa(eTipoMapa.tAOUnited)
        
    End Select
    
End Sub

Private Sub abrirOtroMapa_Click(Index As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 21/09/2021
'*************************************************
    Select Case Index
    
        Case 0 'IAO 1.3
            Call AbrirMapa(eTipoMapa.tIAOold)
            
        Case 1 'IAO 1.4
            Call AbrirMapa(eTipoMapa.tIAOnew)
            
        Case 2 'AO Int
            Call AbrirMapa(eTipoMapa.tInt)
            
        Case 3 'AO Long
            Call AbrirMapa(eTipoMapa.tlong)
            
        Case 4 'WAO Old
            Call AbrirMapa(tWinter_Old)
            
    End Select
End Sub

Private Sub mnuAcerca_Click()
'*************************************************
'Author: Lorwik
'Last modified: 01/05/2021
'*************************************************

    frmAcercade.Show , frmMain
End Sub

Private Sub mnuAutoCompletarSuperficies_Click()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************
    mnuAutoCompletarSuperficies.Checked = (mnuAutoCompletarSuperficies.Checked = False)
    
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

Private Sub mnuMapSize_Click()
    Call frmMapaSize.Show
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
    
    If WalkMode Then _
        Call modGeneral.ToggleWalkMode
    
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
    
    frmOptimizar.Height = 4214
    frmOptimizar.FraOptimizarTodos.Visible = False
End Sub

Private Sub mnuOptimizarMasiv_Click()
'*************************************************
'Author: Lorwik
'Last modified: 01/05/2021
'*************************************************

    frmOptimizar.Show , frmMain
    
    frmOptimizar.Height = 6075
    frmOptimizar.FraOptimizarTodos.Visible = True
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
    
    Call abrirCargarMapa(frmMain.Dialog.filename, TipoMapaActual)
    
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
Exit Sub

mnuReAbrirMapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuReAbrirMapa_Click", Erl)
    Resume Next
End Sub

Private Sub mnuRellenar_Click()
    frmRellenar.Show , frmMain
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

Private Sub mnuVent_Click(Index As Integer)

    Select Case Index
    
        Case 0
            frmZonas.Show , frmMain
            
        Case 1
            frmConsola.Show , frmMain
            
        Case 2
            frmMapa.Show , frmMain
            
        Case 3
            frmPreview.Show , frmMain
            
        Case 4
            frmConfigSup.Show , frmMain
    End Select

End Sub

Private Sub MiniMap_Bloqueos_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_Bloqueos_Click_Err
    
    Minimap_bloqueos.Checked = (Minimap_bloqueos.Checked = False)
    MMiniMap_Bloqueos = Not MMiniMap_Bloqueos

    Exit Sub

MiniMap_Bloqueos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_Bloqueos_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_capa1_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_capa1_Click_Err
    
    Minimap_capa1.Checked = (Minimap_capa1.Checked = False)
    MMiniMap_capa1 = Not MMiniMap_capa1
    Call DibujarMinimapa
    
    Exit Sub

MiniMap_capa1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_capa1_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_capa2_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_capa2_Click_Err
    
    Minimap_capa2.Checked = (Minimap_capa2.Checked = False)
    MMiniMap_capa2 = Not MMiniMap_capa2
    Call DibujarMinimapa

    Exit Sub

MiniMap_capa2_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_capa2_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_capa3_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_capa3_Click_Err
    
    Minimap_capa3.Checked = (Minimap_capa3.Checked = False)
    MMiniMap_capa3 = Not MMiniMap_capa3
    Call DibujarMinimapa
    
    Exit Sub

MiniMap_capa3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_capa3_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_capa4_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_capa4_Click_Err
    
    Minimap_capa4.Checked = (Minimap_capa4.Checked = False)
    MMiniMap_capa4 = Not MMiniMap_capa4
    Call DibujarMinimapa

    Exit Sub

MiniMap_capa4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_capa4_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_ndemapa_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_ndemapa_Click_Err
    
    Minimap_ndemapa.Checked = (Minimap_ndemapa.Checked = False)
    MMiniMap_Nombre = Not MMiniMap_Nombre
    Call DibujarMinimapa
    
    Exit Sub

MiniMap_ndemapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_ndemapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_Npcs_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_Npcs_Click_Err
    
    Minimap_npcs.Checked = (Minimap_npcs.Checked = False)
    MMiniMap_Npcs = Not MMiniMap_Npcs
    Call DibujarMinimapa
    
    Exit Sub

MiniMap_Npcs_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_Npcs_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_objetos_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_objetos_Click_Err
    
    Minimap_objetos.Checked = (Minimap_objetos.Checked = False)
    MMiniMap_objetos = Not MMiniMap_objetos
    Call DibujarMinimapa

    
    Exit Sub

MiniMap_objetos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_objetos_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_particulas_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_particulas_Click_Err
    
    Minimap_particulas.Checked = (Minimap_particulas.Checked = False)
    MMiniMap_particulas = Not MMiniMap_particulas
    Call DibujarMinimapa

    
    Exit Sub

MiniMap_particulas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_particulas_Click", Erl)
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
                                  Y As Single)
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    Call Form_MouseDown(Button, Shift, X, Y)
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
    
    LvBOpcion(0).value = VerBlockeados
    LvBOpcion(1).value = VerTranslados
    LvBOpcion(2).value = VerNpcs
    LvBOpcion(3).value = VerObjetos
    LvBOpcion(4).value = VerTriggers
    LvBOpcion(5).value = VerParticulas
    LvBOpcion(6).value = VerCapa1
    LvBOpcion(7).value = VerCapa2
    LvBOpcion(8).value = VerCapa3
    LvBOpcion(9).value = VerCapa4
    LvBOpcion(10).value = VerGrilla
    
    'Opciones que no van a esta disponibles si iniciamos en modo IAOC
    If ClientSetup.MeMode = eMeMode.ImperiumClasico Or _
        ClientSetup.MeMode = eMeMode.WinterUltimate Then
        mnuZonas.Enabled = False
        mnuZonas.Visible = False
        
        mnuVent(0).Visible = False
        mnuVent(0).Enabled = False
        
        abrirOtroMapa(4).Visible = False 'Desactivamos los mapas Winter Old
        
    Else
    
        mnuOptimizarMasiv.Visible = False
    
    End If
    
    #If VisorMode = 1 Then
    
        mnuGuardarMapa.Visible = False
        mnuGuardarMapaComo.Visible = False
        
        mnuVent(3).Visible = False
        mnuVent(4).Visible = False
        mnuRellenar.Visible = False
        
        For i = 0 To 6
            LvBEdit(i).Visible = False
        Next i
        
        LvBEdit(8).Visible = False
        
        mnuEdicion.Visible = False
        mnuFunciones.Visible = False
        mnuFormatos.Visible = False
        
        minimapaSave.Visible = False
        saveAllMinimap.Visible = False
        
        mnuOtrosAbrirMapa.Visible = False
    
    #End If
    
    Call RefreshMapSize

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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: Lorwik
'Last modified: 26/04/2021
'*************************************************

    Dim tX As Integer
    Dim tY As Integer
    
    If Not MapaCargado Then Exit Sub
    
    Call ConvertCPtoTP(X, Y, tX, tY)
    
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: Lorwik
'Last modified: 26/04/2021
'*************************************************
    Dim tX As Integer
    Dim tY As Integer
    
    'Make sure map is loaded
    If Not MapaCargado Then Exit Sub
    HotKeysAllow = True

    Call ConvertCPtoTP(X, Y, tX, tY)
    
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

Private Sub mnuVerZonas_Click(Index As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 01/04/2021
'*************************************************

    mnuVerZonas(Index).Checked = True

    Select Case Index
    
        Case 0
            mnuVerZonas(1).Checked = False
            mnuVerZonas(2).Checked = False
        
        Case 1
            mnuVerZonas(0).Checked = False
            mnuVerZonas(2).Checked = False
        
        Case 2
            mnuVerZonas(0).Checked = False
            mnuVerZonas(1).Checked = False
    
    End Select

End Sub

Private Sub mnuzonanula_Click()
    Dim X As Integer
    Dim Y As Integer
    
    For X = XMinMapSize To XMaxMapSize
    
        For Y = YMinMapSize To YMaxMapSize
        
            If MapData(X, Y).ZonaIndex = 0 Then
                MsgBox "Se ha encontrado una zona nula en la posicion X: " & X & " Y: " & Y
                Exit Sub
            End If
        
        Next Y
    
    Next X
    
End Sub

Private Sub saveAllMinimap_Click()
    Call RenderizarCuadrantes
End Sub
