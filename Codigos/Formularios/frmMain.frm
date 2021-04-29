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
      Left            =   5040
      TabIndex        =   23
      Top             =   30
      Width           =   4065
      Begin WinterMapEditor.lvButtons_H LvBEdit 
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   0
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   38
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
      Left            =   120
      TabIndex        =   22
      Top             =   30
      Width           =   4815
      Begin WinterMapEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   36
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
      Left            =   18660
      Top             =   10260
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
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   750
      Width           =   19200
   End
   Begin VB.Label lblByLorwik 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "by Lorwik"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   18390
      TabIndex        =   37
      Top             =   0
      Width           =   675
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
      Left            =   16860
      TabIndex        =   11
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
      Left            =   9210
      TabIndex        =   21
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
      Left            =   16095
      TabIndex        =   12
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
      Left            =   15330
      TabIndex        =   13
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
      Left            =   14565
      TabIndex        =   14
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
      Left            =   13800
      TabIndex        =   15
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
      Left            =   13035
      TabIndex        =   16
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
      Left            =   12240
      TabIndex        =   17
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
      Left            =   11505
      TabIndex        =   18
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
      Left            =   10740
      TabIndex        =   19
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
      Left            =   9975
      TabIndex        =   20
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
      Left            =   17625
      TabIndex        =   10
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
      Left            =   18390
      TabIndex        =   9
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
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa [Int]"
         Index           =   1
      End
      Begin VB.Menu mnuArchivoLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarMapa 
         Caption         =   "&Guardar Mapa"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGuardarMapaComo 
         Caption         =   "Guardar Mapa &como..."
      End
      Begin VB.Menu mnuArchivoLine4 
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
         Shortcut        =   ^S
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
      Begin VB.Menu mnuLineEdicion4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCompletarSuperficies 
         Caption         =   "Auto-Completar &Superficies"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
                frmObjs.Show , frmMain
            
            Else
                frmObjs.Visible = False
                
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

        Dialog.filename = PATH_Save & NameMap_Save & (Index + NumMap_Save - 7) & Formato
        
        Call modMapas.NuevoMapa
        
        DoEvents
        Select Case frmMain.Dialog.FilterIndex
        
            Case 1
                If ClientSetup.TipoMapaCargado = eTipoMapa.tWinter Then
                    Call modMapas.Cargar_CSM(Dialog.filename)
                End If
                
            Case 2
                If ClientSetup.TipoMapaCargado = eTipoMapa.tInt Then
                    'Call modMapIO.MapaV2_Cargar(Dialog.filename, True)
                    
                Else
                    'Call modMapIO.MapaV2_Cargar(Dialog.filename)
                    
                End If
            
        End Select
        
        EngineRun = True
        
    End If
    
        Exit Sub
    
errhandler:
        Call MsgBox(Err.Description)
End Sub

Private Sub mnuAbrirMapa_Click(Index As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 25/04/2020
'*************************************************
    Select Case Index
    
        Case 0
            Call AbrirMapa(False)
            
        Case 1
            Call AbrirMapa(True)
            
    End Select
    
End Sub

Private Sub mnuAutoCompletarSuperficies_Click()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************
    mnuAutoCompletarSuperficies.Checked = (mnuAutoCompletarSuperficies.Checked = False)
    
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
    
'    If WalkMode = True Then
'        Call modGeneral.ToggleWalkMode
'    End If
    
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

Private Sub mnuSalir_Click()
    Call CloseMapEditor
    
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
On Error Resume Next
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

    'Temporal
    ClientSetup.MapTam = 1
    Call setMapSize
    
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
    
    If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionIX = tX '+ UserPos.X
        SeleccionIY = tY '+ UserPos.Y
        'DX1.Text = tX
        'DY1.Text = tY
        
    Else
        Call modEdit.ClickEdit(Button, tX, tY)
        
    End If

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
    
     If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionFX = tX '+ TileX
        SeleccionFY = tY '+ TileY
        'DX2.Text = tX
        'DY2.Text = tY
        
    Else
        Call modEdit.ClickEdit(Button, tX, tY)
        
    End If

End Sub
