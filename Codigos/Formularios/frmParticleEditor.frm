VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParticleEditor 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor de Particulas"
   ClientHeight    =   8985
   ClientLeft      =   25995
   ClientTop       =   4800
   ClientWidth     =   8070
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
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
   ScaleHeight     =   8985
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      Caption         =   "Lista de particulas"
      ForeColor       =   &H0000FF00&
      Height          =   5535
      Left            =   240
      TabIndex        =   84
      Top             =   120
      Width           =   2415
      Begin VB.ListBox ListParticulas 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   5130
         ItemData        =   "frmParticleEditor.frx":0000
         Left            =   85
         List            =   "frmParticleEditor.frx":0002
         TabIndex        =   85
         Top             =   240
         Width           =   2250
      End
   End
   Begin VB.Frame frameGrhs 
      BackColor       =   &H00404040&
      Caption         =   "Grh's de la Particula"
      ForeColor       =   &H0000FF00&
      Height          =   5505
      Left            =   5190
      TabIndex        =   82
      Top             =   150
      Width           =   2760
      Begin WinterMapEditor.lvButtons_H cmdLimpiar 
         Height          =   435
         Left            =   1830
         TabIndex        =   86
         Top             =   4920
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   767
         Caption         =   "Limpiar"
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
         cFore           =   16777215
         cFHover         =   16777215
         cBhover         =   0
         cGradient       =   0
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632064
      End
      Begin VB.ListBox lstSelGrhs 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   4545
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   2490
      End
      Begin WinterMapEditor.lvButtons_H cmdEliminar 
         Height          =   435
         Left            =   990
         TabIndex        =   87
         Top             =   4920
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   767
         Caption         =   "Eliminar"
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
         cFore           =   16777215
         cFHover         =   16777215
         cBhover         =   0
         cGradient       =   0
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   192
      End
      Begin WinterMapEditor.lvButtons_H cmdAñadir 
         Height          =   435
         Left            =   120
         TabIndex        =   88
         Top             =   4920
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   767
         Caption         =   "Añadir"
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
         cFore           =   16777215
         cFHover         =   16777215
         cBhover         =   0
         cGradient       =   0
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   49152
      End
   End
   Begin VB.Frame frmfade 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2235
      Left            =   240
      TabIndex        =   75
      Top             =   6600
      Width           =   7680
      Begin VB.TextBox txtfin 
         Height          =   285
         Left            =   1320
         TabIndex        =   77
         Text            =   "0"
         Top             =   90
         Width           =   630
      End
      Begin VB.TextBox txtfout 
         Height          =   300
         Left            =   1320
         TabIndex        =   76
         Text            =   "0"
         Top             =   405
         Width           =   645
      End
      Begin WinterMapEditor.lvButtons_H cmdReferencia 
         Height          =   435
         Left            =   5940
         TabIndex        =   94
         Top             =   1770
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
         Caption         =   "Referencia"
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
         cFore           =   16777215
         cFHover         =   16777215
         cBhover         =   0
         cGradient       =   0
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   8388736
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fade in time"
         Height          =   180
         Left            =   60
         TabIndex        =   80
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fade out time"
         Height          =   300
         Left            =   60
         TabIndex        =   79
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note: The time a particle remains alive is set in the Duration Tab"
         Height          =   585
         Left            =   90
         TabIndex        =   78
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Particle Speed"
      Height          =   855
      Left            =   555
      TabIndex        =   72
      Top             =   6690
      Width           =   1935
      Begin VB.TextBox speed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   73
         Text            =   "0.5"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Render Delay:"
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Frame frameColorSettings 
      BorderStyle     =   0  'None
      Caption         =   "Color Tint Settings"
      Height          =   2175
      Left            =   495
      TabIndex        =   60
      Top             =   6555
      Width           =   3975
      Begin VB.TextBox txtB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   68
         Text            =   "0"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtG 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   67
         Text            =   "0"
         Top             =   1500
         Width           =   375
      End
      Begin VB.TextBox txtR 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   66
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   1440
         ScaleHeight     =   795
         ScaleWidth      =   2355
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   240
         Width           =   2415
      End
      Begin VB.ListBox lstColorSets 
         Height          =   840
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   1215
      End
      Begin VB.HScrollBar BScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   63
         Top             =   1200
         Width           =   3015
      End
      Begin VB.HScrollBar GScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   62
         Top             =   1500
         Width           =   3015
      End
      Begin VB.HScrollBar RScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   61
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B:"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   1800
         Width           =   150
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G:"
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   1500
         Width           =   165
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R:"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   1200
         Width           =   165
      End
   End
   Begin VB.Frame frmSettings 
      BorderStyle     =   0  'None
      Height          =   2190
      Left            =   1080
      TabIndex        =   27
      Top             =   6600
      Width           =   6600
      Begin VB.TextBox txtrx 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   44
         Text            =   "0"
         Top             =   1395
         Width           =   495
      End
      Begin VB.TextBox txtPCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   43
         Text            =   "20"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtX1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   42
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtX2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   41
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtY1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   40
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtY2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   39
         Text            =   "0"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtAngle 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   38
         Text            =   "0"
         Top             =   1605
         Width           =   495
      End
      Begin VB.TextBox vecx1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   37
         Text            =   "-10"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox vecx2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "10"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox vecy1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   35
         Text            =   "-50"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox vecy2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   34
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox life1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   33
         Text            =   "10"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox life2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   32
         Text            =   "50"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox fric 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   31
         Text            =   "5"
         Top             =   840
         Width           =   495
      End
      Begin VB.CheckBox chkAlphaBlend 
         Caption         =   "Alpha Blend"
         Height          =   255
         Left            =   3930
         TabIndex        =   30
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkresize 
         Caption         =   "Resize"
         Height          =   195
         Left            =   1920
         TabIndex        =   29
         Top             =   1920
         Width           =   1245
      End
      Begin VB.TextBox txtry 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   28
         Text            =   "0"
         Top             =   1635
         Width           =   495
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angle:"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   1650
         Width           =   450
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector X1:"
         Height          =   195
         Left            =   1950
         TabIndex        =   58
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector X2:"
         Height          =   195
         Left            =   1950
         TabIndex        =   57
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector Y1:"
         Height          =   195
         Left            =   1950
         TabIndex        =   56
         Top             =   765
         Width           =   750
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector Y2"
         Height          =   195
         Left            =   1950
         TabIndex        =   55
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life Range (L):"
         Height          =   195
         Left            =   3915
         TabIndex        =   54
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life Range (H):"
         Height          =   195
         Left            =   3915
         TabIndex        =   53
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Friction:"
         Height          =   195
         Left            =   3915
         TabIndex        =   52
         Top             =   885
         Width           =   555
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y2:"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   1245
         Width           =   240
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y1:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1005
         Width           =   240
      End
      Begin VB.Label lblPCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# of Particles:"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X1:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   525
         Width           =   240
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X2:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resize Y:"
         Height          =   195
         Left            =   1950
         TabIndex        =   46
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resize X:"
         Height          =   195
         Left            =   1950
         TabIndex        =   45
         Top             =   1440
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Particle Duration"
      Height          =   855
      Left            =   450
      TabIndex        =   23
      Top             =   6645
      Width           =   1935
      Begin VB.TextBox life 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Text            =   "10"
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkNeverDies 
         Caption         =   "Never Dies"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   525
         Width           =   300
      End
   End
   Begin VB.Frame frameSpinSettings 
      BorderStyle     =   0  'None
      Caption         =   "Spin Settings"
      Height          =   1095
      Left            =   465
      TabIndex        =   17
      Top             =   6615
      Width           =   1935
      Begin VB.CheckBox chkSpin 
         Caption         =   "Spin"
         Height          =   255
         Left            =   105
         TabIndex        =   20
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox spin_speedH 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   19
         Text            =   "1"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox spin_speedL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   18
         Text            =   "1"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Speed (H):"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   765
         Width           =   1125
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Speed (L):"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   525
         Width           =   1095
      End
   End
   Begin VB.Frame frameMovement 
      BorderStyle     =   0  'None
      Caption         =   "Movement Settings"
      Height          =   1935
      Left            =   435
      TabIndex        =   6
      Top             =   6615
      Width           =   1935
      Begin VB.CheckBox chkXMove 
         Caption         =   "X Movement"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkYMove 
         Caption         =   "Y Movement"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox move_y2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "0"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox move_y1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox move_x2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "0"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox move_x1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement X1:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   525
         Width           =   1035
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement X2:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   765
         Width           =   1035
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Y1:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1365
         Width           =   1035
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Y2:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1605
         Width           =   1035
      End
   End
   Begin VB.Frame frameGravity 
      BorderStyle     =   0  'None
      Caption         =   "Gravity Settings"
      Height          =   1095
      Left            =   450
      TabIndex        =   0
      Top             =   6630
      Width           =   1935
      Begin VB.CheckBox chkGravity 
         Caption         =   "Gravity Influence"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox txtBounceStrength 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "1"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtGravStrength 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "5"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bounce Strength:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   705
         Width           =   1245
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gravity Strength:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   465
         Width           =   1185
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2670
      Left            =   120
      TabIndex        =   81
      Top             =   6240
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   4710
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Particle Settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravity"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Movement "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Spin "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Speed"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Duration "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Color "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fade"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   3120
      Top             =   300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WinterMapEditor.lvButtons_H cmdSalir 
      Height          =   435
      Left            =   7080
      TabIndex        =   89
      Top             =   5730
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   767
      Caption         =   "Salir"
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
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   0
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   192
   End
   Begin WinterMapEditor.lvButtons_H LvBGuardarParticula 
      Height          =   435
      Left            =   5310
      TabIndex        =   90
      Top             =   5730
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      Caption         =   "Guardar Particula"
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
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   0
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   49152
   End
   Begin WinterMapEditor.lvButtons_H cmdDesaparecer 
      Height          =   435
      Left            =   3540
      TabIndex        =   91
      Top             =   5730
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      Caption         =   "Desaparecer"
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
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   0
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16744576
   End
   Begin WinterMapEditor.lvButtons_H cmdNuevaParticula 
      Height          =   435
      Left            =   1800
      TabIndex        =   92
      Top             =   5730
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      Caption         =   "Nueva particula"
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
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   0
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   49152
   End
   Begin WinterMapEditor.lvButtons_H LvBVistaPrevia 
      Height          =   435
      Left            =   90
      TabIndex        =   93
      Top             =   5730
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      Caption         =   "Vista Previa"
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
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   0
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   65535
   End
End
Attribute VB_Name = "frmParticleEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lstColorSets.AddItem "Bottom Left"
    lstColorSets.AddItem "Top Left"
    lstColorSets.AddItem "Bottom Right"
    lstColorSets.AddItem "Top Right"
    frmSettings.Visible = True
    frmfade.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False

End Sub

Private Sub cmdDesaparecer_Click()
    Particle_Group_Remove_All

End Sub

Private Sub cmdGuardarParticula_Click()
    Call GuardarParticulas

End Sub

Private Sub cmdNuevaParticula_Click()
    Call ModParticulas.NuevaParticula

End Sub

Private Sub cmdVistaPrevia_Click()

    If ListParticulas.ListIndex < 0 Then Exit Sub
    Call CargarParticulasLista

End Sub

Private Sub ListParticulas_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim loopc    As Long

    Dim DataTemp As Boolean

    DataTemp = DataChanged
    
    With StreamData(ListParticulas.ListIndex + 1)
    
        'Set the values
        txtPCount.Text = .NumOfParticles
        txtX1.Text = .x1
        txtY1.Text = .y1
        txtX2.Text = .x2
        txtY2.Text = .y2
        txtAngle.Text = .angle
        vecx1.Text = .vecx1
        vecx2.Text = .vecx2
        vecy1.Text = .vecy1
        vecy2.Text = .vecy2
        life1.Text = .life1
        life2.Text = .life2
        fric.Text = .friction
        chkSpin.value = .spin
        spin_speedL.Text = .spin_speedL
        spin_speedH.Text = .spin_speedH
        txtGravStrength.Text = .grav_strength
        txtBounceStrength.Text = .bounce_strength
        
        chkAlphaBlend.value = .alphaBlend
        chkGravity.value = .gravity
        
        chkXMove.value = .XMove
        chkYMove.value = .YMove
        move_x1.Text = .move_x1
        move_x2.Text = .move_x2
        move_y1.Text = .move_y1
        move_y2.Text = .move_y2
        
        lstSelGrhs.Clear
        
        For loopc = 1 To .NumGrhs
            lstSelGrhs.AddItem .grh_list(loopc)
        Next loopc
    
    End With

End Sub

Private Sub ListParticulas_Click()
    Call CargarParticulasLista

    CurrentGrh.GrhIndex = 0
    CurrentGrh.Started = 0
    Renderizando = eRender.eParticulas
End Sub

Private Sub cmdDelete_Click()

    Dim loopc As Long
    
    If lstSelGrhs.ListIndex >= 0 Then lstSelGrhs.RemoveItem lstSelGrhs.ListIndex
    
    StreamData(ListParticulas.ListIndex + 1).NumGrhs = lstSelGrhs.ListCount
    
    If StreamData(ListParticulas.ListIndex + 1).NumGrhs = 0 Then
        Erase StreamData(ListParticulas.ListIndex + 1).grh_list
    Else
        ReDim StreamData(ListParticulas.ListIndex + 1).grh_list(1 To lstSelGrhs.ListCount) As Long

    End If
    
    For loopc = 1 To StreamData(ListParticulas.ListIndex + 1).NumGrhs
        StreamData(ListParticulas.ListIndex + 1).grh_list(loopc) = lstSelGrhs.List(loopc - 1)
    Next loopc

End Sub

Private Sub LoadStreamFile(StreamFile As String)

    Dim loopc As Long
    
    '****************************
    'load stream types
    '****************************
    TotalStreams = Val(GetVar(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'clear combo box
    ListParticulas.Clear
    
    Dim i          As Long

    Dim GrhListing As String

    'fill StreamData array with info from particle.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).name = GetVar(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = GetVar(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = GetVar(StreamFile, Val(loopc), "X1")
        StreamData(loopc).y1 = GetVar(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).x2 = GetVar(StreamFile, Val(loopc), "X2")
        StreamData(loopc).y2 = GetVar(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).angle = GetVar(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = GetVar(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = GetVar(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = GetVar(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = GetVar(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = GetVar(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = GetVar(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = GetVar(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = GetVar(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = GetVar(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = GetVar(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).alphaBlend = GetVar(StreamFile, Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = GetVar(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = GetVar(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = GetVar(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = GetVar(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = GetVar(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = GetVar(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = GetVar(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = GetVar(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = GetVar(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).life_counter = GetVar(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(GetVar(StreamFile, Val(loopc), "Speed"))
        StreamData(loopc).NumGrhs = GetVar(StreamFile, Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs) As Long
        GrhListing = GetVar(StreamFile, Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = GetVar(str(i), GrhListing, 44)
        Next i
        
        Dim TempSet  As String

        Dim ColorSet As Long
        
        For ColorSet = 1 To 4
            TempSet = GetVar(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).R = GetVar(1, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).G = GetVar(2, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).B = GetVar(3, TempSet, 44)
        Next ColorSet
        
        'fill stream type combo box
        ListParticulas.AddItem loopc & " - " & StreamData(loopc).name
    Next loopc
    
    'set list box index to 1st item
    ListParticulas.ListIndex = 0

End Sub

Private Sub LvBReferenciaPersonaje_Click()
    ReferenciaPJ = Not ReferenciaPJ

End Sub

Private Sub LvBSalir_Click()
    frmParticleEditor.Visible = False

End Sub

Private Sub TabStrip1_Click()

    Select Case TabStrip1.SelectedItem.Index

        Case 1:
            frmSettings.Visible = True
            frameColorSettings.Visible = False
            Frame2.Visible = False
            Frame1.Visible = False
            frameSpinSettings.Visible = False
            frameMovement.Visible = False
            frameGravity.Visible = False
            frmfade.Visible = False

        Case 2:
            frmSettings.Visible = False
            frameColorSettings.Visible = False
            Frame2.Visible = False
            Frame1.Visible = False
            frameSpinSettings.Visible = False
            frameMovement.Visible = False
            frameGravity.Visible = True
            frmfade.Visible = False

        Case 3:
            frmSettings.Visible = False
            frameColorSettings.Visible = False
            Frame2.Visible = False
            Frame1.Visible = False
            frameSpinSettings.Visible = False
            frameMovement.Visible = True
            frameGravity.Visible = False
            frmfade.Visible = False

        Case 4:
            frmSettings.Visible = False
            frameColorSettings.Visible = False
            Frame2.Visible = False
            Frame1.Visible = False
            frameSpinSettings.Visible = True
            frameMovement.Visible = False
            frameGravity.Visible = False
            frmfade.Visible = False

        Case 5:
            frmSettings.Visible = False
            frameColorSettings.Visible = False
            Frame2.Visible = True
            Frame1.Visible = False
            frameSpinSettings.Visible = False
            frameMovement.Visible = False
            frameGravity.Visible = False
            frmfade.Visible = False

        Case 6:
            frmSettings.Visible = False
            frameColorSettings.Visible = False
            Frame2.Visible = False
            Frame1.Visible = True
            frameSpinSettings.Visible = False
            frameMovement.Visible = False
            frameGravity.Visible = False
            frmfade.Visible = False

        Case 7:
            frmSettings.Visible = False
            frameColorSettings.Visible = True
            Frame2.Visible = False
            Frame1.Visible = False
            frameSpinSettings.Visible = False
            frameMovement.Visible = False
            frameGravity.Visible = False
            frmfade.Visible = False

        Case 8:
            frmSettings.Visible = False
            frameColorSettings.Visible = False
            Frame2.Visible = False
            Frame1.Visible = False
            frameSpinSettings.Visible = False
            frameMovement.Visible = False
            frameGravity.Visible = False
            frmfade.Visible = True

    End Select

End Sub

Private Sub vecx1_GotFocus()

    vecx1.SelStart = 0
    vecx1.SelLength = Len(vecx1.Text)

End Sub

Private Sub vecx1_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).vecx1 = vecx1.Text

End Sub

Private Sub vecx2_GotFocus()

    vecx2.SelStart = 0
    vecx2.SelLength = Len(vecx2.Text)

End Sub

Private Sub vecx2_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).vecx2 = vecx2.Text

End Sub

Private Sub vecy1_GotFocus()

    vecy1.SelStart = 0
    vecy1.SelLength = Len(vecy1.Text)

End Sub

Private Sub vecy1_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).vecy1 = vecy1.Text

End Sub

Private Sub vecy2_GotFocus()

    vecy2.SelStart = 0
    vecy2.SelLength = Len(vecy2.Text)

End Sub

Private Sub vecy2_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).vecy2 = vecy2.Text

End Sub

Private Sub life1_GotFocus()

    life1.SelStart = 0
    life1.SelLength = Len(life1.Text)

End Sub

Private Sub life1_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).life1 = life1.Text

End Sub

Private Sub life2_GotFocus()

    life2.SelStart = 0
    life2.SelLength = Len(life2.Text)

End Sub

Private Sub life2_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).life2 = life2.Text

End Sub

Private Sub fric_GotFocus()

    fric.SelStart = 0
    fric.SelLength = Len(fric.Text)

End Sub

Private Sub fric_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).friction = fric.Text

End Sub

Private Sub spin_speedL_GotFocus()

    spin_speedL.SelStart = 0
    spin_speedL.SelLength = Len(spin_speedH.Text)

End Sub

Private Sub spin_speedL_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).spin_speedL = spin_speedL.Text

End Sub

Private Sub spin_speedH_GotFocus()

    spin_speedH.SelStart = 0
    spin_speedH.SelLength = Len(spin_speedH.Text)

End Sub

Private Sub spin_speedH_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).spin_speedH = spin_speedH.Text

End Sub

Private Sub txtPCount_GotFocus()

    txtPCount.SelStart = 0
    txtPCount.SelLength = Len(txtPCount.Text)

End Sub

Private Sub txtPCount_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).NumOfParticles = txtPCount.Text

End Sub

Private Sub txtX1_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).x1 = txtX1.Text

End Sub

Private Sub txtX1_GotFocus()

    txtX1.SelStart = 0
    txtX1.SelLength = Len(txtX1.Text)

End Sub

Private Sub txtY1_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).y1 = txtY1.Text

End Sub

Private Sub txtY1_GotFocus()

    txtY1.SelStart = 0
    txtY1.SelLength = Len(txtY1.Text)

End Sub

Private Sub txtX2_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).x2 = txtX2.Text

End Sub

Private Sub txtX2_GotFocus()

    txtX2.SelStart = 0
    txtX2.SelLength = Len(txtX2.Text)

End Sub

Private Sub txtY2_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).y2 = txtY2.Text

End Sub

Private Sub txtY2_GotFocus()

    txtY2.SelStart = 0
    txtY2.SelLength = Len(txtY2.Text)

End Sub

Private Sub txtAngle_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).angle = txtAngle.Text

End Sub

Private Sub txtAngle_GotFocus()

    txtAngle.SelStart = 0
    txtAngle.SelLength = Len(txtAngle.Text)

End Sub

Private Sub txtGravStrength_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).grav_strength = txtGravStrength.Text

End Sub

Private Sub txtGravStrength_GotFocus()

    txtGravStrength.SelStart = 0
    txtGravStrength.SelLength = Len(txtGravStrength.Text)

End Sub

Private Sub txtBounceStrength_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).bounce_strength = txtBounceStrength.Text

End Sub

Private Sub txtBounceStrength_GotFocus()

    txtBounceStrength.SelStart = 0
    txtBounceStrength.SelLength = Len(txtBounceStrength.Text)

End Sub

Private Sub move_x1_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).move_x1 = move_x1.Text

End Sub

Private Sub move_x1_GotFocus()

    move_x1.SelStart = 0
    move_x1.SelLength = Len(move_x1.Text)

End Sub

Private Sub move_x2_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).move_x2 = move_x2.Text

End Sub

Private Sub move_x2_GotFocus()

    move_x2.SelStart = 0
    move_x2.SelLength = Len(move_x2.Text)

End Sub

Private Sub move_y1_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).move_y1 = move_y1.Text

End Sub

Private Sub move_y1_GotFocus()

    move_y1.SelStart = 0
    move_y1.SelLength = Len(move_y1.Text)

End Sub

Private Sub move_y2_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).move_y2 = move_y2.Text

End Sub

Private Sub move_y2_GotFocus()

    move_y2.SelStart = 0
    move_y2.SelLength = Len(move_y2.Text)

End Sub

Private Sub chkAlphaBlend_Click()

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).alphaBlend = chkAlphaBlend.value

End Sub

Private Sub chkGravity_Click()

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).gravity = chkGravity.value
    
    If chkGravity.value = vbChecked Then
        txtGravStrength.Enabled = True
        txtBounceStrength.Enabled = True
    Else
        txtGravStrength.Enabled = False
        txtBounceStrength.Enabled = False

    End If

End Sub

Private Sub chkXMove_Click()

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).XMove = chkXMove.value
    
    If chkXMove.value = vbChecked Then
        move_x1.Enabled = True
        move_x2.Enabled = True
    Else
        move_x1.Enabled = False
        move_x2.Enabled = False

    End If

End Sub

Private Sub chkYMove_Click()

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).YMove = chkYMove.value
    
    If chkYMove.value = vbChecked Then
        move_y1.Enabled = True
        move_y2.Enabled = True
    Else
        move_y1.Enabled = False
        move_y2.Enabled = False

    End If

End Sub

Private Sub BScroll_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).colortint(lstColorSets.ListIndex).B = BScroll.value
    txtB.Text = BScroll.value
    
    picColor.BackColor = RGB(txtB.Text, txtG.Text, txtR.Text)

End Sub

Private Sub chkNeverDies_Click()

    DataChanged = True
    
    If chkNeverDies.value = vbChecked Then
        life.Enabled = False
        StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).life_counter = -1
    Else
        life.Enabled = True
        StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).life_counter = life.Text

    End If

End Sub

Private Sub chkSpin_Click()

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).spin = chkSpin.value
    
    If chkSpin.value = vbChecked Then
        spin_speedL.Enabled = True
        spin_speedH.Enabled = True
    Else
        spin_speedL.Enabled = False
        spin_speedH.Enabled = False

    End If

End Sub

Private Sub GScroll_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).colortint(lstColorSets.ListIndex).G = GScroll.value
    txtG.Text = GScroll.value
    
    picColor.BackColor = RGB(txtB.Text, txtG.Text, txtR.Text)

End Sub

Private Sub life_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).life_counter = life.Text

End Sub

Private Sub life_GotFocus()

    life.SelStart = 0
    life.SelLength = Len(life.Text)

End Sub

Private Sub lstColorSets_Click()

    Dim DataTemp As Boolean

    DataTemp = DataChanged
    
    RScroll.value = StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).colortint(lstColorSets.ListIndex).R
    GScroll.value = StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).colortint(lstColorSets.ListIndex).G
    BScroll.value = StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).colortint(lstColorSets.ListIndex).B
    
    DataChanged = DataTemp

End Sub

Private Sub RScroll_Change()

    On Error Resume Next

    DataChanged = True
    
    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).colortint(lstColorSets.ListIndex).R = RScroll.value
    txtR.Text = RScroll.value
    
    picColor.BackColor = RGB(txtB.Text, txtG.Text, txtR.Text)

End Sub

Private Sub speed_Change()

    On Error Resume Next

    DataChanged = True
    
    'Arrange decimal separator
    Dim temp As String

    temp = GetVar(1, speed.Text, 44)

    If Not temp = "" Then
        speed.Text = temp & "." & Right(speed.Text, Len(speed.Text) - Len(temp) - 1)
        speed.SelStart = Len(speed.Text)
        speed.SelLength = 0

    End If

    StreamData(frmParticleEditor.ListParticulas.ListIndex + 1).speed = Val(speed.Text)

End Sub

Private Sub speed_GotFocus()

    speed.SelStart = 0
    speed.SelLength = Len(speed.Text)

End Sub

Private Sub lstSelGrhs_Click()
    GrhSelect(0) = frmParticleEditor.lstSelGrhs.List(frmParticleEditor.lstSelGrhs.ListIndex)

End Sub

Private Sub lstSelGrhs_DblClick()

    Call cmdDelete_Click

End Sub

Private Sub cmdAdd_Click()

    Dim loopc As Long
    
    If frmMain.LynxGrh.CellText(, 0) >= 0 Then lstSelGrhs.AddItem frmMain.LynxGrh.CellText(, 0)
    
    StreamData(ListParticulas.ListIndex + 1).NumGrhs = lstSelGrhs.ListCount
    
    ReDim StreamData(ListParticulas.ListIndex + 1).grh_list(1 To lstSelGrhs.ListCount) As Long
    
    For loopc = 1 To StreamData(ListParticulas.ListIndex + 1).NumGrhs
        StreamData(ListParticulas.ListIndex + 1).grh_list(loopc) = lstSelGrhs.List(loopc - 1)
    Next loopc

End Sub

