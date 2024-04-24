VERSION 5.00
Begin VB.Form frmMapSize 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tamaño del Mapa"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9690
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
   ScaleHeight     =   3345
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameAdapted 
      BackColor       =   &H00535353&
      Height          =   2775
      Left            =   90
      TabIndex        =   1
      Top             =   510
      Width           =   9495
      Begin VB.TextBox txtXMax 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "32000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1620
         TabIndex        =   3
         Text            =   "1000"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtYMax 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "32000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1620
         TabIndex        =   2
         Text            =   "1000"
         Top             =   720
         Width           =   1095
      End
      Begin WinterMapEditor.lvButtons_H LvBContinuar 
         Height          =   465
         Left            =   1920
         TabIndex        =   6
         Top             =   2130
         Width           =   2205
         _extentx        =   3889
         _extenty        =   820
         caption         =   "&Continuar"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapSize.frx":0000
         mode            =   0
         value           =   0
         cback           =   8454016
      End
      Begin WinterMapEditor.lvButtons_H LvBSalir 
         Height          =   465
         Left            =   300
         TabIndex        =   7
         Top             =   2130
         Width           =   2085
         _extentx        =   3678
         _extenty        =   820
         caption         =   "&Salir"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapSize.frx":0028
         mode            =   0
         value           =   0
         cback           =   8421631
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tamaño en X:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   150
         TabIndex        =   5
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tamaño en Y:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   150
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Label lblTuModo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tu modo activado solo permite ?x?."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   150
      TabIndex        =   0
      Top             =   30
      Width           =   5130
   End
End
Attribute VB_Name = "frmMapSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LvBSalir_Click()
    Unload Me
End Sub
