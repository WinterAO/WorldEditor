VERSION 5.00
Begin VB.Form frmParticulas 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Particulas"
   ClientHeight    =   4200
   ClientLeft      =   25035
   ClientTop       =   9270
   ClientWidth     =   5115
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   341
   ShowInTaskbar   =   0   'False
   Begin WinterMapEditor.lvButtons_H cmdDel 
      Height          =   555
      Left            =   210
      TabIndex        =   3
      Top             =   3570
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   979
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
   Begin WinterMapEditor.lvButtons_H cmdAdd 
      Height          =   555
      Left            =   2310
      TabIndex        =   2
      Top             =   3570
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   979
      Caption         =   "Agregar"
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
   Begin VB.PictureBox ParticlePic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   2250
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   2805
   End
   Begin WinterMapEditor.LynxGrid LynxParticulas 
      Height          =   3375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   5953
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   5460819
      BackColorBkg    =   5460819
      BackColorEdit   =   14737632
      BackColorSel    =   12937777
      ForeColor       =   12632256
      ForeColorSel    =   8438015
      BackColorEvenRows=   3158064
      CustomColorFrom =   4210752
      CustomColorTo   =   8421504
      GridColor       =   14737632
      FocusRectColor  =   9895934
      GridLines       =   2
      ThemeColor      =   5
      ScrollBars      =   1
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      HotHeaderTracking=   0   'False
   End
End
Attribute VB_Name = "frmParticulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    If cmdAdd.value = True Then
        'LynxParticulas.Enabled = False
        cmdDel.Enabled = False
        
    Else
        LynxParticulas.Enabled = True
        cmdDel.Enabled = True
        
    End If
End Sub

Private Sub cmdDel_Click()
    If cmdDel.value = True Then
        LynxParticulas.Enabled = False
        cmdAdd.Enabled = False
        
    Else
        LynxParticulas.Enabled = True
        cmdAdd.Enabled = True
        
    End If
End Sub

Private Sub Form_Click()
'*************************************************
'Author: Lorwik
'Last modified: 26/04/2021
'*************************************************
    Me.SetFocus
    
End Sub

Private Sub LynxParticulas_Click()
'*************************************************
'Author: Lorwik
'Last modified: 29/04/2021
'*************************************************

    Call CargarInfo

End Sub

Private Sub LynxParticulas_LostFocus()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    HotKeysAllow = True
End Sub

Private Sub LynxParticulas_KeyDown(KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************
    On Error Resume Next
    
    Call CargarInfo
    
End Sub

Private Sub CargarInfo()
    Dim Index As Integer
    
    HotKeysAllow = False
    
    Index = LynxParticulas.CellText(, 0)
    
    ParticlePreview = General_Particle_Create(Index, -1, -1)
End Sub
