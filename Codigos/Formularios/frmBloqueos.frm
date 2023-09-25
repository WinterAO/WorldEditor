VERSION 5.00
Begin VB.Form frmBloqueos 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bloqueos"
   ClientHeight    =   1800
   ClientLeft      =   9645
   ClientTop       =   6090
   ClientWidth     =   4335
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
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   Begin WinterMapEditor.lvButtons_H cQuitarBloqueo 
      Height          =   795
      Left            =   150
      TabIndex        =   1
      Top             =   900
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   1402
      Caption         =   "Quitar Bloqueos"
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
      LockHover       =   1
      cGradient       =   0
      Gradient        =   3
      Mode            =   1
      Value           =   0   'False
      cBack           =   255
   End
   Begin WinterMapEditor.lvButtons_H cInsertarBloqueo 
      Height          =   705
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   1244
      Caption         =   "Insertar Bloqueos"
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
      Mode            =   1
      Value           =   0   'False
      cBack           =   65280
   End
End
Attribute VB_Name = "frmBloqueos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cInsertarBloqueo_Click()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'*************************************************
    cInsertarBloqueo.Tag = vbNullString
    If cInsertarBloqueo.value = True Then
        cQuitarBloqueo.Enabled = False
        
    Else
        cQuitarBloqueo.Enabled = True
        
    End If
End Sub

Private Sub cQuitarBloqueo_Click()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'*************************************************
    cInsertarBloqueo.Tag = vbNullString
    If cQuitarBloqueo.value = True Then
        cInsertarBloqueo.Enabled = False
        
    Else
        cInsertarBloqueo.Enabled = True
        
    End If
End Sub
