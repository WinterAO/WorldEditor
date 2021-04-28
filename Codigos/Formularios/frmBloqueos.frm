VERSION 5.00
Begin VB.Form frmBloqueos 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bloqueos"
   ClientHeight    =   1800
   ClientLeft      =   17310
   ClientTop       =   14325
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
      _extentx        =   7144
      _extenty        =   1402
      caption         =   "Quitar Bloqueos"
      capalign        =   2
      backstyle       =   2
      gradient        =   3
      cgradient       =   0
      cfore           =   16777215
      font            =   "frmBloqueos.frx":0000
      mode            =   1
      value           =   0   'False
      cfhover         =   16777215
      cback           =   255
      cbhover         =   0
      lockhover       =   1
   End
   Begin WinterMapEditor.lvButtons_H cInsertarBloqueo 
      Height          =   705
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   4050
      _extentx        =   7144
      _extenty        =   1244
      caption         =   "Insertar Bloqueos"
      capalign        =   2
      backstyle       =   2
      gradient        =   3
      cgradient       =   0
      cfore           =   16777215
      font            =   "frmBloqueos.frx":0028
      mode            =   1
      value           =   0   'False
      cfhover         =   16777215
      cback           =   65280
      cbhover         =   0
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
