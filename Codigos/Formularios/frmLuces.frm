VERSION 5.00
Begin VB.Form frmLuces 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Luces"
   ClientHeight    =   3570
   ClientLeft      =   24630
   ClientTop       =   10890
   ClientWidth     =   3975
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
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin WinterMapEditor.lvButtons_H lvButtons_H5 
      Height          =   345
      Index           =   0
      Left            =   2310
      TabIndex        =   15
      Top             =   240
      Width           =   495
      _extentx        =   873
      _extenty        =   609
      capalign        =   2
      cgradient       =   0
      font            =   "frmLuces.frx":0000
      mode            =   0
      value           =   0
      cback           =   255
   End
   Begin WinterMapEditor.lvButtons_H cInsertarLuz 
      Height          =   360
      Left            =   1800
      TabIndex        =   9
      Top             =   1680
      Width           =   1665
      _extentx        =   2937
      _extenty        =   635
      caption         =   "Insertar Luz"
      capalign        =   2
      backstyle       =   2
      shape           =   1
      cgradient       =   0
      font            =   "frmLuces.frx":0028
      mode            =   1
      value           =   0
      cback           =   -2147483633
   End
   Begin WinterMapEditor.lvButtons_H cQuitarLuz 
      Height          =   360
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   1665
      _extentx        =   2937
      _extenty        =   635
      caption         =   "Quitar Luz"
      capalign        =   2
      backstyle       =   2
      shape           =   2
      cgradient       =   0
      font            =   "frmLuces.frx":0050
      mode            =   1
      value           =   0
      cback           =   -2147483633
   End
   Begin VB.Frame FraLuzAmbiental 
      Appearance      =   0  'Flat
      BackColor       =   &H00535353&
      Caption         =   "Luz Ambiental"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   3735
      Begin WinterMapEditor.lvButtons_H lvButtons_H1 
         Height          =   360
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1545
         _extentx        =   2725
         _extenty        =   635
         caption         =   "Mañana"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmLuces.frx":0078
         mode            =   2
         value           =   0
         cback           =   8438015
      End
      Begin WinterMapEditor.lvButtons_H lvButtons_H1 
         Height          =   360
         Index           =   1
         Left            =   1920
         TabIndex        =   12
         Top             =   360
         Width           =   1545
         _extentx        =   2725
         _extenty        =   635
         caption         =   "Dia"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmLuces.frx":00A0
         mode            =   2
         value           =   0
         cback           =   16777088
      End
      Begin WinterMapEditor.lvButtons_H lvButtons_H1 
         Height          =   360
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1545
         _extentx        =   2725
         _extenty        =   635
         caption         =   "Tarde"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmLuces.frx":00C8
         mode            =   2
         value           =   0
         cback           =   8421504
      End
      Begin WinterMapEditor.lvButtons_H lvButtons_H1 
         Height          =   360
         Index           =   3
         Left            =   1920
         TabIndex        =   14
         Top             =   840
         Width           =   1545
         _extentx        =   2725
         _extenty        =   635
         caption         =   "Noche"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmLuces.frx":00F0
         mode            =   2
         value           =   0
         cback           =   4210752
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00535353&
      Caption         =   "Rango"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1380
      Begin VB.TextBox cRango 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   105
         TabIndex        =   5
         Text            =   "1"
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(1 al 50)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Frame RGBCOLOR 
      BackColor       =   &H00535353&
      Caption         =   "RGB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1680
      Begin VB.TextBox G 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   600
         TabIndex        =   3
         Text            =   "255"
         Top             =   270
         Width           =   450
      End
      Begin VB.TextBox B 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1095
         TabIndex        =   2
         Text            =   "255"
         Top             =   270
         Width           =   450
      End
      Begin VB.TextBox R 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   105
         TabIndex        =   1
         Text            =   "200"
         Top             =   270
         Width           =   450
      End
   End
   Begin WinterMapEditor.lvButtons_H lvButtons_H5 
      Height          =   345
      Index           =   1
      Left            =   2910
      TabIndex        =   16
      Top             =   240
      Width           =   495
      _extentx        =   873
      _extenty        =   609
      capalign        =   2
      cgradient       =   0
      font            =   "frmLuces.frx":0118
      mode            =   0
      value           =   0
      cback           =   65535
   End
   Begin WinterMapEditor.lvButtons_H lvButtons_H5 
      Height          =   345
      Index           =   2
      Left            =   2910
      TabIndex        =   17
      Top             =   1110
      Width           =   495
      _extentx        =   873
      _extenty        =   609
      capalign        =   2
      cgradient       =   0
      font            =   "frmLuces.frx":0140
      mode            =   0
      value           =   0
      cback           =   12632256
   End
   Begin WinterMapEditor.lvButtons_H lvButtons_H5 
      Height          =   345
      Index           =   3
      Left            =   2310
      TabIndex        =   18
      Top             =   1110
      Width           =   495
      _extentx        =   873
      _extenty        =   609
      capalign        =   2
      cgradient       =   0
      font            =   "frmLuces.frx":0168
      mode            =   0
      value           =   0
      cback           =   16711935
   End
   Begin WinterMapEditor.lvButtons_H lvButtons_H5 
      Height          =   345
      Index           =   4
      Left            =   2910
      TabIndex        =   19
      Top             =   660
      Width           =   495
      _extentx        =   873
      _extenty        =   609
      capalign        =   2
      cgradient       =   0
      font            =   "frmLuces.frx":0190
      mode            =   0
      value           =   0
      cback           =   16777215
   End
   Begin WinterMapEditor.lvButtons_H lvButtons_H5 
      Height          =   345
      Index           =   5
      Left            =   2310
      TabIndex        =   11
      Top             =   660
      Width           =   495
      _extentx        =   873
      _extenty        =   609
      capalign        =   2
      cgradient       =   0
      font            =   "frmLuces.frx":01B8
      mode            =   0
      value           =   0
      cback           =   16776960
   End
End
Attribute VB_Name = "frmLuces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cInsertarLuz_Click()
'*********************************
'Author: Lorwik
'Fecha: 21/03/2012
'*********************************
    If cInsertarLuz.value Then
        cQuitarLuz.Enabled = False
    Else
        cQuitarLuz.Enabled = True
    End If
End Sub

Private Sub cQuitarLuz_Click()
'*********************************
'Author: Lorwik
'Fecha: 21/03/2012
'*********************************
    If cQuitarLuz.value Then
        cInsertarLuz.Enabled = False
    Else
        cInsertarLuz.Enabled = True
    End If
End Sub

Public Sub AccionLuces()
    cInsertarLuz.value = False
    Call cInsertarLuz_Click
    cQuitarLuz.value = False
    Call cQuitarLuz_Click
End Sub

Private Sub lvButtons_H1_Click(Index As Integer)

    If frmMapInfo.chkLuzClimatica.value = Checked Then
        MsgBox "No disponible con la luz base activada"
        Exit Sub
    End If

    Select Case Index
    
        Case 0
            Estado_Actual = Estados(e_estados.Amanecer)
            
        Case 1
            Estado_Actual = Estados(e_estados.MedioDia)
            
        Case 2
            Estado_Actual = Estados(e_estados.Tarde)
            
        Case 3
            Estado_Actual = Estados(e_estados.noche)
    
    End Select
    
    Call Actualizar_Estado
    
End Sub

Private Sub lvButtons_H5_Click(Index As Integer)
    Select Case Index
    
        Case 0
            R = 255
            G = 0
            B = 0
            
        Case 1
            R = 255
            G = 255
            B = 0
            
        Case 2
            R = 192
            G = 192
            B = 192
            
        Case 3
            R = 255
            G = 0
            B = 255
            
        Case 4
            R = 255
            G = 255
            B = 255
            
        Case 5
            R = 127
            G = 255
            B = 255
            
    End Select
End Sub
