VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
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
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   720
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1280
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   18660
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10800
      Left            =   0
      ScaleHeight     =   720
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1280
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   420
      Width           =   19200
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa [Int]"
         Index           =   1
      End
      Begin VB.Menu mnuArchivoLine2 
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
      Begin VB.Menu mnuZonas 
         Caption         =   "Zonas"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub mnuSalir_Click()
    Call CloseMapEditor
End Sub

Private Sub mnuZonas_Click()
    frmZonas.Show , frmMain
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

Private Sub Form_Load()

    'Temporal
    ClientSetup.MapTam = 1
    Call setMapSize

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

    With MainViewPic
    
        .Height = Me.ScaleHeight - 28
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
    
End Sub
