VERSION 5.00
Begin VB.Form frmConfiguracion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración del Editor"
   ClientHeight    =   4785
   ClientLeft      =   14820
   ClientTop       =   8040
   ClientWidth     =   5295
   ClipControls    =   0   'False
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
   ScaleHeight     =   4785
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraMiscelanea 
      Caption         =   "Miscelanea"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   210
      TabIndex        =   5
      Top             =   3150
      Width           =   4965
      Begin VB.CheckBox chkop 
         Caption         =   "Referencia de campo de visión del jugador"
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   330
         Width           =   3705
      End
   End
   Begin VB.Frame FraConfiguraciónDe 
      Caption         =   "Configuración de video"
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
      Height          =   2835
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   4995
      Begin VB.HScrollBar HScDinamicMemory 
         Height          =   285
         Left            =   180
         Max             =   50
         Min             =   1
         TabIndex        =   12
         Top             =   2370
         Value           =   5
         Width           =   3915
      End
      Begin VB.TextBox txtDinamicMemory 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4290
         TabIndex        =   11
         Text            =   "0"
         Top             =   2370
         Width           =   555
      End
      Begin VB.TextBox txtBuffer 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4290
         TabIndex        =   8
         Text            =   "0"
         Top             =   1590
         Width           =   555
      End
      Begin VB.HScrollBar HScTileBuffer 
         Height          =   285
         Left            =   180
         Max             =   50
         Min             =   1
         TabIndex        =   7
         Top             =   1590
         Value           =   5
         Width           =   3915
      End
      Begin VB.ComboBox cmbProcesado 
         Height          =   315
         ItemData        =   "frmConfiguracion.frx":0000
         Left            =   1920
         List            =   "frmConfiguracion.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   400
         Width           =   2895
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Activar sincronización vertical"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   930
         Width           =   3015
      End
      Begin VB.Label lblMemoriaDinamica 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Memoria Dinamica:"
         Height          =   195
         Left            =   300
         TabIndex        =   10
         Top             =   2100
         Width           =   1335
      End
      Begin VB.Label lblTamañoDel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tamaño del Tile Buffer:"
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   1350
         Width           =   1665
      End
      Begin VB.Label lblModoDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modo de procesado:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1470
      End
   End
   Begin WinterMapEditor.lvButtons_H LvBCerrar 
      Height          =   480
      Index           =   13
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Cerrar"
      Top             =   4110
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   847
      Caption         =   "Cerrar"
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
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      ImgSize         =   32
      cBack           =   8421631
   End
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private isChanged As Boolean

Private Sub Form_Load()
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'***************************************************

    If Not leerOpciones Then
        If MsgBox("Hubo un error al leer las opciones ¿Quieres continuar?", vbExclamation + vbYesNo) = vbNo Then
            Unload Me
            Exit Sub
        End If
    End If

    isChanged = False
    
End Sub

Private Function leerOpciones() As Boolean
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'***************************************************

    On Error GoTo leerOpciones_Err

    cmbProcesado.ListIndex = ClientSetup.OverrideVertexProcess
    
    'Campo de visión
    If ClientSetup.CampoVision Then
        chkop(0).value = Checked
        
    Else
        chkop(0).value = Unchecked
        
    End If
    
    'Sincronización vertical
    If ClientSetup.LimiteFPS Then
        chkop(1).value = Checked
        
    Else
        chkop(1).value = Unchecked
        
    End If
    
    'TileBuffer
    HScTileBuffer.value = ClientSetup.TilesBuffer
    txtBuffer.Text = ClientSetup.TilesBuffer
    
    'Memoria dinamica
    HScDinamicMemory.value = ClientSetup.byMemory
    txtDinamicMemory.Text = ClientSetup.byMemory
    
    leerOpciones = True

    Exit Function

leerOpciones_Err:

    Call RegistrarError(Err.Number, Err.Description, "frmConfiguracion.leerOpciones", Erl)
    Resume Next
    
End Function

Private Sub chkop_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'Activa y desactiva las opciones con CheckBox
'***************************************************
    On Error GoTo chkop_MouseUp_Err

    Select Case Index
    
        Case 0 'Campo de visión
        
            If chkop(Index).value = vbUnchecked Then
                ClientSetup.CampoVision = False
                frmMain.ShpScreen.Visible = False
                
            Else
                ClientSetup.CampoVision = True
                frmMain.ShpScreen.Visible = True
                
            End If
            
        Case 1 'Limite de FPS
            If chkop(Index).value = vbUnchecked Then
                ClientSetup.LimiteFPS = False
                
            Else
                ClientSetup.LimiteFPS = True
                
            End If
    
    End Select
    
    'Marcamos que se hicieron cambios
    isChanged = True

    Exit Sub

chkop_MouseUp_Err:

    Call RegistrarError(Err.Number, Err.Description, "frmConfiguracion.chkop_MouseUp", Erl)
    Resume Next
End Sub

Private Sub HScTileBuffer_Change()
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'***************************************************

    txtBuffer.Text = HScTileBuffer.value
    ClientSetup.TilesBuffer = HScTileBuffer.value
    
    TileBufferSize = HScTileBuffer.value
    
    'Marcamos que se hicieron cambios
    isChanged = True
    
End Sub

Private Sub txtBuffer_Change()
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'***************************************************

    If Val(txtBuffer.Text) > HScTileBuffer.Max Then txtBuffer.Text = HScTileBuffer.Max
    If Val(txtBuffer.Text) < HScTileBuffer.Min Then txtBuffer.Text = HScTileBuffer.Min

    HScTileBuffer.value = Val(txtBuffer.Text)
    ClientSetup.TilesBuffer = HScTileBuffer.value
    
    TileBufferSize = HScTileBuffer.value
    
    'Marcamos que se hicieron cambios
    isChanged = True
    
End Sub

Private Sub HScDinamicMemory_Change()
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'***************************************************

    txtDinamicMemory.Text = HScDinamicMemory.value
    ClientSetup.byMemory = HScDinamicMemory.value
    
    'Marcamos que se hicieron cambios
    isChanged = True
End Sub

Private Sub txtDinamicMemory_Change()
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'***************************************************

    If Val(txtDinamicMemory.Text) > HScDinamicMemory.Max Then txtDinamicMemory.Text = HScDinamicMemory.Max
    If Val(txtDinamicMemory.Text) < HScDinamicMemory.Min Then txtDinamicMemory.Text = HScDinamicMemory.Min

    HScDinamicMemory.value = Val(txtDinamicMemory.Text)
    ClientSetup.byMemory = HScDinamicMemory.value
    
    'Marcamos que se hicieron cambios
    isChanged = True
    
End Sub

Private Sub LvBCerrar_Click(Index As Integer)
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'***************************************************

    On Error GoTo LvBCerrar_Click_Err

    If isChanged Then
        If MsgBox("Se hicieron cambios en la configuración ¿Quieres guardar estos cambios?", vbExclamation + vbYesNo) = vbYes Then
        
            Call guardarPerfilVideo
            Call AddtoRichTextBox(frmConsola.StatTxt, "La configuración se guardo satisfactoriamente.", 0, 255, 0, , , , , True)
            
        End If
    End If

    isChanged = False

    Unload Me
    '<EhFooter>
    Exit Sub

LvBCerrar_Click_Err:

    Call AddtoRichTextBox(frmConsola.StatTxt, "Error al guardar la configuración!", 255, 0, 0, , , , , True)
    Call RegistrarError(Err.Number, Err.Description, "frmConfiguracion.LvBCerrar_Click", Erl)
    Resume Next

End Sub
