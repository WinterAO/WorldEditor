VERSION 5.00
Begin VB.Form frmEditIndices 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor de Indices"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2775
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
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtGrhIndex 
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtCapa 
      Height          =   315
      Left            =   270
      TabIndex        =   9
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtAncho 
      Height          =   315
      Left            =   270
      TabIndex        =   6
      Top             =   1530
      Width           =   1065
   End
   Begin VB.TextBox txtAlto 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   1530
      Width           =   1125
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin WinterMapEditor.lvButtons_H LvBGuardar 
      Height          =   465
      Left            =   1440
      TabIndex        =   2
      Top             =   2670
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   820
      Caption         =   "Guardar"
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
      Mode            =   0
      Value           =   0   'False
      cBack           =   65280
   End
   Begin WinterMapEditor.lvButtons_H LvBSalir 
      Height          =   465
      Left            =   240
      TabIndex        =   3
      Top             =   2700
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   820
      Caption         =   "Salir"
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
      Mode            =   0
      Value           =   0   'False
      cBack           =   255
   End
   Begin WinterMapEditor.lvButtons_H LvBRecargarIndices 
      Height          =   435
      Left            =   540
      TabIndex        =   4
      Top             =   3180
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   767
      Caption         =   "Recargar Indices"
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
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632064
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GrhIndex"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   270
      TabIndex        =   12
      Top             =   720
      Width           =   675
   End
   Begin VB.Label lblCapa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Capa"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   300
      TabIndex        =   10
      Top             =   1890
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ancho"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   300
      TabIndex        =   8
      Top             =   1350
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alto"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1470
      TabIndex        =   7
      Top             =   1320
      Width           =   285
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   150
      Width           =   555
   End
End
Attribute VB_Name = "frmEditIndices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RecargarIndices As Boolean

Private Sub Form_Load()
    txtNombre.Text = SupData(frmSuperficies.LynxSuperficies.CellText(, 0)).name
    TxtGrhIndex.Text = SupData(frmSuperficies.LynxSuperficies.CellText(, 0)).Grh
    txtAncho.Text = SupData(frmSuperficies.LynxSuperficies.CellText(, 0)).Width
    txtAlto.Text = SupData(frmSuperficies.LynxSuperficies.CellText(, 0)).Height
    txtCapa.Text = SupData(frmSuperficies.LynxSuperficies.CellText(, 0)).Capa
End Sub

Private Sub LvBGuardar_Click()
    Call WriteVar(IniPath & INITDIR & "indices.ini", "REFERENCIA" & frmSuperficies.LynxSuperficies.CellText(, 0), "Nombre", txtNombre.Text)
    Call WriteVar(IniPath & INITDIR & "indices.ini", "REFERENCIA" & frmSuperficies.LynxSuperficies.CellText(, 0), "GrhIndex", TxtGrhIndex.Text)
    Call WriteVar(IniPath & INITDIR & "indices.ini", "REFERENCIA" & frmSuperficies.LynxSuperficies.CellText(, 0), "Ancho", txtAncho.Text)
    Call WriteVar(IniPath & INITDIR & "indices.ini", "REFERENCIA" & frmSuperficies.LynxSuperficies.CellText(, 0), "Alto", txtAlto.Text)
    Call WriteVar(IniPath & INITDIR & "indices.ini", "REFERENCIA" & frmSuperficies.LynxSuperficies.CellText(, 0), "Capa", txtCapa.Text)
    
    RecargarIndices = True
    
End Sub

Private Sub LvBRecargarIndices_Click()

    Call CargarIndicesSuperficie
    RecargarIndices = False
End Sub

Private Sub LvBSalir_Click()

    If RecargarIndices Then
        Select Case MsgBox("¡ATENCION! No has recargado los indices, no veras los cambios hasta que no los recarges ¿Deseas recargar?", vbYesNo Or vbExclamation, "¡ATENCION!")

            Case vbYes
                Call CargarIndicesSuperficie
                RecargarIndices = False
                
            Case vbNo
                Unload Me

        End Select
        
    Else
        Unload Me
        
    End If

End Sub
