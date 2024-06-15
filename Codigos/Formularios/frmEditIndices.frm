VERSION 5.00
Begin VB.Form frmEditIndices 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor de Indices"
   ClientHeight    =   3750
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
   ScaleHeight     =   250
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
      _extentx        =   2196
      _extenty        =   820
      caption         =   "Guardar"
      capalign        =   2
      backstyle       =   2
      font            =   "frmEditIndices.frx":0000
      cfore           =   16777215
      cfhover         =   16777215
      cbhover         =   0
      cgradient       =   0
      gradient        =   3
      mode            =   0
      value           =   0   'False
      cback           =   65280
   End
   Begin WinterMapEditor.lvButtons_H LvBSalir 
      Height          =   465
      Left            =   240
      TabIndex        =   3
      Top             =   2700
      Width           =   1155
      _extentx        =   2037
      _extenty        =   820
      caption         =   "Salir"
      capalign        =   2
      backstyle       =   2
      font            =   "frmEditIndices.frx":0028
      cfore           =   16777215
      cfhover         =   16777215
      cbhover         =   0
      cgradient       =   0
      gradient        =   3
      mode            =   0
      value           =   0   'False
      cback           =   255
   End
   Begin WinterMapEditor.lvButtons_H LvBRecargarIndices 
      Height          =   405
      Left            =   210
      TabIndex        =   4
      Top             =   3210
      Width           =   2445
      _extentx        =   4313
      _extenty        =   714
      caption         =   "Recargar Indices"
      capalign        =   2
      backstyle       =   2
      font            =   "frmEditIndices.frx":0050
      cfore           =   16777215
      cfhover         =   16777215
      cbhover         =   0
      cgradient       =   0
      gradient        =   3
      mode            =   0
      value           =   0   'False
      cback           =   12632064
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

Private Sub Form_Load()
'***********************************
'Autor: Lorwik
'Fecha: ???
'***********************************

    Call CargarData
    
End Sub

Public Sub CargarData()
'***********************************
'Autor: Lorwik
'Fecha: ???
'***********************************
    On Error GoTo CargarData_Err

    txtNombre.Text = SupData(frmSuperficies.LynxSuperficies.CellText(, 0)).name
    TxtGrhIndex.Text = SupData(frmSuperficies.LynxSuperficies.CellText(, 0)).Grh
    txtAncho.Text = SupData(frmSuperficies.LynxSuperficies.CellText(, 0)).Width
    txtAlto.Text = SupData(frmSuperficies.LynxSuperficies.CellText(, 0)).Height
    txtCapa.Text = SupData(frmSuperficies.LynxSuperficies.CellText(, 0)).Capa
    
    Exit Sub

CargarData_Err:

    Call RegistrarError(Err.Number, Err.Description, "frmEditIndices.CargarData", Erl)
    Resume Next

End Sub

Private Sub LvBGuardar_Click()
'***********************************
'Autor: Lorwik
'Fecha: ???
'***********************************
    On Error GoTo LvBGuardar_Click_Err

    Call WriteVar(IniPath & INITDIR & "indices.ini", "REFERENCIA" & frmSuperficies.LynxSuperficies.CellText(, 0), "Nombre", txtNombre.Text)
    Call WriteVar(IniPath & INITDIR & "indices.ini", "REFERENCIA" & frmSuperficies.LynxSuperficies.CellText(, 0), "GrhIndex", TxtGrhIndex.Text)
    Call WriteVar(IniPath & INITDIR & "indices.ini", "REFERENCIA" & frmSuperficies.LynxSuperficies.CellText(, 0), "Ancho", txtAncho.Text)
    Call WriteVar(IniPath & INITDIR & "indices.ini", "REFERENCIA" & frmSuperficies.LynxSuperficies.CellText(, 0), "Alto", txtAlto.Text)
    Call WriteVar(IniPath & INITDIR & "indices.ini", "REFERENCIA" & frmSuperficies.LynxSuperficies.CellText(, 0), "Capa", txtCapa.Text)

    Exit Sub

LvBGuardar_Click_Err:

    Call RegistrarError(Err.Number, Err.Description, "frmEditIndices.LvBGuardar_Click", Erl)
    Resume Next

End Sub

Private Sub LvBRecargarIndices_Click()
'***********************************
'Autor: Lorwik
'Fecha: ???
'***********************************

    Call CargarIndicesSuperficie
    
End Sub

Private Sub LvBSalir_Click()
'***********************************
'Autor: Lorwik
'Fecha: ???
'***********************************

    Unload Me

End Sub
