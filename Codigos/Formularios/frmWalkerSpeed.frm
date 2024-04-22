VERSION 5.00
Begin VB.Form frmWalkerSpeed 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Velocidad de movimiento"
   ClientHeight    =   645
   ClientLeft      =   22995
   ClientTop       =   6690
   ClientWidth     =   5280
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
   ScaleHeight     =   645
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSpeed 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   4590
      TabIndex        =   1
      Text            =   "0"
      Top             =   150
      Width           =   555
   End
   Begin VB.HScrollBar HScSpeed 
      Height          =   345
      Left            =   180
      Max             =   1000
      Min             =   10
      TabIndex        =   0
      Top             =   150
      Value           =   10
      Width           =   4305
   End
End
Attribute VB_Name = "frmWalkerSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'***************************************************

    Dim valorEntero As Integer
    
    valorEntero = Engine_BaseSpeed * 1000

    HScSpeed.value = valorEntero
    txtSpeed.Text = valorEntero

End Sub

Private Sub HScSpeed_Change()
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'***************************************************

    txtSpeed.Text = HScSpeed.value
    
    Call setSpeedbyOption

End Sub

Private Sub HScSpeed_LostFocus()
'*************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'*************************************************

    HotKeysAllow = True
End Sub

Private Sub txtSpeed_Change()
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'***************************************************

    If txtSpeed.Text > HScSpeed.Max Then txtSpeed.Text = HScSpeed.Max
    If txtSpeed.Text < HScSpeed.Min Then txtSpeed.Text = HScSpeed.Min
    
    HScSpeed.value = Val(txtSpeed)
    
    Call setSpeedbyOption
    
End Sub

Private Sub setSpeedbyOption()
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2024
'***************************************************

    Dim valorDecimal As Double
    valorDecimal = HScSpeed.value / 1000
    
    Engine_BaseSpeed = valorDecimal
    
End Sub
