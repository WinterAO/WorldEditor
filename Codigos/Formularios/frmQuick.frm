VERSION 5.00
Begin VB.Form frmQuick 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quick Superficies"
   ClientHeight    =   4275
   ClientLeft      =   26130
   ClientTop       =   10245
   ClientWidth     =   2400
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
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   0
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   2400
   End
   Begin WinterMapEditor.lvButtons_H LvBLimpiarTodo 
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Top             =   3870
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   609
      Caption         =   "Limpiar Todo"
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
      Mode            =   0
      Value           =   0   'False
      cBack           =   8388736
   End
   Begin WinterMapEditor.lvButtons_H lvAdd 
      Height          =   345
      Left            =   90
      TabIndex        =   2
      Top             =   3450
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   609
      Caption         =   "Añadir"
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
      LockHover       =   1
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   65280
   End
   Begin WinterMapEditor.lvButtons_H LvBLimpiar 
      Height          =   345
      Left            =   1080
      TabIndex        =   3
      Top             =   3450
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      Caption         =   "Limpiar"
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
      LockHover       =   1
      cGradient       =   0
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   255
   End
End
Attribute VB_Name = "frmQuick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lvAdd_Click()
    '*********************************
    'Autor: Lorwik
    'Fecha: 30/04/2024
    'Descripcion: Añade una nueva superficie al inventario Quick
    '*********************************
    
    On Error GoTo lvAdd_Click_Err

    Dim GrhIndex As Long
    Dim FreeSlot As Byte
    
    GrhIndex = Val(frmSuperficies.cGrh.Text)
    
    '¿Es inferior a 1 o superior al numero de Grh disponibles?
    If GrhIndex < 1 Or GrhIndex > grhCount Then
        MsgBox "Superficie no valida."
        Exit Sub
    End If
    
    '¿Ya tiene esa superficie?
    If Inventario.GetItem(GrhIndex) > 0 Then
        MsgBox "Ya tienes esa superficie añadida."
        Exit Sub
    End If
    
    'Buscamos un slot libre
    FreeSlot = Inventario.GetFreeSlot
    
    'Si recibimos un 0 es por que no le queda espacio
    If FreeSlot = 0 Then
        MsgBox "No tienes espación para añadir mas superficies, elimina o reemplaza alguna (o todas)."
        Exit Sub
    End If
    
    'Si ya tenemos el slot, añadimos la superficie
    Call Inventario.SetItem(FreeSlot, Val(frmSuperficies.cGrh.Text))
    
    Exit Sub

lvAdd_Click_Err:

    Call RegistrarError(Err.Number, Err.Description, "frmQuick.lvAdd_Click", Erl)
    Resume Next
End Sub

Private Sub LvBLimpiarTodo_Click()
    '*********************************
    'Autor: Lorwik
    'Fecha: 30/04/2024
    '*********************************
    
    On Error GoTo LvBLimpiarTodo_Click_Err

    If MsgBox("¿¡Estás seguro de que quieres borrar todas las superficies añadidas!?", vbExclamation + vbYesNo) = vbYes Then _
        Call Inventario.ClearAllSlots
    
    Exit Sub

LvBLimpiarTodo_Click_Err:

    Call RegistrarError(Err.Number, Err.Description, "frmQuick.LvBLimpiarTodo_Click", Erl)
    Resume Next
End Sub

Private Sub picInv_Click()
    '*********************************
    'Autor: Lorwik
    'Fecha: 30/04/2024
    '*********************************
    
    '¿Es un slot valido?
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        frmSuperficies.cGrh.Text = Inventario.GrhIndex(Inventario.SelectedItem)
    
End Sub
