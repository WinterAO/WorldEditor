VERSION 5.00
Begin VB.Form frmOBJs 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Objetos"
   ClientHeight    =   5925
   ClientLeft      =   10200
   ClientTop       =   6240
   ClientWidth     =   4440
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
   ScaleHeight     =   395
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cFiltro 
      BackColor       =   &H80000012&
      ForeColor       =   &H80000014&
      Height          =   315
      Left            =   570
      TabIndex        =   3
      Top             =   4170
      Width           =   3765
   End
   Begin VB.ComboBox cNumFunc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   3450
      TabIndex        =   2
      Text            =   "1"
      Top             =   4530
      Width           =   855
   End
   Begin VB.ComboBox cCantFunc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   810
      TabIndex        =   1
      Text            =   "1"
      Top             =   4530
      Width           =   1215
   End
   Begin WinterMapEditor.LynxGrid LynxOBJs 
      Height          =   3975
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   7011
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
   Begin WinterMapEditor.lvButtons_H cAgregarFuncalAzar 
      Height          =   405
      Left            =   150
      TabIndex        =   4
      Top             =   5040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   714
      Caption         =   "Insertar OBJ's al Azar"
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
      cBack           =   65535
   End
   Begin WinterMapEditor.lvButtons_H cInsertarFunc 
      Height          =   795
      Left            =   2400
      TabIndex        =   5
      Top             =   5040
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1402
      Caption         =   "Insertar OBJ's"
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
   Begin WinterMapEditor.lvButtons_H cQuitarFunc 
      Height          =   405
      Left            =   150
      TabIndex        =   6
      Top             =   5460
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   714
      Caption         =   "Quitar OBJ's"
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
      cBack           =   255
   End
   Begin VB.Label lbFiltrar 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar:"
      ForeColor       =   &H80000014&
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   4200
      Width           =   570
   End
   Begin VB.Label lNumFunc 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de OBJ:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Index           =   1
      Left            =   2220
      TabIndex        =   8
      Top             =   4605
      Width           =   1170
   End
   Begin VB.Label lCantFunc 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   4605
      Width           =   675
   End
End
Attribute VB_Name = "frmOBJs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
'*************************************************
'Author: Lorwik
'Last modified: 26/04/2021
'*************************************************
    Me.SetFocus

End Sub

Private Sub LynxOBJs_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2021
    '*************************************************
        
    Call CargarInfo
    
End Sub

Private Sub LynxOBJs_LostFocus()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    HotKeysAllow = True
End Sub

Private Sub LynxOBJs_KeyDown(KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************
On Error Resume Next
    
    Call CargarInfo
End Sub

Private Sub cFiltro_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    If KeyAscii = 13 Then
        Call Filtrar
    End If
End Sub

Private Sub cFiltro_LostFocus()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    HotKeysAllow = True
End Sub

Private Sub cInsertarFunc_Click()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'*************************************************
    If cInsertarFunc.value = True Then
        cQuitarFunc.Enabled = False
        cAgregarFuncalAzar.Enabled = False
        cCantFunc.Enabled = False
    Else
        cQuitarFunc.Enabled = True
        cAgregarFuncalAzar.Enabled = True
        cCantFunc.Enabled = True
    End If
End Sub

Private Sub cAgregarFuncalAzar_Click()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'*************************************************
    On Error Resume Next
    If IsNumeric(cCantFunc.Text) = False Or cCantFunc.Text > 200 Then
        MsgBox "El Valor de Cantidad introducido no es soportado!" & vbCrLf & "El valor maximo es 200.", vbCritical
        Exit Sub
    End If
    
    cAgregarFuncalAzar.Enabled = False
    'Call PonerAlAzar(CInt(cCantFunc.Text))
    cAgregarFuncalAzar.Enabled = True
End Sub

Private Sub cQuitarFunc_Click()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'*************************************************
    If cQuitarFunc.value = True Then
        cInsertarFunc.Enabled = False
        cAgregarFuncalAzar.Enabled = False
        cCantFunc.Enabled = False
        cNumFunc.Enabled = False
        cFiltro.Enabled = False
        LynxOBJs.Enabled = False
        
    Else
        cInsertarFunc.Enabled = True
        cAgregarFuncalAzar.Enabled = True
        cCantFunc.Enabled = True
        cNumFunc.Enabled = True
        cFiltro.Enabled = True
        LynxOBJs.Enabled = True
        
    End If
End Sub

Private Sub CargarInfo()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    Dim GrhObj As Long
    
    HotKeysAllow = False
        
    'Obtiene el numero del Grh
    cNumFunc.Text = LynxOBJs.CellText(, 0)
    
    If frmPreview.Visible Then
        Call fPreviewGrh(ObjData(cNumFunc.Text).GrhIndex)
        Call RenderPreview(True)

    End If
 
End Sub

Private Sub Filtrar()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'Optimized: 27/09/2025
'*************************************************

    Dim vDatos As String
    Dim NumI As Integer
    Dim i As Integer
    Dim k As Long
    
    ' Mantener solo los últimos 5 elementos en el filtro
    If cFiltro.ListCount > 5 Then _
        cFiltro.RemoveItem 0
    
    cFiltro.AddItem cFiltro.Text
    
    LynxOBJs.Clear
    LynxOBJs.Redraw = False
    LynxOBJs.Visible = False
    
    For i = 1 To NumObjs
        vDatos = ObjData(i).name
        NumI = i
        
        ' Usamos InStr en lugar de recorrer cada carácter
        If InStr(1, vDatos & CStr(i), cFiltro.Text, vbTextCompare) > 0 Or LenB(cFiltro.Text) = 0 Then
            LynxOBJs.AddItem NumI
            k = LynxOBJs.Rows - 1
            LynxOBJs.CellText(k, 1) = vDatos
        End If
    Next i
    
    LynxOBJs.Visible = True
    LynxOBJs.Redraw = True
    LynxOBJs.ColForceFit
    
    DoEvents

End Sub

