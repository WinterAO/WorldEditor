VERSION 5.00
Begin VB.Form frmSuperficies 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Superficies"
   ClientHeight    =   6600
   ClientLeft      =   9420
   ClientTop       =   6465
   ClientWidth     =   4455
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkBloquear 
      BackColor       =   &H00424242&
      Caption         =   "Bloquear"
      Height          =   285
      Left            =   3210
      TabIndex        =   12
      Top             =   4980
      Width           =   975
   End
   Begin VB.CheckBox chkAutoCompletarSuperficies 
      BackColor       =   &H00424242&
      Caption         =   "Auto-Completar"
      ForeColor       =   &H8000000B&
      Height          =   345
      Left            =   180
      TabIndex        =   11
      Top             =   4920
      Width           =   1485
   End
   Begin WinterMapEditor.lvButtons_H cQuitarEnTodasLasCapas 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Quitar en Capas 2 y 3"
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
      cBack           =   8421631
   End
   Begin WinterMapEditor.lvButtons_H cQuitarEnEstaCapa 
      Height          =   405
      Left            =   120
      TabIndex        =   7
      Top             =   5340
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   714
      Caption         =   "Quitar en esta Capa"
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
   Begin WinterMapEditor.lvButtons_H cSeleccionarSuperficie 
      Height          =   795
      Left            =   2460
      TabIndex        =   6
      Top             =   5340
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1402
      Caption         =   "Insertar Superficie"
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
   Begin VB.ComboBox cCapas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      ForeColor       =   &H80000014&
      Height          =   315
      ItemData        =   "frmSuperficies.frx":0000
      Left            =   1080
      List            =   "frmSuperficies.frx":0002
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   4560
      Width           =   855
   End
   Begin VB.ComboBox cFiltro 
      BackColor       =   &H80000012&
      ForeColor       =   &H80000014&
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   4170
      Width           =   3765
   End
   Begin VB.ComboBox cGrh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      ForeColor       =   &H80000014&
      Height          =   315
      Left            =   3000
      TabIndex        =   0
      Text            =   "1"
      Top             =   4560
      Width           =   1335
   End
   Begin WinterMapEditor.lvButtons_H LvBEditarIndice 
      Height          =   375
      Left            =   180
      TabIndex        =   9
      Top             =   6180
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      Caption         =   "Editar Indice"
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
      cBack           =   16744576
   End
   Begin WinterMapEditor.LynxGrid LynxSuperficies 
      Height          =   3975
      Left            =   90
      TabIndex        =   10
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
   Begin VB.Label lbGrh 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Sup Actual:"
      ForeColor       =   &H80000014&
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   4620
      Width           =   825
   End
   Begin VB.Label lbCapas 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Capa Actual:"
      ForeColor       =   &H80000014&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   4605
      Width           =   930
   End
   Begin VB.Label lbFiltrar 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar:"
      ForeColor       =   &H80000014&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   4230
      Width           =   480
   End
End
Attribute VB_Name = "frmSuperficies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cSeleccionarSuperficie_Click()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'*************************************************
    If cSeleccionarSuperficie.value = True Then
        cQuitarEnTodasLasCapas.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        
    Else
        cQuitarEnTodasLasCapas.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        
    End If
End Sub

Private Sub cQuitarEnEstaCapa_Click()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'*************************************************
    If cQuitarEnEstaCapa.value = True Then
        LynxSuperficies.Enabled = False
        cFiltro.Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnTodasLasCapas.Enabled = False
        
    Else
        LynxSuperficies.Enabled = True
        cFiltro.Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnTodasLasCapas.Enabled = True
        
    End If
End Sub

Private Sub cQuitarEnTodasLasCapas_Click()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'*************************************************
    If cQuitarEnTodasLasCapas.value = True Then
        cCapas.Enabled = False
        LynxSuperficies.Enabled = False
        cFiltro.Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        
    Else
        cCapas.Enabled = True
        LynxSuperficies.Enabled = True
        cFiltro.Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        
    End If
End Sub

Private Sub Form_Load()
'*************************************************
'Author: Lorwik
'Last modified: 01/05/2021
'*************************************************
    Dim i As Byte
    
    For i = 1 To 4
        cCapas.AddItem i
    Next i
End Sub

Private Sub Form_Click()
'*************************************************
'Author: Lorwik
'Last modified: 26/04/2021
'*************************************************
    Me.SetFocus

End Sub

Private Sub LvBEditarIndice_Click()
    frmEditIndices.Show , frmMain
End Sub

Private Sub LynxSuperficies_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2021
    '*************************************************
        
    Call CargarInfo
    
End Sub

Private Sub LynxSuperficies_LostFocus()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    HotKeysAllow = True
End Sub

Private Sub LynxSuperficies_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub CargarInfo()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    HotKeysAllow = False
    SupActual = LynxSuperficies.CellText(, 0)
        
    'Obtiene el numero del Grh
    cGrh.Text = DameGrhIndex(SupActual)
    
    chkBloquear.value = SupData(SupActual).Block

    'TODO: Faltan movidas aqui
    If SupData(SupActual).Capa <> 0 Then
        If (SupActual) = 0 Then cCapas.tag = cCapas.Text
        cCapas.Text = SupData(SupActual).Capa
            
    Else

        If LenB(cCapas.tag) <> 0 Then
            cCapas.Text = cCapas.tag
            cCapas.tag = vbNullString
                
        End If
            
    End If
        
    'Manda a renderizar la superficie seleccionada
    If frmPreview.Visible Then
        Call fPreviewGrh(cGrh.Text)
        Call RenderPreview

    End If
End Sub

Private Sub Filtrar()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    Dim vDatos As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Long
    
    If cFiltro.ListCount > 5 Then _
        cFiltro.RemoveItem 0
    
    cFiltro.AddItem cFiltro.Text
    
    LynxSuperficies.Clear
    LynxSuperficies.Redraw = False
    LynxSuperficies.Visible = False
    
    For i = 0 To MaxSup
        vDatos = SupData(i).Name
        
        For j = 1 To Len(vDatos)
        
            If UCase$(mid$(vDatos & str(i), j, Len(cFiltro.Text))) = UCase$(cFiltro.Text) Or LenB(cFiltro.Text) = 0 Then
                LynxSuperficies.AddItem i
                k = LynxSuperficies.Rows - 1
                LynxSuperficies.CellText(k, 1) = SupData(i).Grh
                LynxSuperficies.CellText(k, 2) = vDatos
                Exit For
            End If
            
        Next
        
    Next i
    
    LynxSuperficies.Visible = True
    LynxSuperficies.Redraw = True
    LynxSuperficies.ColForceFit
    
    DoEvents

End Sub

Public Function DameGrhIndex(ByVal GrhIn As Long) As Long
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

    DameGrhIndex = SupData(GrhIn).Grh
    
    If SupData(GrhIn).Width > 0 Then
        frmConfigSup.MOSAICO.value = vbChecked
        frmConfigSup.mAncho.Text = SupData(GrhIn).Width
        frmConfigSup.mLargo.Text = SupData(GrhIn).Height
        
    Else
        frmConfigSup.MOSAICO.value = vbUnchecked
        frmConfigSup.mAncho.Text = "0"
        frmConfigSup.mLargo.Text = "0"
        
    End If

End Function
