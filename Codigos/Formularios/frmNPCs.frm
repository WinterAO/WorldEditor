VERSION 5.00
Begin VB.Form frmNPCs 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NPC's"
   ClientHeight    =   5925
   ClientLeft      =   9600
   ClientTop       =   6450
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   395
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   ShowInTaskbar   =   0   'False
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
      ItemData        =   "frmNPCs.frx":0000
      Left            =   840
      List            =   "frmNPCs.frx":0002
      TabIndex        =   6
      Text            =   "1"
      Top             =   4560
      Width           =   1215
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
      ItemData        =   "frmNPCs.frx":0004
      Left            =   3480
      List            =   "frmNPCs.frx":0006
      TabIndex        =   5
      Text            =   "1"
      Top             =   4560
      Width           =   855
   End
   Begin VB.ComboBox cFiltro 
      BackColor       =   &H80000012&
      ForeColor       =   &H80000014&
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Top             =   4200
      Width           =   3765
   End
   Begin WinterMapEditor.lvButtons_H cAgregarFuncalAzar 
      Height          =   405
      Left            =   180
      TabIndex        =   2
      Top             =   5070
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   714
      Caption         =   "Insertar NPC's al Azar"
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
      Left            =   2430
      TabIndex        =   1
      Top             =   5070
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1402
      Caption         =   "Insertar NPC's"
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
   Begin WinterMapEditor.LynxGrid LynxNPCs 
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
   Begin WinterMapEditor.lvButtons_H cQuitarFunc 
      Height          =   405
      Left            =   180
      TabIndex        =   9
      Top             =   5490
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   714
      Caption         =   "Quitar NPC's"
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
   Begin VB.Label lCantFunc 
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
      Left            =   120
      TabIndex        =   8
      Top             =   4635
      Width           =   675
   End
   Begin VB.Label lNumFunc 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de NPC:"
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
      Left            =   2250
      TabIndex        =   7
      Top             =   4635
      Width           =   1170
   End
   Begin VB.Label lbFiltrar 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar:"
      ForeColor       =   &H80000014&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   4230
      Width           =   570
   End
End
Attribute VB_Name = "frmNPCs"
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

Private Sub LynxNPCs_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2021
    '*************************************************
        
    Call CargarInfo
    
End Sub

Private Sub LynxNPCs_LostFocus()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    HotKeysAllow = True
End Sub

Private Sub LynxNPCs_KeyDown(KeyCode As Integer, Shift As Integer)
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
        LynxNPCs.Enabled = False
        
    Else
        cInsertarFunc.Enabled = True
        cAgregarFuncalAzar.Enabled = True
        cCantFunc.Enabled = True
        cNumFunc.Enabled = True
        cFiltro.Enabled = True
        LynxNPCs.Enabled = True
        
    End If
End Sub

Private Sub CargarInfo()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    HotKeysAllow = False
        
    'Obtiene el numero del Grh
    cNumFunc.Text = LynxNPCs.CellText(, 0)
 
End Sub

Private Sub Filtrar()
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'*************************************************

    Dim vDatos As String
    Dim NumI As Integer
    Dim i As Integer
    Dim j As Integer
    Dim K As Long
    
    If cFiltro.ListCount > 5 Then _
        cFiltro.RemoveItem 0
    
    cFiltro.AddItem cFiltro.Text
    LynxNPCs.Clear
    LynxNPCs.Redraw = False
    LynxNPCs.Visible = False
    
    For i = 0 To NumNPCs
        vDatos = NpcData(i).name
        NumI = i
        
        For j = 1 To Len(vDatos)
            If UCase$(mid$(vDatos & str(i), j, Len(cFiltro.Text))) = UCase$(cFiltro.Text) Or LenB(cFiltro.Text) = 0 Then
                LynxNPCs.AddItem NumI
                K = LynxNPCs.Rows - 1
                LynxNPCs.CellText(K, 1) = vDatos
                LynxNPCs.CellText(K, 2) = NpcData(NumI).ELV
                LynxNPCs.CellText(K, 3) = IIf(NpcData(NumI).Hostile = 1, "SI", "NO")
                Exit For
            End If
        Next
        
    Next i
    
    LynxNPCs.Visible = True
    LynxNPCs.Redraw = True
    LynxNPCs.ColForceFit
    
    DoEvents

End Sub

