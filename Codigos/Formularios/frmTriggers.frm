VERSION 5.00
Begin VB.Form frmTriggers 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Triggers"
   ClientHeight    =   5190
   ClientLeft      =   20160
   ClientTop       =   6450
   ClientWidth     =   4440
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   Begin WinterMapEditor.lvButtons_H cInsertarTrigger 
      Height          =   435
      Left            =   300
      TabIndex        =   1
      Top             =   4170
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   767
      Caption         =   "Insertar Trigger"
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
   Begin WinterMapEditor.LynxGrid LynxTriggers 
      Height          =   3975
      Left            =   60
      TabIndex        =   0
      Top             =   60
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
   Begin WinterMapEditor.lvButtons_H cQuitarTrigger 
      Height          =   435
      Left            =   300
      TabIndex        =   2
      Top             =   4680
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   767
      Caption         =   "Quitar Trigger"
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
End
Attribute VB_Name = "frmTriggers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cInsertarTrigger_Click()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'*************************************************
    If cInsertarTrigger.value = True Then
        cQuitarTrigger.Enabled = False
    Else
        cQuitarTrigger.Enabled = True
    End If
End Sub

Private Sub cQuitarTrigger_Click()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'*************************************************
    If cQuitarTrigger.value = True Then
        LynxTriggers.Enabled = False
        cInsertarTrigger.Enabled = False
        
    Else
        LynxTriggers.Enabled = True
        cInsertarTrigger.Enabled = True
        
    End If
End Sub
