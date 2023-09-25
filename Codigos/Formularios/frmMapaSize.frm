VERSION 5.00
Begin VB.Form frmMapaSize 
   Caption         =   "Tamaño del Mapa"
   ClientHeight    =   3165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9480
   Icon            =   "frmMapaSize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   9480
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameAdapted 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   9255
      Begin VB.TextBox txtYMax 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "32000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Text            =   "1000"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtXMax 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "32000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "1000"
         Top             =   240
         Width           =   1095
      End
      Begin WinterMapEditor.lvButtons_H mapResize 
         Height          =   525
         Index           =   0
         Left            =   3000
         TabIndex        =   7
         Top             =   600
         Width           =   765
         _extentx        =   3254
         _extenty        =   714
         caption         =   ""
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapaSize.frx":49D2
         mode            =   0
         value           =   0
         cback           =   -2147483633
      End
      Begin VB.PictureBox Picture1 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   8
         Top             =   0
         Width           =   0
      End
      Begin WinterMapEditor.lvButtons_H mapResize 
         Height          =   525
         Index           =   1
         Left            =   3840
         TabIndex        =   9
         Top             =   600
         Width           =   765
         _extentx        =   1349
         _extenty        =   926
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapaSize.frx":49FA
         mode            =   0
         value           =   0
         cback           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H mapResize 
         Height          =   525
         Index           =   2
         Left            =   4680
         TabIndex        =   10
         Top             =   600
         Width           =   765
         _extentx        =   1349
         _extenty        =   926
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapaSize.frx":4A22
         mode            =   0
         value           =   0
         cback           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H mapResize 
         Height          =   525
         Index           =   3
         Left            =   3000
         TabIndex        =   11
         Top             =   1200
         Width           =   765
         _extentx        =   1349
         _extenty        =   926
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapaSize.frx":4A4A
         mode            =   0
         value           =   0
         cback           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H mapResize 
         Height          =   525
         Index           =   4
         Left            =   3840
         TabIndex        =   12
         Top             =   1200
         Width           =   765
         _extentx        =   1349
         _extenty        =   926
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapaSize.frx":4A72
         mode            =   0
         value           =   0
         cback           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H mapResize 
         Height          =   525
         Index           =   5
         Left            =   4680
         TabIndex        =   13
         Top             =   1200
         Width           =   765
         _extentx        =   1349
         _extenty        =   926
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapaSize.frx":4A9A
         mode            =   0
         value           =   0
         cback           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H mapResize 
         Height          =   525
         Index           =   6
         Left            =   3000
         TabIndex        =   14
         Top             =   1800
         Width           =   765
         _extentx        =   1349
         _extenty        =   926
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapaSize.frx":4AC2
         mode            =   0
         value           =   0
         cback           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H mapResize 
         Height          =   525
         Index           =   7
         Left            =   3840
         TabIndex        =   15
         Top             =   1800
         Width           =   765
         _extentx        =   1349
         _extenty        =   926
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapaSize.frx":4AEA
         mode            =   0
         value           =   0
         cback           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H mapResize 
         Height          =   525
         Index           =   8
         Left            =   4680
         TabIndex        =   16
         Top             =   1800
         Width           =   765
         _extentx        =   1349
         _extenty        =   926
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapaSize.frx":4B12
         mode            =   0
         value           =   0
         cback           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBAplicar 
         Height          =   525
         Left            =   6960
         TabIndex        =   17
         Top             =   1800
         Width           =   2055
         _extentx        =   3625
         _extenty        =   714
         caption         =   "Guardar"
         capalign        =   2
         backstyle       =   2
         gradient        =   3
         cgradient       =   0
         cfore           =   16777215
         font            =   "frmMapaSize.frx":4B3A
         mode            =   0
         value           =   0   'False
         cfhover         =   16777215
         cback           =   255
         cbhover         =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ancla de desplazamiento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tamaño en Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tamaño en X:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "Tu modo activado solo permite ?x?."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5010
   End
End
Attribute VB_Name = "frmMapaSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private newXMax As Integer
Private newYMax As Integer
Private anchor As Byte  '0 to 8

Private Sub Form_Load()
    frameAdapted.Visible = False
    
    If ClientSetup.MeMode = eMeMode.ImperiumClasico Or _
        ClientSetup.MeMode = eMeMode.WinterUltimate Then
        ' solo 100x100
        lblMsg.Caption = frmModo.OptModo(ClientSetup.MeMode).Caption & " solo permite 100x100."
    ElseIf ClientSetup.MeMode = eMeMode.WinterAO Then
        ' solo 1000x1000
        lblMsg.Caption = frmModo.OptModo(ClientSetup.MeMode).Caption & " solo permite 1000x1000."
    ElseIf ClientSetup.MeMode = eMeMode.ArgentumUnited Then
        lblMsg.Caption = frmModo.OptModo(ClientSetup.MeMode).Caption & " permite tamaño adaptado."
        frameAdapted.Visible = True
        
        txtXMax.Text = XMaxMapSize
        txtYMax.Text = YMaxMapSize
        
        newYMax = XMaxMapSize
        newYMax = YMaxMapSize
        LvBAplicar.Enabled = False
        
        Call mapResize_Click(4)
    End If
    
End Sub

Private Sub LvBAplicar_Click()
    Dim res As VbMsgBoxResult
    res = MsgBox("¿Está seguro de que desea cambiar el tamaño del mapa a " & newXMax & " x " & newYMax & "?", vbQuestion + vbOKCancel)
    If res = vbOK Then
        MsgBox "No tan rapido! esto falta :P"
    End If
End Sub

Private Sub CanApply()
    newXMax = Val(txtXMax.Text)
    newYMax = Val(txtYMax.Text)
    
    If newXMax <> XMaxMapSize Or _
        newYMax <> YMaxMapSize Then
        LvBAplicar.Enabled = True
    Else
        LvBAplicar.Enabled = False
    End If
End Sub

Private Sub mapResize_Click(Index As Integer)
 ' 0 1 2
 ' 3 4 5
 ' 6 7 8

Dim LoopC As Byte
For LoopC = 0 To 8
    If Index <> LoopC Then
        mapResize(LoopC).Caption = ""
    End If
Next

Const eArriba = "^"
Const eAbajo = "v"
Const eIzq = "<"
Const eDer = ">"
Const eCentro = "(*)"

mapResize(Index).Caption = eCentro
anchor = Index

Select Case Index
    Case 0
        mapResize(1).Caption = eDer
        mapResize(3).Caption = eAbajo
    Case 1
        mapResize(0).Caption = eIzq
        mapResize(2).Caption = eDer
        mapResize(4).Caption = eAbajo
    Case 2
        mapResize(1).Caption = eIzq
        mapResize(5).Caption = eAbajo
    Case 3
        mapResize(4).Caption = eDer
        mapResize(0).Caption = eArriba
        mapResize(6).Caption = eAbajo
    Case 4
        mapResize(3).Caption = eIzq
        mapResize(5).Caption = eDer
        mapResize(1).Caption = eArriba
        mapResize(7).Caption = eAbajo
    Case 5
        mapResize(4).Caption = eIzq
        mapResize(2).Caption = eArriba
        mapResize(8).Caption = eAbajo
    Case 6
        mapResize(7).Caption = eDer
        mapResize(3).Caption = eArriba
    Case 7
        mapResize(6).Caption = eIzq
        mapResize(8).Caption = eDer
        mapResize(4).Caption = eArriba
    Case 8
        mapResize(7).Caption = eIzq
        mapResize(5).Caption = eArriba
End Select

End Sub

Private Sub txtXMax_Change()
    Call CanApply
End Sub

Private Sub txtYMax_Change()
    Call CanApply
End Sub
