VERSION 5.00
Begin VB.Form frmConfigSup 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración Acanzada de Superficie"
   ClientHeight    =   1110
   ClientLeft      =   13860
   ClientTop       =   6855
   ClientWidth     =   4320
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
   ScaleHeight     =   74
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   288
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox DespMosaic 
      Appearance      =   0  'Flat
      BackColor       =   &H00424242&
      Caption         =   "Desplz. de Mosaico"
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
      Height          =   210
      Left            =   2280
      TabIndex        =   11
      Top             =   480
      UseMaskColor    =   -1  'True
      Value           =   1  'Checked
      Width           =   1920
   End
   Begin VB.CheckBox MOSAICO 
      Appearance      =   0  'Flat
      BackColor       =   &H00424242&
      Caption         =   "Mosaico"
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
      Height          =   165
      Left            =   2280
      TabIndex        =   10
      Top             =   240
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox mLargo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1320
      TabIndex        =   9
      Text            =   "4"
      Top             =   600
      Width           =   420
   End
   Begin VB.TextBox mAncho 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1320
      TabIndex        =   8
      Text            =   "4"
      Top             =   240
      Width           =   420
   End
   Begin VB.TextBox DMAncho 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1800
      TabIndex        =   7
      Text            =   "0"
      Top             =   240
      Width           =   420
   End
   Begin VB.TextBox DMLargo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1800
      TabIndex        =   6
      Text            =   "0"
      Top             =   600
      Width           =   420
   End
   Begin VB.CommandButton cmdDM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   360
      Picture         =   "frmConfigSup.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   240
   End
   Begin VB.CommandButton cmdDM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      Picture         =   "frmConfigSup.frx":02F2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   240
   End
   Begin VB.CommandButton cmdDM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   600
      Picture         =   "frmConfigSup.frx":05E2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   240
   End
   Begin VB.CommandButton cmdDM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   360
      Picture         =   "frmConfigSup.frx":08D1
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   240
   End
   Begin VB.CommandButton cmdDM 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   240
   End
   Begin VB.CheckBox chkActivarPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00424242&
      Caption         =   "Activar Preview"
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
      Height          =   195
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   1725
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
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
      Height          =   195
      Left            =   1080
      TabIndex        =   13
      Top             =   600
      Width           =   105
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   12
      Top             =   240
      Width           =   105
   End
End
Attribute VB_Name = "frmConfigSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkActivarPreview_Click()
    If chkActivarPreview.value = Checked Then
        ClientSetup.Preview = True
        
    Else
        ClientSetup.Preview = False
        
    End If
    
    If Len(ProfileTag) > 0 Then
        Call WriteVar(profileFile(ProfileTag), "MOSTRAR", "Preview", IIf(ClientSetup.Preview = True, "1", "0"))
    End If
End Sub

Private Sub cmdDM_Click(Index As Integer)

    On Error GoTo cmdDM_Click_Err
    
    DespMosaic.value = vbChecked

    Select Case Index

        Case 0 'A
    
            DMLargo.Text = Val(DMLargo.Text) + 1

        Case 1 '<
            DMAncho.Text = Val(DMAncho.Text) + 1

        Case 2 '>
            DMAncho.Text = Val(DMAncho.Text) - 1

        Case 3 'V
            DMLargo.Text = Val(DMLargo.Text) - 1

        Case 4 '0
            DMAncho.Text = 0
            DMLargo.Text = 0

    End Select

    Exit Sub

cmdDM_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmConfigSup.cmdDM_Click", Erl)
    Resume Next
End Sub

Private Sub Form_Load()
    If ClientSetup.Preview = True Then
        frmConfigSup.chkActivarPreview.value = Unchecked
        
    Else
        frmConfigSup.chkActivarPreview.value = Checked
        
    End If
End Sub
