VERSION 5.00
Begin VB.Form frmModo 
   BackColor       =   &H00424242&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
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
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraConfiguraciónDe 
      BackColor       =   &H00535353&
      Caption         =   "Configuración de video"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   5175
      Begin VB.CheckBox chkvSync 
         BackColor       =   &H00535353&
         Caption         =   "Activar sincronización vertical"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   930
         Width           =   3015
      End
      Begin VB.ComboBox cmbProcesado 
         Height          =   315
         ItemData        =   "frmModo.frx":0000
         Left            =   1920
         List            =   "frmModo.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   400
         Width           =   2895
      End
      Begin VB.Label lblModoDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modo de procesado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1470
      End
   End
   Begin VB.Frame FraModo 
      BackColor       =   &H00535353&
      Caption         =   "¿En que modo quieres iniciar el WorldEditor?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin WinterMapEditor.lvButtons_H LvBBoton 
         Height          =   345
         Index           =   1
         Left            =   2310
         TabIndex        =   9
         Top             =   990
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   609
         Caption         =   "Aceptar"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WinterMapEditor.lvButtons_H LvBBoton 
         Height          =   345
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   960
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   609
         Caption         =   "Salir"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.OptionButton OptModo 
         BackColor       =   &H00535353&
         Caption         =   "Modo Winter"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton OptModo 
         BackColor       =   &H00535353&
         Caption         =   "Modo Imperium Classic"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   1
         Top             =   480
         Width           =   2415
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
   End
End
Attribute VB_Name = "frmModo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public ModoElegido As Boolean

Private Sub Form_Load()
    On Error Resume Next
    
    #If Privado = 0 Then
        'Marcamos la opcion
        OptModo(ClientSetup.MeMode).value = True
        
    #Else
        OptModo(eMeMode.WinterAO).value = True
        OptModo(eMeMode.ImperiumClasico).Visible = False
    #End If
    
    cmbProcesado.ListIndex = ClientSetup.OverrideVertexProcess
    
    If ClientSetup.LimiteFPS Then
        chkvSync.value = Checked
        
    Else
        chkvSync.value = Unchecked
        
    End If
End Sub

Private Sub chkvSync_Click()
    If chkvSync.value = Checked Then
        ClientSetup.LimiteFPS = True
        
    Else
        ClientSetup.LimiteFPS = False
        
    End If
End Sub

Private Sub LvBBoton_Click(Index As Integer)
    Select Case Index
    
        Case 0 'Salir
            Call SimpleLogError("Seleccion de modo cancelador, saliendo de WorldEditor.")
            End
        
        Case 1
            ModoElegido = True
            
            ClientSetup.OverrideVertexProcess = cmbProcesado.ListIndex
            
            Call WriteVar(configFile, "CONFIGURACION", "MeMode", CStr(ClientSetup.MeMode))
            Call WriteVar(configFile, "VIDEO", "VertexProcessingOverride", CByte(ClientSetup.OverrideVertexProcess))
            Call WriteVar(configFile, "VIDEO", "LimitarFPS", IIf(ClientSetup.LimiteFPS, "1", "0"))
            
            #If Privado = 0 Then
                Call SimpleLogError("Modo " & OptModo(ClientSetup.MeMode).Caption & " elegido.")
            #End If
            
            Unload Me
    End Select
End Sub

Private Sub OptModo_Click(Index As Integer)

    #If Privado = 1 Then
        ClientSetup.MeMode = eMeMode.WinterAO
        
    #Else
        ClientSetup.MeMode = Index
    #End If
    
End Sub
