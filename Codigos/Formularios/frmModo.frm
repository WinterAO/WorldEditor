VERSION 5.00
Begin VB.Form frmModo 
   BackColor       =   &H00424242&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
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
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraPerfil 
      BackColor       =   &H00535353&
      Caption         =   "Perfil"
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   120
      TabIndex        =   9
      Top             =   540
      Width           =   5115
      Begin VB.ComboBox cmbPerfil 
         Height          =   315
         ItemData        =   "frmModo.frx":0000
         Left            =   210
         List            =   "frmModo.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   300
         Width           =   2895
      End
      Begin WinterMapEditor.lvButtons_H LvBNuevo 
         Height          =   345
         Index           =   2
         Left            =   3270
         TabIndex        =   11
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         Caption         =   "Nuevo"
         CapAlign        =   2
         BackStyle       =   2
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
   End
   Begin VB.Frame FraConfiguraciónDe 
      BackColor       =   &H00535353&
      Caption         =   "Configuración de video"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   5115
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
         ItemData        =   "frmModo.frx":0004
         Left            =   1920
         List            =   "frmModo.frx":0014
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5115
      Begin VB.OptionButton OptModo 
         BackColor       =   &H00535353&
         Caption         =   "Modo Argentum World"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   620
         TabIndex        =   12
         Top             =   630
         Width           =   1965
      End
      Begin VB.OptionButton OptModo 
         BackColor       =   &H00535353&
         Caption         =   "Modo Winter Ultimate"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   2850
         TabIndex        =   8
         Top             =   660
         Width           =   1965
      End
      Begin VB.OptionButton OptModo 
         BackColor       =   &H00535353&
         Caption         =   "Modo Winter"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   2
         Top             =   360
         Width           =   1275
      End
      Begin VB.OptionButton OptModo 
         BackColor       =   &H00535353&
         Caption         =   "Modo Imperium Clasico"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2850
         TabIndex        =   1
         Top             =   360
         Width           =   1965
      End
      Begin VB.PictureBox Picture1 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   7
         Top             =   0
         Width           =   0
      End
   End
   Begin WinterMapEditor.lvButtons_H LvBBoton 
      Height          =   465
      Index           =   1
      Left            =   3000
      TabIndex        =   13
      Top             =   4080
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   820
      Caption         =   "&Continuar"
      CapAlign        =   2
      BackStyle       =   2
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
      cBack           =   8454016
   End
   Begin WinterMapEditor.lvButtons_H LvBBoton 
      Height          =   465
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   820
      Caption         =   "&Salir"
      CapAlign        =   2
      BackStyle       =   2
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
      cBack           =   8421631
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selección de perfil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmModo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Perfiles() As String
Private nPerfiles As Byte

Private Sub cmbPerfil_Click()
    Dim tag As String
    tag = cmbPerfil.Text
    If LenB(tag) > 0 And FileExist(profileFile(tag), vbArchive) Then
        Dim v As Byte
        v = CByte(Val(GetVar(profileFile(tag), "CONFIGURACION", "MeMode")))
        OptModo(v).value = True

        v = CByte(Val(GetVar(profileFile(tag), "VIDEO", "VertexProcessingOverride")))
        cmbProcesado.ListIndex = v
    
        v = CByte(Val(GetVar(profileFile(tag), "VIDEO", "LimitarFPS")))
    
        If (v = 1) Then
            chkvSync.value = Checked
            
        Else
            chkvSync.value = Unchecked
            
        End If
        
        
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    OptModo(3).value = True 'default
    
    Dim i As Byte
        
    nPerfiles = Val(GetVar(profilesFile, "INIT", "profiles"))
    
    cmbPerfil.Clear
    
    For i = 1 To nPerfiles
        Dim tag As String
        tag = GetVar(profilesFile, "PROFILE" & i, "name")
        If LenB(tag) > 0 And FileExist(profileFile(tag), vbArchive) Then
            cmbPerfil.AddItem (tag)
        End If
    Next i
    
    ReDim Perfiles(1 To nPerfiles) As String
    
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
        
            If cmbPerfil.ListIndex < 0 Then
                MsgBox "¡No se ha seleccionado ningún perfil!", vbCritical
                Exit Sub
            End If
        
            ProfileTag = cmbPerfil.List(cmbPerfil.ListIndex)
        
            ModoElegido = True
            
            ClientSetup.OverrideVertexProcess = cmbProcesado.ListIndex
            
            Call WriteVar(profileFile(ProfileTag), "CONFIGURACION", "MeMode", CStr(ClientSetup.MeMode))
            Call WriteVar(profileFile(ProfileTag), "VIDEO", "VertexProcessingOverride", CByte(ClientSetup.OverrideVertexProcess))
            Call WriteVar(profileFile(ProfileTag), "VIDEO", "LimitarFPS", IIf(ClientSetup.LimiteFPS, "1", "0"))
            
            #If Privado = 0 Then
                Call SimpleLogError("Modo " & OptModo(ClientSetup.MeMode).Caption & " elegido.")
            #End If
            
            Unload Me
    End Select
End Sub

Private Sub LvBNuevo_Click(Index As Integer)
    ProfileTag = InputBox("Introduce el nombre para el perfil.")
    
    If ProfileTag = vbNullString Then Exit Sub
    
    Call WriteVar(profilesFile, "INIT", "profiles", nPerfiles + 1)
    Call WriteVar(profilesFile, "PROFILE" & (nPerfiles + 1), "name", ProfileTag)
    
    Call WriteVar(profileFile(ProfileTag), "MOSTRAR", "Capa1", "1")
    Call WriteVar(profileFile(ProfileTag), "MOSTRAR", "Capa2", "1")
    Call WriteVar(profileFile(ProfileTag), "MOSTRAR", "Capa3", "1")
    Call WriteVar(profileFile(ProfileTag), "MOSTRAR", "Capa4", "1")

    ReDim Perfiles(0 To nPerfiles) As String
    
    cmbPerfil.AddItem (ProfileTag)
End Sub

Private Sub OptModo_Click(Index As Integer)

    #If Privado = 1 Then
        ClientSetup.MeMode = eMeMode.ArgentumUnited
        
    #Else
        ClientSetup.MeMode = Index
    #End If
    
End Sub
