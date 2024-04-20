VERSION 5.00
Begin VB.Form frmPerfil 
   BackColor       =   &H00424242&
   BorderStyle     =   0  'None
   Caption         =   "Selección de Perfil"
   ClientHeight    =   4065
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
   ScaleHeight     =   271
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00535353&
      Caption         =   "Carga de Recursos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   5115
      Begin VB.OptionButton OptRecursos 
         BackColor       =   &H00535353&
         Caption         =   "Recursos libres"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   330
         Width           =   1875
      End
      Begin VB.OptionButton OptRecursos 
         BackColor       =   &H00535353&
         Caption         =   "Recursos comprimidos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2580
         TabIndex        =   11
         Top             =   330
         Width           =   2295
      End
   End
   Begin VB.Frame FraPerfil 
      BackColor       =   &H00535353&
      Caption         =   "Perfil"
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
      Height          =   795
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   5115
      Begin VB.ComboBox cmbPerfil 
         Height          =   315
         ItemData        =   "frmModo.frx":0000
         Left            =   210
         List            =   "frmModo.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   2895
      End
      Begin WinterMapEditor.lvButtons_H LvBNuevo 
         Height          =   345
         Index           =   2
         Left            =   3990
         TabIndex        =   6
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "Nuevo"
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
         cBack           =   8454016
      End
      Begin WinterMapEditor.lvButtons_H LvBBorrar 
         Height          =   345
         Index           =   0
         Left            =   3180
         TabIndex        =   13
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "Borrar"
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
         cBack           =   8421631
      End
   End
   Begin VB.Frame FraConfiguraciónDe 
      BackColor       =   &H00535353&
      Caption         =   "Configuración de video"
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
      Height          =   1185
      Left            =   120
      TabIndex        =   0
      Top             =   1410
      Width           =   5115
      Begin VB.CheckBox chkvSync 
         BackColor       =   &H00535353&
         Caption         =   "Activar sincronización vertical"
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
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   750
         Width           =   3015
      End
      Begin VB.ComboBox cmbProcesado 
         Height          =   315
         ItemData        =   "frmModo.frx":0004
         Left            =   1920
         List            =   "frmModo.frx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   345
         Width           =   2895
      End
      Begin VB.Label lblModoDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modo de procesado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   1470
      End
   End
   Begin WinterMapEditor.lvButtons_H LvBBoton 
      Height          =   465
      Index           =   1
      Left            =   3000
      TabIndex        =   7
      Top             =   3450
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
      Left            =   210
      TabIndex        =   8
      Top             =   3450
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
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Perfiles() As String
Private nPerfiles As Byte

Private Sub cmbPerfil_Click()
'*****************************************
'Descripción: Lee la configuración básica del perfil seleccionado.
'
'*****************************************

    Dim tag As String
    tag = cmbPerfil.Text
    
    If LenB(tag) > 0 And FileExist(profileFile(tag), vbArchive) Then
        Dim v As Byte
        v = CByte(Val(GetVar(profileFile(tag), "CONFIGURACION", "MeMode")))

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
    
    Dim i As Byte
    Dim lastProfile As Byte
        
    'Leemos cuantos perfiles hay y cual fue el ultimo que se utilizo
    nPerfiles = Val(GetVar(profilesFile, "INIT", "profiles"))
    lastProfile = Val(GetVar(profilesFile, "INIT", "lastProfile"))
    
    cmbPerfil.Clear
    
    'Leemos todos los perfiles y los añadimos al comboBox
    For i = 1 To nPerfiles
        Dim tag As String
        tag = GetVar(profilesFile, "PROFILE" & i, "name")
        If LenB(tag) > 0 And FileExist(profileFile(tag), vbArchive) Then
            cmbPerfil.AddItem (tag)
        End If
    Next i
    
    ReDim Perfiles(1 To nPerfiles) As String
    
    If lastProfile <= nPerfiles And lastProfile > 0 Then
        cmbPerfil.ListIndex = lastProfile - 1
    End If
    
    
    'Configuración inicial:
    
    cmbProcesado.ListIndex = GetVar(profileFile(cmbPerfil.List(cmbPerfil.ListIndex)), "VIDEO", "VertexProcessingOverride")
    
    If ClientSetup.LimiteFPS Then
        chkvSync.value = Checked
    Else
        chkvSync.value = Unchecked
    End If
    
    ClientSetup.useCompression = CBool(GetVar(profileFile(cmbPerfil.List(cmbPerfil.ListIndex)), "CONFIGURACION", "useCompression"))
    
    If Not ClientSetup.useCompression Then
        OptRecursos(0).value = True
    Else
        OptRecursos(1).value = True
    End If
    
End Sub

Private Sub chkvSync_Click()
    If chkvSync.value = Checked Then
        ClientSetup.LimiteFPS = True
        
    Else
        ClientSetup.LimiteFPS = False
        
    End If
End Sub

Private Sub LvBBorrar_Click(Index As Integer)

    Dim confirmacion As String

    confirmacion = InputBox("¡Cuidado! ¡Estás apunto de eliminar el perfil " & cmbPerfil.List(cmbPerfil.ListIndex) & "! ¿Estás seguro de que quieres borrarlo? Se perdera toda la configuración de este perfil. Escribe el nombre del perfil para confirmar.")
    
    If confirmacion = cmbPerfil.List(cmbPerfil.ListIndex) Then
        MsgBox "Siento decirte que esto aun no esta programado."
        Exit Sub
    Else
        MsgBox "El nombre que escribiste no coincide con el del perfil. El perfil no se eliminara."
        Exit Sub
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
            
            'Lorwik> Esto por ahora no se usa y no se sabe si se volvera a usar.
            ModoElegido = True
            'Call WriteVar(profileFile(ProfileTag), "CONFIGURACION", "MeMode", CStr(ClientSetup.MeMode))
            ' Guarda el índice del perfil seleccionado en "lastProfile"
            
            ClientSetup.OverrideVertexProcess = cmbProcesado.ListIndex
            Call WriteVar(profileFile(ProfileTag), "VIDEO", "VertexProcessingOverride", CByte(ClientSetup.OverrideVertexProcess))
            Call WriteVar(profileFile(ProfileTag), "VIDEO", "LimitarFPS", IIf(ClientSetup.LimiteFPS, "1", "0"))
            
            Call WriteVar(profileFile(ProfileTag), "CONFIGURACION", "useCompression", CByte(ClientSetup.useCompression))
            
            Call WriteVar(profilesFile, "INIT", "lastProfile", cmbPerfil.ListIndex + 1)
            
            Unload Me
    End Select
End Sub

Private Sub LvBNuevo_Click(Index As Integer)
'*****************************************
'Descripción: Crea un nuevo perfil
'
'*****************************************

    ProfileTag = InputBox("Introduce el nombre para el perfil.")
    
    If ProfileTag = vbNullString Then Exit Sub
    
    'Añadimos el nuevo perfil al archivo profile
    Call WriteVar(profilesFile, "INIT", "profiles", nPerfiles + 1)
    Call WriteVar(profilesFile, "PROFILE" & (nPerfiles + 1), "name", ProfileTag)
    
    'Creamos un nuevo archivo profile para el nuevo perfil
    Call WriteVar(profileFile(ProfileTag), "MOSTRAR", "Capa1", "1")
    Call WriteVar(profileFile(ProfileTag), "MOSTRAR", "Capa2", "1")
    Call WriteVar(profileFile(ProfileTag), "MOSTRAR", "Capa3", "1")
    Call WriteVar(profileFile(ProfileTag), "MOSTRAR", "Capa4", "1")

    ReDim Perfiles(1 To nPerfiles) As String
    
    cmbPerfil.AddItem (ProfileTag)
    
End Sub

Private Sub OptRecursos_Click(Index As Integer)

    ClientSetup.useCompression = IIf(OptRecursos(1).value, True, False)

End Sub
