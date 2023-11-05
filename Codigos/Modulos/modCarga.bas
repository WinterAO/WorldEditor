Attribute VB_Name = "modCarga"
Option Explicit

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Enum E_SISTEMA_MUSICA
    CONST_DESHABILITADA = 0
    CONST_MP3 = 1
    CONST_MIDI = 2
End Enum

Public Type tSetupMods

    ' VIDEO
    byMemory    As Integer
    LimiteFPS As Boolean
    OverrideVertexProcess As Byte
    TilesBuffer As Byte
    
    ' AUDIO
    bMusic    As E_SISTEMA_MUSICA
    bSound    As Byte
    bAmbient As Byte
    Invertido As Byte
    MusicVolume As Long
    SoundVolume As Long
    AmbientVol As Long
    
    'MOSTRAR
    Preview As Boolean
    
    'CONFIGURACION
    MeMode As Byte
End Type

Public ClientSetup As tSetupMods

'Path
Public IniPath As String
Public DirRecursos As String
Public DirDats As String

'Recuento de indices
Public grhCount    As Long
Public MaxSup      As Integer
Public NumNPCs     As Integer
Public NumObjs     As Integer

'Constantes
Public Const INITDIR As String = "Init\"

Private Lector As clsIniManager

Public Function profilesFile() As String
    profilesFile = IniPath & INITDIR & "profiles.ini"
End Function

Public Function profileFile(ByVal tag As String) As String
    profileFile = IniPath & INITDIR & "profile-" & tag & ".ini"
End Function

Private Function autoCompletaPath(ByVal Path As String) As String
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'Descripcion: Completa y corrije un path
'*************************************************

    Path = Replace(Path, "/", "\")
    
    If Left(Path, 1) = "\" Then
        ' agrego app.path & path
        Path = App.Path & Path
    End If
    If Right(Path, 1) <> "\" Then
        ' me aseguro que el final sea con "\"
        Path = Path & "\"
    End If
    autoCompletaPath = Path
    
End Function

Public Sub IniciarCabecera()

    With MiCabecera
        .Desc = "WinterAO Resurrection mod Argentum Online by Noland Studios. http://winterao.com.ar"
        .CRC = Rnd * 245
        .MagicWord = Rnd * 92
    End With
    
End Sub

Public Function guardarPerfil() As Boolean

    On Local Error GoTo fileErr:
    
    If Not FileExist(profileFile(ProfileTag), vbArchive) Then
        Exit Function
    End If
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(profileFile(ProfileTag))
    
    Call Lector.ChangeValue("MOSTRAR", "Capa1", IIf(VerCapa1, "1", "0"))
    Call Lector.ChangeValue("MOSTRAR", "Capa2", IIf(VerCapa2, "1", "0"))
    Call Lector.ChangeValue("MOSTRAR", "Capa3", IIf(VerCapa3, "1", "0"))
    Call Lector.ChangeValue("MOSTRAR", "Capa4", IIf(VerCapa4, "1", "0"))
    
    Call Lector.ChangeValue("MOSTRAR", "Translados", IIf(VerTranslados, "1", "0"))
    Call Lector.ChangeValue("MOSTRAR", "Objetos", IIf(VerObjetos, "1", "0"))
    Call Lector.ChangeValue("MOSTRAR", "NPCs", IIf(VerNpcs, "1", "0"))
    Call Lector.ChangeValue("MOSTRAR", "Triggers", IIf(VerTriggers, "1", "0"))
    Call Lector.ChangeValue("MOSTRAR", "Grilla", IIf(VerGrilla, "1", "0"))
    Call Lector.ChangeValue("MOSTRAR", "Particulas", IIf(VerParticulas, "1", "0"))
    Call Lector.ChangeValue("MOSTRAR", "Bloqueos", IIf(VerBlockeados, "1", "0"))
    
    Call Lector.DumpFile(profileFile(ProfileTag))
    
    Exit Function
    
fileErr:

    If Err.Number <> 0 Then
        MsgBox ("Ha ocurrido un error al guardar la configuracion del editor. Error " & Err.Number & " : " & Err.Description)
    End If
End Function

Public Function guardarPerfilMinimap() As Boolean
    '*************************************************
    'Author: Lorwik
    'Last modified: 29/09/2023
    '*************************************************
    
    On Local Error GoTo fileErr:
    
    If Not FileExist(profileFile(ProfileTag), vbArchive) Then
        Exit Function
    End If
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(profileFile(ProfileTag))
    
    Call Lector.ChangeValue("MINIMAPA", "Capa1", IIf(frmMain.Minimap(0).Checked, "1", "0"))
    Call Lector.ChangeValue("MINIMAPA", "Capa2", IIf(frmMain.Minimap(1).Checked, "1", "0"))
    Call Lector.ChangeValue("MINIMAPA", "Capa3", IIf(frmMain.Minimap(2).Checked, "1", "0"))
    Call Lector.ChangeValue("MINIMAPA", "Capa4", IIf(frmMain.Minimap(3).Checked, "1", "0"))
    Call Lector.ChangeValue("MINIMAPA", "NPC", IIf(frmMain.Minimap(4).Checked, "1", "0"))
    Call Lector.ChangeValue("MINIMAPA", "Obj", IIf(frmMain.Minimap(5).Checked, "1", "0"))
    Call Lector.ChangeValue("MINIMAPA", "Bloqueos", IIf(frmMain.Minimap(6).Checked, "1", "0"))
    Call Lector.ChangeValue("MINIMAPA", "Particulas", IIf(frmMain.Minimap(7).Checked, "1", "0"))
    Call Lector.ChangeValue("MINIMAPA", "Nombre", IIf(frmMain.Minimap(8).Checked, "1", "0"))
    Call Lector.ChangeValue("MINIMAPA", "Cuadrantes", IIf(frmMain.Minimap(9).Checked, "1", "0"))
    Call Lector.ChangeValue("MINIMAPA", "Zonas", IIf(frmMain.Minimap(10).Checked, "1", "0"))
    
    Call Lector.DumpFile(profileFile(ProfileTag))
    
    Exit Function
    
fileErr:

    If Err.Number <> 0 Then
        MsgBox ("Ha ocurrido un error al guardar la configuracion del editor. Error " & Err.Number & " : " & Err.Description)
    End If
End Function

Public Function leerPerfil() As Boolean
'**********************************
'Autor: Lorwik
'Fecha: 26/04/2021
'Descripcion: Lee la configuracion de WorldEditor
'**********************************

On Local Error GoTo fileErr:
    
    Dim Profile   As clsIniManager
    Set Profile = New clsIniManager
    
    Dim i As Byte
    Dim NewPath As String
    
    If Not FileExist(profileFile(ProfileTag), vbArchive) Then
        MsgBox "¡No se ha encontrado el archivo de perfil (" & profileFile(ProfileTag) & ") en la carpeta init!", vbOKOnly Or vbExclamation, App.Title
        End
    End If
    
    Call Profile.Initialize(profileFile(ProfileTag))

    With ClientSetup

        .MeMode = Val(Profile.GetValue("CONFIGURACION", "MeMode"))
        
        ' VIDEO
        .LimiteFPS = CBool(Val(Profile.GetValue("VIDEO", "LimitarFPS")))
        .TilesBuffer = Val(Profile.GetValue("VIDEO", "TilesBuffer"))
        .byMemory = Val(Profile.GetValue("VIDEO", "DynamicMemory"))
        .OverrideVertexProcess = Val(Profile.GetValue("VIDEO", "VertexProcessingOverride"))

        '****
        'RUTAS
        '*****************
        
        '-------------------
        'Recursos
        DirRecursos = autoCompletaPath(Profile.GetValue("PATH-" & .MeMode, "DirRecursos"))
        
        If FileExist(DirRecursos, vbDirectory) = False Or DirRecursos = "\" Then
            MsgBox "El directorio de Recursos es incorrecto", vbCritical + vbOKOnly
            
            NewPath = Buscar_Carpeta("Seleccione la carpeta de los recursos de graficos, scripts y fuentes", "")
            Call WriteVar(profileFile(ProfileTag), "PATH-" & .MeMode, "DirRecursos", NewPath)
            DirRecursos = NewPath & "\"
        End If
        
        If FileExist(DirRecursos & "Graficos" & Formato, vbArchive) = False Then
            MsgBox "No se encontro el recursos de graficos en " & DirRecursos & "Graficos" & Formato & "."
            End
        End If
        
        If FileExist(DirRecursos & "Scripts" & Formato, vbArchive) = False Then
            MsgBox "No se encontro el recursos de Scripts."
            End
        End If
        
        If FileExist(DirRecursos & "Fuentes" & Formato, vbArchive) = False Then
            MsgBox "No se encontro el recursos de Fuentes."
            End
        End If
        
        '-------------------
        'Dats
        DirDats = autoCompletaPath(Profile.GetValue("PATH-" & .MeMode, "DirDats"))
        
        If FileExist(DirDats, vbDirectory) = False Or DirDats = "\" Then
            MsgBox "El directorio de Dats es incorrecto", vbCritical + vbOKOnly
            
            NewPath = Buscar_Carpeta("Seleccione la carpeta de los Dats", "")
            Call WriteVar(profileFile(ProfileTag), "PATH-" & .MeMode, "DirDats", NewPath)
            DirDats = NewPath & "\"
        End If
        
        If FileExist(DirDats & "Obj.dat", vbArchive) = False Then
            MsgBox "No se encontro el archivo Obj.dat."
            End
        End If
        
        If FileExist(DirDats & "NPcs.dat", vbArchive) = False Then
            MsgBox "No se encontro el archivo NPCs.dat."
            End
        End If
        
        ' Tamaño de visualizacion
        PantallaX = Val(Profile.GetValue("MOSTRAR", "PantallaX"))
        PantallaY = Val(Profile.GetValue("MOSTRAR", "PantallaY"))
        If PantallaX > 23 Or PantallaX <= 2 Then PantallaX = 23
        If PantallaY > 32 Or PantallaY <= 2 Then PantallaY = 32
        
        ' [GS] 02/10/06
        ' Tamaño de visualizacion en el cliente
        ClienteHeight = Val(Profile.GetValue("MOSTRAR", "ClienteHeight"))
        ClienteWidth = Val(Profile.GetValue("MOSTRAR", "ClienteWidth"))
        If ClienteHeight <= 0 Then ClienteHeight = 13
        If ClienteWidth <= 0 Then ClienteWidth = 17
        
        ' Menu Mostrar
        VerCapa1 = Val(Profile.GetValue("MOSTRAR", "Capa1"))
        VerCapa2 = Val(Profile.GetValue("MOSTRAR", "Capa2"))
        VerCapa3 = Val(Profile.GetValue("MOSTRAR", "Capa3"))
        VerCapa4 = Val(Profile.GetValue("MOSTRAR", "Capa4"))
        VerTranslados = Val(Profile.GetValue("MOSTRAR", "Translados"))
        VerObjetos = Val(Profile.GetValue("MOSTRAR", "Objetos"))
        VerNpcs = Val(Profile.GetValue("MOSTRAR", "NPCs"))
        VerTriggers = Val(Profile.GetValue("MOSTRAR", "Triggers"))
        VerGrilla = Val(Profile.GetValue("MOSTRAR", "Grilla")) ' Grilla
        VerParticulas = Val(Profile.GetValue("MOSTRAR", "Particulas"))
        VerBlockeados = Val(Profile.GetValue("MOSTRAR", "Bloqueos"))
        
        frmMain.Minimap(0).Checked = Val(Profile.GetValue("MINIMAPA", "Capa1"))
        frmMain.Minimap(1).Checked = Val(Profile.GetValue("MINIMAPA", "Capa2"))
        frmMain.Minimap(2).Checked = Val(Profile.GetValue("MINIMAPA", "Capa3"))
        frmMain.Minimap(3).Checked = Val(Profile.GetValue("MINIMAPA", "Capa4"))
        frmMain.Minimap(4).Checked = Val(Profile.GetValue("MINIMAPA", "NPC"))
        frmMain.Minimap(5).Checked = Val(Profile.GetValue("MINIMAPA", "Obj"))
        frmMain.Minimap(6).Checked = Val(Profile.GetValue("MINIMAPA", "Bloqueos"))
        frmMain.Minimap(7).Checked = Val(Profile.GetValue("MINIMAPA", "Particulas"))
        frmMain.Minimap(8).Checked = Val(Profile.GetValue("MINIMAPA", "Nombre"))
        frmMain.Minimap(9).Checked = Val(Profile.GetValue("MINIMAPA", "Cuadrantes"))
        frmMain.Minimap(10).Checked = Val(Profile.GetValue("MINIMAPA", "Zonas"))
        
        MMiniMap_capa1 = frmMain.Minimap(0).Checked
        MMiniMap_capa2 = frmMain.Minimap(1).Checked
        MMiniMap_capa3 = frmMain.Minimap(2).Checked
        MMiniMap_capa4 = frmMain.Minimap(3).Checked
        MMiniMap_Npcs = frmMain.Minimap(4).Checked
        MMiniMap_objetos = frmMain.Minimap(5).Checked
        MMiniMap_particulas = frmMain.Minimap(6).Checked
        MMiniMap_Bloqueos = frmMain.Minimap(7).Checked
        MMiniMap_Nombre = frmMain.Minimap(8).Checked
        
        ' AUDIO
        .bMusic = CByte(Val(Profile.GetValue("AUDIO", "MUSICA")))
        .bSound = CByte(Val(Profile.GetValue("AUDIO", "SONIDO")))
        .bAmbient = CByte(Val(Profile.GetValue("AUDIO", "AMBIENT")))
        .MusicVolume = CLng(Val(Profile.GetValue("AUDIO", "VOLMUSICA")))
        .SoundVolume = CLng(Val(Profile.GetValue("AUDIO", "VOLAUDIO")))
        .AmbientVol = CLng(Val(Profile.GetValue("AUDIO", "VOLAMBIENT")))
        
    End With

    Set Profile = Nothing

    leerPerfil = True

  Exit Function
  
fileErr:

    leerPerfil = False

    If Err.Number <> 0 Then
       MsgBox ("Ha ocurrido un error al cargar la configuracion de perfil del cliente. Error " & Err.Number & " : " & Err.Description)
       End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    End If

End Function

'-----------------------------------------------------------
'Carga de Indices
'-----------------------------------------------------------

Public Sub LoadGrhData()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Graficos
'*************************************
On Error GoTo ErrorHandler:

    Dim Grh         As Long
    Dim Frame       As Long
    Dim fileVersion As Long
    Dim fileBuff    As clsByteBuffer
    Dim InfoHead    As INFOHEADER
    Dim buffer()    As Byte

    InfoHead = File_Find(DirRecursos & "Scripts" & Formato, LCase$("graficos.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("graficos.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
    
        fileVersion = fileBuff.getLong
        
        grhCount = fileBuff.getLong
        
        ReDim GrhData(0 To grhCount) As GrhData
        
        While Grh <> grhCount
            Grh = fileBuff.getLong

            With GrhData(Grh)
            
                '.active = True
                .NumFrames = fileBuff.getInteger
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To .NumFrames)
                
                If .NumFrames > 1 Then
                
                    For Frame = 1 To .NumFrames
                        .Frames(Frame) = fileBuff.getLong
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then GoTo ErrorHandler
                    Next Frame
                    
                    .speed = fileBuff.getSingle
                    If .speed <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                    
                Else
                    
                    .FileNum = fileBuff.getLong
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = fileBuff.getInteger
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = fileBuff.getInteger
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .sX = fileBuff.getInteger
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    .sY = fileBuff.getInteger
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                    
                End If
                
            End With
            
        Wend
        
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
Exit Sub

ErrorHandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Graficos.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor
        End If
        
    End If
    
End Sub

Public Sub CargarMinimapa()

    Dim fileBuff    As clsByteBuffer
    Dim InfoHead    As INFOHEADER
    Dim buffer()    As Byte
    Dim i           As Long
    
    InfoHead = File_Find(DirRecursos & "Scripts" & Formato, LCase$("minimap.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("minimap.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        For i = 1 To grhCount
            If Grh_Check(i) Then
                GrhData(i).mini_map_color = fileBuff.getLong
            End If
        Next i
        
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
End Sub

Public Sub CargarCabezas()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Cabezas
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim InfoHead    As INFOHEADER
    Dim i           As Integer
    Dim NumHeads As Integer
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(DirRecursos & "Scripts" & modCompression.Formato, LCase$("Head.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Head.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        NumHeads = fileBuff.getInteger()  'cantidad de cabezas
    
        ReDim heads(0 To NumHeads) As tHead
                
        For i = 1 To NumHeads
            heads(i).Std = fileBuff.getByte()
            heads(i).Texture = fileBuff.getInteger()
            heads(i).startX = fileBuff.getInteger()
            heads(i).startY = fileBuff.getInteger()
        Next i
        
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Head.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor
        End If
        
    End If
    
End Sub

Public Sub CargarCascos()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Cascos
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i           As Integer
    Dim NumCascos As Integer
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(DirRecursos & "Scripts" & modCompression.Formato, LCase$("Helmet.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Helmet.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
    
        NumCascos = fileBuff.getInteger()   'cantidad de cascos
             
        ReDim Cascos(0 To NumCascos) As tHead
             
        For i = 1 To NumCascos
            Cascos(i).Std = fileBuff.getByte()
            Cascos(i).Texture = fileBuff.getInteger()
            Cascos(i).startX = fileBuff.getInteger()
            Cascos(i).startY = fileBuff.getInteger()
        Next i
         
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Helmet.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor
        End If
        
    End If
    
End Sub

Sub CargarCuerpos()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Cuerpos
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i           As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(DirRecursos & "Scripts" & modCompression.Formato, LCase$("Personajes.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Personajes.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
    
        'num de cabezas
        NumCuerpos = fileBuff.getInteger()
    
        'Resize array
        ReDim BodyData(0 To NumCuerpos) As BodyData
        ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
        For i = 1 To NumCuerpos
            MisCuerpos(i).Body(1) = fileBuff.getLong()
            MisCuerpos(i).Body(2) = fileBuff.getLong()
            MisCuerpos(i).Body(3) = fileBuff.getLong()
            MisCuerpos(i).Body(4) = fileBuff.getLong()
            MisCuerpos(i).HeadOffsetX = fileBuff.getInteger()
            MisCuerpos(i).HeadOffsetY = fileBuff.getInteger()
            
            If MisCuerpos(i).Body(1) Then
                Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
                Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
                Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
                Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
                
                BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
                BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
            End If
        Next i
    
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Personajes.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor
        End If
        
    End If
    
End Sub

Sub CargarAnimArmas()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Armas
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i As Long
    Dim NumWeaponAnims As Integer
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(DirRecursos & "Scripts" & modCompression.Formato, LCase$("Armas.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Armas.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
    
        'num de armas
        NumWeaponAnims = fileBuff.getInteger()
        
        'Resize array
        ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
        ReDim Weapons(1 To NumWeaponAnims) As tIndiceArmas
        
        For i = 1 To NumWeaponAnims
            Weapons(i).weapon(1) = fileBuff.getLong()
            Weapons(i).weapon(2) = fileBuff.getLong()
            Weapons(i).weapon(3) = fileBuff.getLong()
            Weapons(i).weapon(4) = fileBuff.getLong()
            
            If Weapons(i).weapon(1) Then
            
                Call InitGrh(WeaponAnimData(i).WeaponWalk(1), Weapons(i).weapon(1), 0)
                Call InitGrh(WeaponAnimData(i).WeaponWalk(2), Weapons(i).weapon(2), 0)
                Call InitGrh(WeaponAnimData(i).WeaponWalk(3), Weapons(i).weapon(3), 0)
                Call InitGrh(WeaponAnimData(i).WeaponWalk(4), Weapons(i).weapon(4), 0)
            
            End If
        Next i
    
        Erase buffer
    End If
    
    Set fileBuff = Nothing

errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Armas.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor
        End If
        
    End If

End Sub

Sub CargarAnimEscudos()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Escudos
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim InfoHead    As INFOHEADER
    Dim i As Long
    Dim NumEscudosAnims As Integer
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(DirRecursos & "Scripts" & modCompression.Formato, LCase$("Escudos.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Escudos.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
    
        'num de escudos
        NumEscudosAnims = fileBuff.getInteger()
        
        'Resize array
        ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
        ReDim Shields(1 To NumEscudosAnims) As tIndiceEscudos
        
        For i = 1 To NumEscudosAnims
            Shields(i).shield(1) = fileBuff.getLong()
            Shields(i).shield(2) = fileBuff.getLong()
            Shields(i).shield(3) = fileBuff.getLong()
            Shields(i).shield(4) = fileBuff.getLong()
            
            If Shields(i).shield(1) Then
            
                Call InitGrh(ShieldAnimData(i).ShieldWalk(1), Shields(i).shield(1), 0)
                Call InitGrh(ShieldAnimData(i).ShieldWalk(2), Shields(i).shield(2), 0)
                Call InitGrh(ShieldAnimData(i).ShieldWalk(3), Shields(i).shield(3), 0)
                Call InitGrh(ShieldAnimData(i).ShieldWalk(4), Shields(i).shield(4), 0)
            
            End If
        Next i
    
        Erase buffer
    End If
    
    Set fileBuff = Nothing

errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Escudos.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor
        End If
        
    End If
    
End Sub


Private Function Grh_Check(ByVal grh_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= grhCount Then
        Grh_Check = GrhData(grh_index).NumFrames
    End If
End Function

Public Sub CargarIndicesSuperficie()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'Descripcion: Carga los indices de Superficie
'*************************************************

On Error GoTo Fallo
    Dim Leer As New clsIniManager
    Dim i As Integer
    Dim K As Long
    
    If FileExist(IniPath & INITDIR & "indices.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'indices.ini'", vbCritical
        End
    End If
    
    Leer.Initialize IniPath & INITDIR & "indices.ini"
    MaxSup = Leer.GetValue("INIT", "Referencias")
    
    ReDim SupData(MaxSup) As SupData
    
    frmSuperficies.LynxSuperficies.Clear
    frmSuperficies.LynxSuperficies.Redraw = False
    frmSuperficies.LynxSuperficies.Visible = False

    frmSuperficies.LynxSuperficies.AddColumn "Indice", 0
    frmSuperficies.LynxSuperficies.AddColumn "Grh", 0
    frmSuperficies.LynxSuperficies.AddColumn "Nombre", 3
    
    For i = 0 To MaxSup
        SupData(i).name = Leer.GetValue("REFERENCIA" & i, "Nombre")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
        
        frmSuperficies.LynxSuperficies.AddItem i
        K = frmSuperficies.LynxSuperficies.Rows - 1
        frmSuperficies.LynxSuperficies.CellText(K, 1) = SupData(i).Grh
        frmSuperficies.LynxSuperficies.CellText(K, 2) = SupData(i).name
    Next
    
    frmSuperficies.LynxSuperficies.Visible = True
    frmSuperficies.LynxSuperficies.Redraw = True
    frmSuperficies.LynxSuperficies.ColForceFit
    
    DoEvents
    
    Set Leer = Nothing
    
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de \indices.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
    
End Sub

Public Sub CargarIndicesNPC()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'Descripcion: Carga los indices de NPCs
'*************************************************
On Error Resume Next
'On Error GoTo Fallo

    If FileExist(DirDats & "NPCs.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & DirDats, vbCritical
        Call CloseMapEditor
    End If

    Dim Trabajando As String
    Dim NPC As Long
    Dim Hostil As String
    Dim Leer As New clsIniManager
    Dim K As Long
    
    Call Leer.Initialize(DirDats & "NPCs.dat")
    NumNPCs = Val(Leer.GetValue("INIT", "NumNPCs"))
    
    ReDim NpcData(NumNPCs) As NpcData
    Trabajando = "Dats\NPCs.dat"
    
    frmNPCs.LynxNPCs.Clear
    frmNPCs.LynxNPCs.Redraw = False
    frmNPCs.LynxNPCs.Visible = False
    
    frmNPCs.LynxNPCs.AddColumn "Num", 0
    frmNPCs.LynxNPCs.AddColumn "Nombre", 2
    frmNPCs.LynxNPCs.AddColumn "Nivel", 0
    frmNPCs.LynxNPCs.AddColumn "Hostil", 0
    
    For NPC = 1 To NumNPCs
        With NpcData(NPC)
            .name = CStr(Leer.GetValue("NPC" & NPC, "Name"))
            .ELV = Val(Leer.GetValue("NPC" & NPC, "ELV"))
            .Hostile = Val(Leer.GetValue("NPC" & NPC, "Hostile"))
            .NpcType = Val(Leer.GetValue("NPC" & NPC, "NPCType"))
            
            .Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
            .Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
            .WeaponAnim = Val(Leer.GetValue("NPC" & NPC, "WeaponAnim"))
            .CascoAnim = Val(Leer.GetValue("NPC" & NPC, "CascoAnim"))
            .ShieldAnim = Val(Leer.GetValue("NPC" & NPC, "ShieldAnim"))
            .Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))
            
            frmNPCs.LynxNPCs.AddItem NPC
            
            K = frmNPCs.LynxNPCs.Rows - 1
            frmNPCs.LynxNPCs.CellText(K, 1) = .name
            frmNPCs.LynxNPCs.CellText(K, 2) = .ELV
            
            Hostil = IIf(.Hostile = 1, "SI", "NO")
            frmNPCs.LynxNPCs.CellText(K, 3) = Hostil
            
        End With
    Next
    
    frmNPCs.LynxNPCs.Visible = True
    frmNPCs.LynxNPCs.Redraw = True
    frmNPCs.LynxNPCs.ColForceFit
    
    DoEvents
    
    Set Leer = Nothing
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & " de " & Trabajando & " en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

Public Sub CargarIndicesOBJ()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'*************************************************

On Error GoTo Fallo

    Dim Obj As Integer
    Dim K As Long
    Dim Leer As New clsIniManager

    If FileExist(DirDats & "\OBJ.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'OBJ.dat' en " & DirDats, vbCritical
        End
    End If

    Call Leer.Initialize(DirDats & "\OBJ.dat")
    
    frmOBJs.LynxOBJs.Clear
    frmOBJs.LynxOBJs.Redraw = False
    frmOBJs.LynxOBJs.Visible = False
    
    NumObjs = Val(Leer.GetValue("INIT", "NumOBJs"))
    ReDim ObjData(1 To NumObjs) As ObjData
    
    frmOBJs.LynxOBJs.AddColumn "Num", 0
    frmOBJs.LynxOBJs.AddColumn "Nombre", 2
    
    For Obj = 1 To NumObjs
        frmCarga.lblStatus.Caption = "Cargando Datos de Objetos..." & Obj & "/" & NumObjs
        DoEvents
        
        With ObjData(Obj)
    
            .name = Leer.GetValue("OBJ" & Obj, "Name")
            .GrhIndex = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
            .ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
            .Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
            .Info = Leer.GetValue("OBJ" & Obj, "Info")
            .WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
            .Texto = Leer.GetValue("OBJ" & Obj, "Texto")
            .GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))
            .Cerrada = Val(Leer.GetValue("OBJ" & Obj, "Cerrada"))
            .Subtipo = Val(Leer.GetValue("OBJ" & Obj, "Subtipo"))
            
            frmOBJs.LynxOBJs.AddItem Obj
            K = frmOBJs.LynxOBJs.Rows - 1
            frmOBJs.LynxOBJs.CellText(K, 1) = .name
        
        End With
    Next Obj
    
    frmOBJs.LynxOBJs.Visible = True
    frmOBJs.LynxOBJs.Redraw = True
    frmOBJs.LynxOBJs.ColForceFit
    
    DoEvents
    
    Set Leer = Nothing
    Exit Sub
    
Fallo:
    MsgBox "Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

Public Sub CargarIndicesTriggers()
'*************************************************
'Author: Lorwik
'Last modified: 29/04/2021
' Carga los indices de Triggers
'*************************************************

On Error GoTo Fallo

    Dim K As Long

    If FileExist(IniPath & INITDIR & "Triggers.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Triggers.ini' en " & IniPath & INITDIR & "Triggers.ini", vbCritical
        End
    End If
    
    Dim NumT As Integer
    Dim T As Integer
    Dim Leer As New clsIniManager
    
    Call Leer.Initialize(IniPath & INITDIR & "Triggers.ini")
    
    frmTriggers.LynxTriggers.Clear
    frmTriggers.LynxTriggers.Redraw = False
    frmTriggers.LynxTriggers.Visible = False
    
    frmTriggers.LynxTriggers.AddColumn "Num", 0
    frmTriggers.LynxTriggers.AddColumn "Nombre", 2
    
    NumT = Val(Leer.GetValue("INIT", "NumTriggers"))
    For T = 1 To NumT
        frmTriggers.LynxTriggers.AddItem T
        K = frmTriggers.LynxTriggers.Rows - 1
        frmTriggers.LynxTriggers.CellText(K, 1) = Leer.GetValue("Trig" & T, "Name")
    Next T

    frmTriggers.LynxTriggers.Visible = True
    frmTriggers.LynxTriggers.Redraw = True
    frmTriggers.LynxTriggers.ColForceFit
    
    DoEvents

    Set Leer = Nothing
    Exit Sub
    
Fallo:
    MsgBox "Error al intentar cargar el Trigger " & T & " de Triggers.ini en " & IniPath & INITDIR & "Triggers.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub
