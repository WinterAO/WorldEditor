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
    CampoVision As Boolean
    useCompression As Boolean
    GuardadoAuto As Boolean
    IntervaloGuardado As Byte
    
    'AUTO CAPTURAR
    AutoCapturarSuperficies As Boolean
    AutoCapturarTraslados As Boolean
    AutoCapturarNPCs As Boolean
    AutoCapturarObjs As Boolean
    AutoCapturarParticulas As Boolean
    
End Type

Public ClientSetup As tSetupMods

'Path
Public IniPath As String
Public dirRecursos_Compressed As String
Public dirRecursos_Uncompressed As String
Public dirDats As String

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
    
    With ClientSetup
    
        Call Lector.ChangeValue("CONFIGURACION", "GuardadoAuto", IIf(ClientSetup.GuardadoAuto, "1", "0"))
        Call Lector.ChangeValue("CONFIGURACION", "IntervalGuardado", ClientSetup.IntervaloGuardado)
        
        Call Lector.ChangeValue("AUTO-CAPTURAR", "Superficies", IIf(.AutoCapturarSuperficies, "1", "0"))
        Call Lector.ChangeValue("AUTO-CAPTURAR", "Traslados", IIf(.AutoCapturarTraslados, "1", "0"))
        Call Lector.ChangeValue("AUTO-CAPTURAR", "NPCs", IIf(.AutoCapturarNPCs, "1", "0"))
        Call Lector.ChangeValue("AUTO-CAPTURAR", "Objetos", IIf(.AutoCapturarObjs, "1", "0"))
        Call Lector.ChangeValue("AUTO-CAPTURAR", "Particulas", IIf(.AutoCapturarParticulas, "1", "0"))
    
    End With
    
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

Public Function guardarPerfilVideo() As Boolean

    On Local Error GoTo fileErr:
    
    If Not FileExist(profileFile(ProfileTag), vbArchive) Then
        Exit Function
    End If
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(profileFile(ProfileTag))
    
    Call Lector.ChangeValue("VIDEO", "VertexProcessingOverride", CByte(ClientSetup.OverrideVertexProcess))
    Call Lector.ChangeValue("VIDEO", "LimitarFPS", IIf(ClientSetup.LimiteFPS, "1", "0"))
    Call Lector.ChangeValue("VIDEO", "TilesBuffer", Val(ClientSetup.TilesBuffer))
    Call Lector.ChangeValue("VIDEO", "DynamicMemory", Val(ClientSetup.byMemory))
    
    Call Lector.DumpFile(profileFile(ProfileTag))
    
    Exit Function
    
fileErr:

    If Err.Number <> 0 Then
        MsgBox ("Ha ocurrido un error al guardar la configuracion de video del editor. Error " & Err.Number & " : " & Err.Description)
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

        'CONFIGURACION GENERAL
        .MeMode = Val(Profile.GetValue("CONFIGURACION", "MeMode"))
        .GuardadoAuto = Val(Profile.GetValue("CONFIGURACION", "GuardadoAuto"))
        .IntervaloGuardado = Val(Profile.GetValue("CONFIGURACION", "IntervalGuardado"))
        
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
        If ClientSetup.useCompression Then
            dirRecursos_Compressed = autoCompletaPath(Profile.GetValue("PATH-" & .MeMode, "dirRecursos_Compressed"))
            
            If FileExist(dirRecursos_Compressed, vbDirectory) = False Or dirRecursos_Compressed = "\" Then
                MsgBox "El directorio de recursos comprimidos es incorrecto", vbCritical + vbOKOnly
                
                NewPath = Buscar_Carpeta("Seleccione el directorio donde se encuentran los recursos comprimidos de gráficos, scripts y fuentes", "")
                Call WriteVar(profileFile(ProfileTag), "PATH-" & .MeMode, "dirRecursos_Compressed", NewPath)
                dirRecursos_Compressed = NewPath & "\"
            End If
            
            If FileExist(dirRecursos_Compressed & "Graficos" & Formato, vbArchive) = False Then
                MsgBox "No se encontro el recurso de graficos en " & dirRecursos_Compressed & "Graficos" & Formato & "."
                End
            End If
            
            If FileExist(dirRecursos_Compressed & "Scripts" & Formato, vbArchive) = False Then
                MsgBox "No se encontro el recurso de Scripts."
                End
            End If
            
            If FileExist(dirRecursos_Compressed & "Fuentes" & Formato, vbArchive) = False Then
                MsgBox "No se encontro el recurso de Fuentes."
                End
            End If
            
        Else
            dirRecursos_Uncompressed = autoCompletaPath(Profile.GetValue("PATH-" & .MeMode, "dirRecursos_Uncompressed"))
            
            If FileExist(dirRecursos_Uncompressed, vbDirectory) = False Or dirRecursos_Uncompressed = "\" Then
                MsgBox "El directorio de recursos libres es incorrecto", vbCritical + vbOKOnly
                
                NewPath = Buscar_Carpeta("Seleccione el directorio donde se encuentran las carpetas de los recursos libres de gráficos, scripts y fuentes", "")
                Call WriteVar(profileFile(ProfileTag), "PATH-" & .MeMode, "dirRecursos_Uncompressed", NewPath)
                dirRecursos_Uncompressed = NewPath & "\"
            End If
            
            If FileExist(dirRecursos_Uncompressed & "\Graficos", vbDirectory) = False Then
                MsgBox "No se encontro el directorio de graficos en " & dirRecursos_Uncompressed & "Graficos" & Formato & "."
                End
            End If
            
            If FileExist(dirRecursos_Uncompressed & "\Scripts", vbDirectory) = False Then
                MsgBox "No se encontro el directorio de Scripts."
                End
            End If
            
            If FileExist(dirRecursos_Uncompressed & "\Fuentes", vbDirectory) = False Then
                MsgBox "NNo se encontro el directorio de Fuentes."
                End
            End If
            
        End If
        
        '-------------------
        'Dats
        dirDats = autoCompletaPath(Profile.GetValue("PATH-" & .MeMode, "DirDats"))
        
        If FileExist(dirDats, vbDirectory) = False Or dirDats = "\" Then
            MsgBox "El directorio de Dats es incorrecto", vbCritical + vbOKOnly
            
            NewPath = Buscar_Carpeta("Seleccione la carpeta de los Dats", "")
            Call WriteVar(profileFile(ProfileTag), "PATH-" & .MeMode, "DirDats", NewPath)
            dirDats = NewPath & "\"
        End If
        
        If FileExist(dirDats & "Obj.dat", vbArchive) = False Then
            MsgBox "No se encontro el archivo Obj.dat."
            End
        End If
        
        If FileExist(dirDats & "NPcs.dat", vbArchive) = False Then
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
        
        ' AUTO-CAPTURA
        .AutoCapturarSuperficies = CBool(Val(Profile.GetValue("AUTO-CAPTURAR", "Superficies")))
        .AutoCapturarTraslados = CBool(Val(Profile.GetValue("AUTO-CAPTURAR", "Traslados")))
        .AutoCapturarNPCs = CBool(Val(Profile.GetValue("AUTO-CAPTURAR", "NPCs")))
        .AutoCapturarObjs = CBool(Val(Profile.GetValue("AUTO-CAPTURAR", "Objetos")))
        .AutoCapturarParticulas = CBool(Val(Profile.GetValue("AUTO-CAPTURAR", "Particulas")))
        
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
'Fecha: 20/04/2024
'*************************************

    If ClientSetup.useCompression Then
        Call LoadGrhData_Compressed
        
    Else
        Call LoadGrhData_Uncompressed
        
    End If
    
End Sub

Private Function LoadGrhData_Compressed() As Boolean
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Graficos del archivo comprimido
'*************************************
On Error GoTo ErrorHandler:

    Dim Grh         As Long
    Dim k           As Long
    Dim frame       As Long
    Dim fileVersion As Long
    Dim fileBuff    As clsByteBuffer
    Dim InfoHead    As INFOHEADER
    Dim buffer()    As Byte

    InfoHead = File_Find(dirRecursos_Compressed & "Scripts" & Formato, LCase$("graficos.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        With frmGrh.LynxGrh
        
            .Clear
            .Redraw = False
            .Visible = False
            .AddColumn "Grh", 0
            .AddColumn "Tipo", 0
        
        End With
    
        Extract_File_Memory Scripts, LCase$("graficos.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
    
        fileVersion = fileBuff.getLong
        
        grhCount = fileBuff.getLong
        
        ReDim GrhData(0 To grhCount) As GrhData
        
        While Grh <> grhCount
        
            Grh = fileBuff.getLong
            
            frmGrh.LynxGrh.AddItem Grh
            k = frmGrh.LynxGrh.Rows - 1
            frmGrh.LynxGrh.CellText(k, 1) = Grh

            With GrhData(Grh)
            
                '.active = True
                .NumFrames = fileBuff.getInteger
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To .NumFrames)
                
                If .NumFrames > 1 Then
                
                    frmGrh.LynxGrh.CellText(k, 1) = "ANIMACION"
                
                    For frame = 1 To .NumFrames
                        .Frames(frame) = fileBuff.getLong
                        If .Frames(frame) <= 0 Or .Frames(frame) > grhCount Then GoTo ErrorHandler
                    Next frame
                    
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
                
                    frmGrh.LynxGrh.CellText(k, 1) = ""
                    
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
    
    With frmGrh.LynxGrh
        .Visible = True
        .Redraw = True
        .ColForceFit
    End With
    
    LoadGrhData_Compressed = True
    
Exit Function

ErrorHandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Graficos.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor
        End If
        
    End If
    
    With frmGrh.LynxGrh
        .Visible = True
        .Redraw = True
        .ColForceFit
    End With
    
    Erase buffer
    
    LoadGrhData_Compressed = False
    
End Function

Public Function LoadGrhData_Uncompressed() As Boolean
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Graficos
'*************************************

    On Error GoTo ErrorHandler
    
    Dim k           As Long

    Dim Grh         As Long

    Dim fileVersion As Long

    Dim frame       As Long

    Dim handle      As Integer
    
    If Not FileExist(dirRecursos_Uncompressed & "\Scripts\graficos.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Graficos.ind."
        LoadGrhData_Uncompressed = False
        Exit Function

    End If
    
    'Open files
    handle = FreeFile()
    Open dirRecursos_Uncompressed & "\Scripts\Graficos.ind" For Binary Access Read As handle
    
    With frmGrh.LynxGrh
        
        .Clear
        .Redraw = False
        .Visible = False
        .AddColumn "Grh", 0
        .AddColumn "Tipo", 0
        
    End With
    
    Get handle, , fileVersion
        
    Get handle, , grhCount

    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData

    While Grh <> grhCount
    
        Get handle, , Grh
        
        frmGrh.LynxGrh.AddItem Grh
        k = frmGrh.LynxGrh.Rows - 1
        frmGrh.LynxGrh.CellText(k, 1) = Grh
    
        With GrhData(Grh)

            If Grh <> 0 Then
            
                Grh = Grh
               
                'Get number of frames
                Get handle, , .NumFrames

                If .NumFrames <= 0 Then GoTo ErrorHandler
            
                ReDim .Frames(1 To .NumFrames)
            
                If .NumFrames > 1 Then

                    frmGrh.LynxGrh.CellText(k, 1) = "ANIMACION"

                    For frame = 1 To .NumFrames
                        Get handle, , .Frames(frame)

                        If .Frames(frame) <= 0 Or .Frames(frame) > grhCount Then GoTo ErrorHandler
                    Next frame
                
                    Get handle, , .speed

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
                
                    frmGrh.LynxGrh.CellText(k, 1) = ""
                    
                    'Read in normal GRH data
                    Get handle, , .FileNum

                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelWidth

                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelHeight

                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , GrhData(Grh).sX

                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .sY

                    If .sY < 0 Then GoTo ErrorHandler
                
                    'Compute width and height
                    .TileWidth = .pixelWidth / 32
                    .TileHeight = .pixelHeight / 32
                
                    .Frames(1) = Grh

                End If

            End If

        End With
        
    Wend
    
    Close handle
    
    With frmGrh.LynxGrh
        .Visible = True
        .Redraw = True
        .ColForceFit

    End With
    
    DoEvents
    
    LoadGrhData_Uncompressed = True
    
    Exit Function

ErrorHandler:
    Close handle
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Graficos.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor

        End If
        
    End If
    
    With frmGrh.LynxGrh
        .Visible = True
        .Redraw = True
        .ColForceFit

    End With
    
    LoadGrhData_Uncompressed = False

End Function

Public Sub CargarMinimapa()
'*************************************
'Autor: Lorwik
'Fecha: 20/04/2024
'*************************************

    If ClientSetup.useCompression Then
        Call CargarMinimapa_Compressed
    Else
        Call CargarMinimapa_Uncompressed
    End If

End Sub

Private Function CargarMinimapa_Compressed() As Boolean
'*************************************
'Autor: Lorwik
'Fecha: ????
'*************************************

    On Error GoTo ErrorHandler
    
    Dim fileBuff    As clsByteBuffer
    Dim InfoHead    As INFOHEADER
    Dim buffer()    As Byte
    Dim i           As Long
    
    InfoHead = File_Find(dirRecursos_Compressed & "Scripts" & Formato, LCase$("minimap.bin"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("minimap.bin"), buffer()
        
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
    
    CargarMinimapa_Compressed = True
    
    Exit Function
    
ErrorHandler:

    CargarMinimapa_Compressed = False
    
End Function

Private Function CargarMinimapa_Uncompressed() As Boolean
    '*************************************
    'Autor: Lorwik
    'Fecha: ????
    '*************************************

    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    Dim N As Integer
    
    If Not FileExist(dirRecursos_Uncompressed & "\Scripts\minimap.bin", vbArchive) Then
        MsgBox "No se ha encontrado el archivo minimap.bin."
        CargarMinimapa_Uncompressed = False
        Exit Function

    End If
    
    N = FreeFile
    Open dirRecursos_Uncompressed & "\Scripts\minimap.bin" For Binary Access Read As #N
        
    For i = 1 To grhCount

        If Grh_Check(i) Then
            Get #N, , GrhData(i).mini_map_color

        End If
        
    Next i

    Close #N
    
    CargarMinimapa_Uncompressed = True
    
    Exit Function
    
ErrorHandler:

    CargarMinimapa_Uncompressed = False
    
End Function

Public Sub CargarCabezas()
'*************************************
'Autor: Lorwik
'Fecha: 20/04/2024
'*************************************

    If ClientSetup.useCompression Then
        Call CargarCabezas_Compressed
    Else
        Call CargarCabezas_Uncompressed
    End If
End Sub

Private Function CargarCabezas_Compressed() As Boolean
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Cabezas
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim InfoHead    As INFOHEADER
    Dim i           As Integer
    Dim NumHeads    As Integer
    Dim fileBuff    As clsByteBuffer
    
    InfoHead = File_Find(dirRecursos_Compressed & "Scripts" & modCompression.Formato, LCase$("Head.ind"))
    
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
    CargarCabezas_Compressed = True
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Head.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor
        End If
        
    End If
    
    CargarCabezas_Compressed = False
    
End Function

Private Function CargarCabezas_Uncompressed() As Boolean
    On Error GoTo ErrorHandler:
    
    Dim N            As Integer

    Dim i            As Integer
    
    Dim NumHeads     As Integer
    
    If Not FileExist(dirRecursos_Uncompressed & "\Scripts\head.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo head.ind."
        CargarCabezas_Uncompressed = False
        Exit Function
    End If

    N = FreeFile
    Open dirRecursos_Uncompressed & "\Scripts\head.ind" For Binary Access Read As #N

    'num de cabezas
    Get #N, , NumHeads

    'Resize array
    ReDim heads(0 To NumHeads) As tHead
            
        For i = 1 To NumHeads
            Get #N, , heads(i).Std
            Get #N, , heads(i).Texture
            Get #N, , heads(i).startX
            Get #N, , heads(i).startY
            
        Next i

    Close #N

    CargarCabezas_Uncompressed = True

    Exit Function

ErrorHandler:
    Close #N
    'MsgBox "Error " & Err.Number & " durante la carga de Head.ind!"
    CargarCabezas_Uncompressed = False
    Resume
    
End Function

Public Sub CargarCascos()
'*************************************
'Autor: Lorwik
'Fecha: 20/04/2024
'*************************************

    If ClientSetup.useCompression Then
        Call CargarCascos_Compressed
    Else
        Call CargarCascos_Uncompressed
    End If
End Sub

Private Function CargarCascos_Compressed() As Boolean
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
    
    InfoHead = File_Find(dirRecursos_Compressed & "Scripts" & modCompression.Formato, LCase$("Helmet.ind"))
    
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
    CargarCascos_Compressed = True
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Helmet.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor
        End If
        
    End If
    
    CargarCascos_Compressed = False
    
End Function

Public Function CargarCascos_Uncompressed() As Boolean
    On Error GoTo ErrorHandler:
    
    Dim N          As Integer

    Dim i          As Integer
    
    Dim NumCascos  As Integer

    If Not FileExist(dirRecursos_Uncompressed & "\Scripts\helmet.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo helmet.ind."
        CargarCascos_Uncompressed = False
        Exit Function
    End If

    N = FreeFile
    Open dirRecursos_Uncompressed & "\Scripts\helmet.ind" For Binary Access Read As #N

    'num de cascos
    Get #N, , NumCascos

    'Resize array
    ReDim Cascos(0 To NumCascos) As tHead
    
    For i = 1 To NumCascos
        Get #N, , Cascos(i).Std
        Get #N, , Cascos(i).Texture
        Get #N, , Cascos(i).startX
        Get #N, , Cascos(i).startY
            
    Next i
         
    Close #N

    CargarCascos_Uncompressed = True

    Exit Function

ErrorHandler:
    Close #N
    'MsgBox "Error " & Err.Number & " durante la carga de Helmet.ind!"
    CargarCascos_Uncompressed = False
    Resume
    
End Function

Public Sub CargarCuerpos()
'*************************************
'Autor: Lorwik
'Fecha: 20/04/2024
'*************************************

    If ClientSetup.useCompression Then
        Call CargarCuerpos_Compressed
    Else
        Call CargarCuerpos_Uncompressed
    End If
End Sub

Private Function CargarCuerpos_Compressed() As Boolean
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
    
    InfoHead = File_Find(dirRecursos_Compressed & "Scripts" & modCompression.Formato, LCase$("Personajes.ind"))
    
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
                
                BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
                BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
            End If
        Next i
    
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    CargarCuerpos_Compressed = True
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Personajes.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor
        End If
        
    End If
    
    CargarCuerpos_Compressed = False
    
End Function

Public Function CargarCuerpos_Uncompressed() As Boolean

    On Error GoTo ErrorHandler

    Dim N            As Integer

    Dim i            As Long
    
    Dim NumCuerpos   As Integer

    Dim MisCuerpos() As tIndiceCuerpo
    
    If Not FileExist(dirRecursos_Uncompressed & "\Scripts\personajes.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Personajes.ind."
        CargarCuerpos_Uncompressed = False
        Exit Function
    End If

    N = FreeFile()
    Open dirRecursos_Uncompressed & "\Scripts\Personajes.ind" For Binary Access Read As #N

    'num de cabezas
    Get #N, , NumCuerpos

    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo

    For i = 1 To NumCuerpos

        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
            Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
            Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
            Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
                
            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY

        End If
        
    Next i

    Close #N
    
    CargarCuerpos_Uncompressed = True
    
    Exit Function

ErrorHandler:
    Close #N
    'MsgBox "Error " & Err.Number & " durante la carga de Personajes.ind!"
    CargarCuerpos_Uncompressed = False
    Resume
    
End Function

Public Sub CargarAnimArmas()
'*************************************
'Autor: Lorwik
'Fecha: 20/04/2024
'*************************************

    If ClientSetup.useCompression Then
        CargarAnimArmas_Compressed
    Else
        CargarAnimArmas_Uncompressed
    End If

End Sub

Private Function CargarAnimArmas_Compressed()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Armas del archivo comprimido
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i As Long
    Dim NumWeaponAnims As Integer
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(dirRecursos_Compressed & "Scripts" & modCompression.Formato, LCase$("Armas.ind"))
    
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
    
    CargarAnimArmas_Compressed = True

errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Armas.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor
        End If
        
    End If
    
    CargarAnimArmas_Compressed = False

End Function

Private Function CargarAnimArmas_Uncompressed() As Boolean
    '*************************************
    'Autor: Lorwik
    'Fecha: ???
    'Descripción: Carga el index de Armas
    '*************************************
    
    On Error GoTo errhandler:

    Dim N              As Integer

    Dim i              As Long

    Dim NumWeaponAnims As Integer
    
    If Not FileExist(dirRecursos_Uncompressed & "\Scripts\armas.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Armas.ind."
        CargarAnimArmas_Uncompressed = False
        Exit Function

    End If
    
    N = FreeFile
    Open dirRecursos_Uncompressed & "\Scripts\Armas.ind" For Binary Access Read As #N
    
    'num de armas
    Get #N, , NumWeaponAnims
        
    'Resize array
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    ReDim Weapons(1 To NumWeaponAnims) As tIndiceArmas
        
    For i = 1 To NumWeaponAnims
        Get #N, , Weapons(i)
            
        If Weapons(i).weapon(1) Then
            
            Call InitGrh(WeaponAnimData(i).WeaponWalk(1), Weapons(i).weapon(1), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(2), Weapons(i).weapon(2), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(3), Weapons(i).weapon(3), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(4), Weapons(i).weapon(4), 0)

        End If

    Next i
    
    Close #N
    
    CargarAnimArmas_Uncompressed = True
    
    Exit Function

errhandler:
    Close #N
    'MsgBox "Error " & Err.Number & " durante la carga de Armas.ind!"
    CargarAnimArmas_Uncompressed = False
    Resume
    
End Function

Public Sub CargarAnimEscudos()
'*************************************
'Autor: Lorwik
'Fecha: 20/04/2024
'*************************************

    If ClientSetup.useCompression Then
        Call CargarAnimEscudos_Compressed
    Else
        Call CargarAnimEscudos_Uncompressed
    End If

End Sub

Private Function CargarAnimEscudos_Compressed()
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
    
    InfoHead = File_Find(dirRecursos_Compressed & "Scripts" & modCompression.Formato, LCase$("Escudos.ind"))
    
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
    CargarAnimEscudos_Compressed = True

errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Escudos.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseMapEditor
        End If
        
    End If
    
    CargarAnimEscudos_Compressed = False
    
End Function

Public Function CargarAnimEscudos_Uncompressed() As Boolean
    '*************************************
    'Autor: Lorwik
    'Fecha: ???
    'Descripción: Carga el index de Escudos
    '*************************************
    
    On Error GoTo errhandler:

    Dim N               As Integer

    Dim i               As Long

    Dim NumEscudosAnims As Integer

    If Not FileExist(dirRecursos_Uncompressed & "\Scripts\escudos.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Escudos.ind."
        CargarAnimEscudos_Uncompressed = False
        Exit Function

    End If

    N = FreeFile
    Open dirRecursos_Uncompressed & "\Scripts\escudos.ind" For Binary Access Read As #N

    'num de escudos
    Get #N, , NumEscudosAnims
        
    'Resize array
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    ReDim Shields(1 To NumEscudosAnims) As tIndiceEscudos
        
    For i = 1 To NumEscudosAnims
        Get #N, , Shields(i)
            
        If Shields(i).shield(1) Then
            
            Call InitGrh(ShieldAnimData(i).ShieldWalk(1), Shields(i).shield(1), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(2), Shields(i).shield(2), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(3), Shields(i).shield(3), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(4), Shields(i).shield(4), 0)
            
        End If
            
    Next i
    
    Close #N
    
    CargarAnimEscudos_Uncompressed = True
    
    Exit Function

errhandler:
    Close #N
    'MsgBox "Error " & Err.Number & " durante la carga de Escudos.ind!"
    CargarAnimEscudos_Uncompressed = False
    Resume
    
End Function

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
    Dim k As Long
    
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
        k = frmSuperficies.LynxSuperficies.Rows - 1
        frmSuperficies.LynxSuperficies.CellText(k, 1) = SupData(i).Grh
        frmSuperficies.LynxSuperficies.CellText(k, 2) = SupData(i).name
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

    If FileExist(dirDats & "NPCs.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & dirDats, vbCritical
        Call CloseMapEditor
    End If

    Dim Trabajando As String
    Dim NPC As Long
    Dim Hostil As String
    Dim Leer As New clsIniManager
    Dim k As Long
    
    Call Leer.Initialize(dirDats & "NPCs.dat")
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
            
            k = frmNPCs.LynxNPCs.Rows - 1
            frmNPCs.LynxNPCs.CellText(k, 1) = .name
            frmNPCs.LynxNPCs.CellText(k, 2) = .ELV
            
            Hostil = IIf(.Hostile = 1, "SI", "NO")
            frmNPCs.LynxNPCs.CellText(k, 3) = Hostil
            
        End With
    Next
    
    frmNPCs.LynxNPCs.Visible = True
    frmNPCs.LynxNPCs.Redraw = True
    frmNPCs.LynxNPCs.ColForceFit
    
    DoEvents
    
    Set Leer = Nothing
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & " de " & Trabajando & " en " & dirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

Public Sub CargarIndicesOBJ()
'*************************************************
'Author: Lorwik
'Last modified: 28/04/2021
'*************************************************

On Error GoTo Fallo

    Dim Obj As Integer
    Dim k As Long
    Dim Leer As New clsIniManager

    If FileExist(dirDats & "\OBJ.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'OBJ.dat' en " & dirDats, vbCritical
        End
    End If

    Call Leer.Initialize(dirDats & "\OBJ.dat")
    
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
            .OBJType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
            .Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
            .Info = Leer.GetValue("OBJ" & Obj, "Info")
            .WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
            .Texto = Leer.GetValue("OBJ" & Obj, "Texto")
            .GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))
            .Cerrada = Val(Leer.GetValue("OBJ" & Obj, "Cerrada"))
            .Subtipo = Val(Leer.GetValue("OBJ" & Obj, "Subtipo"))
            
            frmOBJs.LynxOBJs.AddItem Obj
            k = frmOBJs.LynxOBJs.Rows - 1
            frmOBJs.LynxOBJs.CellText(k, 1) = .name
        
        End With
    Next Obj
    
    frmOBJs.LynxOBJs.Visible = True
    frmOBJs.LynxOBJs.Redraw = True
    frmOBJs.LynxOBJs.ColForceFit
    
    DoEvents
    
    Set Leer = Nothing
    Exit Sub
    
Fallo:
    MsgBox "Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & dirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

Public Sub CargarIndicesTriggers()
'*************************************************
'Author: Lorwik
'Last modified: 29/04/2021
' Carga los indices de Triggers
'*************************************************

On Error GoTo Fallo

    Dim k As Long

    If FileExist(IniPath & INITDIR & "Triggers.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Triggers.ini' en " & IniPath & INITDIR & "Triggers.ini", vbCritical
        End
    End If
    
    Dim NumT As Integer
    Dim t As Integer
    Dim Leer As New clsIniManager
    
    Call Leer.Initialize(IniPath & INITDIR & "Triggers.ini")
    
    frmTriggers.LynxTriggers.Clear
    frmTriggers.LynxTriggers.Redraw = False
    frmTriggers.LynxTriggers.Visible = False
    
    frmTriggers.LynxTriggers.AddColumn "Num", 0
    frmTriggers.LynxTriggers.AddColumn "Nombre", 2
    
    NumT = Val(Leer.GetValue("INIT", "NumTriggers"))
    For t = 1 To NumT
        frmTriggers.LynxTriggers.AddItem t
        k = frmTriggers.LynxTriggers.Rows - 1
        frmTriggers.LynxTriggers.CellText(k, 1) = Leer.GetValue("Trig" & t, "Name")
    Next t

    frmTriggers.LynxTriggers.Visible = True
    frmTriggers.LynxTriggers.Redraw = True
    frmTriggers.LynxTriggers.ColForceFit
    
    DoEvents

    Set Leer = Nothing
    Exit Sub
    
Fallo:
    MsgBox "Error al intentar cargar el Trigger " & t & " de Triggers.ini en " & IniPath & INITDIR & "Triggers.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub
