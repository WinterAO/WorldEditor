Attribute VB_Name = "modCarga"
Option Explicit

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Enum eMeMode
    WinterAO
    ImperiumClasico
    WinterUltimate
End Enum

Public Enum eTipoMapa
    tInt
    tlong
    tWinter
    tIAOClasico
    tIAOnew
    tIAOold
    tWinter_Old
End Enum

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
    MapTam As Byte
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

Public Function configFile() As String
    configFile = IniPath & INITDIR & "configuracion.ini"
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
        If eMeMode.WinterAO Then
            .Desc = "WinterAO Resurrection mod Argentum Online by Noland Studios. http://winterao.com.ar"
            
        ElseIf eMeMode.ImperiumClasico Then
            .Desc = "Imperium Clasico mod Argentum Online by Comunidad Winter. http://imperiumclasico.com.ar"
            
        End If
        
        .CRC = Rnd * 245
        .MagicWord = Rnd * 92
    End With
    
End Sub

Public Function pre_leerConfiguracion() As Boolean
'**********************************
'Autor: Lorwik
'Fecha: 26/04/2021
'Descripcion: Pre-Lee la configuracion de WorldEditor
'**********************************

On Local Error GoTo fileErr:
    
    Dim i As Byte
    Dim Lector   As clsIniManager
    Set Lector = New clsIniManager
    
    If Not FileExist(configFile, vbArchive) Then
        MsgBox "¡No se ha encontrado el archivo de configuracion en la carpeta init!", vbOKOnly Or vbExclamation, App.Title
        End
    End If
    
    Call Lector.Initialize(configFile)
    
    With ClientSetup
    
        .byMemory = Lector.GetValue("VIDEO", "DynamicMemory")
        .OverrideVertexProcess = CByte(Lector.GetValue("VIDEO", "VertexProcessingOverride"))
        .LimiteFPS = CBool(Lector.GetValue("VIDEO", "LimitarFPS"))
        .TilesBuffer = Val(Lector.GetValue("VIDEO", "TilesBuffer"))
        .Preview = CBool(Lector.GetValue("MOSTRAR", "Preview"))
        .MeMode = Lector.GetValue("CONFIGURACION", "MeMode")
        
    End With

    Set Lector = Nothing

    pre_leerConfiguracion = True

  Exit Function
  
fileErr:

    pre_leerConfiguracion = False

    If Err.Number <> 0 Then
       MsgBox ("Ha ocurrido un error al pre-cargar la configuracion del cliente. Error " & Err.Number & " : " & Err.Description)
       End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    End If

End Function

Public Function leerConfiguracion() As Boolean
'**********************************
'Autor: Lorwik
'Fecha: 26/04/2021
'Descripcion: Lee la configuracion de WorldEditor
'**********************************

On Local Error GoTo fileErr:
    
    Dim Lector   As clsIniManager
    Set Lector = New clsIniManager
    Dim i As Byte
    Dim NewPath As String
    
    If Not FileExist(configFile, vbArchive) Then
        MsgBox "¡No se ha encontrado el archivo de configuracion en la carpeta init!", vbOKOnly Or vbExclamation, App.Title
        End
    End If
    
    Call Lector.Initialize(configFile)
    
    With ClientSetup

        '****
        'RUTAS
        '*****************
        
        '-------------------
        'Recursos
        DirRecursos = autoCompletaPath(Lector.GetValue("PATH-" & namePerfil & "-" & .MeMode, "DirRecursos"))
        
        If FileExist(DirRecursos, vbDirectory) = False Or DirRecursos = "\" Then
            MsgBox "El directorio de Recursos es incorrecto", vbCritical + vbOKOnly
            
            NewPath = Buscar_Carpeta("DirRecursos", "")
            Call WriteVar(configFile, "PATH-" & namePerfil & "-" & .MeMode, "DirRecursos", NewPath)
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
        DirDats = autoCompletaPath(Lector.GetValue("PATH-" & namePerfil & "-" & .MeMode, "DirDats"))
        
        If FileExist(DirDats, vbDirectory) = False Or DirDats = "\" Then
            MsgBox "El directorio de Dats es incorrecto", vbCritical + vbOKOnly
            
            NewPath = Buscar_Carpeta("DirDats", "")
            Call WriteVar(configFile, "PATH-" & namePerfil & "-" & .MeMode, "DirDats", NewPath)
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
        PantallaX = Val(Lector.GetValue("MOSTRAR", "PantallaX"))
        PantallaY = Val(Lector.GetValue("MOSTRAR", "PantallaY"))
        If PantallaX > 23 Or PantallaX <= 2 Then PantallaX = 23
        If PantallaY > 32 Or PantallaY <= 2 Then PantallaY = 32
        
        ' [GS] 02/10/06
        ' Tamaño de visualizacion en el cliente
        ClienteHeight = Val(Lector.GetValue("MOSTRAR", "ClienteHeight"))
        ClienteWidth = Val(Lector.GetValue("MOSTRAR", "ClienteWidth"))
        If ClienteHeight <= 0 Then ClienteHeight = 13
        If ClienteWidth <= 0 Then ClienteWidth = 17
        
        ' Menu Mostrar
        VerCapa1 = Val(Lector.GetValue("MOSTRAR", "Capa1"))
        VerCapa2 = Val(Lector.GetValue("MOSTRAR", "Capa2"))
        VerCapa3 = Val(Lector.GetValue("MOSTRAR", "Capa3"))
        VerCapa4 = Val(Lector.GetValue("MOSTRAR", "Capa4"))
        VerTranslados = Val(Lector.GetValue("MOSTRAR", "Translados"))
        VerObjetos = Val(Lector.GetValue("MOSTRAR", "Objetos"))
        VerNpcs = Val(Lector.GetValue("MOSTRAR", "NPCs"))
        VerTriggers = Val(Lector.GetValue("MOSTRAR", "Triggers"))
        VerGrilla = Val(Lector.GetValue("MOSTRAR", "Grilla")) ' Grilla
        VerParticulas = Val(Lector.GetValue("MOSTRAR", "Particulas"))
        VerBlockeados = Val(Lector.GetValue("MOSTRAR", "Bloqueos"))
        
        frmMain.Minimap_capa1.Checked = Val(Lector.GetValue("MINIMAP", "Capa1"))
        MMiniMap_capa1 = frmMain.Minimap_capa1.Checked
        frmMain.Minimap_capa2.Checked = Val(Lector.GetValue("MINIMAP", "Capa2"))
        MMiniMap_capa2 = frmMain.Minimap_capa2.Checked
        frmMain.Minimap_capa3.Checked = Val(Lector.GetValue("MINIMAP", "Capa3"))
        MMiniMap_capa3 = frmMain.Minimap_capa3.Checked
        frmMain.Minimap_capa4.Checked = Val(Lector.GetValue("MINIMAP", "Capa4"))
        MMiniMap_capa4 = frmMain.Minimap_capa4.Checked
        frmMain.Minimap_objetos.Checked = Val(Lector.GetValue("MINIMAP", "Obj"))
        MMiniMap_objetos = frmMain.Minimap_objetos.Checked
        frmMain.Minimap_npcs.Checked = Val(Lector.GetValue("MINIMAP", "NPC"))
        MMiniMap_Npcs = frmMain.Minimap_npcs.Checked
        frmMain.Minimap_particulas.Checked = Val(Lector.GetValue("MINIMAP", "Particulas"))
        MMiniMap_particulas = frmMain.Minimap_particulas.Checked
        frmMain.Minimap_ndemapa.Checked = Val(Lector.GetValue("MINIMAP", "Nombre"))
        MMiniMap_Nombre = frmMain.Minimap_ndemapa.Checked
        frmMain.Minimap_bloqueos.Checked = Val(Lector.GetValue("MINIMAP", "Bloqueos"))
        MMiniMap_Bloqueos = frmMain.Minimap_bloqueos.Checked
        
        ' AUDIO
        .bMusic = CByte(Lector.GetValue("AUDIO", "MUSICA"))
        .bSound = CByte(Lector.GetValue("AUDIO", "SONIDO"))
        .bAmbient = CByte(Lector.GetValue("AUDIO", "AMBIENT"))
        .MusicVolume = CLng(Lector.GetValue("AUDIO", "VOLMUSICA"))
        .SoundVolume = CLng(Lector.GetValue("AUDIO", "VOLAUDIO"))
        .AmbientVol = CLng(Lector.GetValue("AUDIO", "VOLAMBIENT"))
        
    End With

    Set Lector = Nothing

    leerConfiguracion = True

  Exit Function
  
fileErr:

    leerConfiguracion = False

    If Err.Number <> 0 Then
       MsgBox ("Ha ocurrido un error al cargar la configuracion del cliente. Error " & Err.Number & " : " & Err.Description)
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
    Dim LaCabecera  As tCabecera
    Dim fileBuff    As clsByteBuffer
    Dim InfoHead    As INFOHEADER
    Dim buffer()    As Byte
    
    InfoHead = File_Find(DirRecursos & "Scripts" & Formato, LCase$("Graficos.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Graficos.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        fileVersion = fileBuff.getLong
        
        grhCount = fileBuff.getLong
        
        ReDim GrhData(0 To grhCount) As GrhData
        
        While Grh < grhCount
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
                    
                    '.Trans = fileBuff.getByte
                    'If .Trans < 0 Then GoTo ErrorHandler
                    
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
'************************************
'Autor: Lorwik
'Fecha: ???
'************************************

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
    Dim LaCabecera As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(DirRecursos & "Scripts" & modCompression.Formato, LCase$("Personajes.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Personajes.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        'num de cabezas
        NumCuerpos = fileBuff.getInteger()
    
        'Resize array
        ReDim BodyData(0 To NumCuerpos) As tBodyData
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
            Call MsgBox("El archivo Personajes.ind no existe. ")
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
    
    frmObjs.LynxOBJs.Clear
    frmObjs.LynxOBJs.Redraw = False
    frmObjs.LynxOBJs.Visible = False
    
    NumObjs = Val(Leer.GetValue("INIT", "NumOBJs"))
    ReDim ObjData(1 To NumObjs) As ObjData
    
    frmObjs.LynxOBJs.AddColumn "Num", 0
    frmObjs.LynxOBJs.AddColumn "Nombre", 2
    
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
        
        frmObjs.LynxOBJs.AddItem Obj
        K = frmObjs.LynxOBJs.Rows - 1
        frmObjs.LynxOBJs.CellText(K, 1) = .name
        
        End With
    Next Obj
    
    frmObjs.LynxOBJs.Visible = True
    frmObjs.LynxOBJs.Redraw = True
    frmObjs.LynxOBJs.ColForceFit
    
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
