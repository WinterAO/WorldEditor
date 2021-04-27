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
End Enum

Public Enum eTipoMapa
    tInt
    tLong
    tWinter
    tIAOClasico
    tIAOnew
    tIAOold
    tWinter_Old
End Enum

Public Type tSetupMods

    ' VIDEO
    byMemory    As Integer
    LimiteFPS As Boolean
    OverrideVertexProcess As Byte
    TilesBuffer As Byte
    
    'MOSTRAR
    MapTam As Byte
    Preview As Boolean
    
    'CONFIGURACION
    MeMode As Byte
    
    'OTROS
    TipoMapaCargado As eTipoMapa
End Type

Public ClientSetup As tSetupMods

'Path
Public IniPath As String
Public DirRecursos As String
Public DirDats As String

'Recuento de indices
Public grhCount    As Long
Public MaxSup      As Integer

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
        .Desc = "WinterAO Resurrection mod Argentum Online by Noland Studios. http://winterao.com.ar"
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
        DirRecursos = autoCompletaPath(Lector.GetValue("PATH" & .MeMode, "DirRecursos"))
        
        If FileExist(DirRecursos, vbDirectory) = False Or DirRecursos = "\" Then
            MsgBox "El directorio de Recursos es incorrecto", vbCritical + vbOKOnly
            
            NewPath = Buscar_Carpeta("DirRecursos", "")
            Call WriteVar(configFile, "PATH" & .MeMode, "DirRecursos", NewPath)
            DirRecursos = NewPath & "\"
        End If
        
        If FileExist(DirRecursos & "Graficos" & Formato, vbArchive) = False Then
            MsgBox "No se encontro el recursos de graficos."
            End
        End If
        
        If FileExist(DirRecursos & "Scripts" & Formato, vbArchive) = False Then
            MsgBox "No se encontro el recursos de Scripts."
            End
        End If
        
        '-------------------
        'Dats
        DirDats = autoCompletaPath(Lector.GetValue("PATH" & .MeMode, "DirDats"))
        
        If FileExist(DirDats, vbDirectory) = False Or DirDats = "\" Then
            MsgBox "El directorio de Dats es incorrecto", vbCritical + vbOKOnly
            
            NewPath = Buscar_Carpeta("DirDats", "")
            Call WriteVar(configFile, "PATH" & .MeMode, "DirDats", NewPath)
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
        frmMain.Minimap_capa2.Checked = Val(Lector.GetValue("MINIMAP", "Capa2"))
        frmMain.Minimap_capa3.Checked = Val(Lector.GetValue("MINIMAP", "Capa3"))
        frmMain.Minimap_capa4.Checked = Val(Lector.GetValue("MINIMAP", "Capa4"))
        frmMain.Minimap_objetos.Checked = Val(Lector.GetValue("MINIMAP", "Obj"))
        frmMain.Minimap_npcs.Checked = Val(Lector.GetValue("MINIMAP", "NPC"))
        frmMain.Minimap_particulas.Checked = Val(Lector.GetValue("MINIMAP", "Particulas"))
        frmMain.Minimap_ndemapa.Checked = Val(Lector.GetValue("MINIMAP", "Nombre"))
        frmMain.Minimap_bloqueos.Checked = Val(Lector.GetValue("MINIMAP", "Bloqueos"))
        
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
    Dim Buffer()    As Byte
    
    InfoHead = File_Find(DirRecursos & "Scripts.WAO", LCase$("Graficos.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Graficos.ind"), Buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader Buffer
        
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
                    
                    .Trans = fileBuff.getByte
                    If .Trans < 0 Then GoTo ErrorHandler
                    
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                    
                End If
                
            End With
            
        Wend
        
        Erase Buffer
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
    Dim Buffer()    As Byte
    Dim i           As Long
    
    InfoHead = File_Find(DirRecursos & "Scripts" & Formato, LCase$("minimap.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("minimap.ind"), Buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader Buffer
        
        For i = 1 To grhCount
            If Grh_Check(i) Then
                GrhData(i).mini_map_color = fileBuff.getLong
            End If
        Next i
        
        Erase Buffer
    End If
    
    Set fileBuff = Nothing
    
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

''
' Carga los indices de Superficie
'

Public Sub CargarIndicesSuperficie()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
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
    'frmMain.lListado(0).Clear
    
    frmSuperficies.LynxSuperficies.Clear
    frmSuperficies.LynxSuperficies.Redraw = False
    frmSuperficies.LynxSuperficies.Visible = False

    frmSuperficies.LynxSuperficies.AddColumn "Grh", 0
    frmSuperficies.LynxSuperficies.AddColumn "Nombre", 3
    
    For i = 0 To MaxSup
        SupData(i).name = Leer.GetValue("REFERENCIA" & i, "Nombre")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))

        frmSuperficies.LynxSuperficies.AddItem SupData(i).Grh
        K = frmSuperficies.LynxSuperficies.Rows - 1
        frmSuperficies.LynxSuperficies.CellText(K, 1) = SupData(i).name
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
