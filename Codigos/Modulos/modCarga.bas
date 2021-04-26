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

'Constantes
Public Const INITDIR As String = "\Init\"

Public Function configFile() As String
    configFile = App.Path & INITDIR & "configuracion.ini"
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
        .LimiteFPS = CBool(Lector.GetValue("Video", "LimitarFPS"))
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
    
    InfoHead = File_Find(DirRecursos & "Scripts.WAO", LCase$("Graficos.ind"))
    
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
                    
                    .Trans = fileBuff.getByte
                    If .Trans < 0 Then GoTo ErrorHandler
                    
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
