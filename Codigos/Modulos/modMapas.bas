Attribute VB_Name = "modMapas"
Option Explicit

'**********************
'Formato WinterAO
'**********************
Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroParticulas As Long
    NumeroLuces As Long
    NumeroZonas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
    NumeroData As Integer
End Type

Private Type tDatosBloqueados
    X As Integer
    Y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    Y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    Y As Integer
    Trigger As Integer
End Type

Private Type tDatosZonas
    X As Integer
    Y As Integer
    Zona As Integer
End Type

Public Type tDatosLuces
    r As Integer
    g As Integer
    b As Integer
    range As Byte
    X As Integer
    Y As Integer
End Type

Private Type tDatosParticulas
    X As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    Y As Integer
    NPCIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDat
    map_name As String
    battle_mode As Boolean
    backup_mode As Boolean
    restrict_mode As String
    music_number As String
    zone As String
    terrain As String
    ambient As String
    lvlMinimo As String
    RoboNpcsPermitido As Boolean
    InvocarSinEfecto As Boolean
    OcultarSinEfecto As Boolean
    ResuSinEfecto As Boolean
    MagiaSinEfecto As Boolean
    InviSinEfecto As Boolean
    LuzBase As Long
    version As Long
    NoTirarItems As Boolean
End Type

Public MapSize As tMapSize
Private MapDat() As tMapDat
'**********************

Public Sub AbrirMapa(Optional ByVal IntMode As Boolean = False)
    frmMain.Dialog.CancelError = True
    On Error GoTo ErrHandler
    
    DeseaGuardarMapa frmMain.Dialog.filename
    
    frmMain.ObtenerNombreArchivo False
    
    If Len(frmMain.Dialog.filename) < 3 Then Exit Sub
    
        'If WalkMode = True Then
        '    Call modGeneral.ToggleWalkMode
        'End If
        
        Call modMapas.NuevoMapa
        
        Call AbrirunMapa(frmMain.Dialog.filename, IntMode)
        
        DoEvents
        'frmMain.mnuReAbrirMapa.Enabled = True
        EngineRun = True
    
    Exit Sub
ErrHandler:
End Sub

Public Sub AbrirunMapa(ByVal Path As String, Optional ByVal Mode As Boolean = False)
'********************************
'Autor: Lorwik
'Fecha: 23/03/2021
'********************************

    If frmMain.Dialog.FilterIndex = 1 Then
        If ClientSetup.MeMode = eMeMode.WinterAO Then
            If Mode Then
                'Call modMapas.Cargar_CSM_Old(Path)
                
            Else
                Call modMapas.Cargar_CSM(Path)
                
            End If
                
'        #If Privado = 0 Then
'        ElseIf ClientSetup.WeMode = eWeMode.ImperiumClasico Then
'            Call modMapImpC.Cargar_MapImpClasico(Path)
'        #End If
'
        End If
    Else
        'Call MapaV2_Cargar(Path, Mode)
            
    End If
End Sub

''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional Path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

'    If MapInfo.Changed = 1 Then
'        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
'            GuardarMapa Path
'        End If
'    End If
End Sub

''
' Guarda el Mapa
'
' @param Path Especifica el path del mapa

Public Sub GuardarMapa(Optional Path As String)
'*************************************************
'Author: Lorwik
'Last modified: 26/04/2021
'*************************************************

    frmMain.Dialog.CancelError = True
On Error GoTo ErrHandler
    
    If LenB(Path) = 0 Then
        frmMain.ObtenerNombreArchivo True
        Path = frmMain.Dialog.filename
        If LenB(Path) = 0 Then Exit Sub
    End If
    
    If frmMain.Dialog.FilterIndex = 1 Then
        
        Select Case ClientSetup.MeMode
        
            Case eMeMode.WinterAO
                'Call Save_CSM(Path)
                
        End Select
        
            
    ElseIf frmMain.Dialog.FilterIndex = 2 Then
        'Call MapaV2_Guardar(Path)
        
    End If

ErrHandler:
End Sub

Public Sub NuevoMapa()
    '***************************************************
    'Author: Lorwik
    'Last Modification: 26/04/2021
    'Descripcion: Limpia todo el mapa a uno nuevo
    '***************************************************
    
    Dim Y As Integer
    Dim X As Integer
    Dim i As Byte
    
    'Volvemos a setear el tamaño del mapa
    Call setMapSize
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            With MapData(X, Y)
            
                .Graphic(1).GrhIndex = 1
                
                ' Bloqueos
                .Blocked = 0
        
                ' Capas 2, 3 y 4
                .Graphic(2).GrhIndex = 0
                .Graphic(3).GrhIndex = 0
                .Graphic(4).GrhIndex = 0
                
                ' OBJs
                .OBJInfo.ObjIndex = 0
                .OBJInfo.Amount = 0
                .ObjGrh.GrhIndex = 0
        
                ' Translados
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
                
                ' Triggers
                .Trigger = 0
        
                .Particle_Group_Index = 0
                
                Call Engine_Long_To_RGB_List(MapData(X, Y).Engine_Light(), -1)
                
                .Light.active = False
                .Light.range = 0
                .Light.map_x = 0
                .Light.map_y = 0
                .Light.RGBCOLOR.a = 0
                .Light.RGBCOLOR.r = 0
                .Light.RGBCOLOR.g = 0
                .Light.RGBCOLOR.b = 0
                
                If ClientSetup.MeMode = eMeMode.WinterAO Then _
                    .ZonaIndex = 0
                
                For i = 0 To 3
                    .Engine_Light(i) = 0
                Next i

                InitGrh .Graphic(1), 1

            End With
            
        Next X
    Next Y
    
    'Borramos todas las luces
    Call LightRemoveAll
    
    If ClientSetup.MeMode = eMeMode.WinterAO Then
        CantZonas = 0
        ReDim MapZonas(CantZonas) As tMapInfo
    
        frmZonas.LstZona.Clear
    
        Call NuevaZona(CantZonas)
                
        'Call MapZona_Actualizar(CantZonas)
        
        frmZonas.LstZona.ListIndex = 0
        
        'Call MapZona_Actualizar(frmZonas.LstZona.ListIndex + 1)
           
    Else
    
        MapInfo.MapVersion = 0
        MapInfo.name = "Mapa Desconocido"
        MapInfo.Music = 0
        MapInfo.ambient = 0
        MapInfo.PK = True
        MapInfo.MagiaSinEfecto = 0
        MapInfo.InviSinEfecto = 0
        MapInfo.ResuSinEfecto = 0
        MapInfo.Terreno = "BOSQUE"
        MapInfo.Zona = "CAMPO"
        MapInfo.Restringir = "No"
        MapInfo.NoEncriptarMP = 0
        MapInfo.LuzBase = 0
            
        'Call MapInfo_Actualizar
    End If
    
    Estado_Actual = Estados(e_estados.MedioDia)
    Call Actualizar_Estado
    
    'Set changed flag
    MapInfo.Changed = 0
    frmMain.MousePointer = 0
    
End Sub

Public Sub NuevaZona(ByVal id As Integer)
'*****************************************
'Autor: Lorwik
'Fecha: 01/04/2021
'Descripción: Crea una nueva zona
'*****************************************

    CantZonas = CantZonas + 1
    
    ReDim Preserve MapZonas(CantZonas) As tMapInfo

    Call ResetearZona(id)
    
    frmZonas.LstZona.AddItem (CantZonas & "- " & MapZonas(CantZonas).name)
End Sub

Public Sub ResetearZona(ByVal id As Integer)
'*****************************************
'Autor: Lorwik
'Fecha: 01/04/2021
'Descripción: Resetea las propiedades de la zona indicada
'*****************************************

    If id > CantZonas Then
        MsgBox "Error al resetear la zona. El ID de la zona seleccionada es superior al numero de zonas existentes.", vbCritical
        Exit Sub
    End If

    With MapZonas(CantZonas)
        .MapVersion = 0
        .name = "Zona Desconocida"
        .Music = 0
        .ambient = 0
        .PK = True
        .MagiaSinEfecto = 0
        .InviSinEfecto = 0
        .ResuSinEfecto = 0
        .Terreno = "BOSQUE"
        .Zona = "CAMPO"
        .Restringir = "No"
        .NoEncriptarMP = 0
        .LuzBase = 0
    End With
    
End Sub

Public Sub EliminarZona()
'*****************************************
'Autor: Lorwik
'Fecha: 01/04/2021
'Descripción: Elimina la ultima zona creada
'*****************************************
    
    If CantZonas = 1 Then
        MsgBox "El numero de zonas llego al mnimo. No puedes eliminar mas zonas."
        Exit Sub
    End If
    
    'Primero se resetea la zona
    Call ResetearZona(CantZonas)

    frmZonas.LstZona.RemoveItem frmZonas.LstZona.ListIndex
    
    CantZonas = CantZonas - 1
    
    ReDim Preserve MapZonas(CantZonas) As tMapInfo
    
End Sub

Public Sub ActualizarZonaList()
'*****************************************
'Autor: Lorwik
'Fecha: 02/04/2021
'Descripción: Actualiza la lista de zonas
'*****************************************

    Dim i As Integer

    frmZonas.LstZona.Clear
        
    For i = 1 To CantZonas
        frmZonas.LstZona.AddItem (i & "- " & MapZonas(i).name)
            
    Next i
    
End Sub

Sub Cargar_CSM(ByVal Map As String)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 14/03/2021
    'Descripcion: Carga los mapas de WinterAO
    '***************************************************
    
    On Error GoTo ErrorHandler
    
    Dim fh As Integer
    Dim File As Integer
    Dim MH As tMapHeader
    Dim Blqs() As tDatosBloqueados
    Dim L1() As Long
    Dim L2() As tDatosGrh
    Dim L3() As tDatosGrh
    Dim L4() As tDatosGrh
    Dim Triggers() As tDatosTrigger
    Dim Particulas() As tDatosParticulas
    Dim Luces() As tDatosLuces
    Dim Zonas() As tDatosZonas
    Dim Objetos() As tDatosObjs
    Dim NPCs() As tDatosNPC
    Dim TEs() As tDatosTE
    
    Dim LaCabecera As tCabecera
    
    Dim i As Long
    Dim j As Long
    DoEvents
        
    ClientSetup.TipoMapaCargado = eTipoMapa.tWinter
        
    'Change mouse icon
    frmMain.MousePointer = 11
        
    fh = FreeFile
    Open Map For Binary Access Read As fh
    
        Get #fh, , LaCabecera
    
        Get #fh, , MH
        Get #fh, , MapSize
        
        ReDim MapDat(MH.NumeroData) As tMapDat
        Get #fh, , MapDat
        
        With MapSize
            If Not .XMax = XMaxMapSize Or Not .YMax = YMaxMapSize Then
                ReDim MapData(.XMin To .XMax, .YMin To .YMax)
            End If
            ReDim L1(.XMin To .XMax, .YMin To .YMax)
        End With
                   
        Get #fh, , L1
        
        With MH
            If .NumeroBloqueados > 0 Then
                ReDim Blqs(1 To .NumeroBloqueados)
                Get #fh, , Blqs
                For i = 1 To .NumeroBloqueados
                    MapData(Blqs(i).X, Blqs(i).Y).Blocked = 1
                Next i
            End If
            
            If .NumeroLayers(2) > 0 Then
                ReDim L2(1 To .NumeroLayers(2))
                Get #fh, , L2
                For i = 1 To .NumeroLayers(2)
                    InitGrh MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).GrhIndex
                Next i
            End If
            
            If .NumeroLayers(3) > 0 Then
                ReDim L3(1 To .NumeroLayers(3))
                Get #fh, , L3
                For i = 1 To .NumeroLayers(3)
                    InitGrh MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).GrhIndex
                Next i
            End If
            
            If .NumeroLayers(4) > 0 Then
                ReDim L4(1 To .NumeroLayers(4))
                Get #fh, , L4
                For i = 1 To .NumeroLayers(4)
                    InitGrh MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).GrhIndex
                Next i
            End If
            
            If .NumeroTriggers > 0 Then
                ReDim Triggers(1 To .NumeroTriggers)
                Get #fh, , Triggers
                For i = 1 To .NumeroTriggers
                    MapData(Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
                Next i
            End If
            
            If .NumeroParticulas > 0 Then
                ReDim Particulas(1 To .NumeroParticulas)
                Get #fh, , Particulas
'                For i = 1 To .NumeroParticulas
'                    MapData(Particulas(i).X, Particulas(i).Y).Particle_Index = Particulas(i).Particula
'                    Call General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).Y)
'
'                    MapData(Particulas(i).X, Particulas(i).Y).Particle_Group_Index = General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).Y)
'                Next i
            End If
'
            If .NumeroLuces > 0 Then
                ReDim Luces(1 To .NumeroLuces)
                Dim p As Byte
                Get #fh, , Luces
'                For i = 1 To .NumeroLuces
'
'                    With MapData(Luces(i).X, Luces(i).Y)
'                        .Light.range = Luces(i).range ' Changed by: Project Administrator at: 4/26/2021-20:26:18 on machine: DESKTOP-BR8H09Q
'                        .Light.RGBCOLOR.a = 255
'                        .Light.RGBCOLOR.R = Luces(i).R
'                        .Light.RGBCOLOR.G = Luces(i).G
'                        .Light.RGBCOLOR.B = Luces(i).B
'
'                    End With
'
'                    Call Create_Light_To_Map(Luces(i).X, Luces(i).Y, Luces(i).range, Luces(i).R, Luces(i).G, Luces(i).B)
'                Next i
'
'                Call LightRenderAll
            End If
'
            If .NumeroZonas > 0 Then
                ReDim Zonas(1 To .NumeroZonas)
                Get #fh, , Zonas
                For i = 1 To .NumeroZonas
                    MapData(Zonas(i).X, Zonas(i).Y).ZonaIndex = Zonas(i).Zona
                Next i
            End If
            
            If .NumeroOBJs > 0 Then
                ReDim Objetos(1 To .NumeroOBJs)
                Get #fh, , Objetos
'                For i = 1 To .NumeroOBJs
'                    MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.ObjIndex = Objetos(i).ObjIndex
'                    MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.Amount = Objetos(i).ObjAmmount
'                    If MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.ObjIndex > numObjs Then
'                        InitGrh MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, 20299
'                    Else
'                        InitGrh MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, ObjData(MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.ObjIndex).GrhIndex
'                    End If
'                Next i
            End If
                
            If .NumeroNPCs > 0 Then
                ReDim NPCs(1 To .NumeroNPCs)
                Get #fh, , NPCs
'                For i = 1 To .NumeroNPCs
'                    If NPCs(i).NPCIndex > 0 Then
'                        MapData(NPCs(i).X, NPCs(i).Y).NPCIndex = NPCs(i).NPCIndex
'                        Call MakeChar(NextOpenChar(), NpcData(NPCs(i).NPCIndex).Body, NpcData(NPCs(i).NPCIndex).Head, NpcData(NPCs(i).NPCIndex).Heading, NPCs(i).X, NPCs(i).Y)
'                    End If
'                Next i
            End If
    
            If .NumeroTE > 0 Then
                ReDim TEs(1 To .NumeroTE)
                Get #fh, , TEs
'                For i = 1 To .NumeroTE
'                    MapData(TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
'                    MapData(TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
'                    MapData(TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
'                Next i
            End If
             
        End With
    
    Close fh
    
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            If L1(i, j) > 0 Then
                InitGrh MapData(i, j).Graphic(1), L1(i, j)
            End If
        Next i
    Next j
    
'    'MapInfo_Cargar Map
'    frmMapInfo.txtMapVersion.Text = MapInfo.MapVersion
'
    Call Pestanas(Map, ".csm")
'
'    'Change mouse icon
    frmMain.MousePointer = 0
'
'    ' Vacio deshacer
'    modEdicion.Deshacer_Clear
'
'    Call CSMInfoCargar
'
'    'Set changed flag
'    MapInfo.Changed = 0
'
    MapaCargado = True
'
'    Call DibujarMinimapa ' Radar
'
    Call AddtoRichTextBox(frmConsola.StatTxt, "Mapa " & Map & " cargado...", 0, 255, 0)

ErrorHandler:
    If fh <> 0 Then Close fh
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Error en el Mapa " & Map & ", se ha generado un informe de errores en: " & App.Path & "\Logs.txt", 255, 0, 0)
    
    File = FreeFile
    
    Call RegistrarError(Err.Number, Err.Description, "modMapas.Cargar_CSM", Erl)
End Sub

''
' Calcula la orden de Pestanas
'
' @param Map Especifica path del mapa

Public Sub Pestanas(ByVal Map As String, Optional ByVal MapFormat As String = ".map")
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
On Error Resume Next
    Dim loopc As Integer
    
    For loopc = Len(Map) To 1 Step -1
        If mid(Map, loopc, 1) = "\" Then
            PATH_Save = Left(Map, loopc)
            Exit For
        End If
    Next
    
    Map = Right(Map, Len(Map) - (Len(PATH_Save)))
    
    MapaActual = ReadField(1, Right(Map, Len(Map) - 4), Asc("."))
    'If frmCopiarBordes.Visible Then Call frmCopiarBordes.Inicializar
    
    For loopc = Len(Left(Map, Len(Map) - 4)) To 1 Step -1
        If IsNumeric(mid(Left(Map, Len(Map) - 4), loopc, 1)) = False Then
            NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - loopc)
            NameMap_Save = Left(Map, loopc)
            Exit For
        End If
    Next
    
    For loopc = (NumMap_Save - 4) To (NumMap_Save + 8)
            If FileExist(PATH_Save & NameMap_Save & loopc & MapFormat, vbArchive) = True Then
                frmMain.MapPest(loopc - NumMap_Save + 4).Visible = True
                frmMain.MapPest(loopc - NumMap_Save + 4).Enabled = True
                frmMain.MapPest(loopc - NumMap_Save + 4).Caption = NameMap_Save & loopc
            Else
                frmMain.MapPest(loopc - NumMap_Save + 4).Visible = False
            End If
    Next
    
End Sub
