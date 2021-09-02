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

    R As Integer
    G As Integer
    B As Integer
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

Public MapSize   As tMapSize

Private MapDat() As tMapDat
'*************************
'Fin formato carga Winter
'*************************

'-----------------------------------------

'**********************
'Formato Imperium Clasico
'**********************

Private Type tMapHeaderIAC
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueadosIAC
    X As Integer
    Y As Integer
End Type

Private Type tDatosGrhIAC
    X As Integer
    Y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTriggerIAC
    X As Integer
    Y As Integer
    Trigger As Integer
End Type

Private Type tDatosLucesIAC
    R As Integer
    G As Integer
    B As Integer
    range As Byte
    X As Integer
    Y As Integer
End Type

Private Type tDatosParticulasIAC
    X As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPCIAC
    X As Integer
    Y As Integer
    NPCIndex As Integer
End Type

Private Type tDatosObjsIAC
    X As Integer
    Y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTEIAC
    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSizeIAC
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDatIAC
    map_name As String
    battle_mode As Boolean
    backup_mode As Boolean
    restrict_mode As String
    music_number As String
    zone As String
    terrain As String
    ambient As String
    lvlMinimo As String
    LuzBase As Long
    version As Long
    NoTirarItems As Boolean
End Type

Public MapSizeIAC As tMapSizeIAC
Private MapDatIAC As tMapDatIAC

'*************************
'Fin formato carga Imperium Clasico
'*************************

Public Sub AbrirMapa(Optional ByVal IntMode As Boolean = False)
    frmMain.Dialog.CancelError = True

    On Error GoTo AbrirMapa_Err
    
    DeseaGuardarMapa frmMain.Dialog.filename
    
    frmMain.ObtenerNombreArchivo False
    
    If Len(frmMain.Dialog.filename) < 3 Then Exit Sub
    
    If WalkMode = True Then _
        Call modGeneral.ToggleWalkMode
        
    Call modMapas.NuevoMapa
        
    Call AbrirunMapa(frmMain.Dialog.filename, IntMode)
        
    DoEvents
    frmMain.mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
    Exit Sub
    
AbrirMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapas.AbrirMapa", Erl)
    Resume Next
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
            
        ElseIf ClientSetup.MeMode = eMeMode.ImperiumClasico Then
            #If Privado = 0 Then
                Call modMapas.Cargar_MapImpClasico(Path, tIAOClasico)
            #End If
            
        End If

    Else
        Call modMapas.Cargar_Map(Path, Mode)
            
    End If

End Sub

Public Sub DeseaGuardarMapa(Optional Path As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    'Nos pregunta donde guardar el mapa en caso de modificarlo
    '*************************************************

    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            GuardarMapa Path

        End If

    End If

End Sub

Public Sub GuardarMapa(Optional Path As String)
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2021
    'Guarda el Mapa
    '*************************************************

    frmMain.Dialog.CancelError = True

    On Error GoTo errhandler
    
    If LenB(Path) = 0 Then
        frmMain.ObtenerNombreArchivo True
        Path = frmMain.Dialog.filename

        If LenB(Path) = 0 Then Exit Sub

    End If
    
    If frmMain.Dialog.FilterIndex = 1 Then
        
        Select Case ClientSetup.MeMode
        
            Case eMeMode.WinterAO
                Call Save_CSM(Path)
                
        End Select
            
    ElseIf frmMain.Dialog.FilterIndex = 2 Then
        Call Guardar_Map(Path)
        
    End If

errhandler:

End Sub

Public Sub NuevoMapa()
    '***************************************************
    'Author: Lorwik
    'Last Modification: 26/04/2021
    'Descripcion: Limpia todo el mapa a uno nuevo
    '***************************************************
    
    Dim Y     As Integer

    Dim X     As Integer

    Dim i     As Byte

    Dim LoopC As Integer
    
    frmMain.mnuReAbrirMapa.Enabled = False
    
    MapaCargado = False
    
    For LoopC = 0 To frmMain.MapPest.Count - 1
        frmMain.MapPest(LoopC).Enabled = False
    Next
    
    frmMain.MousePointer = 11
    
    'Volvemos a setear el tamaño del mapa
    Call setMapSize
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            With MapData(X, Y)
            
                .Graphic(1).GrhIndex = 1
                
                ' Bloqueos
                .bLocked = 0
        
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
                .Light.RGBCOLOR.R = 0
                .Light.RGBCOLOR.G = 0
                .Light.RGBCOLOR.B = 0
                
                If ClientSetup.MeMode = eMeMode.WinterAO Then .ZonaIndex = 0
                
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
        
        frmZonas.LstZona.ListIndex = 0
        
        Call MapZona_Actualizar(frmZonas.LstZona.ListIndex + 1)
           
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
            
        Call MapInfo_Actualizar

    End If
    
    Call DibujarMinimapa
    
    Estado_Actual = Estados(e_estados.MedioDia)
    Call Actualizar_Estado
    
    'Set changed flag
    MapInfo.Changed = 0
    frmMain.MousePointer = 0
    
    MapaCargado = True
    EngineRun = True
    
End Sub

Public Sub MapInfo_Actualizar()
    '*************************************************
    'Author: Lorwik
    'Last modified: 29/04/2021
    ' Actualiza el formulario de MapInfo
    '*************************************************

    On Error Resume Next

    With frmMapInfo
        .txtMapNombre.Text = MapInfo.name
        .txtMapMusica.Text = MapInfo.Music
        .txtMapTerreno.Text = MapInfo.Terreno
        .txtMapZona.Text = MapInfo.Zona
        .txtMapRestringir.Text = MapInfo.Restringir
        '   .chkMapBackup.value = MapInfo.BackUp
        .chkMapPK.value = IIf(MapInfo.PK = True, 1, 0)
        .TxtAmbient.Text = MapInfo.ambient
        .TxtlvlMinimo = MapInfo.lvlMinimo
        .chkMapMagiaSinEfecto.value = MapInfo.MagiaSinEfecto
        .chkMapInviSinEfecto.value = IIf(MapInfo.InviSinEfecto, vbChecked, vbUnchecked)
        .chkInvocarSin.value = MapInfo.InvocarSinEfecto
        .chkOcultarSin.value = MapInfo.OcultarSinEfecto
        .chkMapResuSinEfecto.value = IIf(MapInfo.ResuSinEfecto, vbChecked, vbUnchecked)
        .txtMapVersion = MapInfo.MapVersion
        .ChkMapNpc.value = MapInfo.RoboNpcsPermitido

    End With

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

Public Sub ActualizarZonaList(ByVal id As Integer)
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
    
    frmZonas.LstZona.ListIndex = id
    
End Sub

Public Sub MapZona_Actualizar(ByVal id As Integer)
    
    Dim tR As Byte

    Dim tG As Byte

    Dim tB As Byte
    
    Call ConvertLongToRGB(MapZonas(id).LuzBase, tR, tG, tB)

    With frmMapInfo
        .txtMapNombre.Text = MapZonas(id).name
        .txtMapMusica.Text = MapZonas(id).Music
        .txtMapTerreno.Text = MapZonas(id).Terreno
        .txtMapZona = MapZonas(id).Zona
        .txtMapRestringir = MapZonas(id).Restringir
        '   .chkMapBackup.value = MapZonas(ID).BackUp
        .chkMapPK.value = IIf(MapZonas(id).PK = True, 1, 0)
        .TxtAmbient.Text = MapZonas(id).ambient
        .TxtlvlMinimo = MapZonas(id).lvlMinimo
        .chkMapMagiaSinEfecto.value = IIf(MapZonas(id).MagiaSinEfecto, vbChecked, vbUnchecked)
        .chkMapInviSinEfecto.value = IIf(MapZonas(id).InviSinEfecto, vbChecked, vbUnchecked)
        .chkInvocarSin.value = IIf(MapZonas(id).InvocarSinEfecto, vbChecked, vbUnchecked)
        .chkOcultarSin.value = IIf(MapZonas(id).OcultarSinEfecto, vbChecked, vbUnchecked)
        .chkMapResuSinEfecto.value = IIf(MapZonas(id).ResuSinEfecto, vbChecked, vbUnchecked)
        .txtMapVersion = MapZonas(id).MapVersion
        .ChkMapNpc.value = IIf(MapZonas(id).RoboNpcsPermitido, vbChecked, vbUnchecked)
        
        If MapZonas(id).LuzBase = 0 Then
            .chkLuzClimatica = vbUnchecked
        Else
            .chkLuzClimatica = vbChecked

        End If
        
        .PicColorMap.BackColor = MapZonas(id).LuzBase
        .LuzMapa.Text = tR & "-" & tG & "-" & tB

    End With
    
End Sub

'#######################################
'CARGAR DE MAPAS FORMATO WINTER ACTUAL
'#######################################

Sub Cargar_CSM(ByVal Map As String)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 14/03/2021
    'Descripcion: Carga los mapas de WinterAO
    '***************************************************
    
    On Error GoTo ErrorHandler
    
    Dim fh           As Integer

    Dim File         As Integer

    Dim MH           As tMapHeader

    Dim Blqs()       As tDatosBloqueados

    Dim L1()         As Long

    Dim L2()         As tDatosGrh

    Dim L3()         As tDatosGrh

    Dim L4()         As tDatosGrh

    Dim Triggers()   As tDatosTrigger

    Dim Particulas() As tDatosParticulas

    Dim Luces()      As tDatosLuces

    Dim Zonas()      As tDatosZonas

    Dim Objetos()    As tDatosObjs

    Dim NPCs()       As tDatosNPC

    Dim TEs()        As tDatosTE
    
    Dim LaCabecera   As tCabecera
    
    Dim i            As Long

    Dim j            As Long

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
                If Blqs(i).X > XMinMapSize And Blqs(i).X < XMaxMapSize And Blqs(i).Y > YMinMapSize And Blqs(i).Y < YMaxMapSize Then _
                    MapData(Blqs(i).X, Blqs(i).Y).bLocked = 1
            Next i

        End If
            
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2

            For i = 1 To .NumeroLayers(2)
                If L2(i).X > XMinMapSize And L2(i).X < XMaxMapSize And L2(i).Y > YMinMapSize And L2(i).Y < YMaxMapSize Then _
                    InitGrh MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).GrhIndex
            Next i

        End If
            
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3

            For i = 1 To .NumeroLayers(3)
                If L3(i).X > XMinMapSize And L3(i).X < XMaxMapSize And L3(i).Y > YMinMapSize And L3(i).Y < YMaxMapSize Then _
                    InitGrh MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).GrhIndex
            Next i

        End If
            
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4

            For i = 1 To .NumeroLayers(4)
                If L4(i).X > XMinMapSize And L4(i).X < XMaxMapSize And L4(i).Y > YMinMapSize And L4(i).Y < YMaxMapSize Then _
                    InitGrh MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).GrhIndex
            Next i

        End If
            
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers

            For i = 1 To .NumeroTriggers
                If Triggers(i).X > XMinMapSize And Triggers(i).X < XMaxMapSize And Triggers(i).Y > YMinMapSize And Triggers(i).Y < YMaxMapSize Then _
                    MapData(Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
            Next i

        End If
            
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas

            For i = 1 To .NumeroParticulas
                If Particulas(i).X > XMinMapSize And Particulas(i).X < XMaxMapSize And Particulas(i).Y > YMinMapSize And Particulas(i).Y < YMaxMapSize Then
                    MapData(Particulas(i).X, Particulas(i).Y).Particle_Index = Particulas(i).Particula
                    Call General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).Y)
                End If
            Next i

        End If

        '
        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)

            Dim p As Byte

            Get #fh, , Luces

            For i = 1 To .NumeroLuces

                If Luces(i).X > XMinMapSize And Luces(i).X < XMaxMapSize And Luces(i).Y > YMinMapSize And Luces(i).Y < YMaxMapSize Then
                    With MapData(Luces(i).X, Luces(i).Y)
                        .Light.range = Luces(i).range ' Changed by: Project Administrator at: 4/26/2021-20:26:18 on machine: DESKTOP-BR8H09Q
                        .Light.RGBCOLOR.a = 255
                        .Light.RGBCOLOR.R = Luces(i).R
                        .Light.RGBCOLOR.G = Luces(i).G
                        .Light.RGBCOLOR.B = Luces(i).B
    
                    End With
    
                    Call Create_Light_To_Map(Luces(i).X, Luces(i).Y, Luces(i).range, Luces(i).R, Luces(i).G, Luces(i).B)
                End If
            Next i

            Call LightRenderAll

        End If

        '
        If .NumeroZonas > 0 Then
            ReDim Zonas(1 To .NumeroZonas)
            Get #fh, , Zonas

            For i = 1 To .NumeroZonas
                If Zonas(i).X > XMinMapSize And Zonas(i).X < XMaxMapSize + 1 And Zonas(i).Y > YMinMapSize And Zonas(i).Y < YMaxMapSize + 1 Then _
                    MapData(Zonas(i).X, Zonas(i).Y).ZonaIndex = Zonas(i).Zona
            Next i

        End If
            
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos

            For i = 1 To .NumeroOBJs
                MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.ObjIndex = Objetos(i).ObjIndex
                MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.Amount = Objetos(i).ObjAmmount

                If MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.ObjIndex > NumObjs Then
                    InitGrh MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, 20299
                Else
                    InitGrh MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, ObjData(MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.ObjIndex).GrhIndex

                End If

            Next i

        End If
                
        If .NumeroNPCs > 0 Then
            ReDim NPCs(1 To .NumeroNPCs)
            Get #fh, , NPCs

            For i = 1 To .NumeroNPCs

                If NPCs(i).NPCIndex > 0 Then
                    MapData(NPCs(i).X, NPCs(i).Y).NPCIndex = NPCs(i).NPCIndex
                    Call MakeChar(NextOpenChar(), NpcData(NPCs(i).NPCIndex).Body, NpcData(NPCs(i).NPCIndex).Head, NpcData(NPCs(i).NPCIndex).Heading, NPCs(i).X, NPCs(i).Y)

                End If

            Next i

        End If
    
        If .NumeroTE > 0 Then
            ReDim TEs(1 To .NumeroTE)
            Get #fh, , TEs

            For i = 1 To .NumeroTE
                MapData(TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
                MapData(TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
                MapData(TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
            Next i

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
    
    'MapInfo_Cargar Map
    frmMapInfo.txtMapVersion.Text = MapInfo.MapVersion
    
    Call Pestanas(Map, ".csm")

    'Change mouse icon
    frmMain.MousePointer = 0
    
    ' Vacio deshacer
    '    modEdicion.Deshacer_Clear
    
    Call CSMInfoCargar
    
    '    'Set changed flag
    MapInfo.Changed = 0

    MapaCargado = True
    
    Call DibujarMinimapa
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Mapa " & Map & " cargado...", 0, 255, 0)

    Exit Sub

ErrorHandler:

    If fh <> 0 Then Close fh
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Error en el Mapa " & Map & ", se ha generado un informe de errores en: " & App.Path & "\Logs.txt", 255, 0, 0)
    
    File = FreeFile
    
    Call RegistrarError(Err.Number, Err.Description, "modMapas.Cargar_CSM", Erl)

End Sub

Public Sub CSMInfoCargar()
    '**********************************
    'Autor: Lorwik
    'Fecha: 14/03/2021
    'Descripcion: Cargar la informacion de los mapas de WinterAO.
    '**********************************

    Dim i  As Integer

    Dim tR As Byte

    Dim tG As Byte

    Dim tB As Byte

    CantZonas = UBound(MapDat)
    
    If CantZonas < 1 Then Exit Sub
    
    ReDim MapZonas(CantZonas) As tMapInfo
    
    frmZonas.LstZona.Clear

    For i = 1 To CantZonas

        With MapZonas(i)
        
            .name = MapDat(i).map_name
            .Music = MapDat(i).music_number
        
            .MagiaSinEfecto = MapDat(i).MagiaSinEfecto
            .InviSinEfecto = MapDat(i).InviSinEfecto
            .ResuSinEfecto = MapDat(i).ResuSinEfecto
            .RoboNpcsPermitido = MapDat(i).RoboNpcsPermitido
            .InvocarSinEfecto = MapDat(i).InvocarSinEfecto
            .OcultarSinEfecto = MapDat(i).OcultarSinEfecto
            
            .lvlMinimo = Val(MapDat(i).lvlMinimo)
            .LuzBase = MapDat(i).LuzBase
            
            If MapDat(i).LuzBase <> 0 Then
                frmMapInfo.chkLuzClimatica = Checked
                Call ConvertLongToRGB(MapDat(i).LuzBase, tR, tG, tB)
                
                Estado_Custom.a = 255
                Estado_Custom.R = tR
                Estado_Custom.G = tG
                Estado_Custom.B = tB
                
                Call Actualizar_Estado
                
                frmMapInfo.LuzMapa.Text = tR & "-" & tG & "-" & tB
                frmMapInfo.PicColorMap.BackColor = .LuzBase
                
            Else
                frmMapInfo.chkLuzClimatica = Unchecked
                
            End If
            
            .MapVersion = MapDat(i).version
            
            If MapDat(i).battle_mode = True Then
                .PK = True
            Else
                .PK = False

            End If
            
            .ambient = MapDat(i).ambient
            
            .Terreno = MapDat(i).terrain
            .Zona = MapDat(i).zone
            .Restringir = MapDat(i).restrict_mode
            .BackUp = MapDat(i).backup_mode
        
            frmZonas.LstZona.AddItem (i & "- " & .name)
        
        End With

    Next i
    
    Call MapInfo_Actualizar

End Sub

'#######################################
'GUARDADO DE MAPAS FORMATO WINTER ACTUAL
'#######################################

Public Function Save_CSM(ByVal MapRoute As String, _
                         Optional ByVal Fusion As Boolean = False) As Boolean

    On Error GoTo ErrorHandler

    Dim fh           As Integer

    Dim MH           As tMapHeader

    Dim Blqs()       As tDatosBloqueados

    Dim L1()         As Long

    Dim L2()         As tDatosGrh

    Dim L3()         As tDatosGrh

    Dim L4()         As tDatosGrh

    Dim Triggers()   As tDatosTrigger

    Dim Luces()      As tDatosLuces

    Dim Particulas() As tDatosParticulas

    Dim Zonas()      As tDatosZonas

    Dim Objetos()    As tDatosObjs

    Dim NPCs()       As tDatosNPC

    Dim TEs()        As tDatosTE
        
    Dim i            As Integer

    Dim j            As Integer
    
    If NoSobreescribir = False Then
        If FileExist(MapRoute, vbNormal) = True Then
            If MsgBox("¿Desea sobrescribir " & MapRoute & "?", vbCritical + vbYesNo) = vbNo Then
                Exit Function
                
            Else
                'Kill MapRoute
                
            End If

        End If

    End If
    
    frmMain.MousePointer = 11
    MapSize.XMax = XMaxMapSize
    MapSize.XMin = XMinMapSize
    MapSize.YMax = YMaxMapSize
    MapSize.YMin = YMinMapSize
    
    If Fusion Then
        MapSize.XMax = 1100
        MapSize.YMax = 1100

    End If
    
    ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax)
    
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            
            If Fusion Then

                With SuperMapData(i, j)

                    If .bLocked Then
                        MH.NumeroBloqueados = MH.NumeroBloqueados + 1
                        ReDim Preserve Blqs(1 To MH.NumeroBloqueados)
                        Blqs(MH.NumeroBloqueados).X = i
                        Blqs(MH.NumeroBloqueados).Y = j

                    End If
                    
                    L1(i, j) = .Graphic(1).GrhIndex
                    
                    If .Graphic(2).GrhIndex > 0 Then
                        MH.NumeroLayers(2) = MH.NumeroLayers(2) + 1
                        ReDim Preserve L2(1 To MH.NumeroLayers(2))
                        L2(MH.NumeroLayers(2)).X = i
                        L2(MH.NumeroLayers(2)).Y = j
                        L2(MH.NumeroLayers(2)).GrhIndex = .Graphic(2).GrhIndex

                    End If
                    
                    If .Graphic(3).GrhIndex > 0 Then
                        MH.NumeroLayers(3) = MH.NumeroLayers(3) + 1
                        ReDim Preserve L3(1 To MH.NumeroLayers(3))
                        L3(MH.NumeroLayers(3)).X = i
                        L3(MH.NumeroLayers(3)).Y = j
                        L3(MH.NumeroLayers(3)).GrhIndex = .Graphic(3).GrhIndex

                    End If
                    
                    If .Graphic(4).GrhIndex > 0 Then
                        MH.NumeroLayers(4) = MH.NumeroLayers(4) + 1
                        ReDim Preserve L4(1 To MH.NumeroLayers(4))
                        L4(MH.NumeroLayers(4)).X = i
                        L4(MH.NumeroLayers(4)).Y = j
                        L4(MH.NumeroLayers(4)).GrhIndex = .Graphic(4).GrhIndex

                    End If
                    
                    If .Trigger > 0 Then
                        MH.NumeroTriggers = MH.NumeroTriggers + 1
                        ReDim Preserve Triggers(1 To MH.NumeroTriggers)
                        Triggers(MH.NumeroTriggers).X = i
                        Triggers(MH.NumeroTriggers).Y = j
                        Triggers(MH.NumeroTriggers).Trigger = .Trigger

                    End If
                    
                    If .Particle_Index > 0 Then
                        MH.NumeroParticulas = MH.NumeroParticulas + 1
                        ReDim Preserve Particulas(1 To MH.NumeroParticulas)
                        Particulas(MH.NumeroParticulas).X = i
                        Particulas(MH.NumeroParticulas).Y = j
                        Particulas(MH.NumeroParticulas).Particula = .Particle_Index
    
                    End If
                   
                    '¿Hay luz activa en este punto?
                    If .Light.range > 0 Then
                        MH.NumeroLuces = MH.NumeroLuces + 1
                        ReDim Preserve Luces(1 To MH.NumeroLuces)
                        
                        Luces(MH.NumeroLuces).R = .Light.RGBCOLOR.R
                        Luces(MH.NumeroLuces).G = .Light.RGBCOLOR.G
                        Luces(MH.NumeroLuces).B = .Light.RGBCOLOR.B
                        Luces(MH.NumeroLuces).range = .Light.range
                        Luces(MH.NumeroLuces).X = i
                        Luces(MH.NumeroLuces).Y = j

                    End If
                    
                    If .ZonaIndex > 0 Then
                        MH.NumeroZonas = MH.NumeroZonas + 1
                        ReDim Preserve Zonas(1 To MH.NumeroZonas)
                        Zonas(MH.NumeroZonas).X = i
                        Zonas(MH.NumeroZonas).Y = j
                        Zonas(MH.NumeroZonas).Zona = .ZonaIndex

                    End If
                    
                    If .OBJInfo.ObjIndex > 0 Then
                        MH.NumeroOBJs = MH.NumeroOBJs + 1
                        ReDim Preserve Objetos(1 To MH.NumeroOBJs)
                        Objetos(MH.NumeroOBJs).ObjIndex = .OBJInfo.ObjIndex
                        Objetos(MH.NumeroOBJs).ObjAmmount = .OBJInfo.Amount
                        Objetos(MH.NumeroOBJs).X = i
                        Objetos(MH.NumeroOBJs).Y = j

                    End If
                    
                    If .NPCIndex > 0 Then
                        MH.NumeroNPCs = MH.NumeroNPCs + 1
                        ReDim Preserve NPCs(1 To MH.NumeroNPCs)
                        NPCs(MH.NumeroNPCs).NPCIndex = .NPCIndex
                        NPCs(MH.NumeroNPCs).X = i
                        NPCs(MH.NumeroNPCs).Y = j

                    End If
                    
                    If .TileExit.Map > 0 Then
                        MH.NumeroTE = MH.NumeroTE + 1
                        ReDim Preserve TEs(1 To MH.NumeroTE)
                        TEs(MH.NumeroTE).DestM = .TileExit.Map
                        TEs(MH.NumeroTE).DestX = .TileExit.X
                        TEs(MH.NumeroTE).DestY = .TileExit.Y
                        TEs(MH.NumeroTE).X = i
                        TEs(MH.NumeroTE).Y = j

                    End If
                    
                End With
                
            Else

                With MapData(i, j)

                    If .bLocked Then
                        MH.NumeroBloqueados = MH.NumeroBloqueados + 1
                        ReDim Preserve Blqs(1 To MH.NumeroBloqueados)
                        Blqs(MH.NumeroBloqueados).X = i
                        Blqs(MH.NumeroBloqueados).Y = j

                    End If
                    
                    L1(i, j) = .Graphic(1).GrhIndex
                    
                    If .Graphic(2).GrhIndex > 0 Then
                        MH.NumeroLayers(2) = MH.NumeroLayers(2) + 1
                        ReDim Preserve L2(1 To MH.NumeroLayers(2))
                        L2(MH.NumeroLayers(2)).X = i
                        L2(MH.NumeroLayers(2)).Y = j
                        L2(MH.NumeroLayers(2)).GrhIndex = .Graphic(2).GrhIndex

                    End If
                    
                    If .Graphic(3).GrhIndex > 0 Then
                        MH.NumeroLayers(3) = MH.NumeroLayers(3) + 1
                        ReDim Preserve L3(1 To MH.NumeroLayers(3))
                        L3(MH.NumeroLayers(3)).X = i
                        L3(MH.NumeroLayers(3)).Y = j
                        L3(MH.NumeroLayers(3)).GrhIndex = .Graphic(3).GrhIndex

                    End If
                    
                    If .Graphic(4).GrhIndex > 0 Then
                        MH.NumeroLayers(4) = MH.NumeroLayers(4) + 1
                        ReDim Preserve L4(1 To MH.NumeroLayers(4))
                        L4(MH.NumeroLayers(4)).X = i
                        L4(MH.NumeroLayers(4)).Y = j
                        L4(MH.NumeroLayers(4)).GrhIndex = .Graphic(4).GrhIndex

                    End If
                    
                    If .Trigger > 0 Then
                        MH.NumeroTriggers = MH.NumeroTriggers + 1
                        ReDim Preserve Triggers(1 To MH.NumeroTriggers)
                        Triggers(MH.NumeroTriggers).X = i
                        Triggers(MH.NumeroTriggers).Y = j
                        Triggers(MH.NumeroTriggers).Trigger = .Trigger

                    End If
                    
                    If .Particle_Index > 0 Then
                        MH.NumeroParticulas = MH.NumeroParticulas + 1
                        ReDim Preserve Particulas(1 To MH.NumeroParticulas)
                        Particulas(MH.NumeroParticulas).X = i
                        Particulas(MH.NumeroParticulas).Y = j
                        Particulas(MH.NumeroParticulas).Particula = .Particle_Index
                        Debug.Print .Particle_Index
    
                    End If
                   
                    '¿Hay luz activa en este punto?
                    If .Light.range > 0 Then
                        MH.NumeroLuces = MH.NumeroLuces + 1
                        ReDim Preserve Luces(1 To MH.NumeroLuces)
                        
                        Luces(MH.NumeroLuces).R = .Light.RGBCOLOR.R
                        Luces(MH.NumeroLuces).G = .Light.RGBCOLOR.G
                        Luces(MH.NumeroLuces).B = .Light.RGBCOLOR.B
                        Luces(MH.NumeroLuces).range = .Light.range
                        Luces(MH.NumeroLuces).X = i
                        Luces(MH.NumeroLuces).Y = j

                    End If
                    
                    If .ZonaIndex > 0 And .ZonaIndex <= CantZonas Then
                    
                        MH.NumeroZonas = MH.NumeroZonas + 1
                        ReDim Preserve Zonas(1 To MH.NumeroZonas)
                        Zonas(MH.NumeroZonas).X = i
                        Zonas(MH.NumeroZonas).Y = j
                        Zonas(MH.NumeroZonas).Zona = .ZonaIndex

                    End If

                    If .OBJInfo.ObjIndex > 0 Then
                        MH.NumeroOBJs = MH.NumeroOBJs + 1
                        ReDim Preserve Objetos(1 To MH.NumeroOBJs)
                        Objetos(MH.NumeroOBJs).ObjIndex = .OBJInfo.ObjIndex
                        Objetos(MH.NumeroOBJs).ObjAmmount = .OBJInfo.Amount
                        Objetos(MH.NumeroOBJs).X = i
                        Objetos(MH.NumeroOBJs).Y = j

                    End If
                    
                    If .NPCIndex > 0 Then
                        MH.NumeroNPCs = MH.NumeroNPCs + 1
                        ReDim Preserve NPCs(1 To MH.NumeroNPCs)
                        NPCs(MH.NumeroNPCs).NPCIndex = .NPCIndex
                        NPCs(MH.NumeroNPCs).X = i
                        NPCs(MH.NumeroNPCs).Y = j

                    End If
                    
                    If .TileExit.Map > 0 Then
                        MH.NumeroTE = MH.NumeroTE + 1
                        ReDim Preserve TEs(1 To MH.NumeroTE)
                        TEs(MH.NumeroTE).DestM = .TileExit.Map
                        TEs(MH.NumeroTE).DestX = .TileExit.X
                        TEs(MH.NumeroTE).DestY = .TileExit.Y
                        TEs(MH.NumeroTE).X = i
                        TEs(MH.NumeroTE).Y = j

                    End If
                    
                End With
            
            End If

        Next i
    Next j
 
    MH.NumeroData = CantZonas
    
    Call CSMInfoSave
              
    fh = FreeFile
    Open MapRoute For Binary As fh
        
    Put #fh, , MiCabecera
        
    Put #fh, , MH
    Put #fh, , MapSize
    Put #fh, , MapDat
    Put #fh, , L1
    
    With MH

        If .NumeroBloqueados > 0 Then Put #fh, , Blqs

        If .NumeroLayers(2) > 0 Then Put #fh, , L2

        If .NumeroLayers(3) > 0 Then Put #fh, , L3

        If .NumeroLayers(4) > 0 Then Put #fh, , L4

        If .NumeroTriggers > 0 Then Put #fh, , Triggers

        If .NumeroParticulas > 0 Then Put #fh, , Particulas

        If .NumeroLuces > 0 Then Put #fh, , Luces

        If .NumeroZonas > 0 Then Put #fh, , Zonas

        If .NumeroOBJs > 0 Then Put #fh, , Objetos

        If .NumeroNPCs > 0 Then Put #fh, , NPCs

        If .NumeroTE > 0 Then Put #fh, , TEs

    End With
    
    Close fh
    
    Call Pestanas(MapRoute, ".csm")
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
    NoSobreescribir = False
    
    Save_CSM = True
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Mapa " & MapRoute & " guardado...", 0, 255, 0)
    Exit Function

ErrorHandler:

    If fh <> 0 Then Close fh

End Function

Private Sub CSMInfoSave()
    '**********************************
    'Autor: Lorwik
    'Fecha: 14/03/2021
    'Descripcion: Guarda la informacion de los mapas de WinterAO.
    '**********************************

    Dim i As Integer

    ReDim MapDat(CantZonas) As tMapDat

    For i = 1 To CantZonas

        With MapZonas(i)

            MapDat(i).map_name = .name
            MapDat(i).music_number = .Music
            
            MapDat(i).MagiaSinEfecto = .MagiaSinEfecto
            MapDat(i).InviSinEfecto = .InviSinEfecto
            MapDat(i).ResuSinEfecto = .ResuSinEfecto
            MapDat(i).LuzBase = .LuzBase
            MapDat(i).RoboNpcsPermitido = .RoboNpcsPermitido
            MapDat(i).OcultarSinEfecto = .OcultarSinEfecto
            MapDat(i).InvocarSinEfecto = .InvocarSinEfecto
            
            MapDat(i).lvlMinimo = .lvlMinimo
            
            If frmMapInfo.chkLuzClimatica = Checked Then
                MapDat(i).LuzBase = .LuzBase
            Else
                MapDat(i).LuzBase = 0

            End If
            
            MapDat(i).version = .MapVersion
            
            If .PK = True Then
                MapDat(i).battle_mode = True
            Else
                MapDat(i).battle_mode = False

            End If
            
            MapDat(i).ambient = .ambient
            MapDat(i).terrain = .Terreno
            MapDat(i).zone = .Zona
            MapDat(i).restrict_mode = .Restringir
            MapDat(i).backup_mode = .BackUp

        End With
        
    Next i
    
End Sub

'#######################################
'CARGA DE MAPAS FORMATO ARGENTUM
'#######################################

Public Sub Cargar_Map(ByVal Map As String, Optional ByVal EsInteger As Boolean = False)
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************

    On Error Resume Next

    Dim LoopC       As Integer

    Dim TempInt     As Integer

    Dim Body        As Integer

    Dim Head        As Integer

    Dim Heading     As Byte

    Dim Y           As Integer

    Dim X           As Integer

    Dim i           As Byte

    Dim ByFlags     As Byte

    Dim FreeFileMap As Long

    Dim FreeFileInf As Long

    DoEvents
    
    'Change mouse icon
    frmMain.MousePointer = 11
       
    'Con esto, le digo al WE que estamos usando mapas de tipo integer,
    'lo uso mas que nada para que no crashee cargar los mapas siguientes en las Pestañas.
    If EsInteger Then
        ClientSetup.TipoMapaCargado = eTipoMapa.tInt
        
    Else
        ClientSetup.TipoMapaCargado = eTipoMapa.tLong
        
    End If
    
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    Map = Left$(Map, Len(Map) - 4)
    Map = Map & ".inf"
    
    FreeFileInf = FreeFile
    Open Map For Binary As FreeFileInf
    Seek FreeFileInf, 1
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    
    'Cabecera inf
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt

    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            With MapData(X, Y)
            
                Get FreeFileMap, , ByFlags
                .bLocked = (ByFlags And 1)
            
                'Layer 1
                If EsInteger Then
                    Get FreeFileMap, , .Graphic(1).GrhIndexInt
                    Call InitGrh(.Graphic(1), .Graphic(1).GrhIndexInt)
                    
                Else
                    Get FreeFileMap, , .Graphic(1).GrhIndex
                    Call InitGrh(.Graphic(1), .Graphic(1).GrhIndex)
                    
                End If
            
                'Layer 2 used?
                If ByFlags And 2 Then
                    
                    If EsInteger Then
                        Get FreeFileMap, , .Graphic(2).GrhIndexInt
                        Call InitGrh(.Graphic(2), .Graphic(2).GrhIndexInt)
                        
                    Else
                        Get FreeFileMap, , .Graphic(2).GrhIndex
                        Call InitGrh(.Graphic(2), .Graphic(2).GrhIndex)
                        
                    End If
 
                Else
                
                    .Graphic(2).GrhIndex = 0
                    
                End If
                
                'Layer 3 used?
                If ByFlags And 4 Then
                    
                    If EsInteger Then
                        Get FreeFileMap, , .Graphic(3).GrhIndexInt
                        Call InitGrh(.Graphic(3), .Graphic(3).GrhIndexInt)
                        
                    Else
                        Get FreeFileMap, , .Graphic(3).GrhIndex
                        Call InitGrh(.Graphic(3), .Graphic(3).GrhIndex)
                        
                    End If

                Else
                
                    .Graphic(3).GrhIndex = 0
                    
                End If
                
                'Layer 4 used?
                If ByFlags And 8 Then
                    
                    If EsInteger Then
                        Get FreeFileMap, , .Graphic(4).GrhIndexInt
                        Call InitGrh(.Graphic(3), .Graphic(3).GrhIndexInt)
                        
                    Else
                        Get FreeFileMap, , .Graphic(4).GrhIndex
                        Call InitGrh(.Graphic(4), .Graphic(4).GrhIndex)
                        
                    End If

                Else
                    
                    .Graphic(4).GrhIndex = 0

                End If
             
                'Trigger used?
                If ByFlags And 16 Then
                    Get FreeFileMap, , .Trigger
                Else
                    .Trigger = 0

                End If
            
                'Cargamos el archivo ".INF"
                Get FreeFileInf, , ByFlags
            
                If ByFlags And 1 Then
                    
                    With .TileExit
                    
                        Get FreeFileInf, , .Map
                        Get FreeFileInf, , .X
                        Get FreeFileInf, , .Y
                    
                    End With

                End If
    
                If ByFlags And 2 Then
                
                    'Get and make NPC
                    Get FreeFileInf, , .NPCIndex
    
                    If .NPCIndex < 0 Then
                        .NPCIndex = 0
                        
                    Else
                        Body = NpcData(.NPCIndex).Body
                        Head = NpcData(.NPCIndex).Head
                        Heading = NpcData(.NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)
                        
                    End If

                End If
    
                If ByFlags And 4 Then
                    
                    'Get and make Object
                    Get FreeFileInf, , .OBJInfo.ObjIndex
                    Get FreeFileInf, , .OBJInfo.Amount

                    If .OBJInfo.ObjIndex > 0 Then Call InitGrh(.ObjGrh, ObjData(.OBJInfo.ObjIndex).GrhIndex)

                End If
            
            End With
    
        Next X
    Next Y
    
    'Close files
    Close FreeFileMap
    Close FreeFileInf
    
    Call Pestanas(Map, ".map")
    
    Map = Left$(Map, Len(Map) - 4) & ".dat"
    
    Call MapInfo_Cargar(Map)
    
    With frmMain
    
        frmMapInfo.txtMapVersion.Text = MapInfo.MapVersion
        
        'Set changed flag
        MapInfo.Changed = 0
        
        'Change mouse icon
        .MousePointer = 0
    
    End With
    
    MapaCargado = True

End Sub

Public Sub MapInfo_Cargar(ByVal Archivo As String)
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************

    On Error GoTo MapInfo_Cargar_Err
    
    Dim Leer      As New clsIniManager

    Dim LoopC     As Integer

    Dim Path      As String

    Dim MapTitulo As String
    
    MapTitulo = Empty
    Leer.Initialize Archivo

    For LoopC = Len(Archivo) To 1 Step -1

        If mid(Archivo, LoopC, 1) = "\" Then
            Path = Left(Archivo, LoopC)
            Exit For

        End If

    Next
    Archivo = Right(Archivo, Len(Archivo) - (Len(Path)))
    MapTitulo = UCase(Left(Archivo, Len(Archivo) - 4))

    MapInfo.name = Leer.GetValue(MapTitulo, "Name")
    MapInfo.Music = Leer.GetValue(MapTitulo, "MusicNum")
    MapInfo.MagiaSinEfecto = Val(Leer.GetValue(MapTitulo, "MagiaSinEfecto"))
    MapInfo.InviSinEfecto = Val(Leer.GetValue(MapTitulo, "InviSinEfecto"))
    MapInfo.ResuSinEfecto = Val(Leer.GetValue(MapTitulo, "ResuSinEfecto"))
    MapInfo.NoEncriptarMP = Val(Leer.GetValue(MapTitulo, "NoEncriptarMP"))
    
    If Val(Leer.GetValue(MapTitulo, "Pk")) = 0 Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False

    End If
    
    MapInfo.Terreno = Leer.GetValue(MapTitulo, "Terreno")
    MapInfo.Zona = Leer.GetValue(MapTitulo, "Zona")
    MapInfo.Restringir = Leer.GetValue(MapTitulo, "Restringir")
    MapInfo.BackUp = Val(Leer.GetValue(MapTitulo, "BACKUP"))
    
    Call MapInfo_Actualizar
    
    Set Leer = Nothing
    Exit Sub
    
MapInfo_Cargar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapas.MapInfo_Cargar", Erl)

    Resume Next
    
End Sub

'#######################################
'GUARDADO DE MAPAS FORMATO ARGENTUM
'#######################################

Public Sub Guardar_Map(ByVal SaveAs As String)
    '*************************************************
    'Author:Lorwik
    'Last modified: 01/05/2021
    '*************************************************

    On Error GoTo ErrorSave

    Dim FreeFileMap As Long

    Dim FreeFileInf As Long

    Dim LoopC       As Long

    Dim TempInt     As Integer

    Dim Y           As Long

    Dim X           As Long

    Dim ByFlags     As Byte

    If FileExist(SaveAs, vbNormal) = True Then
        
        If NoSobreescribir = False Then
            If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            Else
                Call Kill(SaveAs)

            End If
        
        Else
            Call Kill(SaveAs)
            
        End If
        
    End If

    frmMain.MousePointer = 11

    ' y borramos el .inf tambien
    If FileExist(Left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Call Kill(Left$(SaveAs, Len(SaveAs) - 4) & ".inf")

    End If

    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1

    SaveAs = Left$(SaveAs, Len(SaveAs) - 4)
    SaveAs = SaveAs & ".inf"

    'Open .inf file
    FreeFileInf = FreeFile
    Open SaveAs For Binary As FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    
    ' Version del Mapa
    If frmMapInfo.txtMapVersion.Text < 32767 Then
        frmMapInfo.txtMapVersion.Text = frmMapInfo.txtMapVersion + 1

    End If

    Put FreeFileMap, , CInt(frmMapInfo.txtMapVersion.Text)
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            With MapData(X, Y)
            
                ByFlags = 0
                
                If .bLocked = 1 Then ByFlags = ByFlags Or 1
                
                If .Graphic(2).GrhIndex Then ByFlags = ByFlags Or 2
                If .Graphic(3).GrhIndex Then ByFlags = ByFlags Or 4
                If .Graphic(4).GrhIndex Then ByFlags = ByFlags Or 8

                If .Trigger Then ByFlags = ByFlags Or 16
                    
                Put FreeFileMap, , ByFlags
                    
                If ClientSetup.TipoMapaCargado = eTipoMapa.tInt Then
                    Put FreeFileMap, , .Graphic(1).GrhIndexInt
                    
                Else
                    Put FreeFileMap, , .Graphic(1).GrhIndex

                End If
                
                For LoopC = 2 To 4
                    
                    If ClientSetup.TipoMapaCargado = eTipoMapa.tInt Then
                        If .Graphic(LoopC).GrhIndex Then Put FreeFileMap, , .Graphic(LoopC).GrhIndexInt
                    Else

                        If .Graphic(LoopC).GrhIndex Then Put FreeFileMap, , .Graphic(LoopC).GrhIndex

                    End If

                Next LoopC
                    
                If .Trigger Then Put FreeFileMap, , .Trigger
                
                'Escribimos el archivo ".INF"
                ByFlags = 0
                    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                If .NPCIndex Then ByFlags = ByFlags Or 2
                
                If .OBJInfo.ObjIndex Then ByFlags = ByFlags Or 4
                    
                Put FreeFileInf, , ByFlags
                    
                If .TileExit.Map Then
                    Put FreeFileInf, , .TileExit.Map
                    Put FreeFileInf, , .TileExit.X
                    Put FreeFileInf, , .TileExit.Y

                End If
                    
                If .NPCIndex Then Put FreeFileInf, , CInt(.NPCIndex)
                    
                If .OBJInfo.ObjIndex Then
                    Put FreeFileInf, , .OBJInfo.ObjIndex
                    Put FreeFileInf, , .OBJInfo.Amount

                End If
            
            End With
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close FreeFileInf

    Call Pestanas(SaveAs, ".map")

    'write .dat file
    SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
    Call MapInfo_Guardar(SaveAs)

    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
    NoSobreescribir = False

    Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & Err.Number & " - " & Err.Description

End Sub

'#######################################
'GUARDADO DE MAPAS FORMATO IMPERIUM CLASICO
'#######################################

Sub Cargar_MapImpClasico(ByVal Map As String, ByVal Tipo As eTipoMapa)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 14/03/2021
    'Descripcion: Carga los mapas de ImperiumAO de la 1.3 y 1.4
    '***************************************************
    
    On Error GoTo ErrorHandler
    
    Dim fh              As Integer
    Dim File            As Integer
    Dim MH              As tMapHeader
    Dim Blqs()          As tDatosBloqueados
    Dim L1()            As Integer
    Dim Ln1()           As Long
    Dim L2()            As tDatosGrh
    Dim L3()            As tDatosGrh
    Dim L4()            As tDatosGrh
    Dim Triggers()      As tDatosTrigger
    Dim Luces()         As tDatosLucesIAC
    Dim Particulas()    As tDatosParticulas
    Dim Objetos()       As tDatosObjs
    Dim NPCs()          As tDatosNPC
    Dim TEs()           As tDatosTE
    
    Dim i               As Long
    Dim j               As Long
    Dim tR              As Byte
    Dim tG              As Byte
    Dim tB              As Byte
    
    fh = FreeFile
    
    Open Map For Binary Access Read As fh
    
        Get #fh, , MH
        Get #fh, , MapSize
        
        If Tipo = eTipoMapa.tIAOnew Then
            TipoMapaCargado = eTipoMapa.tIAOnew
            Call CaptionWorldEditor(Map, False, "ImperiumAO 1.4")
            
            Get #fh, , MapDatIAO
            ReDim Ln1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
            Get #fh, , Ln1
            
        ElseIf Tipo = eTipoMapa.tIAOold Then
            TipoMapaCargado = eTipoMapa.tIAOold
            Call CaptionWorldEditor(Map, False, "ImperiumAO 1.3")
            
            Get #fh, , MapDatIAOOld
            ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Integer
            Get #fh, , L1
            
        Else
            Exit Sub
            
        End If
        
        With MH
            If .NumeroBloqueados > 0 Then
                ReDim Blqs(1 To .NumeroBloqueados)
                Get #fh, , Blqs
                For i = 1 To .NumeroBloqueados
                    MapData(Blqs(i).X, Blqs(i).Y).bLocked = 1
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
                For i = 1 To .NumeroParticulas
                    MapData(Particulas(i).X, Particulas(i).Y).Particle_Index = Particulas(i).Particula
                    Call General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).Y)
                    
                    'MapData(Particulas(i).X, Particulas(i).y).Particle_Group_Index = General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).y)
                Next i
            End If
                
            If .NumeroLuces > 0 Then
                ReDim Luces(1 To .NumeroLuces)
                Dim p As Byte
                Get #fh, , Luces
                For i = 1 To .NumeroLuces
                
                    MapData(Luces(i).X, Luces(i).Y).Light.range = Luces(i).Rango
                    
                    Call ConvertLongToRGB(Luces(i).color, tR, tG, tB)
                
                    Call Create_Light_To_Map(Luces(i).X, Luces(i).Y, Luces(i).Rango, tR, tG, tB)
                Next i
                
                Call LightRenderAll
            End If
                
            If .NumeroOBJs > 0 Then
                ReDim Objetos(1 To .NumeroOBJs)
                Get #fh, , Objetos
                For i = 1 To .NumeroOBJs
                    MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.ObjIndex = Objetos(i).ObjIndex
                    MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.Amount = Objetos(i).ObjAmmount
                    If MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.ObjIndex > NumObjs Then
                        InitGrh MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, 20299
                    Else
                        InitGrh MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, ObjData(MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.ObjIndex).GrhIndex
                    End If
                Next i
            End If
                
            If .NumeroNPCs > 0 Then
                ReDim NPCs(1 To .NumeroNPCs)
                Get #fh, , NPCs
                For i = 1 To .NumeroNPCs
                    If NPCs(i).NPCIndex > 0 Then
                        MapData(NPCs(i).X, NPCs(i).Y).NPCIndex = NPCs(i).NPCIndex
                        Call MakeChar(NextOpenChar(), NpcData(NPCs(i).NPCIndex).Body, NpcData(NPCs(i).NPCIndex).Head, NpcData(NPCs(i).NPCIndex).Heading, NPCs(i).X, NPCs(i).Y)
                    End If
                Next i
            End If
    
            If .NumeroTE > 0 Then
                ReDim TEs(1 To .NumeroTE)
                Get #fh, , TEs
                For i = 1 To .NumeroTE
                    MapData(TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
                    MapData(TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
                    MapData(TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
                Next i
            End If
            
        End With
    
    Close fh
    
    
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            If Tipo = eTipoMapa.tIAOnew Then
                If Ln1(i, j) > 0 Then
                    InitGrh MapData(i, j).Graphic(1), Ln1(i, j)
                End If
                
            ElseIf Tipo = eTipoMapa.tIAOold Then
                If L1(i, j) > 0 Then
                    InitGrh MapData(i, j).Graphic(1), L1(i, j)
                End If
                
            End If
        Next i
    Next j
    
    'MapInfo_Cargar Map
    frmMapInfo.txtMapVersion.Text = MapInfo.MapVersion
    
    Call Pestanas(Map, ".csm")

    ' Vacia el Deshacer
    modEdicion.Deshacer_Clear
    
    'Change mouse icon
    frmMain.MousePointer = 0
    
    Call CSMInfoCargarIAO
    
    'Set changed flag
    MapInfo.Changed = 0

    MapaCargado = True
    
    Call DibujarMinimapa ' Radar
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa " & Map & " cargado...", 0, 255, 0)
ErrorHandler:

    If fh <> 0 Then Close fh
    Call AddtoRichTextBox(frmMain.StatTxt, "Error en el Mapa " & Map & ", se ha generado un informe de errores en: " & App.Path & "\Logs.txt", 255, 0, 0)
    File = FreeFile
    Open App.Path & "\Logs.txt" For Output As #File
        Print #File, Err.Description
    Close #File
    
End Sub

Public Sub MapInfo_Guardar(ByVal Archivo As String)
'*************************************************
'Author: Lorwik
'Last modified: 01/05/2021
'Guardar Informacion del Mapa (.dat)
'*************************************************

    Dim MapTitulo As String

    If LenB(MapTitulo) = 0 Then
        MapTitulo = NameMap_Save
    End If

    Call WriteVar(Archivo, MapTitulo, "Name", MapInfo.name)
    Call WriteVar(Archivo, MapTitulo, "MusicNum", MapInfo.Music)
    Call WriteVar(Archivo, MapTitulo, "MagiaSinefecto", Val(MapInfo.MagiaSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "InviSinEfecto", Val(MapInfo.InviSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "ResuSinEfecto", Val(MapInfo.ResuSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "NoEncriptarMP", Val(MapInfo.NoEncriptarMP))

    Call WriteVar(Archivo, MapTitulo, "Terreno", MapInfo.Terreno)
    Call WriteVar(Archivo, MapTitulo, "Zona", MapInfo.Zona)
    Call WriteVar(Archivo, MapTitulo, "Restringir", MapInfo.Restringir)
    Call WriteVar(Archivo, MapTitulo, "BackUp", str(MapInfo.BackUp))

    If MapInfo.PK Then
        Call WriteVar(Archivo, MapTitulo, "Pk", "0")
        
    Else
        Call WriteVar(Archivo, MapTitulo, "Pk", "1")
        
    End If
End Sub

Public Sub Pestanas(ByVal Map As String, Optional ByVal MapFormat As String = ".map")

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    'Calcula la orden de Pestanas
    '*************************************************
    On Error Resume Next

    Dim LoopC As Integer
    
    For LoopC = Len(Map) To 1 Step -1

        If mid(Map, LoopC, 1) = "\" Then
            PATH_Save = Left(Map, LoopC)
            Exit For

        End If

    Next
    
    Map = Right(Map, Len(Map) - (Len(PATH_Save)))
    
    MapaActual = ReadField(1, Right(Map, Len(Map) - 4), Asc("."))
    'If frmCopiarBordes.Visible Then Call frmCopiarBordes.Inicializar
    
    For LoopC = Len(Left(Map, Len(Map) - 4)) To 1 Step -1

        If IsNumeric(mid(Left(Map, Len(Map) - 4), LoopC, 1)) = False Then
            NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - LoopC)
            NameMap_Save = Left(Map, LoopC)
            Exit For

        End If

    Next
    
    For LoopC = (NumMap_Save - 4) To (NumMap_Save + 8)

        If FileExist(PATH_Save & NameMap_Save & LoopC & MapFormat, vbArchive) = True Then
            frmMain.MapPest(LoopC - NumMap_Save + 4).Visible = True
            frmMain.MapPest(LoopC - NumMap_Save + 4).Enabled = True
            frmMain.MapPest(LoopC - NumMap_Save + 4).Caption = NameMap_Save & LoopC
        Else
            frmMain.MapPest(LoopC - NumMap_Save + 4).Visible = False

        End If

    Next
    
End Sub
