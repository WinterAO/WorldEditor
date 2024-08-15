Attribute VB_Name = "modMapasWAO"
Option Explicit
'/////////////////////////////////////////////////////////////////////
'Lectura, guardado y otras features del formato de mapas WinterAO
'/////////////////////////////////////////////////////////////////////

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

    x As Integer
    y As Integer

End Type

Private Type tDatosGrh

    x As Integer
    y As Integer
    GrhIndex As Long

End Type

Private Type tDatosTrigger

    x As Integer
    y As Integer
    Trigger As Integer

End Type

Private Type tDatosZonas

    x As Integer
    y As Integer
    Zona As Integer

End Type

Public Type tDatosLuces

    R As Integer
    G As Integer
    B As Integer
    range As Byte
    x As Integer
    y As Integer

End Type

Private Type tDatosParticulas

    x As Integer
    y As Integer
    Particula As Long

End Type

Private Type tDatosNPC

    x As Integer
    y As Integer
    NPCIndex As Integer

End Type

Private Type tDatosObjs

    x As Integer
    y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer

End Type

Private Type tDatosTE

    x As Integer
    y As Integer
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
    music_number As Integer
    zone As String
    terrain As String
    Ambient As Integer
    lvlMinimo As Integer
    lvlMaximo As Integer
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
    
    Dim i            As Long

    Dim j            As Long

    DoEvents
        
    'Change mouse icon
    frmMain.MousePointer = 11
        
    fh = FreeFile
    Open Map For Binary Access Read As fh
    
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
                If Blqs(i).x > XMinMapSize And Blqs(i).x < XMaxMapSize And Blqs(i).y > YMinMapSize And Blqs(i).y < YMaxMapSize Then _
                    MapData(Blqs(i).x, Blqs(i).y).bLocked = 1
            Next i

        End If
            
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2

            For i = 1 To .NumeroLayers(2)
                If L2(i).x > XMinMapSize And L2(i).x < XMaxMapSize And L2(i).y > YMinMapSize And L2(i).y < YMaxMapSize Then _
                    InitGrh MapData(L2(i).x, L2(i).y).Graphic(2), L2(i).GrhIndex
            Next i

        End If
            
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3

            For i = 1 To .NumeroLayers(3)
                If L3(i).x > XMinMapSize And L3(i).x < XMaxMapSize And L3(i).y > YMinMapSize And L3(i).y < YMaxMapSize Then _
                    InitGrh MapData(L3(i).x, L3(i).y).Graphic(3), L3(i).GrhIndex
            Next i

        End If
            
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4

            For i = 1 To .NumeroLayers(4)
                If L4(i).x > XMinMapSize And L4(i).x < XMaxMapSize And L4(i).y > YMinMapSize And L4(i).y < YMaxMapSize Then _
                    InitGrh MapData(L4(i).x, L4(i).y).Graphic(4), L4(i).GrhIndex
            Next i

        End If
            
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers

            For i = 1 To .NumeroTriggers
                If Triggers(i).x > XMinMapSize And Triggers(i).x < XMaxMapSize And Triggers(i).y > YMinMapSize And Triggers(i).y < YMaxMapSize Then _
                    MapData(Triggers(i).x, Triggers(i).y).Trigger = Triggers(i).Trigger
            Next i

        End If
            
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas

            For i = 1 To .NumeroParticulas
                If Particulas(i).x > XMinMapSize And Particulas(i).x < XMaxMapSize And Particulas(i).y > YMinMapSize And Particulas(i).y < YMaxMapSize Then
                    MapData(Particulas(i).x, Particulas(i).y).Particle_Index = Particulas(i).Particula
                    Call General_Particle_Create(Particulas(i).Particula, Particulas(i).x, Particulas(i).y)
                End If
            Next i

        End If

        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)

            Dim p As Byte

            Get #fh, , Luces

            For i = 1 To .NumeroLuces

                If Luces(i).x > XMinMapSize And Luces(i).x < XMaxMapSize And Luces(i).y > YMinMapSize And Luces(i).y < YMaxMapSize Then
                    With MapData(Luces(i).x, Luces(i).y)
                        .Light.range = Luces(i).range
                        .Light.RGBCOLOR.A = 255
                        .Light.RGBCOLOR.R = Luces(i).R
                        .Light.RGBCOLOR.G = Luces(i).G
                        .Light.RGBCOLOR.B = Luces(i).B

                    End With

                    Call LucesRedondas.Create_Light_To_Map(Luces(i).x, Luces(i).y, RGBA_From_Comp(Luces(i).R, Luces(i).G, Luces(i).B), Luces(i).range)
                End If
            Next i

            Call LucesRedondas.LightRenderAll

        End If

        '
        If .NumeroZonas > 0 Then
            ReDim Zonas(1 To .NumeroZonas)
            Get #fh, , Zonas

            For i = 1 To .NumeroZonas
                If Zonas(i).x > XMinMapSize And Zonas(i).x < XMaxMapSize + 1 And Zonas(i).y > YMinMapSize And Zonas(i).y < YMaxMapSize + 1 Then _
                    MapData(Zonas(i).x, Zonas(i).y).ZonaIndex = Zonas(i).Zona
            Next i

        End If
            
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos

            For i = 1 To .NumeroOBJs
                MapData(Objetos(i).x, Objetos(i).y).OBJInfo.ObjIndex = Objetos(i).ObjIndex
                MapData(Objetos(i).x, Objetos(i).y).OBJInfo.Amount = Objetos(i).ObjAmmount

                If MapData(Objetos(i).x, Objetos(i).y).OBJInfo.ObjIndex > NumObjs Then
                    InitGrh MapData(Objetos(i).x, Objetos(i).y).ObjGrh, 20299
                Else
                    InitGrh MapData(Objetos(i).x, Objetos(i).y).ObjGrh, ObjData(MapData(Objetos(i).x, Objetos(i).y).OBJInfo.ObjIndex).GrhIndex

                End If

            Next i

        End If
                
        If .NumeroNPCs > 0 Then
            ReDim NPCs(1 To .NumeroNPCs)
            Get #fh, , NPCs

            For i = 1 To .NumeroNPCs

                If NPCs(i).NPCIndex > 0 Then
                    MapData(NPCs(i).x, NPCs(i).y).NPCIndex = NPCs(i).NPCIndex
                    Call Char_Make(NextOpenChar(), NpcData(NPCs(i).NPCIndex).Body, NpcData(NPCs(i).NPCIndex).Head, NpcData(NPCs(i).NPCIndex).Heading, _
                                    NPCs(i).x, NPCs(i).y, NpcData(NPCs(i).NPCIndex).WeaponAnim, NpcData(NPCs(i).NPCIndex).ShieldAnim, NpcData(NPCs(i).NPCIndex).CascoAnim)

                End If

            Next i

        End If
    
        If .NumeroTE > 0 Then
            ReDim TEs(1 To .NumeroTE)
            Get #fh, , TEs

            For i = 1 To .NumeroTE
                MapData(TEs(i).x, TEs(i).y).TileExit.Map = TEs(i).DestM
                MapData(TEs(i).x, TEs(i).y).TileExit.x = TEs(i).DestX
                MapData(TEs(i).x, TEs(i).y).TileExit.y = TEs(i).DestY
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
    
    Call Pestanas(Map, ".csm") '

    'Change mouse icon
    frmMain.MousePointer = 0
    
    Call CSMInfoCargar
    
    '    'Set changed flag
    MapInfo.Changed = 0
    
    Call coloresZona

    MapaCargado = True
    
    Call DibujarMinimapa
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Mapa " & Map & " cargado...", 0, 255, 0, , , True)

    Exit Sub

ErrorHandler:

    If fh <> 0 Then Close fh
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Error en el Mapa " & Map & ", se ha generado un informe de errores en: " & App.Path & "\Logs.txt", 255, 0, 0)
    
    File = FreeFile
    
    Call RegistrarError(Err.Number, Err.Description, "modMapas.Cargar_CSM", Erl)

End Sub

Public Function Save_CSM(ByVal MapRoute As String) As Boolean

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
    
    ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax)
    
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax

                With MapData(i, j)

                    If .bLocked Then
                        MH.NumeroBloqueados = MH.NumeroBloqueados + 1
                        ReDim Preserve Blqs(1 To MH.NumeroBloqueados)
                        Blqs(MH.NumeroBloqueados).x = i
                        Blqs(MH.NumeroBloqueados).y = j

                    End If
                    
                    L1(i, j) = .Graphic(1).GrhIndex
                    
                    If .Graphic(2).GrhIndex > 0 Then
                        MH.NumeroLayers(2) = MH.NumeroLayers(2) + 1
                        ReDim Preserve L2(1 To MH.NumeroLayers(2))
                        L2(MH.NumeroLayers(2)).x = i
                        L2(MH.NumeroLayers(2)).y = j
                        L2(MH.NumeroLayers(2)).GrhIndex = .Graphic(2).GrhIndex

                    End If
                    
                    If .Graphic(3).GrhIndex > 0 Then
                        MH.NumeroLayers(3) = MH.NumeroLayers(3) + 1
                        ReDim Preserve L3(1 To MH.NumeroLayers(3))
                        L3(MH.NumeroLayers(3)).x = i
                        L3(MH.NumeroLayers(3)).y = j
                        L3(MH.NumeroLayers(3)).GrhIndex = .Graphic(3).GrhIndex

                    End If
                    
                    If .Graphic(4).GrhIndex > 0 Then
                        MH.NumeroLayers(4) = MH.NumeroLayers(4) + 1
                        ReDim Preserve L4(1 To MH.NumeroLayers(4))
                        L4(MH.NumeroLayers(4)).x = i
                        L4(MH.NumeroLayers(4)).y = j
                        L4(MH.NumeroLayers(4)).GrhIndex = .Graphic(4).GrhIndex

                    End If
                    
                    If .Trigger > 0 Then
                        MH.NumeroTriggers = MH.NumeroTriggers + 1
                        ReDim Preserve Triggers(1 To MH.NumeroTriggers)
                        Triggers(MH.NumeroTriggers).x = i
                        Triggers(MH.NumeroTriggers).y = j
                        Triggers(MH.NumeroTriggers).Trigger = .Trigger

                    End If
                    
                    If .Particle_Index > 0 Then
                        MH.NumeroParticulas = MH.NumeroParticulas + 1
                        ReDim Preserve Particulas(1 To MH.NumeroParticulas)
                        Particulas(MH.NumeroParticulas).x = i
                        Particulas(MH.NumeroParticulas).y = j
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
                        Luces(MH.NumeroLuces).x = i
                        Luces(MH.NumeroLuces).y = j

                    End If
                    
                    If .ZonaIndex > 0 And .ZonaIndex <= CantZonas Then
                    
                        MH.NumeroZonas = MH.NumeroZonas + 1
                        ReDim Preserve Zonas(1 To MH.NumeroZonas)
                        Zonas(MH.NumeroZonas).x = i
                        Zonas(MH.NumeroZonas).y = j
                        Zonas(MH.NumeroZonas).Zona = .ZonaIndex

                    End If

                    If .OBJInfo.ObjIndex > 0 Then
                        MH.NumeroOBJs = MH.NumeroOBJs + 1
                        ReDim Preserve Objetos(1 To MH.NumeroOBJs)
                        Objetos(MH.NumeroOBJs).ObjIndex = .OBJInfo.ObjIndex
                        Objetos(MH.NumeroOBJs).ObjAmmount = .OBJInfo.Amount
                        Objetos(MH.NumeroOBJs).x = i
                        Objetos(MH.NumeroOBJs).y = j

                    End If
                    
                    If .NPCIndex > 0 Then
                        MH.NumeroNPCs = MH.NumeroNPCs + 1
                        ReDim Preserve NPCs(1 To MH.NumeroNPCs)
                        NPCs(MH.NumeroNPCs).NPCIndex = .NPCIndex
                        NPCs(MH.NumeroNPCs).x = i
                        NPCs(MH.NumeroNPCs).y = j

                    End If
                    
                    If .TileExit.Map > 0 Then
                        MH.NumeroTE = MH.NumeroTE + 1
                        ReDim Preserve TEs(1 To MH.NumeroTE)
                        TEs(MH.NumeroTE).DestM = .TileExit.Map
                        TEs(MH.NumeroTE).DestX = .TileExit.x
                        TEs(MH.NumeroTE).DestY = .TileExit.y
                        TEs(MH.NumeroTE).x = i
                        TEs(MH.NumeroTE).y = j

                    End If
                    
                End With

        Next i
    Next j
 
    MH.NumeroData = CantZonas
    
    Call CSMInfoSave
              
    fh = FreeFile
    Open MapRoute For Binary As fh
    
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
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Mapa " & MapRoute & " guardado...", 0, 255, 0, , , True)
    Exit Function

ErrorHandler:

    If fh <> 0 Then Close fh

End Function

Private Sub CSMInfoCargar()
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
    
    ReDim MapZonas(CantZonas) As tZonaInfo
    
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
            .lvlMaximo = Val(MapDat(i).lvlMaximo)
            .LuzBase = MapDat(i).LuzBase
            
            If MapDat(i).LuzBase <> 0 Then
                frmMapInfo.chkLuzClimatica = Checked
                Call ConvertLongToRGB(MapDat(i).LuzBase, tR, tG, tB)
                
                Estado_Custom.A = 255
                Estado_Custom.R = tR
                Estado_Custom.G = tG
                Estado_Custom.B = tB
                
                Call Actualizar_Estado
                
                frmMapInfo.LuzMapa.Text = tR & "-" & tG & "-" & tB
                frmMapInfo.PicColorMap.BackColor = .LuzBase
                
            Else
                frmMapInfo.chkLuzClimatica = Unchecked
                
            End If
            
            If MapDat(i).battle_mode = True Then
                .PK = True
            Else
                .PK = False

            End If
            
            .Ambient = MapDat(i).Ambient
            
            .Terreno = MapDat(i).terrain
            .Zona = MapDat(i).zone
            .Restringir = MapDat(i).restrict_mode
            .BackUp = MapDat(i).backup_mode
        
            frmZonas.LstZona.AddItem (i & " - " & .name)
        
        End With

    Next i
    
    Call MapInfo_Actualizar

End Sub

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
            MapDat(i).lvlMaximo = .lvlMaximo
            
            If frmMapInfo.chkLuzClimatica = Checked Then
                MapDat(i).LuzBase = .LuzBase
            Else
                MapDat(i).LuzBase = 0

            End If
            
            If .PK = True Then
                MapDat(i).battle_mode = True
            Else
                MapDat(i).battle_mode = False

            End If
            
            MapDat(i).Ambient = .Ambient
            MapDat(i).terrain = .Terreno
            MapDat(i).zone = .Zona
            MapDat(i).restrict_mode = .Restringir
            MapDat(i).backup_mode = .BackUp

        End With
        
    Next i
    
End Sub

Public Function Exportar_Zonas(ByVal ZonRoute As String) As Boolean

    '**********************************
    'Autor: Lorwik
    'Fecha: 18/05/2024
    'Descripcion: Exporta las zonas a un archivo binario
    '**********************************
    On Error GoTo Exportar_Zonas_Err

    Dim fh           As Integer

    Dim MH           As tMapHeader

    Dim Zonas()      As tDatosZonas

    Dim i            As Integer

    Dim j            As Integer
    
    If NoSobreescribir = False Then
        If FileExist(ZonRoute, vbNormal) = True Then
            If MsgBox("¿Desea sobrescribir " & ZonRoute & "?", vbCritical + vbYesNo) = vbNo Then
                Exit Function
                
            Else
                'Kill ZonRoute
                
            End If

        End If

    End If
    
    frmMain.MousePointer = 11
    MapSize.XMax = XMaxMapSize
    MapSize.XMin = XMinMapSize
    MapSize.YMax = YMaxMapSize
    MapSize.YMin = YMinMapSize
    
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax

            With MapData(i, j)
                    
                If .ZonaIndex > 0 And .ZonaIndex <= CantZonas Then
                    
                    MH.NumeroZonas = MH.NumeroZonas + 1
                    ReDim Preserve Zonas(1 To MH.NumeroZonas)
                    Zonas(MH.NumeroZonas).x = i
                    Zonas(MH.NumeroZonas).y = j
                    Zonas(MH.NumeroZonas).Zona = .ZonaIndex

                End If

            End With

        Next i
    Next j
 
    MH.NumeroData = CantZonas
    
    Call CSMInfoSave
              
    fh = FreeFile
    Open ZonRoute For Binary As fh
        
    Put #fh, , MH
    Put #fh, , MapSize
    Put #fh, , MapDat

    If MH.NumeroZonas > 0 Then Put #fh, , Zonas

    Close fh
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
    NoSobreescribir = False
    
    Exit Function

    Exportar_Zonas = True

    Exit Function

Exportar_Zonas_Err:

    Exportar_Zonas = False
    Call AddtoRichTextBox(frmConsola.StatTxt, "Error al exportar las zonas.", 255, 0, 0)
    Call RegistrarError(Err.Number, Err.Description, "modMapasWAO.Exportar_Zonas", Erl)

    Resume Next

End Function

Public Function Importar_Zonas(ByVal Path As String) As Boolean

    '**********************************
    'Autor: Lorwik
    'Fecha: 18/05/2024
    'Descripcion: Importa zonas desde un archivo binario
    '**********************************
    On Error GoTo Importar_Zonas_Err

    Dim fh         As Integer

    Dim File       As Integer

    Dim MH         As tMapHeader

    Dim Zonas()    As tDatosZonas
    
    Dim i          As Long

    Dim j          As Long

    DoEvents
        
    'Change mouse icon
    frmMain.MousePointer = 11
        
    fh = FreeFile
    Open Path For Binary Access Read As fh
    
    Get #fh, , MH
    Get #fh, , MapSize
        
    ReDim MapDat(MH.NumeroData) As tMapDat
    Get #fh, , MapDat

    If MH.NumeroZonas > 0 Then
        ReDim Zonas(1 To MH.NumeroZonas)
        Get #fh, , Zonas

        For i = 1 To MH.NumeroZonas

            If Zonas(i).x > XMinMapSize And Zonas(i).x < XMaxMapSize + 1 And Zonas(i).y > YMinMapSize And Zonas(i).y < YMaxMapSize + 1 Then MapData(Zonas(i).x, Zonas(i).y).ZonaIndex = Zonas(i).Zona
        
        Next i

    End If
              
    Close fh

    'Change mouse icon
    frmMain.MousePointer = 0
    
    Call CSMInfoCargar
    
    'Set changed flag
    MapInfo.Changed = 0
    
    Call coloresZona

    Importar_Zonas = True
    
    Call DibujarMinimapa

    Exit Function

Importar_Zonas_Err:

    Importar_Zonas = False
    Call AddtoRichTextBox(frmConsola.StatTxt, "Error en la carga del mapa " & Map, 255, 0, 0, , , True)
    Call RegistrarError(Err.Number, Err.Description, "modMapasWAO.Importar_Zonas", Erl)

    Resume Next

End Function
