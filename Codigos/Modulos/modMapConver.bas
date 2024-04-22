Attribute VB_Name = "ModMapConver"
Option Explicit

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
    Ambient As String
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

'#######################################
'CARGAR DE MAPAS FORMATO WINTER ACTUAL
'#######################################

Sub Cargar_ConverCSM(ByVal Map As String)
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

        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)

            Dim p As Byte

            Get #fh, , Luces

            For i = 1 To .NumeroLuces

                If Luces(i).X > XMinMapSize And Luces(i).X < XMaxMapSize And Luces(i).Y > YMinMapSize And Luces(i).Y < YMaxMapSize Then
                    With MapData(Luces(i).X, Luces(i).Y)
                        .Light.range = Luces(i).range
                        .Light.RGBCOLOR.A = 255
                        .Light.RGBCOLOR.R = Luces(i).R
                        .Light.RGBCOLOR.G = Luces(i).G
                        .Light.RGBCOLOR.B = Luces(i).B

                    End With

                    Call LucesRedondas.Create_Light_To_Map(Luces(i).X, Luces(i).Y, RGBA_From_Comp(Luces(i).R, Luces(i).G, Luces(i).B), Luces(i).range)
                End If
            Next i

            Call LucesRedondas.LightRenderAll

        End If

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
                    Call Char_Make(NextOpenChar(), NpcData(NPCs(i).NPCIndex).Body, NpcData(NPCs(i).NPCIndex).Head, NpcData(NPCs(i).NPCIndex).Heading, _
                                    NPCs(i).X, NPCs(i).Y, NpcData(NPCs(i).NPCIndex).WeaponAnim, NpcData(NPCs(i).NPCIndex).ShieldAnim, NpcData(NPCs(i).NPCIndex).CascoAnim)

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
    
    Call Pestanas(Map, ".csm") '

    'Change mouse icon
    frmMain.MousePointer = 0
    
    Call CSMInfoCargar
    
    '    'Set changed flag
    MapInfo.Changed = 0
    
    Call coloresZona

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
        
            frmZonas.LstZona.AddItem (i & "- " & .name)
        
        End With

    Next i
    
    Call MapInfo_Actualizar

End Sub
