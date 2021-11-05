Attribute VB_Name = "modOtroMapas"
Option Explicit
'***************************
'Map format .CSM
'***************************
Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
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

Private Type tDatosLucesIAO
    X As Integer
    Y As Integer
    color As Long
    Rango As Byte
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

Private Type tMapDatIAOold
    map_name As String * 64
    battle_mode As Byte
    backup_mode As Byte
    restrict_mode As String * 4
    music_number As String * 16
    zone As String * 16
    terrain As String * 16
    Ambient As String * 16
    base_light As Long
End Type

Private Type tMapDatIAO
    map_name As String * 64
    battle_mode As Byte
    backup_mode As Byte
    restrict_mode As String * 4
    music_number As String * 16
    zone As String * 16
    terrain As String * 16
    Ambient As String * 16
    base_light As Long
    lvlMinimo As Long
    extra1 As Long
    extra2 As Long
    extra3 As String * 32
End Type

Public MapDatIAO As tMapDatIAO
Public MapDatIAOOld As tMapDatIAOold

Sub Cargar_MapIAO(ByVal Map As String, ByVal Tipo As eTipoMapa)
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
    Dim Luces()         As tDatosLucesIAO
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
            
            Get #fh, , MapDatIAO
            ReDim Ln1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
            Get #fh, , Ln1
            
        ElseIf Tipo = eTipoMapa.tIAOold Then
            
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
    
    'Change mouse icon
    frmMain.MousePointer = 0
    
    If Tipo = eTipoMapa.tIAOold Then
        Call CSMInfoCargarIAOold
    End If
    
    'Set changed flag
    MapInfo.Changed = 0

    MapaCargado = True
    
    TipoMapaActual = Tipo
    
    Call DibujarMinimapa ' Radar
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Mapa " & Map & " cargado...", 0, 255, 0)
    
ErrorHandler:

    If fh <> 0 Then Close fh
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Error en el Mapa " & Map & ", se ha generado un informe de errores en: " & App.Path & "\Logs.txt", 255, 0, 0)
    
    File = FreeFile
    
    Call RegistrarError(Err.Number, Err.Description, "modMapas.Cargar_CSM", Erl)

End Sub

Public Sub CSMInfoCargarIAOold()
'**********************************
'Autor: Lorwik
'Fecha: 14/03/2021
'Descripcion: Cargar la informacion de los mapas de IAO 1.3
'**********************************

    Dim tR As Byte
    Dim tG As Byte
    Dim tB As Byte
    
    MapInfo.name = MapDatIAOOld.map_name
    MapInfo.Music = MapDatIAOOld.music_number
    
    MapInfo.lvlMinimo = 1
    frmMapInfo.chkLuzClimatica = Unchecked
    
    MapInfo.MapVersion = 0
    
    If MapDatIAOOld.battle_mode = True Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False
    End If
    
    MapInfo.Ambient = MapDatIAOOld.Ambient
    
    MapInfo.Terreno = MapDatIAOOld.terrain
    MapInfo.Zona = MapDatIAOOld.zone
    MapInfo.Restringir = MapDatIAOOld.restrict_mode
    MapInfo.BackUp = MapDatIAOOld.backup_mode
    
    Call MapInfo_Actualizar
End Sub
