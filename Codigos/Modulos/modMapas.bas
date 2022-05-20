Attribute VB_Name = "modMapas"
Option Explicit

Public TipoMapaActual As Byte 'Indica que tipo de mapa se abrio actualmente

'/////////////////////////////////////////////////////////////////////
'Lectura, guardado y otras features del formato de mapas Argentum y
'otras funciones relacionadas con mapas en general
'/////////////////////////////////////////////////////////////////////

Public Sub AbrirMapa(ByRef Tipo As Byte)

    frmMain.Dialog.CancelError = True

    On Error GoTo AbrirMapa_Err
    
    DeseaGuardarMapa frmMain.Dialog.filename
    
    frmMain.ObtenerNombreArchivo False
    
    If Len(frmMain.Dialog.filename) < 3 Then Exit Sub
    
    If WalkMode = True Then _
        Call modGeneral.ToggleWalkMode
        
    Call modMapas.NuevoMapa
        
    Call abrirCargarMapa(frmMain.Dialog.filename, Tipo)
        
    DoEvents
    frmMain.mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
    Exit Sub
    
AbrirMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapas.AbrirMapa", Erl)
    Resume Next
End Sub

Public Sub abrirCargarMapa(ByVal Path As String, ByRef Tipo As Byte)

    If frmMain.Dialog.FilterIndex = 1 Then
    
        Select Case Tipo
        
            Case eTipoMapa.tWinter_Old
                'Call modMapas.Cargar_CSM_Old(frmMain.Dialog.filename)
                MsgBox "Función no disponible por el momento."
                Exit Sub
                
            Case eTipoMapa.tWinter
                Call modMapasWAO.Cargar_CSM(frmMain.Dialog.filename)
                
            Case eTipoMapa.tIAOClasico
                #If Privado = 0 Then
                    Call modMapasIAC.Cargar_MapImpClasico(frmMain.Dialog.filename)
                #End If
                
            Case eTipoMapa.tIAOold
                Call Cargar_MapIAO(frmMain.Dialog.filename, eTipoMapa.tIAOold)
                
            Case eTipoMapa.tIAOnew
                Call Cargar_MapIAO(frmMain.Dialog.filename, eTipoMapa.tIAOnew)
            
        End Select

    Else
    
        Select Case Tipo
        
            Case eTipoMapa.tInt
                Call modMapas.Cargar_Map(frmMain.Dialog.filename, True)
                
            Case eTipoMapa.tlong
                Call modMapas.Cargar_Map(frmMain.Dialog.filename, False)
                
        End Select
            
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
                
            Case eMeMode.ImperiumClasico
                Call Save_MapImpClasico(Path)
                
            Case eMeMode.WinterUltimate
                Call Save_MapImpClasico(Path)
                
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
    
    Dim y     As Integer

    Dim X     As Integer

    Dim i     As Byte

    Dim LoopC As Integer
    
    frmMain.mnuReAbrirMapa.Enabled = False
    
    MapaCargado = False
    
    For LoopC = 0 To frmMain.MapPest.Count - 1
        frmMain.MapPest(LoopC).Enabled = False
    Next
    
    frmMain.MousePointer = 11
    
    Call setMapSize
        
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            With MapData(X, y)
            
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
                .TileExit.y = 0
                
                ' Triggers
                .Trigger = 0
        
                .Particle_Group_Index = 0
                
                Call Engine_Long_To_RGB_List(MapData(X, y).Engine_Light(), -1)
                
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
    Next y
    
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
        MapInfo.Ambient = 0
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
        .TxtAmbient.Text = MapInfo.Ambient
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
        .Ambient = 0
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
        .TxtAmbient.Text = MapZonas(id).Ambient
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
'CARGA DE MAPAS FORMATO ARGENTUM
'#######################################

Public Sub Cargar_Map(ByVal Map As String, Optional ByVal EsInteger As Boolean = False)
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/05/2021
    '*************************************************

    On Error Resume Next

    Dim LoopC       As Integer

    Dim tempint     As Integer

    Dim Body        As Integer

    Dim Head        As Integer

    Dim Heading     As Byte

    Dim y           As Integer

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
        TipoMapaActual = eTipoMapa.tInt
        
    Else
        TipoMapaActual = eTipoMapa.tlong
        
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
    Get FreeFileMap, , tempint
    Get FreeFileMap, , tempint
    Get FreeFileMap, , tempint
    Get FreeFileMap, , tempint
    
    'Cabecera inf
    Get FreeFileInf, , tempint
    Get FreeFileInf, , tempint
    Get FreeFileInf, , tempint
    Get FreeFileInf, , tempint
    Get FreeFileInf, , tempint

    'Load arrays
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            With MapData(X, y)
            
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
                        Get FreeFileInf, , .y
                    
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
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, X, y)
                        
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
    Next y
    
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

    Dim tempint     As Integer

    Dim y           As Long

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
    Put FreeFileMap, , tempint
    Put FreeFileMap, , tempint
    Put FreeFileMap, , tempint
    Put FreeFileMap, , tempint
    
    'inf Header
    Put FreeFileInf, , tempint
    Put FreeFileInf, , tempint
    Put FreeFileInf, , tempint
    Put FreeFileInf, , tempint
    Put FreeFileInf, , tempint
    
    'Write .map file
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            With MapData(X, y)
            
                ByFlags = 0
                
                If .bLocked = 1 Then ByFlags = ByFlags Or 1
                
                If .Graphic(2).GrhIndex Then ByFlags = ByFlags Or 2
                If .Graphic(3).GrhIndex Then ByFlags = ByFlags Or 4
                If .Graphic(4).GrhIndex Then ByFlags = ByFlags Or 8

                If .Trigger Then ByFlags = ByFlags Or 16
                    
                Put FreeFileMap, , ByFlags
                    
                If TipoMapaActual = eTipoMapa.tInt Then
                    Put FreeFileMap, , .Graphic(1).GrhIndexInt
                    
                Else
                    Put FreeFileMap, , .Graphic(1).GrhIndex

                End If
                
                For LoopC = 2 To 4
                    
                    If TipoMapaActual = eTipoMapa.tInt Then
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
                    Put FreeFileInf, , .TileExit.y

                End If
                    
                If .NPCIndex Then Put FreeFileInf, , CInt(.NPCIndex)
                    
                If .OBJInfo.ObjIndex Then
                    Put FreeFileInf, , .OBJInfo.ObjIndex
                    Put FreeFileInf, , .OBJInfo.Amount

                End If
            
            End With
            
        Next X
    Next y
    
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
