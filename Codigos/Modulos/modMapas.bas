Attribute VB_Name = "modMapas"
Option Explicit

Public colorZona() As RGBA
Public UserMap As Integer

'/////////////////////////////////////////////////////////////////////
'Lectura, guardado y otras features del formato de mapas Argentum y
'otras funciones relacionadas con mapas en general
'/////////////////////////////////////////////////////////////////////

Public Sub AbrirMapa(ByVal bBig As Boolean)

    frmMain.Dialog.CancelError = True

    On Error GoTo AbrirMapa_Err
    
    DeseaGuardarMapa frmMain.Dialog.filename
    
    frmMain.ObtenerNombreArchivo False
    
    If Len(frmMain.Dialog.filename) < 3 Then Exit Sub
    
    If WalkMode = True Then _
        Call modGeneral.ToggleWalkMode
        
'    If bBig Then
'        Call setMapSize(1000, 1000)
'
'    Else
'        Call setMapSize(100, 100)
'
'    End If
        
    Call modMapas.NuevoMapa
        
    If bBig Then
        Call abrirCargarMapa(frmMain.Dialog.filename)
        
    Else
        Call ModMapConver.Cargar_ConverCSM(frmMain.Dialog.filename)
    
    End If
    
    UserPos.x = XMaxMapSize / 2
    UserPos.y = YMaxMapSize / 2
        
    DoEvents
    frmMain.mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
    Exit Sub
    
AbrirMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapas.AbrirMapa", Erl)
    Resume Next
End Sub

Public Sub abrirCargarMapa(ByVal path As String)
    
    Dim ind As Integer
    ind = InStrRev(path, "\") + 5
    UserMap = mid$(path, ind, Len(path) - ind - 3)
    
    Call modMapasWAO.Cargar_CSM(frmMain.Dialog.filename)

End Sub

Public Sub DeseaGuardarMapa(Optional path As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    'Nos pregunta donde guardar el mapa en caso de modificarlo
    '*************************************************

    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            GuardarMapa path

        End If

    End If

End Sub

Public Sub GuardarMapa(Optional path As String)
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/04/2021
    'Guarda el Mapa
    '*************************************************

    frmMain.Dialog.CancelError = True

    On Error GoTo errhandler
    
    If LenB(path) = 0 Then
        frmMain.ObtenerNombreArchivo True
        path = frmMain.Dialog.filename

        If LenB(path) = 0 Then Exit Sub

    End If
    
    Call Save_CSM(path)
    
    Call ShowMessageScreen("Mapa guardado.")
                
errhandler:

End Sub

Public Sub NuevoMapa()
    '***************************************************
    'Author: Lorwik
    'Last Modification: 26/04/2021
    'Descripcion: Limpia todo el mapa a uno nuevo
    '***************************************************
    
    Dim y     As Integer

    Dim x     As Integer

    Dim i     As Byte

    Dim loopc As Integer
    
    frmMain.mnuReAbrirMapa.Enabled = False
    
    MapaCargado = False
    
    For loopc = 0 To frmMain.MapPest.Count - 1
        frmMain.MapPest(loopc).Enabled = False
    Next
    
    frmMain.MousePointer = 11
        
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
        
            With MapData(x, y)
            
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
                
                .CharIndex = 0
                .NPCIndex = 0
        
                ' Traslados
                .TileExit.Map = 0
                .TileExit.x = 0
                .TileExit.y = 0
                
                ' Triggers
                .Trigger = 0
        
                .Particle_Group_Index = 0
                .Particle_Index = 0
                
                Call Long_2_RGBAList(MapData(x, y).Light_Value(), -1)
                
                .Light.active = False
                .Light.range = 0
                .Light.map_x = 0
                .Light.map_y = 0
                .Light.RGBCOLOR.A = 0
                .Light.RGBCOLOR.R = 0
                .Light.RGBCOLOR.G = 0
                .Light.RGBCOLOR.B = 0
                
                .ZonaIndex = 0
                
                For i = 0 To 3
                    .Light_Value(i).A = 0
                    .Light_Value(i).R = 0
                    .Light_Value(i).G = 0
                    .Light_Value(i).B = 0
                Next i

                InitGrh .Graphic(1), 1

            End With
            
        Next x
    Next y
    
    'Borramos todas las luces
    Call LucesRedondas.LightRemoveAll(False)
    
    Call ResetearZonas
    
    Call DibujarMinimapa
    
    Estado_Actual = Estados(e_estados.MedioDia)
    Call Actualizar_Estado
    
    'Set changed flag
    MapInfo.Changed = 0
    frmMain.MousePointer = 0
    
    'Vaciamos la cola de movimiento
    keysMovementPressedQueue.Clear
    
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
        .txtNivelMaximo = MapInfo.lvlMaximo
        .chkMapMagiaSinEfecto.value = MapInfo.MagiaSinEfecto
        .chkMapInviSinEfecto.value = IIf(MapInfo.InviSinEfecto, vbChecked, vbUnchecked)
        .chkInvocarSin.value = MapInfo.InvocarSinEfecto
        .chkOcultarSin.value = MapInfo.OcultarSinEfecto
        .chkMapResuSinEfecto.value = IIf(MapInfo.ResuSinEfecto, vbChecked, vbUnchecked)
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

    If id <= CantZonas Then Call ResetearZona(id)
    
    frmZonas.LstZona.AddItem (CantZonas & " - " & MapZonas(CantZonas).name)
    
    ReDim Preserve colorZona(CantZonas) As RGBA
    
    Call coloresZona(CantZonas)

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
        MsgBox "El numero de zonas llego al minimo. No puedes eliminar mas zonas."
        Exit Sub

    End If
    
    'Primero se resetea la zona
    Call ResetearZona(CantZonas)

    frmZonas.LstZona.RemoveItem frmZonas.LstZona.ListIndex
    
    CantZonas = CantZonas - 1
    
    ReDim Preserve MapZonas(CantZonas) As tMapInfo
    ReDim Preserve colorZona(CantZonas) As RGBA
    
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
        frmZonas.LstZona.AddItem (i & " - " & MapZonas(i).name)
            
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
        .txtNivelMaximo = MapZonas(id).lvlMaximo
        .chkMapMagiaSinEfecto.value = IIf(MapZonas(id).MagiaSinEfecto, vbChecked, vbUnchecked)
        .chkMapInviSinEfecto.value = IIf(MapZonas(id).InviSinEfecto, vbChecked, vbUnchecked)
        .chkInvocarSin.value = IIf(MapZonas(id).InvocarSinEfecto, vbChecked, vbUnchecked)
        .chkOcultarSin.value = IIf(MapZonas(id).OcultarSinEfecto, vbChecked, vbUnchecked)
        .chkMapResuSinEfecto.value = IIf(MapZonas(id).ResuSinEfecto, vbChecked, vbUnchecked)
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

Public Sub coloresZona(Optional ByVal zonaID As Long = -1)
    '*****************************************
    'Autor: Lorwik
    'Fecha: 16/05/2024
    'Descripción: Asigna colores aleatorios a zonas
    '*****************************************

    Dim i As Long

    If CantZonas = 0 Then Exit Sub
    
    ' Redimensiona colorZona si es necesario
    If UBound(colorZona) < CantZonas Then
        ReDim Preserve colorZona(CantZonas) As RGBA
    End If

    ' Asigna colores aleatorios
    If zonaID = -1 Then
        ' Asignar colores a todas las zonas
        For i = 1 To CantZonas
            colorZona(i) = GenerarColorAleatorio()
        Next i
    Else
        ' Asignar color a una zona específica
        colorZona(zonaID) = GenerarColorAleatorio()
    End If
End Sub

Private Function GenerarColorAleatorio() As RGBA
    '*****************************************
    'Autor: Lorwik
    'Fecha: 16/05/2024
    '*****************************************
    ' Genera un color aleatorio en formato RGBA
    Dim RedValue As Byte
    Dim GreenValue As Byte
    Dim BlueValue As Byte
    
    RedValue = Int(Rnd * 256) ' Valor aleatorio entre 0 y 255
    GreenValue = Int(Rnd * 256)
    BlueValue = Int(Rnd * 256)
    
    GenerarColorAleatorio = RGBA_From_Comp(RedValue, GreenValue, BlueValue, 255)
End Function

Public Sub Pestanas(ByVal Map As String, Optional ByVal MapFormat As String = ".map")

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    'Calcula la orden de Pestanas
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
    
    MapaActual = ReadField(1, Right(Map, Len(Map)), Asc("."))
    
    nMapaActual = ReadField(1, Right(Map, Len(Map) - 4), Asc("."))
    'If frmCopiarBordes.Visible Then Call frmCopiarBordes.Inicializar
    
    For loopc = Len(Left(Map, Len(Map) - 4)) To 1 Step -1

        If IsNumeric(mid(Left(Map, Len(Map) - 4), loopc, 1)) = False Then
            NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - loopc)
            NameMap_Save = Left(Map, loopc)
            Exit For

        End If

    Next
    
    For loopc = (NumMap_Save - 4) To (NumMap_Save + 6)

        If FileExist(PATH_Save & NameMap_Save & loopc & MapFormat, vbArchive) = True Then
            frmMain.MapPest(loopc - NumMap_Save + 4).Visible = True
            frmMain.MapPest(loopc - NumMap_Save + 4).Enabled = True
            frmMain.MapPest(loopc - NumMap_Save + 4).Caption = NameMap_Save & loopc
        Else
            frmMain.MapPest(loopc - NumMap_Save + 4).Visible = False

        End If

    Next
    
End Sub

Public Sub ResetearZonas()
'***************************************
'Autor: Lorwik
'Fecha: 08/06/2024
'***************************************

    CantZonas = 0
    ReDim MapZonas(CantZonas) As tMapInfo
    
    frmZonas.LstZona.Clear
    
    Call NuevaZona(CantZonas)
        
    frmZonas.LstZona.ListIndex = 0
        
    Call MapZona_Actualizar(frmZonas.LstZona.ListIndex + 1)
    
End Sub
