Attribute VB_Name = "modEdicion"
Option Explicit

Public Enum eEstadoSelect
    nada = 0
    Copiado
End Enum

Public EstadoSelect As Byte

''
' Elimita una capa completa del mapa
'
' @param Capa Especifica la capa


Public Sub Quitar_Capa(ByVal Capa As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If EditWarning Then Exit Sub
    
    '*****************************************************************
    'Clears one layer
    '*****************************************************************
    
    Dim y As Integer
    Dim x As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            If Capa = 1 Then
                MapData(x, y).Graphic(Capa).GrhIndex = 1
            Else
                MapData(x, y).Graphic(Capa).GrhIndex = 0
            End If
        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1
End Sub

''
' Coloca la misma superficie seleccionada en todo el mapa
'

Public Sub Superficie_Todo()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim y As Integer
    Dim x As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
    
            If frmConfigSup.MOSAICO.value = vbChecked Then
                Dim aux As Integer
                aux = Val(frmSuperficies.cGrh.Text) + _
                ((y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (x Mod frmConfigSup.mAncho)
                 MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                'Setup GRH
                InitGrh MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)), aux
            Else
                'Else Place graphic
                MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = Val(frmSuperficies.cGrh.Text)
                'Setup GRH
                InitGrh MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)), Val(frmSuperficies.cGrh.Text)
            End If
    
        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Bloquea los Bordes del Mapa
'

Public Sub Bloquear_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Dim y As Long
    Dim x As Long
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
                MapData(x, y).bLocked = 1
            End If
        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1
    
End Sub

''
' Desbloquea los Bordes del Mapa
'

Public Sub Desbloquear_Bordes()
'*************************************************
'Author: Lorwik
'Last modified: 20/05/2022
'*************************************************
    Dim y As Long
    Dim x As Long
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
                MapData(x, y).bLocked = 0
            End If
        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Modifica los bloqueos de todo mapa
'
' @param Valor Especifica el estado de Bloqueo que se asignara


Public Sub Bloqueo_Todo(ByVal Valor As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If EditWarning Then Exit Sub
    
    
    Dim y As Integer
    Dim x As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            MapData(x, y).bLocked = Valor
        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Elimita todo lo que se encuentre en los bordes del mapa
'

Public Sub Quitar_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If EditWarning Then Exit Sub
    
    '*****************************************************************
    'Clears a border in a room with current GRH
    '*****************************************************************
    
    Dim y As Integer
    Dim x As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
    
            If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
            
                MapData(x, y).Graphic(1).GrhIndex = 1
                InitGrh MapData(x, y).Graphic(1), 1
                MapData(x, y).bLocked = 0
                
                 'Erase NPCs
                If MapData(x, y).NPCIndex > 0 Then
                    Char_Erase MapData(x, y).CharIndex
                    MapData(x, y).NPCIndex = 0
                End If
    
                'Erase Objs
                MapData(x, y).OBJInfo.ObjIndex = 0
                MapData(x, y).OBJInfo.Amount = 0
                MapData(x, y).ObjGrh.GrhIndex = 0
    
                'Clear exits
                MapData(x, y).TileExit.Map = 0
                MapData(x, y).TileExit.x = 0
                MapData(x, y).TileExit.y = 0
                
                ' Triggers
                MapData(x, y).Trigger = 0
    
            End If
    
        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Coloca la superficie seleccionada en todos los bordes
'

Public Sub Superficie_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    Dim y As Integer
    Dim x As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
    
            If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    
              If frmConfigSup.MOSAICO.value = vbChecked Then
                Dim aux As Integer
                aux = Val(frmSuperficies.cGrh.Text) + _
                ((y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (x Mod frmConfigSup.mAncho)
                If frmBloqueos.cInsertarBloqueo.value = True Then
                    MapData(x, y).bLocked = 1
                Else
                    MapData(x, y).bLocked = 0
                End If
                MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                'Setup GRH
                InitGrh MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)), aux
              Else
                'Else Place graphic
                If frmBloqueos.cInsertarBloqueo.value = True Then
                    MapData(x, y).bLocked = 1
                Else
                    MapData(x, y).bLocked = 0
                End If
                
                MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = Val(frmSuperficies.cGrh.Text)
                
                'Setup GRH
        
                InitGrh MapData(x, y).Graphic(Val(frmSuperficies.cCapas.Text)), Val(frmSuperficies.cGrh.Text)
            End If
                 'Erase NPCs
                If MapData(x, y).NPCIndex > 0 Then
                    Char_Erase MapData(x, y).CharIndex
                    MapData(x, y).NPCIndex = 0
                End If
    
                'Erase Objs
                MapData(x, y).OBJInfo.ObjIndex = 0
                MapData(x, y).OBJInfo.Amount = 0
                MapData(x, y).ObjGrh.GrhIndex = 0
    
                'Clear exits
                MapData(x, y).TileExit.Map = 0
                MapData(x, y).TileExit.x = 0
                MapData(x, y).TileExit.y = 0
    
            End If
    
        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Coloca la zona seleccionada en todos los bordes
'

Public Sub Zonas_Bordes()
'*************************************************
'Author: Lorwik
'Last modified: 25/05/2022
'*************************************************

    Dim y As Integer
    Dim x As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
    
            If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    
                MapData(x, y).ZonaIndex = frmZonas.LstZona.ListIndex + 1
    
            End If
    
        Next x
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Elimita los NPCs del mapa
'
' @param Hostiles Indica si elimita solo hostiles o solo npcs no hostiles

Public Sub Quitar_NPCs(ByVal Hostiles As Boolean, ByVal Zona As Boolean)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Quitar_NPCs_Err

    Dim y As Integer
    Dim x As Integer

    'Mensaje de Alerta
    If EditWarning Then Exit Sub

    'Si es el mapa general...
    If Not Zona Then

        For y = YMinMapSize To YMaxMapSize
            For x = XMinMapSize To XMaxMapSize
                
                If Not Hostiles Then
                    If MapData(x, y).NPCIndex > 0 Then
                        Call Char_Erase(MapData(x, y).CharIndex)
                        MapData(x, y).NPCIndex = 0
        
                    End If
                Else

                    If MapData(x, y).NPCIndex > 500 Then
                        Call Char_Erase(MapData(x, y).CharIndex)
                        MapData(x, y).NPCIndex = 0
        
                    End If
                End If
            
            Next x
        Next y
    
    Else 'Si es solo la zona...
    
        Dim zonaNumber As Integer
        
        zonaNumber = InputBox("Indica el numero de la zona.", "Eliminar NPC's de la zona.")
        
        For y = YMinMapSize To YMaxMapSize
            For x = XMinMapSize To XMaxMapSize
            
                If Not Hostiles Then
                    If zonaNumber = MapData(x, y).ZonaIndex Then
                        If MapData(x, y).NPCIndex > 0 Then
                            Call Char_Erase(MapData(x, y).CharIndex)
                            MapData(x, y).NPCIndex = 0
        
                        End If
                    End If
                Else

                    If zonaNumber = MapData(x, y).ZonaIndex Then
                        If MapData(x, y).NPCIndex > 500 Then
                            Call Char_Erase(MapData(x, y).CharIndex)
                            MapData(x, y).NPCIndex = 0
        
                        End If
                    End If
                End If
                
            Next x
            
        Next y
    
    End If

    'Set changed flag
    MapInfo.Changed = 1

    Exit Sub

Quitar_NPCs_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Quitar_NPCs", Erl)
    Resume Next
    
End Sub

''
' Manda una advertencia de Edicion Critica
'
' @return   Nos devuelve si acepta o no el cambio

Public Function EditWarning() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If MsgBox(MSGDang, vbExclamation + vbYesNo) = vbNo Then
        EditWarning = True
        
    Else
        EditWarning = False
        
    End If
End Function

''
' Acciona la operacion al hacer doble click en una posicion del mapa
'
' @param tX Especifica la posicion X en el mapa
' @param tY Espeficica la posicion Y en el mapa

Public Sub DobleClick(tX As Integer, tY As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 27/04/2021
'Basado en el codigo de GS
'*************************************************
    Dim Formato As String
    
    Formato = ".csm"
    
    ' Traslados
    Dim tTrans As WorldPos
    tTrans = MapData(tX, tY).TileExit
    If tTrans.Map > 0 Then
        If LenB(frmMain.Dialog.filename) <> 0 Then
            If FileExist(PATH_Save & NameMap_Save & tTrans.Map & Formato, vbArchive) = True Then
                Call modMapas.NuevoMapa
                frmMain.Dialog.filename = PATH_Save & NameMap_Save & tTrans.Map & Formato
                
                Call abrirCargarMapa(frmMain.Dialog.filename)
                
                UserPos.x = tTrans.x
                UserPos.y = tTrans.y
                
                If WalkMode = True Then
                    MoveCharbyPos UserCharIndex, UserPos.x, UserPos.y
                    CharList(UserCharIndex).Heading = SOUTH
                End If
                
                frmMain.mnuReAbrirMapa.Enabled = True
            End If
        End If
    End If
End Sub

''
' Realiza una operacion de edicion aislada sobre el mapa
'
' @param Button Indica el estado del Click del mouse
' @param tX Especifica la posicion X en el mapa
' @param tY Especifica la posicion Y en el mapa

Public Sub ClickEdit(Button As Integer, tX As Integer, tY As Integer)
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2021
    'Basado en el codigo de GS
    '*************************************************

    On Error GoTo ClickEdit_Err
    
    Dim NPCIndex As Integer

    Dim ObjIndex As Integer

    Dim Head     As Integer

    Dim Body     As Integer
    
    Dim Arma     As Integer
    
    Dim Escudo   As Integer
    
    Dim Casco    As Integer

    Dim Heading  As Byte

    Dim loopc    As Integer

    If tY < YMinMapSize Or tY > YMaxMapSize Then Exit Sub
    If tX < XMinMapSize Or tX > XMaxMapSize Then Exit Sub
    
    Call ActualizarMinimapa(tX, tY)

    If Button = 0 Then
        SobreY = tY
        SobreX = tX
        
    End If

    With MapData(tX, tY)

        Select Case Button
        
                'Click Derecho
            Case vbRightButton
            
                ' Posicion
                Call AddtoRichTextBox(frmConsola.StatTxt, "Posición X:" & tX & ", Y:" & tY, 255, 255, 255, False, False, True)
                
                ' Bloqueos
                If .bLocked = 1 Then Call AddtoRichTextBox(frmConsola.StatTxt, " (BLOQ)", 255, 255, 255, False, False, True)
                
                ' Traslados
                If MapData(tX, tY).TileExit.Map <> 0 Then
                    If ClientSetup.AutoCapturarTraslados Then
                        frmTraslados.tTMapa.Text = MapData(tX, tY).TileExit.Map
                        frmTraslados.tTX.Text = MapData(tX, tY).TileExit.x
                        frmTraslados.tTY = MapData(tX, tY).TileExit.y

                    End If

                    frmConsola.StatTxt.Text = frmConsola.StatTxt.Text & " (Trans.: " & MapData(tX, tY).TileExit.Map & "," & MapData(tX, tY).TileExit.x & "," & MapData(tX, tY).TileExit.y & ")"

                End If
                
                ' NPCs
                If .NPCIndex > 0 Then
                    If .NPCIndex > 499 Then
                        Call AddtoRichTextBox(frmConsola.StatTxt, " (NPC-Hostil: " & .NPCIndex & " - " & NpcData(.NPCIndex).name & ")", 255, 255, 255, False, False, True)
                
                    Else
                        Call AddtoRichTextBox(frmConsola.StatTxt, " (NPC: " & .NPCIndex & " - " & NpcData(.NPCIndex).name & ")", 255, 255, 255, False, False, True)
                
                    End If
                    
                    If ClientSetup.AutoCapturarNPCs Then _
                        frmNPCs.cNumFunc.Text = .NPCIndex

                End If
                
                ' OBJs

                If .OBJInfo.ObjIndex > 0 Then
                    Call AddtoRichTextBox(frmConsola.StatTxt, " (Obj: " & .OBJInfo.ObjIndex & " - " & ObjData(.OBJInfo.ObjIndex).name & " - Cant.:" & .OBJInfo.Amount & ")", 255, 255, 255, False, False, True)
                    
                    If ClientSetup.AutoCapturarObjs Then _
                        frmOBJs.cNumFunc.Text = .OBJInfo.ObjIndex
                    
                End If
                
                ' Capas
                Call AddtoRichTextBox(frmConsola.StatTxt, "Capa1: " & .Graphic(1).GrhIndex & " - Capa2: " & .Graphic(2).GrhIndex & " - Capa3: " & .Graphic(3).GrhIndex & " - Capa4: " & .Graphic(4).GrhIndex, 255, 255, 255, False, False, True)

                If ClientSetup.AutoCapturarSuperficies And frmSuperficies.cSeleccionarSuperficie.value = False Then
                    If .Graphic(4).GrhIndex <> 0 Then
                        frmSuperficies.cCapas.Text = 4
                        frmSuperficies.cGrh.Text = .Graphic(4).GrhIndex
                
                    ElseIf .Graphic(3).GrhIndex <> 0 Then
                        frmSuperficies.cCapas.Text = 3
                        frmSuperficies.cGrh.Text = .Graphic(3).GrhIndex
                
                    ElseIf .Graphic(2).GrhIndex <> 0 Then
                        frmSuperficies.cCapas.Text = 2
                        frmSuperficies.cGrh.Text = .Graphic(2).GrhIndex
                
                    ElseIf .Graphic(1).GrhIndex <> 0 Then
                        frmSuperficies.cCapas.Text = 1
                        frmSuperficies.cGrh.Text = .Graphic(1).GrhIndex
                
                    End If
                    
                End If
                
                ' Particulas
                Call AddtoRichTextBox(frmConsola.StatTxt, "Particula: " & .Particle_Index, 255, 255, 255, False, False, True)
                
                If ClientSetup.AutoCapturarParticulas Then
                    If .Particle_Index > 0 Then _
                        frmParticulas.cParticula.Text = .Particle_Index
                End If
                
                Exit Sub
            
                'Click Izquierdo
            Case vbLeftButton
            
                '########################
                'SUPERFICIES
                '########################
                
                'Borrar Capas 2-3
                If frmSuperficies.cQuitarEnTodasLasCapas.value = True Then
                    MapInfo.Changed = 1 'Set changed flag
                    
                    For loopc = 2 To 3
                        .Graphic(loopc).GrhIndex = 0
                    Next loopc
                        
                    Exit Sub

                End If
            
                'Borrar Capa actual
                If frmSuperficies.cQuitarEnEstaCapa.value = True Then
                    If Val(frmSuperficies.cCapas.Text) = 1 Then
                        If .Graphic(1).GrhIndex <> 1 Then

                            MapInfo.Changed = 1 'Set changed flag
                            .Graphic(1).GrhIndex = 1
                            Exit Sub

                        End If
                        
                    ElseIf .Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex <> 0 Then

                        MapInfo.Changed = 1 'Set changed flag
                        .Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = 0
                        Exit Sub

                    End If

                End If
                
                '########################
                'SUPERFICIES
                '########################
                If frmSuperficies.cSeleccionarSuperficie.value = True Then
                    
                    If frmConfigSup.MOSAICO.value = vbChecked Then

                        Dim aux As Long

                        Dim dy  As Long

                        Dim dX  As Long

                        If frmConfigSup.DespMosaic.value = vbChecked Then
                            dy = Val(frmConfigSup.DMLargo)
                            dX = Val(frmConfigSup.DMAncho.Text)
                            
                        Else
                            dy = 0
                            dX = 0

                        End If
                            
                        If frmSuperficies.chkAutoCompletarSuperficies.value = vbUnchecked Then

                            MapInfo.Changed = 1 'Set changed flag
                            aux = Val(frmSuperficies.cGrh.Text) + (((tY + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((tX + dX) Mod frmConfigSup.mAncho.Text)
                            
                            If .Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex <> aux Or .bLocked <> frmBloqueos.Visible Then
                                .Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                                InitGrh .Graphic(Val(frmSuperficies.cCapas.Text)), aux
                                If frmSuperficies.chkBloquear.value Then .bLocked = 1
                            End If
                            
                        Else

                            MapInfo.Changed = 1 'Set changed flag
                            
                            Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
                            
                            tXX = tX
                            tYY = tY
                            desptile = 0
                            
                            If tXX > XMaxMapSize Then tXX = XMaxMapSize
                            If tXX < XMinMapSize Then tXX = XMinMapSize
                            If tYY > YMaxMapSize Then tYY = YMaxMapSize
                            If tYY < YMinMapSize Then tYY = YMinMapSize
                            
                            For i = 1 To frmConfigSup.mLargo.Text
                                For j = 1 To frmConfigSup.mAncho.Text

                                    If tYY >= YMinMapSize And tYY <= YMaxMapSize Then
                                        If tXX >= YMinMapSize And tXX <= XMaxMapSize Then
                                            aux = Val(frmSuperficies.cGrh.Text) + desptile
                                            MapData(tXX, tYY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                                            InitGrh MapData(tXX, tYY).Graphic(Val(frmSuperficies.cCapas.Text)), aux
                                            tXX = tXX + 1
                                            desptile = desptile + 1
                                            
                                            If frmSuperficies.chkBloquear.value Then MapData(tXX - 1, tYY).bLocked = 1

                                        End If

                                    End If

                                Next
                                tXX = tX
                                tYY = tYY + 1
                            Next
                            
                            tYY = tY
                            
                        End If
                      
                    Else
                    
                        'Else Place graphic
                        If .bLocked <> frmBloqueos.Visible Or .Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex <> Val(frmSuperficies.cGrh.Text) Then

                            MapInfo.Changed = 1 'Set changed flag
                            .Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = Val(frmSuperficies.cGrh.Text)
                            'Setup GRH
                            InitGrh .Graphic(Val(frmSuperficies.cCapas.Text)), Val(frmSuperficies.cGrh.Text)
                            
                            If frmSuperficies.chkBloquear.value Then .bLocked = 1
                            
                        End If
                        
                    End If
                    
                End If
                
                '########################
                'TRASLADOS
                '########################
                If frmTraslados.cInsertarTrans.value = True Then
                    If Cfg_TrOBJ > 0 And Cfg_TrOBJ <= 0 And frmTraslados.cInsertarTransOBJ.value = True Then
                        If ObjData(Cfg_TrOBJ).OBJType = 19 Then

                            MapInfo.Changed = 1 'Set changed flag
                            InitGrh .ObjGrh, ObjData(Cfg_TrOBJ).GrhIndex
                            .OBJInfo.ObjIndex = Cfg_TrOBJ
                            .OBJInfo.Amount = 1

                        End If

                    End If
                    
                    If Val(frmTraslados.tTMapa.Text) < 0 Or Val(frmTraslados.tTMapa.Text) > 9000 Then
                        MsgBox "Valor de Mapa invalido", vbCritical + vbOKOnly
                        Exit Sub
                        
                    ElseIf Val(frmTraslados.tTX.Text) < 0 Or Val(frmTraslados.tTX.Text) > XMaxMapSize Then
                        MsgBox "Valor de X invalido", vbCritical + vbOKOnly
                        Exit Sub
                        
                    ElseIf Val(frmTraslados.tTY.Text) < 0 Or Val(frmTraslados.tTY.Text) > XMaxMapSize Then
                        MsgBox "Valor de Y invalido", vbCritical + vbOKOnly
                        Exit Sub
                        
                    End If
                    
                    If frmTraslados.cUnionManual.value = True Then

                        MapInfo.Changed = 1 'Set changed flag
                        .TileExit.Map = Val(frmTraslados.tTMapa.Text)
                        
                        If tX >= 90 Then ' 21 ' derecha
                            .TileExit.x = 12
                            .TileExit.y = tY
                            
                        ElseIf tX <= 11 Then ' 9 ' izquierda
                            .TileExit.x = 91
                            .TileExit.y = tY
                            
                        End If
                        
                        If tY >= 91 Then ' 94 '''' hacia abajo
                            .TileExit.y = 11
                            .TileExit.x = tX
                            
                        ElseIf tY <= 10 Then ''' hacia arriba
                            .TileExit.y = 90
                            .TileExit.x = tX
                            
                        End If
                        
                    Else

                        MapInfo.Changed = 1 'Set changed flag
                        .TileExit.Map = Val(frmTraslados.tTMapa.Text)
                        .TileExit.x = Val(frmTraslados.tTX.Text)
                        .TileExit.y = Val(frmTraslados.tTY.Text)

                    End If
                        
                ElseIf frmTraslados.cQuitarTrans.value = True Then
                    MapInfo.Changed = 1 'Set changed flag
                    .TileExit.Map = 0
                    .TileExit.x = 0
                    .TileExit.y = 0

                End If
                
                '########################
                'BLOQUEOS
                '########################
                If frmBloqueos.cInsertarBloqueo.value = True Then
                    If .bLocked <> 1 Then

                        MapInfo.Changed = 1 'Set changed flag
                        .bLocked = 1

                    End If
                    
                ElseIf frmBloqueos.cQuitarBloqueo.value = True Then

                    If .bLocked <> 0 Then

                        MapInfo.Changed = 1 'Set changed flag
                        .bLocked = 0

                    End If

                End If
                
                '########################
                'NPC's
                '########################
                If frmNPCs.cInsertarFunc.value = True Then
                    If frmNPCs.cNumFunc.Text > 0 Then
                        NPCIndex = frmNPCs.cNumFunc.Text

                        If NPCIndex <> .NPCIndex Then

                            MapInfo.Changed = 1 'Set changed flag
                            Body = NpcData(NPCIndex).Body
                            Head = NpcData(NPCIndex).Head
                            Casco = NpcData(NPCIndex).CascoAnim
                            Arma = NpcData(NPCIndex).WeaponAnim
                            Escudo = NpcData(NPCIndex).ShieldAnim
                            Heading = NpcData(NPCIndex).Heading
                            Call Char_Make(NextOpenChar(), Body, Head, Heading, tX, tY, Arma, Escudo, Casco)
                            .NPCIndex = NPCIndex

                        End If

                    End If
                    
                ElseIf frmNPCs.cInsertarFunc.value = True Then

                    If frmNPCs.cNumFunc.Text > 0 Then
                        NPCIndex = frmNPCs.cNumFunc.Text

                        If NPCIndex <> (.NPCIndex) Then

                            MapInfo.Changed = 1 'Set changed flag
                            Body = NpcData(NPCIndex).Body
                            Head = NpcData(NPCIndex).Head
                            Casco = NpcData(NPCIndex).CascoAnim
                            Arma = NpcData(NPCIndex).WeaponAnim
                            Escudo = NpcData(NPCIndex).ShieldAnim
                            Heading = NpcData(NPCIndex).Heading
                            Call Char_Make(NextOpenChar(), Body, Head, Heading, tX, tY, Arma, Escudo, Casco)
                            .NPCIndex = NPCIndex

                        End If

                    End If
                    
                ElseIf frmNPCs.cQuitarFunc.value = True Then

                    If .NPCIndex > 0 Then

                        MapInfo.Changed = 1 'Set changed flag
                        .NPCIndex = 0
                        Call Char_Erase(.CharIndex)

                    End If

                End If
                
                '########################
                'OBJETOS
                '########################
                If frmOBJs.cInsertarFunc.value = True Then ' Insertar Objeto
                    If frmOBJs.cNumFunc.Text > 0 Then
                        ObjIndex = frmOBJs.cNumFunc.Text
                        
                        If .OBJInfo.ObjIndex <> ObjIndex Or .OBJInfo.Amount <> Val(frmOBJs.cCantFunc.Text) Then

                            MapInfo.Changed = 1 'Set changed flag
                            InitGrh .ObjGrh, ObjData(ObjIndex).GrhIndex
                            .OBJInfo.ObjIndex = ObjIndex
                            .OBJInfo.Amount = Val(frmOBJs.cCantFunc.Text)
                            
                            Select Case ObjData(ObjIndex).OBJType

                                Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                    .Graphic(3) = .ObjGrh

                            End Select
                            
                        End If

                    End If

                ElseIf frmOBJs.cQuitarFunc.value = True Then ' Quitar Objeto
                
                    If .OBJInfo.ObjIndex <> 0 Or .OBJInfo.Amount <> 0 Then

                        MapInfo.Changed = 1 'Set changed flag
                        If .Graphic(3).GrhIndex = .ObjGrh.GrhIndex Then .Graphic(3).GrhIndex = 0
                        .ObjGrh.GrhIndex = 0
                        .OBJInfo.ObjIndex = 0
                        .OBJInfo.Amount = 0

                    End If
                    
                End If
                
                '########################
                'TRIGGERS
                '########################
                If frmTriggers.cInsertarTrigger.value = True Then ' Insertar Trigger
                
                    If .Trigger <> frmTriggers.LynxTriggers.CellText(, 0) Then

                        MapInfo.Changed = 1 'Set changed flag
                        .Trigger = frmTriggers.LynxTriggers.CellText(, 0)

                    End If
                    
                ElseIf frmTriggers.cQuitarTrigger.value = True Then ' Quitar Trigger
                
                    If .Trigger <> 0 Then

                        MapInfo.Changed = 1 'Set changed flag
                        .Trigger = 0

                    End If
                    
                End If
                
                '########################
                'PARTICULAS
                '########################
                If frmParticulas.cmdAdd.value = True Then ' Insertar Particle

                    MapInfo.Changed = 1
                    General_Particle_Create CLng(frmParticulas.LynxParticulas.CellText(, 0)), tX, tY, CLng(-1)
                    .Particle_Index = CLng(frmParticulas.LynxParticulas.CellText(, 0))
                    
                ElseIf frmParticulas.cmdDel.value = True Then ' Quitar Particle

                    MapInfo.Changed = 1 'Set changed flag
                    .Particle_Index = 0
                    .Particle_Group_Index = 0
                    
                End If
                
                '########################
                'PARTICULAS
                '########################
                If frmLuces.cInsertarLuz.value Then
                    If Val(frmLuces.cRango = 0) Then Exit Sub
                    
                    Dim TmpColor As RGBA
                    TmpColor = RGBA_From_Comp(Val(frmLuces.R), Val(frmLuces.G), Val(frmLuces.B))
                    
                    Call LucesRedondas.Create_Light_To_Map(tX, tY, TmpColor, Val(frmLuces.cRango))
                    Call LucesRedondas.LightRenderAll
                    
                    With MapData(tX, tY).Light
                        .active = True
                        .range = frmLuces.cRango
                        .RGBCOLOR.A = 255
                        .RGBCOLOR.R = Val(frmLuces.R)
                        .RGBCOLOR.G = Val(frmLuces.G)
                        .RGBCOLOR.B = Val(frmLuces.B)
                        
                    End With
                    
                    MapInfo.Changed = 1 'Set changed flag
                    
                ElseIf frmLuces.cQuitarLuz.value Then
                
                    With MapData(tX, tY).Light
                        .range = 0
                        .RGBCOLOR.A = 255
                        .RGBCOLOR.R = Val(frmLuces.R)
                        .RGBCOLOR.G = Val(frmLuces.G)
                        .RGBCOLOR.B = Val(frmLuces.B)
                        
                    End With
        
                    LucesRedondas.Delete_Light_To_Map tX, tY
        
                    MapInfo.Changed = 1 'Set changed flag
                    
                End If
                
                '########################
                'ZONAS
                '########################
                If frmZonas.LvBPintar.value Then
                    .ZonaIndex = frmZonas.LstZona.ListIndex + 1
                    MapInfo.Changed = 1
                    
                ElseIf frmZonas.LvBQuitar.value Then
                    .ZonaIndex = 0
                    MapInfo.Changed = 1
                    
                End If
            
        End Select
        
    End With
    
    Exit Sub

ClickEdit_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.ClickEdit", Erl)

    Resume Next

End Sub

Public Sub DePegar()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim x As Integer
    Dim y As Integer

    For x = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
             MapData(x + DeSeleccionOX, y + DeSeleccionOY) = DeSeleccionMap(x, y)
        Next
    Next
End Sub

Public Sub PegarSeleccion() '(mx As Integer, my As Integer)
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Static UltimoX As Integer
    Static UltimoY As Integer
    
    Dim x As Integer
    Dim y As Integer
    
    If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    
    EstadoSelect = eEstadoSelect.nada
    
    UltimoX = SobreX
    UltimoY = SobreY
    
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For x = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
            DeSeleccionMap(x, y) = MapData(x + SobreX, y + SobreY)
        Next
    Next
    
    For x = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
             MapData(x + SobreX, y + SobreY).bLocked = SeleccionMap(x, y).bLocked
             MapData(x + SobreX, y + SobreY).CharIndex = SeleccionMap(x, y).CharIndex
             MapData(x + SobreX, y + SobreY).Light_Value(0) = SeleccionMap(x, y).Light_Value(0)
             MapData(x + SobreX, y + SobreY).Light_Value(1) = SeleccionMap(x, y).Light_Value(1)
             MapData(x + SobreX, y + SobreY).Light_Value(2) = SeleccionMap(x, y).Light_Value(2)
             MapData(x + SobreX, y + SobreY).Light_Value(3) = SeleccionMap(x, y).Light_Value(3)
             MapData(x + SobreX, y + SobreY).fX = SeleccionMap(x, y).fX
             MapData(x + SobreX, y + SobreY).FxIndex = SeleccionMap(x, y).FxIndex
             MapData(x + SobreX, y + SobreY).Graphic(1) = SeleccionMap(x, y).Graphic(1)
             MapData(x + SobreX, y + SobreY).Graphic(2) = SeleccionMap(x, y).Graphic(2)
             MapData(x + SobreX, y + SobreY).Graphic(3) = SeleccionMap(x, y).Graphic(3)
             MapData(x + SobreX, y + SobreY).Graphic(4) = SeleccionMap(x, y).Graphic(4)
             MapData(x + SobreX, y + SobreY).Light = SeleccionMap(x, y).Light
             MapData(x + SobreX, y + SobreY).NPCIndex = SeleccionMap(x, y).NPCIndex
             MapData(x + SobreX, y + SobreY).ObjGrh = SeleccionMap(x, y).ObjGrh
             MapData(x + SobreX, y + SobreY).OBJInfo = SeleccionMap(x, y).OBJInfo
             MapData(x + SobreX, y + SobreY).Particle_Group_Index = SeleccionMap(x, y).Particle_Group_Index
             MapData(x + SobreX, y + SobreY).Particle_Index = SeleccionMap(x, y).Particle_Index
             MapData(x + SobreX, y + SobreY).TileExit = SeleccionMap(x, y).TileExit
             MapData(x + SobreX, y + SobreY).Trigger = SeleccionMap(x, y).Trigger
             
             If frmMain.mnuCopiarZonas.Checked Then _
                MapData(x + SobreX, y + SobreY).ZonaIndex = SeleccionMap(x, y).ZonaIndex
        Next
    Next
    Seleccionando = False
End Sub

Public Sub AccionSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************

    Dim x As Integer
    Dim y As Integer
    
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For x = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(x, y) = MapData(x + SeleccionIX, y + SeleccionIY)
        Next
    Next
    
    For x = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
           ClickEdit vbLeftButton, SeleccionIX + x, SeleccionIY + y
        Next
    Next
    Seleccionando = False
End Sub

Public Sub CortarSeleccion()
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    
    On Error GoTo CortarSeleccion_Err
    
    CopiarSeleccion
    Dim x     As Integer
    Dim y     As Integer
    Dim vacio As MapBlock
    
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For x = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(x, y) = MapData(x + SeleccionIX, y + SeleccionIY)
        Next
    Next

    For x = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            MapData(x + SeleccionIX, y + SeleccionIY) = vacio
        Next
    Next
    
    Seleccionando = False

    Exit Sub

CortarSeleccion_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.CortarSeleccion", Erl)
    Resume Next
    
End Sub

Public Sub CopiarSeleccion(Optional ByVal Borde As Boolean = False)
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    'podria usar copy mem , pero por las dudas no XD
    
    On Error GoTo CopiarSeleccion_Err
    
    Dim x As Integer
    Dim y As Integer
    Dim i As Byte
    
    EstadoSelect = eEstadoSelect.Copiado
    
    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock
    
    If Not Borde Then

        For x = 0 To SeleccionAncho - 1
            For y = 0 To SeleccionAlto - 1
                SeleccionMap(x, y) = MapData(x + SeleccionIX, y + SeleccionIY)
            Next
        Next
        
    Else
    
        For x = 0 To SeleccionAncho - 1
            For y = 0 To SeleccionAlto - 1

                With SeleccionMap(x, y)
                    .bLocked = MapData(x + SeleccionIX, y + SeleccionIY).bLocked
                    .Trigger = MapData(x + SeleccionIX, y + SeleccionIY).Trigger
                    .Particle_Index = MapData(x + SeleccionIX, y + SeleccionIY).Particle_Index
                    .Particle_Group_Index = MapData(x + SeleccionIX, y + SeleccionIY).Particle_Group_Index
                    .OBJInfo = MapData(x + SeleccionIX, y + SeleccionIY).OBJInfo
                    .ObjGrh = MapData(x + SeleccionIX, y + SeleccionIY).ObjGrh
                    .NPCIndex = MapData(x + SeleccionIX, y + SeleccionIY).NPCIndex
                    .Light = MapData(x + SeleccionIX, y + SeleccionIY).Light

                    For i = 1 To 4
                        .Graphic(i) = MapData(x + SeleccionIX, y + SeleccionIY).Graphic(i)
                    Next i

                    .FxIndex = MapData(x + SeleccionIX, y + SeleccionIY).FxIndex
                    .fX = MapData(x + SeleccionIX, y + SeleccionIY).fX

                    For i = 0 To 3
                        .Light_Value(i) = MapData(x + SeleccionIX, y + SeleccionIY).Light_Value(i)
                    Next i

                    .CharIndex = MapData(x + SeleccionIX, y + SeleccionIY).CharIndex
                End With
            Next
        Next
        
    End If
    
    Exit Sub

CopiarSeleccion_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.CopiarSeleccion", Erl)
    Resume Next
End Sub


