Attribute VB_Name = "modEdit"
Option Explicit

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
    
    If frmMain.Dialog.FilterIndex = 1 Then
        Formato = ".csm"
    ElseIf frmMain.Dialog.FilterIndex = 2 Then
        Formato = ".map"
    End If
    
    ' Translados
    Dim tTrans As WorldPos
    tTrans = MapData(tX, tY).TileExit
    If tTrans.Map > 0 Then
        If LenB(frmMain.Dialog.filename) <> 0 Then
            If FileExist(PATH_Save & NameMap_Save & tTrans.Map & Formato, vbArchive) = True Then
                Call modMapas.NuevoMapa
                frmMain.Dialog.filename = PATH_Save & NameMap_Save & tTrans.Map & Formato
                
                Select Case frmMain.Dialog.FilterIndex
                
                    Case 1
                        If ClientSetup.MeMode = eMeMode.WinterAO Then
                            Call modMapas.Cargar_CSM(frmMain.Dialog.filename)
                        End If
                
                End Select
                
                'modMapIO.AbrirunMapa frmMain.Dialog.filename
                UserPos.X = tTrans.X
                UserPos.Y = tTrans.Y
                
                If WalkMode = True Then
                    MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
                    CharList(UserCharIndex).Heading = SOUTH
                End If
                
                'frmMain.mnuReAbrirMapa.Enabled = True
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
    Dim Head As Integer
    Dim Body As Integer
    Dim Heading As Byte
    Dim loopc As Integer

    If tY < YMinMapSize Or tY > YMaxMapSize Then Exit Sub
    If tX < XMinMapSize Or tX > XMaxMapSize Then Exit Sub

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
                
                ' NPCs
'                If .NPCIndex > 0 Then
'                    If .NPCIndex > 499 Then
'                        Call AddtoRichTextBox(frmConsola.StatTxt, " (NPC-Hostil: " & .NPCIndex & " - " & NpcData(.NPCIndex).name & ")", 255, 255, 255, False, False, True)
'
'                    Else
'                        Call AddtoRichTextBox(frmConsola.StatTxt, " (NPC: " & .NPCIndex & " - " & NpcData(.NPCIndex).name & ")", 255, 255, 255, False, False, True)
'
'                    End If
'                End If
                
                ' OBJs
                'If .OBJInfo.ObjIndex > 0 Then _
                    Call AddtoRichTextBox(frmConsola.StatTxt, " (Obj: " & .OBJInfo.ObjIndex & " - " & ObjData(.OBJInfo.ObjIndex).name & " - Cant.:" & .OBJInfo.Amount & ")", 255, 255, 255, False, False, True)
            
                ' Capas
                Call AddtoRichTextBox(frmConsola.StatTxt, "Capa1: " & .Graphic(1).GrhIndex & " - Capa2: " & .Graphic(2).GrhIndex & " - Capa3: " & .Graphic(3).GrhIndex & " - Capa4: " & .Graphic(4).GrhIndex, 255, 255, 255, False, False, True)
'                If frmMain.mnuAutoCapturarSuperficie.Checked = True And frmMain.cSeleccionarSuperficie.value = False Then
'                    If .Graphic(4).GrhIndex <> 0 Then
'                        frmMain.cCapas.Text = 4
'                        frmMain.cGrh.Text = .Graphic(4).GrhIndex
'
'                    ElseIf .Graphic(3).GrhIndex <> 0 Then
'                        frmMain.cCapas.Text = 3
'                        frmMain.cGrh.Text = .Graphic(3).GrhIndex
'
'                    ElseIf .Graphic(2).GrhIndex <> 0 Then
'                        frmMain.cCapas.Text = 2
'                        frmMain.cGrh.Text = .Graphic(2).GrhIndex
'
'                    ElseIf .Graphic(1).GrhIndex <> 0 Then
'                        frmMain.cCapas.Text = 1
'                        frmMain.cGrh.Text = .Graphic(1).GrhIndex
'
'                    End If
'                End If
                
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
                
                'Insertar Superficie
                If frmSuperficies.cSeleccionarSuperficie.value = True Then
                    
                    If frmConfigSup.MOSAICO.value = vbChecked Then
                        Dim aux As Long
                        Dim dy As Integer
                        Dim dX As Integer
                        
                        If frmConfigSup.DespMosaic.value = vbChecked Then
                            dy = Val(frmConfigSup.DMLargo)
                            dX = Val(frmConfigSup.DMAncho.Text)
                            
                        Else
                            dy = 0
                            dX = 0
                            
                        End If
                            
                        If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                            MapInfo.Changed = 1 'Set changed flag
                            aux = Val(frmSuperficies.cGrh.Text) + _
                            (((tY + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((tX + dX) Mod frmConfigSup.mAncho.Text)
                            
                             If .Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex <> aux Or .bLocked <> frmMain.LvBEdit(2).value Then
                                .Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                                InitGrh .Graphic(Val(frmSuperficies.cCapas.Text)), aux
                                
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
                                    
                                        If tXX >= XMinMapSize And tXX <= XMaxMapSize Then
                                        
                                            aux = Val(frmSuperficies.cGrh.Text) + desptile
                                            MapData(tXX, tYY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                                            InitGrh MapData(tXX, tYY).Graphic(Val(frmSuperficies.cCapas.Text)), aux
                                            tXX = tXX + 1
                                            desptile = desptile + 1
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
                        If .bLocked <> frmMain.LvBEdit(2).value Or .Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex <> Val(frmSuperficies.cGrh.Text) Then
                            MapInfo.Changed = 1 'Set changed flag
                            .Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = Val(frmSuperficies.cGrh.Text)
                            'Setup GRH
                            InitGrh .Graphic(Val(frmSuperficies.cCapas.Text)), Val(frmSuperficies.cGrh.Text)
                            
                        End If
                        
                    End If
                    
                End If
                
                '########################
                'TRASLADOS
                '########################
                If frmTraslados.cInsertarTrans.value = True Then
'                    If Cfg_TrOBJ > 0 And Cfg_TrOBJ <= 0 And frmMain.cInsertarTransOBJ.value = True Then
'                        If ObjData(Cfg_TrOBJ).ObjType = 19 Then
'                            MapInfo.Changed = 1 'Set changed flag
'                            InitGrh .ObjGrh, ObjData(Cfg_TrOBJ).GrhIndex
'                            .OBJInfo.ObjIndex = Cfg_TrOBJ
'                            .OBJInfo.Amount = 1
'                        End If
'                    End If
                    
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
                            .TileExit.X = 12
                            .TileExit.Y = tY
                            
                        ElseIf tX <= 11 Then ' 9 ' izquierda
                            .TileExit.X = 91
                            .TileExit.Y = tY
                            
                        End If
                        
                        If tY >= 91 Then ' 94 '''' hacia abajo
                            .TileExit.Y = 11
                            .TileExit.X = tX
                            
                        ElseIf tY <= 10 Then ''' hacia arriba
                            .TileExit.Y = 90
                            .TileExit.X = tX
                            
                        End If
                        
                    Else
                        MapInfo.Changed = 1 'Set changed flag
                        .TileExit.Map = Val(frmTraslados.tTMapa.Text)
                        .TileExit.X = Val(frmTraslados.tTX.Text)
                        .TileExit.Y = Val(frmTraslados.tTY.Text)
                    End If
                        
                ElseIf frmTraslados.cQuitarTrans.value = True Then
                    MapInfo.Changed = 1 'Set changed flag
                    .TileExit.Map = 0
                    .TileExit.X = 0
                    .TileExit.Y = 0
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
                            Heading = NpcData(NPCIndex).Heading
                            Call MakeChar(NextOpenChar(), Body, Head, Heading, tX, tY)
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
                            Heading = NpcData(NPCIndex).Heading
                            Call MakeChar(NextOpenChar(), Body, Head, Heading, tX, tY)
                            .NPCIndex = NPCIndex
                        End If
                    End If
                    
                ElseIf frmNPCs.cQuitarFunc.value = True Then
                    If .NPCIndex > 0 Then
                        MapInfo.Changed = 1 'Set changed flag
                        .NPCIndex = 0
                        Call EraseChar(.CharIndex)
                    End If
                End If
                
                '########################
                'OBJETOS
                '########################
                If frmObjs.cInsertarFunc.value = True Then ' Insertar Objeto
                    If frmObjs.cNumFunc.Text > 0 Then
                        ObjIndex = frmObjs.cNumFunc.Text
                        
                        If .OBJInfo.ObjIndex <> ObjIndex Or MapData(tX, tY).OBJInfo.Amount <> Val(frmObjs.cCantFunc.Text) Then
                            MapInfo.Changed = 1 'Set changed flag
                            InitGrh MapData(tX, tY).ObjGrh, ObjData(ObjIndex).GrhIndex
                            .OBJInfo.ObjIndex = ObjIndex
                            .OBJInfo.Amount = Val(frmObjs.cCantFunc.Text)
                            
                            Select Case ObjData(ObjIndex).ObjType
                                Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                    .Graphic(3) = .ObjGrh
                            End Select
                            
                        End If
                    End If
                ElseIf frmObjs.cQuitarFunc.value = True Then ' Quitar Objeto
                
                    If .OBJInfo.ObjIndex <> 0 Or .OBJInfo.Amount <> 0 Then
                        MapInfo.Changed = 1 'Set changed flag
                        
                        If .Graphic(3).GrhIndex = .ObjGrh.GrhIndex Then .Graphic(3).GrhIndex = 0
                        
                        .ObjGrh.GrhIndex = 0
                        .OBJInfo.ObjIndex = 0
                        .OBJInfo.Amount = 0
                    End If
                    
                End If
            
        End Select
        
    End With
    
    Exit Sub

ClickEdit_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdit.ClickEdit", Erl)
    Resume Next

End Sub
