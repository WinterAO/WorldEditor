Attribute VB_Name = "modEdicion"
Option Explicit

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
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If Capa = 1 Then
                MapData(X, y).Graphic(Capa).GrhIndex = 1
            Else
                MapData(X, y).Graphic(Capa).GrhIndex = 0
            End If
        Next X
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
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            If frmConfigSup.MOSAICO.value = vbChecked Then
                Dim aux As Integer
                aux = Val(frmSuperficies.cGrh.Text) + _
                ((y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (X Mod frmConfigSup.mAncho)
                 MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                'Setup GRH
                InitGrh MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)), aux
            Else
                'Else Place graphic
                MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = Val(frmSuperficies.cGrh.Text)
                'Setup GRH
                InitGrh MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)), Val(frmSuperficies.cGrh.Text)
            End If
    
        Next X
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
    Dim X As Long
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
                MapData(X, y).bLocked = 1
            End If
        Next X
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
    Dim X As Long
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
                MapData(X, y).bLocked = 0
            End If
        Next X
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
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            MapData(X, y).bLocked = Valor
        Next X
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
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
            
                MapData(X, y).Graphic(1).GrhIndex = 1
                InitGrh MapData(X, y).Graphic(1), 1
                MapData(X, y).bLocked = 0
                
                 'Erase NPCs
                If MapData(X, y).NPCIndex > 0 Then
                    EraseChar MapData(X, y).CharIndex
                    MapData(X, y).NPCIndex = 0
                End If
    
                'Erase Objs
                MapData(X, y).OBJInfo.ObjIndex = 0
                MapData(X, y).OBJInfo.Amount = 0
                MapData(X, y).ObjGrh.GrhIndex = 0
    
                'Clear exits
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0
                
                ' Triggers
                MapData(X, y).Trigger = 0
    
            End If
    
        Next X
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
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    
              If frmConfigSup.MOSAICO.value = vbChecked Then
                Dim aux As Integer
                aux = Val(frmSuperficies.cGrh.Text) + _
                ((y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (X Mod frmConfigSup.mAncho)
                If frmBloqueos.cInsertarBloqueo.value = True Then
                    MapData(X, y).bLocked = 1
                Else
                    MapData(X, y).bLocked = 0
                End If
                MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                'Setup GRH
                InitGrh MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)), aux
              Else
                'Else Place graphic
                If frmBloqueos.cInsertarBloqueo.value = True Then
                    MapData(X, y).bLocked = 1
                Else
                    MapData(X, y).bLocked = 0
                End If
                
                MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = Val(frmSuperficies.cGrh.Text)
                
                'Setup GRH
        
                InitGrh MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)), Val(frmSuperficies.cGrh.Text)
            End If
                 'Erase NPCs
                If MapData(X, y).NPCIndex > 0 Then
                    EraseChar MapData(X, y).CharIndex
                    MapData(X, y).NPCIndex = 0
                End If
    
                'Erase Objs
                MapData(X, y).OBJInfo.ObjIndex = 0
                MapData(X, y).OBJInfo.Amount = 0
                MapData(X, y).ObjGrh.GrhIndex = 0
    
                'Clear exits
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0
    
            End If
    
        Next X
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
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    
                MapData(X, y).ZonaIndex = frmZonas.LstZona.ListIndex + 1
    
            End If
    
        Next X
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Elimita los NPCs del mapa
'
' @param Hostiles Indica si elimita solo hostiles o solo npcs no hostiles

Public Sub Quitar_NPCs(ByVal Hostiles As Boolean)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Quitar_NPCs_Err

    Dim y As Integer
    Dim X As Integer

    If EditWarning Then Exit Sub

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            If Not Hostiles Then
                If MapData(X, y).NPCIndex > 0 Then
                    Call EraseChar(MapData(X, y).CharIndex)
                    MapData(X, y).NPCIndex = 0
    
                End If
            Else
                    If MapData(X, y).NPCIndex > 500 Then
                    Call EraseChar(MapData(X, y).CharIndex)
                    MapData(X, y).NPCIndex = 0
    
                End If
            End If
        
        Next X
    Next y

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

Private Function EditWarning() As Boolean
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
