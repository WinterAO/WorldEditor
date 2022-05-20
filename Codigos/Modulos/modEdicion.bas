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
Dim y As Integer
Dim X As Integer

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
Dim y As Integer
Dim X As Integer

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
