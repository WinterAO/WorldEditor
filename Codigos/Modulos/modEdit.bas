Attribute VB_Name = "modEdit"
Option Explicit

''
' Acciona la operacion al hacer doble click en una posicion del mapa
'
' @param tX Especifica la posicion X en el mapa
' @param tY Espeficica la posicion Y en el mapa

Sub DobleClick(tX As Integer, tY As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
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
