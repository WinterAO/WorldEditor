Attribute VB_Name = "modMinimapa"
Option Explicit

Public MMiniMap_capa1      As Boolean
Public MMiniMap_capa2      As Boolean
Public MMiniMap_capa3      As Boolean
Public MMiniMap_capa4      As Boolean
Public MMiniMap_Npcs       As Boolean
Public MMiniMap_objetos    As Boolean
Public MMiniMap_Bloqueos   As Boolean
Public MMiniMap_particulas As Boolean
Public MMiniMap_Nombre     As Boolean

Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Sub DibujarMinimapa(Optional ByVal Refrescar = False)

    Dim map_x As Integer
    Dim map_y As Integer
    Dim XMin As Long
    Dim XMax As Long
    Dim YMin As Long
    Dim YMax As Long
    
    Dim picMapahDC As Long
    
    'Dibujamos todo el minimapa al completo
    If Refrescar = False Then
    
        frmMapa.picMapa.BackColor = vbBlack
        picMapahDC = frmMapa.picMapa.hDC
            
        XMin = XMinMapSize
        XMax = XMaxMapSize
        YMin = YMinMapSize
        YMax = YMaxMapSize

        For map_y = YMin To YMax
            For map_x = YMin To XMax
            
                If MMiniMap_capa1 Then
                    If MapData(map_x, map_y).Graphic(1).GrhIndex > 0 Then
                        SetPixel picMapahDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).mini_map_color
    
                    End If
    
                End If
                
                If MMiniMap_capa2 Then
                    If MapData(map_x, map_y).Graphic(2).GrhIndex > 0 Then
                        SetPixel picMapahDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(2).GrhIndex).mini_map_color
    
                    End If
    
                End If
            
                If MMiniMap_capa3 Then
                    If MapData(map_x, map_y).Graphic(3).GrhIndex > 0 Then
                        SetPixel picMapahDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(3).GrhIndex).mini_map_color
    
                    End If
    
                End If
            
                If MMiniMap_capa4 Then
                    If MapData(map_x, map_y).Graphic(4).GrhIndex > 0 Then
                        SetPixel picMapahDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).mini_map_color
    
                    End If
    
                End If
            
                If MMiniMap_Npcs Then
                    If MapData(map_x, map_y).NPCIndex > 0 Then
                        SetPixel picMapahDC, map_x - 1, map_y - 1, vbYellow
    
                    End If
    
                End If
            
                If MMiniMap_objetos Then
                    If MapData(map_x, map_y).OBJInfo.ObjIndex > 0 Then
                        SetPixel picMapahDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).ObjGrh.GrhIndex).mini_map_color
    
                    End If
    
                End If
            
                If MMiniMap_Bloqueos Then
                    If MapData(map_x, map_y).Blocked > 0 Then
                        SetPixel picMapahDC, map_x - 1, map_y - 1, vbRed
    
                    End If
    
                End If
            
                If MMiniMap_particulas Then
                    If MapData(map_x, map_y).Particle_Index > 0 Then
                        SetPixel picMapahDC, map_x - 1, map_y - 1, vbWhite
    
                    End If
    
                End If
    
                If MMiniMap_Nombre Then
                    frmMapa.picMapa.CurrentX = 30
                    frmMapa.picMapa.CurrentY = 26
                    'frmMapa.picMapa.Print frmMain.MapPest(7).Caption
    
                End If
            Next map_x
        Next map_y
    
    End If
    
    'frmMain.UserM.Left = (UserPos.X * 2) - 2
    'frmMain.UserM.Top = (UserPos.Y * 2) - 2
    frmMapa.ApuntadorRadar.Left = (UserPos.X) - 9
    frmMapa.ApuntadorRadar.Top = (UserPos.Y) - 8
    
    'Refrescamos
    'frmMain.Minimap.Refresh
End Sub

Public Sub ActualizarMinimapa(ByVal tX As Integer, ByVal tY As Integer)

    If tY < YMinMapSize Or tY > YMaxMapSize Then Exit Sub
    If tX < XMinMapSize Or tX > XMaxMapSize Then Exit Sub
    
'    If frmMain.cSeleccionarSuperficie.value = True Then
'        If MapData(tX, tY).Graphic(Val(frmMain.cCapas.Text)).GrhIndex > 0 Then _
'            SetPixel frmMapa.picMapa.hDC, tX - 1, tY - 1, GrhData(MapData(tX, tY).Graphic(Val(frmMain.cCapas.Text)).GrhIndex).mini_map_color
'
'    ElseIf frmMain.cQuitarEnEstaCapa.value = True Then
'        SetPixel frmMapa.picMapa.hDC, tX - 1, tY - 1, 0
'
'    End If

End Sub

