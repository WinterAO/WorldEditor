Attribute VB_Name = "mDx8_Luces"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 14/05/10
'Blisse-AO | Light Engine, Read the _
    #LightEngine to Set the type of Lights
'***************************************************

Option Base 0

Private Type tLight
    RGBCOLOR As D3DCOLORVALUE
    active As Boolean
    map_x As Integer
    map_y As Integer
    range As Byte
End Type
 
Private Light_List() As tLight
Private NumLights As Integer

Public Enum e_estados
    Amanecer = 0
    MedioDia
    Tarde
    noche
End Enum

Public Estados(0 To 3) As D3DCOLORVALUE
Public Estado_Actual As D3DCOLORVALUE
Public Estado_Custom As D3DCOLORVALUE

Public Function Create_Light_To_Map(ByVal map_x As Integer, ByVal map_y As Integer, Optional range As Byte = 3, Optional ByVal Red As Byte = 255, Optional ByVal Green As Byte = 255, Optional ByVal Blue As Byte = 255)
    NumLights = NumLights + 1
   
    ReDim Preserve Light_List(1 To NumLights) As tLight
   
    Light_List(NumLights).RGBCOLOR.R = Red
    Light_List(NumLights).RGBCOLOR.G = Green
    Light_List(NumLights).RGBCOLOR.B = Blue
    Light_List(NumLights).RGBCOLOR.a = 255
    Light_List(NumLights).range = range
    Light_List(NumLights).active = True
    Light_List(NumLights).map_x = map_x
    Light_List(NumLights).map_y = map_y
   
    Call LightRender(NumLights)
End Function

Public Function Delete_Light_To_Map(ByVal X As Integer, ByVal Y As Integer)
   
    Dim i As Long
   
    For i = 1 To NumLights
        If Light_List(i).map_x = X And Light_List(i).map_y = Y Then
            Delete_Light_To_Index i
            
            Exit Function
        End If
    Next i
 
End Function

Public Function Delete_Light_To_Index(ByVal light_index As Integer)
'************************************
'Autor: Lorwik
'Fecha: 14/08/2020
'Descripción: Primero desactivamos una luz concreta y luego reordenamos el array
'************************************

    'Borramos la luz
    Light_List(light_index).active = False
    'Reordamos el Aray
    If light_index = NumLights Then
        Do Until Light_List(NumLights).active
            NumLights = NumLights - 1
            If NumLights = 0 Then
                Call Actualizar_Estado
                Exit Function
                
            End If
        Loop
        ReDim Preserve Light_List(1 To NumLights)
    
    End If
 
    Call Actualizar_Estado

End Function

#If LightEngine = 1 Then '   Luces Radiales

Private Sub LightRender(ByVal light_index As Integer)
 
 On Local Error Resume Next
 
    If light_index = 0 Then Exit Sub
    If Light_List(light_index).active = False Then Exit Sub
   
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim Ya As Integer
    Dim Xa As Integer
   
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE
   
    Dim XCoord As Integer
    Dim YCoord As Integer
   
    AmbientColor.a = Estado_Actual.a
    AmbientColor.R = Estado_Actual.R
    AmbientColor.G = Estado_Actual.G
    AmbientColor.B = Estado_Actual.B

    LightColor = Light_List(light_index).RGBCOLOR
       
    min_x = Light_List(light_index).map_x - Light_List(light_index).range
    max_x = Light_List(light_index).map_x + Light_List(light_index).range
    min_y = Light_List(light_index).map_y - Light_List(light_index).range
    max_y = Light_List(light_index).map_y + Light_List(light_index).range
    
    Dim TEMP_COLOR As D3DCOLORVALUE
    
    For Ya = min_y To max_y
        For Xa = min_x To max_x
            If InMapBounds(Xa, Ya) Then
                XCoord = Xa * 32
                YCoord = Ya * 32
                Call Engine_Get_ARGB(MapData(Xa, Ya).Engine_Light(0), TEMP_COLOR)
                
                MapData(Xa, Ya).Engine_Light(0) = LightCalculate(Light_List(light_index).range, Light_List(light_index).map_x * 32, Light_List(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).Engine_Light(0), LightColor, TEMP_COLOR)
 
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32
                Call Engine_Get_ARGB(MapData(Xa, Ya).Engine_Light(3), TEMP_COLOR)
                                
                MapData(Xa, Ya).Engine_Light(3) = LightCalculate(Light_List(light_index).range, Light_List(light_index).map_x * 32, Light_List(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).Engine_Light(3), LightColor, TEMP_COLOR)
                       
                XCoord = Xa * 32
                YCoord = Ya * 32 + 32
                Call Engine_Get_ARGB(MapData(Xa, Ya).Engine_Light(1), TEMP_COLOR)
                MapData(Xa, Ya).Engine_Light(1) = LightCalculate(Light_List(light_index).range, Light_List(light_index).map_x * 32, Light_List(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).Engine_Light(1), LightColor, TEMP_COLOR)
   
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32 + 32
                Call Engine_Get_ARGB(MapData(Xa, Ya).Engine_Light(2), TEMP_COLOR)
                MapData(Xa, Ya).Engine_Light(2) = LightCalculate(Light_List(light_index).range, Light_List(light_index).map_x * 32, Light_List(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).Engine_Light(2), LightColor, TEMP_COLOR)
               
            End If
        Next Xa
    Next Ya
End Sub

Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim XDist As Single
    Dim YDist As Single
    Dim VertexDist As Single
    Dim pRadio As Integer
   
    Dim CurrentColor As D3DCOLORVALUE
   
    pRadio = cRadio * 32
   
    XDist = LightX + 16 - XCoord
    YDist = LightY + 16 - YCoord
   
    VertexDist = Sqr(XDist * XDist + YDist * YDist)
   
    If VertexDist <= pRadio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, VertexDist / pRadio) 'aca hay algo mal ;) Ambient color ;)
        LightCalculate = D3DColorXRGB(Round(CurrentColor.R), Round(CurrentColor.G), Round(CurrentColor.B))
    Else
        LightCalculate = TileLight
    End If
End Function

#Else 'Luces Normales

Private Sub LightRender(ByVal light_index As Integer)

    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim ia As Single
    Dim i As Integer
    Dim color(3) As Long
    Dim Ya As Integer
    Dim Xa As Integer

    Dim XCoord As Integer
    Dim YCoord As Integer
    
    With Light_List(light_index)
    
        color(0) = D3DColorARGB(255, .RGBCOLOR.R, .RGBCOLOR.G, .RGBCOLOR.B)
        color(1) = color(0)
        color(2) = color(0)
        color(3) = color(0)
    
        'Set up light borders
        min_x = .map_x - .range
        min_y = .map_y - .range
        max_x = .map_x + .range
        max_y = .map_y + .range
    
    End With
    
    'Arrange corners
    
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).Engine_Light(2) = color(2)
    End If
    
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).Engine_Light(1) = color(1)
    End If
    
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).Engine_Light(0) = color(0)
    End If
    
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).Engine_Light(3) = color(3)
    End If
    
    'Arrange borders
    
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).Engine_Light(1) = color(1)
            MapData(X, min_y).Engine_Light(2) = color(2)
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).Engine_Light(0) = color(0)
            MapData(X, max_y).Engine_Light(3) = color(3)
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).Engine_Light(2) = color(2)
            MapData(min_x, Y).Engine_Light(3) = color(3)
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).Engine_Light(0) = color(0)
            MapData(max_x, Y).Engine_Light(1) = color(1)
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).Engine_Light(0) = color(0)
                MapData(X, Y).Engine_Light(1) = color(1)
                MapData(X, Y).Engine_Light(2) = color(2)
                MapData(X, Y).Engine_Light(3) = color(3)
            End If
        Next Y
    Next X
    
    
End Sub

#End If 'Terminamos de Seleccionar las luces

Public Sub DeInit_LightEngine()
    'Kill Font's
    Erase Light_List()
    
    'Exit, The works is done.
    Exit Sub
End Sub

Public Function LightRenderAll()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim i As Long

    If NumLights = 0 Then Exit Function

    For i = 1 To UBound(Light_List)
        LightRender i
    Next i

End Function

Public Sub LightRemoveAll()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************

    Dim i As Long
    
    If NumLights = 0 Then Exit Sub
    Debug.Print "Numero de luces: " & NumLights & "-" & UBound(Light_List)
    For i = 1 To UBound(Light_List)
        Delete_Light_To_Index i
    Next i

End Sub

Public Sub Init_MeteoEngine()
'***************************************************
'Author: Standelf
'Last Modification: 15/05/10
'Initializate
'***************************************************
    With Estados(e_estados.Amanecer)
        .a = 255
        .R = 230
        .G = 200
        .B = 200
    End With
    
    With Estados(e_estados.MedioDia)
        .a = 255
        .R = 255
        .G = 255
        .B = 255
    End With
    
    With Estados(e_estados.Tarde)
        .a = 255
        .R = 200
        .G = 200
        .B = 200
    End With
  
    With Estados(e_estados.noche)
        .a = 255
        .R = 165
        .G = 165
        .B = 165
    End With
    
End Sub

Public Sub Actualizar_Estado()
'***************************************************
'Author: Lorwik
'Last Modification: 09/08/2020
'Actualiza el estado del clima y del dia
'***************************************************
    Dim X As Integer, Y As Integer
    Dim tR As Byte
    Dim tG As Byte
    Dim tB As Byte
    
    '******************
    'MODO WINTER
    '******************
    If ClientSetup.MeMode = eMeMode.WinterAO Then
    
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                If MapZonas(MapData(X, Y).ZonaIndex).LuzBase <> 0 Then
                
                    Call ConvertLongToRGB(MapZonas(MapData(X, Y).ZonaIndex).LuzBase, tR, tG, tB)
                    
                    With Estado_Custom
                        .a = 255
                        .R = tR
                        .G = tG
                        .B = tB
                    End With
                    
                    Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Custom)
            
                Else
                
                    Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
                
                End If
            Next Y
        Next X
        
    Else
    '******************
    'MODO IMPC
    '******************
    
        '¿El mapa tiene su propia luz?
        If MapInfo.LuzBase <> 0 Then
        
            Call ConvertLongToRGB(MapZonas(MapData(X, Y).ZonaIndex).LuzBase, tR, tG, tB)
                    
            With Estado_Custom
                .a = 255
                .R = tR
                .G = tG
                .B = tB
            End With
        
            For X = XMinMapSize To XMaxMapSize
                For Y = YMinMapSize To YMaxMapSize
                    Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Custom)
                Next Y
            Next X
            
            Call LightRenderAll
            
            Exit Sub
        End If
            
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
            Next Y
        Next X
        
    End If
    
    Call LightRenderAll

End Sub



