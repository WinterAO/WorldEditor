VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLucesRedondas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tLight
    RGBCOLOR As RGBA
    active As Boolean
    map_x As Integer
    map_y As Integer
    range As Byte
End Type
 
Private Light_List() As tLight

Private NumLights As Integer

Friend Function Create_Light_To_Map(ByVal map_x As Integer, ByVal map_y As Integer, color As RGBA, Optional range As Byte = 3, Optional ByVal Render As Boolean = True)
    
    On Error GoTo Create_Light_To_Map_Err
    
    NumLights = NumLights + 1
   
    ReDim Preserve Light_List(1 To NumLights) As tLight
   
    'Le damos color
    Light_List(NumLights).RGBCOLOR = color
       
    'Alpha (Si borras esto RE KB!!)
    Light_List(NumLights).RGBCOLOR.A = 255
    
    'Un rango determinado
    Light_List(NumLights).range = range
       
    'La activamos para que se pueda renderizar
    Light_List(NumLights).active = True
       
    'Pos en el mapa
    Light_List(NumLights).map_x = map_x
    Light_List(NumLights).map_y = map_y
   
    If Render Then _
        Call LightRender(NumLights)
        
    Exit Function

Create_Light_To_Map_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsLucesRedondas.Create_Light_To_Map", Erl)
    Resume Next
        
End Function

Public Function Delete_Light_To_Map(ByVal x As Integer, ByVal y As Integer)
   
    Dim i As Long
   
    For i = 1 To NumLights
        If Light_List(i).map_x = x And Light_List(i).map_y = y Then
            Delete_Light_To_Index i
            
            Exit Function
        End If
    Next i
 
End Function

Public Function Delete_Light_To_Index(ByVal light_index As Integer, Optional ByVal Actualizar As Boolean = True)
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
                Exit Function
            End If
        Loop
        ReDim Preserve Light_List(1 To NumLights)
    
    End If
 
    If Actualizar Then _
        Call Actualizar_Estado

End Function

Private Sub LightRender(ByVal light_index As Integer)
 
    On Error GoTo LightRender_Err
    
    If light_index = 0 Then Exit Sub
    If Light_List(light_index).active = False Then Exit Sub
       
    Dim min_x        As Long

    Dim min_y        As Long

    Dim max_x        As Long

    Dim max_y        As Long

    Dim color        As Long

    Dim Ya           As Long

    Dim Xa           As Long
       
    Dim TileLight    As RGBA

    Dim LightColor   As RGBA
       
    Dim XCoord       As Integer

    Dim YCoord       As Integer
    
    Dim LightX       As Integer
    
    Dim LightY       As Integer
    
    Dim radius       As Long
    
    Dim range        As Integer

    LightColor = Light_List(light_index).RGBCOLOR
    
    LightX = Light_List(light_index).map_x * 32 + 16
    LightY = Light_List(light_index).map_y * 32 + 16

    range = Light_List(light_index).range
    
    radius = range * 32 + 16
    radius = radius * radius
           
    min_x = Light_List(light_index).map_x - range
    max_x = Light_List(light_index).map_x + range
    min_y = Light_List(light_index).map_y - range
    max_y = Light_List(light_index).map_y + range

    
    For Ya = min_y To max_y
        For Xa = min_x To max_x

            If InMapBounds(Xa, Ya) Then
                XCoord = Xa * 32
                YCoord = Ya * 32
                MapData(Xa, Ya).Light_Value(0) = LightCalculate(radius, LightX, LightY, XCoord, YCoord, MapData(Xa, Ya).Light_Value(0), LightColor)
     
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32
                MapData(Xa, Ya).Light_Value(3) = LightCalculate(radius, LightX, LightY, XCoord, YCoord, MapData(Xa, Ya).Light_Value(3), LightColor)
                           
                XCoord = Xa * 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).Light_Value(1) = LightCalculate(radius, LightX, LightY, XCoord, YCoord, MapData(Xa, Ya).Light_Value(1), LightColor)
       
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).Light_Value(2) = LightCalculate(radius, LightX, LightY, XCoord, YCoord, MapData(Xa, Ya).Light_Value(2), LightColor)
                   
            End If

        Next Xa
    Next Ya

    
    Exit Sub

LightRender_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsLucesRedondas.LightRender", Erl)
    Resume Next
End Sub

Private Function LightCalculate(ByVal Radio2 As Long, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As RGBA, LightColor As RGBA) As RGBA
    
    On Error GoTo LightCalculate_Err
    
    Dim XDist        As Single
    Dim YDist        As Single
    Dim VertexDist2  As Single

    XDist = LightX - XCoord
    YDist = LightY - YCoord

    VertexDist2 = XDist * XDist + YDist * YDist

    If VertexDist2 <= Radio2 Then
        Call LerpRGBA(LightCalculate, LightColor, TileLight, VertexDist2 / Radio2)
    Else
        LightCalculate = TileLight
    End If

    
    Exit Function

LightCalculate_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsLucesRedondas.LightCalculate", Erl)
    Resume Next
    
End Function

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

Public Sub LightRemoveAll(ByVal Actualizar As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************

    Dim i As Long
    
    If NumLights = 0 Then Exit Sub
    Debug.Print "Numero de luces a destruir: " & NumLights & "-" & UBound(Light_List)
    For i = 1 To UBound(Light_List)
        Delete_Light_To_Index i, Actualizar
    Next i

End Sub
