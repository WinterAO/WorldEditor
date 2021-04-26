Attribute VB_Name = "modEngineColors"
Option Explicit

Public Normal_RGBList(3) As Long
Public temp_rgb(3) As Long
Public AlphaTecho As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, source As Any, ByVal Length As Long)

Public Sub Engine_D3DColor_To_RGB_List(rgb_list() As Long, color As D3DCOLORVALUE)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 14/05/10
'Blisse-AO | Set a D3DColorValue to a RGB List
'***************************************************
    rgb_list(0) = D3DColorARGB(color.a, color.R, color.G, color.B)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Sub Engine_Long_To_RGB_List(rgb_list() As Long, long_color As Long)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 16/05/10
'Blisse-AO | Set a Long Color to a RGB List
'***************************************************
    rgb_list(0) = long_color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Sub ConvertLongToRGB(ByVal value As Long, R As Byte, G As Byte, B As Byte)
    R = value Mod 256
    G = Int(value / 256) Mod 256
    B = Int(value / 256 / 256) Mod 256
End Sub

Public Function SetARGB_Alpha(rgb_list() As Long, Alpha As Byte) As Long()

    '***************************************************
    'Author: Juan Manuel Couso (Cucsifae)
    'Last Modification: 29/08/18
    'Obtiene un ARGB list le modifica el alpha y devuelve una copia
    '***************************************************
    Dim TempColor        As D3DCOLORVALUE
    Dim tempARGB(0 To 3) As Long

    'convertimos el valor del rgb list a D3DCOLOR
    Call ARGBtoD3DCOLORVALUE(rgb_list(1), TempColor)

    'comprobamos ue no se salga del rango permitido
    If Alpha > 255 Then Alpha = 255
    If Alpha < 0 Then Alpha = 0
    
    'seteamos el alpha
    TempColor.a = Alpha
    
    'generamos el nuevo RGB_List
    Call Engine_D3DColor_To_RGB_List(tempARGB(), TempColor)

    SetARGB_Alpha = tempARGB()

End Function

Public Function ARGBtoD3DCOLORVALUE(ByVal ARGB As Long, ByRef color As D3DCOLORVALUE)
    Dim dest(3) As Byte
    CopyMemory dest(0), ARGB, 4
    color.a = dest(3)
    color.R = dest(2)
    color.G = dest(1)
    color.B = dest(0)
End Function

Public Function ARGB(ByVal R As Long, ByVal G As Long, ByVal B As Long, ByVal a As Long) As Long
        
    Dim c As Long
        
    If a > 127 Then
        a = a - 128
        c = a * 2 ^ 24 Or &H80000000
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or B
    Else
        c = a * 2 ^ 24
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or B
    End If
    
    ARGB = c

End Function

