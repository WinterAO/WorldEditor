Attribute VB_Name = "modMeteorologia"
Option Explicit

Public Enum e_estados
    Amanecer = 0
    MedioDia
    Tarde
    noche
End Enum

Public Estados(0 To 3) As RGBA
Public Estado_Actual As RGBA
Public Estado_Custom As RGBA

Public Sub Init_MeteoEngine()
'***************************************************
'Author: Standelf
'Last Modification: 15/05/10
'Initializate
'***************************************************
    With Estados(e_estados.Amanecer)
        .A = 255
        .R = 230
        .G = 200
        .B = 200
    End With
    
    With Estados(e_estados.MedioDia)
        .A = 255
        .R = 255
        .G = 255
        .B = 255
    End With
    
    With Estados(e_estados.Tarde)
        .A = 255
        .R = 200
        .G = 200
        .B = 200
    End With
  
    With Estados(e_estados.noche)
        .A = 255
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
    Dim X  As Integer, Y As Integer
    
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            
            If MapZonas(MapData(X, Y).ZonaIndex).LuzBase <> 0 Then '¿La zona tiene su propia luz?
            
                Call Long_2_RGBAList(MapData(X, Y).Light_Value(), MapZonas(MapData(X, Y).ZonaIndex).LuzBase)
                
            Else
                Call RGBAList(MapData(X, Y).Light_Value(), Estado_Actual.R, Estado_Actual.G, Estado_Actual.B, Estado_Actual.A)
                
            End If
                
        Next Y
    Next X
            
    Call LucesRedondas.LightRenderAll

End Sub

