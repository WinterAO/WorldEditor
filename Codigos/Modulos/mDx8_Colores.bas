Attribute VB_Name = "mDx8_Colores"
' ****************************************************
' Para una fácil conversión entre RGBA(4 bytes) y Long
' Nota: No uso D3DCOLORVALUE porque usa 4 singles
' ****************************************************

Option Explicit

Type RGBA
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

' Desvanecimiento en Techos
Public ColorTecho As Byte
Public temp_rgb(3) As RGBA

' Titulos en el render (Nombre de mapa, subida de lvl, etc)
Public renderText As String
Public renderTextPk As String
Public renderFont As Integer
Public colorRender As Byte
Public render_msg(3) As RGBA

Public Type tColor

    R As Byte
    G As Byte
    B As Byte

End Type

'Colores del mapa
Public COLOR_WHITE(3) As RGBA

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)

Sub ConvertLongToRGB(ByVal value As Long, R As Byte, G As Byte, B As Byte)
    R = value Mod 256
    G = Int(value / 256) Mod 256
    B = Int(value / 256 / 256) Mod 256
End Sub

Sub Long_2_RGBA(Dest As RGBA, ByVal src As Long)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo Long_2_RGBA_Err
    
    Call CopyMemory(Dest, src, 4)
    
    Exit Sub

Long_2_RGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.Long_2_RGBA", Erl)
    Resume Next
    
End Sub

Function RGBA_2_Long(color As RGBA) As Long
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBA_2_Long_Err
    
    Call CopyMemory(RGBA_2_Long, color, 4)
    
    Exit Function

RGBA_2_Long_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_2_Long", Erl)
    Resume Next
    
End Function

Function RGBA_From_Long(ByVal color As Long) As RGBA
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBA_From_Long_Err
    
    Call CopyMemory(RGBA_From_Long, color, 4)
    
    Exit Function

RGBA_From_Long_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_From_Long", Erl)
    Resume Next
    
End Function

Function RGBA_From_Comp(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal A As Byte = 255) As RGBA
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBA_From_Comp_Err
    
    RGBA_From_Comp.R = R
    RGBA_From_Comp.G = G
    RGBA_From_Comp.B = B
    RGBA_From_Comp.A = A
    
    Exit Function

RGBA_From_Comp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_From_Comp", Erl)
    Resume Next
    
End Function

Function RGBA_From_vbColor(ByVal color As Long) As RGBA
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBA_From_Long_Err

    Call Long_2_RGBA(RGBA_From_vbColor, color)

    RGBA_From_vbColor.A = RGBA_From_vbColor.R
    RGBA_From_vbColor.R = RGBA_From_vbColor.B
    RGBA_From_vbColor.B = RGBA_From_vbColor.A
    RGBA_From_vbColor.A = 255
    
    Exit Function

RGBA_From_Long_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_From_Long", Erl)
    Resume Next
    
End Function

Sub SetRGBA(color As RGBA, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal A As Byte = 255)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo SetRGBA_Err
    
    color.R = R
    color.G = G
    color.B = B
    color.A = A
    
    Exit Sub

SetRGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.SetRGBA", Erl)
    Resume Next
    
End Sub

Sub Long_2_RGBAList(Dest() As RGBA, ByVal src As Long)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo Long_2_RGBAList_Err
    
    Dim i As Long
    
    For i = 0 To 3
        Call Long_2_RGBA(Dest(i), src)
    Next
    
    Exit Sub

Long_2_RGBAList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.Long_2_RGBAList", Erl)
    Resume Next
    
End Sub

Sub RGBAList(Dest() As RGBA, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal A As Byte = 255)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBAList_Err
    
    Dim i As Long
    
    For i = 0 To 3
        Call SetRGBA(Dest(i), R, G, B, A)
    Next
    
    Exit Sub

RGBAList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBAList", Erl)
    Resume Next
    
End Sub


Sub RGBA_ToList(Dest() As RGBA, color As RGBA)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBAList_Err
    
    Dim i As Long
    
    For i = 0 To 3
        Call SetRGBA(Dest(i), color.R, color.G, color.B, color.A)
    Next
    
    Exit Sub

RGBAList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_ToList", Erl)
    Resume Next
    
End Sub

Sub Copy_RGBAList(src() As RGBA, Dest() As RGBA)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo Copy_RGBAList_Err
    
    Dim i As Long
    
    For i = 0 To 3
        Dest(i) = src(i)
    Next
    
    Exit Sub

Copy_RGBAList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.Copy_RGBAList", Erl)
    Resume Next
    
End Sub

Sub LerpRGBA(Dest As RGBA, A As RGBA, B As RGBA, ByVal Factor As Single)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo LerpRGBA_Err
    
    Dim InvFactor As Single: InvFactor = (1 - Factor)

    Dest.R = A.R * InvFactor + B.R * Factor
    Dest.G = A.G * InvFactor + B.G * Factor
    Dest.B = A.B * InvFactor + B.B * Factor
    Dest.A = A.A * InvFactor + B.A * Factor
    
    Exit Sub

LerpRGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.LerpRGBA", Erl)
    Resume Next
    
End Sub

Sub LerpRGB(Dest As RGBA, A As RGBA, B As RGBA, ByVal Factor As Single)
    '***************************************************
    'Author: Martín Trionfetti (HarThaoS)
    '***************************************************
    
    On Error GoTo LerpRGB_Err
    
    Dim InvFactor As Single: InvFactor = (1 - Factor)

    Dest.R = A.R * InvFactor + B.R * Factor
    Dest.G = A.G * InvFactor + B.G * Factor
    Dest.B = A.B * InvFactor + B.B * Factor
    
    Exit Sub

LerpRGB_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.LerpRGB", Erl)
    Resume Next
    
End Sub

Sub ModulateRGBA(Dest As RGBA, A As RGBA, B As RGBA)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo ModulateRGBA_Err
    
    Dest.R = CLng(A.R) * B.R \ 255
    Dest.G = CLng(A.G) * B.G \ 255
    Dest.B = CLng(A.B) * B.B \ 255
    Dest.A = CLng(A.A) * B.A \ 255
    
    Exit Sub

ModulateRGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.ModulateRGBA", Erl)
    Resume Next
    
End Sub

Sub AddRGBA(Dest As RGBA, A As RGBA, B As RGBA)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo AddRGBA_Err
    
    Dest.R = Min(CLng(A.R) + CLng(B.R), 255)
    Dest.G = Min(CLng(A.G) + CLng(B.G), 255)
    Dest.B = Min(CLng(A.B) + CLng(B.B), 255)
    Dest.A = Min(CLng(A.A) + CLng(B.A), 255)
    
    Exit Sub

AddRGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.AddRGBA", Erl)
    Resume Next
    
End Sub

Function vbColor_2_Long(color As Long) As Long
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo vbColor_2_Long_Err
    
    Dim TmpColor As RGBA
    Call Long_2_RGBA(TmpColor, color)

    TmpColor.A = TmpColor.R
    TmpColor.R = TmpColor.B
    TmpColor.B = TmpColor.A
    TmpColor.A = 255
    
    vbColor_2_Long = RGBA_2_Long(TmpColor)
    
    Exit Function

vbColor_2_Long_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.vbColor_2_Long", Erl)
    Resume Next
    
End Function

Sub Copy_RGBAList_WithAlpha(Dest() As RGBA, src() As RGBA, ByVal Alpha As Byte)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo Copy_RGBAList_WithAlpha_Err
    
    Dim i As Long
    
    For i = 0 To 3
        Dest(i) = src(i)
        Dest(i).A = Alpha
    Next
    
    Exit Sub

Copy_RGBAList_WithAlpha_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.Copy_RGBAList_WithAlpha", Erl)
    Resume Next
    
End Sub

Function RGBA_ToString(color As RGBA) As String
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBA_ToString_Err
    
    RGBA_ToString = "RGBA(" & color.R & ", " & color.G & ", " & color.B & ", " & color.A & ")"
    
    Exit Function

RGBA_ToString_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_ToString", Erl)
    Resume Next
    
End Function

