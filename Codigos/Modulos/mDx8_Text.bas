Attribute VB_Name = "mDx8_Text"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)
    
Private Type CharVA
    X As Integer
    Y As Integer
    W As Integer
    h As Integer
    
    Tx1 As Single
    Tx2 As Single
    Ty1 As Single
    Ty2 As Single
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type

Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
End Type

Private cfonts(1 To 2) As CustomFont ' _Default2 As CustomFont

Public Sub Text_Render_Special(ByVal intX As Integer, ByVal intY As Integer, ByRef strText As String, ByVal lngColor As Long, Optional bolCentred As Boolean = False, Optional Font As Integer = 1)  ' GSZAO
'*****************************************************************
'Text_Render_Special by ^[GS]^
'*****************************************************************
    Dim i As Long
    If LenB(strText) <> 0 Then

        Call Engine_Long_To_RGB_List(temp_rgb(), lngColor)

        Call Engine_Render_Text(SpriteBatch, cfonts(Font), strText, intX, intY, temp_rgb(), bolCentred, , , Font)
        
    End If
    
End Sub ' GSZAO

Private Function Es_Emoticon(ByVal ascii As Byte) As Boolean ' GSZAO
'*****************************************************************
'Emoticones by ^[GS]^
'*****************************************************************
    Es_Emoticon = False
    If (ascii = 129 Or ascii = 137 Or ascii = 141 Or ascii = 143 Or ascii = 144 Or ascii = 157 Or ascii = 160) Then
        Es_Emoticon = True
    End If
End Function ' GSZAO

Private Sub Engine_Render_Text(ByRef Batch As clsBatch, _
                                ByRef UseFont As CustomFont, _
                                ByVal Text As String, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByRef color() As Long, _
                                Optional ByVal Center As Boolean = False, _
                                Optional ByVal Alpha As Byte = 255, _
                                Optional ByVal ParseEmoticons As Boolean = False, _
                                Optional Font As Integer = 1)
                                
'*****************************************************************
'Render text with a custom font
'*****************************************************************
    Dim TempVA As CharVA
    Dim tempstr() As String
    Dim Count As Integer
    Dim ascii() As Byte
    Dim i As Long
    Dim j As Long
    Dim yOffset As Single
    
    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    
     'WyroX: Agregado para evitar dibujar emojis en los nombres de los personajes
    If ParseEmoticons Then
        'Analizar mensaje, palabra por palabra... GSZAO
        
        tempstr = Split(Text, Chr$(32))
        Text = vbNullString

        For i = 0 To UBound(tempstr)
            If tempstr(i) = ":)" Or tempstr(i) = "=)" Then
                tempstr(i) = Chr$(129)
            ElseIf tempstr(i) = ":@" Or tempstr(i) = "=@" Then
                tempstr(i) = Chr$(137)
            ElseIf tempstr(i) = ":(" Or tempstr(i) = "=(" Then
                tempstr(i) = Chr$(141)
            ElseIf tempstr(i) = "^^" Or tempstr(i) = "^_^" Then
                tempstr(i) = Chr$(143)
            ElseIf tempstr(i) = ":D" Or tempstr(i) = "=D" Then
                tempstr(i) = Chr$(144)
            ElseIf tempstr(i) = "xD" Or tempstr(i) = "XD" Then
                tempstr(i) = Chr$(157)
            ElseIf tempstr(i) = ":S" Or tempstr(i) = "=S" Then
                tempstr(i) = Chr$(160)
            End If
            Text = Text & Chr$(32) & tempstr(i)
        Next
        ' Made by ^[GS]^ for GSZAO
    End If
    
    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)

    'Set the texture
    Call Batch.SetTexture(UseFont.Texture)
    
    If Center Then
        X = X - CInt(Engine_GetTextWidth(cfonts(Font), Text) * 0.5)
    End If
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)
        If Len(tempstr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            Count = 0
        
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
        
            'Loop through the characters
            For j = 1 To Len(tempstr(i))

                Call CopyMemory(TempVA, UseFont.HeaderInfo.CharVA(ascii(j - 1)), 24) 'this number represents the size of "CharVA" struct
                
                TempVA.X = X + Count
                TempVA.Y = Y + yOffset
                
                'Set the colors
                If Es_Emoticon(ascii(j - 1)) Then ' GSZAO los colores no afectan a los emoticones!
                    
                    If (ascii(j - 1) <> 157) Then
                        Count = Count + 5   ' Los emoticones tienen tamano propio (despues hay que cargarlos "correctamente" para evitar hacer esto)
                    End If
                    
                End If
                Call Batch.SetAlpha(False)
                Call Batch.Draw(TempVA.X, TempVA.Y, TempVA.W, TempVA.h, color, TempVA.Tx1, TempVA.Ty1, TempVA.Tx2, TempVA.Ty2)

                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(ascii(j - 1))
                
            Next j
            
        End If
    Next i

End Sub

Private Function Engine_GetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth
'***************************************************
Dim i As Integer
Dim Len_text As Long

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    Len_text = Len(Text)
    
    'Loop through the text
    For i = 1 To Len_text
        
        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
        
    Next i

End Function

Public Function Engine_AnchoTexto(ByVal Font As Byte, ByVal Text As String) As Integer
'**************************************
'Autor: Lorwik
'Fecha: 22/05/2020
'Descripcion: Devuelve la posicion de ocupa el texto
'**************************************


    Engine_AnchoTexto = Engine_GetTextWidth(cfonts(Font), Text)

End Function

Sub Engine_Init_FontTextures()
    '*****************************************************************
    'Init the custom font textures
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures
    '*****************************************************************
    On Error GoTo eDebug:
    
    Dim i       As Long
    Dim TexInfo As D3DXIMAGE_INFO_A
    Dim InfoHead As INFOHEADER
    Dim Data() As Byte
        
    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '*** Default font ***
    For i = 1 To UBound(cfonts)

        InfoHead = File_Find(DirRecursos & "Fuentes" & modCompression.Formato, "font" & CStr(i) & ".png")
        
        If InfoHead.lngFileSize <> 0 Then

            Extract_File_Memory Fuentes, "font" & LCase$(CStr(i) & ".png"), Data()
                
            'Set the texture
            Set cfonts(i).Texture = DirectD3D8.CreateTextureFromFileInMemoryEx(DirectDevice, _
                                                                       Data(0), UBound(Data) + 1, _
                                                                       D3DX_DEFAULT, _
                                                                       D3DX_DEFAULT, _
                                                                       0, _
                                                                       0, _
                                                                       D3DFMT_UNKNOWN, _
                                                                       D3DPOOL_MANAGED, _
                                                                       D3DX_FILTER_POINT, _
                                                                       D3DX_FILTER_POINT, _
                                                                       &HFF000000, _
                                                                       ByVal 0, _
                                                                       ByVal 0)
        
            Erase Data
        End If
        
        'Store the size of the texture
        cfonts(i).TextureSize.X = TexInfo.Width
        cfonts(i).TextureSize.Y = TexInfo.Height
    Next
    
    Exit Sub

eDebug:

    If Err.Number = "-2005529767" Then
        MsgBox "Error en la textura de fuente utilizada " & DirRecursos & "Font.png.", vbCritical
        End
    End If
    
    End

End Sub

Sub Engine_Init_FontSettings()
    '*****************************************************************
    'Init the custom font settings
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontSettings
    '*****************************************************************
    Dim LoopChar  As Long
    Dim Row       As Single
    Dim u         As Single
    Dim v         As Single
    Dim i         As Long
    Dim j         As Integer
    Dim fileBuff  As clsByteBuffer
    Dim InfoHead  As INFOHEADER
    Dim buffer()  As Byte
    
    '*** Default font ***
    
    For i = 1 To UBound(cfonts)
    
        'Load the header information
        InfoHead = File_Find(DirRecursos & "\Fuentes" & modCompression.Formato, LCase$("Font" & i & ".dat"))
    
        If InfoHead.lngFileSize <> 0 Then
        
            Extract_File_Memory srcFileType.Fuentes, LCase$("Font" & i & ".dat"), buffer()
        
            Set fileBuff = New clsByteBuffer
            
            fileBuff.initializeReader buffer
        
            With cfonts(i).HeaderInfo
                .BitmapWidth = fileBuff.getLong()
                .BitmapHeight = fileBuff.getLong()
                .CellWidth = fileBuff.getLong()
                .CellHeight = fileBuff.getLong()
                .BaseCharOffset = fileBuff.getByte()
                
                For j = 0 To 255
                    .CharWidth(j) = fileBuff.getByte()
                Next j
            
            End With
            
            'Calculate some common values
            cfonts(i).CharHeight = cfonts(i).HeaderInfo.CellHeight - 4
            cfonts(i).RowPitch = cfonts(i).HeaderInfo.BitmapWidth \ cfonts(i).HeaderInfo.CellWidth
            cfonts(i).ColFactor = cfonts(i).HeaderInfo.CellWidth / cfonts(i).HeaderInfo.BitmapWidth
            cfonts(i).RowFactor = cfonts(i).HeaderInfo.CellHeight / cfonts(i).HeaderInfo.BitmapHeight
            
            'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
            For LoopChar = 0 To 255
                
                'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
                Row = (LoopChar - cfonts(i).HeaderInfo.BaseCharOffset) \ cfonts(i).RowPitch
                u = ((LoopChar - cfonts(i).HeaderInfo.BaseCharOffset) - (Row * cfonts(i).RowPitch)) * cfonts(i).ColFactor
                v = Row * cfonts(i).RowFactor
        
                'Set the verticies
                With cfonts(i).HeaderInfo.CharVA(LoopChar)
                    .X = 0
                    .Y = 0
                    .W = cfonts(i).HeaderInfo.CellWidth
                    .h = cfonts(i).HeaderInfo.CellHeight
                    .Tx1 = u
                    .Ty1 = v
                    .Tx2 = u + cfonts(i).ColFactor
                    .Ty2 = v + cfonts(i).RowFactor
                End With
                
            Next LoopChar
            
            Erase buffer
            
            Set fileBuff = Nothing
        End If
    Next i
End Sub

Public Sub DrawText(ByVal X As Integer, _
                    ByVal Y As Integer, _
                    ByVal Text As String, _
                    ByVal color As Long, _
                    Optional Center As Boolean = False, _
                    Optional Font As Integer = 1)

    Dim aux(3) As Long

    Call Engine_Long_To_RGB_List(aux(), color)
    Call Engine_Render_Text(SpriteBatch, cfonts(Font), Text, X, Y, aux(), Center, , , Font)

End Sub





