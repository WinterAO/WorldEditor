Attribute VB_Name = "modTileEngine"
Option Explicit

Public EngineRun As Boolean

Public OffsetCounterX As Single
Public OffsetCounterY As Single

Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer

'Tamano de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime As Single
Public timerTicksPerFrame As Single

Public FPS As Long
Public FramesPerSecCounter As Long
Public FPSLastCheck As Long

'*************
'<<<<< PRIVADAS >>>>>
'****

'Tamano del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private MouseTileX As Integer
Private MouseTileY As Integer

Private DrawBuffer As cDIBSection

'*************
'<<<<< CONSTANTES >>>>>>
'****

'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

Public Const DEGREETORADIAN As Single = 0.01745329251994 'Pi / 180

'Grafico que se muestra si hay error en un Grh
Public Const GRH_ERROR As Long = 22512

Private Const MOVEMENT_SPEED As Single = 1

'*************
'<<<<< API >>>>>>
'****

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Sub ConvertCPtoTP(ByVal viewPortX As Integer, _
                  ByVal viewPortY As Integer, _
                  ByRef tX As Integer, _
                  ByRef tY As Integer)
                  
    On Error GoTo ConvertCPtoTP_Error
    '******************************************
    'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
    '******************************************

    tX = (UserPos.x + viewPortX \ TilePixelWidth - WindowTileWidth \ 2)
    tY = (UserPos.y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2)
    
    On Error GoTo 0
    Exit Sub

ConvertCPtoTP_Error:
    Call RegistrarError(Err.Number, Err.Description, "ConvertCPtoTP", Erl)
    Resume Next
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

    '¿Es un Grh invalido?
    If GrhIndex <= 0 Or GrhIndex > grhCount Then GrhIndex = GRH_ERROR
    
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.speed = GrhData(Grh.GrhIndex).speed
End Sub

Sub MoveCharbyPos(CharIndex As Integer, nX As Integer, nY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
    Dim x As Integer
    Dim y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As Byte
    
    With CharList(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
        addx = nX - x
        addy = nY - y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        End If
        
        If Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        End If
        
        If Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        End If
        
        If Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.x = nX
        .Pos.y = nY
        MapData(x, y).CharIndex = 0
        
        .MoveOffset.x = -1 * (TilePixelWidth * addx)
        .MoveOffset.y = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
    End With

End Sub

Public Sub InitTileEngine(ByVal setTilePixelHeight As Integer, _
                          ByVal setTilePixelWidth As Integer, _
                          ByVal pixelsToScrollPerFrameX As Integer, _
                          pixelsToScrollPerFrameY As Integer)
    '***************************************************
    'Author: Aaron Perkins
    'Last Modification: 08/14/07
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Configures the engine to start running.
    '***************************************************

    On Error GoTo ErrorHandler:

    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    
    Call ChangeView

    'Tamaño del mapa por defecto 1000 x 1000 ("mundo continuo")
    Call setMapSize(1000, 1000)
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    Call LoadGraphics

    Exit Sub
    
ErrorHandler:

    Call RegistrarError(Err.Number, Err.Description, "Mod_TileEngine.InitTileEngine", Erl)
    Call CloseMapEditor
    
End Sub

Public Sub setMapSize(ByVal XMax As Integer, ByVal YMax As Integer)
    '***********************************
    'Seteamos el tamaño del mapa
    '***********************************
    
    If XMaxMapSize <> XMax Or YMaxMapSize <> YMax Then
    
        XMaxMapSize = XMax
        YMaxMapSize = YMax
        
        'Set intial user position
        UserPos.x = (XMax / 2) + 25
        UserPos.y = (YMax / 2) - 25
        
        MinXBorder = XMinMapSize + (ClienteWidth \ 2)
        MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
        MinYBorder = YMinMapSize + (ClienteHeight \ 2)
        MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)
        
        ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    End If
    
End Sub

Public Sub ChangeView()
    WindowTileHeight = Round(frmMain.MainViewPic.ScaleHeight / 32, 0)
    WindowTileWidth = Round(frmMain.MainViewPic.ScaleWidth / 32, 0)
    
    frmMapa.ApuntadorRadar.Height = WindowTileHeight
    frmMapa.ApuntadorRadar.Width = WindowTileWidth

    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2
    
End Sub

Public Sub LoadGraphics()
    Call SurfaceDB.Initialize(DirectD3D8, ClientSetup.byMemory)
End Sub

Sub ShowNextFrame()

    On Error GoTo ErrorHandler:

    If EngineRun Then
    
        Dim cX        As Integer

        Dim cY        As Integer

        Dim Cuadrante As Integer
        
        Call Engine_BeginScene
        
        Call DesvanecimientoMsg
        
        If UserMoving Then
        
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.x <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame
    
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
                    OffsetCounterX = 0
                    AddtoUserPos.x = 0
                    UserMoving = False
                    
                End If
                    
            End If
                
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame
    
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.y = 0
                    UserMoving = False
                    
                End If
    
            End If

        End If
        
        '****** Update screen ******
        Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
        
        ' Calculamos los FPS y los mostramos
        Call Engine_Update_FPS
        Call DrawText(10, 5, "FPS: " & modTileEngine.FPS, COLOR_WHITE, False)
        Call DrawText(10, 20, "Mouse: " & MousePos, COLOR_WHITE, False)
        
        Call ObtenerCuadrante(Cuadrante, cX, cY)
        Call DrawText(10, 35, "Cuadrante: " & Cuadrante & " X:" & cX & " Y: " & cY, COLOR_WHITE, False)
        
        If ContadorTiles > 0 Then Call DrawText(10, 50, "Tiles: " & ContadorTiles, COLOR_WHITE, False)
        
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * Engine_BaseSpeed
            
        Call Engine_EndScene(MainScreenRect, 0)

    End If

ErrorHandler:

    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        
        Call mDx8_Engine.Engine_DirectX8_Init
        
        Call LoadGraphics
    
    End If

End Sub

Public Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim Start_Time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        Call QueryPerformanceFrequency(timer_freq)
    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Sub Device_Textured_Render(ByVal x As Single, ByVal y As Single, _
                                  ByVal Width As Integer, ByVal Height As Integer, _
                                  ByVal sX As Integer, ByVal sY As Integer, _
                                  ByVal tex As Long, _
                                  ByRef color() As RGBA, _
                                  Optional ByVal Alpha As Boolean = False, _
                                  Optional ByVal angle As Single = 0, _
                                  Optional ByVal ScaleX As Single = 1!, _
                                  Optional ByVal ScaleY As Single = 1!)

        Dim Texture As Direct3DTexture8
        
        Dim TextureWidth As Long, TextureHeight As Long
        Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
        
        With SpriteBatch

                Call .SetTexture(Texture)
                    
                Call .SetAlpha(Alpha)
                
                If TextureWidth <> 0 And TextureHeight <> 0 Then
                    Call .Draw(x, y, Width * ScaleX, Height * ScaleY, color, sX / TextureWidth, sY / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, angle)
                Else
                    Call .Draw(x, y, TextureWidth * ScaleX, TextureHeight * ScaleY, color, , , , , angle)
                End If
                
        End With
        
End Sub

Function LegalPos(x As Integer, y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************

    LegalPos = True
    
    'Check to see if its out of bounds
    If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        LegalPos = False
        Exit Function
    End If
    
    'Check to see if its blocked
    If MapData(x, y).bLocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    'Check for character
    If MapData(x, y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If

End Function

Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Sub Draw_Grh(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByRef Color_List() As RGBA, ByVal Animate As Byte, Optional ByVal Alpha As Boolean = False, Optional ByVal angle As Single = 0, Optional ByVal ScaleX As Single = 1!, Optional ByVal ScaleY As Single = 1!)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Long
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
On Error GoTo Error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.speed) * MOVEMENT_SPEED

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - (.pixelWidth * ScaleX - TilePixelWidth) \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        Call Device_Textured_Render(x, y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha, angle, ScaleX, ScaleY)
        
    End With
    
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'Call Log_Engine("Error in Draw_Grh, " & Err.Description & ", (" & Err.number & ")")
        MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log.", vbCritical
        Call CloseMapEditor
    End If
End Sub

Public Sub DrawHead(ByVal Head As Integer, ByVal x As Integer, ByVal y As Integer, Light() As RGBA, ByVal Heading As Byte, Optional ByVal EsCabeza As Boolean = True, Optional ByVal Alpha As Boolean = False, Optional ByVal angle As Single = 0, Optional ByVal ScaleX As Single = 1!, Optional ByVal ScaleY As Single = 1!)

    Dim textureX1 As Integer
    Dim textureX2 As Integer
    Dim textureY1 As Integer
    Dim textureY2 As Integer
    Dim OffsetX As Integer
    Dim OffsetY As Integer
    Dim Texture As Long

    If EsCabeza Then
        If heads(Head).Texture <= 0 Then Exit Sub
        Texture = heads(Head).Texture
    Else
        If Cascos(Head).Texture <= 0 Then Exit Sub
        Texture = Cascos(Head).Texture
    End If
    
    textureX2 = 27
    textureY2 = 32
 
    If EsCabeza Then
        textureX1 = heads(Head).startX - textureX2
        textureY1 = ((Heading - 2) * textureY2) + heads(Head).startY
    Else
        textureX1 = Cascos(Head).startX - textureX2 + 1
        textureY1 = ((Heading - 2) * textureY2) + Cascos(Head).startY + 2
    End If
    
    Device_Textured_Render x - OffsetX + 3, y - OffsetY + 4, textureX2, textureY2, (textureX2 + textureX1), (textureY2 + textureY1), Texture, Light, Alpha, angle, ScaleX, ScaleY

End Sub

Sub Draw_GrhIndex(ByVal GrhIndex As Long, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByRef Color_List() As RGBA, Optional ByVal angle As Single = 0, Optional ByVal Alpha As Boolean = False)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - (.pixelWidth - TilePixelWidth) \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        'Draw
        Call Device_Textured_Render(x, y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha)
    End With
    
End Sub

Public Sub PrepareDrawBuffer()
    Set DrawBuffer = New cDIBSection
    'El tamanio del buffer es arbitrario = 1024 x 1024
    Call DrawBuffer.Create(1024, 1024)
    
End Sub

Public Sub CleanDrawBuffer()
    Set DrawBuffer = Nothing
    
End Sub

Sub RenderScreen(ByVal tilex As Integer, _
                 ByVal tiley As Integer, _
                 ByVal PixelOffsetX As Integer, _
                 ByVal PixelOffsetY As Integer)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    '**************************************************************
    
    On Error GoTo RenderScreen_Err
    
    Dim y                As Long     'Keeps track of where on map we are

    Dim x                As Long     'Keeps track of where on map we are
    
    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen
    
    Dim MinY             As Long     'Start Y pos on current map

    Dim MaxY             As Long     'End Y pos on current map

    Dim MinX             As Long     'Start X pos on current map

    Dim MaxX             As Long     'End X pos on current map
    
    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen
    
    Dim minXOffset       As Integer

    Dim minYOffset       As Integer
    
    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    Dim ElapsedTime      As Single

    Dim ColorFinal(3)    As Long
    
    Dim Grh              As Grh                  'Temp Grh for show tile and blocked
    
    Dim bCapa            As Byte
    
    Dim Sobre            As Long
    
    ElapsedTime = Engine_ElapsedTime()
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    MinY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize * 2 ' WyroX: Parche para que no desaparezcan techos y arboles
    MinX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If MinY < XMinMapSize Then
        minYOffset = YMinMapSize - MinY
        MinY = YMinMapSize

    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If MinX < XMinMapSize Then
        minXOffset = XMinMapSize - MinX
        MinX = XMinMapSize

    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    If Val(frmSuperficies.cCapas.Text) >= 1 And (frmSuperficies.cCapas.Text) <= 4 Then
        bCapa = Val(frmSuperficies.cCapas.Text)
        
    Else
        bCapa = 1

    End If
    
    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX

            '###################
            'CAPAS
            '###################

            If InMapBounds(x, y) Then
    
                PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
                PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY
                
                'Layer 1 **********************************
                If MapData(x, y).Graphic(1).GrhIndex <> 0 And VerCapa1 Then Call Draw_Grh(MapData(x, y).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Light_Value(), 1)
    
                'Layer 2 **********************************
                If MapData(x, y).Graphic(2).GrhIndex <> 0 And VerCapa2 Then Call Draw_Grh(MapData(x, y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Light_Value(), 1)
            
            End If
        
            ScreenX = ScreenX + 1
        Next
    
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next
    
    '<----- Layer Obj, Char, 3 ----->
    ScreenY = minYOffset - TileBufferSize

    For y = MinY To MaxY
        
        ScreenX = minXOffset - TileBufferSize

        For x = MinX To MaxX

            If InMapBounds(x, y) Then
            
                PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
                PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
                
                With MapData(x, y)
                
                    'Object Layer ***********************************
                    If .ObjGrh.GrhIndex <> 0 And VerObjetos Then Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, .Light_Value(), 1)

                    'Char layer**************************************
                    If .CharIndex <> 0 And VerNpcs Then Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)

                    'Layer 3 *****************************************
                    If .Graphic(3).GrhIndex <> 0 And VerCapa3 Then Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Light_Value(), 1)

                    'Particulas **************************************
                    If .Particle_Group_Index And VerParticulas Then Call mDx8_Particulas.Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetYTemp + 16)
                    
                End With
                
            End If
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y
    
    '<----- Layer 4 ----->
    ScreenY = minYOffset - TileBufferSize

    For y = MinY To MaxY

        ScreenX = minXOffset - TileBufferSize

        For x = MinX To MaxX
            
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
            
            'Layer 4
            If VerCapa4 Then If MapData(x, y).Graphic(4).GrhIndex Then Call Draw_Grh(MapData(x, y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Light_Value(), 1)
            
            If MapData(x, y).TileExit.Map <> 0 And VerTranslados Then
                Grh.GrhIndex = 3
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, COLOR_WHITE(), 1)
                        
            End If
                
            'Show blocked tiles
            If VerBlockeados And MapData(x, y).bLocked = 1 Then
                Grh.GrhIndex = 4
                Grh.FrameCounter = 1
                Grh.Started = 0
                    
                Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, COLOR_WHITE(), 1)
                        
            End If
                
            If VerGrilla Then
                Grh.GrhIndex = 2
                Grh.FrameCounter = 1
                Grh.Started = 0
                    
                Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, COLOR_WHITE(), 0)
                        
            End If

            If VerTriggers Then If MapData(x, y).Trigger > 0 Then Call DrawText(PixelOffsetXTemp + 5, PixelOffsetYTemp - 13, MapData(x, y).Trigger, COLOR_WHITE, False, 2)
                
            If frmMain.mnuverZonas(0).Checked Then 'Zona actual
                If MapData(x, y).ZonaIndex = frmZonas.LstZona.ListIndex + 1 And MapData(x, y).ZonaIndex > 0 Then Call DrawText(PixelOffsetXTemp + 7, PixelOffsetYTemp + 7, "z" & MapData(x, y).ZonaIndex, COLOR_WHITE, False, 1)
                            
            ElseIf frmMain.mnuverZonas(1).Checked Then 'Todas las zonas
                    
                If MapData(x, y).ZonaIndex > 0 Then
                    Call RGBA_ToList(temp_rgb(), colorZona(MapData(x, y).ZonaIndex))
                    Call DrawText(PixelOffsetXTemp + 7, PixelOffsetYTemp + 7, "z" & MapData(x, y).ZonaIndex, temp_rgb(), False, 1)
                End If

            End If
                    
            If Seleccionando Then
                If x >= SeleccionIX And y >= SeleccionIY Then
                    If x <= SeleccionFX And y <= SeleccionFY Then
                        Grh.GrhIndex = 2
                        Grh.FrameCounter = 1
                        Grh.Started = 0
                                
                        Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Light_Value(), 1)

                    End If

                End If

            End If
             
            ScreenX = ScreenX + 1
            
        Next x

        ScreenY = ScreenY + 1
    Next y
    
    'Preview al pegar
    If EstadoSelect = eEstadoSelect.Copiado Then
        Dim i As Byte

        For i = 1 To 4
            For y = 0 To SeleccionAlto - 1
                For x = 0 To SeleccionAncho - 1

                    If y + SobreY > YMaxMapSize Then Exit For
                    If x + SobreX > XMaxMapSize Then Exit For
                    Call Draw_Grh(SeleccionMap(x, y).Graphic(1), (x + SobreX - UserPos.x + HalfWindowTileWidth) * 32 + PixelOffsetX, (y + SobreY - UserPos.y + HalfWindowTileHeight) * 32 + PixelOffsetY, 1, SeleccionMap(x, y).Light_Value(), 1)
                Next x
            Next y
        Next i
                
    End If
    
    'Preview al seleccionar una superficie
    If frmSuperficies.cSeleccionarSuperficie.value = True Then
        Dim mGrh As Grh
        Dim aux  As Long
        Dim dy   As Integer
        Dim dX   As Integer

        If frmConfigSup.DespMosaic.value = vbChecked Then
            dy = Val(frmConfigSup.DMLargo.Text)
            dX = Val(frmConfigSup.DMAncho.Text)
        Else
            dy = 0
            dX = 0

        End If
        
        If frmConfigSup.mAncho.Text = "0" Then frmConfigSup.mAncho.Text = "1"
        If frmConfigSup.mLargo.Text = "0" Then frmConfigSup.mLargo.Text = "1"

        If frmSuperficies.chkAutoCompletarSuperficies.value = vbUnchecked Then
            aux = Val(frmSuperficies.cGrh.Text) + (((SobreY + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((SobreX + dX) Mod frmConfigSup.mAncho.Text)
            mGrh.GrhIndex = aux
            InitGrh mGrh, aux
            Call Draw_Grh(mGrh, (SobreX - UserPos.x + HalfWindowTileWidth) * 32 + PixelOffsetX, (SobreY - UserPos.y + HalfWindowTileHeight) * 32 + PixelOffsetY, 1, COLOR_WHITE(), 1, False)
            
        Else

            For x = 1 To frmConfigSup.mAncho.Text
                For y = 1 To frmConfigSup.mLargo.Text
                    aux = Val(frmSuperficies.cGrh.Text) + (y - 1) * frmConfigSup.mAncho.Text + x - 1
                    mGrh.GrhIndex = aux
                    InitGrh mGrh, aux
                    Call Draw_Grh(mGrh, (SobreX - UserPos.x + HalfWindowTileWidth + x - 1) * 32 + PixelOffsetX, (SobreY - UserPos.y + HalfWindowTileHeight + y - 1) * 32 + PixelOffsetY, 1, COLOR_WHITE(), 1, False)
            
                Next y
            Next x

        End If
            
    End If
    
    If colorRender <> 240 Then
        Call Draw_GrhIndex(34027, frmMain.MainViewPic.ScaleHeight - 64, 150, 1, render_msg())
        Call DrawText(frmMain.MainViewPic.ScaleHeight - 64, 105, renderText, render_msg(), True, 2)
    End If
    
    Exit Sub
    
RenderScreen_Err:

    If Err.Number Then
        Call RegistrarError(Err.Number, Err.Description, "Mod_TileEngine.RenderScreen", Erl)

    End If
    
End Sub

Private Sub CharRender(ByVal CharIndex As Long, _
                       ByVal PixelOffsetX As Integer, _
                       ByVal PixelOffsetY As Integer)
    '*******************************************************
    'Esto forma parte del RenderScreen.
    'Dibuja todo aquello que tenga cuerpo (por asi decirlo)
    'Bichos y PJ
    '*******************************************************
    Dim moved As Boolean
    Dim Pos   As Integer
    Dim line  As String
    
    With CharList(CharIndex)

        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If
        
        If .Heading = 0 Then Exit Sub
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY

        'Dibujamos el cuerpo
        If .Body.Walk(.Heading).GrhIndex Then _
                    Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, COLOR_WHITE(), 1)
            
        If .Head Then Call DrawHead(.Head, PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y + OFFSET_HEAD, COLOR_WHITE(), .Heading, True)
            
        'Draw Helmet
        If .Casco Then Call DrawHead(.Casco, PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y + OFFSET_HEAD, COLOR_WHITE(), .Heading, False)
                
        'Draw Weapon
        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, COLOR_WHITE(), 1)
                
        'Draw Shield
        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, COLOR_WHITE(), 1)
    End With
End Sub

Public Sub fPreviewGrh(ByVal GrhIn As Long)
    '*************************************************
    'Author: Unkwown
    'Last modified: 22/05/06
    '*************************************************
    
    On Error GoTo fPreviewGrh_Err

    If Val(GrhIn) < 1 Then
        frmSuperficies.cGrh.Text = grhCount
        Exit Sub

    End If

    If Val(GrhIn) > grhCount Then
        frmSuperficies.cGrh.Text = 1
        Exit Sub

    End If

    'Change CurrentGrh
    CurrentGrh.GrhIndex = GrhIn
    CurrentGrh.Started = 1
    CurrentGrh.FrameCounter = 1

    Exit Sub

fPreviewGrh_Err:
    'Call RegistrarError(Err.Number, Err.Description, "modPaneles.fPreviewGrh", Erl)

    Resume Next
    
End Sub

Public Sub RenderPreview(Optional ByVal SinMosaico As Boolean = False)
    '*************************************************
    'Author: Lorwik
    'Last modified: 29/04/2023
    '*************************************************
    
    On Error Resume Next

    Dim DestRect As RECT
    
    Dim i        As Integer

    Dim j        As Integer

    Dim ww       As Integer

    Dim hh       As Integer

    Dim Cont     As Integer
    
    With DestRect
        .Bottom = frmPreview.PreviewGrh.Height
        .Right = frmPreview.PreviewGrh.Width

    End With
    
    'Si el Render no esta activo, salimos
    If Not frmPreview.PreviewGrh.Visible Then Exit Sub
    
    'Clear the inventory window
    Call Engine_BeginScene
    
    If frmConfigSup.MOSAICO = vbUnchecked Or SinMosaico Then
        Call Draw_GrhIndex(CurrentGrh.GrhIndex, frmPreview.PreviewGrh.Height / 2, frmPreview.PreviewGrh.Width - 100, 1, COLOR_WHITE(), 0)
        
    Else
    
        hh = Val(frmConfigSup.mLargo)
        ww = Val(frmConfigSup.mAncho)
        
        For i = 1 To hh
            For j = 1 To ww
            
                Call Draw_GrhIndex(CurrentGrh.GrhIndex, j * 32, i * 32, 0, COLOR_WHITE())

                If Cont < hh * ww Then Cont = Cont + 1
                CurrentGrh.GrhIndex = CurrentGrh.GrhIndex + 1
            Next
        Next
        
        CurrentGrh.GrhIndex = CurrentGrh.GrhIndex - Cont

    End If
    
    frmPreview.PreviewGrh.AutoRedraw = False

    Call Engine_EndScene(DestRect, frmPreview.PreviewGrh.hWnd)

    Call DrawBuffer.LoadPictureBlt(frmPreview.PreviewGrh.hDC)

    frmPreview.PreviewGrh.AutoRedraw = True

    Call DrawBuffer.PaintPicture(frmPreview.PreviewGrh.hDC, 0, 0, frmPreview.PreviewGrh.Width, frmPreview.PreviewGrh.Height, 0, 0, vbSrcCopy)

End Sub

Public Sub RenderParticlePreview()
'***********************************************
'Autor: Lorwik
'Fecha: ????
'Descripcion: Renderiza la preview de la particula seleccionada
'***********************************************

    Dim DestRect     As RECT
    
    Dim i As Integer, j As Integer
    Dim Cont As Integer
    
    With DestRect
        .Bottom = frmParticulas.ParticlePic.ScaleHeight
        .Right = frmParticulas.ParticlePic.ScaleWidth

    End With
    
    'Clear the inventory window
    Call Engine_BeginScene
    
    If ParticlePreview <> 0 Then _
        Call mDx8_Particulas.Particle_Group_Render(ParticlePreview, frmParticulas.ParticlePic.ScaleWidth / 2, frmParticulas.ParticlePic.ScaleHeight / 2)
    
    frmParticulas.ParticlePic.AutoRedraw = False

    Call Engine_EndScene(DestRect, frmParticulas.ParticlePic.hWnd)

    Call DrawBuffer.LoadPictureBlt(frmParticulas.ParticlePic.hDC)

    frmParticulas.ParticlePic.AutoRedraw = True

    Call DrawBuffer.PaintPicture(frmParticulas.ParticlePic.hDC, 0, 0, frmParticulas.ParticlePic.Width, frmParticulas.ParticlePic.Height, 0, 0, vbSrcCopy)
End Sub

Function NextOpenChar() As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
    Dim LoopC As Integer
    
    LoopC = 1
    Do While CharList(LoopC).active
        LoopC = LoopC + 1
    Loop
    
    NextOpenChar = LoopC

End Function

Public Sub MapCapture(ByRef Format As Boolean, _
                      ByVal ConBordes As Boolean, _
                      mx As Integer, _
                      my As Integer)
    
    On Error GoTo MapCapture_Err

    '*************************************************
    'Author: Torres Patricio(Pato)
    'Last modified:12/03/11
    '*************************************************
    
    Dim D3DWindow        As D3DPRESENT_PARAMETERS
    Dim y                As Long     'Keeps track of where on map we are
    Dim x                As Long     'Keeps track of where on map we are
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    Dim Grh              As Grh      'Temp Grh for show tile and blocked
    Static srcRect       As RECT

    Dim MaxX             As Integer
    Dim MaxY             As Integer
    Dim MinX             As Integer
    Dim MinY             As Integer
    
    MinX = mx + 1 - 8
    MinY = my + 1 - 8
    MaxX = MinX + 100 + 8
    MaxY = MinY + 100 + 8
    
    If MinX < 1 Then MinX = 1
    If MinY < 1 Then MinY = 1
    
    If MaxX > XMinMapSize Then MaxX = XMaxMapSize
    If MaxY > YMinMapSize Then MaxY = YMaxMapSize
    
    DoEvents
    frmRenderer.picMap.Width = 800
    frmRenderer.picMap.Height = 800
    
    With srcRect
        .Left = 0
        .Top = 0
        .Right = frmRenderer.picMap.ScaleWidth
        .Bottom = frmRenderer.picMap.ScaleHeight
    End With
    
    Call Engine_BeginScene
    
    DoEvents

    For y = MinY To MaxY
        For x = MinX To MaxX
        
            'Layer 1 **********************************
            If MapData(x, y).Graphic(1).GrhIndex <> 0 Then
                PixelOffsetXTemp = (x - 1) * 32
                PixelOffsetYTemp = (y - 1) * 32
                
                Call Draw_Grh(MapData(x, y).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 0, MapData(x, y).Light_Value(), 1)

            End If

            '******************************************
        
        Next x
    Next y
    
    'Draw floor layer 2
    For y = MinY To MaxY
        For x = MinX To MaxX
            PixelOffsetXTemp = (x - 1) * 32
            PixelOffsetYTemp = (y - 1) * 32
            
            'Layer 2 **********************************
            If (MapData(x, y).Graphic(2).GrhIndex <> 0) And VerCapa2 Then
                Call Draw_Grh(MapData(x, y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Light_Value(), 1)

            End If

            '******************************************
        
        Next x
    Next y

    'Draw Transparent Layers
    For y = MinY To MaxY
        For x = MinX To MaxX
                
            PixelOffsetXTemp = (x - MinX) * 32
            PixelOffsetYTemp = (y - MinY) * 32
        
            With MapData(x, y)
            
                'Object Layer **********************************
                If (.ObjGrh.GrhIndex <> 0) And VerObjetos Then
                    Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Light_Value(), 1)

                End If

                '***********************************************
            
                'Layer 3 *****************************************
                If (.Graphic(3).GrhIndex <> 0) And VerCapa3 Then
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Light_Value(), 1)

                End If

                '************************************************
            
            End With

        Next x
    Next y
    
    'Draw layer 4
    For y = MinY To MaxY
        For x = MinX To MaxX

            With MapData(x, y)
            
                'Layer 4 **********************************
                If (.Graphic(4).GrhIndex <> 0) And VerCapa4 Then
                    
                    'Draw
                    Call Draw_Grh(.Graphic(4), (x - MinX) * 32, (y - MinY) * 32, 1, MapData(x, y).Light_Value(), 1)
                    
                End If

                '**********************************
            
            End With

        Next x
    Next y

    'Draw trans, bloqs, triggers and select tiles
    For y = MinY To MaxY
        For x = MinX To MaxX

            With MapData(x, y)
                PixelOffsetXTemp = (x - MinX) * 32
                PixelOffsetYTemp = (y - MinY) * 32
            
                '**********************************
                Grh.FrameCounter = 1
                Grh.Started = 0

                If (.TileExit.Map <> 0) And VerTranslados Then
                    Grh.GrhIndex = 3
                
                    Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Light_Value(), 0)

                End If
            
                'Show blocked tiles
                If (.bLocked = 1) And VerBlockeados Then
                    Grh.GrhIndex = 4
                    Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Light_Value(), 0)

                End If

                '******************************************
            
            End With

        Next x
    Next y

    ToWorldMap2 = True
    DoEvents
    
    'Get timing info
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * Engine_BaseSpeed
    
    Call Engine_EndScene(srcRect, frmRenderer.picMap.hWnd)

    DoEvents
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''Guardo la imagen''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Call frmRenderer.Capturar_Imagen(frmRenderer.picMap, frmRenderer.picMap)
     
    'Si no existe la carpeta de MiniMapas, la hacemos.
    If Not FileExist(App.Path & "\Render\", vbDirectory) Then
        Call MkDir(App.Path & "\Render\")
    End If
    
    If Not FileExist(App.Path & "\Resources\", vbDirectory) Then
        Call MkDir(App.Path & "\Resources\")
    End If
    
    If Not FileExist(App.Path & "\Resources\Minimapas\", vbDirectory) Then
        Call MkDir(App.Path & "\Resources\Minimapas\")
    End If
    
    'SaveImageAs App.Path & "\Render\" & MapName & ".png", FrmRender.picMap.hdc, FrmRender.picMap.ScaleWidth, FrmRender.picMap.ScaleHeight, CLng(100)
    
    SavePicture frmRenderer.picMap.Image, App.Path & "\Render\Partials\Mapa" & UserMap & "-" & MinX & "-" & MinY & "-" & XMaxMapSize & "-" & YMaxMapSize & ".bmp"
    'FrmRender.PicGrande.PaintPicture FrmRender.picMap.Image, ((minX - 1) \ 200) * 800, ((minY - 1) \ 200) * 800, 800, 800
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    DoEvents
    Exit Sub

MapCapture_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.MapCapture", Erl)
    Resume Next
    
End Sub

Public Sub DesvanecimientoMsg()
'*****************************************************************
'Author: FrankoH
'Last Modify Date: 04/09/2019
'DESVANECIMIENTO DE LOS TEXTOS DEL RENDER
'*****************************************************************
    Static lastmovement As Long
    
    If GetTickCount - lastmovement > 1 Then
        lastmovement = GetTickCount
    Else
        Exit Sub
    End If

    If LenB(renderText) Then
        If Not Val(colorRender) = 0 Then colorRender = colorRender - 1
    ElseIf LenB(renderText) = 0 Then
        Exit Sub
    Else
        If Not Val(colorRender) = 240 Then colorRender = colorRender + 1
    End If
    
    If Not Val(colorRender) = 240 Then
        Call RGBAList(render_msg(), 255, 255, 255, colorRender)
    End If
    
    If colorRender = 0 Then renderMsgReset
    
End Sub

Public Sub renderMsgReset()

    renderFont = 1
    renderText = vbNullString
    renderTextPk = vbNullString

End Sub

Public Sub ShowMessageScreen(ByVal Message As String)

    Call renderMsgReset

    renderText = Message
    colorRender = 240

End Sub
