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

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************

    tX = (UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2) + 1
    tY = (UserPos.y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2) + 1
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
    Dim X As Integer
    Dim y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As Byte
    
    With CharList(CharIndex)
        X = .Pos.X
        y = .Pos.y
        
        addX = nX - X
        addY = nY - y
        
        If Sgn(addX) = 1 Then
            nHeading = eDireccion.EAST
        End If
        
        If Sgn(addX) = -1 Then
            nHeading = eDireccion.WEST
        End If
        
        If Sgn(addY) = -1 Then
            nHeading = eDireccion.NORTH
        End If
        
        If Sgn(addY) = 1 Then
            nHeading = eDireccion.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.y = nY
        MapData(X, y).CharIndex = 0
        
        .MoveOffset.X = -1 * (TilePixelWidth * addX)
        .MoveOffset.y = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
    End With

End Sub

Public Sub InitTileEngine(ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer)
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

    Call setMapSize
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    Call LoadGraphics

    Exit Sub
    
ErrorHandler:

    Call RegistrarError(Err.Number, Err.Description, "Mod_TileEngine.InitTileEngine", Erl)
    Call CloseMapEditor
    
End Sub

Public Sub setMapSize()
    '***********************************
    'Seteamos el tamaño del mapa
    '***********************************
    
    If ClientSetup.MapTam = 0 Then
        XMaxMapSize = 100
        YMaxMapSize = 100
        
        'Set intial user position
        UserPos.X = 50
        UserPos.y = 50
        
    Else
        XMaxMapSize = 1100
        YMaxMapSize = 1100
        
        'Set intial user position
        UserPos.X = 550
        UserPos.y = 550
        
    End If
    
    MinXBorder = XMinMapSize + (ClienteWidth \ 2)
    MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
    MinYBorder = YMinMapSize + (ClienteHeight \ 2)
    MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    With MapSize
        .XMax = XMaxMapSize
        .XMin = XMinMapSize
        .YMax = YMaxMapSize
        .YMin = YMinMapSize
        
        ReDim MapData(.XMin To .XMax, .YMin To .YMax)
    End With
End Sub

Public Sub ChangeView()
    WindowTileHeight = Round(frmMain.MainViewPic.ScaleHeight / 32, 0)
    WindowTileWidth = Round(frmMain.MainViewPic.ScaleWidth / 32, 0)

    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2
    
End Sub

Public Sub LoadGraphics()
    Call SurfaceDB.Initialize(DirectD3D8, ClientSetup.byMemory)
End Sub

Sub ShowNextFrame()

On Error GoTo ErrorHandler:

    If EngineRun Then
        
        Call Engine_BeginScene
        
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.X <> 0 Then
            OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
    
            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                OffsetCounterX = 0
                AddtoUserPos.X = 0
    
            End If
                    
        End If
                
        '****** Move screen Up and Down if needed ******
        If AddtoUserPos.y <> 0 Then
            OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame
    
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.y) Then
                OffsetCounterY = 0
                AddtoUserPos.y = 0
                        
            End If
    
        End If
        
        '****** Update screen ******
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
        
        ' Calculamos los FPS y los mostramos
        Call Engine_Update_FPS
        Call DrawText(10, 5, "FPS: " & modTileEngine.FPS, -1, False)
        'Call DrawText(10, 20, "Mouse: " & MousePos, -1, False)
        
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

Public Sub Device_Textured_Render(ByVal X As Single, ByVal y As Single, _
                                  ByVal Width As Integer, ByVal Height As Integer, _
                                  ByVal sX As Integer, ByVal sY As Integer, _
                                  ByVal tex As Long, _
                                  ByRef color() As Long, _
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
                    Call .Draw(X, y, Width * ScaleX, Height * ScaleY, color, sX / TextureWidth, sY / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, angle)
                Else
                    Call .Draw(X, y, TextureWidth * ScaleX, TextureHeight * ScaleY, color, , , , , angle)
                End If
                
        End With
        
End Sub

Function LegalPos(X As Integer, y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************

    LegalPos = True
    
    'Check to see if its out of bounds
    If X < XMinMapSize Or X > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        LegalPos = False
        Exit Function
    End If
    
    'Check to see if its blocked
    If MapData(X, y).bLocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    'Check for character
    If MapData(X, y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If

End Function

Function InMapBounds(ByVal X As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, ByVal Animate As Byte, Optional ByVal Alpha As Boolean = False, Optional ByVal angle As Single = 0, Optional ByVal ScaleX As Single = 1!, Optional ByVal ScaleY As Single = 1!)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Long
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
On Error GoTo Error

    If Grh.GrhIndex > grhCount Or GrhData(Grh.GrhIndex).NumFrames = 0 And GrhData(Grh.GrhIndex).FileNum = 0 Then
        Call InitGrh(Grh, GRH_ERROR) ' 23829
        Call AddtoRichTextBox(frmConsola.StatTxt, "Error en Grh. Posicion: X:" & X & " Y:" & y, 255, 0, 0)
    End If

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
                X = X - (.pixelWidth * ScaleX - TilePixelWidth) \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        Call Device_Textured_Render(X, y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha, angle, ScaleX, ScaleY)
        
    End With
    
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        #If Desarrollo = 0 Then
            Call RegistrarError(Err.Number, "Error in Draw_Grh, " & Err.Description, "Draw_Grh", Erl)
            MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
            Call CloseMapEditor
        
        #Else
            Debug.Print "Error en Draw_Grh en el grh" & CurrentGrhIndex & ", " & Err.Description & ", (" & Err.Number & ")"
        #End If
    End If
End Sub

Sub Draw_GrhIndex(ByVal GrhIndex As Long, ByVal X As Integer, ByVal y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, Optional ByVal Alpha As Boolean = False)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth - TilePixelWidth) \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        'Draw
        Call Device_Textured_Render(X, y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha)
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
    Dim X                As Long     'Keeps track of where on map we are
    
    Dim screenminY       As Integer  'Start Y pos on current screen
    Dim screenmaxY       As Integer  'End Y pos on current screen
    Dim screenminX       As Integer  'Start X pos on current screen
    Dim screenmaxX       As Integer  'End X pos on current screen
    
    Dim minY             As Integer  'Start Y pos on current map
    Dim maxY             As Integer  'End Y pos on current map
    Dim minX             As Integer  'Start X pos on current map
    Dim maxX             As Integer  'End X pos on current map
    
    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen
    
    Dim minXOffset       As Integer
    Dim minYOffset       As Integer
    
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    Dim ElapsedTime      As Single
    Dim ColorFinal(3)    As Long
    
    Dim Grh              As Grh                  'Temp Grh for show tile and blocked
    
    ElapsedTime = Engine_ElapsedTime()
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize * 2 ' WyroX: Parche para que no desaparezcan techos y arboles
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
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

    'Draw floor layer
    For y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            
            PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY
            
            'Layer 1 **********************************
            If VerCapa1 Then
                If MapData(X, y).Graphic(1).GrhIndex <> 0 Then _
                    Call Draw_Grh(MapData(X, y).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, y).Engine_Light(), 1)
            End If
            '******************************************

            'Layer 2 **********************************
            If VerCapa2 Then
                If MapData(X, y).Graphic(2).GrhIndex <> 0 Then _
                    Call Draw_Grh(MapData(X, y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, y).Engine_Light(), 1)
            End If
            '******************************************
            
            ScreenX = ScreenX + 1
        Next
    
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next
    
    '<----- Layer Obj, Char, 3 ----->
    ScreenY = minYOffset - TileBufferSize

    For y = minY To maxY
        
        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX
            If InMapBounds(X, y) Then
            
                PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
                PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
                
                With MapData(X, y)
                
                    'Object Layer **********************************
                   ' If .ObjGrh.GrhIndex <> 0 Then _
                        Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                    '***********************************************

                    'Char layer********************************
                    'If .CharIndex <> 0 Then Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                    '*************************************************

                    'Layer 3 *****************************************
                    If VerCapa3 Then
                        If .Graphic(3).GrhIndex <> 0 Then _
                            Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, y).Engine_Light(), 1)
                    End If
                    '************************************************
                    
'                    'Particulas
'                    If .Particle_Group_Index Then
'
'                        'Solo las renderizamos si estan cerca del area de vision.
'                        If EstaDentroDelArea(X, Y) Then
'                            Call mDx8_Particulas.Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetYTemp + 16)
'                        End If
'
'                    End If
                    
                End With
                
            End If
            
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next y
    
    '<----- Layer 4 ----->
    ScreenY = minYOffset - TileBufferSize

    For y = minY To maxY

        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX
            
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
            
            'Layer 4
            If VerCapa4 Then
                If MapData(X, y).Graphic(4).GrhIndex Then _
                    Call Draw_Grh(MapData(X, y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, temp_rgb(), 1)
            End If
            
            If MapData(X, y).TileExit.Map <> 0 And VerTranslados Then
                Grh.GrhIndex = 3
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, Normal_RGBList(), 1)
                        
            End If
                
            'Show blocked tiles
            If VerBlockeados And MapData(X, y).bLocked = 1 Then
                Grh.GrhIndex = 4
                Grh.FrameCounter = 1
                Grh.Started = 0
                    
                Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, Normal_RGBList(), 1)
                        
            End If
                
            If VerGrilla Then
                Grh.GrhIndex = 2
                Grh.FrameCounter = 1
                Grh.Started = 0
                    
                Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, Normal_RGBList(), 0)
                        
            End If

            If VerTriggers Then '4978
                If MapData(X, y).Trigger > 0 Then _
                    Call DrawText(PixelOffsetXTemp + 5, PixelOffsetYTemp - 13, MapData(X, y).Trigger, -1, False, 2)
            End If
            
            ScreenX = ScreenX + 1
            
        Next X

        ScreenY = ScreenY + 1
    Next y
    
RenderScreen_Err:

    If Err.Number Then
        Call RegistrarError(Err.Number, Err.Description, "Mod_TileEngine.RenderScreen", Erl)
    End If
    
End Sub

Public Sub RenderPreview()
'***********************************************
'Autor: Lorwik
'Fecha: ????
'Descripcion: Renderiza la preview de la superficie, objeto, etc seleccionada
'***********************************************

    Dim destRect     As RECT
    
    Dim i As Integer, j As Integer
    Dim Cont As Integer
    
    With destRect
        .Bottom = frmPreview.PreviewGrh.Height
        .Right = frmPreview.PreviewGrh.Width
    End With
    
    'Clear the inventory window
    Call Engine_BeginScene

    If frmConfigSup.MOSAICO.value = vbUnchecked Or frmSuperficies.Visible = True Then
        Call Draw_GrhIndex(CurrentGrh.GrhIndex, frmPreview.PreviewGrh.Height / 2, frmPreview.PreviewGrh.Width - 50, 1, Normal_RGBList(), 0)
 
    Else
        For i = 1 To CInt(Val(frmConfigSup.mLargo))
            For j = 1 To CInt(Val(frmConfigSup.mAncho))
            
                Call Draw_GrhIndex(CurrentGrh.GrhIndex, (j - 1) * 32, (i - 1) * 32, 1, Normal_RGBList(), 0)
                
                If Cont < CInt(Val(frmConfigSup.mLargo)) * CInt(Val(frmConfigSup.mAncho)) Then _
                    Cont = Cont + 1: CurrentGrh.GrhIndex = CurrentGrh.GrhIndex + 1
            Next j
        Next i
        
        CurrentGrh.GrhIndex = CurrentGrh.GrhIndex - Cont
    End If
    
    frmPreview.PreviewGrh.AutoRedraw = False

    Call Engine_EndScene(destRect, frmPreview.PreviewGrh.hWnd)

    Call DrawBuffer.LoadPictureBlt(frmPreview.PreviewGrh.hdc)

    frmPreview.PreviewGrh.AutoRedraw = True

    Call DrawBuffer.PaintPicture(frmPreview.PreviewGrh.hdc, 0, 0, frmPreview.PreviewGrh.Width, frmPreview.PreviewGrh.Height, 0, 0, vbSrcCopy)
End Sub


