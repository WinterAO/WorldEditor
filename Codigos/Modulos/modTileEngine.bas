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

    tX = (UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2)
    tY = (UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2)
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
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As Byte
    
    With CharList(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        addX = nX - X
        addY = nY - Y
        
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
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffset.X = -1 * (TilePixelWidth * addX)
        .MoveOffset.Y = -1 * (TilePixelHeight * addY)
        
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

    If ClientSetup.MeMode = eMeMode.WinterAO Or _
        ClientSetup.MeMode = eMeMode.ArgentumUnited Then
        Call setMapSize(1000, 1000)
    Else
        Call setMapSize(100, 100)
    End If
    
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
        UserPos.X = (XMax / 2) + 25
        UserPos.Y = (YMax / 2) - 25
        
        MinXBorder = XMinMapSize + (ClienteWidth \ 2)
        MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
        MinYBorder = YMinMapSize + (ClienteHeight \ 2)
        MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)
        
        ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
        
        Call frmMain.RefreshMapSize
    
    End If
    
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
    
        Dim cX As Integer
        Dim cY As Integer
        Dim Cuadrante As Integer
        
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
        If AddtoUserPos.Y <> 0 Then
            OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
    
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                OffsetCounterY = 0
                AddtoUserPos.Y = 0
                        
            End If
    
        End If
        
        '****** Update screen ******
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
        
        ' Calculamos los FPS y los mostramos
        Call Engine_Update_FPS
        Call DrawText(10, 5, "FPS: " & modTileEngine.FPS, -1, False)
        Call DrawText(10, 20, "Mouse: " & MousePos, -1, False)
        
        Call ObtenerCuadrante(Cuadrante, cX, cY)
        Call DrawText(10, 35, "Cuadrante: " & Cuadrante & " X:" & cX & " Y: " & cY, -1, False)
        
        If ContadorTiles > 0 Then _
            Call DrawText(10, 50, "Tiles: " & ContadorTiles, -1, False)
        
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

Public Sub Device_Textured_Render(ByVal X As Single, ByVal Y As Single, _
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
                    Call .Draw(X, Y, Width * ScaleX, Height * ScaleY, color, sX / TextureWidth, sY / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, angle)
                Else
                    Call .Draw(X, Y, TextureWidth * ScaleX, TextureHeight * ScaleY, color, , , , , angle)
                End If
                
        End With
        
End Sub

Function LegalPos(X As Integer, Y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************

    LegalPos = True
    
    'Check to see if its out of bounds
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        LegalPos = False
        Exit Function
    End If
    
    'Check to see if its blocked
    If MapData(X, Y).bLocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    'Check for character
    If MapData(X, Y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If

End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, ByVal Animate As Byte, Optional ByVal Alpha As Boolean = False, Optional ByVal angle As Single = 0, Optional ByVal ScaleX As Single = 1!, Optional ByVal ScaleY As Single = 1!)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Long
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
On Error GoTo Error

    If Grh.GrhIndex > grhCount Or GrhData(Grh.GrhIndex).NumFrames = 0 And GrhData(Grh.GrhIndex).FileNum = 0 Then
        Call InitGrh(Grh, GRH_ERROR) ' 23829
        Call AddtoRichTextBox(frmConsola.StatTxt, "Error en Grh. Posicion: X:" & X & " Y:" & Y, 255, 0, 0)
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
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha, angle, ScaleX, ScaleY)
        
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

Sub Draw_GrhIndex(ByVal GrhIndex As Long, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, Optional ByVal Alpha As Boolean = False)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth - TilePixelWidth) \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        'Draw
        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha)
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
    
    Dim Y                As Long     'Keeps track of where on map we are

    Dim X                As Long     'Keeps track of where on map we are
    
    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen
    
    Dim minY             As Long     'Start Y pos on current map

    Dim maxY             As Long     'End Y pos on current map

    Dim minX             As Long     'Start X pos on current map

    Dim maxX             As Long     'End X pos on current map
    
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
    
    '###################
    'PREVIEW
    '###################
    If Val(frmSuperficies.cCapas.Text) >= 1 And (frmSuperficies.cCapas.Text) <= 4 Then
        bCapa = Val(frmSuperficies.cCapas.Text)
        
    Else
        bCapa = 1

    End If
    
    'Draw floor layer
    For Y = screenminY To screenmaxY
    
        For X = screenminX To screenmaxX
            
            'Previsualización
            '*******************************
            If SobreX = X And SobreY = Y Then
                            
                ' Pone Grh !
                Sobre = -1
    
                If frmSuperficies.cSeleccionarSuperficie.value = True And ClientSetup.Preview = True Then
                    Sobre = MapData(X, Y).Graphic(bCapa).GrhIndex
    
                    If frmConfigSup.MOSAICO.value = vbChecked Then

                        Dim aux As Long

                        Dim dy  As Integer

                        Dim dX  As Integer
    
                        If frmConfigSup.DespMosaic.value = vbChecked Then
                            dy = Val(frmConfigSup.DMLargo.Text)
                            dX = Val(frmConfigSup.DMAncho.Text)
                        Else
                            dy = 0
                            dX = 0
    
                        End If
    
                        If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                            aux = Val(frmSuperficies.cGrh.Text) + (((Y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dX) Mod frmConfigSup.mAncho.Text)
    
                            If MapData(X, Y).Graphic(bCapa).GrhIndex <> aux Then
                                MapData(X, Y).Graphic(bCapa).GrhIndex = aux
                                InitGrh MapData(X, Y).Graphic(bCapa), aux
    
                            End If
    
                        Else
                            aux = Val(frmSuperficies.cGrh.Text) + (((Y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dX) Mod frmConfigSup.mAncho.Text)
    
                            If MapData(X, Y).Graphic(bCapa).GrhIndex <> aux Then
                                MapData(X, Y).Graphic(bCapa).GrhIndex = aux
                                InitGrh MapData(X, Y).Graphic(bCapa), aux
    
                            End If
    
                        End If
    
                    Else
    
                        If MapData(X, Y).Graphic(bCapa).GrhIndex <> Val(frmSuperficies.cGrh.Text) Then
                            MapData(X, Y).Graphic(bCapa).GrhIndex = Val(frmSuperficies.cGrh.Text)
                            InitGrh MapData(X, Y).Graphic(bCapa), Val(frmSuperficies.cGrh.Text)
    
                        End If
    
                    End If
    
                End If
    
            Else
                
                Sobre = -1
                
            End If

            '###################
            'CAPAS
            '###################

            If InMapBounds(X, Y) Then
    
                PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
                PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY
                
                'Layer 1 **********************************
                If MapData(X, Y).Graphic(1).GrhIndex <> 0 And VerCapa1 Then Call Draw_Grh(MapData(X, Y).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)
    
                'Layer 2 **********************************
                If MapData(X, Y).Graphic(2).GrhIndex <> 0 And VerCapa2 Then Call Draw_Grh(MapData(X, Y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)
                
                If Sobre >= 0 Then
                    If MapData(X, Y).Graphic(bCapa).GrhIndex <> Sobre Then
                        MapData(X, Y).Graphic(bCapa).GrhIndex = Sobre
                        InitGrh MapData(X, Y).Graphic(bCapa), Sobre
                            
                        If MapData(X, Y).Graphic(bCapa).GrhIndex = GRH_ERROR Then MapData(X, Y).Graphic(bCapa).GrhIndex = 0

                    End If
                    
                End If
                

            
            End If
        
            ScreenX = ScreenX + 1
        Next
    
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next
    
    '<----- Layer Obj, Char, 3 ----->
    ScreenY = minYOffset - TileBufferSize

    For Y = minY To maxY
        
        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX

            If InMapBounds(X, Y) Then
            
                PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
                PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
                
                With MapData(X, Y)
                
                    'Object Layer ***********************************
                    If .ObjGrh.GrhIndex <> 0 And VerObjetos Then Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)

                    'Char layer**************************************
                    If .CharIndex <> 0 And VerNpcs Then Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)

                    'Layer 3 *****************************************
                    If .Graphic(3).GrhIndex <> 0 And VerCapa3 Then Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)

                    'Particulas **************************************
                    If .Particle_Group_Index And VerParticulas Then Call mDx8_Particulas.Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetYTemp + 16)
                    
                End With
                
            End If
            
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y
    
    '<----- Layer 4 ----->
    ScreenY = minYOffset - TileBufferSize

    For Y = minY To maxY

        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX
            
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
            
            'Layer 4
            If VerCapa4 Then If MapData(X, Y).Graphic(4).GrhIndex Then Call Draw_Grh(MapData(X, Y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)
            
            If MapData(X, Y).TileExit.Map <> 0 And VerTranslados Then
                Grh.GrhIndex = 3
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, Normal_RGBList(), 1)
                        
            End If
                
            'Show blocked tiles
            If VerBlockeados And MapData(X, Y).bLocked = 1 Then
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

            If VerTriggers Then If MapData(X, Y).Trigger > 0 Then Call DrawText(PixelOffsetXTemp + 5, PixelOffsetYTemp - 13, MapData(X, Y).Trigger, -1, False, 2)
                
            If ClientSetup.MeMode = eMeMode.WinterAO Or _
                ClientSetup.MeMode = eMeMode.ArgentumUnited Then
                
                If frmMain.mnuVerZonas(0).Checked Then 'Zona actual
                    If MapData(X, Y).ZonaIndex = frmZonas.LstZona.ListIndex + 1 And MapData(X, Y).ZonaIndex > 0 Then Call DrawText(PixelOffsetXTemp + 7, PixelOffsetYTemp + 7, "z" & MapData(X, Y).ZonaIndex, -1, False, 1)
                            
                ElseIf frmMain.mnuVerZonas(1).Checked Then 'Todas las zonas
                    
                    If MapData(X, Y).ZonaIndex > 0 Then
                        Call DrawText(PixelOffsetXTemp + 7, PixelOffsetYTemp + 7, "z" & MapData(X, Y).ZonaIndex, colorZona(MapData(X, Y).ZonaIndex), False, 1)
                    End If

                End If

            End If
                    
            If Seleccionando Then
                If X >= SeleccionIX And Y >= SeleccionIY Then
                    If X <= SeleccionFX And Y <= SeleccionFY Then
                        Grh.GrhIndex = 2
                        Grh.FrameCounter = 1
                        Grh.Started = 0
                                
                        Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)

                    End If

                End If

            End If
             
            ScreenX = ScreenX + 1
            
        Next X

        ScreenY = ScreenY + 1
    Next Y
    
    If EstadoSelect = eEstadoSelect.Copiado Then
        
        Dim sX As Long
        Dim sY As Long
        Dim xX As Long
        Dim yY As Long
                
        For yY = 0 To SeleccionAncho - 1
            For xX = 0 To SeleccionAlto - 1
                            
                ScreenX = (xX + SobreX - UserPos.X) + 21
                ScreenY = yY + SobreY - UserPos.Y + 11
        
                Call Draw_Grh(SeleccionMap(xX, yY).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, MapData(X, Y).Engine_Light(), 1)
                Call Draw_Grh(SeleccionMap(xX, yY).Graphic(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, MapData(X, Y).Engine_Light(), 1)
                sX = sX + 32
            Next xX
            sX = 0
            sY = sY + 32
        Next yY
                        
        sX = 0
        sY = 0
                
    End If
    
    
RenderScreen_Err:

    If Err.Number Then
        Call RegistrarError(Err.Number, Err.Description, "Mod_TileEngine.RenderScreen", Erl)

    End If
    
End Sub

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'*******************************************************
'Esto forma parte del RenderScreen.
'Dibuja todo aquello que tenga cuerpo (por asi decirlo)
'Bichos y PJ
'*******************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    
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
            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 1, False)
        'Dibujamos la Cabeza
        'If .Head Then _
            Call DrawHead(.Head, PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, LightIluminado(), , X, Y)

    End With
End Sub

Public Sub RenderPreview()
'***********************************************
'Autor: Lorwik
'Fecha: ????
'Descripcion: Renderiza la preview de la superficie, objeto, etc seleccionada
'***********************************************

    Dim DestRect     As RECT
    
    Dim i As Integer, j As Integer
    Dim Cont As Integer
    Dim aux As Long
    Dim W As Long, h As Long
    
    With DestRect
        .Bottom = frmPreview.PreviewGrh.ScaleHeight
        .Right = frmPreview.PreviewGrh.ScaleWidth
    End With
    
    'Clear the inventory window
    Call Engine_BeginScene

    If frmSuperficies.Visible = True Then

       If frmConfigSup.MOSAICO.value = vbUnchecked Then
       
           'Call Draw_GrhIndex(CurrentGrh.GrhIndex, (GrhData(CurrentGrh.GrhIndex).pixelWidth) / 2, (GrhData(CurrentGrh.GrhIndex).pixelHeight), 1, Normal_RGBList(), 0)
           
           h = frmConfigSup.mLargo.Text
           If h <= 0 Then h = 1
           W = frmConfigSup.mAncho.Text
           If W <= 0 Then W = 1
           
           aux = Val(CurrentGrh.GrhIndex) + (((1 + 1) Mod h) * W) + ((1 + 1) Mod W)
           Call Draw_GrhIndex(aux, 0, 0, 0, Normal_RGBList(), 0)
    
       Else
           For i = 1 To CInt(Val(frmConfigSup.mLargo))
               For j = 1 To CInt(Val(frmConfigSup.mAncho))
               
                   Call Draw_GrhIndex((CurrentGrh.GrhIndex), j * 32, i * 32, 0, Normal_RGBList(), 0)
                   
                   If Cont < CInt(Val(frmConfigSup.mLargo)) * CInt(Val(frmConfigSup.mAncho)) Then _
                       Cont = Cont + 1: CurrentGrh.GrhIndex = CurrentGrh.GrhIndex + 1
                       
               Next j
           Next i
           
           CurrentGrh.GrhIndex = CurrentGrh.GrhIndex - Cont
       End If
    End If
    
    frmPreview.PreviewGrh.AutoRedraw = False

    Call Engine_EndScene(DestRect, frmPreview.PreviewGrh.hwnd)

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

    Call Engine_EndScene(DestRect, frmParticulas.ParticlePic.hwnd)

    Call DrawBuffer.LoadPictureBlt(frmParticulas.ParticlePic.hDC)

    frmParticulas.ParticlePic.AutoRedraw = True

    Call DrawBuffer.PaintPicture(frmParticulas.ParticlePic.hDC, 0, 0, frmParticulas.ParticlePic.Width, frmParticulas.ParticlePic.Height, 0, 0, vbSrcCopy)
End Sub

Sub MakeChar(CharIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, X As Integer, Y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
On Error Resume Next

    'Update LastChar
    If CharIndex > LastChar Then LastChar = CharIndex
    NumChars = NumChars + 1
    
    With CharList(CharIndex)
    
    'Update head, body, ect.
    If Body > 0 Then _
    .Body = BodyData(Body)
    
    'If Head > 0 Then _
        .Head = HeadData(Head)
        
    .Heading = Heading
    
    'Reset moving stats
    .Moving = 0
    .MoveOffset.X = 0
    .MoveOffset.Y = 0
    
    'Update position
    .Pos.X = X
    .Pos.Y = Y
    
    'Make active
    .active = 1
    
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex

End Sub

Sub EraseChar(CharIndex As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
    If CharIndex = 0 Then Exit Sub
    'Make un-active
    CharList(CharIndex).active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until CharList(LastChar).active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0
    
    'Update NumChars
    NumChars = NumChars - 1

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

