Attribute VB_Name = "mDx8_Engine"
'CSEH: RegistrarError
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' No matter what you do with DirectX8, you will need to start with
' the DirectX8 object. You will need to create a new instance of
' the object, using the New keyword, rather than just getting a
' pointer to it, since there's nowhere to get a pointer from yet (duh!).

Public DirectX As New DirectX8

' The D3DX8 object contains lots of helper functions, mostly math
' to make Direct3D alot easier to use. Notice we create a new
' instance of the object using the New keyword.
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8

' The Direct3DDevice8 represents our rendering device, which could
' be a hardware or a software device. The great thing is we still
' use the same object no matter what it is
Public DirectDevice As Direct3DDevice8

' The D3DDISPLAYMODE type structure that holds
' the information about your current display adapter.
Public DispMode  As D3DDISPLAYMODE
    
' The D3DPRESENT_PARAMETERS type holds a description of the way
' in which DirectX will display it's rendering.
Public D3DWindow As D3DPRESENT_PARAMETERS

Public SurfaceDB As New clsTextureManager
Public SpriteBatch As New clsBatch

Private Viewport As D3DVIEWPORT8
Private Projection As D3DMATRIX
Private View As D3DMATRIX

Public Engine_BaseSpeed As Single

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamano muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Public ScreenWidth As Long
Public ScreenHeight As Long

Public MainScreenRect As RECT

Public Type TLVERTEX
    x As Single
    y As Single
    Z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Private EndTime As Long

Public Sub Engine_DirectX8_Init()
    On Error GoTo EngineHandler:

    Dim StError As String

    ' Initialize all DirectX objects.
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
    
    StError = "No se pudo inicializar el motor grafico. Por favor, verifique si tiene sus librerias y sus controladores actualizados."
 
    If ClientSetup.OverrideVertexProcess > 0 Then
        
        Select Case ClientSetup.OverrideVertexProcess
            
            Case 1:
                If Not Engine_Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
                    Call MsgBox(StError)
                    End
                End If
            
            
            Case 2:
                If Not Engine_Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                    Call MsgBox(StError)
                    End
                End If

            
            Case 3:
                If Not Engine_Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                    Call MsgBox(StError)
                    End
                End If
        End Select
        
    Else
    
        'Detectamos el modo de renderizado mas compatible con tu PC.
        If Not Engine_Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not Engine_Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not Engine_Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
            
                    Call MsgBox(StError)
                    End
                
                End If
            End If
        End If
    End If

    'Seteamos la matriz de proyeccion.
    Call D3DXMatrixOrthoOffCenterLH(Projection, 0, ScreenWidth, ScreenHeight, 0, -1#, 1#)
    Call D3DXMatrixIdentity(View)
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
    Call DirectDevice.SetTransform(D3DTS_VIEW, View)

    ' Set rendering options
    Call Engine_Init_RenderStates
    
    'Carga dinamica de texturas por defecto.
    Set SurfaceDB = New clsTextureManager
    
    'Sprite batching.
    Set SpriteBatch = New clsBatch
    Call SpriteBatch.Initialise(2000)
    
    'Inicializamos el resto de sistemas.
    Call Engine_DirectX8_Aditional_Init
    
    EndTime = timeGetTime
    
    Exit Sub
EngineHandler:
    
    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_DirectX8_Init", Erl)
    
    Call CloseMapEditor
End Sub

Private Function Engine_Init_DirectDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean

On Error GoTo ErrorDevice:

    'Establecemos el tamaño del render a lo maximo que podria alcanzar con su resolucion
    ScreenWidth = (Screen.Width \ Screen.TwipsPerPixelX)
    ScreenHeight = (Screen.Height \ Screen.TwipsPerPixelY)

    ' Retrieve the information about your current display adapter.
    Call DirectD3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    
        ' Fill the D3DPRESENT_PARAMETERS type, describing how DirectX should
    ' display it's renders.
    With D3DWindow
        .Windowed = True

        ' The swap effect determines how the graphics get from the backbuffer to the screen.
        ' D3DSWAPEFFECT_DISCARD:
        '   Means that every time the render is presented, the backbuffer
        '   image is destroyed, so everything must be rendered again.
        .SwapEffect = D3DSWAPEFFECT_DISCARD
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = ScreenWidth
        .BackBufferHeight = ScreenHeight
        .hDeviceWindow = frmMain.MainViewPic.hWnd
    End With
    
    If Not DirectDevice Is Nothing Then
        Set DirectDevice = Nothing
    End If
    
    ' Create the rendering device.
    ' Here we request a Hardware or Mixed rasterization.
    ' If your computer does not have this, the request may fail, so use
    ' D3DDEVTYPE_REF instead of D3DDEVTYPE_HAL if this happens. A real
    ' program would be able to detect an error and automatically switch device.
    ' We also request software vertex processing, which means the CPU has to
    Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DWindow.hDeviceWindow, D3DCREATEFLAGS, D3DWindow)
    
    'Lo pongo xq es bueno saberlo...
    Select Case D3DCREATEFLAGS
    
        Case D3DCREATE_MIXED_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: MIXED"
        
        Case D3DCREATE_HARDWARE_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: HARDWARE"
            
        Case D3DCREATE_SOFTWARE_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: SOFTWARE"
            
    End Select
    
    'Everything was successful
    Engine_Init_DirectDevice = True
    
    Exit Function
    
ErrorDevice:
    
    'Destroy the D3DDevice so it can be remade
    Set DirectDevice = Nothing

    'Return a failure
    Engine_Init_DirectDevice = False
    
    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_Init_DirectDevice", Erl)
    
End Function

Private Sub Engine_Init_RenderStates()

On Error GoTo EngineHandler:

    'Set the render states
    With DirectDevice
    
        Call .SetVertexShader(D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
        Call .SetRenderState(D3DRS_LIGHTING, False)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)
        Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        Call .SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
        
    End With
    
    Exit Sub
    
EngineHandler:
    
    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_Init_RenderStates", Erl)
End Sub

'CSEH: RegistrarError
Public Sub Engine_DirectX8_End()
'***************************************************
'Author: Standelf
'Last Modification: 26/05/2010
'Destroys all DX objects
'***************************************************
On Error Resume Next

On Error GoTo EngineHandler:

    Dim i As Byte
    
    '   DeInit Lights
    Call LucesRedondas.DeInit_LightEngine
    
    '   Clean Particles
    Call Particle_Group_Remove_All
    
    '   Clean Texture
    If Not DirectDevice Is Nothing Then _
        Call DirectDevice.SetTexture(0, Nothing)
    
    '   Borrar DBI Surface
    Call CleanDrawBuffer
    
    Set DirectD3D8 = Nothing
    Set DirectD3D = Nothing
    Set DirectX = Nothing
    Set DirectDevice = Nothing
    Set SpriteBatch = Nothing
    Set Sound = Nothing
    Set LucesRedondas = Nothing
    Set Inventario = Nothing
    
    '   Erase Data
    Erase MapData()
    Erase CharList()
    
    Exit Sub
    
EngineHandler:
    
    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_DirectX8_End", Erl)
    
End Sub

Public Sub Engine_DirectX8_Aditional_Init()
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/12/2010
'**************************************************************

    On Error GoTo EngineHandler:

    FPS = 101
    FramesPerSecCounter = 101

    TileBufferSize = ClientSetup.TilesBuffer
    Engine_BaseSpeed = 0.5
    
    With MainScreenRect
        .Bottom = frmMain.MainViewPic.ScaleHeight
        .Right = frmMain.MainViewPic.ScaleWidth
    End With

    'Inicializamos y cargamos los graficos de las Fonts.
    Call mDx8_Text.Engine_Init_FontTextures
    
    If Not prgRun Then
        
        Set LucesRedondas = New clsLucesRedondas
        Set Inventario = New clsGraphicalInventory
        
        ' Seteamos algunos colores por adelantado y unica vez.
        Call RGBAList(COLOR_WHITE(), 255, 255, 255, 255)
        
        ' Inicializamos otros sistemas.
        Call mDx8_Text.Engine_Init_FontSettings
        Call Init_MeteoEngine
        
        ' Inicializa DIB surface, un buffer usado para dejar imagenes estaticas en PictureBox
        Call PrepareDrawBuffer
        
        'Inicializamos el inventario grafico
        Call Inventario.Initialize(DirectD3D8, frmQuick.picInv, MAX_INVENTORY_SLOTS, , INVENTORY_SIZE, INVENTORY_SIZE, , , , , True)
        
    End If
    
    Exit Sub
    
EngineHandler:
    
    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_DirectX8_Aditional_Init", Erl)
    
End Sub

Public Sub Engine_Draw_Line(x1 As Single, y1 As Single, x2 As Single, y2 As Single, Optional color As Long = -1, Optional Color2 As Long = -1)

On Error GoTo Error
    
    Call Long_2_RGBAList(temp_rgb(), color)
    
    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(x1, y1, x2, y2, temp_rgb())
    
Exit Sub

Error:
    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_Draw_Line", Erl)
    
End Sub

Public Sub Engine_Draw_Point(x1 As Single, y1 As Single, Optional color As Long = -1)

On Error GoTo Error
    
    Call Long_2_RGBAList(temp_rgb(), color)
    
    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(x1, y1, 0, 1, temp_rgb(), 0, 0)
    
Exit Sub

Error:
    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_Draw_Point", Erl)
    
End Sub

Public Function Engine_ElapsedTime() As Long
'**************************************************************
'Gets the time that past since the last call
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ElapsedTime
'**************************************************************

    On Error GoTo Engine_ElapsedTime_Err

    Dim Start_Time As Long

    'Get current time
    Start_Time = timeGetTime

    'Calculate elapsed time
    Engine_ElapsedTime = Start_Time - EndTime

    'Get next end time
    EndTime = Start_Time
    
    Exit Function

Engine_ElapsedTime_Err:

    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_ElapsedTime", Erl)
End Function

Public Function Engine_PixelPosX(ByVal x As Long) As Long
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosX
'*****************************************************************
    On Error GoTo Engine_PixelPosX_Err

    Engine_PixelPosX = (x - 1) * 32
    
    Exit Function
    
Engine_PixelPosX_Err:

    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_PixelPosX", Erl)

End Function

Public Function Engine_PixelPosY(ByVal y As Long) As Long
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosY
'*****************************************************************
    On Error GoTo Engine_PixelPosY_Err

    Engine_PixelPosY = (y - 1) * 32
    
    Exit Function
    
Engine_PixelPosY_Err:

    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_PixelPosY", Erl)
End Function

Public Sub Engine_Draw_Box(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, color As Long)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | Render Box
'***************************************************
    On Error GoTo Engine_Draw_Box_Err

    Call Long_2_RGBAList(temp_rgb(), color)

    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(x, y, Width, ByVal Height, temp_rgb())
    
    Exit Sub
    
Engine_Draw_Box_Err:

    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_Draw_Box", Erl)

End Sub

Private Function Engine_Collision_Between(ByVal value As Single, ByVal Bound1 As Single, ByVal Bound2 As Single) As Byte
'*****************************************************************
'Find if a value is between two other values (used for line collision)
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Between
'*****************************************************************
    On Error GoTo Engine_Collision_Between_Err

    'Checks if a value lies between two bounds
    If Bound1 > Bound2 Then
        If value >= Bound2 Then
            If value <= Bound1 Then Engine_Collision_Between = 1
        End If
    Else
        If value >= Bound1 Then
            If value <= Bound2 Then Engine_Collision_Between = 1
        End If
    End If
    
    Exit Function
    
Engine_Collision_Between_Err:

    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_Collision_Between", Erl)

End Function

Public Function Engine_Collision_Line(ByVal L1X1 As Long, ByVal L1Y1 As Long, ByVal L1X2 As Long, ByVal L1Y2 As Long, ByVal L2X1 As Long, ByVal L2Y1 As Long, ByVal L2X2 As Long, ByVal L2Y2 As Long) As Byte
'*****************************************************************
'Check if two lines intersect (return 1 if true)
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Line
'*****************************************************************

    On Error GoTo Engine_Collision_Line_Err
    
    Dim m1 As Single
    Dim M2 As Single
    Dim b1 As Single
    Dim b2 As Single
    Dim IX As Single

    'This will fix problems with vertical lines
    If L1X1 = L1X2 Then L1X1 = L1X1 + 1
    If L2X1 = L2X2 Then L2X1 = L2X1 + 1

    'Find the first slope
    m1 = (L1Y2 - L1Y1) / (L1X2 - L1X1)
    b1 = L1Y2 - m1 * L1X2

    'Find the second slope
    M2 = (L2Y2 - L2Y1) / (L2X2 - L2X1)
    b2 = L2Y2 - M2 * L2X2
    
    'Check if the slopes are the same
    If M2 - m1 = 0 Then
    
        If b2 = b1 Then
            'The lines are the same
            Engine_Collision_Line = 1
        Else
            'The lines are parallel (can never intersect)
            Engine_Collision_Line = 0
        End If
        
    Else
        
        'An intersection is a point that lies on both lines. To find this, we set the Y equations equal and solve for X.
        'M1X+B1 = M2X+B2 -> M1X-M2X = -B1+B2 -> X = B1+B2/(M1-M2)
        IX = ((b2 - b1) / (m1 - M2))
        
        'Check for the collision
        If Engine_Collision_Between(IX, L1X1, L1X2) Then
            If Engine_Collision_Between(IX, L2X1, L2X2) Then Engine_Collision_Line = 1
        End If
        
    End If
    
    Exit Function
    
Engine_Collision_Line_Err:

    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_Collision_Line", Erl)

End Function

Public Function Engine_Collision_LineRect(ByVal sX As Long, ByVal sY As Long, ByVal sW As Long, ByVal sH As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Byte
'*****************************************************************
'Check if a line intersects with a rectangle (returns 1 if true)
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_LineRect
'*****************************************************************
    On Error GoTo Engine_Collision_LineRect_Err

    'Top line
    If Engine_Collision_Line(sX, sY, sX + sW, sY, x1, y1, x2, y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If
    
    'Right line
    If Engine_Collision_Line(sX + sW, sY, sX + sW, sY + sH, x1, y1, x2, y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Bottom line
    If Engine_Collision_Line(sX, sY + sH, sX + sW, sY + sH, x1, y1, x2, y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Left line
    If Engine_Collision_Line(sX, sY, sX, sY + sW, x1, y1, x2, y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    Exit Function

Engine_Collision_LineRect_Err:

Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_Collision_LineRect", Erl)

End Function

Function Engine_Collision_Rect(ByVal x1 As Integer, ByVal y1 As Integer, ByVal Width1 As Integer, ByVal Height1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal Width2 As Integer, ByVal Height2 As Integer) As Boolean
'*****************************************************************
'Check for collision between two rectangles
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Rect
'*****************************************************************
    On Error GoTo Engine_Collision_Rect_Err

    If x1 + Width1 >= x2 Then
        If x1 <= x2 + Width2 Then
            If y1 + Height1 >= y2 Then
                If y1 <= y2 + Height2 Then
                    Engine_Collision_Rect = True
                End If
            End If
        End If
    End If

    Exit Function

Engine_Collision_Rect_Err:

    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_Collision_Rect", Erl)

End Function

Public Sub Engine_BeginScene(Optional ByVal color As Long = 0)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | DD Clear & BeginScene
'***************************************************

    On Error GoTo Engine_BeginScene_Err

    Call DirectDevice.BeginScene
    Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, color, 1#, 0)
    Call SpriteBatch.Begin
    
    Exit Sub
    
Engine_BeginScene_Err:

    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_BeginScene", Erl)

End Sub

Public Sub Engine_EndScene(ByRef DestRect As RECT, Optional ByVal hWndDest As Long = 0)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | DD EndScene & Present
'***************************************************

On Error GoTo DeviceHandler:

    Call SpriteBatch.Flush
    
    Call DirectDevice.EndScene
        
    If hWndDest = 0 Then
        Call DirectDevice.Present(DestRect, ByVal 0&, ByVal 0&, ByVal 0&)
    
    Else
        Call DirectDevice.Present(DestRect, ByVal 0, hWndDest, ByVal 0)
    
    End If
    
    Exit Sub
    
DeviceHandler:

    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then

        Call mDx8_Engine.Engine_DirectX8_Init

        Call LoadGraphics

    End If

End Sub

Public Sub Engine_Update_FPS()
    '***************************************************
    'Author: ???
    'Last Modification: ????
    'Calculate $ Limitate (if active) FPS.
    '***************************************************
    On Error GoTo Engine_Update_FPS_Err

    If ClientSetup.LimiteFPS Then
        While (GetTickCount - FPSLastCheck) \ 10 < FramesPerSecCounter
            Call Sleep(5)
        Wend
    End If

    If FPSLastCheck + 1000 < timeGetTime Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 1
        FPSLastCheck = timeGetTime
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1

    End If
    
    Exit Sub

Engine_Update_FPS_Err:

    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_Update_FPS", Erl)

End Sub

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'More info: http://www.vbgore.com/GameClient.TileEn ... e_GetAngle" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************

    Dim SideA As Single
    Dim SideC As Single
 
    On Error GoTo ErrOut
 
    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then
 
        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90
 
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then
 
        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360
 
            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)
 
    'Side B = CenterY
 
    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)
 
    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583
 
    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
 
    'Exit function
 
Exit Function
 
    'Check for error
ErrOut:
 
    'Return a 0 saying there was an error
    Engine_GetAngle = 0
 
Exit Function

End Function

Public Sub Engine_Get_ARGB(color As Long, Data As D3DCOLORVALUE)
'**************************************************************
'Author: Standelf
'Last Modify Date: 18/10/2012
'**************************************************************
    On Error GoTo Engine_Get_ARGB_Err
    
    Dim A As Long, R As Long, G As Long, B As Long
        
    If color < 0 Then
        A = ((color And (&H7F000000)) / (2 ^ 24)) Or &H80&
    Else
        A = color / (2 ^ 24)
    End If
    
    R = (color And &HFF0000) / (2 ^ 16)
    G = (color And &HFF00&) / (2 ^ 8)
    B = (color And &HFF&)
    
    With Data
        .A = A
        .R = R
        .G = G
        .B = B
    End With
    
    Exit Sub
        
Engine_Get_ARGB_Err:

    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_Get_ARGB", Erl)

End Sub
