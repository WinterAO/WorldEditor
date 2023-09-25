Attribute VB_Name = "mDx8_Particulas"
'*************************************************************
'Este modulo contiene TODOS los procedimientos que conforma
'el Sistema de Particulas ORE.
'*************************************************************

Option Explicit

Private Type RGB
    R As Long
    G As Long
    B As Long
End Type

Public Type Stream
    name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    alphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    speed As Single
    life_counter As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
End Type

Private Type Particle
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
End Type

Private Type Particle_Group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    Particle_Count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alphaBlend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    live As Long
    liv1 As Integer
    liveend As Long
    
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    
    'Added by Juan Martin Sotuyo Dodero
    speed As Single
    life_counter As Long
End Type

Public particle_group_list() As Particle_Group
Dim particle_group_count As Long
Dim particle_group_last As Long

Public TotalStreams As Integer
Public StreamData() As Stream

Public Const PI As Single = 3.14159265358979

Public Sub CargarParticulas()
    Dim buffer()    As Byte
    Dim InfoHead    As INFOHEADER
    Dim loopc As Long
    Dim i As Long
    Dim ColorSet As Long
    Dim LaCabecera  As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(DirRecursos & "Scripts" & modCompression.Formato, LCase$("Particulas.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Particulas.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        frmParticulas.LynxParticulas.Clear
        frmParticulas.LynxParticulas.Redraw = False
        frmParticulas.LynxParticulas.Visible = False
        
        frmParticulas.LynxParticulas.AddColumn "Particula", 0
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong

        TotalStreams = fileBuff.getInteger
        
        'resize StreamData array
        ReDim StreamData(1 To TotalStreams) As Stream
        
        'fill StreamData array with info from Particles.ini
        For loopc = 1 To TotalStreams
            With StreamData(loopc)
                .name = loopc
                
                frmParticulas.LynxParticulas.AddItem loopc
                
                .NumOfParticles = fileBuff.getLong
                .NumGrhs = fileBuff.getLong
                .id = fileBuff.getLong
                .x1 = fileBuff.getLong
                .y1 = fileBuff.getLong
                .x2 = fileBuff.getLong
                .y2 = fileBuff.getLong
                .angle = fileBuff.getLong
                .vecx1 = fileBuff.getLong
                .vecx2 = fileBuff.getLong
                .vecy1 = fileBuff.getLong
                .vecy2 = fileBuff.getLong
                .life1 = fileBuff.getLong
                .life2 = fileBuff.getLong
                .friction = fileBuff.getLong
                .spin = fileBuff.getByte
                .spin_speedL = fileBuff.getSingle
                .spin_speedH = fileBuff.getSingle
                .alphaBlend = fileBuff.getByte
                .gravity = fileBuff.getByte
                .grav_strength = fileBuff.getLong
                .bounce_strength = fileBuff.getLong
                .XMove = fileBuff.getByte
                .YMove = fileBuff.getByte
                .move_x1 = fileBuff.getLong
                .move_x2 = fileBuff.getLong
                .move_y1 = fileBuff.getLong
                .move_y2 = fileBuff.getLong
                .speed = fileBuff.getSingle
                .life_counter = fileBuff.getLong
                
                ReDim .grh_list(1 To .NumGrhs)
                
                For i = 1 To .NumGrhs
                    .grh_list(i) = fileBuff.getLong
                Next i
                
                For ColorSet = 1 To 4
                    .colortint(ColorSet - 1).R = fileBuff.getLong
                    .colortint(ColorSet - 1).G = fileBuff.getLong
                    .colortint(ColorSet - 1).B = fileBuff.getLong
                Next ColorSet
    
            End With
        Next loopc
    
    End If
    
    frmParticulas.LynxParticulas.Visible = True
    frmParticulas.LynxParticulas.Redraw = True
    frmParticulas.LynxParticulas.ColForceFit
    
    DoEvents
    
    Set fileBuff = Nothing

End Sub

Public Function General_Particle_Create(ByVal ParticulaInd As Long, _
                                        ByVal X As Integer, _
                                        ByVal Y As Integer, _
                                        Optional ByVal particle_life As Long = 0) As Long

    Dim rgb_list(0 To 3) As Long
    
    If ParticulaInd = 0 Then Exit Function

    With StreamData(ParticulaInd)
        rgb_list(0) = RGB(.colortint(0).R, .colortint(0).G, .colortint(0).B)
        rgb_list(1) = RGB(.colortint(1).R, .colortint(1).G, .colortint(1).B)
        rgb_list(2) = RGB(.colortint(2).R, .colortint(2).G, .colortint(2).B)
        rgb_list(3) = RGB(.colortint(3).R, .colortint(3).G, .colortint(3).B)
    
        General_Particle_Create = Particle_Group_Create(X, Y, .grh_list, rgb_list(), .NumOfParticles, ParticulaInd, .alphaBlend, IIf(particle_life = 0, .life_counter, particle_life), .speed, , .x1, .y1, .angle, .vecx1, .vecx2, .vecy1, .vecy2, .life1, .life2, .friction, .spin_speedL, .gravity, .grav_strength, .bounce_strength, .x2, .y2, .XMove, .move_x1, .move_x2, .move_y1, .move_y2, .YMove, .spin_speedH, .spin)

    End With

End Function

Public Function Particle_Group_Remove(ByVal Particle_Group_Index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(Particle_Group_Index) Then
        Particle_Group_Destroy Particle_Group_Index
        Particle_Group_Remove = True
    End If
End Function

Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
    
    Particle_Group_Remove_All = True
End Function

Public Sub Particle_Group_Render(ByVal Particle_Group_Index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martin Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    
    If Particle_Group_Index > UBound(particle_group_list) Then Exit Sub
    
    'If GetTickCount - particle_group_list(Particle_Group_Index).live > (particle_group_list(Particle_Group_Index).liv1 * 25) And Not particle_group_list(Particle_Group_Index).liv1 = -1 Then
    '    Call Particle_Group_Destroy(Particle_Group_Index)
    '    Exit Sub
    'End If
        
    With particle_group_list(Particle_Group_Index)
    
        'Set colors
        temp_rgb(0) = .rgb_list(0)
        temp_rgb(1) = .rgb_list(1)
        temp_rgb(2) = .rgb_list(2)
        temp_rgb(3) = .rgb_list(3)

        'See if it is time to move a particle
        .frame_counter = .frame_counter + timerTicksPerFrame
        If .frame_counter > .frame_speed Then
            .frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
            
        'If it's still alive render all the particles inside
        For loopc = 1 To .Particle_Count

            'Render particle
            Particle_Render .particle_stream(loopc), _
                        screen_x, screen_y, _
                        .grh_index_list(Round(RandomNumber(1, .grh_index_count), 0)), _
                        temp_rgb(), _
                        .alphaBlend, no_move, _
                        .x1, .y1, .angle, _
                        .vecx1, .vecx2, _
                        .vecy1, .vecy2, _
                        .life1, .life2, _
                        .fric, .spin_speedL, _
                        .gravity, .grav_strength, _
                        .bounce_strength, .x2, _
                        .y2, .XMove, _
                        .move_x1, .move_x2, _
                        .move_y1, .move_y2, _
                        .YMove, .spin_speedH, _
                        .spin
        Next loopc
                
        If no_move = False Then
            'Update the group alive counter
            If .never_die = False Then
                .alive_counter = .alive_counter - 1
            End If
        End If
        
    End With
    
End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alphaBlend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martin Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************

    With temp_particle
    
        If no_move = False Then
        
            If .alive_counter = 0 Then
            
                'Start new particle
                Call InitGrh(.Grh, grh_index)
                .X = RandomNumber(x1, x2) - 16
                .Y = RandomNumber(y1, y2) - 16
                .vector_x = RandomNumber(vecx1, vecx2)
                .vector_y = RandomNumber(vecy1, vecy2)
                .alive_counter = RandomNumber(life1, life2)
                .friction = fric
                
            Else
                
                'Continue old particle
                'Do gravity
                If gravity = True Then
                    
                    .vector_y = .vector_y + grav_strength
                    
                    If .Y > 0 Then
                        'bounce
                        .vector_y = bounce_strength
                    End If
                    
                End If
                
                'Do rotation
                If spin Then .angle = .angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
                If .angle >= 360 Then
                    .angle = 0
                End If
                
                If XMove = True Then .vector_x = RandomNumber(move_x1, move_x2)
                If YMove = True Then .vector_y = RandomNumber(move_y1, move_y2)
            End If
            
            'Add in vector
            .X = .X + (.vector_x \ .friction)
            .Y = .Y + (.vector_y \ .friction)
        
            'decrement counter
             .alive_counter = .alive_counter - 1
        End If
        
        'Draw it
        If .Grh.GrhIndex Then
            Call Draw_Grh(.Grh, .X + screen_x, .Y + screen_y, 1, rgb_list(), 1, True, .angle)
        End If
        
    End With
    
End Sub

Private Function Particle_Group_Next_Open() As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    If particle_group_last = 0 Then
        Particle_Group_Next_Open = 1
        Exit Function
    End If
    
    loopc = 1

    Do Until particle_group_list(loopc).active = False

        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If

        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc
    
    Exit Function
    
ErrorHandler:
    Particle_Group_Next_Open = 1

End Function

Private Function Particle_Group_Check(ByVal Particle_Group_Index As Long) As Boolean

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    
    'check index
    If Particle_Group_Index > 0 And Particle_Group_Index <= particle_group_last Then
        If particle_group_list(Particle_Group_Index).active Then
            Particle_Group_Check = True
        End If
    End If

End Function

Private Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal Particle_Count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alphaBlend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'Returns the particle_group_index if successful, else 0
'Modified by Juan Martin Sotuyo Dodero
'Modified by Augusto Jose Rando
'**************************************************************
    
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Call Particle_Group_Make(Particle_Group_Create, map_x, map_y, Particle_Count, stream_type, grh_index_list(), rgb_list(), alphaBlend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin)
        End If
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Call Particle_Group_Make(Particle_Group_Create, map_x, map_y, Particle_Count, stream_type, grh_index_list(), rgb_list(), alphaBlend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin)
    End If

End Function

Private Function Particle_Group_Find(ByVal id As Long) As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    'Find the index related to the handle
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
        loopc = 1

    Do Until particle_group_list(loopc).id = id

        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If

        loopc = loopc + 1
        
    Loop
    
    Particle_Group_Find = loopc
    
    Exit Function
    
ErrorHandler:
    Particle_Group_Find = 0

End Function

Private Function Particle_Get_Type(ByVal Particle_Group_Index As Long) As Byte

    On Error GoTo ErrorHandler:
    
    Particle_Get_Type = particle_group_list(Particle_Group_Index).stream_type
    
    Exit Function
    
ErrorHandler:
    Particle_Get_Type = 0

End Function

Private Sub Particle_Group_Destroy(ByVal Particle_Group_Index As Long)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    On Error Resume Next

    Dim temp As Particle_Group
    Dim i    As Integer

    With particle_group_list(Particle_Group_Index)

        If .map_x > 0 And .map_y > 0 Then
            MapData(.map_x, .map_y).Particle_Group_Index = 0
        End If

    End With

    particle_group_list(Particle_Group_Index) = temp
    
    'Update array size
    If Particle_Group_Index = particle_group_last Then

        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1

            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub

            End If

        Loop
        
        ReDim Preserve particle_group_list(1 To particle_group_last) As Particle_Group

    End If

    particle_group_count = particle_group_count - 1

End Sub

Private Sub Particle_Group_Make(ByVal Particle_Group_Index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal Particle_Count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alphaBlend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martin Sotuyo Dodero
'*****************************************************************
    'Update array size
    If Particle_Group_Index > particle_group_last Then
        particle_group_last = Particle_Group_Index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    
    With particle_group_list(Particle_Group_Index)
        'Make active
        .active = True
        
        'Map pos
        If (map_x <> -1) And (map_y <> -1) Then
            .map_x = map_x
            .map_y = map_y
        End If
        
        'Grh list
        ReDim .grh_index_list(1 To UBound(grh_index_list))
        .grh_index_list() = grh_index_list()
        .grh_index_count = UBound(grh_index_list)
        
        'Sets alive vars
        If alive_counter = -1 Then
            .alive_counter = -1
            .liv1 = -1
            .never_die = True
        Else
            .alive_counter = alive_counter
            .liv1 = alive_counter
            .never_die = False
        End If
        
        'alpha blending
        .alphaBlend = alphaBlend
        
        'stream type
        .stream_type = stream_type
        
        'speed
        .frame_speed = frame_speed
        
        .x1 = x1
        .y1 = y1
        .x2 = x2
        .y2 = y2
        .angle = angle
        .vecx1 = vecx1
        .vecx2 = vecx2
        .vecy1 = vecy1
        .vecy2 = vecy2
        .life1 = life1
        .life2 = life2
        .fric = fric
        .spin = spin
        .spin_speedL = spin_speedL
        .spin_speedH = spin_speedH
        .gravity = gravity
        .grav_strength = grav_strength
        .bounce_strength = bounce_strength
        .XMove = XMove
        .YMove = YMove
        .move_x1 = move_x1
        .move_x2 = move_x2
        .move_y1 = move_y1
        .move_y2 = move_y2
        
        .rgb_list(0) = rgb_list(0)
        .rgb_list(1) = rgb_list(1)
        .rgb_list(2) = rgb_list(2)
        .rgb_list(3) = rgb_list(3)
        
        'handle
        .id = id
        
        .live = GetTickCount()
        
        'create particle stream
        .Particle_Count = Particle_Count
        ReDim .particle_stream(1 To Particle_Count)
    
    End With
    
    'plot particle group on map
    If (map_x <> -1 And map_x <> 0) And (map_y <> -1 And map_x <> 0) Then
        MapData(map_x, map_y).Particle_Group_Index = Particle_Group_Index
    End If
    
End Sub

Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).Particle_Group_Index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function


