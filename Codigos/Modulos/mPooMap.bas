Attribute VB_Name = "mPooMap"
Option Explicit

Sub Map_MoveTo(ByVal Direccion As E_Heading)
      '***************************************************
      'Author: Alejandro Santos (AlejoLp)
      'Last Modify Date: 06/28/2008
      'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
      ' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
      ' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
      ' 06/28/2008: NicoNZ - Saque lo que impedia que si el usuario estaba paralizado se ejecute el sub.
      '***************************************************

      Dim LegalOk As Boolean
      Static lastmovement As Long
    
      Select Case Direccion

            Case E_Heading.NORTH
                  LegalOk = Map_LegalPos(UserPos.X, UserPos.Y - 1)

            Case E_Heading.EAST
                  LegalOk = Map_LegalPos(UserPos.X + 1, UserPos.Y)

            Case E_Heading.SOUTH
                  LegalOk = Map_LegalPos(UserPos.X, UserPos.Y + 1)

            Case E_Heading.WEST
                  LegalOk = Map_LegalPos(UserPos.X - 1, UserPos.Y)
                        
      End Select

      If LegalOk Then
           Call DibujarMinimapa(True)

          Call Char_MovebyHead(UserCharIndex, Direccion)
          Call Char_MoveScreen(Direccion)

      Else

        If (CharList(UserCharIndex).Heading <> Direccion) Then
            If MainTimer.Check(TimersIndex.ChangeHeading) Then
                Call Char_SetHeading(UserCharIndex, Direccion)
            End If
        End If

      End If
        
End Sub

Sub Char_MovebyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
    '*****************************************************************
    'Starts the movement of a character in nHeading direction
    '*****************************************************************

    Dim addx As Integer
    Dim addy As Integer
        
    Dim X    As Integer
    Dim Y    As Integer
        
    Dim nX   As Integer
    Dim nY   As Integer
    
    If (CharIndex <= 0) Then Exit Sub
    
    With CharList(CharIndex)
        
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move

        Select Case nHeading

            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
                                
        End Select
        
        nX = X + addx
        nY = Y + addy
                
        '// Miqueas : Agrego este parchesito para evitar un run time
        If Not (Map_InBounds(nX, nY)) Then Exit Sub

        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        
        MapData(X, Y).CharIndex = 0
         
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy

    End With

'    If CharIndex <> UserCharIndex Then
'        If Not EstaDentroDelArea(nX, nY) Then
'            Call Char_Erase(CharIndex)
'        End If
'    End If

End Sub

Function Map_LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
    '*****************************************************************
    'Author: ZaMa
    'Last Modification: 06/04/2020
    'Checks to see if a tile position is legal, including if there is a casper in the tile
    '*****************************************************************

    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then Exit Function
    
    If WalkMode Then
    
        'Tile Bloqueado?
        If (Map_GetBlocked(X, Y)) Then Exit Function
        
        'CharIndex = (Char_MapPosExits(CInt(X), CInt(Y)))
            
        'Hay un personaje?
        If (CharIndex > 0) Then
        
            If (Map_GetBlocked(UserPos.X, UserPos.Y)) Then
                    
                Exit Function
    
            End If
                
        End If
        
    End If
    
    Map_LegalPos = True
      
End Function

Function Map_InBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
      '*****************************************************************
      'Checks to see if a tile position is in the maps bounds
      '*****************************************************************

      If (X < XMinMapSize) Or (X > XMaxMapSize) Or (Y < YMinMapSize) Or (Y > YMaxMapSize) Then
            Map_InBounds = False

            Exit Function

      End If
    
      Map_InBounds = True
End Function

Public Function Map_GetBlocked(ByVal X As Integer, ByVal Y As Integer) As Boolean
      '*****************************************************************
      'Author: Aaron Perkins - Modified by Juan Martin Sotuyo Dodero
      'Last Modify Date: 10/07/2002
      'Checks to see if a tile position is blocked
      '*****************************************************************

      If (Map_InBounds(X, Y)) Then
            Map_GetBlocked = (MapData(X, Y).bLocked)
      End If

End Function

Sub Char_MoveScreen(ByVal nHeading As E_Heading)
    '******************************************
    'Starts the screen moving in a direction
    '******************************************

    Dim X  As Integer
    Dim Y  As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move

    Select Case nHeading

        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1

    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y

    'Check to see if its out of bounds
    If (tX < MinXBorder) Or (tX > MaxXBorder) Or (tY < MinYBorder) Or (tY > MaxYBorder) Then

        Exit Sub

    Else
        
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
                
        'bTecho = Char_Techo
               
    End If

End Sub

Public Sub Char_SetHeading(ByVal CharIndex As Long, ByVal Heading As Byte)
 
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify by Miqueas150 Date: 24/02/2013
    'Changes the character heading
    '*****************************************************************
    
    'Make sure it's a legal char_index
 
    If Char_Check(CharIndex) Then
 
        With CharList(CharIndex)
               
            .Heading = Heading
 
            Exit Sub
 
        End With
 
    End If
 
End Sub

Private Function Char_Check(ByVal char_index As Integer) As Boolean

    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martin Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (CharList(char_index).Heading > 0)
    End If
    
End Function

Public Sub Char_UserPos()
 
    Dim X As Integer

    Dim Y As Integer
     
    If Char_Check(UserCharIndex) Then
        
        '// Damos valor a las variables asi sacamos la pos del usuario.
        Call Char_MapPosGet(UserCharIndex, X, Y)
               
        'Call frmMain.ActualizarCoordenadas(X, Y)

        Call DibujarMinimapa(True)
 
        Exit Sub
 
    End If

End Sub

Private Sub Char_MapPosGet(ByVal CharIndex As Long, ByRef X As Integer, ByRef Y As Integer)
                                
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 13/12/2013
    '// By Miqueas150
    '
    '*****************************************************************
        
    'Make sure it's a legal char_index
      
    With CharList(CharIndex)
                  
        'Get map pos
        X = .Pos.X
        Y = .Pos.Y
        
    End With
 
End Sub
