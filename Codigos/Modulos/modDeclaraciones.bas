Attribute VB_Name = "modDeclaraciones"
Option Explicit

Public HotKeysAllow As Boolean
Public prgRun As Boolean

'Map sizes in tiles
Public XMaxMapSize As Integer
Public YMaxMapSize As Integer
Public Const XMinMapSize As Integer = 1
Public Const YMinMapSize As Integer = 1

'Where the map borders are.. Set during load
Public MinXBorder As Integer
Public MaxXBorder As Integer
Public MinYBorder As Integer
Public MaxYBorder As Integer

' DX Config
Public PantallaX As Integer
Public PantallaY As Integer

' Client Config
Public ClienteHeight As Integer
Public ClienteWidth As Integer

'Heading Constants
Public Enum eDireccion
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Long
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Contiene info acerca de donde se puede encontrar un grh tamano y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    speed As Single
    
    Trans As Byte
End Type

'Posicion en un mapa
Public Type Position
    X As Integer
    Y As Integer
End Type

'Holds a world position
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Holds info about a object
Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

Private Type tLight
    RGBCOLOR As D3DCOLORVALUE
    active As Boolean
    map_x As Integer
    map_y As Integer
    range As Byte
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    
    Engine_Light(0 To 3) As Long
    Light As tLight
    
    Particle_Index As Integer
    Particle_Group_Index As Long 'Particle Engine
    
    fX As Grh
    FxIndex As Integer
    
    ZonaIndex As Integer
End Type

'Hold info about each map
Public Type tMapInfo
    Music As String
    name As String
    MapVersion As Integer
    PK As Boolean
    MagiaSinEfecto As Byte
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    LuzBase As Long
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    Changed As Byte ' flag for WorldEditor
    RoboNpcsPermitido As Byte
    InvocarSinEfecto As Byte
    OcultarSinEfecto As Byte
    lvlMinimo As Byte
    ambient As String
    NoEncriptarMP As Byte
End Type

' Lista de Cuerpos body.dat
Public Type tBodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

Public Type tHeadData
    Head(1 To 4) As Grh
End Type

'Hold info about a character
Public Type Char
    active As Byte
    Heading As Byte
    Pos As Position

    Body As tBodyData
    Head As tHeadData
    
    Moving As Byte
    MoveOffset As Position
    
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
End Type

'**********Arrays Publicas************
Public GrhData() As GrhData 'Holds all the grh data
Public MapData() As MapBlock 'Holds map data for current map
Public BodyData() As tBodyData
Public HeadData() As tHeadData
Public CharList(1 To 10000) As Char 'Holds info about all characters on map
Public MapZonas() As tMapInfo
'************************************

Public CantZonas As Integer
Public UserPos As Position 'Posicion
Public MapInfo As tMapInfo 'Holds map info for current map

'Variables de estado
Public AddtoUserPos As Position 'For moving user
Public UserCharIndex As Integer
Public WalkMode As Boolean
Public dLastWalk As Double

'********** OUTSIDE FUNCTIONS ***********

'For KeyInput
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
