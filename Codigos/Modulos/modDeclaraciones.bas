Attribute VB_Name = "modDeclaraciones"
Option Explicit

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

Public GrhData() As GrhData 'Holds all the grh data

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

Public UserPos As Position 'Posicion

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

Public MapData() As MapBlock 'Holds map data for current map

