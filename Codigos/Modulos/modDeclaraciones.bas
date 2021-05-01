Attribute VB_Name = "modDeclaraciones"
Option Explicit

'Totals
Global NumChars As Integer
Global LastChar As Integer

Public Const MSGMod As String = "Este mapa há sido modificado." & vbCrLf & "Si no lo guardas perderas todos los cambios ¿Deseas guardarlo?"
Public Const MSGDang As String = "CUIDADO! Este comando puede arruinar el mapa." & vbCrLf & "¿Estas seguro que desea continuar?"

Public HotKeysAllow As Boolean
Public PATH_Save As String
Public NumMap_Save As Integer
Public NameMap_Save As String
Public MapaActual As Integer
Public MapaCargado As Boolean
Public MousePos As String
Public ParticlePreview As Long
Public NoSobreescribir As Boolean

' Objeto de Translado
Public Cfg_TrOBJ As Integer

Public prgRun As Boolean

Public SeleccionIX As Integer
Public SeleccionFX As Integer
Public SeleccionIY As Integer
Public SeleccionFY As Integer
Public SeleccionAncho As Integer
Public SeleccionAlto As Integer
Public Seleccionando As Boolean
Public SeleccionMap() As MapBlock

Public DeSeleccionOX As Integer
Public DeSeleccionOY As Integer
Public DeSeleccionIX As Integer
Public DeSeleccionFX As Integer
Public DeSeleccionIY As Integer
Public DeSeleccionFY As Integer
Public DeSeleccionAncho As Integer
Public DeSeleccionAlto As Integer
Public DeSeleccionando As Boolean
Public DeSeleccionMap() As MapBlock

'Ver Capas/Npc/Objetos/etc
Public VerBlockeados As Boolean
Public VerTriggers As Boolean
Public VerGrilla As Boolean ' grilla
Public VerParticulas As Boolean
Public VerCapa1 As Boolean
Public VerCapa2 As Boolean
Public VerCapa3 As Boolean
Public VerCapa4 As Boolean
Public VerTranslados As Boolean
Public VerObjetos As Boolean
Public VerNpcs As Boolean

Public SobreX As Integer ' Posicion X bajo el Cursor
Public SobreY As Integer   ' Posicion Y bajo el Cursor

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
    GrhIndexInt As Integer
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
    
    mini_map_color As Long
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

' Cuerpos body.dat
Public Type tIndiceCuerpo
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

' Lista de Cuerpos body.dat
Public Type tBodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Long
End Type

'Heads list
Public Type tHeadData
    Head(0 To 4) As Grh
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
    bLocked As Byte
    
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

Type SupData
    name As String
    Grh As Long
    Width As Byte
    Height As Byte
    Block As Boolean
    Capa As Byte
End Type

Public Type NpcData
    name As String
    ELV As Integer
    Hostile As Byte
    Body As Integer
    Head As Integer
    Heading As Byte
    NpcType As Byte
End Type

Public Type ObjData
    name As String 'Nombre del obj
    ObjType As Integer 'Tipo enum que determina cuales son las caract del obj
    GrhIndex As Long ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    Info As String
    Ropaje As Integer 'Indice del grafico del ropaje
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    Texto As String
End Type

'**********Arrays Publicas************
Public GrhData() As GrhData 'Holds all the grh data
Public BodyData() As tBodyData
Public HeadData() As tHeadData
Public MapData() As MapBlock 'Holds map data for current map
Public SuperMapData() As MapBlock
Public CharList(1 To 10000) As Char 'Holds info about all characters on map
Public MapZonas() As tMapInfo
Public SupData() As SupData
Public NpcData() As NpcData
Public ObjData() As ObjData
'************************************

Public CantZonas As Integer
Public UserPos As Position 'Posicion
Public MapInfo As tMapInfo 'Holds map info for current map
Public CurrentGrh As Grh

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
