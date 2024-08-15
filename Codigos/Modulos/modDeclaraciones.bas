Attribute VB_Name = "modDeclaraciones"
Option Explicit

'#######################
'Recursos
'#######################
Public Const PasswordResources = "$FlLrjB3JoliHdAPKA8&YaJR5"

'#######################
'Configuración
'#######################
Public ModoElegido       As Boolean

Public ProfileTag        As String ' Perfil seleccionado

'#######################
'Objetos publicos
'#######################

Public LucesRedondas     As clsLucesRedondas

Public Inventario        As clsGraphicalInventory

'The main timer of the Carga.
Public MainTimer         As clsTimer

'Totals
Global NumChars          As Integer

Global LastChar          As Integer

'#######################
'Avisos
'#######################
Public Const MSGMod      As String = "Este mapa há sido modificado." & vbCrLf & "Si no lo guardas perderas todos los cambios ¿Deseas guardarlo?"

Public Const MSGDang     As String = "¡CUIDADO! Esta función podria arruinar el mapa." & vbCrLf & "¿Estas seguro que quieres continuar?"

'Renderizado del mapa
Public Radio             As Byte

Public ToWorldMap2       As Boolean

Public HotKeysAllow      As Boolean

Public PATH_Save         As String

Public NumMap_Save       As Integer

Public NameMap_Save      As String

Public nMapaActual       As Integer

Public MapaActual        As String

Public MapaCargado       As Boolean

Public MousePos          As String

Public ParticlePreview   As Long

Public NoSobreescribir   As Boolean

Public SupActual         As Integer

' Objeto de Translado
Public Cfg_TrOBJ         As Integer

'Bucle principal
Public prgRun            As Boolean

'Seleccion
Public SeleccionIX       As Integer

Public SeleccionFX       As Integer

Public SeleccionIY       As Integer

Public SeleccionFY       As Integer

Public SeleccionAncho    As Integer

Public SeleccionAlto     As Integer

Public Seleccionando     As Boolean

Public SeleccionMap()    As MapBlock

Public DeSeleccionOX     As Integer

Public DeSeleccionOY     As Integer

Public DeSeleccionIX     As Integer

Public DeSeleccionFX     As Integer

Public DeSeleccionIY     As Integer

Public DeSeleccionFY     As Integer

Public DeSeleccionAncho  As Integer

Public DeSeleccionAlto   As Integer

Public DeSeleccionando   As Boolean

Public DeSeleccionMap()  As MapBlock

Public CopyX             As Long

Public CopyY             As Long

Public SobreX            As Integer ' Posicion X bajo el Cursor

Public SobreY            As Integer   ' Posicion Y bajo el Cursor

'Map sizes in tiles
Public XMaxMapSize       As Integer

Public YMaxMapSize       As Integer

Public Const XMinMapSize As Integer = 1

Public Const YMinMapSize As Integer = 1

'Where the map borders are.. Set during load
Public MinXBorder        As Integer

Public MaxXBorder        As Integer

Public MinYBorder        As Integer

Public MaxYBorder        As Integer

' DX Config
Public PantallaX         As Integer

Public PantallaY         As Integer

' Client Config
Public ClienteHeight     As Integer

Public ClienteWidth      As Integer

'Contador de tiles
Public ContadorTiles     As Long

Public Const OFFSET_HEAD As Integer = -34

' Tamaño del inventario Quick
Public Const MAX_INVENTORY_SLOTS As Byte = 35

Public Const INVENTORY_SIZE As Byte = 64

'###########################
' Constantes de intervalo
'###########################
Public Enum eIntervalos
    INT_CHANGE_HEADING = 300
End Enum

'Direcciones
Public Enum E_Heading

    SOUTH = 1
    NORTH = 2
    WEST = 3
    EAST = 4

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
    
    mini_map_color As Long

End Type

'Posicion en un mapa
Public Type Position

    X As Integer
    Y As Integer

End Type

'Lista de cuerpos
Public Type BodyData

    Walk(E_Heading.SOUTH To E_Heading.EAST) As Grh
    HeadOffset As Position

End Type

'Lista de cabezas
Public Type HeadData

    Head(E_Heading.SOUTH To E_Heading.EAST) As Grh
    offset As Position

End Type

'Lista de las animaciones de las armas
Type WeaponAnimData

    WeaponWalk(E_Heading.SOUTH To E_Heading.EAST) As Grh

End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData

    ShieldWalk(E_Heading.SOUTH To E_Heading.EAST) As Grh

End Type

'Holds a world position
Public Type WorldPos

    Map As Integer
    X As Integer
    Y As Integer

End Type

'Lista de cabezas
Public Type tHead

    Std As Byte
    Texture As Integer
    startX As Integer
    startY As Integer

End Type

Public heads()  As tHead

Public Cascos() As tHead

Public Type tIndiceCuerpo

    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer

End Type

Public Type tIndiceAtaque

    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer

End Type

Public Type tIndiceFx

    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer

End Type

Public Type tIndiceArmas

    weapon(1 To 4) As Long

End Type

Public Type tIndiceEscudos

    shield(1 To 4) As Long

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
    
    Light_Value(3) As RGBA
    Light As tLight
    
    Particle_Index As Integer
    Particle_Group_Index As Long 'Particle Engine
    
    fX As Grh
    FxIndex As Integer
    
    ZonaIndex As Integer

End Type

'Hold info about each map
Public Type tZonaInfo

    Music As Integer
    name As String
    PK As Boolean
    MagiaSinEfecto As Boolean
    InviSinEfecto As Boolean
    ResuSinEfecto As Boolean
    LuzBase As Long
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Boolean
    RoboNpcsPermitido As Boolean
    InvocarSinEfecto As Boolean
    OcultarSinEfecto As Boolean
    lvlMinimo As Integer
    lvlMaximo As Integer
    Ambient As Integer
    
    Changed As Byte ' flag for WorldEditor
End Type

'Hold info about a character
Public Type Char

    active As Byte
    Heading As Byte
    Pos As Position

    Body As BodyData
    Head As Integer
    Casco As Integer
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    
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
    WeaponAnim As Integer
    CascoAnim As Integer
    ShieldAnim As Integer
    Heading As Byte
    NpcType As Byte

End Type

Public Type ObjData

    name As String 'Nombre del obj
    OBJType As Integer 'Tipo enum que determina cuales son las caract del obj
    GrhIndex As Long ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    Info As String
    Ropaje As Integer 'Indice del grafico del ropaje
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    Texto As String
    Cerrada As Byte
    Subtipo As Byte

End Type

'Quick Superficies
Type Inventory
    name As String
    GrhIndex As Long
End Type

'**********Arrays Publicas************
Public GrhData()            As GrhData 'Holds all the grh data

Public BodyData()           As BodyData

Public HeadData()           As HeadData

Public WeaponAnimData()     As WeaponAnimData

Public ShieldAnimData()     As ShieldAnimData

Public CascoAnimData()      As HeadData

Public MapData()            As MapBlock 'Holds map data for current map

Public CharList(1 To 10000) As Char 'Holds info about all characters on map

Public MapZonas()           As tZonaInfo

Public SupData()            As SupData

Public NpcData()            As NpcData

Public ObjData()            As ObjData
'************************************

Public CantZonas            As Integer

Public UserMoving           As Byte

Public UserPos              As Position 'Posicion

Public MapInfo              As tZonaInfo 'Holds map info for current map

Public CurrentGrh           As Grh

'Variables de estado
Public AddtoUserPos         As Position 'For moving user

Public UserCharIndex        As Integer

Public WalkMode             As Boolean

Public dLastWalk            As Double

'********** OUTSIDE FUNCTIONS ***********

'For KeyInput
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'Para ejecutar el browser y programas externos
Public Const SW_SHOWNORMAL As Long = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
