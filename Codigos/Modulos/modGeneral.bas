Attribute VB_Name = "modGeneral"
Option Explicit

Private m_Jpeg     As clsJpeg

Private m_FileName As String

Public keysMovementPressedQueue As clsArrayList

'Escribe y Lee archivos de texto plano
Private Declare Function writeprivateprofilestring _
                Lib "kernel32" _
                Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                    ByVal lpKeyname As Any, _
                                                    ByVal lpString As String, _
                                                    ByVal lpFileName As String) As Long

Private Declare Function getprivateprofilestring _
                Lib "kernel32" _
                Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                  ByVal lpKeyname As Any, _
                                                  ByVal lpdefault As String, _
                                                  ByVal lpreturnedstring As String, _
                                                  ByVal nSize As Long, _
                                                  ByVal lpFileName As String) As Long

Public Function Form_Caption() As String
    Form_Caption = "WinterMapEditor versión: " & App.Major & "." & App.Minor & "." & App.Revision

End Function

Sub Main()

    'Ruta principal
    IniPath = App.path & "\"
    
    Call Application.DeleteFile(Application.GetErrorLogFilename())

    'Call modCarga.pre_leerConfiguracion 'Leemos la config basica para elegir un modo
    
    frmPerfil.Show
    
    Do While ModoElegido = False
        DoEvents
    Loop
    
    frmCarga.Show
    
    Call GenerateContra
    Call modCarga.leerPerfil 'Leemos el resto de la config
    
    frmCarga.lblStatus.Caption = "Iniciando motor grafico (1/3)."
    DoEvents
    Call mDx8_Engine.Engine_DirectX8_Init
    
    frmCarga.lblStatus.Caption = "Iniciando motor grafico (2/3)."
    DoEvents
    Call InitTileEngine(32, 32, 8, 8)
    
    frmCarga.lblStatus.Caption = "Iniciando motor grafico (3/3)."
    DoEvents
    Call mDx8_Engine.Engine_DirectX8_Aditional_Init
    
    frmCarga.lblStatus.Caption = "Iniciando motor de Sonido."
    DoEvents
    Set Sound = New clsSoundEngine

    If Not Sound.Initialize_Engine(frmMain.hWnd, dirRecursos_Compressed, False, True, True, ClientSetup.SoundVolume, ClientSetup.MusicVolume, ClientSetup.Invertido) Then
        MsgBox "¡No se ha logrado iniciar el engine de DirectSound! Reinstale los últimos controladores de DirectX. No habrá soporte de audio en el editor.", vbCritical, "Advertencia"
        
    End If
    
    'Carga de indices
    '------------------------
    frmCarga.lblStatus.Caption = "Cargando Indice de Graficos."
    DoEvents
    Call modCarga.LoadGrhData
    
    frmCarga.lblStatus.Caption = "Cargando Minimapa."
    DoEvents
    Call modCarga.CargarMinimapa
    
    frmCarga.lblStatus.Caption = "Cargando Cabezas."
    DoEvents
    Call modCarga.CargarCabezas
    
    frmCarga.lblStatus.Caption = "Cargando Cascos."
    DoEvents
    Call modCarga.CargarCascos
    
    frmCarga.lblStatus.Caption = "Cargando Cuerpos."
    DoEvents
    Call modCarga.CargarCuerpos
    
    frmCarga.lblStatus.Caption = "Cargando Cabezas."
    DoEvents
    Call modCarga.CargarAnimArmas
    
    frmCarga.lblStatus.Caption = "Cargando Cabezas."
    DoEvents
    Call modCarga.CargarAnimEscudos
    
    frmCarga.lblStatus.Caption = "Cargando Indice de NPC's."
    DoEvents
    Call modCarga.CargarIndicesNPC
    
    frmCarga.lblStatus.Caption = "Cargando Indice de Objetos."
    DoEvents
    Call modCarga.CargarIndicesOBJ
    
    frmCarga.lblStatus.Caption = "Cargando Triggers."
    DoEvents
    Call modCarga.CargarIndicesTriggers
    
    frmCarga.lblStatus.Caption = "Cargando Particulas."
    DoEvents
    Call mDx8_Particulas.CargarParticulas
    
    frmCarga.lblStatus.Caption = "Cargando Indice de Superficies."
    DoEvents
    Call modCarga.CargarIndicesSuperficie
    
    frmCarga.lblStatus.Caption = "Iniciando movimiento sensual."
    DoEvents
    'Esto es para el movimiento suave de pjs, para que el pj termine de hacer el movimiento antes de empezar otro
    Set keysMovementPressedQueue = New clsArrayList
    Call keysMovementPressedQueue.Initialize(1, 4)
    
    Set MainTimer = New clsTimer
    '------------------------
     
    Call modMapas.NuevoMapa
    Unload frmCarga
    frmMain.Show
    
    prgRun = True
    
    ' Intervals
    LoadTimerIntervals
    
    Do While prgRun
        
        If frmMain.WindowState <> vbMinimized And frmMain.Visible Then
            Call ShowNextFrame
            
            If frmParticulas.Visible Then Call RenderParticlePreview
            
            Call Sound.Sound_Render
            
            Call CheckKeys
            
            If CurrentGrh.GrhIndex = 0 Then InitGrh CurrentGrh, 1

        End If
    
        DoEvents
        
    Loop
    
    Call CloseMapEditor

End Sub

Public Sub CloseMapEditor()
    '************************************
    'Author: Lorwik
    'Last Modify Date: 26/04/2021
    '************************************

    On Error Resume Next

    Dim mifrm As Form

    Call guardarPerfil

    'Eliminamos los Array
    Erase GrhData
    Erase BodyData
    Erase HeadData
    Erase MapData
    Erase CharList
    Erase MapZonas
    Erase SupData
    Erase NpcData
    Erase ObjData

    EngineRun = False
    
    'Stop tile engine
    Call Engine_DirectX8_End

    'Destruimos los objetos publicos creados
    Set SurfaceDB = Nothing
    Set MainTimer = Nothing

    For Each mifrm In Forms

        Unload mifrm
    Next
    
    'Allow MainLoop to close program
    If prgRun = True Then
        prgRun = False

    End If
    
    End

End Sub

Private Sub AddMovementToKeysMovementPressedQueue()
    If GetKeyState(vbKeyUp) < 0 Then
        If keysMovementPressedQueue.itemExist(vbKeyUp) = False Then keysMovementPressedQueue.Add (vbKeyUp) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(vbKeyUp) Then keysMovementPressedQueue.Remove (vbKeyUp) ' Remueve la tecla que teniamos presionada
    End If

    If GetKeyState(vbKeyDown) < 0 Then
        If keysMovementPressedQueue.itemExist(vbKeyDown) = False Then keysMovementPressedQueue.Add (vbKeyDown) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(vbKeyDown) Then keysMovementPressedQueue.Remove (vbKeyDown) ' Remueve la tecla que teniamos presionada
    End If

    If GetKeyState(vbKeyLeft) < 0 Then
        If keysMovementPressedQueue.itemExist(vbKeyLeft) = False Then keysMovementPressedQueue.Add (vbKeyLeft) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(vbKeyLeft) Then keysMovementPressedQueue.Remove (vbKeyLeft) ' Remueve la tecla que teniamos presionada
    End If

    If GetKeyState(vbKeyRight) < 0 Then
        If keysMovementPressedQueue.itemExist(vbKeyRight) = False Then keysMovementPressedQueue.Add (vbKeyRight) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(vbKeyRight) Then keysMovementPressedQueue.Remove (vbKeyRight) ' Remueve la tecla que teniamos presionada
    End If
End Sub

Private Sub CheckKeys()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************

    If HotKeysAllow = False Then Exit Sub
    
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    
    If UserMoving = 0 Then
        Call AddMovementToKeysMovementPressedQueue
        
        'Move Up
        If keysMovementPressedQueue.GetLastItem() = vbKeyUp Then
            Call Map_MoveTo(NORTH)
            Call Char_UserPos
        End If
                
        'Move Right
        If keysMovementPressedQueue.GetLastItem() = vbKeyRight Then
            Call Map_MoveTo(EAST)
            Call Char_UserPos
        End If
            
        'Move down
        If keysMovementPressedQueue.GetLastItem() = vbKeyDown Then
            Call Map_MoveTo(SOUTH)
            Call Char_UserPos
        End If
            
        'Move left
        If keysMovementPressedQueue.GetLastItem() = vbKeyLeft Then
            Call Map_MoveTo(WEST)
            Call Char_UserPos
        End If
    End If
    
End Sub

Public Sub ToggleWalkMode()

    '*************************************************
    'Author: Unkwown
    'Last modified: 01/04/2024 - Lorwik
    '*************************************************
    On Error GoTo ToggleWalkMode_Err

    If WalkMode = False Then
        WalkMode = True
        If Not frmWalkerSpeed Then frmWalkerSpeed.Show , frmMain
        
    Else
        frmMain.mnuModoCaminata.Checked = False
        WalkMode = False
        
    End If
    
    If Not WalkMode Then
        'Erase character
        Call Char_Erase(UserCharIndex)
        MapData(UserPos.x, UserPos.y).CharIndex = 0
        
    Else

        'MakeCharacter
        If LegalPos(UserPos.x, UserPos.y) Then
            Call Char_Make(NextOpenChar(), 107, 1, SOUTH, UserPos.x, UserPos.y, 1, 11, 81)
            UserCharIndex = MapData(UserPos.x, UserPos.y).CharIndex
            frmMain.mnuModoCaminata.Checked = True
            
        Else
            MsgBox "ERROR: Ubicacion ilegal."
            WalkMode = False
            
        End If

    End If
    
    Exit Sub
    
ToggleWalkMode_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.ToggleWalkMode", Erl)

    Resume Next

End Sub

Public Sub ObtenerCuadranteCompleto(ByRef Cuadrante As Integer, _
                            ByRef tX As Integer, _
                            ByRef tY As Integer)
    '*****************************************************
    'Autor: Lorwik
    'Fecha: 03/04/2021
    'Descripción: Actualiza las coordenadas ya sean totales o por cuadrantes
    '*****************************************************

    Dim cx As Integer

    Dim cy As Integer
    
    cx = Fix((UserPos.x / 100))
    cy = Fix((UserPos.y / 100))
    
    tX = UserPos.x - (cx * 100)
    tY = UserPos.y - (cy * 100)
    
    Cuadrante = cx * cy

End Sub

Public Function CalcularCuadrante(ByVal PosX As Integer, ByVal PosY As Integer) As Integer
    '*****************************************************
    'Autor: Lorwik
    'Fecha: 16/05/2024
    'Descripción: Devuelve el numero del cuadrante actual
    '*****************************************************

    ' Determina el número de cuadrantes en una fila y columna
    Dim CuadrantesPorFila As Integer
    Dim CuadrantesPorColumna As Integer
    CuadrantesPorFila = XMaxMapSize \ 100
    CuadrantesPorColumna = YMaxMapSize \ 100
    
    ' Calcula la posición del cuadrante en el mapa
    Dim CuadranteX As Integer
    Dim CuadranteY As Integer
    CuadranteX = PosX \ 100
    CuadranteY = PosY \ 100
    
    ' Calcula el número único del cuadrante
    CalcularCuadrante = (CuadranteY * CuadrantesPorFila) + CuadranteX + 1
    
End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                     ByVal Text As String, _
                     Optional ByVal Red As Integer = -1, _
                     Optional ByVal Green As Integer, _
                     Optional ByVal Blue As Integer, _
                     Optional ByVal bold As Boolean = False, _
                     Optional ByVal italic As Boolean = False, _
                     Optional ByVal bCrLf As Boolean = True, _
                     Optional ByVal Alignment As Byte = rtfLeft, _
                     Optional ByVal bFecha As Boolean = True)
    
    '****************************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D apperance!
    '****************************************************
    'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
    'Juan Martin Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
    'Jopi 17/08/2019 : Consola transparente.
    'Jopi 17/08/2019 : Ahora podes especificar el alineamiento del texto.
    'Lorwik 20/03/2024: Ahora puedes mostrar la hora en la que se imprimio el mensaje
    '****************************************************

    Dim horaActual As String

    Dim hora       As Integer

    Dim minutos    As Integer
    
    ' Obtener la hora actual en formato de cadena de caracteres
    horaActual = Time
    
    ' Extraer la hora y los minutos
    hora = Hour(horaActual)
    minutos = Minute(horaActual)

    With RichTextBox
    
        If bFecha Then Text = hora & ":" & minutos & "> " & Text
        
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF

        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        ' 0 = Left
        ' 1 = Center
        ' 2 = Right
        .SelAlignment = Alignment

        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        
        .SelText = Text

        ' Esto arregla el bug de las letras superponiendose la consola del frmMain
        If Not RichTextBox = frmConsola.StatTxt Then RichTextBox.Refresh

    End With

End Sub

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal value As String)
    '*****************************************************************
    'Escribe en un archivo de texto plano
    '*****************************************************************
    writeprivateprofilestring Main, Var, value, File
    
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

    '*****************************************************************
    'Lee de un archivo de texto plano
    '*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    '*****************************************************************
    'Comprueba si existe el archivo o directorio
    '*****************************************************************

    FileExist = (Dir$(File, FileType) <> "")

End Function

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String

    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    Dim i         As Integer

    Dim lastPos   As Integer

    Dim CurChar   As String * 1

    Dim FieldNum  As Integer

    Dim Seperator As String
    
    Seperator = Chr(SepASCII)
    lastPos = 0
    FieldNum = 0
    
    For i = 1 To Len(Text)
        CurChar = mid(Text, i, 1)

        If CurChar = Seperator Then
            FieldNum = FieldNum + 1

            If FieldNum = Pos Then
                ReadField = mid(Text, lastPos + 1, (InStr(lastPos + 1, Text, Seperator, vbTextCompare) - 1) - (lastPos))
                Exit Function

            End If

            lastPos = i

        End If

    Next i

    FieldNum = FieldNum + 1
    
    If FieldNum = Pos Then
        ReadField = mid(Text, lastPos + 1)

    End If

End Function

Function Buscar_Carpeta(Optional Titulo As String, _
                        Optional Path_Inicial As Variant) As String
                        
    '******************************************************************
    ' Funcción que abre el cuadro de dialogo y retorna la ruta
    '******************************************************************
  
    On Local Error GoTo errFunction
      
    Dim objShell  As Object

    Dim objFolder As Object

    Dim o_Carpeta As Object
      
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
      
    On Error Resume Next

    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder(0, Titulo, 0, Path_Inicial)
      
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
      
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.path
  
    Exit Function
    'Error
errFunction:
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString
    Call RegistrarError(Err.Number, Err.Description, "Buscar_Carpeta", Erl)
  
End Function

Public Function RandomNumber(ByVal LowerBound As Variant, _
                             ByVal UpperBound As Variant) As Single
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    Randomize Timer
    
    RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound

    If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function

Private Sub LoadTimerIntervals()
    '***************************************************
    'Author: Lorwik
    'Last Modification: 31/03/2024
    'Set the intervals of timers
    '***************************************************
    
    With MainTimer
    
        Call .SetInterval(TimersIndex.ChangeHeading, eIntervalos.INT_CHANGE_HEADING)
    
        'Init timers
        Call .Start(TimersIndex.ChangeHeading)
    
    End With

End Sub

Public Function Client_Screenshot(ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long) As Boolean
    '*******************************
    'Autor: ???
    'Fecha: ???
    '*******************************

    On Error GoTo ErrorHandler

    Dim i     As Long

    Dim Index As Long

    i = 1
    
    Set m_Jpeg = New clsJpeg
    
    '80 Quality
    m_Jpeg.Quality = 100
    
    'Sample the cImage by hDC
    m_Jpeg.SampleHDC hDC, Width, Height
    
    m_FileName = App.path & "\Render\Fotos\WinterME_Foto"
    
    If Dir$(App.path & "\Render\Fotos", vbDirectory) = vbNullString Then
        MkDir (App.path & "\Render\Fotos")

    End If
    
    Do While Dir$(m_FileName & Trim$(str$(i)) & ".jpg") <> vbNullString
        i = i + 1
        DoEvents
    Loop
    
    Index = i
    
    m_Jpeg.Comment = "WinterMapEditor" & " - " & Format$(Date, "dd/mm/yyyy") & " - " & Format$(Time, "hh:mm AM/PM")
    
    'Save the JPG file
    m_Jpeg.SaveFile m_FileName & Trim$(str$(Index)) & ".jpg"
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "¡Captura realizada con exito! Se guardo en " & m_FileName & Trim$(str$(Index)) & ".jpg", 204, 193, 155, 0, 1, , , True)
    
    Set m_Jpeg = Nothing
    
    Client_Screenshot = True
    
    Exit Function

ErrorHandler:
    Call AddtoRichTextBox(frmConsola.StatTxt, "¡Error en la captura!", 204, 193, 155, 0, 1, , , True)
    Client_Screenshot = False
    
End Function

Public Function Max(ByVal A As Variant, ByVal B As Variant) As Variant
    
    On Error GoTo max_Err
    

    If A > B Then
        Max = A
    Else
        Max = B

    End If

    
    Exit Function

max_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.max", Erl)
    Resume Next
    
End Function

Public Function Min(ByVal A As Double, ByVal B As Double) As Variant
    
    On Error GoTo min_Err
    

    If A < B Then
        Min = A
    Else
        Min = B

    End If

    
    Exit Function

min_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.min", Erl)
    Resume Next
    
End Function

