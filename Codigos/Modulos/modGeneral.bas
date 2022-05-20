Attribute VB_Name = "modGeneral"
Option Explicit

'Escribe y Lee archivos de texto plano
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function Form_Caption() As String
    Form_Caption = "WinterMapEditor versión: " & App.Major & "." & App.Minor & "." & App.Revision
End Function

Sub Main()

    'Ruta principal
    IniPath = App.Path & "\"

    Call modCarga.pre_leerConfiguracion 'Leemos la config basica para elegir un modo
    
    frmModo.Show
    
    Do While frmModo.ModoElegido = False
        DoEvents
    Loop
    
    frmCarga.Show
    
    Call GenerateContra
    Call modCarga.leerConfiguracion 'Leemos el resto de la config
    Call modCarga.IniciarCabecera
    
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
    If Not Sound.Initialize_Engine(frmMain.hwnd, DirRecursos, False, (ClientSetup.bSound > 0), (ClientSetup.bMusic <> CONST_DESHABILITADA), ClientSetup.SoundVolume, ClientSetup.MusicVolume, ClientSetup.Invertido) Then
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
    
    frmCarga.lblStatus.Caption = "Cargando Cuerpos."
    DoEvents
    Call modCarga.CargarCuerpos
    
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
    '------------------------
     
    Call modMapas.NuevoMapa
    Unload frmCarga
    frmMain.Show
    
    prgRun = True
    
    Do While prgRun
        
        If frmMain.WindowState <> vbMinimized And frmMain.Visible Then
            Call ShowNextFrame
            
            If frmParticulas.Visible Then Call RenderParticlePreview
            
            Call Sound.Sound_Render
            
            Call CheckKeys
            
            If CurrentGrh.GrhIndex = 0 Then _
                InitGrh CurrentGrh, 1
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

    'Call GuardarConfiguracion

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

    For Each mifrm In Forms
        Unload mifrm
    Next
    
    'Allow MainLoop to close program
    If prgRun = True Then
        prgRun = False
    End If
    
    End

End Sub

Public Sub CheckKeys()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************

    If HotKeysAllow = False Then Exit Sub
    '[Loopzer]
    'If GetKeyState(vbKeyControl) < 0 Then
    '    If Seleccionando Then
    '        If GetKeyState(vbKeyC) < 0 Then CopiarSeleccion
    '        If GetKeyState(vbKeyX) < 0 Then CortarSeleccion
    '        If GetKeyState(vbKeyB) < 0 Then BlockearSeleccion
    '        If GetKeyState(vbKeyD) < 0 Then AccionSeleccion
    ''    Else
    '        If GetKeyState(vbKeyS) < 0 Then DePegar ' GS
    '        If GetKeyState(vbKeyV) < 0 Then PegarSeleccion
    '    End If
    'End If
    '[/Loopzer]
    
    If GetKeyState(vbKeyUp) < 0 Then
        If UserPos.y < YMinMapSize Then Exit Sub ' 10
        If LegalPos(UserPos.X, UserPos.y - 1) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.y = UserPos.y - 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
            dLastWalk = GetTickCount
            
        ElseIf WalkMode = False Then
            UserPos.y = UserPos.y - 1

        End If
        
        Call DibujarMinimapa(True)
        frmMain.SetFocus
        Exit Sub

    End If

    If GetKeyState(vbKeyRight) < 0 Then
        If UserPos.X > XMaxMapSize Then Exit Sub ' 89
        If LegalPos(UserPos.X + 1, UserPos.y) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.X = UserPos.X + 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
            dLastWalk = GetTickCount
            
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X + 1
            
        End If
        
        Call DibujarMinimapa(True)
        frmMain.SetFocus
        Exit Sub

    End If

    If GetKeyState(vbKeyDown) < 0 Then
        If UserPos.y > YMaxMapSize Then Exit Sub ' 92
        
        If LegalPos(UserPos.X, UserPos.y + 1) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.y = UserPos.y + 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
            dLastWalk = GetTickCount
            
        ElseIf WalkMode = False Then
            UserPos.y = UserPos.y + 1
            
        End If
        
        Call DibujarMinimapa(True)
        frmMain.SetFocus
        Exit Sub
        
    End If

    If GetKeyState(vbKeyLeft) < 0 Then
        If UserPos.X < XMinMapSize Then Exit Sub ' 12
        If LegalPos(UserPos.X - 1, UserPos.y) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.X = UserPos.X - 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X - 1

        End If

        Call DibujarMinimapa(True)
        frmMain.SetFocus
        Exit Sub

    End If
    
End Sub

Public Sub ToggleWalkMode()
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************
    On Error GoTo ToggleWalkMode_Err

    If WalkMode = False Then
        WalkMode = True
        
    Else
        frmMain.mnuModoCaminata.Checked = False
        WalkMode = False
        
    End If
    
    If Not WalkMode Then
        'Erase character
        Call EraseChar(UserCharIndex)
        MapData(UserPos.X, UserPos.y).CharIndex = 0
        
    Else
        'MakeCharacter
        If LegalPos(UserPos.X, UserPos.y) Then
            Call MakeChar(NextOpenChar(), 1, 1, SOUTH, UserPos.X, UserPos.y)
            UserCharIndex = MapData(UserPos.X, UserPos.y).CharIndex
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

Public Sub ObtenerCuadrante(ByRef Cuadrante As Integer, ByRef tX As Integer, ByRef tY As Integer)
'*****************************************************
'Autor: Lorwik
'Fecha: 03/04/2021
'Descripción: Actualiza las coordenadas ya sean totales o por cuadrantes
'*****************************************************

    Dim cX As Integer
    Dim cY As Integer
    
    cX = Fix((UserPos.X / 100))
    cY = Fix((UserPos.y / 100))
    
    tX = UserPos.X - (cX * 100)
    tY = UserPos.y - (cY * 100)
    
    Cuadrante = cX * cY

End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                    ByVal Text As String, _
                    Optional ByVal Red As Integer = -1, _
                    Optional ByVal Green As Integer, _
                    Optional ByVal Blue As Integer, _
                    Optional ByVal bold As Boolean = False, _
                    Optional ByVal italic As Boolean = False, _
                    Optional ByVal bCrLf As Boolean = True, _
                    Optional ByVal Alignment As Byte = rtfLeft)
    
'****************************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D apperance!
'****************************************************
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martin Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'Jopi 17/08/2019 : Consola transparente.
'Jopi 17/08/2019 : Ahora podes especificar el alineamiento del texto.
'****************************************************
    With RichTextBox
        
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

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
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
    Dim i As Integer
    Dim lastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
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
      
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
      
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
      
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder( _
                            0, _
                            Titulo, _
                            0, _
                            Path_Inicial)
      
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
      
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path
  
Exit Function
'Error
errFunction:
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString
    Call RegistrarError(Err.Number, Err.Description, "Buscar_Carpeta", Erl)
  
End Function

Public Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
    Randomize Timer
    
    RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
    If RandomNumber > UpperBound Then RandomNumber = UpperBound
End Function
