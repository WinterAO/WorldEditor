Attribute VB_Name = "modGeneral"
Option Explicit

'Escribe y Lee archivos de texto plano
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function Form_Caption() As String
    Form_Caption = "WinterMapEditor versión: " & App.Major & "." & App.Minor & "." & App.Revision
End Function

Sub Main()

    Call modCarga.pre_leerConfiguracion 'Leemos la config basica para elegir un modo
    
    frmModo.Show
    
    Do While frmModo.ModoElegido = False
        DoEvents
    Loop
    
    frmCarga.Show
    
    Call modCarga.leerConfiguracion 'Leemos el resto de la config
    Call GenerateContra
    Call modCarga.IniciarCabecera
    
    'Tile Engine
    Call InitTileEngine(32, 32, 8, 8)
    
    'Carga de indices
    '------------------------
    frmCarga.lblStatus.Caption = "Cargando Indice de Graficos."
    DoEvents
    Call LoadGrhData
    '------------------------
     
    Unload frmCarga
    frmMain.Show
    
    prgRun = True
    
    Do While prgRun
        
        If frmMain.WindowState <> vbMinimized And frmMain.Visible Then
            'TODO: La movida de las teclas y el engine aqui.
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

    'EngineRun = False
    
    'Stop tile engine
    'Call Engine_DirectX8_End

    'Destruimos los objetos publicos creados
    'Set SurfaceDB = Nothing

    For Each mifrm In Forms
        Unload mifrm
    Next
    
    'Allow MainLoop to close program
    If prgRun = True Then
        prgRun = False
    End If
    
    End

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

