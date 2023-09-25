VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsByteBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                                    ByRef Destination As Any, _
                                    ByRef source As Any, _
                                    ByVal Length As Long)

Private Data() As Byte

Private CurrentPos As Long
Private lastPos As Long
Private hFile As Long

Private Const INVALID_FILE_HANDLER As Integer = -1
Private Const DEFAULT_MAX_SIZE_FILE As Long = 65535
Private Const BYTE_SIZE As Byte = 1
Private Const BOOL_SIZE As Byte = 2
Private Const INTEGER_SIZE As Byte = 2
Private Const LONG_SIZE As Byte = 4
Private Const SINGLE_SIZE As Byte = 4
Private Const DOUBLE_SIZE As Byte = 8

Public Sub initializeReader(ByRef arrayByte() As Byte)
lastPos = UBound(arrayByte)
ReDim Data(lastPos) As Byte

Call CopyMemory(Data(0), arrayByte(0), lastPos + 1)
CurrentPos = 0
End Sub

Public Sub initializeWriter(ByVal fileHandler As Integer)
    ReDim Data(DEFAULT_MAX_SIZE_FILE * 20) As Byte

    hFile = fileHandler
    CurrentPos = 0
    lastPos = -1
End Sub

Public Sub getBytes(ByRef Destination() As Byte, Optional ByVal Length As Long = -1)
If Length >= 0 Then
    Call CopyMemory(Destination(0), Data(CurrentPos), Length)
Else
    Call CopyMemory(Destination(0), Data(0), lastPos + 1)
End If
End Sub

Public Function getByte() As Byte
getByte = Data(CurrentPos)
CurrentPos = CurrentPos + BYTE_SIZE
End Function

Public Function getBoolean() As Boolean
Call CopyMemory(getBoolean, Data(CurrentPos), BOOL_SIZE)
CurrentPos = CurrentPos + BOOL_SIZE
End Function

Public Function getInteger() As Integer
Call CopyMemory(getInteger, Data(CurrentPos), INTEGER_SIZE)
CurrentPos = CurrentPos + INTEGER_SIZE
End Function

Public Function getLong() As Long
Call CopyMemory(getLong, Data(CurrentPos), LONG_SIZE)
CurrentPos = CurrentPos + LONG_SIZE
End Function

Public Function getSingle() As Single
Call CopyMemory(getSingle, Data(CurrentPos), SINGLE_SIZE)
CurrentPos = CurrentPos + SINGLE_SIZE
End Function

Public Function getDouble() As Double
Call CopyMemory(getDouble, Data(CurrentPos), DOUBLE_SIZE)
CurrentPos = CurrentPos + DOUBLE_SIZE
End Function

Public Function getString(Optional ByVal Length As Integer = -1) As String
    Dim ret() As Byte

    If Length < 0 Then
        Length = getInteger()
        getString = getString(Length)
    Else

        If Length > 0 Then
            ReDim ret(Length - 1) As Byte
       
            Call CopyMemory(ret(0), Data(CurrentPos), Length)
       
            getString = StrConv(ret, vbUnicode)
            CurrentPos = CurrentPos + Length
        End If
    End If

End Function

Public Sub putByte(ByVal value As Byte)
Data(lastPos + 1) = value
lastPos = lastPos + BYTE_SIZE
End Sub

Public Sub putBoolean(ByVal value As Boolean)
Call CopyMemory(Data(lastPos + 1), value, BOOL_SIZE)
lastPos = lastPos + BOOL_SIZE
End Sub

Public Sub putInteger(ByVal value As Integer)
Call CopyMemory(Data(lastPos + 1), value, INTEGER_SIZE)
lastPos = lastPos + INTEGER_SIZE
End Sub

Public Sub putLong(ByVal value As Long)
Call CopyMemory(Data(lastPos + 1), value, LONG_SIZE)
lastPos = lastPos + LONG_SIZE
End Sub

Public Sub putSingle(ByVal value As Single)
Call CopyMemory(Data(lastPos + 1), value, SINGLE_SIZE)
lastPos = lastPos + SINGLE_SIZE
End Sub

Public Sub putDouble(ByVal value As Double)
Call CopyMemory(Data(lastPos + 1), value, DOUBLE_SIZE)
lastPos = lastPos + DOUBLE_SIZE
End Sub

Public Sub putString(ByRef str As String, Optional ByVal withLength As Boolean = True)
Dim Length As Long

Length = Len(str)

If withLength Then
    Call putInteger(Length)
    Call putString(str, False)
Else
    If Length > 0 Then
        Call CopyMemory(Data(lastPos + 1), ByVal StrPtr(StrConv(str, vbFromUnicode)), Length)
   
        lastPos = lastPos + Length
    End If
End If
End Sub

Public Sub getVoid(ByVal Length As Integer)
CurrentPos = CurrentPos + Length
End Sub

Public Sub putVoid(ByVal Length As Integer)
lastPos = lastPos + Length
End Sub

Public Sub clearData()
ReDim Data(DEFAULT_MAX_SIZE_FILE) As Byte

CurrentPos = 0
lastPos = -1
hFile = -1
End Sub

Public Function getLastPos() As Long
    getLastPos = lastPos
End Function

Public Function getCurrentPos() As Long
    getCurrentPos = CurrentPos
End Function

Public Function EOF() As Boolean
    EOF = (CurrentPos > UBound(Data))
End Function

Public Sub saveBuffer()
    Dim buf() As Byte

    If hFile > 0 Then
        ReDim buf(lastPos) As Byte
   
        Call CopyMemory(buf(0), Data(0), lastPos + 1)
        Put hFile, , buf
    End If
End Sub

Private Sub Class_Initialize()
    hFile = INVALID_FILE_HANDLER
End Sub

Private Sub Class_Terminate()
On Error Resume Next

    Erase Data()
End Sub




