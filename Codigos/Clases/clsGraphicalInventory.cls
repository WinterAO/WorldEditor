VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martin Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Private InventoryOffset As Long             'Number of lines we scrolled down from topmost
Private InvSelectedItem As Long             'Currently selected item
Private MouseDownSelectedItem As Long
Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer
Private prevSelItem As Long
Private ShowValue As Boolean
Private ValueFormated As Boolean

Private ImgContainer As Boolean
Private InventoryTickCount As Long

' Variables drag drop
Private TempItem As Integer
Private TempSlot As Integer

Private MouseX As Integer
Private MouseY As Integer
Private DragItem As Byte
Private Drag As Byte

Private DrawBuffer As cDIBSection ' Para guardar temporalmente una "foto" del inventario (WyroX)

Private isInventory As Boolean ' Para separar al inventario del main, de los otros

Private Inventory() As Inventory    'User's inventory

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************

    Set DrawBuffer = Nothing

End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Public Sub ScrollInventory(ByVal Up As Boolean)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 14/01/2010
'Scrolls the graphical inventory up / down
'14/01/2010: ZaMa - Now redraws inventory only if needed
'***************************************************

    'Only move if there is something to show up /down
    If CanScroll(Up) Then
    
        InventoryOffset = InventoryOffset + IIf(Up, 1, -1)
        
        'Redraw the inventory if needed
        Call DrawInventory
        
    End If

End Sub

Public Function CanScroll(ByVal Up As Boolean)

    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 14/01/2010
    'Determines wether inventory can be scrolled up / down
    '***************************************************
    
    If Up Then
        CanScroll = (InventoryOffset + (InventoryWindow.ScaleHeight \ ownTilePixelHeight)) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) < UBound(Inventory)
    
    Else
        CanScroll = InventoryOffset > 0

    End If
    
End Function

Public Sub DeselectItem()
'***************************************************
'Author: ZaMa
'Last Modify Date: 30/11/2009
'Deselect the currently selected item
'***************************************************
    If InvSelectedItem = 0 Then Exit Sub
    
    Dim ClearSlot As Byte
        ClearSlot = InvSelectedItem
    
    'Select nothing
    InvSelectedItem = 0
    
    'Redraw old item to deselect it
    Call DrawInventory

End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get GrhIndex(ByVal Slot As Byte) As Long
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    GrhIndex = Inventory(Slot).GrhIndex
End Property

Public Property Get MaxObjs() As Byte
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Retrieves the capacity of the Inventory
'***************************************************
    MaxObjs = UBound(Inventory)
End Property

Public Sub SetMaxObjs(ByVal MaxObjs As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Set the capacity of the Inventary
'***************************************************
    If UBound(Inventory) = MaxObjs Then Exit Sub
    
    ReDim Preserve Inventory(1 To MaxObjs) As Inventory
End Sub

Public Sub ToggleShowValue(ByVal bShow As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 11/12
'Indicates if the amount of the items in the inventory is shown
'***************************************************
    ShowValue = bShow
    
    Dim Slot As Integer
    ' Update Inventory
        Call DrawInventory
End Sub

Public Sub ClearAllSlots()
'***************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Clear All Slots.
'***************************************************

    Dim Slot As Long
    For Slot = 1 To Me.MaxObjs
        Call SetItem(Slot, 0)
    Next Slot

End Sub

Public Sub SetItem(ByVal Slot As Byte, ByVal eGrhIndex As Long)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If Slot < 1 Or Slot > UBound(Inventory) Then Exit Sub
    
    With Inventory(Slot)
        .GrhIndex = eGrhIndex
    End With
    
    'If InventoryWindow.Visible = False Then InventoryWindow.Visible = True
    
    'Render inventory slot (we don't need to render the whole inventory)
    Call DrawInventory
End Sub

Public Function GetItem(ByVal GrhIndex As Long) As Byte
'**************************************************
'Autor: Lorwik
'Fecha: 30/04/2024
'Descripcion: Busca si existe el GrhIndex pasado por
'parametros y si es así devuelve el slot donde se encuentra
'**************************************************

    Dim i As Byte
    
    'Buscamos
    For i = 1 To MAX_INVENTORY_SLOTS
        
        'Si coincide devolvemos la posicion
        If GrhIndex = Inventory(i).GrhIndex Then
            GetItem = i
            Exit Function
        End If
    
    Next i
    
    'Si no hubo resultados devolvemos 0
    GetItem = 0

End Function

Public Function GetFreeSlot() As Byte
'**************************************************
'Autor: Lorwik
'Fecha: 30/04/2024
'Descripcion: Busca un slot libre
'**************************************************

    Dim i As Byte
    
    'Buscamos
    For i = 1 To MAX_INVENTORY_SLOTS
        
        'Si coincide devolvemos la posicion
        If Inventory(i).GrhIndex = 0 Then
            GetFreeSlot = i
            Exit Function
        End If
    
    Next i
    
    'Si no hubo resultados devolvemos 0
    GetFreeSlot = 0

End Function

Private Function ClickItem(ByVal x As Long, ByVal y As Long) As Long

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Selects the item clicked if it's valid and return's it's index
    '***************************************************
    Dim TempItem As Long
    Dim temp_x   As Long
    Dim temp_y   As Long
    
    temp_x = x \ ownTilePixelWidth
    temp_y = y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    TempSlot = TempItem

    'Make sure it's within limits
    If TempItem <= UBound(Inventory) Then

        'Make sure slot isn't empty
        If Inventory(TempItem).GrhIndex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If

    End If

End Function

Public Sub DrawInventory()

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 18/01/2020
    'Renders the inventory to the given PictureBox.
    '***************************************************
    Dim loopc        As Long
    Dim x            As Integer
    Dim y            As Integer
    Dim DestRect     As RECT
    Dim ColorFinal(3) As RGBA
    Dim i As Byte
    Dim GrhIndex As Long
    
    With DestRect
        .Bottom = InventoryWindow.ScaleHeight
        .Right = InventoryWindow.ScaleWidth
    End With
    
    'Clear the inventory window
    Call Engine_BeginScene

    For loopc = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 To UBound(Inventory)
            
        startdX = ((loopc - 1) Mod (frmQuick.picInv.ScaleWidth / 32)) * 32
        startdY = ((loopc - 1) \ (frmQuick.picInv.ScaleWidth / 32)) * 32

        For i = 0 To 3
            ColorFinal(i) = COLOR_WHITE(i)
        Next i
            
        If Inventory(loopc).GrhIndex Then

            ' Draw
            If loopc <> DragItem Then

                GrhIndex = Inventory(loopc).GrhIndex
                
                With GrhData(GrhIndex)
                
                    Call Device_Textured_Render_Inv(startdX, startdY, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, ColorFinal())
                
                End With
                
                If loopc = SelectedItem Then _
                    Call Draw_GrhIndex(2, startdX, startdY, 0, COLOR_WHITE(), 0, False) 'Dibuja el fondo cuando lo seleccionamos
                
                'Numero del Grh
                Call DrawText(startdX, startdY, Inventory(loopc).GrhIndex, COLOR_WHITE)

            End If

        End If

    Next loopc

    If DragItem <> 0 And Drag = 1 And isInventory Then
        Call Draw_GrhIndex(Inventory(DragItem).GrhIndex, MouseX - 16, MouseY - 16, 0, ColorFinal())
        
        ' Si estamos arrastrando un item, dibujamos normal
        InventoryWindow.AutoRedraw = False
        Call Engine_EndScene(DestRect, InventoryWindow.hWnd)
        
    Else
        ' WyroX:
        ' Brujeria necesaria para que no quede negro el inventario.
        ' En realidad: con el AutoRedraw desactivado, copiamos una "foto" del inventario a un buffer en memoria.
        ' Activamos el AutoDraw, lo copiamos de vuelta al inventario y lo dejamos como imagen estatica.

        InventoryWindow.AutoRedraw = False

        Call Engine_EndScene(DestRect, InventoryWindow.hWnd)

        Call DrawBuffer.LoadPictureBlt(InventoryWindow.hDC)

        InventoryWindow.AutoRedraw = True

        Call DrawBuffer.PaintPicture(InventoryWindow.hDC, 0, 0, InventoryWindow.ScaleWidth, InventoryWindow.ScaleHeight, 0, 0, vbSrcCopy)

        InventoryWindow.Picture = InventoryWindow.Image

    End If

End Sub

Public Sub DrawDragAndDrop()

    If ((DragItem <> 0 And Drag = 1) And (timeGetTime - InventoryTickCount > 40)) Then

        Call Inventario.DrawInventory

    End If

End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 13/09/19
    'Implements the mouse move event of the inventory picture box
    'Displays a ToolTip of the item under the mouse
    'Now we show more information about certain items (Recox)
    '***************************************************
    Dim temp_x   As Integer
    Dim temp_y   As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's areanv
    If x < 16 Then x = 16
    If y < 16 Then y = 16
    If x > InventoryWindow.ScaleWidth - 16 Then x = InventoryWindow.ScaleWidth - 16
    If y > InventoryWindow.ScaleHeight - 16 Then y = InventoryWindow.ScaleHeight - 16
        
    temp_x = x \ ownTilePixelWidth
    temp_y = y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
    TempSlot = TempItem

    MouseX = CInt(x)
    MouseY = CInt(y)

End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton And isInventory Then
        If x < 0 Or y < 0 Or x > InventoryWindow.ScaleWidth Or y > InventoryWindow.ScaleHeight Then Exit Sub

        Dim temp_x As Integer
        Dim temp_y As Integer

        temp_x = x \ 32
        temp_y = y \ 32

        'Get the currently clickced item
        DragItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1

        If Not (DragItem <= 25 And DragItem >= 1) Then
            DragItem = 0
            Exit Sub
        Else

            If Inventory(DragItem).GrhIndex = 0 Then
                DragItem = 0
                Exit Sub

            End If

        End If

        If isInventory Then
            Drag = 1
            
        End If

    End If

End Sub

Public Sub Initialize(ByRef DirectD3D As D3DX8, ByRef InvPic As PictureBox, ByVal MaxObjs As Byte, Optional ByVal FontSize As Integer = 7, Optional ByVal TileWidth As Integer = 32, Optional ByVal TileHeight As Integer = 32, Optional ByVal startX As Integer = 0, Optional ByVal startY As Integer = 0, Optional ByVal bImgContainer As Boolean = True, Optional ByVal bShowText As Boolean = True, Optional thisInventory As Boolean = False)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 03/12/09
    'Sets the reference to the picture box on which to render the inventory
    '03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
    '***************************************************
    
    Set InventoryWindow = InvPic
    isInventory = thisInventory
    
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
    ' Inicializar buffer
    Set DrawBuffer = New cDIBSection
    Call DrawBuffer.Create(InvPic.ScaleWidth, InvPic.ScaleHeight)
    
    ReDim Inventory(1 To MaxObjs) As Inventory
    
    'Make sure DirectD3D was correctly initialized
    If DirectD3D Is Nothing Then Exit Sub
    
    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    
    ImgContainer = bImgContainer
    ShowValue = bShowText
    
    If ImgContainer Then InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item

End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Implements the mouse up event of the inventory picture box
    'Check outs which item was clicked
    '***************************************************
    
    'Store previously selected item
    Dim prevSelItem As Long

    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(MouseX, MouseY)

    If Drag Then

        If Not DragItem = TempSlot And Not (TempSlot < 1 Or DragItem < 1) Then
            'Call Protocol_Write.WriteMoveItem(DragItem, TempSlot, eMoveType.Inventory)
            InvSelectedItem = TempSlot
        End If

        DragItem = 0
        Drag = 0

    End If

    Call DrawInventory

End Sub

Public Function GetPos(ByVal x As Integer, ByVal y As Integer) As Byte
    
    Dim temp_x As Integer
    Dim temp_y As Integer
    
    temp_x = x \ 32
    temp_y = y \ 32
    
    GetPos = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1

End Function



