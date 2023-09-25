VERSION 5.00
Begin VB.Form frmColorPicker 
   BackColor       =   &H00424242&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color ambiental"
   ClientHeight    =   5370
   ClientLeft      =   23550
   ClientTop       =   8445
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   422
   ShowInTaskbar   =   0   'False
   Begin WinterMapEditor.lvButtons_H Command1 
      Height          =   435
      Left            =   1650
      TabIndex        =   26
      Top             =   4890
      Width           =   2475
      _extentx        =   4366
      _extenty        =   767
      caption         =   "Seleccionar"
      capalign        =   2
      backstyle       =   2
      font            =   "frmColorPicker.frx":0000
      cfore           =   16777215
      cfhover         =   16777215
      cbhover         =   0
      cgradient       =   0
      gradient        =   3
      mode            =   0
      value           =   0   'False
      cback           =   65280
   End
   Begin VB.PictureBox picBigBox 
      Height          =   3870
      Left            =   225
      MousePointer    =   99  'Custom
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   465
      Width           =   3885
   End
   Begin VB.PictureBox picThinBox 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00C0C0FF&
      ForeColor       =   &H00C0FFC0&
      Height          =   3840
      Left            =   4260
      Picture         =   "frmColorPicker.frx":0028
      ScaleHeight     =   252
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   17
      Top             =   465
      Width           =   315
   End
   Begin VB.TextBox txtHexColor 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4995
      MaxLength       =   6
      TabIndex        =   12
      Text            =   "HexColor"
      Top             =   4230
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   11
      Text            =   "B"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   5280
      TabIndex        =   10
      Text            =   "G"
      Top             =   3345
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   5280
      TabIndex        =   9
      Text            =   "R"
      Top             =   2940
      Width           =   450
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   8
      Text            =   "Brightness"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   7
      Text            =   "Saturation"
      Top             =   2145
      Width           =   420
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   5265
      TabIndex        =   6
      Text            =   "Hue"
      Top             =   1755
      Width           =   435
   End
   Begin VB.OptionButton objOption 
      BackColor       =   &H00404040&
      Caption         =   "B:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   4785
      TabIndex        =   5
      Top             =   3735
      Width           =   495
   End
   Begin VB.OptionButton objOption 
      BackColor       =   &H00404040&
      Caption         =   "G:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4740
      TabIndex        =   4
      Top             =   3345
      Width           =   510
   End
   Begin VB.OptionButton objOption 
      BackColor       =   &H00404040&
      Caption         =   "R:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4740
      TabIndex        =   3
      Top             =   3000
      Width           =   495
   End
   Begin VB.OptionButton objOption 
      BackColor       =   &H00404040&
      Caption         =   "B:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4815
      TabIndex        =   2
      Top             =   2520
      Width           =   480
   End
   Begin VB.OptionButton objOption 
      BackColor       =   &H00404040&
      Caption         =   "S:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   1
      Top             =   2070
      Width           =   480
   End
   Begin VB.OptionButton objOption 
      BackColor       =   &H00404040&
      Caption         =   "H:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   0
      Top             =   1785
      Width           =   465
   End
   Begin VB.Label lblElColor 
      BackStyle       =   0  'Transparent
      Caption         =   "El color negro #000000 es ignorado por el sistema, si quieres establecer un negro, elige uno menos puro."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   240
      TabIndex        =   25
      Top             =   4440
      Width           =   4515
   End
   Begin VB.Line linTriang2Falling 
      BorderColor     =   &H00FFFFFF&
      X1              =   318
      X2              =   327
      Y1              =   184
      Y2              =   189
   End
   Begin VB.Line linTriang2Rising 
      BorderColor     =   &H00FFFFFF&
      X1              =   325
      X2              =   334
      Y1              =   195
      Y2              =   189
   End
   Begin VB.Line linTriang2Vert 
      BorderColor     =   &H00FFFFFF&
      X1              =   318
      X2              =   315
      Y1              =   185
      Y2              =   200
   End
   Begin VB.Label lblThinContainer 
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000007&
      Height          =   3870
      Left            =   4200
      TabIndex        =   24
      Top             =   465
      Width           =   525
   End
   Begin VB.Line linTriang1Falling 
      BorderColor     =   &H00FFFFFF&
      X1              =   277
      X2              =   282
      Y1              =   251
      Y2              =   256
   End
   Begin VB.Line linTriang1Rising 
      BorderColor     =   &H00FFFFFF&
      X1              =   277
      X2              =   282
      Y1              =   261
      Y2              =   256
   End
   Begin VB.Line linTriang1Vert 
      BorderColor     =   &H00FFFFFF&
      X1              =   277
      X2              =   277
      Y1              =   251
      Y2              =   261
   End
   Begin VB.Label lblComplementaryColor 
      BackColor       =   &H80000017&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   5745
      TabIndex        =   23
      ToolTipText     =   "Complementary Color (adds 180° to Hue Angle)."
      Top             =   645
      Width           =   435
   End
   Begin VB.Label lblSuffix 
      BackColor       =   &H00404040&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   5790
      TabIndex        =   21
      Top             =   2550
      Width           =   210
   End
   Begin VB.Label lblSuffix 
      BackColor       =   &H00404040&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   5775
      TabIndex        =   20
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label lblSuffix 
      BackColor       =   &H00404040&
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   5775
      TabIndex        =   19
      Top             =   1755
      Width           =   210
   End
   Begin VB.Label lblOldColor 
      BackColor       =   &H00FFFF80&
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      Top             =   1005
      Width           =   900
   End
   Begin VB.Label lblNewColor 
      Appearance      =   0  'Flat
      BackColor       =   &H0099CCDD&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   495
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   15
      Left            =   5475
      TabIndex        =   14
      Top             =   1845
      Width           =   15
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4785
      TabIndex        =   13
      Top             =   4245
      Width           =   195
   End
   Begin VB.Label lblContainer 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1050
      Left            =   4755
      TabIndex        =   22
      Top             =   465
      Width           =   690
   End
End
Attribute VB_Name = "frmColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim blnDrag                        As Boolean, intSystemColorAngleMax1530 As Integer, bteSaturationMax255 As Byte, bteBrightnessMax255 As Byte

Dim mSngRValue                     As Single, mSngGValue As Single, mSngBValue As Single

Dim blnNotFirstTimeMarker          As Boolean, mBteMarkerOldX As Integer, mBteMarkerOldY As Integer

Dim arLongMarkerColorStore(11, 11) As Long, arsPicPath() As String

Dim mBlnRecentThinBoxPress         As Boolean, mBlnBigBoxReady As Boolean

'Welcome to use, improve and share this utility. It gives you more control than the standard vb-colorpicker.
'Anna-Carin who created this program gives it away for free.
'A SMALL BUG TO FIX IS THAT THE NUDGE FUNCTION OF THE ARROW MARKERS LOSES FOCUS.

'API TO PAINT PIXELS IN picBoxes.
Private Declare Function SetPixelV _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal color As Long) As Byte 'Painting by API is good and fast.

'FINDS THE REAL PATH FOR MyDocuments
Private Declare Function SHGetFolderPath _
                Lib "shfolder" _
                Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, _
                                          ByVal nFolder As Long, _
                                          ByVal hToken As Long, _
                                          ByVal dwReserved As Long, _
                                          ByVal lpszPath As String) As Long

'Shelling html
Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

Private Type HSL 'IS USED FOR THE HSL FUNCTION FROM THE WEBSITE VBspeed.

    Hue As Integer 'FROM 0 To 360.
    Saturation As Byte
    Luminance As Byte

End Type

Function Shell(Program As String, _
               Optional ShowCmd As Long = vbNormalNoFocus, _
               Optional ByVal WorkDir As Variant) As Long

    Dim FirstSpace As Integer, Slash As Integer
    
    If Left(Program, 1) = """" Then
        FirstSpace = InStr(2, Program, """")

        If FirstSpace <> 0 Then
            Program = mid(Program, 2, FirstSpace - 2) & mid(Program, FirstSpace + 1)
            FirstSpace = FirstSpace - 1

        End If

    Else
        FirstSpace = InStr(Program, " ")

    End If
    
    If FirstSpace = 0 Then FirstSpace = Len(Program) + 1
    
    If IsMissing(WorkDir) Then

        For Slash = FirstSpace - 1 To 1 Step -1

            If mid(Program, Slash, 1) = "\" Then Exit For
        Next
    
        If Slash = 0 Then
            WorkDir = CurDir
        ElseIf Slash = 1 Or mid(Program, Slash - 1, 1) = ":" Then
            WorkDir = Left(Program, Slash)
        Else
            WorkDir = Left(Program, Slash - 1)

        End If

    End If
    
    Shell = ShellExecute(0, vbNullString, Left(Program, FirstSpace - 1), LTrim(mid(Program, FirstSpace)), WorkDir, ShowCmd)

    If Shell < 32 Then VBA.Shell Program, ShowCmd 'To raise Error

End Function

Private Sub Command1_Click()
    If ClientSetup.MeMode = eMeMode.WinterAO Then
        MapZonas(frmZonas.LstZona.ListIndex + 1).LuzBase = RGB(Text1(3), Text1(4), Text1(5))
        
    Else
        MapInfo.LuzBase = RGB(Text1(3), Text1(4), Text1(5))
        
    End If
    
    frmMapInfo.LuzMapa.Text = frmColorPicker.Text1(3).Text & "-" & frmColorPicker.Text1(4).Text & "-" & frmColorPicker.Text1(5).Text

    Call AddtoRichTextBox(frmConsola.StatTxt, "Luz de mapa aceptada. Luz: " & RGB(Text1(3), Text1(4), Text1(5)) & ".", 255, 255, 255, False, True, True)

    Call Actualizar_Estado

    Call frmMapInfo.CambiarColorMap

    Me.Visible = False
End Sub

Public Sub Form_Load()

    Dim udtAngelSaturationBrightness As HSL, bteValdRadioKnapp As Byte

    'ReDim Preserve arsPicPath(1) 'arsPicPath NEEDS A FIRST INITIALISATION TO ENABLE THE USE OF Ubound LATER.
    
    Dim Ctr                          As Byte, bteExtraWidth As Byte, bteExtraHeight As Byte

    'STYLING THE FORM. STRANGELY FAILED TO SWITCH THE SCALEMODE TO PIXLES. 1 pixel=20 twips.
    
    mBlnRecentThinBoxPress = True 'TO GET RID OF GREY SQUARES IN THE PICTURE.
    'frmColorPicker.ScaleMode = vbPixels 'RESEMBLING PIXELS.
    frmColorPicker.Width = 6420
    frmColorPicker.Height = 5805
    
    For Ctr = 0 To 2
        Text1(Ctr).Move 351, 117 + Ctr * 25, 30, 21
        objOption(Ctr).Move 320, 120 + Ctr * 25, 33, 17
    Next Ctr
    
    For Ctr = 3 To 5
        Text1(Ctr).Move 350, 196 + (Ctr - 3) * 26, 30, 21
        objOption(Ctr).Move 320, 198 + (Ctr - 3) * 26, 33, 17
    Next Ctr
    
    txtHexColor.Move 336, 281, 56, 19

    Label1.Move 319, 283, 13, 14 'tecknet #
        
    lblNewColor.Move 322, 33, 58, 33
    lblOldColor.Move 322, 66, 58, 33
    lblContainer.Move 321, 32, 60, 68
    lblOldColor.BackColor = lblNewColor.BackColor 'STARTS AT THE SAME COLOR.
    
    picThinBox.Picture = Nothing 'LOADS WITH NOTHING TO GET GET THE JPG-IMAGE OUT OF SIGHT.
    picThinBox.ScaleMode = vbPixels
    
    'CALCULATES THE FRAME WIDTH BEFORE STYLING.
    bteExtraWidth = picThinBox.Width - picThinBox.ScaleWidth 'OUTER MEASURE MINUS ACTUAL INNER MEASURE = FRAMEWIDTH.
    bteExtraHeight = picThinBox.Height - picThinBox.ScaleHeight 'OUTER MEASURE MINUS ACTUAL INNER MEASURE = FRAMEWIDTH.
    picThinBox.Move 284, 31, 19 + bteExtraWidth, 256 + bteExtraHeight 'FRAMES ARE 4 UNITS BROAD. CURIOSITY FACT IS THAT TEH FRAMES OF ALL VBCONTROLS EXCEPT FOR forms ARE MEASURED FROM THE FRAME CENTER, SO YOU ACTUALLY GET HALF THE WIDTH, BUT IT WORKS SINCE VB USE THE SAME LOGIC ALL THE WAY.
    
    lblThinContainer.BackStyle = 0 'Transparent
    lblThinContainer.Left = 284 - 10: lblThinContainer.Top = picThinBox.Top: lblThinContainer.Width = picThinBox.Width + 20: lblThinContainer.Height = picThinBox.Height
        
    linTriang1Vert.x1 = 277: linTriang1Vert.x2 = 277: linTriang1Vert.y1 = 251: linTriang1Vert.y2 = 261
    linTriang1Rising.x1 = 277: linTriang1Rising.x2 = 283: linTriang1Rising.y1 = 261: linTriang1Rising.y2 = 256
    linTriang1Falling.x1 = 277: linTriang1Falling.x2 = 283: linTriang1Falling.y1 = 251: linTriang1Falling.y2 = 256
    
    linTriang2Vert.x1 = 314: linTriang2Vert.x2 = 314: linTriang2Vert.y1 = 251: linTriang2Vert.y2 = 261
    linTriang2Rising.x1 = 309: linTriang2Rising.x2 = 314: linTriang2Rising.y2 = 261: linTriang2Rising.y1 = 256
    linTriang2Falling.x1 = 309: linTriang2Falling.x2 = 314: linTriang2Falling.y2 = 251: linTriang2Falling.y1 = 256
    
    picBigBox.Width = 256 + 4 '256 INCREASING BY 4 SINCE VB PROBABLY CHEATS THE SAME WAY AS IT DID IN picThinBox.
    picBigBox.Height = 256 + 4
    picBigBox.ScaleWidth = 256
    picBigBox.ScaleHeight = 256
    picBigBox.Left = 13
    picBigBox.Top = 31

    'objOption(0) = True 'STATES Hue AS DEFAULT. ***ATT!!!!!  THIS BOOTS THE CLICK ROUTINE TO DECORATE ThinBox AND BigBox.
    Call SplitlblNewColorToRGBboxes 'ALSO THE SYSTEM CONSTANTS OF RGB GETS UPDATED.
    udtAngelSaturationBrightness = RGBToHSL201(lblNewColor.BackColor, True) 'TRUE MEANS THAT HSL IS UPDATING BOTH THE textboxes AND THE systemConstants.
    
    objOption(bteValdRadioKnapp) = True

End Sub

Private Sub objOption_Click(Index As Integer) 'Choosing modus.

    Dim Ctr As Integer
    
    If Index <> 9 And txtHexColor.Left = 286 Then ' Restore HexBox & Combo1.

        'MsgBox "Move HexBox"
        For Ctr = 286 To 336
            txtHexColor.Move Ctr, 281, 56, 20
        Next Ctr

        DoEvents 'Problems with visual jam.

    End If
    
    If Index = 0 Then 'MsgBox "Hue"
        'picThinBox.Visible = True
        Call PaintThinBox(0)
        mBteMarkerOldX = bteSaturationMax255: mBteMarkerOldY = 255 - bteBrightnessMax255

        If mBlnBigBoxReady = True Then Call picBigBox_Colorize 'NO, MAKE THIS EASIER - REDRAW ONLY IF setup HAS FINISHED.
        Call picBigBox_Colorize

    End If
    
    If Index = 1 Then
        'picThinBox.Visible = True
        Call PaintThinBox(1)
        mBteMarkerOldX = intSystemColorAngleMax1530 / 6: mBteMarkerOldY = 255 - bteBrightnessMax255
        Call picBigBox_Colorize 'Speciell design.

    End If
    
    If Index = 2 Then ' "Brightness"
        'MsgBox "Saturation" 'cOLOR ANGEL IS NOW HORIZONTAL FROM LEFT TO RIGHT. TEXTBOXES ARE NOW IMPORTANT.
        'picThinBox.Visible = True
        Call PaintThinBox(2)
        mBteMarkerOldX = intSystemColorAngleMax1530 / 6: mBteMarkerOldY = 255 - bteSaturationMax255
        Call picBigBox_Colorize 'Speciell design.

        'Set objAnyPictureBox = Nothing 'Kanske sparar minne.
    End If
    
    If Index = 3 Then ' "R"
        'picThinBox.Visible = True
        'Call ColorCirkel("Red") 'Speciell design.
        'Call SampleMarkerBackground   'Sparar bakgrunden bakom Marker om där finns någon.
        'Call PaintMarker(mBteMarkerOldX, mBteMarkerOldY) 'Fyll i Markeren igen (om där finns någon).
        Call opt3RedPaintPicThinBox(ByVal Text1(4), Text1(5))
        mBteMarkerOldX = Text1(5): mBteMarkerOldY = 255 - Text1(4)
        Call picBigBox_Colorize 'Speciell design.
        
    End If
    
    If Index = 4 Then ' "G"
        'picThinBox.Visible = True
        'Call ColorCirkel("Green") 'Speciell design.
        'Call SampleMarkerBackground   'Sparar bakgrunden bakom Marker om där finns någon.
        'Call PaintMarker(mBteMarkerOldX, mBteMarkerOldY) 'Fyll i Markeren igen (om där finns någon).
        Call opt4GreenPaintPicThinBox(ByVal Text1(3), Text1(5))
        mBteMarkerOldX = Text1(5): mBteMarkerOldY = 255 - Text1(3)
        Call picBigBox_Colorize 'Speciell design.
    
    End If
    
    If Index = 5 Then ' "B"
        'picThinBox.Visible = True
        'Call ColorCirkel("Blue") 'Speciell design.
        'Call SampleMarkerBackground   'Sparar bakgrunden bakom Marker om där finns någon.
        'Call PaintMarker(mBteMarkerOldX, mBteMarkerOldY) 'Fyll i Markeren igen (om där finns någon).
        Call opt5BluePaintPicThinBox(ByVal Text1(3), Text1(4))
        mBteMarkerOldX = Text1(3): mBteMarkerOldY = 255 - Text1(4)
        Call picBigBox_Colorize 'Speciell design.
    
    End If
    
    Call imgArrowsModeDepending 'MOVING imgArrows
    
    picThinBox.Refresh

End Sub

Public Sub picBigBox_Colorize()

    Dim blnFadeToGrey As Boolean, R As Single, G As Single, B As Single
    
    picBigBox.Visible = False 'HIDES THE SLOW DRAWIING PROCEDURE.
    picBigBox.AutoRedraw = True 'ELSE YOU WONT SEE ANYTHING.
    'If blnNotFirstTimeMarker = True Then Call EraseMarker 'raderar Marker eftersom en ny bakgrundsfärg (11x11) behöver samplas när denna procedur är klar.
    
    'Set objAnyPictureBox = picBigBox 'Ritar om canvas.
    '*****     ********     **********     **********
    If objOption(0) Then 'IN CASE Option(0) WE SHALL FETCH a fully saturated version of color AND MAKE A 3-D FADE.'
        Call Bigbox3D 'NEW VERSION

        'Call FadeToGrey(objAnyPictureBox, 255, False) 'Alla systemkonstanter borde vara uppdaterade redan.
        'Call FadeToBlack(objAnyPictureBox, 255, 255)
    End If
    
    If objOption(1) Then
        'Call RainBowSurface(objAnyPictureBox, 255, blnFadeToGrey, True)
        Call RainBowBigbox(vbFalse, vbTrue) 'FadeToGrey=False & FadeToBlack=True

    End If

    If objOption(2) Then
        'Call RainBowSurface(objAnyPictureBox, 255, True, False)
        Call RainBowBigbox(vbTrue, vbFalse) 'FadeToGrey= true & FadeToBlack=false

    End If
    
    If objOption(3) Then
        Call opt3RedPaintPicBigBox

    End If

    If objOption(4) Then
        Call opt4GreenPaintPicBigBox

    End If

    If objOption(5) Then
        Call opt5BluePaintPicBigBox

    End If
    
    picBigBox.Visible = True 'SHOWS THE PICBOX AFTER THE SLOW DECORATION.
    picThinBox.Visible = True 'IS NEEDED TO SHOW IN CASE THE FORMER MODE WAS POSTCARDVIEW WHICH THUS HIDES ThinBox.
    
    If blnNotFirstTimeMarker = True Then 'IN CASE THERE IS A marker-coordinate...
        Call SampleMarkerBackground   'SAVES THE BACKGROUND OF MARKER IF THERE IS ANY.
        Call PaintMarker(mBteMarkerOldX, mBteMarkerOldY) 'REPAINT THE MARKER (if there is any).
        lblNewColor.BackColor = picBigBox.Point(mBteMarkerOldX, mBteMarkerOldY)

    End If

    If mBlnBigBoxReady = False Then 'PLACES A MARKER AT CORRECT LOCATION AT THE SETUP STAGE.
        blnNotFirstTimeMarker = True 'PASSWORD.

        'MODE DEPENDING NEW MARKER POSITION.
        If objOption(0) Then mBteMarkerOldX = bteSaturationMax255:   mBteMarkerOldY = 255 - bteBrightnessMax255 'Transmitting logical values.
        If objOption(1) Then mBteMarkerOldX = intSystemColorAngleMax1530 / 6: mBteMarkerOldY = 255 - bteBrightnessMax255 'Transmitting logical values.
        If objOption(2) Then mBteMarkerOldX = intSystemColorAngleMax1530 / 6: mBteMarkerOldY = 255 - bteSaturationMax255 'Transmitting logical values.
    
        Call SampleMarkerBackground   'SAVES THE BACKGROUND OF MARKER IF THERE IS ANY.
        Call PaintMarker(mBteMarkerOldX, mBteMarkerOldY) 'REPAINT THE MARKER (if there is any).
        mBlnBigBoxReady = True 'NOW AT LEAST THE FIRST SPONTANEOUS REDRAW HAS FINISHED.

    End If

End Sub

Private Sub picBigBox_MouseMove(Knapp As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    'PROBLEM: GIF-IMAGES ETC WONT REACT WHEN I SAVE THE OLD IMAGE AS A MATRIX. ON THE OTHER HAND I CAN PAINT OVER GIFS.
    Dim lngColor As Long, udtAngelSaturationBrightness As HSL
    
    If blnDrag = False Then Exit Sub 'Baile if mousebutton is not held down.
    
    If X > 255 Then X = 255 'LIMITER.
    If X < 0 Then X = 0
    If Y > 255 Then Y = 255
    If Y < 0 Then Y = 0
    
    'PASTE THE MARKER ON THE LOCATION OF X,Y.*******
    'HIDE THE MARKER FOR CONVENIENS. LET THE MARKER FOLLOW IF THE MOUSEBUTTON IS PRESSED. FIRST ERASE THE OLD MARKER.
    If objOption(0) Then lngColor = HSLToRGB(intSystemColorAngleMax1530, ByVal X, ByVal 255 - Y, True) 'CONVERT AND UPDATE TEXTBOXES.
    If objOption(1) Then lngColor = HSLToRGB(ByVal X * 6, ByVal bteSaturationMax255, ByVal 255 - Y, True): Call PaintThinBox(1) 'CONVERT AND UPDATE TEXTBOXES.
    If objOption(2) Then lngColor = HSLToRGB(ByVal X * 6, ByVal 255 - Y, ByVal bteBrightnessMax255, True): Call PaintThinBox(2) 'CONVERT AND UPDATE TEXTBOXES.
    
    If objOption(3) Then Call BigBoxOpt3Reaction(ByVal X, Y) 'CONVERT AND UPDATE TEXTBOXES.
    If objOption(4) Then Call BigBoxOpt4Reaction(ByVal X, Y) 'CONVERT AND UPDATE TEXTBOXES.
    If objOption(5) Then Call BigBoxOpt5Reaction(ByVal X, Y) 'CONVERT AND UPDATE TEXTBOXES.
    
    mBlnRecentThinBoxPress = False

End Sub

Private Sub picBigBox_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    Dim lngColor As Long

    If mBlnBigBoxReady = False Then MsgBox "mBlnBigBoxReady = False i BigBox MouseDown! There are no colors to show in bigbox": Exit Sub 'Bail if no color in bigbox.
    'picBigBox.Enabled = False 'FAST DOUBLE CLICKS JAMS THE PICTURE. TRY TO DEACTIVATE THE FIRST CLICK.
    blnDrag = True

    'HIDING THE MARKER NOT TO RISK OF GETTING JAM IN MY PROBE.
    If blnNotFirstTimeMarker = True Then
        Call EraseMarker

    End If
    
    If objOption(0) Then lngColor = HSLToRGB(intSystemColorAngleMax1530, ByVal X, ByVal 255 - Y, True) 'CONVERT AND UPDATE TEXTBOXES.
    If objOption(1) Then
        lngColor = HSLToRGB(ByVal X * 6, ByVal bteSaturationMax255, ByVal 255 - Y, True) 'CONVERT AND UPDATE TEXTBOXES.
        Call FadeThinBoxToGrey 'REPAINT ThinBox - FADE SATURATED COLORS; THE SYSTEM CONSTANTS ARE ALREADY UPDATED.
        picThinBox.Refresh

    End If

    If objOption(2) Then
        picThinBox.BackColor = HSLToRGB(ByVal X * 6, ByVal 255 - Y, 255, False) 'SETTING THE BRIGHT COLOR THAT IS TO BE FADED. CONVERTING AND UPDATING TEXTBOXES.
        lngColor = HSLToRGB(ByVal X * 6, ByVal 255 - Y, ByVal bteBrightnessMax255, True)  'UPDATING THE REAL, NONSATURATED SYSTEM CONSTANTS AND lblNewColor.
        'Call FadeToBlack(picThinBox, 19, 255) 'Redraw ThinBox - fade strong colors; System constants are already updated.
        Call FadeThinBoxToBlack 'REPAINTING ThinBox - FADE SATURATED COLORS ; THE SYSTEM CONSTANTS ARE ALREADY UPDATED.
        picThinBox.Refresh

    End If
    
    If objOption(3) Then
        Call BigBoxOpt3Reaction(ByVal X, Y)

    End If

    If objOption(4) Then
        Call BigBoxOpt4Reaction(ByVal X, Y)

    End If

    If objOption(5) Then
        Call BigBoxOpt5Reaction(ByVal X, Y)

    End If
    
    'THESE FILTER OPTIONS ARE SOEWHAT UNPREDICTABLE SO I FEEL MY WAY.
    'vbSrcInvert FOLLOWED BY vbDstInvert APPARENTLY GIVES A TRANSPARENT PICTURE.
    '?OLD CODE? IN CASE OF THE CURSOR COLLIDING WITH THE MARKER, THE MARKER HAS TO ENTIRELY ERASED AND ENTIRELY REPAINTED.
    mBteMarkerOldX = X 'Already changing here in order to get the correct position in SampleMarkerBackground .
    mBteMarkerOldY = Y 'Will be used by erasemarker.
    Call SampleMarkerBackground 'Saving the new backround behind marker now when there's no Cursor in the way.
    
    blnNotFirstTimeMarker = True
    'picBigBox.Enabled = True 'trying to get rid of the artefacts that arise at fast double clicks.

End Sub

Private Sub picBigBox_MouseUp(Knapp As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

    If mBlnBigBoxReady = False Then MsgBox "mBlnBigBoxReady = False!": Exit Sub 'Baile if no color in bigbox.
    
    blnDrag = False

    If X > 255 Then X = 255 'LIMITER
    If X < 0 Then X = 0
    If Y > 255 Then Y = 255
    If Y < 0 Then Y = 0
    
    mBteMarkerOldX = X
    mBteMarkerOldY = Y
    Call SampleMarkerBackground
    Call PaintMarker(X, Y) 'PAINT MARKER ON ITS NEW LOCATION.

    'picBigBox.Enabled = True 'trying to get rid of the artefacts that arise at fast double clicks.
End Sub

Public Sub EraseMarker()

    'If blnSetup
    Dim CtrY As Byte, CtrX As Byte

    For CtrY = 0 To 10
        For CtrX = 0 To 10
            picBigBox.PSet (mBteMarkerOldX - 5 + CtrX, mBteMarkerOldY - 5 + CtrY), arLongMarkerColorStore(CtrX, CtrY)
        Next CtrX
    Next CtrY

End Sub

Private Sub picThinBox_MouseDown(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    ' set flag to start drawing
    mBlnRecentThinBoxPress = True
    blnDrag = True: Call picThinBox_MouseMove(Button, Shift, X, Y) 'REUSING THE UPDATE ROUTINES.

End Sub

Private Sub lblThinContainer_MouseDown(Button As Integer, _
                                       Shift As Integer, _
                                       X As Single, _
                                       Y As Single)

    Dim sngScaleConst As Single

    sngScaleConst = Screen.TwipsPerPixelY 'GIVING ME THE ACTUAL SIZE OF THE PIXELS OF THE SCREEN, HERE = 15.
    
    mBlnRecentThinBoxPress = True
    Y = Y / sngScaleConst 'CONVERTING FROM THE UNIT TWIP TO PIXELS. ATT! PROBLEM! SHOULD BE /20 BUT IS 15.
    blnDrag = True: Call picThinBox_MouseMove(Button, Shift, X, Y) 'REUSING THE UPDATE ROUTINES.

End Sub

Private Sub lblThinContainer_MouseMove(Knapp As Integer, _
                                       Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    Y = Y / 15 'CONVERTING FROM THE UNIT TWIP TO PIXELS. ATT! PROBLEM! SHOULD BE /20 BUT IS 15.
    Call picThinBox_MouseMove(Knapp, Shift, X, Y)

End Sub

Private Sub picThinBox_MouseMove(Knapp As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)

    Dim lngColor As Long, udtAngelSaturationBrightness As HSL
    
    If blnDrag = False Then Exit Sub

    'If Text1(1) = "Saturation" Then Text1(1) = 100 'The program har been started recently.
    If Y < 0 Then Y = 0 'LIMITER
    If Y > 255 Then Y = 255
    'imgArrows.Top = Y + 28 'Animering
    Call TriangelMove(Y) 'ANIMATION
    
    If objOption(0) Then lngColor = HSLToRGB((255 - Y) * 6, ByVal bteSaturationMax255, ByVal bteBrightnessMax255, True): Exit Sub 'Convert and update textboxes.
    If objOption(1) Then lngColor = HSLToRGB(ByVal intSystemColorAngleMax1530, 255 - Y, ByVal bteBrightnessMax255, True): Exit Sub 'Convert and update textboxes.
    If objOption(2) Then lngColor = HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal bteSaturationMax255, 255 - Y, True) 'Convert and update textboxes.
    If objOption(3) Then
        Text1(3) = 255 - Y: udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True) 'Convert and update textboxes.
        lblNewColor.BackColor = RGB(Text1(3), Text1(4), Text1(5))

    End If

    If objOption(4) Then
        Text1(4) = 255 - Y: udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True) 'Convert and update textboxes.
        lblNewColor.BackColor = RGB(Text1(3), Text1(4), Text1(5))

    End If

    If objOption(5) Then
        Text1(5) = 255 - Y: udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True) 'Convert and update textboxes.
        lblNewColor.BackColor = RGB(Text1(3), Text1(4), Text1(5))

    End If
    
End Sub

Private Sub picThinBox_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    ' set flag to start drawing
    blnDrag = False
    Call picBigBox_Colorize '(mSngRValue, mSngGValue, mSngBValue)

End Sub

Private Sub lblThinContainer_MouseUp(Button As Integer, _
                                     Shift As Integer, _
                                     X As Single, _
                                     Y As Single)
    ' set flag to start drawing
    Y = Y / 20 'CONVERTING FROM THE UNIT TWIP TO PIXELS.
    blnDrag = False
    Call picBigBox_Colorize '(mSngRValue, mSngGValue, mSngBValue)

End Sub

Public Sub FadeThinBoxToGrey()

    Dim sng255saturation As Single, sngLokalBrightness As Single, X As Byte, Y As Integer ', YCtr As Integer
    
    sng255saturation = 255: sngLokalBrightness = bteBrightnessMax255
        
    For X = 0 To 19
        Y = 0 'Sets YCtr for making a new countdown.

        Do 'Interesting if there would raise an error, thus a leap directly to EndSub.
            SetPixelV picThinBox.hDC, X, Y, HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal Round(sng255saturation - sng255saturation * Y / 255), ByVal sngLokalBrightness, False)
            Y = Y + 1
        Loop While Y < 256 'Because Y gets to big when the loop has finished.

    Next X

End Sub

Public Sub Bigbox3D()

    Dim sngLokalSaturation As Single, sngLokalBrightness As Single, YRADNOLL As Integer

    Dim sngR256delToBlack  As Single, sngG256delToBlack As Single, sngB256delToBlack As Single

    Dim R                  As Single, G As Single, B As Single, lColor As Long, Y As Integer, X As Integer
    
    sngLokalSaturation = 255: sngLokalBrightness = 255 'There is a need for intense start color.

    'If R > G Then lSuperior = R Else lSuperior = G 'Det skulle gå att halvera denna rutin medelst en superior, men koden blir då svårare att fatta.
    'If B > lSuperior Then lSuperior = B
    '********* Firstly a single fade from saturated to grey on the uppermost row.
    For X = 0 To 255
        SetPixelV picBigBox.hDC, X, YRADNOLL, HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal Round(sngLokalSaturation * X / 255), ByVal sngLokalBrightness, False)
    Next X 'Resets Y for a new row.

    '********* Here will be an FADE TO BLACK for all columns ********

    For X = 255 To 0 Step -1
        'If blnVertical = True Then R = Ro: G = Go: B = Bo ' If line is vertical the reset for a new round.
        lColor = picBigBox.Point(X, 0) 'Reading the uppermost pixel which is to be faded.
        R = lColor And &HFF
        G = (lColor And &HFF00&) \ &H100&
        B = (lColor And &HFF0000) \ &H10000
        sngR256delToBlack = R / 255  'The fraction blocks which lead down to black.
        sngG256delToBlack = G / 255
        sngB256delToBlack = B / 255
        
        For Y = 0 To 255 'Interesting if there would raise an error, thus a leap back to EndSub.
            'objAnyPictureBox.PSet (X, Y), RGB(R, G, B)
            SetPixelV picBigBox.hDC, X, Y, RGB(R, G, B) 'Painting with API.
            R = R - sngR256delToBlack 'Darkening the shade one of a 256:th.
            G = G - sngG256delToBlack
            B = B - sngB256delToBlack
        Next Y
        
        Y = Y - 1 'Because that Y gets too big when the loop is completed.
    Next X

End Sub

Public Sub FadeThinBoxToBlack()

    Dim sngR256delToBlack As Single, sngG256delToBlack As Single, sngB256delToBlack As Single

    Dim R                 As Single, G As Single, B As Single, lColor As Long, X As Byte, Y As Integer
    
    For X = 0 To 19
        lColor = picThinBox.Point(X, 0) 'Reads the uppermost pixel MAX LIGHT which is to be faded.
        R = lColor And &HFF
        G = (lColor And &HFF00&) \ &H100&
        B = (lColor And &HFF0000) \ &H10000
        sngR256delToBlack = R / 255  'Fractions which leads down to black.
        sngG256delToBlack = G / 255
        sngB256delToBlack = B / 255
        'If blnVertical = True Then R = Ro: G = Go: B = Bo 'Om Vertical linje så återställer sig originalfärgen för en ny runda.
        
        For Y = 0 To 255 'Interesting if the is an error, thus a jump directly to EndSub.
            'objAnyPictureBox.PSet (X, Y), RGB(R, G, B)
            SetPixelV picThinBox.hDC, X, Y, RGB(R, G, B) 'Painting with API.
            R = R - sngR256delToBlack 'Darkening the shade of one 256th.
            G = G - sngG256delToBlack
            B = B - sngB256delToBlack
        Next Y

        Y = Y - 1 'Because Y gets too big when loop is complete.
    Next X

End Sub

Public Sub RainBowBigbox(blnFadeToGrey, _
                         blnFadeToBlack) 'Is used by both radiobutton 1 & 2.

    Dim Ctr        As Byte, blnUpdateTextBoxes As Boolean, bteK4243 As Byte

    Dim Saturation As Single, Luminance As Single

    Static intNODE As Integer, YCtr As Integer, XCtr As Integer, intRainbowAngle As Integer

    'There is no risk for getting dull shades since I use the native principal by adding/subtracting values against at constant FF-component.
    'The algoritm gives med decimal values which increases the importance for mathematical models for choosing color, not pic.point.
    'XCtr = X
    intRainbowAngle = 0 'Protects the systemcolorangel
    
    If blnFadeToGrey = vbTrue And blnFadeToBlack = vbFalse Then
        Saturation = 255
        Luminance = bteBrightnessMax255 'Starting value fully saturated. Brightness is to be the same for the whole of bigbox.
    Else
        Saturation = bteSaturationMax255
        Luminance = 255 'Fading from fully bright.

    End If
    
    'For intRainbowAngle = 0 To 1529
    bteK4243 = 42 'Has to alternate between 42 and 43 pixels per colorfield to make even at 256 pixels.
    'For intNODE = 0 To 1275 Step 255

    XCtr = 0 'To255

    For YCtr = 0 To 255

        Do 'X loopen 0 To 255.

            '1 Red in in direction towards yellow. Green is counting up.
            For Ctr = 1 To bteK4243  'Has to alternate between 42 and 43 pixels per colorfield to make even at 256 pixels.

                If blnFadeToBlack Then Luminance = 255 - YCtr 'Round(bteBrightnessMax255 - (bteBrightnessMax255 / 255 * YCtr))
                If blnFadeToGrey Then Saturation = 255 - YCtr 'Round(bteSaturationMax255 - (bteSaturationMax255 / 255 * YCtr))
                intRainbowAngle = intNODE + ((254 * (Ctr - 1)) / (bteK4243 - 1)) 'Wonderful solution: this logic about going from zero to the full value (here 254) I have been seeking for a long time.
                SetPixelV picBigBox.hDC, XCtr, YCtr, HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, False)
                XCtr = XCtr + 1
            Next Ctr '

            If bteK4243 = 43 Then bteK4243 = 42 Else bteK4243 = 43
            intNODE = intNODE + 255 'Bistabile switch.
        Loop While XCtr < 255

        intRainbowAngle = 0 'Painting the last fully red which lies outside the logic.
        picBigBox.PSet (XCtr, YCtr), HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, blnUpdateTextBoxes)
        intNODE = 0: XCtr = 0: Next YCtr

End Sub

Public Sub RainBowThinBox() 'By swapping the XY-vvalues at the call you can paint either horisontal or vertical.

    Dim Ctr           As Byte, blnUpdateTextBoxes As Boolean, bteK4243 As Byte

    Dim blnHorizontal As Boolean, Saturation As Single, Luminance As Single

    Static intNODE    As Integer, YCtr As Integer, XCtr As Integer, intRainbowAngle As Integer

    'There is no risk for getting dull shades since I use the native principal by adding/subtracting values against at constant FF-component.
    'The algoritm gives med decimal values which increases the importance for mathematical models for choosing color, not pic.point.
    'picThinBox.ScaleMode = vbPixels
    intRainbowAngle = 0 'Protecting systemcolorangel
    Saturation = 255: Luminance = 255 'Fully shining colors.
    'If blnFadeToGrey = True And blnFadeToBlack = False Then Saturation = 255: Luminance = bteBrightnessMax255 'Starting value is full saturation. Brightness is to be the same for the whole bigbox.
    'Horizontal or vertical kan be chosen by intKoordSuperior/intKoordInferior.
    'YCtr = 255: If XCtr = YCtr Then blnHorizontal = True
    
    'For intRainbowAngle = 0 To 1529
    bteK4243 = 42 'Has to alternate between 42 and 43 pixels per colorfield to make even at 256 pixels.
    'For intNODE = 0 To 1275 Step 255
    
    'Vertical
    For XCtr = 0 To 19

        Do 'Y loopen 255 To 0.

            '1 Red in in direction towards yellow. Green is counting up.
            For Ctr = 1 To bteK4243  'Has to alternate between 42 and 43 pixels per colorfield to make even at 256 pixels.
                intRainbowAngle = intNODE + ((254 * (Ctr - 1)) / (bteK4243 - 1)) 'Wonderful solution: this logic about going from zero to the full value (here 254) I have been seeking for a long time.
                'objAnyPictureBox.PSet (XCtr, YCtr), HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, blnUpdatetextBoxes)
                SetPixelV picThinBox.hDC, XCtr, YCtr, HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, blnUpdateTextBoxes)
                YCtr = YCtr - 1
            Next Ctr '

            If bteK4243 = 43 Then bteK4243 = 42 Else bteK4243 = 43
            intNODE = intNODE + 255 'Bistabile switch.
        Loop While YCtr > 0

        intRainbowAngle = 0 'Painting the last fully red which is outside the logic of the routine.
        SetPixelV picThinBox.hDC, XCtr, YCtr, HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, blnUpdateTextBoxes)
        intNODE = 0
        YCtr = 255
    Next XCtr

    'End If

End Sub

Private Sub picThinBox_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim Y As Integer, intDirektion As Integer

    'MsgBox "Är i rutinen picThinBox_KeyDown!"
    'Y = imgArrows.Top - 28

    If objOption(0) Then
        If KeyCode = vbKeyUp Then
            intDirektion = 1
            Call NudgeHueValue(ByVal intDirektion)
            Call picBigBox_Colorize

        End If

        If KeyCode = vbKeyDown Then
            intDirektion = -1
            Call NudgeHueValue(ByVal intDirektion)
            Call picBigBox_Colorize

        End If

    End If

    If objOption(1) Then '******
        MsgBox "Add code for radio1! Probably just writing in textbox Saturation!"

    End If

    If objOption(2) Then '*****
        MsgBox "Add code for radio2!"

    End If

End Sub

Public Sub NudgeHueValue(ByVal intDirektion)

    '1530 levels. The triangels are moving every sixth step and are lying on the byte level of 1530/6.
    'RGBtxtboxes tells the nudge level:
    Dim lngColor As Long

    'NudgeValue goes from ZERO to 1536.
    intSystemColorAngleMax1530 = intSystemColorAngleMax1530 + intDirektion 'Calculating the new value of intSystemColorAngleMax1530, thus +1 or -1.

    If intSystemColorAngleMax1530 > 1530 Then intSystemColorAngleMax1530 = 1530 'Limiter.
    If intSystemColorAngleMax1530 < 0 Then intSystemColorAngleMax1530 = 0

    lngColor = HSLToRGB(ByVal intSystemColorAngleMax1530, bteSaturationMax255, bteBrightnessMax255, True) 'lngColor as a function of HSLToRGB. System constants are being updated at the same time.
    Call TriangelMove(255 - (intSystemColorAngleMax1530 / 1530 * 255))  'Moving the triangel.

End Sub

Public Sub SampleMarkerBackground()

    Dim CtrX As Byte, CtrY As Byte

    'Saving the background behind Marker.
    If blnNotFirstTimeMarker = True Then

        For CtrY = 0 To 10
            For CtrX = 0 To 10
                arLongMarkerColorStore(CtrX, CtrY) = picBigBox.Point(mBteMarkerOldX - 5 + CtrX, mBteMarkerOldY - 5 + CtrY)
            Next CtrX
        Next CtrY

    End If

End Sub

Public Sub PaintMarker(X, Y)

    If bteBrightnessMax255 < 200 Then 'White marker if the surroundings are grey.
        picBigBox.Circle (X, Y), 5, vbWhite
        Exit Sub

    End If

    If Text1(0) < 26 Or Text1(0) > 200 Then 'Shades of blue.
        If bteSaturationMax255 > 70 Then ' And bteSaturationMax255 < 150 Then 'White marker if the surroundings are grey..
            picBigBox.Circle (X, Y), 5, vbWhite
            Exit Sub

        End If

    End If

End Sub

Public Sub SplitlblNewColorToRGBboxes() 'Updating the system constants and textboxes regarding to RGB.
    mSngRValue = lblNewColor.BackColor And &HFF: Text1(3) = mSngRValue
    mSngGValue = (lblNewColor.BackColor And &HFF00&) \ &H100&: Text1(4) = mSngGValue
    mSngBValue = (lblNewColor.BackColor And &HFF0000) \ &H10000: Text1(5) = mSngBValue

End Sub

Private Function RGBToHSL201(ByVal RGBValue As Long, _
                             ByVal blnUpdateTextBoxes As Boolean) As HSL

    Dim R    As Long, G As Long, B As Long

    Dim lMax As Long, lMin As Long, lDiff As Long, lSum As Long

    R = RGBValue And &HFF&
    G = (RGBValue And &HFF00&) \ &H100&
    B = (RGBValue And &HFF0000) \ &H10000

    If R > G Then lMax = R: lMin = G Else lMax = G: lMin = R 'Finds the Superior and inferior components.
    If B > lMax Then lMax = B Else If B < lMin Then lMin = B

    lDiff = lMax - lMin
    lSum = lMax + lMin
    'Luminance, thus brightness' Adobe photoshop uses the logic that the site VBspeed regards (regarded) as too primitive = superior decides the level of brightness.
    RGBToHSL201.Luminance = lMax / 255 * 100

    'Saturation******
    If lMax <> 0 Then 'Protecting from the impossible operation of division by zero.
        RGBToHSL201.Saturation = 100 * lDiff / lMax 'The logic of Adobe Photoshops is this simple.
    Else
        RGBToHSL201.Saturation = 0

    End If

    'Hue ************** R is situated at the angel of 360 eller noll degrees; G vid 120 degrees; B vid 240 degrees. intSystemColorAngleMax1530
    Dim q As Single

    If lDiff = 0 Then q = 0 Else q = 60 / lDiff 'Protecting from the impossible operation of division by zero.

    Select Case lMax

        Case R

            If G < B Then
                RGBToHSL201.Hue = 360& + q * (G - B)
                intSystemColorAngleMax1530 = (360& + q * (G - B)) * 4.25 'Converting from degrees to my resolution of detail.
            Else
                RGBToHSL201.Hue = q * (G - B)
                intSystemColorAngleMax1530 = (q * (G - B)) * 4.25

            End If

        Case G
            RGBToHSL201.Hue = 120& + q * (B - R) ' (R - G)
            intSystemColorAngleMax1530 = (120& + q * (B - R)) * 4.25

        Case B
            RGBToHSL201.Hue = 240& + q * (R - G)
            intSystemColorAngleMax1530 = (240& + q * (R - G)) * 4.25
    End Select 'The case of B was missing.

    'End If
    If blnUpdateTextBoxes = True Then

        'txtHexColor = Hex$(R * 65536 + G * 256 + B): txtHexColor.Refresh 'Applying to internetstandard<>VBstandard
        If R < &H10 Then
            txtHexColor = Right$("00000" & Hex$(R * 65536 + G * 256 + B), 6) 'Adds letters of zero to the left which is a necessary so called padding.
        Else
            txtHexColor = Hex$(R * 65536 + G * 256 + B)

        End If

        txtHexColor.Refresh 'End of hexabox routine.
        Text1(0) = Round(intSystemColorAngleMax1530 / 1530 * 360)

        If lMax = 0 Then
            bteSaturationMax255 = 0 'Protecting from the impossible operation of division by zero.
        Else
            bteSaturationMax255 = 255 * lDiff / lMax
            Text1(1) = RGBToHSL201.Saturation '= saturation both 0 To 255 and 0 To 100%.

        End If

        bteBrightnessMax255 = lMax: Text1(2) = RGBToHSL201.Luminance '=Brighness both 0 To 255 and 0 To 100%.

    End If

End Function

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) Then 'Limiting the numerical textboxes (Text1[x]) to just register numerical enters.
        KeyAscii = 0

    End If

End Sub

Sub Text1_LostFocus(Index As Integer)

    Dim udtAngelSaturationBrightness As HSL, lngColor As Long 'Has to take care of intSystemColorAngleMax1530 0 To 1529.

    'If mBlnBigBoxReady = False Then Exit Sub 'Even the computers own enters are giving undesired calls to this routine.
    mBlnBigBoxReady = False 'Gives me fresh coordinates, but only in the RBG-model at this stage.
    blnNotFirstTimeMarker = False '-"-

    'HAVE TO ADD THE FUNCTIONALITY: img.Pilars position is totally dependent of the actual mode.
    If Index = 0 Then 'The user adjusted Hue so RGB will be aproximately calculated.
        If Text1(0) > 360 Then MsgBox "An integer between 0 and 360 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(0) = 360 'Checking both the precense of decimals and numbers greater than 360.

        '    If Text1(0) < 0 Then MsgBox "An integer between 0 and 360 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(0) = 0 'Checking both the precense of decimals and numbers greater than 360.
        If Text1(0) <> Round(Text1(0)) Then MsgBox "An integer between 0 and 360 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(0) = Round(Text1(0))  'Checking both the precense of decimals and numbers greater than 360.
    
        lngColor = HSLToRGB(Text1(0) / 360 * 255 * 6, bteSaturationMax255, bteBrightnessMax255, True)

        'imgArrows.Top = 255 - (intSystemColorAngleMax1530 / 1530 * 255) + 28 'Animating imgArrows
    End If

    If Index = 1 Then 'The user adjusted Saturation so RGB will be aproximately calculated.
        If Text1(1) > 100 Then MsgBox "An number between 0 and 100 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(1) = 100 'Checking both the precense of decimals and numbers greater than 360.
        If Text1(1) < 0 Then MsgBox "An number between 0 and 100 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(1) = 0 'Checking both the precense of decimals and numbers greater than 360.

        lngColor = HSLToRGB(intSystemColorAngleMax1530, Text1(1) / 100 * 255, bteBrightnessMax255, True)

    End If

    If Index = 2 Then 'The user adjusted Luminance so RGB will be aproximately calculated.
        If Text1(2) > 100 Then MsgBox "An number between 0 and 100 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(2) = 100 'Checking both the precense of decimals and numbers greater than 360.
        If Text1(2) < 0 Then MsgBox "An number between 0 and 100 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(2) = 0 'Checking both the precense of decimals and numbers greater than 360.

        lngColor = HSLToRGB(intSystemColorAngleMax1530, bteSaturationMax255, Text1(2) / 100 * 255, True)

    End If

    'ByVal RGBValue As Long, ByVal blnUpdateTextBoxes As Boolean
    If Index > 2 Then 'The user adjusted RGB so HSL is to calculated aproximately.
        'lngColor = RGB(Text1(3), Text1(4), Text1(5))
        udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True)

    End If

    'Justera imgArrows beroende på modus.
    'If objOption(0) Then imgArrows.Top = 255 - (intSystemColorAngleMax1530 / 1530 * 255) + 28 'Flyttar imgArrows
    'If objOption(1) Then imgArrows.Top = 255 - (Text1(1) * 2.55) + 28  'Flyttar imgArrows
    'If objOption(2) Then imgArrows.Top = 255 - (Text1(2) * 2.55) + 28  'Flyttar imgArrows
    Call imgArrowsModeDepending

    Call picBigBox_Colorize   '(mSngRValue, mSngGValue, mSngBValue)'Redrawing BigBox

End Sub

Private Sub txtHexColor_KeyPress(KeyAscii As Integer) 'Limits the textbox to numerics and A-F and capitals and to six pieces of letters.

    'The limitation to six letters probably has to be done in the vb editor poreperties window (see Greg Perry VB in 6 days).
    'txtHexColor.Refresh 'Otherwise the limitation of length wont work because i doesnt undestand its current length.'
    'If Len(txtHexColor) > 6 Then Exit Sub 'Max sex letters in the box.
    If (KeyAscii > 64 And KeyAscii < 71) Then Exit Sub 'A-F are OK.
    If (KeyAscii > 96 And KeyAscii < 103) Then KeyAscii = KeyAscii - 32: Exit Sub 'a-f becomes A-F. OK.
    If (KeyAscii > 47 And KeyAscii < 58) Then Exit Sub 'Numerics are OK.

    KeyAscii = 0 'All other letters are unwanted.

End Sub

Private Sub txtHexColor_LostFocus()

    'Dim udtAngelSaturationBrightness As HSL, lngColor As Long 'Must take care of intSystemColorAngleMax1530 0 To 1529.
    'On Error GoTo Bajs
    Dim sShift As String 'OBS! Must shift RGB into BGR to fit vb-standard.

    sShift = txtHexColor: sShift = mid(sShift, 5) & mid(sShift, 3, 2) & mid(sShift, 1, 2) 'Shifting RGB to BGR.
    lblNewColor.BackColor = ("&H" + sShift) 'OBS! Must shift RGB into BGR to fit vb-standard.
    Call SplitlblNewColorToRGBboxes 'Automatic update of the RGB textboxes.
    Call Text1_LostFocus(3) 'Simulating that the user adjusted the RGBtxtboxes->Total update. 3 means that the RedTextbox has been adjusted.
    Exit Sub
Bajs:
    MsgBox "An error occured while translating hexnumber to decimal number!"

End Sub

Public Function HSLToRGB(ByVal intLocalColorAngle As Integer, _
                         ByVal Saturation As Long, _
                         ByVal Luminance As Long, _
                         ByVal blnUpdateTextBoxes As Boolean) As Long

    Dim R As Long, G As Long, B As Long, lMax As Byte, lMid As Byte, lMin As Long, q As Single

    lMax = Luminance
    lMin = (255 - Saturation) * lMax / 255 '255 - (Saturation * lMax / 255)
    q = (lMax - lMin) / 255

    Select Case intLocalColorAngle

        Case 0 To 255
            lMid = (intLocalColorAngle - 0) * q + lMin
            R = lMax: G = lMid: B = lMin

        Case 256 To 510 'This period surpasses the node border with one unit - over to gren color. CHECK by F8.
            lMid = -(intLocalColorAngle - 255) * q + lMax '-(intLocalColorAngle - 256) * q + lMin
            R = lMid: G = lMax: B = lMin

        Case 511 To 765
            lMid = (intLocalColorAngle - 510) * q + lMin
            R = lMin: G = lMax: B = lMid

        Case 766 To 1020
            lMid = -(intLocalColorAngle - 765) * q + lMax
            R = lMin: G = lMid: B = lMax

        Case 1021 To 1275
            lMid = (intLocalColorAngle - 1020) * q + lMin
            R = lMid: G = lMin: B = lMax

        Case 1276 To 1530
            lMid = -(intLocalColorAngle - 1275) * q + lMax
            R = lMax: G = lMin: B = lMid

        Case Else
            MsgBox "Error occured in HSLToRGB. intSystemColorAngleMax1530= " & str(intLocalColorAngle)

    End Select

    mSngRValue = R: mSngGValue = G: mSngBValue = B 'Updating the sustem constants automatically. Perhaps must exclude this to give them protection.
    HSLToRGB = B * &H10000 + G * &H100& + R 'Delivers lngColor in VB-format.

    If blnUpdateTextBoxes = True Then 'Then the calling routine is not any of the complex automatic routines for fading etc.
        'Since this is a single time called session I can safely update my system constants and convert my hifgh resolution system constants to textbox dito.
        Text1(0) = Round(intLocalColorAngle / 255 / 6 * 360)
        Text1(1) = Round(Saturation / 255 * 100)
        Text1(2) = Round(Luminance / 255 * 100)
        Text1(3) = mSngRValue
        Text1(4) = mSngGValue
        Text1(5) = mSngBValue
        Text1(0).Refresh
        Text1(1).Refresh
        Text1(2).Refresh
        Text1(3).Refresh
        Text1(4).Refresh
        Text1(5).Refresh

        'txtHexColor = Hex$(mSngRValue * 65536 + mSngGValue * 256 + mSngBValue): txtHexColor.Refresh 'Applies to internetstandard<>VBstandard
        If mSngRValue < &H10 Then
            txtHexColor = Right$("00000" & Hex$(mSngRValue * 65536 + mSngGValue * 256 + mSngBValue), 6) 'Padding with zeroletters to the left.
        Else
            txtHexColor = Hex$(mSngRValue * 65536 + mSngGValue * 256 + mSngBValue)

        End If

        txtHexColor.Refresh 'End of the Hexabox routine.
        lblNewColor.BackColor = HSLToRGB
        lblNewColor.Refresh
        intSystemColorAngleMax1530 = intLocalColorAngle 'Sometims there is only a mouse Y coordinate tha is delivered from the calling routinen.
        bteSaturationMax255 = Saturation
        bteBrightnessMax255 = Luminance

    End If

End Function

Private Sub lblNewColor_Click()

    Dim udtAngelSaturationBrightness As HSL ', lngColor As Long

    'MsgBox "Klick!"
    lblOldColor.BackColor = lblNewColor.BackColor
    mBlnBigBoxReady = False 'Delivers fresh coordinates, but only in the HSL-model at this stage.
    blnNotFirstTimeMarker = False '-"-
    'Call Form_Load
    'Call SplitlblNewColorToRGBboxes 'Also the system constants RGB are updated.
    udtAngelSaturationBrightness = RGBToHSL201(lblNewColor.BackColor, True) 'True means that HSL is updating both the textboxes and the system constants.
    
    Call TriangelMove(255 - (intSystemColorAngleMax1530 / 1530 * 255) + 28) 'Animates the triangeln.
    Call picBigBox_Colorize   'Redraw BigBox

End Sub

Private Sub lblOldColor_Click()

    Dim udtAngelSaturationBrightness As HSL ', lngColor As Long

    'MsgBox "Klick!"
    lblNewColor.BackColor = lblOldColor.BackColor
    mBlnBigBoxReady = False 'Delivers fresh coordinates, but only in the HSL-model at this stage.
    blnNotFirstTimeMarker = False '-"-
    'Call Form_Load
    'Call SplitlblNewColorToRGBboxes 'Also the system constants RGB are updated.
    udtAngelSaturationBrightness = RGBToHSL201(lblNewColor.BackColor, True) 'True means that HSL is updating both the textboxes and the system constants.
    'imgArrows.Top = 255 - (intSystemColorAngleMax1530 / 1530 * 255) + 28 'Animates imgArrows
    Call TriangelMove(255 - (intSystemColorAngleMax1530 / 1530 * 255) + 28) 'Animates the triangel.
    Call picBigBox_Colorize   'Rita om BigBox

End Sub

Private Sub imgArrowsModeDepending()

    'AdjustingJusterar imgArrows depending on current mode.
    If objOption(0) Then Call TriangelMove(255 - (intSystemColorAngleMax1530 / 1530 * 255)) 'Animating the triangel.
    If objOption(1) Then Call TriangelMove(255 - (Text1(1) * 2.55))  'Animating the triangel.
    If objOption(2) Then Call TriangelMove(255 - (Text1(2) * 2.55))  'Animating the triangel.
    If objOption(3) Then Call TriangelMove(255 - Text1(3))   'Animating the triangel.

End Sub

Private Sub lblComplementaryColor_Click(Index As Integer)

    If Text1(0) < 180 Then Text1(0) = Text1(0) + 180 Else Text1(0) = Text1(0) - 180
    Call Text1_LostFocus(0) 'Noll stands for Hue.

End Sub

Private Sub PaintThinBox(Index As Integer)

    Dim blnFadeToGrey As Boolean, blnFadeToBlack As Boolean

    If Index = 0 Then 'MsgBox "Hue"
        'Set objAnyPictureBox = picThinBox
        'Call RainBowSurface(objAnyPictureBox, 19, blnFadeToGrey, blnFadeToBlack)
        Call RainBowThinBox

    End If

    If Index = 1 Then
        'MsgBox "Saturation" 'ColorVinkeln ligger nu vågrätt från vänster till höger. Textrutorna är viktiga.
        'Set objAnyPictureBox = picThinBox
        'picThinBox.Cls
        'Call FadeToGrey(objAnyPictureBox, 19, 255)
        Call FadeThinBoxToGrey
        picThinBox.Refresh

        'Set objAnyPictureBox = Nothing 'Kanske sparar minne.
    End If

    If Index = 2 Then ' "Brightness"
        'MsgBox "Saturation" 'Color angel is now horisontal from left to right. Textboxes are important.
        'Set objAnyPictureBox = picThinBox
        'Crucial to give the thin box maximunm lightness as a starting point for the lightness fade.
        picThinBox.BackColor = HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal bteSaturationMax255, ByVal 255, False) 'Delivers a lighter shade of the active color. 'Setting the whole square for easy fading.
        'Call FadeToBlack(objAnyPictureBox, 19, 255)
        Call FadeThinBoxToBlack
        picThinBox.Refresh

    End If

    picThinBox.Visible = True
    'Set objAnyPictureBox = Nothing 'Perhaps saves some memory.

End Sub

Private Sub TriangelMove(Y)
    linTriang1Vert.y1 = Y + 28: linTriang1Vert.y2 = Y + 28 + 10
    linTriang1Rising.y1 = Y + 28 + 10: linTriang1Rising.y2 = Y + 28 + 4
    linTriang1Falling.y1 = Y + 28: linTriang1Falling.y2 = Y + 28 + 6

    linTriang2Vert.y1 = Y + 28: linTriang2Vert.y2 = Y + 28 + 10
    linTriang2Rising.y2 = Y + 28 + 10: linTriang2Rising.y1 = Y + 28 + 5
    linTriang2Falling.y2 = Y + 28: linTriang2Falling.y1 = Y + 28 + 5

End Sub

Public Sub opt3RedPaintPicThinBox(ByVal G, B)

    Dim bteX As Byte, intCtr As Integer

    'Painting picThinBox (19, 255)
    For bteX = 0 To 19
        For intCtr = 0 To 255
            SetPixelV picThinBox.hDC, bteX, intCtr, RGB(255 - intCtr, G, B) 'Painting with API.
        Next intCtr
    Next bteX

End Sub

Public Sub opt4GreenPaintPicThinBox(ByVal R, B)

    Dim bteX As Byte, intCtr As Integer

    'Painting picThinBox (19, 255)
    For bteX = 0 To 19
        For intCtr = 0 To 255
            SetPixelV picThinBox.hDC, bteX, intCtr, RGB(R, 255 - intCtr, B) 'Painting by API.
        Next intCtr
    Next bteX

End Sub

Public Sub opt5BluePaintPicThinBox(ByVal R, G)

    Dim bteX As Byte, intCtr As Integer

    'Painting picThinBox (19, 255)
    For bteX = 0 To 19
        For intCtr = 0 To 255
            SetPixelV picThinBox.hDC, bteX, intCtr, RGB(R, G, 255 - intCtr) 'Painting by API.
        Next intCtr
    Next bteX

End Sub

Public Sub BigBoxOpt3Reaction(ByVal X, Y)

    Dim udtAngelSaturationBrightness As HSL

    lblNewColor.BackColor = picBigBox.Point(X, Y): lblNewColor.Refresh
    Call SplitlblNewColorToRGBboxes 'Updating the module global mSngRValue etc.
    Call opt3RedPaintPicThinBox(ByVal mSngGValue, mSngBValue)
    picThinBox.Refresh
    udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True) 'Causes an update of all the constants.

End Sub

Public Sub BigBoxOpt4Reaction(ByVal X, Y)

    Dim udtAngelSaturationBrightness As HSL

    lblNewColor.BackColor = picBigBox.Point(X, Y): lblNewColor.Refresh
    Call SplitlblNewColorToRGBboxes 'Updating the module global mSngRValue etc.
    Call opt4GreenPaintPicThinBox(ByVal mSngRValue, mSngBValue)
    picThinBox.Refresh
    udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True) 'Causes an update of all the constants. The letter of three stands for the RED-box.

End Sub

Public Sub BigBoxOpt5Reaction(ByVal X, Y)

    Dim udtAngelSaturationBrightness As HSL

    lblNewColor.BackColor = picBigBox.Point(X, Y): lblNewColor.Refresh
    Call SplitlblNewColorToRGBboxes 'Updating the module global mSngRValue etc.
    Call opt5BluePaintPicThinBox(ByVal mSngRValue, mSngGValue)
    picThinBox.Refresh
    udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True) 'Causes an update of all the constants. The letter of three stands for the RED-box..

End Sub

Private Sub opt3RedPaintPicBigBox()

    Dim R As Single, G As Single, B As Single

    'Paint the picBigBox
    R = Text1(3) 'Red

    For B = 255 To 0 Step -1
        For G = 255 To 0 Step -1 'Interesting if there is an error, thus a jump directly to EndSub.
            SetPixelV picBigBox.hDC, B, 255 - G, RGB(R, G, B) 'Painting by API.
        Next G

        'G = G - 1 'Because that G becomes too big when the loop has finishes.
    Next B

End Sub

Private Sub opt4GreenPaintPicBigBox()

    Dim R As Single, G As Single, B As Single

    'Paint picBigBox
    G = Text1(4) 'Green

    For B = 255 To 0 Step -1
        For R = 255 To 0 Step -1 'Interesting if there is an error, thus a jump directly to EndSub.
            SetPixelV picBigBox.hDC, B, 255 - R, RGB(R, G, B) 'Painting by API.
        Next R

        R = R - 1 'Because that R becomes too big when the loop has finishes.
    Next B

End Sub

Private Sub opt5BluePaintPicBigBox()

    Dim R As Single, G As Single, B As Single

    'Paint picBigBox
    B = Text1(5) 'Blue

    For R = 255 To 0 Step -1
        For G = 255 To 0 Step -1 'Interesting if there is an error, thus a jump directly to EndSub.
            SetPixelV picBigBox.hDC, R, 255 - G, RGB(R, G, B) 'Ritar medelst API.
        Next G

        G = G - 1 'Because that G becomes too big when the loop has finishes..
    Next R

End Sub

Private Sub MoveHexBox()

    Dim Ctr As Integer

    'MsgBox "Move HexBox"
    For Ctr = 336 To 286 Step -1
        txtHexColor.Move Ctr, 281, 56, 20
    Next Ctr

End Sub
