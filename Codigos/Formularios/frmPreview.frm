VERSION 5.00
Begin VB.Form frmPreview 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preview"
   ClientHeight    =   3540
   ClientLeft      =   9420
   ClientTop       =   12540
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PreviewGrh 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   0
      ScaleHeight     =   3510
      ScaleWidth      =   4425
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub fPreviewGrh(ByVal GrhIn As Long)
'*************************************************
'Author: Unkwown
'Last modified: 22/05/06
'*************************************************

    If Val(GrhIn) < 1 Then
      frmSuperficies.cGrh.Text = grhCount
      Exit Sub
    End If
    
    If Val(GrhIn) > grhCount Then
      frmSuperficies.cGrh.Text = 1
      Exit Sub
    End If
    
    'Change CurrentGrh
    CurrentGrh.GrhIndex = GrhIn
    CurrentGrh.Started = 1
    CurrentGrh.FrameCounter = 1
    CurrentGrh.speed = GrhData(CurrentGrh.GrhIndex).speed

End Sub
