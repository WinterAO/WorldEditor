VERSION 5.00
Begin VB.Form frmMinimapa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Minimapa"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   1485
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
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   99
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox minimapa 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   0
      ScaleHeight     =   99
      ScaleMode       =   0  'User
      ScaleWidth      =   99
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1485
   End
End
Attribute VB_Name = "frmMinimapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
