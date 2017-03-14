VERSION 5.00
Begin VB.Form frmWire3Dxyz 
   Caption         =   "Wireframe 3D xyz"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   677
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdInitCube 
      Caption         =   "Initialize Cube"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox pctWire3D 
      Height          =   7695
      Left            =   600
      ScaleHeight     =   509
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   653
      TabIndex        =   0
      Top             =   600
      Width           =   9855
   End
End
Attribute VB_Name = "frmWire3Dxyz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
