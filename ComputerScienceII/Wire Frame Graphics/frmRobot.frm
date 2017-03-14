VERSION 5.00
Begin VB.Form frmRobot 
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   607
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   706
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Height          =   375
      Left            =   9360
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   8895
      Left            =   0
      ScaleHeight     =   589
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   589
      TabIndex        =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "frmRobot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
