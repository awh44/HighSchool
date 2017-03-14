VERSION 5.00
Begin VB.Form frmStartup 
   Caption         =   "Fractals"
   ClientHeight    =   4215
   ClientLeft      =   2115
   ClientTop       =   3420
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   2565
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton cmdBoxes 
         Caption         =   "Boxes"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton cmdSierpinski 
         Caption         =   "Sierpinski Triangle"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTree 
         Caption         =   "Trees"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdTriangles 
         Caption         =   "Diamond"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdSquares 
         Caption         =   "Squares"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   2400
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBoxes_Click()
frmBoxes.Show
End Sub

Private Sub cmdSierpinski_Click()
frmSierpinski.Show
End Sub

Private Sub cmdSquares_Click()
frmSquares.Show
End Sub

Private Sub cmdTree_Click()
frmTree.Show
End Sub

Private Sub cmdTriangles_Click()
frmTriangles.Show
End Sub
