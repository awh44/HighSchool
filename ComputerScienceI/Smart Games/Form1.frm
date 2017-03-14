VERSION 5.00
Begin VB.Form frmTicTacToe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic-Tac-Toe"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraFirst 
      Caption         =   "First Move"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6480
      TabIndex        =   15
      Top             =   960
      Width           =   2055
      Begin VB.OptionButton optComputer 
         Caption         =   "Computer"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optPlayer 
         Caption         =   "Player"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Game"
      Height          =   735
      Left            =   6480
      TabIndex        =   14
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Game"
      Height          =   735
      Left            =   6480
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdComputer 
      Caption         =   "Computer"
      Height          =   735
      Left            =   6480
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Line Diagonal2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   360
      X2              =   6000
      Y1              =   6000
      Y2              =   360
   End
   Begin VB.Line Diagonal1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   240
      X2              =   6000
      Y1              =   240
      Y2              =   6120
   End
   Begin VB.Line Column3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   5280
      X2              =   5280
      Y1              =   120
      Y2              =   6240
   End
   Begin VB.Line Column2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   3120
      X2              =   3120
      Y1              =   120
      Y2              =   6120
   End
   Begin VB.Line Column1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   960
      X2              =   960
      Y1              =   120
      Y2              =   6120
   End
   Begin VB.Line Row3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   120
      X2              =   6240
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Row2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   120
      X2              =   6240
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Row1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   120
      X2              =   6240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   12
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblMove 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   6480
      TabIndex        =   10
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Computer's move:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Line Line4 
      X1              =   -120
      X2              =   6360
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line3 
      X1              =   -120
      X2              =   6360
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      X1              =   4200
      X2              =   4200
      Y1              =   6600
      Y2              =   -240
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2040
      Y1              =   6600
      Y2              =   -240
   End
   Begin VB.Label lblBoard 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   8
      Left            =   4320
      TabIndex        =   8
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label lblBoard 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   7
      Left            =   2160
      TabIndex        =   7
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label lblBoard 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label lblBoard 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   5
      Left            =   4320
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblBoard 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   4
      Left            =   2160
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblBoard 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblBoard 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   2
      Left            =   4320
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblBoard 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblBoard 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmTicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Austin Herring
'Start date: December 14, 2009
'Completion date: December 22, 2009
'Smart Games Project
'frmTicTacToe
'Description: a simulation of a smart tic-tac-toe game.
    'Computer goes first.
    'We will program a smart offensive strategy.
    
Dim n As Integer     'store which move computer is on (1, 2, 3, 4, or 5)
Dim Game_over As Boolean ' 2-valued (true/false) variable
Dim Move_made As Boolean
Dim XO As Boolean
Dim position As Integer




Private Sub Check_Block()
If lblBoard(0) = "O" And lblBoard(1) = "O" And lblBoard(2) = "" And Move_made = False Then
    lblBoard(2) = "X"
    Move_made = True
End If
If lblBoard(0) = "O" And lblBoard(2) = "O" And lblBoard(1) = "" And Move_made = False Then
    lblBoard(1) = "X"
    Move_made = True
End If
If lblBoard(1) = "O" And lblBoard(2) = "O" And lblBoard(0) = "" And Move_made = False Then
    lblBoard(0) = "X"
    Move_made = True
End If

If lblBoard(3) = "O" And lblBoard(4) = "O" And lblBoard(5) = "" And Move_made = False Then
    lblBoard(5) = "X"
    Move_made = True
End If
If lblBoard(3) = "O" And lblBoard(5) = "O" And lblBoard(4) = "" And Move_made = False Then
    lblBoard(4) = "X"
    Move_made = True
End If
If lblBoard(4) = "O" And lblBoard(5) = "O" And lblBoard(3) = "" And Move_made = False Then
    lblBoard(3) = "X"
    Move_made = True
End If

If lblBoard(6) = "O" And lblBoard(7) = "O" And lblBoard(8) = "" And Move_made = False Then
    lblBoard(8) = "X"
    Move_made = True
End If
If lblBoard(6) = "O" And lblBoard(8) = "O" And lblBoard(7) = "" And Move_made = False Then
    lblBoard(7) = "X"
    Move_made = True
End If
If lblBoard(7) = "O" And lblBoard(8) = "O" And lblBoard(6) = "" And Move_made = False Then
    lblBoard(6) = "X"
    Move_made = True
End If

If lblBoard(0) = "O" And lblBoard(6) = "O" And lblBoard(3) = "" And Move_made = False Then
    lblBoard(3) = "X"
    Move_made = True
End If
If lblBoard(1) = "O" And lblBoard(7) = "O" And lblBoard(4) = "" And Move_made = False Then
    lblBoard(4) = "X"
    Move_made = True
End If
If lblBoard(2) = "O" And lblBoard(8) = "O" And lblBoard(5) = "" And Move_made = False Then
    lblBoard(5) = "X"
    Move_made = True
End If

If lblBoard(0) = "O" And lblBoard(3) = "O" And lblBoard(6) = "" And Move_made = False Then
    lblBoard(6) = "X"
    Move_made = True
End If
If lblBoard(1) = "O" And lblBoard(4) = "O" And lblBoard(7) = "" And Move_made = False Then
    lblBoard(7) = "X"
    Move_made = True
End If
If lblBoard(2) = "O" And lblBoard(5) = "O" And lblBoard(8) = "" And Move_made = False Then
    lblBoard(8) = "X"
    Move_made = True
End If

If lblBoard(3) = "O" And lblBoard(6) = "O" And lblBoard(0) = "" And Move_made = False Then
    lblBoard(0) = "X"
    Move_made = True
End If
If lblBoard(4) = "O" And lblBoard(7) = "O" And lblBoard(1) = "" And Move_made = False Then
    lblBoard(1) = "X"
    Move_made = True
End If
If lblBoard(5) = "O" And lblBoard(8) = "O" And lblBoard(2) = "" And Move_made = False Then
    lblBoard(2) = "X"
    Move_made = True
End If

If lblBoard(0) = "O" And lblBoard(8) = "O" And lblBoard(4) = "" And Move_made = False Then
    lblBoard(4) = "X"
    Move_made = True
End If
If lblBoard(0) = "O" And lblBoard(4) = "O" And lblBoard(8) = "" And Move_made = False Then
    lblBoard(8) = "X"
    Move_made = True
End If
If lblBoard(4) = "O" And lblBoard(8) = "O" And lblBoard(0) = "" And Move_made = False Then
    lblBoard(0) = "X"
    Move_made = True
End If

If lblBoard(2) = "O" And lblBoard(6) = "O" And lblBoard(4) = "" And Move_made = False Then
    lblBoard(4) = "X"
    Move_made = True
End If
If lblBoard(2) = "O" And lblBoard(4) = "O" And lblBoard(6) = "" And Move_made = False Then
    lblBoard(6) = "X"
    Move_made = True
End If
If lblBoard(4) = "O" And lblBoard(6) = "O" And lblBoard(2) = "" And Move_made = False Then
    lblBoard(2) = "X"
    Move_made = True
End If

End Sub

Private Sub Check_Winner()
'check row 1, (index 0, 1, 2)
If lblBoard(0) = "X" And lblBoard(1) = "X" And lblBoard(2) = "" And Game_over = False Then
    lblBoard(2) = "X"
    Game_over = True
    lblResult = "Lose"
    Row1.Visible = True
End If
If lblBoard(0) = "X" And lblBoard(2) = "X" And lblBoard(1) = "" And Game_over = False Then
    lblBoard(1) = "X"
    Game_over = True
    lblResult = "Lose"
    Row1.Visible = True
End If
If lblBoard(1) = "X" And lblBoard(2) = "X" And lblBoard(0) = "" And Game_over = False Then
    lblBoard(0) = "X"
    Game_over = True
    lblResult = "Lose"
    Row1.Visible = True
End If

If lblBoard(3) = "X" And lblBoard(4) = "X" And lblBoard(5) = "" And Game_over = False Then
    lblBoard(5) = "X"
    Game_over = True
    lblResult = "Lose"
    Row2.Visible = True
End If
If lblBoard(3) = "X" And lblBoard(5) = "X" And lblBoard(4) = "" And Game_over = False Then
    lblBoard(4) = "X"
    Game_over = True
    lblResult = "Lose"
    Row2.Visible = True
End If
If lblBoard(4) = "X" And lblBoard(5) = "X" And lblBoard(3) = "" And Game_over = False Then
    lblBoard(3) = "X"
    Game_over = True
    lblResult = "Lose"
    Row2.Visible = True
End If

If lblBoard(6) = "X" And lblBoard(7) = "X" And lblBoard(8) = "" And Game_over = False Then
    lblBoard(8) = "X"
    Game_over = True
    lblResult = "Lose"
    Row3.Visible = True
End If
If lblBoard(6) = "X" And lblBoard(8) = "X" And lblBoard(7) = "" And Game_over = False Then
    lblBoard(7) = "X"
    Game_over = True
    lblResult = "Lose"
    Row3.Visible = True
End If
If lblBoard(7) = "X" And lblBoard(8) = "X" And lblBoard(6) = "" And Game_over = False Then
    lblBoard(6) = "X"
    Game_over = True
    lblResult = "Lose"
    Row3.Visible = True
End If

If lblBoard(0) = "X" And lblBoard(6) = "X" And lblBoard(3) = "" And Game_over = False Then
    lblBoard(3) = "X"
    Game_over = True
    lblResult = "Lose"
    Column1.Visible = True
End If
If lblBoard(1) = "X" And lblBoard(7) = "X" And lblBoard(4) = "" And Game_over = False Then
    lblBoard(4) = "X"
    Game_over = True
    lblResult = "Lose"
    Column2.Visible = True
End If
If lblBoard(2) = "X" And lblBoard(8) = "X" And lblBoard(5) = "" And Game_over = False Then
    lblBoard(5) = "X"
    Game_over = True
    lblResult = "Lose"
    Column3.Visible = True
End If

If lblBoard(0) = "X" And lblBoard(3) = "X" And lblBoard(6) = "" And Game_over = False Then
    lblBoard(6) = "X"
    Game_over = True
    lblResult = "Lose"
    Column1.Visible = True
End If
If lblBoard(1) = "X" And lblBoard(4) = "X" And lblBoard(7) = "" And Game_over = False Then
    lblBoard(7) = "X"
    Game_over = True
    lblResult = "Lose"
    Column2.Visible = True
End If
If lblBoard(2) = "X" And lblBoard(5) = "X" And lblBoard(8) = "" And Game_over = False Then
    lblBoard(8) = "X"
    Game_over = True
    lblResult = "Lose"
    Column3.Visible = True
End If

If lblBoard(3) = "X" And lblBoard(6) = "X" And lblBoard(0) = "" And Game_over = False Then
    lblBoard(0) = "X"
    Game_over = True
    lblResult = "Lose"
    Column1.Visible = True
End If
If lblBoard(4) = "X" And lblBoard(7) = "X" And lblBoard(1) = "" And Game_over = False Then
    lblBoard(1) = "X"
    Game_over = True
    lblResult = "Lose"
    Column2.Visible = True
End If
If lblBoard(5) = "X" And lblBoard(8) = "X" And lblBoard(2) = "" And Game_over = False Then
    lblBoard(2) = "X"
    Game_over = True
    lblResult = "Lose"
    Column3.Visible = True
End If

If lblBoard(0) = "X" And lblBoard(8) = "X" And lblBoard(4) = "" And Game_over = False Then
    lblBoard(4) = "X"
    Game_over = True
    lblResult = "Lose"
    Diagonal1.Visible = True
End If
If lblBoard(0) = "X" And lblBoard(4) = "X" And lblBoard(8) = "" And Game_over = False Then
    lblBoard(8) = "X"
    Game_over = True
    lblResult = "Lose"
    Diagonal1.Visible = True
End If
If lblBoard(4) = "X" And lblBoard(8) = "X" And lblBoard(0) = "" And Game_over = False Then
    lblBoard(0) = "X"
    Game_over = True
    lblResult = "Lose"
    Diagonal1.Visible = True
End If

If lblBoard(2) = "X" And lblBoard(6) = "X" And lblBoard(4) = "" And Game_over = False Then
    lblBoard(4) = "X"
    Game_over = True
    lblResult = "Lose"
    Diagonal2.Visible = True
End If
If lblBoard(2) = "X" And lblBoard(4) = "X" And lblBoard(6) = "" And Game_over = False Then
    lblBoard(6) = "X"
    Game_over = True
    lblResult = "Lose"
    Diagonal2.Visible = True
End If
If lblBoard(4) = "X" And lblBoard(6) = "X" And lblBoard(2) = "" And Game_over = False Then
    lblBoard(2) = "X"
    Game_over = True
    lblResult = "Lose"
    Diagonal2.Visible = True
End If



End Sub

Private Sub cmdComputer_Click()
Randomize



n = Val(lblMove)

If n = 1 Then 'comp's first move
    Do
        position = Int(9 * Rnd)              '0-8
    Loop Until position = 0 Or position = 2 Or position = 6 Or position = 8
    lblBoard(position) = "X"
End If

If n = 2 Then
    Player_winner
    If Game_over = False Then Check_Winner
    If Game_over = False Then Check_Block
    If Game_over = False And Move_made = False Then
        Do
            XO = False
            position = Int(9 * Rnd)
            check_XO
        Loop Until (position = 0 Or position = 2 Or position = 6 Or position = 8) And Len(lblBoard(position)) = 0 And XO = False
        lblBoard(position) = "X"
    End If
End If

If n = 3 Then
    Player_winner
    If Game_over = False Then Check_Winner
    If Game_over = False Then Check_Block
    If Game_over = False And Move_made = False Then        ' if game isn't over, make next corner move
        Do
            XO = False
            position = Int(9 * Rnd)
            check_XO
        Loop Until (position = 0 Or position = 2 Or position = 6 Or position = 8) And Len(lblBoard(position)) = 0 And XO = False
        lblBoard(position) = "X"
    End If
End If

If n = 4 Then
    Player_winner
    If Game_over = False Then Check_Winner
    If Game_over = False Then Check_Block
    If Game_over = False And Move_made = False Then      ' if game isn't over, make next corner move
        Do
            position = Int(9 * Rnd)
        Loop Until (position = 0 Or position = 2 Or position = 6 Or position = 8) And Len(lblBoard(position)) = 0
        lblBoard(position) = "X"
    End If
End If

If n = 5 Then
    Player_winner
    If Game_over = False Then Check_Winner
    If Game_over = False Then
        For c = 0 To 8
            If lblBoard(c) = "" Then
                lblBoard(c) = "X"
                lblResult = "Tie"
            End If
        Next c
    End If
End If

If n <> 5 Then lblMove = Val(lblMove) + 1

Move_made = False
Game_over = False



End Sub

Private Sub cmdNew_Click()
For c = 0 To 8
    lblBoard(c) = ""
Next c
cmdStart_Click
lblResult = ""
Game_over = False
Move_made = False
XO = False
Row1.Visible = False
Row2.Visible = False
Row3.Visible = False
Column1.Visible = False
Column2.Visible = False
Column3.Visible = False
Diagonal1.Visible = False
Diagonal2.Visible = False

End Sub

Private Sub cmdStart_Click()

If optComputer = True Then
    Do
        position = Int(9 * Rnd)              '0-8
    Loop Until position = 0 Or position = 2 Or position = 6 Or position = 8
    lblBoard(position) = "X"
    lblMove = 2
End If

If optPlayer = True Then
    lblMove = 1
End If



cmdStart.Enabled = False
End Sub

Private Sub Form_Load()
optComputer = True
End Sub

Private Sub lblBoard_Click(Index As Integer)

If optPlayer = True And cmdStart.Enabled = False And Len(lblBoard(Index)) = 0 Then
    lblBoard(Index) = "O"
    If lblMove = 5 Then lblResult = "Tie"
    Computer_second
End If

If lblMove > 1 And optComputer = True And Len(lblBoard(Index)) = 0 Then
    If Len(lblBoard(Index)) = 0 Then lblBoard(Index) = "O"
    cmdComputer_Click
    Player_winner
End If
End Sub

Private Sub Computer_second()

Randomize

If lblMove = 1 Then
    If lblBoard(4) = "O" Then
        Do
            position = Int(9 * Rnd)
        Loop Until (position = 0 Or position = 2 Or position = 6 Or position = 8) And Len(lblBoard(position)) = 0
        lblBoard(position) = "X"
    Else
        lblBoard(4) = "X"
    End If
End If

If lblMove = 2 Then
    Player_winner
    If Game_over = False Then Check_Winner
    If Game_over = False Then Check_Block
    If Game_over = False And Move_made = False Then
        If (lblBoard(0) = "O" And lblBoard(8) = "O") Or (lblBoard(2) = "O" And lblBoard(6) = "O") Then
            Do
                position = Int(9 * Rnd)
            Loop Until (position = 1 Or position = 3 Or position = 5 Or position = 7) And Len(lblBoard(position)) = 0
            lblBoard(position) = "X"
        Else
            Do
                position = Int(9 * Rnd)
            Loop Until (position = 0 Or position = 2 Or position = 6 Or position = 8) And Len(lblBoard(position)) = 0
            lblBoard(position) = "X"
        End If
    End If
End If

If lblMove = 3 Then
    Player_winner
    If Game_over = False Then Check_Winner
    If Game_over = False Then Check_Block

    
    If Game_over = False And Move_made = False Then
        If (lblBoard(0) = "O" And lblBoard(8) = "O") Or (lblBoard(2) = "O" And lblBoard(6) = "O") Then
            Do
                position = Int(9 * Rnd)
            Loop Until (position = 1 Or position = 3 Or position = 5 Or position = 7) And Len(lblBoard(position)) = 0
            lblBoard(position) = "X"
        Else
            Do
                position = Int(9 * Rnd)
            Loop Until (position = 0 Or position = 2 Or position = 6 Or position = 8) And Len(lblBoard(position)) = 0
            lblBoard(position) = "X"
        End If
    End If
    
End If

If lblMove = 4 Then
    Player_winner
    If Game_over = False Then Check_Winner
    If Game_over = False Then Check_Block
    If Game_over = False And Move_made = False Then
        Do
            position = Int(9 * Rnd)
        Loop Until (position = 0 Or position = 2 Or position = 6 Or position = 8) And Len(lblBoard(position)) = 0
        lblBoard(position) = "X"
    End If
End If

    
If lblMove <> 5 Then lblMove = Val(lblMove) + 1
Move_made = False


End Sub

Private Sub Player_winner()         'at one point, there was a way for the player to win....so this isn't completely useless
If lblBoard(0) = "O" And lblBoard(1) = "O" And lblBoard(2) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(0) = "O" And lblBoard(2) = "O" And lblBoard(1) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(1) = "O" And lblBoard(2) = "O" And lblBoard(0) = "O" Then
    lblResult = "Win"
    Game_over = True
End If

If lblBoard(3) = "O" And lblBoard(4) = "O" And lblBoard(5) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(3) = "O" And lblBoard(5) = "O" And lblBoard(4) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(4) = "O" And lblBoard(5) = "O" And lblBoard(3) = "O" Then
    lblResult = "Win"
    Game_over = True
End If

If lblBoard(6) = "O" And lblBoard(7) = "O" And lblBoard(8) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(6) = "O" And lblBoard(8) = "O" And lblBoard(7) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(7) = "O" And lblBoard(8) = "O" And lblBoard(6) = "O" Then
    lblResult = "Win"
    Game_over = True
End If

If lblBoard(0) = "O" And lblBoard(6) = "O" And lblBoard(3) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(1) = "O" And lblBoard(7) = "O" And lblBoard(4) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(2) = "O" And lblBoard(8) = "O" And lblBoard(5) = "O" Then
    lblResult = "Win"
    Game_over = True
End If

If lblBoard(0) = "O" And lblBoard(3) = "O" And lblBoard(6) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(1) = "O" And lblBoard(4) = "O" And lblBoard(7) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(2) = "O" And lblBoard(5) = "O" And lblBoard(8) = "O" Then
    lblResult = "Win"
    Game_over = True
End If

If lblBoard(3) = "O" And lblBoard(6) = "O" And lblBoard(0) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(4) = "O" And lblBoard(7) = "O" And lblBoard(1) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(5) = "O" And lblBoard(8) = "O" And lblBoard(2) = "O" Then
    lblResult = "Win"
    Game_over = True
End If

If lblBoard(0) = "O" And lblBoard(8) = "O" And lblBoard(4) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(0) = "O" And lblBoard(4) = "O" And lblBoard(8) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(4) = "O" And lblBoard(8) = "O" And lblBoard(0) = "O" Then
    lblResult = "Win"
    Game_over = True
End If

If lblBoard(2) = "O" And lblBoard(6) = "O" And lblBoard(4) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(2) = "O" And lblBoard(4) = "O" And lblBoard(6) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
If lblBoard(4) = "O" And lblBoard(6) = "O" And lblBoard(2) = "O" Then
    lblResult = "Win"
    Game_over = True
End If
End Sub

Private Sub check_XO()

If lblBoard(0) = "O" And lblBoard(1) = "X" And position = 2 Then
    XO = True
End If
If lblBoard(0) = "O" And lblBoard(2) = "X" And position = 1 Then
    XO = True
End If
If lblBoard(1) = "O" And lblBoard(2) = "X" And position = 0 Then
    XO = True
End If

If lblBoard(3) = "O" And lblBoard(4) = "X" And position = 5 Then
    XO = True
End If
If lblBoard(3) = "O" And lblBoard(5) = "X" And position = 4 Then
    XO = True
End If
If lblBoard(4) = "O" And lblBoard(5) = "X" And position = 3 Then
    XO = True
End If

If lblBoard(6) = "O" And lblBoard(7) = "X" And position = 8 Then
    XO = True
End If
If lblBoard(6) = "O" And lblBoard(8) = "X" And position = 7 Then
    XO = True
End If
If lblBoard(7) = "O" And lblBoard(8) = "X" And position = 6 Then
    XO = True
End If

If lblBoard(0) = "O" And lblBoard(6) = "X" And position = 3 Then
    XO = True
End If
If lblBoard(1) = "O" And lblBoard(7) = "X" And position = 4 Then
    XO = True
End If
If lblBoard(2) = "O" And lblBoard(8) = "X" And position = 4 Then
    XO = True
End If

If lblBoard(0) = "O" And lblBoard(3) = "X" And position = 6 Then
    XO = True
End If
If lblBoard(1) = "O" And lblBoard(4) = "X" And position = 7 Then
    XO = True
End If
If lblBoard(2) = "O" And lblBoard(5) = "X" And position = 8 Then
    XO = True
End If

If lblBoard(3) = "O" And lblBoard(6) = "X" And position = 0 Then
    XO = True
End If
If lblBoard(4) = "O" And lblBoard(7) = "X" And position = 1 Then
    XO = True
End If
If lblBoard(5) = "O" And lblBoard(8) = "X" And position = 2 Then
    XO = True
End If

If lblBoard(0) = "O" And lblBoard(8) = "X" And position = 4 Then
    XO = True
End If
If lblBoard(0) = "O" And lblBoard(4) = "X" And position = 8 Then
    XO = True
End If
If lblBoard(4) = "O" And lblBoard(8) = "X" And position = 0 Then
    XO = True
End If

If lblBoard(2) = "O" And lblBoard(6) = "X" And position = 4 Then
    XO = True
End If
If lblBoard(2) = "O" And lblBoard(4) = "X" And position = 6 Then
    XO = True
End If
If lblBoard(4) = "O" And lblBoard(6) = "X" And position = 2 Then
    XO = True
End If








If lblBoard(0) = "X" And lblBoard(1) = "O" And position = 2 Then
    XO = True
End If
If lblBoard(0) = "X" And lblBoard(2) = "O" And position = 1 Then
    XO = True
End If
If lblBoard(1) = "X" And lblBoard(2) = "O" And position = 0 Then
    XO = True
End If

If lblBoard(3) = "X" And lblBoard(4) = "O" And position = 5 Then
    XO = True
End If
If lblBoard(3) = "X" And lblBoard(5) = "O" And position = 4 Then
    XO = True
End If
If lblBoard(4) = "X" And lblBoard(5) = "O" And position = 3 Then
    XO = True
End If

If lblBoard(6) = "X" And lblBoard(7) = "O" And position = 8 Then
    XO = True
End If
If lblBoard(6) = "X" And lblBoard(8) = "O" And position = 7 Then
    XO = True
End If
If lblBoard(7) = "X" And lblBoard(8) = "O" And position = 6 Then
    XO = True
End If

If lblBoard(0) = "X" And lblBoard(6) = "O" And position = 3 Then
    XO = True
End If
If lblBoard(1) = "X" And lblBoard(7) = "O" And position = 4 Then
    XO = True
End If
If lblBoard(2) = "X" And lblBoard(8) = "O" And position = 4 Then
    XO = True
End If

If lblBoard(0) = "X" And lblBoard(3) = "O" And position = 6 Then
    XO = True
End If
If lblBoard(1) = "X" And lblBoard(4) = "O" And position = 7 Then
    XO = True
End If
If lblBoard(2) = "X" And lblBoard(5) = "O" And position = 8 Then
    XO = True
End If

If lblBoard(3) = "X" And lblBoard(6) = "O" And position = 0 Then
    XO = True
End If
If lblBoard(4) = "X" And lblBoard(7) = "O" And position = 1 Then
    XO = True
End If
If lblBoard(5) = "X" And lblBoard(8) = "O" And position = 2 Then
    XO = True
End If

If lblBoard(0) = "X" And lblBoard(8) = "O" And position = 4 Then
    XO = True
End If
If lblBoard(0) = "X" And lblBoard(4) = "O" And position = 8 Then
    XO = True
End If
If lblBoard(4) = "X" And lblBoard(8) = "O" And position = 0 Then
    XO = True
End If

If lblBoard(2) = "X" And lblBoard(6) = "O" And position = 4 Then
    XO = True
End If
If lblBoard(2) = "X" And lblBoard(4) = "O" And position = 6 Then
    XO = True
End If
If lblBoard(4) = "X" And lblBoard(6) = "O" And position = 2 Then
    XO = True
End If



End Sub


Private Sub Check_corners()

End Sub
