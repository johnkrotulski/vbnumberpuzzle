VERSION 5.00
Begin VB.Form Puzzle 
   Caption         =   "Puzzle Game"
   ClientHeight    =   4485
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame background 
      BackColor       =   &H000080FF&
      Height          =   4092
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4452
      Begin VB.Frame frameEnd 
         Height          =   4092
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   4452
         Begin VB.Frame frameRecord 
            Caption         =   "Record Book"
            Height          =   2772
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   4212
            Begin VB.Label lblHistory 
               Height          =   252
               Index           =   9
               Left            =   120
               TabIndex        =   32
               Top             =   2400
               Width           =   3972
            End
            Begin VB.Label lblHistory 
               Height          =   252
               Index           =   8
               Left            =   120
               TabIndex        =   31
               Top             =   2160
               Width           =   3972
            End
            Begin VB.Label lblHistory 
               Height          =   252
               Index           =   7
               Left            =   120
               TabIndex        =   30
               Top             =   1920
               Width           =   3972
            End
            Begin VB.Label lblHistory 
               Height          =   252
               Index           =   6
               Left            =   120
               TabIndex        =   29
               Top             =   1680
               Width           =   3972
            End
            Begin VB.Label lblHistory 
               Height          =   252
               Index           =   5
               Left            =   120
               TabIndex        =   28
               Top             =   1440
               Width           =   3972
            End
            Begin VB.Label lblHistory 
               Height          =   252
               Index           =   4
               Left            =   120
               TabIndex        =   27
               Top             =   1200
               Width           =   3972
            End
            Begin VB.Label lblHistory 
               Height          =   252
               Index           =   3
               Left            =   120
               TabIndex        =   26
               Top             =   960
               Width           =   3972
            End
            Begin VB.Label lblHistory 
               Height          =   252
               Index           =   2
               Left            =   120
               TabIndex        =   25
               Top             =   720
               Width           =   3972
            End
            Begin VB.Label lblHistory 
               Height          =   252
               Index           =   1
               Left            =   120
               TabIndex        =   24
               Top             =   480
               Width           =   3972
            End
            Begin VB.Label lblHistory 
               Height          =   252
               Index           =   0
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Width           =   3972
            End
         End
         Begin VB.CommandButton cmdName 
            Caption         =   "enter"
            Height          =   252
            Left            =   3720
            TabIndex        =   21
            Top             =   600
            Width           =   612
         End
         Begin VB.TextBox txtName 
            Height          =   288
            Left            =   1920
            TabIndex        =   20
            Top             =   600
            Width           =   1692
         End
         Begin VB.Label lblName 
            Caption         =   "Please Enter your name"
            Height          =   252
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1812
         End
         Begin VB.Label lblMoves 
            Caption         =   "You have completed it in      moves"
            Height          =   252
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   4212
         End
      End
      Begin VB.CommandButton Box 
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
         Index           =   16
         Left            =   3360
         TabIndex        =   16
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   15
         Left            =   2280
         TabIndex        =   15
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   14
         Left            =   1200
         TabIndex        =   14
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   13
         Left            =   120
         TabIndex        =   13
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   12
         Left            =   3360
         TabIndex        =   12
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   11
         Left            =   2280
         TabIndex        =   11
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   10
         Left            =   1200
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   9
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   8
         Left            =   3360
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   7
         Left            =   2280
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   6
         Left            =   1200
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   4
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   2
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Box 
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
         Index           =   3
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label lblnmoves 
      Height          =   252
      Left            =   0
      TabIndex        =   33
      Top             =   4200
      Width           =   4452
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRestart 
         Caption         =   "&Restart"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuMoves 
         Caption         =   "&Moves"
      End
   End
End
Attribute VB_Name = "Puzzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Counter As Integer


Private Sub Form_Load()

frameEnd.Visible = False

For temp = 1 To 16
Box(temp).Visible = True
Next temp

Dim startArray(0 To 15) As Integer

For i = 0 To 15
    startArray(i) = i
  Next i
  
Randomize

For constInt = 15 To 0 Step -1

tempint = Int(constInt * Rnd)
swapInt = startArray(constInt)
startArray(constInt) = startArray(tempint)
startArray(tempint) = swapInt
Box(constInt + 1).Caption = startArray(constInt)
If startArray(constInt) = 0 Then
    Box(constInt + 1).Visible = False
    End If
Next constInt

Counter = 0

End Sub

Private Sub Box_Click(Index As Integer)
 
Counter = Counter + 1
lblnmoves = ""

Select Case Index

    Case 1
        Call Change(2, 1)
        Call Change(5, 1)
    Case 2
        Call Change(1, 2)
        Call Change(6, 2)
        Call Change(3, 2)
    Case 3
        Call Change(2, 3)
        Call Change(7, 3)
        Call Change(4, 3)
    Case 4
        Call Change(3, 4)
        Call Change(8, 4)
    Case 5
        Call Change(1, 5)
        Call Change(6, 5)
        Call Change(9, 5)
    Case 6
        Call Change(5, 6)
        Call Change(2, 6)
        Call Change(7, 6)
        Call Change(10, 6)
    Case 7
        Call Change(6, 7)
        Call Change(8, 7)
        Call Change(3, 7)
        Call Change(11, 7)
    Case 8
        Call Change(7, 8)
        Call Change(4, 8)
        Call Change(12, 8)
    Case 9
        Call Change(5, 9)
        Call Change(10, 9)
        Call Change(13, 9)
    Case 10
        Call Change(6, 10)
        Call Change(9, 10)
        Call Change(11, 10)
        Call Change(14, 10)
    Case 11
        Call Change(10, 11)
        Call Change(7, 11)
        Call Change(12, 11)
        Call Change(15, 11)
    Case 12
        Call Change(11, 12)
        Call Change(8, 12)
        Call Change(16, 12)
    Case 13
        Call Change(9, 13)
        Call Change(14, 13)
    Case 14
        Call Change(15, 14)
        Call Change(13, 14)
        Call Change(10, 14)
    Case 15
        Call Change(14, 15)
        Call Change(11, 15)
        Call Change(16, 15)
    Case 16
        Call Change(15, 16)
        Call Change(12, 16)
        If (Box(1).Caption = 1) And (Box(2).Caption = 2) And (Box(3).Caption = 3) And (Box(4).Caption = 4) And (Box(5).Caption = 5) And (Box(6).Caption = 6) And (Box(7).Caption = 7) And (Box(8).Caption = 8) And (Box(9).Caption = 9) And (Box(10).Caption = 10) And (Box(11).Caption = 11) And (Box(12).Caption = 12) And (Box(13).Caption = 13) And (Box(14).Caption = 14) And (Box(15).Caption = 15) And (Box(16).Caption = 0) Then
            frameEnd.Visible = True
            lblMoves.Caption = "You have completed the puzzle in " & Str(Counter) & " moves"
        End If

    End Select
    
    
        
    
End Sub


Private Sub Change(X As Integer, Y As Integer)

If Box(X).Caption = "0" Then

temp = Box(X).Caption
Box(X).Caption = Box(Y).Caption
Box(Y).Caption = temp
Box(Y).Visible = False
Box(X).Visible = True
Box(X).SetFocus

End If

End Sub

Private Sub cmdName_Click()

If getcounter(lblHistory(9)) > Counter Then
    lblHistory(9) = txtName & " " & Counter & " moves"
End If

For i = 8 To 0 Step -1
    If getcounter(lblHistory(i)) > getcounter(lblHistory(i + 1)) Then
        temp = lblHistory(i)
        lblHistory(i) = lblHistory(i + 1)
        lblHistory(i + 1) = temp
    End If
Next i
    
frameRecord.Visible = True

End Sub


Private Sub mnuMoves_Click()
    lblnmoves = "   " & Counter & " moves"
End Sub

Private Sub mnuRestart_Click()
    Call Form_Load
End Sub

Private Function getcounter(tempstr As String) As Integer

If tempstr = "" Then

    getcounter = 9999
Else

    tempint = Val(Left(Right(tempstr, 7), 1))
    
    incint = 10
    i = 8

    Do While tempchar >= "0" And tempchar <= "9"
       tempint = tempint + (Val(tempchar) * incint)
       tempchar = Left(Right(tempstr, i), 1)
       incint = incint * 10
       i = i + 1
    Loop
   
    getcounter = tempint
End If
   
End Function
