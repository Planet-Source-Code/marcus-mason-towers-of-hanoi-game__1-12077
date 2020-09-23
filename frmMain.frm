VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Towers of Hanoi"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDone 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2640
      Top             =   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Happy_lobster@yahoo.co.uk"
      ForeColor       =   &H000080FF&
      Height          =   435
      Left            =   3360
      TabIndex        =   4
      Top             =   2760
      Width           =   4065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Marcus Mason"
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   2550
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moves:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label lblMoves 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1320
      TabIndex        =   0
      Top             =   2280
      Width           =   195
   End
   Begin VB.Line lneDisc 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   10
      Index           =   7
      X1              =   24
      X2              =   152
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Line lneDisc 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   10
      Index           =   6
      X1              =   32
      X2              =   144
      Y1              =   112
      Y2              =   112
   End
   Begin VB.Line lneDisc 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   10
      Index           =   5
      X1              =   40
      X2              =   136
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line lneDisc 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   10
      Index           =   4
      X1              =   48
      X2              =   128
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Line lneDisc 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   10
      Index           =   3
      X1              =   56
      X2              =   120
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Line lneDisc 
      BorderColor     =   &H000080FF&
      BorderWidth     =   10
      Index           =   2
      X1              =   64
      X2              =   112
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Line lneDisc 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Index           =   1
      X1              =   72
      X2              =   104
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Line lneTower 
      BorderWidth     =   10
      Index           =   2
      X1              =   240
      X2              =   240
      Y1              =   16
      Y2              =   144
   End
   Begin VB.Line lneTower 
      BorderWidth     =   10
      Index           =   1
      X1              =   88
      X2              =   88
      Y1              =   16
      Y2              =   144
   End
   Begin VB.Line Line2 
      BorderWidth     =   10
      X1              =   16
      X2              =   464
      Y1              =   144
      Y2              =   144
   End
   Begin VB.Line lneTower 
      BorderWidth     =   10
      Index           =   3
      X1              =   392
      X2              =   392
      Y1              =   16
      Y2              =   144
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoMove As Integer
Dim Stack(3, 10) As Integer
Dim DiscCount(3) As Integer
Dim StackNo(10) As Integer
Dim LineWidth(10) As Integer

Dim OriginalTower As Integer
Dim OriginalCount As Integer

Dim DiscNumber As Integer
Dim TowerGap As Integer
Dim AllowMove As Boolean
Dim CurrentTower As Integer
Dim TotalDiscs As Integer
Dim Moves As Integer

Private Sub Form_Load()
    Dim C As Integer
    
    TotalDiscs = 7
    'Set initial stack settings
    For C = 1 To TotalDiscs
        StackNo(C) = 1
        Stack(1, C) = (TotalDiscs + 1) - C
        DiscCount(1) = TotalDiscs
        LineWidth(C) = lneDisc(C).X2 - lneDisc(C).X1
    Next
    Moves = 0
    
    'Get gap value between towers
    TowerGap = (lneTower(2).X1 - lneTower(1).X1) / 2
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim C As Integer

    If Button = 1 Then
    
        'Calculate which disc has been clicked on
        DiscNumber = 0
        For C = 1 To TotalDiscs
            With lneDisc(C)
                If X > .X1 Then
                    If X < .X2 Then
                        If Y > .Y1 - 8 Then
                            If Y < .Y1 + 8 Then
                                DiscNumber = C
                            End If
                        End If
                    End If
                End If
            End With
        Next
    
        If DiscNumber = 0 Then Exit Sub
        
        'Check if this is the top disc
        CurrentTower = StackNo(DiscNumber)
        
        If Stack(CurrentTower, DiscCount(CurrentTower)) = DiscNumber Then
            
            'Move to top of tower
            With lneDisc(DiscNumber)
                .Y1 = lneTower(CurrentTower).Y1 + 8
                .Y2 = .Y1
            End With
            
            'Setup variables ready to move disk
            OriginalTower = CurrentTower
            OriginalCount = DiscCount(CurrentTower)
            
            Stack(CurrentTower, DiscCount(CurrentTower)) = 0
            DiscCount(CurrentTower) = DiscCount(CurrentTower) - 1
    
            AllowMove = True
            
        End If
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        If AllowMove = True Then
            
            'See what peg the mouse is over
            If X < lneTower(1).X1 + TowerGap Then
                CurrentTower = 1
            ElseIf X >= lneTower(1).X1 + TowerGap And X < lneTower(2).X1 + TowerGap Then
                CurrentTower = 2
            ElseIf X >= lneTower(2).X1 + TowerGap Then
                CurrentTower = 3
            End If
            
            'Move the disk over the correct peg
            With lneDisc(DiscNumber)
                .X1 = lneTower(CurrentTower).X1 - (LineWidth(DiscNumber) / 2)
                .X2 = .X1 + LineWidth(DiscNumber)
            End With
            
        End If
    End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim C As Single
    
    'User has finished thier move
    If Button = 1 Then
        If AllowMove = True Then
            If DiscCount(CurrentTower) > 0 Then
            
                'Check if move is allowed
                If Stack(CurrentTower, DiscCount(CurrentTower)) < DiscNumber Then
                    
                    'Move not allowed
                    Beep
                    
                    'Restore old variables
                    DiscCount(OriginalTower) = DiscCount(OriginalTower) + 1
                    StackNo(DiscNumber) = OriginalTower
                    Stack(OriginalTower, DiscCount(OriginalTower)) = DiscNumber
    
                    'Replace disk
                    With lneDisc(DiscNumber)
                        .X1 = lneTower(OriginalTower).X1 - (LineWidth(DiscNumber) / 2)
                        .X2 = .X1 + LineWidth(DiscNumber)
                        
                        For C = .Y1 To lneTower(OriginalTower).Y2 - (DiscCount(OriginalTower) * 16)
                            .Y1 = C
                            .Y2 = C
                        Next
                    End With
                
                Else
                    'Do the move
                    DoMove
                End If
            Else
                'Do the move
                DoMove
            End If

            'Reset variables
            DiscNumber = 0
            AllowMove = False
        End If
    End If
    
    
End Sub

Sub DoMove()
    Dim C As Single

    'Set variables
    DiscCount(CurrentTower) = DiscCount(CurrentTower) + 1
    StackNo(DiscNumber) = CurrentTower
    Stack(CurrentTower, DiscCount(CurrentTower)) = DiscNumber
    
    'Animate disk falling
    With lneDisc(DiscNumber)
        For C = .Y1 To lneTower(CurrentTower).Y2 - (DiscCount(CurrentTower) * 16)
            .Y1 = C
            .Y2 = C
            .Refresh
        Next
    End With
    
    'Increment no of moves
    If OriginalTower <> CurrentTower Then
        Moves = Moves + 1
        lblMoves = Moves
    End If
    
    'Check if completed
    If CurrentTower > 1 And DiscCount(CurrentTower) = TotalDiscs Then
        MsgBox "Hoorah you win!", 48
        
        'Show finished routine
        tmrDone.Enabled = True
    End If
    
End Sub


Private Sub tmrDone_Timer()
    Dim C As Integer
    Dim FirstColour As Long
    
    'Colour disks with colours flashing upwards
    FirstColour = lneDisc(1).BorderColor
    For C = 1 To TotalDiscs - 1
        lneDisc(C).BorderColor = lneDisc(C + 1).BorderColor
    Next
    lneDisc(TotalDiscs).BorderColor = FirstColour
    
    
End Sub
