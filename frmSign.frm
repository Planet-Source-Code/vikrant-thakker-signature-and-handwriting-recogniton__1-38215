VERSION 5.00
Begin VB.Form frmSign 
   Caption         =   "Anveshak - A Spying tool"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   765
      Width           =   7485
      Begin VB.Timer TmrStatusPause 
         Left            =   4860
         Top             =   90
      End
      Begin VB.Timer Tmr 
         Left            =   5670
         Top             =   45
      End
      Begin VB.ComboBox cmbLetList 
         Height          =   315
         Left            =   135
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   2385
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         BackColor       =   &H00E84A4A&
         DrawWidth       =   8
         ForeColor       =   &H00FFFFFF&
         Height          =   2805
         Left            =   2430
         MouseIcon       =   "frmSign.frx":0000
         MousePointer    =   99  'Custom
         ScaleHeight     =   185
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   277
         TabIndex        =   4
         Top             =   1305
         Width           =   4185
      End
      Begin VB.TextBox txtSign 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   3
         Top             =   540
         Width           =   1725
      End
      Begin Recog.cmd cmdSubmit 
         Height          =   555
         Left            =   4635
         TabIndex        =   12
         Top             =   450
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   979
         BTYPE           =   5
         TX              =   "&Recognize Me"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         FCOL            =   0
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status"
         Height          =   330
         Left            =   135
         TabIndex        =   7
         Top             =   2790
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblRes 
         Caption         =   "Result"
         Height          =   375
         Left            =   2295
         TabIndex        =   6
         Top             =   4365
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Signature    :"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   495
         TabIndex        =   2
         Top             =   1575
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name  :"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   495
         TabIndex        =   1
         Top             =   540
         Width           =   1770
      End
   End
   Begin Recog.cmd cmdBack 
      Height          =   555
      Left            =   4815
      TabIndex        =   9
      Top             =   5760
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "&Back"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
   End
   Begin Recog.cmd cmdHelp 
      Height          =   555
      Left            =   2745
      TabIndex        =   10
      Top             =   5760
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "&Help"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
   End
   Begin Recog.cmd cmdCLS 
      Height          =   555
      Left            =   675
      TabIndex        =   11
      Top             =   5760
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "&Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Signature Recognition"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8070
   End
End
Attribute VB_Name = "frmSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
Unload Me
frmMain.Show
End Sub

Private Sub cmdRecog_Click()
If Trim(txtSign.Text) = "" Then
    MsgBox "Enter the UserName", vbOKOnly, "Anveshak"
End If
End Sub


Private Sub cmdCLS_Click()
picDraw.Cls
lblRes.Visible = False
End Sub

Private Sub cmdExit_Click()
Ans = MsgBox("Do you want to Exit ?", vbYesNo, "Anveshak")
If Ans = vbYes Then
    End
Else
    Exit Sub
End If
End
End Sub

Private Sub cmdHelp_Click()
MsgBox "Sorry for inconvinience ! UNDER CONSTRUCTION !!!", vbOKOnly, "Anveshak"
End Sub


Private Sub cmdSubmit_Click()
' Displays the result
If Trim(lblRes.Caption) = Trim(txtSign.Text) Then
    MsgBox "Login Successfull !", vbOKOnly, "Anveshak"
ElseIf Trim(lblRes.Caption) <> Trim(txtSign.Text) Then
    MsgBox "Invalid User !", vbOKOnly, "Anveshak"
End If

End Sub

Private Sub Form_Activate()
' If no characters are teached to the computer, then display an appropriate message
If rsChar.RecordCount = 0 Then
MsgBox "Note if this is your first time using the program you should probably " & vbNewLine & "edit every letter because right now they are custom to my handwriting." & vbNewLine & "So you should edit it to yours to make it more accurate.  Also a letter " & vbNewLine & "is finished once you release the mouse so be careful on letter like 'i'.", vbOKOnly, "Anveshak"
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Call LoadAll
For i = 0 To rsChar.RecordCount - 1 Step 1
    cmbLetList.AddItem Letter(i), i
Next i
End Sub


Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Store the current X and Y values when the mouse button is clicked down
' Enable and start the timer, so that we can know the time taken by user to draw a character

Call MouseDown
Tmr.Enabled = True
picDraw.CurrentX = X
picDraw.CurrentY = Y
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Direc As Integer
Dim BuffX As Integer, BuffY As Integer
Static Count As Integer

If WriteLet = True Then
    Count = Count + 1
    If Count Mod 2 = 0 Then
        If NumLet < 200 Then
            BuffX = X   ' Stores the new current value of X
            BuffY = Y   ' Stores the new current value of Y
' The direction value is calculated and stored in the variable Direc

            Direc = Direction(HoldX, HoldY, BuffX, BuffY)
            HoldX = X
            HoldY = Y
            
            picDraw.Line -(BuffX, BuffY)
            
    
            LetterMovement(NumLet) = Direc
            
            NumLet = NumLet + 1
        Else
            lblStatus.Caption = "Letter Limit"
        End If
    End If
End If
End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Dim F As Integer
ReDim Letter(-1 To rsChar.RecordCount) As String
ReDim Score(rsChar.RecordCount) As Single
frmOCR.Cls


'stretch or compress array to fit 100 into the BuffArray
Distort = NumLet / 100
For i = 0 To 100
    BuffArray(i) = LetterMovement(Int(i * Distort))
Next i


If WriteFile = False Then
    'calculate the score for each letter
    For F = 0 To rsChar.RecordCount - 1
        Dim Total As Integer
        For i = 0 To 100

            
            If BuffArray(i) > Alphabet(F).Direc(i) Then
                Difference = BuffArray(i) - Alphabet(F).Direc(i)
            Else
                Difference = Alphabet(F).Direc(i) - BuffArray(i)
            End If
            
            'exceptions because the where the circle ends
            If BuffArray(i) = 0 And Alphabet(F).Direc(i) = 15 Then
                Difference = 1
            ElseIf BuffArray(i) = 0 And Alphabet(F).Direc(i) = 14 Then
                Difference = 2
            ElseIf BuffArray(i) = 1 And Alphabet(F).Direc(i) = 15 Then
                Difference = 2
            ElseIf BuffArray(i) = 1 And Alphabet(F).Direc(i) = 14 Then
                Difference = 3
            End If
            
            Score(F) = Score(F) + (8 - Difference)
            Total = Total + 8
        Next i
        'put score into percent
        Score(F) = Score(F) / Total * 100
        Total = 0
        
If rsChar.RecordCount > 0 Then
rsChar.MoveFirst
    rsChar.Move (F)
       ' frmOCR.Print rsChar!Char & ":  " & CInt(Score(F)) & "%"
       'lblRes.Caption = rsChar!Char & ":  " & CInt(Score(F)) & "%"
        lblRes.Visible = False
        lblRes.Caption = rsChar!Char
End If
        
      '  frmOCR.Print Letter(F) & ":  " & CInt(Score(F)) & "%"
    Next F
    
    Highest = 0
    HighScore = Score(0)
    
    For i = 1 To rsChar.RecordCount - 1
        If Score(i) > HighScore Then
            Highest = i
            HighScore = Mid(Score(i), 1, 2)
        End If
    Next i
    frmOCR.Print ""
    If HighScore < 50 Then
        If rsChar.RecordCount > 0 Then
            rsChar.MoveFirst
            rsChar.Move (Highest)
          '  frmOCR.Print "?" & rsChar!Char & "?  Percent: "; CInt(HighScore) & "%"
          '  lblRes.Caption = "?" & rsChar!Char & "?  Percent: " & CInt(HighScore) & "%"
            lblRes.Visible = False
            lblRes.Caption = rsChar!Char
        End If
            'frmOCR.Print "?" & Letter(Highest) & "?  Percent: "; CInt(HighScore) & "%"
      '      TS.Speak (Letter(Highest))
    Else
    If rsChar.RecordCount > 0 Then
            rsChar.MoveFirst
            rsChar.Move (Highest)
            'frmOCR.Print rsChar!Char & "  Percent: "; CInt(HighScore) & "%"
            lblRes.Visible = False
            lblRes.Caption = rsChar!Char
        End If
           ' frmOCR.Print Letter(Highest) & "  Percent: "; CInt(HighScore) & "%"
       '     TS.Speak (Letter(Highest))
    End If
    
    
    lblStatus.Caption = "Drawing"

End If
WriteFile = False
WriteLet = False  ' This tells the computer that we have completed drawing a character
NumLet = 0

End Sub

Private Sub Tmr_Timer()
tm = tm + 1
End Sub

Private Sub tmrStatusPause_Timer()
lblStatus.Caption = "Drawing"
TmrStatusPause.Interval = 0     ' This increments the variable tm by 1, so as to get the value of total time taken to draw the character
End Sub

