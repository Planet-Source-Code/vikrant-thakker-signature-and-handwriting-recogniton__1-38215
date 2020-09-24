VERSION 5.00
Begin VB.Form frmOCR 
   AutoRedraw      =   -1  'True
   Caption         =   "Anveshak - A Spying tool"
   ClientHeight    =   6570
   ClientLeft      =   1155
   ClientTop       =   1005
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   2  'CenterScreen
   Begin Recog.cmd cmdHelp 
      Height          =   510
      Left            =   1800
      TabIndex        =   12
      Top             =   5715
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   900
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
   End
   Begin VB.Frame Frame2 
      Height          =   4110
      Left            =   4725
      TabIndex        =   5
      Top             =   1125
      Width           =   4695
      Begin VB.ComboBox cmbLetList 
         Height          =   315
         Left            =   2385
         TabIndex        =   18
         Text            =   "Select Letter"
         Top             =   3645
         Visible         =   0   'False
         Width           =   1800
      End
      Begin Recog.cmd cmdSubmit 
         Height          =   555
         Left            =   450
         TabIndex        =   9
         Top             =   2385
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   979
         BTYPE           =   5
         TX              =   "&Submit"
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
      Begin VB.CheckBox chkSound 
         Caption         =   "&Enable Sound"
         Height          =   330
         Left            =   450
         TabIndex        =   8
         Top             =   3600
         Width           =   1545
      End
      Begin Recog.cmd cmdCLS 
         Height          =   555
         Left            =   2385
         TabIndex        =   16
         Top             =   2385
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.Label lblRes 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "AvantGarde Bk BT"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1590
         Left            =   450
         TabIndex        =   20
         Top             =   675
         Visible         =   0   'False
         Width           =   3750
      End
      Begin VB.Label Label7 
         Caption         =   "List of taught characters"
         Height          =   285
         Left            =   2385
         TabIndex        =   17
         Top             =   3285
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Result is :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E84A4A&
         Height          =   375
         Left            =   405
         TabIndex        =   10
         Top             =   315
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4110
      Left            =   45
      TabIndex        =   4
      Top             =   1125
      Width           =   4695
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DrawWidth       =   8
         ForeColor       =   &H00C00000&
         Height          =   2445
         Left            =   945
         MouseIcon       =   "frmOCR.frx":0000
         MousePointer    =   99  'Custom
         ScaleHeight     =   161
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   192
         TabIndex        =   6
         Top             =   765
         Width           =   2910
      End
      Begin VB.Label Label2 
         Caption         =   "The speed at which you write should be the same as when you edited"
         Height          =   735
         Left            =   2610
         TabIndex        =   14
         Top             =   3285
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Here"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E84A4A&
         Height          =   330
         Left            =   945
         TabIndex        =   7
         Top             =   360
         Width           =   1545
      End
   End
   Begin VB.TextBox txtChar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7785
      TabIndex        =   3
      Top             =   180
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Timer Tmr 
      Interval        =   1
      Left            =   450
      Top             =   2655
   End
   Begin VB.CommandButton cmdWorks 
      Caption         =   "About"
      Height          =   375
      Left            =   6615
      TabIndex        =   2
      Top             =   5850
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrStatusPause 
      Left            =   450
      Top             =   3600
   End
   Begin Recog.cmd cmdBack 
      Height          =   510
      Left            =   3105
      TabIndex        =   13
      Top             =   5715
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   900
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
   End
   Begin Recog.cmd cmdAbout 
      Height          =   510
      Left            =   4410
      TabIndex        =   19
      Top             =   5715
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   900
      BTYPE           =   5
      TX              =   "&About"
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Character Recognition"
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
      Height          =   825
      Left            =   45
      TabIndex        =   15
      Top             =   90
      Width           =   9420
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Help Section :"
      Height          =   375
      Left            =   630
      TabIndex        =   11
      Top             =   5940
      Width           =   1140
   End
   Begin VB.Label lblStatus 
      Caption         =   "Drawing"
      Height          =   255
      Left            =   6615
      TabIndex        =   1
      Top             =   5310
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   255
      Left            =   7920
      TabIndex        =   0
      Top             =   5310
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmOCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkSound_Click()
MsgBox "Sorry for inconvinience ! UNDER CONSTRUCTION !!!", vbOKOnly, "Anveshak"
End Sub

Private Sub cmdAbout_Click()
Unload Me
frmAbout.Show
End Sub

Private Sub cmdBack_Click()
Unload Me
frmMain.Show
End Sub

Private Sub cmdCLS_Click()
picDraw.Cls
lblRes.Visible = False
End Sub

Private Sub cmdHelp_Click()
MsgBox "Sorry for inconvinience ! UNDER CONSTRUCTION !!!", vbOKOnly, "Anveshak"
End Sub

Private Sub cmdSubmit_Click()
lblRes.Visible = True     ' Display the result
End Sub


Private Sub Form_Activate()
' If no characters are taught to computer, then display an appropriate message
If rsChar.RecordCount = 0 Then
MsgBox "Note if this is your first time using the program you should probably " & vbNewLine & "edit every letter because right now they are custom to my handwriting." & vbNewLine & "So you should edit it to yours to make it more accurate.  Also a letter " & vbNewLine & "is finished once you release the mouse so be careful on letter like 'i'.", vbOKOnly, "Anveshak"
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Call LoadAll
'Add the list of taught characters in the combo box
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
tm = tm + 1    ' This increments the variable tm by 1, so as to get the value of total time taken to draw the character
End Sub

Private Sub tmrStatusPause_Timer()
lblStatus.Caption = "Drawing"
TmrStatusPause.Interval = 0
End Sub
