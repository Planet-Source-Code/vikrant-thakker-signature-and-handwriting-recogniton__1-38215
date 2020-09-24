VERSION 5.00
Begin VB.Form frmTeach 
   Caption         =   "Anveshak - A Spying tool"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbLetList 
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Text            =   "Select Letter"
      Top             =   5445
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer tmrStatusPause 
      Left            =   405
      Top             =   3510
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   5445
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdWorks 
      Caption         =   "About"
      Height          =   375
      Left            =   5085
      TabIndex        =   12
      Top             =   5445
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6660
      TabIndex        =   11
      Top             =   5445
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Timer Tmr 
      Interval        =   1
      Left            =   405
      Top             =   2565
   End
   Begin VB.Frame Frame1 
      Height          =   4110
      Left            =   0
      TabIndex        =   7
      Top             =   1035
      Width           =   4695
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DrawWidth       =   7
         ForeColor       =   &H00C00000&
         Height          =   2445
         Left            =   945
         MousePointer    =   2  'Cross
         ScaleHeight     =   161
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   192
         TabIndex        =   8
         Top             =   765
         Width           =   2910
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Here"
         Height          =   330
         Left            =   945
         TabIndex        =   10
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "The speed at which you write should be the same as when you edited"
         Height          =   735
         Left            =   2610
         TabIndex        =   9
         Top             =   3285
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4110
      Left            =   4680
      TabIndex        =   1
      Top             =   1035
      Width           =   4695
      Begin VB.TextBox txtChar 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   540
         TabIndex        =   20
         Top             =   1845
         Width           =   3660
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "&Enable Sound"
         Height          =   330
         Left            =   540
         TabIndex        =   3
         Top             =   315
         Width           =   1545
      End
      Begin Recog.cmd cmdEdit 
         Height          =   555
         Left            =   540
         TabIndex        =   2
         Top             =   2700
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   979
         BTYPE           =   5
         TX              =   "&Teach"
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
      Begin Recog.cmd cmdCLS 
         Height          =   555
         Left            =   2430
         TabIndex        =   4
         Top             =   2700
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
      Begin VB.Label Label7 
         Caption         =   "Teach"
         Height          =   375
         Left            =   540
         TabIndex        =   21
         Top             =   1485
         Width           =   1950
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Result is :"
         Height          =   375
         Left            =   585
         TabIndex        =   6
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label lblRes 
         Caption         =   "a"
         Height          =   330
         Left            =   1485
         TabIndex        =   5
         Top             =   3600
         Visible         =   0   'False
         Width           =   825
      End
   End
   Begin Recog.cmd cmdHelp 
      Height          =   510
      Left            =   1755
      TabIndex        =   0
      Top             =   5625
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
   Begin Recog.cmd cmdBack 
      Height          =   510
      Left            =   3060
      TabIndex        =   15
      Top             =   5625
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
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   5850
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblStatus 
      Caption         =   "Drawing"
      Height          =   255
      Left            =   6300
      TabIndex        =   18
      Top             =   5985
      Width           =   1260
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Help Section :"
      Height          =   375
      Left            =   585
      TabIndex        =   17
      Top             =   5850
      Width           =   1140
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Teach Computer"
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
      TabIndex        =   16
      Top             =   90
      Width           =   9330
   End
End
Attribute VB_Name = "frmTeach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This form is concerned with the teaching part of the Anveshak

'Option Explicit
Dim BuffArray(100) As String    ' Creates an array of 100 elements
Dim strFile As String           ' Stores the complete line in the array
Dim Distort As Single
Dim Difference As Integer       ' Gets the difference of the previous and the current position of mouse
Dim Score() As Single           ' After comparison with the array, each unit is increased when a matched direction is found
Dim i As Integer
Dim F As Integer
Dim Highest As Integer          ' The alphabet that has the highest score
Dim HighScore As Single         ' The amount of highest score

Private Sub cmdBack_Click()
Unload Me
frmMain.Show
End Sub

Private Sub cmdClear_Click()
picDraw.Cls
End Sub

Private Sub cmdCLS_Click()
picDraw.Cls
lblRes.Visible = False
End Sub

Private Sub cmdEdit_Click()
ReDim Letter(-1 To rsChar.RecordCount) As String
'If cmbLetList.ListIndex <> -1 Then
If Trim(txtChar.Text) <> "" Then
    WriteFile = True
    picDraw.Cls
'    lblStatus.Caption = "Editing: " & Letter(cmbLetList.ListIndex)
    lblStatus.Caption = "Teaching: " & txtChar.Text
Else
    MsgBox "Please write a lettet to teach.", vbCritical, "Anveshak"
End If
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

Private Sub cmdWorks_Click()
frmWorks.Show 1    ' Opens and Displays the form frmWorks
End Sub

Private Sub Form_Activate()
If rsChar.RecordCount = 0 Then
MsgBox "Note if this is your first time using the program you should probably " & vbNewLine & "edit every letter because right now they are custom to my handwriting." & vbNewLine & "So you should edit it to yours to make it more accurate.  Also a letter " & vbNewLine & "is finished once you release the mouse so be careful on letter like 'i'.", vbOKOnly, "Anveshak"
End If
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
Dim BuffArray(100) As String
Dim strFile As String
Dim Distort As Single
Dim Difference As Integer
Dim Score() As Single
Dim i As Integer
Dim F As Integer
Dim Highest As Integer
Dim HighScore As Single
ReDim Letter(-1 To rsChar.RecordCount) As String
ReDim Score(rsChar.RecordCount) As Single
frmOCR.Cls

If cmbLetList.ListIndex <> -1 Then
'stretch or compress array to fit 100 into the BuffArray
Distort = NumLet / 100
For i = 0 To 100
    BuffArray(i) = LetterMovement(Int(i * Distort))
    If WriteFile = True Then
        Alphabet(cmbLetList.ListIndex).Direc(i) = LetterMovement(Int(i * Distort))
    End If
Next i
ElseIf cmbLetList.ListIndex = -1 Then
Distort = NumLet / 100
For i = 0 To 100
    BuffArray(i) = LetterMovement(Int(i * Distort))
    If WriteFile = True Then
        Alphabet(0).Direc(i) = LetterMovement(Int(i * Distort))
    End If
Next i
End If '

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


    lblStatus.Caption = "Drawing" '
Else
rsChar.AddNew
   
    For F = 0 To rsChar.RecordCount
        For i = 0 To 100
            If Val(Alphabet(F).Direc(i)) < 10 Then
                strFile = strFile & "0"
            End If
            strFile = strFile & Alphabet(F).Direc(i)
                If i = 100 Then
                    strFile = strFile & Alphabet(F).Direc(i)
                    Tmr.Enabled = False
      
            rsChar!String = strFile
            rsChar!Char = txtChar.Text
            rsChar!Time = tm
            rsChar.Update
                End If
        Next i
        If F <> rsChar.RecordCount Then
            strFile = strFile & vbNewLine
        End If
    Next F
    
    lblStatus.Caption = "Letter Saved"
    TmrStatusPause.Interval = 1000
    
    Call LoadAll   ' This will reallocate the elements in array
   
'End If

'Call Teach
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
TmrStatusPause.Interval = 0
End Sub

