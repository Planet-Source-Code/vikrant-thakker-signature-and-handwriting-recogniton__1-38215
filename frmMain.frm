VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Anveshak - A spying tool"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5820
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin Recog.cmd cmdRecog 
         Height          =   600
         Left            =   2970
         TabIndex        =   1
         Top             =   1665
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   1058
         BTYPE           =   5
         TX              =   "Character Recognition"
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
      Begin Recog.cmd cmdSign 
         Height          =   600
         Left            =   2970
         TabIndex        =   2
         Top             =   2565
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   1058
         BTYPE           =   5
         TX              =   "Signature Recognition"
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
      Begin Recog.cmd cmdTeach 
         Height          =   600
         Left            =   2970
         TabIndex        =   3
         Top             =   3510
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   1058
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
      Begin Recog.cmd cmdQuit 
         Height          =   600
         Left            =   4680
         TabIndex        =   4
         Top             =   4905
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   1058
         BTYPE           =   5
         TX              =   "&Quit"
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
         Height          =   600
         Left            =   1215
         TabIndex        =   6
         Top             =   4950
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   1058
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
      Begin VB.Image Image2 
         Height          =   1050
         Left            =   1080
         Picture         =   "frmMain.frx":0CCA
         Top             =   0
         Width           =   6000
      End
      Begin VB.Image Image1 
         Height          =   6000
         Left            =   0
         Picture         =   "frmMain.frx":1899
         Top             =   0
         Width           =   1050
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Main Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   4095
         TabIndex        =   5
         Top             =   180
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Character and Signature Recognition Software
' Developed by Vikrant Thakker
' email : vikrant_thakker@yahoo.com, vikrant_thakker@hotmail.com
' India

' This is the Main Screen that lets you navigate through
' different parts of the software...

' Date : 11/08/02


Private Sub cmdAbout_Click()
Unload Me
frmAbout.Show
End Sub

Private Sub cmdQuit_Click()
Ans = MsgBox("Do you want to Quit ?", vbYesNo, "Anveshak")
If Ans = vbYes Then
       End
ElseIf Ans = vbNo Then
    Exit Sub
End If
End Sub

Private Sub cmdRecog_Click()
Unload Me
Load frmOCR
frmOCR.Show
End Sub

Private Sub cmdSign_Click()
Unload Me
Load frmSign
frmSign.Show
End Sub

Private Sub cmdTeach_Click()
Unload Me
Load frmTeach
frmTeach.Show
End Sub

