VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   Caption         =   "About Anveshak"
   ClientHeight    =   7560
   ClientLeft      =   4635
   ClientTop       =   1875
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   504
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   503
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Recog.cmd cmdOK 
      Height          =   465
      Left            =   5940
      TabIndex        =   17
      Top             =   7020
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   820
      BTYPE           =   5
      TX              =   "Back"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   2895
      Left            =   4230
      TabIndex        =   4
      Top             =   4005
      Width           =   3255
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   " MAIDAN GARHI, NEW DELHI - 110068"
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   12
         Top             =   2565
         Width           =   3075
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "IGNOU"
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   11
         Top             =   2340
         Width           =   3075
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "INDIRA GANDHI NATIONAL OPEN UNI."
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   10
         Top             =   2115
         Width           =   3075
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Project For"
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   9
         Top             =   1890
         Width           =   3075
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Enroll No  :  000445495,  Sem.  :  6th"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   1485
         Width           =   3075
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Name        :  Vikrant Thakker"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   1080
         Width           =   3120
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "The Spying Tool"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   45
         TabIndex        =   6
         Top             =   450
         Width           =   3120
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Anveshak"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   3120
      End
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Vikrant Thakker"
      ForeColor       =   &H00C0C0FF&
      Height          =   345
      Left            =   90
      TabIndex        =   16
      Top             =   3690
      Width           =   3975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Regards,"
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   90
      TabIndex        =   15
      Top             =   3330
      Width           =   3975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWorks.frx":0000
      ForeColor       =   &H0000FF00&
      Height          =   2955
      Left            =   90
      TabIndex        =   14
      Top             =   315
      Width           =   3975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Dear Users,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   45
      TabIndex        =   13
      Top             =   90
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWorks.frx":0329
      ForeColor       =   &H0000FF00&
      Height          =   3135
      Left            =   4260
      TabIndex        =   3
      Top             =   300
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Program Logic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   4275
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWorks.frx":05C3
      ForeColor       =   &H0000FF00&
      Height          =   2775
      Left            =   90
      TabIndex        =   0
      Top             =   4755
      Width           =   3975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
Load frmMain
frmMain.Show
End Sub
