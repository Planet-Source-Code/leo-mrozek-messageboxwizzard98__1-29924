VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   Caption         =   "About . . ."
   ClientHeight    =   4920
   ClientLeft      =   1665
   ClientTop       =   1815
   ClientWidth     =   5250
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4920
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRead 
      Caption         =   "&ReadMe File"
      Height          =   315
      Left            =   3878
      TabIndex        =   10
      Top             =   4073
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   315
      Left            =   3878
      TabIndex        =   2
      Top             =   4493
      Width           =   1155
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0442
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   788
      TabIndex        =   9
      Top             =   2813
      Width           =   3675
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Round Lake, Illinois  60073"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1058
      TabIndex        =   8
      Top             =   2513
      Width           =   3135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "317 W. Treehouse Lane"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1058
      TabIndex        =   7
      Top             =   2273
      Width           =   3135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Leo Mrozek"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1058
      TabIndex        =   6
      Top             =   2033
      Width           =   3135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To Register send $25.00 U.S. Check or Money Order to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   578
      TabIndex        =   5
      Top             =   1493
      Width           =   3915
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":04E0
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   165
      TabIndex        =   4
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label lblCopyRight 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright Caption"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   158
      TabIndex        =   3
      Top             =   1193
      Width           =   4935
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   158
      TabIndex        =   1
      Top             =   953
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Message Box Wizard 98 for Visual Basic 6.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   795
      Left            =   818
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload frmAbout
End Sub

Private Sub cmdRead_Click()
    Dim iret As Long
    Dim sFile As String
    
    sFile = App.Path & "\MsgWiz.RTF"
    iret = ShellExecute(Me.hwnd, vbNullString, sFile, _
        vbNullString, "c:\", SW_SHOWNORMAL)
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    lblCopyRight.Caption = App.LegalCopyright
End Sub
Private Sub Form_Paint()
    Dim i

    ScaleMode = vbPixels
    DrawStyle = 5
    DrawWidth = 1
    For i = 1 To ScaleHeight Step ScaleHeight \ 64
        Line (-1, i - 1)-(ScaleWidth, i + ScaleHeight \ 64), _
           RGB(0, 0, 255 - i * 255 \ ScaleHeight), BF
    Next i
End Sub

