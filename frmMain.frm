VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MessageBoxWizzard98"
   ClientHeight    =   5550
   ClientLeft      =   1245
   ClientTop       =   1890
   ClientWidth     =   8055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5550
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk 
      Caption         =   "User Response Expected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   180
      TabIndex        =   16
      ToolTipText     =   "This selects whether a Select Case is created for each button"
      Top             =   4560
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.Frame frHelpFile 
      Caption         =   "Help File and Context ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   60
      TabIndex        =   50
      Top             =   2880
      Width           =   2835
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   1020
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtHelpFile 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2595
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "ID and File Name must both be entered otherwise it is ignored."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   60
         TabIndex        =   52
         Top             =   960
         Width           =   2715
      End
      Begin VB.Label Label8 
         Caption         =   "Context ID:"
         Height          =   255
         Left            =   180
         TabIndex        =   51
         Top             =   660
         Width           =   795
      End
   End
   Begin VB.Timer Timer1 
      Left            =   6900
      Top             =   720
   End
   Begin VB.Frame Frame2 
      Caption         =   "Declare Variables:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   49
      Top             =   2220
      Width           =   2835
      Begin VB.OptionButton optDeclare 
         Caption         =   "No"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optDeclare 
         Caption         =   "Yes"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdViewCode 
      Caption         =   "&View Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   26
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdCode 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   27
      Top             =   4920
      Width           =   1395
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3060
      TabIndex        =   25
      Top             =   4920
      Width           =   1995
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   3000
      ScaleHeight     =   615
      ScaleWidth      =   4935
      TabIndex        =   38
      Top             =   2640
      Width           =   4995
      Begin VB.OptionButton optButton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Abort/Retry/Ignore"
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   24
         Top             =   360
         Width           =   1755
      End
      Begin VB.OptionButton optButton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Retry/Cancel"
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   23
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton optButton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Yes/No/Cancel"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optButton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Yes/No"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   21
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton optButton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK/Cancel"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   360
         Width           =   1155
      End
      Begin VB.OptionButton optButton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK Only"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   19
         Top             =   120
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1155
      Left            =   3000
      ScaleHeight     =   1095
      ScaleWidth      =   4935
      TabIndex        =   28
      Top             =   3660
      Width           =   4995
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click on Icon to Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1080
         TabIndex        =   53
         Top             =   840
         Width           =   1965
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Left            =   3720
         Top             =   120
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   4020
         Picture         =   "frmMain.frx":0442
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Selected Icon"
         Height          =   195
         Index           =   6
         Left            =   3780
         TabIndex        =   41
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "None"
         Height          =   195
         Index           =   5
         Left            =   4065
         TabIndex        =   40
         Top             =   780
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   3060
         Picture         =   "frmMain.frx":0884
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "None"
         Height          =   195
         Index           =   4
         Left            =   3120
         TabIndex        =   35
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Information"
         Height          =   195
         Index           =   3
         Left            =   2220
         TabIndex        =   34
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exclamation"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   33
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Question"
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   32
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Critical"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   31
         Top             =   600
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2280
         Picture         =   "frmMain.frx":0CC6
         Top             =   60
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   1500
         Picture         =   "frmMain.frx":1108
         Top             =   60
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   720
         Picture         =   "frmMain.frx":154A
         Top             =   60
         Width           =   480
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":198C
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type Message Box:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   29
      Top             =   60
      Width           =   2835
      Begin VB.OptionButton optBoxType 
         Caption         =   "Input Box"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   1500
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
      Begin VB.OptionButton optBoxType 
         Caption         =   "Message Box"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1515
      End
   End
   Begin VB.Frame frDefButton 
      Caption         =   "Default Button:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   43
      Top             =   1440
      Width           =   2835
      Begin VB.OptionButton optDefButton 
         Caption         =   "4th"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   11
         Top             =   300
         Width           =   555
      End
      Begin VB.OptionButton optDefButton 
         Caption         =   "3rd"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   1380
         TabIndex        =   10
         Top             =   300
         Width           =   555
      End
      Begin VB.OptionButton optDefButton 
         Caption         =   "2nd"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   9
         Top             =   300
         Width           =   615
      End
      Begin VB.OptionButton optDefButton 
         Caption         =   "1st"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   555
      End
   End
   Begin VB.Frame frDefValue 
      Caption         =   "Input Box Default Value:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   44
      Top             =   1440
      Visible         =   0   'False
      Width           =   2835
      Begin VB.TextBox txtDefValue 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.Frame frModal 
      Caption         =   "Modal Style:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   37
      Top             =   720
      Width           =   2835
      Begin VB.OptionButton optModal 
         Caption         =   "None"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optModal 
         Caption         =   "Application"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optModal 
         Caption         =   "System"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame frPos 
      Caption         =   "Positioning (in Twips):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   45
      Top             =   720
      Visible         =   0   'False
      Width           =   2835
      Begin VB.TextBox txtPos 
         Height          =   285
         Index           =   1
         Left            =   1740
         TabIndex        =   6
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txtPos 
         Height          =   285
         Index           =   0
         Left            =   420
         TabIndex        =   5
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "Y:"
         Height          =   255
         Left            =   1500
         TabIndex        =   47
         Top             =   255
         Width           =   195
      End
      Begin VB.Label Label6 
         Caption         =   "X:"
         Height          =   255
         Left            =   180
         TabIndex        =   46
         Top             =   255
         Width           =   195
      End
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   3000
      TabIndex        =   17
      Text            =   "(Insert Title Here)"
      Top             =   240
      Width           =   4995
   End
   Begin VB.TextBox txtMessage 
      Height          =   1395
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Text            =   "frmMain.frx":1DCE
      Top             =   900
      Width           =   4995
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7020
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":20FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2418
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2732
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2D66
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3080
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":339A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":36B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRegister 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Unregistered"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   60
      TabIndex        =   48
      Top             =   4980
      Width           =   2835
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   42
      Top             =   30
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Button Style:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   39
      Top             =   2400
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Style:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   36
      Top             =   3420
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   30
      Top             =   660
      Width           =   825
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPreviewMessage 
         Caption         =   "&Preview Message"
      End
      Begin VB.Menu mnuViewCode 
         Caption         =   "&View Code"
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "&Register"
      End
      Begin VB.Menu mnuVarNames 
         Caption         =   "Variable &Names"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionRTL 
         Caption         =   "Right to &Left Reading"
      End
      Begin VB.Menu mnuOptionRJ 
         Caption         =   "&Right Justify"
      End
      Begin VB.Menu mnuOptionForeGnd 
         Caption         =   "&Force Foreground"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Sub cmdCode_Click()
    Connect.Hide
End Sub

Private Sub cmdPreview_Click()
    If optBoxType(0).Value Then
        ' MsgBox Code
        Call LoadCodeMsg
        If Len(txtHelpFile) = 0 Then
            Response = MsgBox(txtMessage, bStyle, txtTitle)
        Else
            If Len(txtID) = 0 Then
                Response = MsgBox(txtMessage, bStyle, txtTitle)
            Else
                Response = MsgBox(txtMessage, bStyle, txtTitle, txtHelpFile, txtID)
            End If
        End If
    Else
        ' InputBox Code
        Call LoadCodeInput
        If Len(txtPos(0)) = 0 And Len(txtPos(1)) = 0 Then
            If Len(txtHelpFile) = 0 Then
                Response = InputBox(aStyle(0), aStyle(1), aStyle(2))
            Else
                If Len(txtID) = 0 Then
                    Response = InputBox(aStyle(0), aStyle(1), aStyle(2))
                Else
                    Response = InputBox(aStyle(0), aStyle(1), aStyle(2), , , txtHelpFile, txtID)
                End If
            End If
        ElseIf Len(txtPos(0)) <> 0 And Len(txtPos(1)) <> 0 Then
            If Len(txtHelpFile) = 0 Then
                Response = InputBox(aStyle(0), aStyle(1), aStyle(2), txtPos(0), txtPos(1))
            Else
                If Len(txtID) = 0 Then
                    Response = InputBox(aStyle(0), aStyle(1), aStyle(2), txtPos(0), txtPos(1))
                Else
                    Response = InputBox(aStyle(0), aStyle(1), aStyle(2), txtPos(0), txtPos(1), txtHelpFile, txtID)
                End If
            End If
        ElseIf Len(txtPos(0)) = 0 And Len(txtPos(1)) <> 0 Then
            If Len(txtHelpFile) = 0 Then
                Response = InputBox(aStyle(0), aStyle(1), aStyle(2), , txtPos(1))
            Else
                If Len(txtID) = 0 Then
                    Response = InputBox(aStyle(0), aStyle(1), aStyle(2), , txtPos(1))
                Else
                    Response = InputBox(aStyle(0), aStyle(1), aStyle(2), , txtPos(1), txtHelpFile, txtID)
                End If
            End If
        ElseIf Len(txtPos(0)) <> 0 And Len(txtPos(1)) = 0 Then
            If Len(txtHelpFile) = 0 Then
                Response = InputBox(aStyle(0), aStyle(1), aStyle(2), txtPos(0))
            Else
                If Len(txtID) = 0 Then
                    Response = InputBox(aStyle(0), aStyle(1), aStyle(2), txtPos(0))
                Else
                    Response = InputBox(aStyle(0), aStyle(1), aStyle(2), txtPos(0), , txtHelpFile, txtID)
                End If
            End If
        ElseIf Len(txtPos(0)) <> 0 And Len(txtPos(1)) <> 0 Then
            If Len(txtHelpFile) = 0 Then
                Response = InputBox(aStyle(0), aStyle(1), aStyle(2), txtPos(0), txtPos(1))
            Else
                If Len(txtID) = 0 Then
                    Response = InputBox(aStyle(0), aStyle(1), aStyle(2), txtPos(0), txtPos(1))
                Else
                    Response = InputBox(aStyle(0), aStyle(1), aStyle(2), txtPos(0), txtPos(1), txtHelpFile, txtID)
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdViewCode_Click()
    
    If chk.Value = 1 Then
        bUserResponse = True
    Else
        bUserResponse = False
    End If
    
    strMsgText = txtMessage.Text
    strMsgTitle = txtTitle.Text
    mbMsgBoxType = optBoxType(0).Value
    lngXPos = CLng(Val(txtPos(0).Text))
    lngYPos = CLng(Val(txtPos(1).Text))
    bDeclareVars = optDeclare(0).Value
    strDefValue = txtDefValue.Text
    strHelpFile = txtHelpFile.Text
    strHelpFileID = txtID.Text
    
    If optBoxType(0).Value Then
        Call LoadCodeMsg
        Load frmPreview
        frmPreview.Show
    Else
        Call LoadCodeInput
        frmPreview.Show
    End If
End Sub

Private Sub Form_Load()
    ' Code to retrive from Registry and compare for registered number
    Dim Registered As String
    Dim sPath As String
    Dim sPath16 As String
    Dim sPath32 As String
    Dim sIcon As String
    Dim sVarNames(0 To 2) As String
    
    Registered = GetINIStr("UserSettings", "RegisterNumber", "Unregistered")
    sVarNames(0) = GetINIStr("UserSettings", "RTL", 0)
    sVarNames(1) = GetINIStr("UserSettings", "RJ", 0)
    sVarNames(2) = GetINIStr("UserSettings", "ForeGnd", 0)
    If sVarNames(0) = "1" Then mnuOptionRTL.Checked = True
    If sVarNames(1) = "1" Then mnuOptionRJ.Checked = True
    If sVarNames(2) = "1" Then mnuOptionForeGnd.Checked = True
    
    
    If Registered = "KXP4410" Then
        lblRegister.Caption = "Registered"
        lblRegister.Font.Size = 10
        lblRegister.ForeColor = &H8000000F
        lblRegister.BackColor = &H8000000F
        Timer1.Interval = 0
        Timer1.Enabled = False
        mnuRegister.Visible = False
    Else
        lblRegister.Caption = "Unregistered"
        Timer1.Interval = 60000
        Timer1.Enabled = True
        mnuRegister.Visible = True
    End If
    
    Image1(0).Picture = ImageList1.ListImages(1).Picture
    Image1(1).Picture = ImageList1.ListImages(2).Picture
    Image1(2).Picture = ImageList1.ListImages(3).Picture
    Image1(3).Picture = ImageList1.ListImages(4).Picture
    Image1(4).Picture = ImageList1.ListImages(9).Picture

End Sub

Private Sub Image1_Click(Index As Integer)
    Select Case Index
        Case 0  ' Critical
            IconStyle = "vbCritical"
            Image1(5).Picture = Image1(0).Picture
            Label2(5).Caption = Label2(0).Caption
        Case 1  ' Question
            IconStyle = "vbQuestion"
            Image1(5).Picture = Image1(1).Picture
            Label2(5).Caption = Label2(1).Caption
        Case 2  ' Exclamation
            IconStyle = "vbExclamation"
            Image1(5).Picture = Image1(2).Picture
            Label2(5).Caption = Label2(2).Caption
        Case 3  ' Information
            IconStyle = "vbInformation"
            Image1(5).Picture = Image1(3).Picture
            Label2(5).Caption = Label2(3).Caption
        Case 4  ' None
            IconStyle = ""
            Image1(5).Picture = Image1(4).Picture
            Label2(5).Caption = Label2(4).Caption
    End Select
End Sub
Private Sub LoadCodeInput()
    Dim x As Integer
    
    bStyle = 0
    For x = 0 To UBound(aStyle)
        aStyle(x) = ""
    Next
    aStyle(0) = txtMessage
    aStyle(1) = txtTitle
    aStyle(2) = txtDefValue
    aStyle(3) = Val(txtPos(0))
    aStyle(4) = Val(txtPos(1))
End Sub
Private Sub LoadCodeMsg()
    Dim x As Integer
    Dim sVarNames(0 To 2) As String
    
    bStyle = 0
    For x = 0 To UBound(aStyle)
        aStyle(x) = ""
    Next
    If optButton(0).Value Then
        aStyle(0) = "vbOKOnly"
        bStyle = 0
    ElseIf optButton(1).Value Then
        aStyle(0) = "vbOKCancel"
        bStyle = 1
    ElseIf optButton(2).Value Then
        aStyle(0) = "vbYesNo"
        bStyle = 4
    ElseIf optButton(3).Value Then
        aStyle(0) = "vbYesNoCancel"
        bStyle = 3
    ElseIf optButton(4).Value Then
        aStyle(0) = "vbRetryCancel"
        bStyle = 5
    ElseIf optButton(5).Value Then
        aStyle(0) = "vbAbortRetryIgnore"
        bStyle = 2
    End If
        
    If optModal(2).Value Then
        aStyle(1) = ""
    ElseIf optModal(0).Value Then
        aStyle(1) = "vbSystemModal"
        bStyle = bStyle + 4096
    ElseIf optModal(1).Value Then
        aStyle(1) = "vbApplicationModal"
        bStyle = bStyle + 0
    End If
      
    Select Case Label2(5).Caption
        Case "Critical"
            aStyle(2) = "vbCritical"
            bStyle = bStyle + 16
        Case "Question"
            aStyle(2) = "vbQuestion"
            bStyle = bStyle + 32
        Case "Exclamation"
            aStyle(2) = "vbExclamation"
            bStyle = bStyle + 48
        Case "Information"
            aStyle(2) = "vbInformation"
            bStyle = bStyle + 64
        Case "None"
    End Select
        
    If optDefButton(0).Value Then
        aStyle(3) = "vbDefaultButton1"
        bStyle = bStyle + 0
    ElseIf optDefButton(1).Value Then
        aStyle(3) = "vbDefaultButton2"
        bStyle = bStyle + 256
    ElseIf optDefButton(2).Value Then
        aStyle(3) = "vbDefaultButton3"
        bStyle = bStyle + 512
    ElseIf optDefButton(3).Value Then
        aStyle(3) = "vbDefaultButton4"
        bStyle = bStyle + 768
    End If
    
    If optDefButton(3).Enabled Then
        aStyle(3) = "vbMsgBoxHelpButton"
        bStyle = bStyle + 16384
    End If
    
    sVarNames(0) = GetINIStr("UserSettings", "RTL", 0)
    sVarNames(1) = GetINIStr("UserSettings", "RJ", 0)
    sVarNames(2) = GetINIStr("UserSettings", "ForeGnd", 0)
    
    If sVarNames(0) = "1" Then
        aStyle(3) = aStyle(3) & " + vbMsgBoxRTLReading"
        bStyle = bStyle + 1048576
    End If
    
    If sVarNames(1) = "1" Then
        aStyle(3) = aStyle(3) & " + vbMsgBoxRight"
        bStyle = bStyle + 524288
    End If
    
    If sVarNames(2) = "1" Then
        aStyle(3) = aStyle(3) & " + vbMsgBoxSetForeGround"
        bStyle = bStyle + 65536
    End If
    
    sStyle = ""
    For x = 0 To 3
        sStyle = sStyle & aStyle(x)
        If Len(aStyle(x)) > 0 And x <> 3 Then
            sStyle = sStyle & " + "
        End If
    Next
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuClose_Click()
    Call cmdCode_Click
End Sub

Private Sub mnuOptionForeGnd_Click()
    Dim iGood As Integer
    mnuOptionForeGnd.Checked = Not mnuOptionForeGnd.Checked
    iGood = WriteINIStr("UserSettings", "ForeGnd", IIf(mnuOptionForeGnd.Checked, 1, 0))
End Sub

Private Sub mnuOptionRJ_Click()
    Dim iGood As Integer
    mnuOptionRJ.Checked = Not mnuOptionRJ.Checked
    iGood = WriteINIStr("UserSettings", "RJ", IIf(mnuOptionRJ.Checked, 1, 0))
End Sub

Private Sub mnuOptionRTL_Click()
    Dim iGood As Integer
    mnuOptionRTL.Checked = Not mnuOptionRTL.Checked
    iGood = WriteINIStr("UserSettings", "RTL", IIf(mnuOptionRTL.Checked, 1, 0))
End Sub

Private Sub mnuPreviewMessage_Click()
    Call cmdPreview_Click
End Sub
Private Sub mnuRegister_Click()
    Dim iDone As Integer
    Dim Default As String
    
    Title = "Registration Number"
    Default = "Unregistered"
    Msg = ""
    Msg = Msg & "Please enter the Registration Number received" & vbCrLf
    Msg = Msg & "to mark you program as registered." & vbCrLf
    Response = InputBox(Msg, Title, Default)
    
    If UCase(Response) = "KXP4410" Then
        iDone = WriteINIStr("UserSettings", "RegisterNumber", "KXP4410")
        mnuRegister.Visible = False
        lblRegister.Caption = "Registered"
        lblRegister.Font.Size = 10
        lblRegister.ForeColor = &H8000000F
        lblRegister.BackColor = &H8000000F
        Timer1.Interval = 0
        Timer1.Enabled = False
        Style = vbOKOnly + vbApplicationModal + vbInformation + vbDefaultButton1
        Title = "ThankYou"
        Msg = "Thankyou for registering MessageBoxWizard98." & vbCrLf
        Msg = Msg & "    The prompt for registering has been turned off." & vbCrLf
        Response = MsgBox(Msg, Style, Title)
    Else
        iDone = WriteINIStr("UserSettings", "RegisterNumber", "Unregistered")
        mnuRegister.Visible = True
    End If
End Sub

Private Sub mnuVarNames_Click()
    frmVars.Show 1
End Sub

Private Sub mnuViewCode_Click()
    Call cmdViewCode_Click
End Sub

Private Sub optBoxType_Click(Index As Integer)
    Select Case Index
        Case 0
            frDefValue.Visible = False
            frDefButton.Visible = True
            frPos.Visible = False
            frModal.Visible = True
            Picture1.Enabled = True
            Picture2.Enabled = True
            chk.Enabled = True
            chk.Value = 1
        Case 1
            frDefValue.Visible = True
            frDefButton.Visible = False
            frPos.Visible = True
            frModal.Visible = False
            Picture1.Enabled = False
            Picture2.Enabled = False
            chk.Enabled = False
            chk.Value = 1
    End Select
    
End Sub

Private Sub optButton_Click(Index As Integer)
    Dim x As Integer
    
    intOptButton = Index
    
    For x = 0 To 2
        optDefButton(x).Enabled = True
    Next
    Select Case Index
        Case 0
            If optDefButton(1).Value Or optDefButton(2).Value Then
                optDefButton(0).Value = True
            End If
            optDefButton(1).Enabled = False
            optDefButton(2).Enabled = False
        Case 1
            If optDefButton(2).Value Then
                optDefButton(0).Value = True
            End If
            optDefButton(2).Enabled = False
        Case 2
            If optDefButton(2).Value Then
                optDefButton(0).Value = True
            End If
            optDefButton(2).Enabled = False
        Case 3
            ' Uses all Default Buttons
        Case 4
            If optDefButton(2).Value Then
                optDefButton(0).Value = True
            End If
            optDefButton(2).Enabled = False
        Case 5
    End Select
End Sub

Private Sub Timer1_Timer()
    Static iCount As Integer
    iCount = iCount + 1
    If iCount = 2 Then
        frmAbout.Show 1
        iCount = 0
    End If
End Sub

Private Sub txtHelpFile_Change()
    If Len(txtHelpFile) > 0 And Len(txtID) > 0 Then
        optDefButton(3).Enabled = True
    Else
        optDefButton(3).Enabled = False
    End If
End Sub

Private Sub txtID_Change()
    If Len(txtHelpFile) > 0 And Len(txtID) > 0 Then
        optDefButton(3).Enabled = True
    Else
        optDefButton(3).Enabled = True
    End If
End Sub

Private Sub txtMessage_GotFocus()
    txtMessage.SelStart = 0
    txtMessage.SelLength = Len(txtMessage)
End Sub

Private Sub txtPos_GotFocus(Index As Integer)
    txtPos(Index).SelStart = 0
    txtPos(Index).SelLength = Len(txtPos(Index))
End Sub
Private Sub txtTitle_GotFocus()
    txtTitle.SelStart = 0
    txtTitle.SelLength = Len(txtTitle)
End Sub
Private Sub txtDefValue_GotFocus()
    txtDefValue.SelStart = 0
    txtDefValue.SelLength = Len(txtDefValue)
End Sub
