VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReadme 
   Caption         =   "ReadMe"
   ClientHeight    =   5940
   ClientLeft      =   270
   ClientTop       =   1680
   ClientWidth     =   10545
   Icon            =   "frmRead.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5940
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtb 
      Height          =   5295
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   9340
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmRead.frx":0442
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4740
      TabIndex        =   0
      Top             =   5520
      Width           =   1695
   End
End
Attribute VB_Name = "frmReadme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdClose_Click()
    Unload frmReadme
End Sub

Private Sub Form_Load()
    rtb.LoadFile App.Path & "\Readme.RTF", rtfRTF
End Sub
