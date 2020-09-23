VERSION 5.00
Begin VB.Form frmVars 
   Caption         =   "Change Variable Names"
   ClientHeight    =   1395
   ClientLeft      =   1170
   ClientTop       =   1515
   ClientWidth     =   5760
   Icon            =   "frmVars.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1395
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   60
      Top             =   900
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   435
      Left            =   2940
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtNewName 
      Height          =   315
      Left            =   3900
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox cboVarName 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtVarName 
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "New Name:"
      Height          =   195
      Left            =   3900
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Current Name:"
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Variable:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboVarName_Click()
    Timer1.Interval = 1
    Timer1.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload frmVars
End Sub

Private Sub cmdSave_Click()
    Dim iGood As Integer
    If Len(txtNewName) <> 0 Then
        iGood = WriteINIStr("UserSettings", cboVarName.Text, txtNewName)
        If iGood = -1 Then
            txtVarName = txtNewName
            txtNewName = ""
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim sDone As String
    
    cboVarName.AddItem "Message"
    cboVarName.AddItem "Style"
    cboVarName.AddItem "Title"
    cboVarName.AddItem "X-Position"
    cboVarName.AddItem "Y-Position"
    cboVarName.AddItem "Help File"
    cboVarName.AddItem "Help Context ID"
    cboVarName.AddItem "Return Value"
    cboVarName.AddItem "Default Value"
    
    sDone = GetINIStr("UserSettings", "Message", "sMsg")
    sDone = GetINIStr("UserSettings", "Style", "lStyle")
    sDone = GetINIStr("UserSettings", "Title", "sTitle")
    sDone = GetINIStr("UserSettings", "X-Position", "i_xPos")
    sDone = GetINIStr("UserSettings", "Y-Position", "i_yPos")
    sDone = GetINIStr("UserSettings", "Help File", "sHelpFile")
    sDone = GetINIStr("UserSettings", "Help Context ID", "iHelpFileID")
    sDone = GetINIStr("UserSettings", "Return Value", "Response")
    sDone = GetINIStr("UserSettings", "Default Value", "sDefault")


End Sub
Private Sub Timer1_Timer()
    Dim status As Long
    Dim x As Integer
    
    status = SendMessage(cboVarName.hwnd, CB_GETDROPPEDSTATE, 0, 0&)
    If status = 0 Then
        If cboVarName.ListIndex <> -1 Then
            txtVarName = GetINIStr("UserSettings", cboVarName.Text, "Undefined")
        End If
        Timer1.Enabled = False
    End If
End Sub

Private Sub cboVarName_DropDown()
    Timer1.Interval = 1
    Timer1.Enabled = True
End Sub

Private Sub cboVarName_GotFocus()
    Dim lResult As Long
    lResult = SendMessage(GetFocus(), CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub txtNewName_Change()
    If Len(txtNewName) = 0 Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Sub
