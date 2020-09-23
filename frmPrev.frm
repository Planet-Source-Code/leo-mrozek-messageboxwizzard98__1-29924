VERSION 5.00
Begin VB.Form frmPreview 
   Caption         =   "Preview Code"
   ClientHeight    =   3735
   ClientLeft      =   1080
   ClientTop       =   1515
   ClientWidth     =   8610
   Icon            =   "frmPrev.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3735
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Co&py Code"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3180
      Width           =   1395
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3180
      Width           =   1395
   End
   Begin VB.TextBox txtPreview 
      Height          =   2715
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   8415
   End
   Begin VB.Label lblPreview 
      Caption         =   "Message Box Code"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Copy Code to copy all code to clipboard or select portion needed and copy to clipboard using Ctrl-C."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   3060
      TabIndex        =   3
      Top             =   3180
      Width           =   5415
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sString As String
Dim sVarNames(0 To 8) As String
Dim mintShowResponseCode As Integer

Private Sub cmdClose_Click()
    Unload frmPreview
End Sub
Private Sub cmdPaste_Click()
    Const CF_TEXT = 1
    Clipboard.Clear                     ' Clear Clipboard.
    Clipboard.SetText txtPreview.Text   ' Put text on Clipboard.
    
    Style = vbOKOnly + vbApplicationModal + vbInformation + vbDefaultButton1
    Title = "Code Copied"
    Msg = "Your code for this Message has been copied to the Clipboard. " & vbCrLf
    Msg = "Close this Add-in, open your code to the location required for " & vbCrLf
    Msg = Msg & "   the message box code and use Ctrl-V to paste the code " & vbCrLf
    Msg = Msg & "                                 into your project." & vbCrLf
    Response = MsgBox(Msg, Style, Title)
    Unload frmPreview
    
End Sub

Private Sub Form_Load()
    sString = ""
    If mbMsgBoxType Then
        lblPreview.Caption = "Message Box Source Code"
        Call MsgBoxCode
    Else
        lblPreview.Caption = "Input Box Source Code"
        Call InputBoxCode
    End If
    
    txtPreview = sString
    
End Sub
Private Sub InputBoxCode()
    Dim sTemp As String
    Dim x As Integer
    Dim y As Integer
    Dim sMsg() As String
        
    sVarNames(0) = GetINIStr("UserSettings", "Message", "strMsg")
    sVarNames(1) = GetINIStr("UserSettings", "Style", "lngStyle")
    sVarNames(2) = GetINIStr("UserSettings", "Title", "strTitle")
    sVarNames(3) = GetINIStr("UserSettings", "X-Position", "int_xPos")
    sVarNames(4) = GetINIStr("UserSettings", "Y-Position", "int_yPos")
    sVarNames(5) = GetINIStr("UserSettings", "Help File", "strHelpFile")
    sVarNames(6) = GetINIStr("UserSettings", "Help Context ID", "intHelpFileID")
    sVarNames(7) = GetINIStr("UserSettings", "Return Value", "vntResponse")
    sVarNames(8) = GetINIStr("UserSettings", "Default Value", "strDefault")
    
    sString = ""
    If bDeclareVars Then
        sString = "Dim " & sVarNames(0) & " as String" & vbCrLf
        sString = sString & "Dim " & sVarNames(8) & " as String" & vbCrLf
        sString = sString & "Dim " & sVarNames(2) & " as String" & vbCrLf
        If Len(lngXPos) <> 0 Then
            sString = sString & "Dim " & sVarNames(3) & " as Integer" & vbCrLf
        End If
        If Len(lngYPos) <> 0 Then
            sString = sString & "Dim " & sVarNames(4) & " as Integer" & vbCrLf
        End If
        If Len(strHelpFile) = 0 Then
            sString = sString & "Dim " & sVarNames(7) & " as Variant" & vbCrLf
        Else
            sString = sString & "Dim " & sVarNames(7) & " as Variant" & vbCrLf
            sString = sString & "Dim " & sVarNames(5) & " as String" & vbCrLf
            sString = sString & "Dim " & sVarNames(6) & " as Integer" & vbCrLf
        End If
    End If
    
    sString = sString & vbCrLf
    sString = sString & sVarNames(2) & " = " & """"
    sString = sString & strMsgTitle
    sString = sString & """" & vbCrLf

    sString = sString & sVarNames(8) & " = " & """"
    sString = sString & strDefValue
    sString = sString & """" & vbCrLf

    If Len(lngXPos) = 0 And Len(lngYPos) = 0 Then
        ' None to include in code
    ElseIf Len(lngXPos) <> 0 And Len(lngYPos) <> 0 Then
        sString = sString & sVarNames(3) & " = " & """"
        sString = sString & Val(lngXPos) & """" & vbCrLf
        sString = sString & sVarNames(4) & " = " & """"
        sString = sString & Val(lngYPos) & """" & vbCrLf
    ElseIf Len(lngXPos) = 0 And Len(lngYPos) <> 0 Then
        sString = sString & sVarNames(4) & " = " & """" & lngYPos & """" & vbCrLf
    ElseIf Len(lngYPos) = 0 And Len(lngXPos) <> 0 Then
        sString = sString & sVarNames(3) & " = " & """" & lngXPos & """" & vbCrLf
    End If

    If Len(strHelpFile) = 0 Then
    Else
        sString = sString & sVarNames(5) & " = " & """"
        sString = sString & strHelpFile
        sString = sString & """" & vbCrLf
        sString = sString & sVarNames(6) & " = " & strHelpFileID & vbCrLf
    End If
    Call ParseMsg(sVarNames(0))
    
    If Len(lngXPos) = 0 And Len(lngYPos) = 0 Then
        If Len(strHelpFile) = 0 Then
            sString = sString & sVarNames(7) & " = InputBox(" & sVarNames(0) & ", " & sVarNames(2) & ", " & sVarNames(8) & ")" & vbCrLf
        Else
            sString = sString & sVarNames(7) & " = InputBox(" & sVarNames(0) & ", " & sVarNames(2) & ", " & sVarNames(8) & ", , , " & sVarNames(5) & ", " & sVarNames(6) & ")" & vbCrLf
        End If
    ElseIf Len(lngXPos) <> 0 And Len(lngYPos) <> 0 Then
        If Len(strHelpFile) = 0 Then
            sString = sString & sVarNames(7) & " = InputBox(" & sVarNames(0) & ", " & sVarNames(2) & ", " & sVarNames(8) & ", " & sVarNames(3) & ", " & sVarNames(4) & ")" & vbCrLf
        Else
            sString = sString & sVarNames(7) & " = InputBox(" & sVarNames(0) & ", " & sVarNames(2) & ", " & sVarNames(8) & ", " & sVarNames(3) & ", " & sVarNames(4) & ", " & sVarNames(5) & ", " & sVarNames(6) & ")" & vbCrLf
        End If
    ElseIf Len(lngXPos) = 0 And Len(lngYPos) <> 0 Then
        If Len(strHelpFile) = 0 Then
            sString = sString & sVarNames(7) & " = InputBox(" & sVarNames(0) & ", " & sVarNames(2) & ", " & sVarNames(8) & ", ," & sVarNames(4) & ")" & vbCrLf
        Else
            sString = sString & sVarNames(7) & " = InputBox(" & sVarNames(0) & ", " & sVarNames(2) & ", " & sVarNames(8) & ", ," & sVarNames(4) & ", " & sVarNames(5) & ", " & sVarNames(6) & ")" & vbCrLf
        End If
    ElseIf Len(lngYPos) = 0 And Len(lngXPos) <> 0 Then
        If Len(strHelpFile) = 0 Then
            sString = sString & sVarNames(7) & " = InputBox(" & sVarNames(0) & ", " & sVarNames(2) & ", " & sVarNames(8) & "," & sVarNames(3) & ")" & vbCrLf
        Else
            sString = sString & sVarNames(7) & " = InputBox(" & sVarNames(0) & ", " & sVarNames(2) & ", " & sVarNames(8) & ", " & sVarNames(3) & ", , " & sVarNames(5) & ", " & sVarNames(6) & ")" & vbCrLf
        End If
    End If
    sString = sString & "If Len(" & sVarNames(7) & ") = 0 Then Exit Sub   ' User did not respond"
End Sub
Private Sub MsgBoxCode()
    sVarNames(0) = GetINIStr("UserSettings", "Message", "sMsg")
    sVarNames(1) = GetINIStr("UserSettings", "Style", "lStyle")
    sVarNames(2) = GetINIStr("UserSettings", "Title", "sTitle")
    sVarNames(3) = GetINIStr("UserSettings", "X-Position", "i_xPos")
    sVarNames(4) = GetINIStr("UserSettings", "Y-Position", "i_yPos")
    sVarNames(5) = GetINIStr("UserSettings", "Help File", "sHelpFile")
    sVarNames(6) = GetINIStr("UserSettings", "Help Context ID", "iHelpFileID")
    sVarNames(7) = GetINIStr("UserSettings", "Return Value", "Response")
    sVarNames(8) = GetINIStr("UserSettings", "Default Value", "sDefault")
    
    If bUserResponse Then
        mintShowResponseCode = 1
    Else
        mintShowResponseCode = 0
    End If
    
    sString = ""
    
    If bDeclareVars Then
        sString = "Dim " & sVarNames(0) & " as String" & vbCrLf
        sString = sString & "Dim " & sVarNames(1) & " as Long" & vbCrLf
        sString = sString & "Dim " & sVarNames(2) & " as String" & vbCrLf
        
        If Len(strHelpFile) = 0 Then
            If mintShowResponseCode = 1 Then
                sString = sString & "Dim " & sVarNames(7) & " as Integer" & vbCrLf & vbCrLf
            Else
                sString = sString & vbCrLf
            End If
        Else
            sString = sString & "Dim " & sVarNames(7) & " as Integer" & vbCrLf
            sString = sString & "Dim " & sVarNames(5) & " as String" & vbCrLf
            sString = sString & "Dim " & sVarNames(6) & " as Integer" & vbCrLf & vbCrLf
        End If
    End If
    
    sString = sString & sVarNames(1) & " = " & sStyle & vbCrLf
    
    sString = sString & sVarNames(2) & " = " & """"
    sString = sString & strMsgTitle
    sString = sString & """" & vbCrLf
    
    If Len(strHelpFile) = 0 Then
    Else
        sString = sString & sVarNames(5) & " = " & """"
        sString = sString & strHelpFile
        sString = sString & """" & vbCrLf
        sString = sString & sVarNames(6) & " = " & strHelpFileID & vbCrLf
    End If
    
    Call ParseMsg(sVarNames(0))
    
    If Len(strHelpFile) = 0 Then
        If mintShowResponseCode = 1 Then
            sString = sString & sVarNames(7) & " = MsgBox(" & sVarNames(0) & ", " & sVarNames(1) & ", " & sVarNames(2) & ")" & vbCrLf
        Else
            sString = sString & "MsgBox " & sVarNames(0) & ", " & sVarNames(1) & ", " & sVarNames(2) & vbCrLf
        End If
    Else
        sString = sString & sVarNames(7) & " = MsgBox(" & sVarNames(0) & ", " & sVarNames(1) & ", " & sVarNames(2) & ", " & sVarNames(5) & ", " & sVarNames(6) & ")" & vbCrLf
    End If
    
    If mintShowResponseCode = 1 Then
        
        sString = sString & "Select Case " & sVarNames(7) & vbCrLf
        If intOptButton = 0 Then
            sString = sString & "    Case vbOK" & vbCrLf & vbCrLf
        ElseIf intOptButton = 1 Then
            sString = sString & "    Case vbOK" & vbCrLf & vbCrLf
            sString = sString & "    Case vbCancel" & vbCrLf & vbCrLf
        ElseIf intOptButton = 2 Then
            sString = sString & "    Case vbYes" & vbCrLf & vbCrLf
            sString = sString & "    Case vbNo" & vbCrLf & vbCrLf
        ElseIf intOptButton = 3 Then
            sString = sString & "    Case vbYes" & vbCrLf & vbCrLf
            sString = sString & "    Case vbNo" & vbCrLf & vbCrLf
            sString = sString & "    Case vbCancel" & vbCrLf & vbCrLf
        ElseIf intOptButton = 4 Then
            sString = sString & "    Case vbRetry" & vbCrLf & vbCrLf
            sString = sString & "    Case vbCancel" & vbCrLf & vbCrLf
        ElseIf intOptButton = 5 Then
            sString = sString & "    Case vbAbort" & vbCrLf & vbCrLf
            sString = sString & "    Case vbRetry" & vbCrLf & vbCrLf
            sString = sString & "    Case vbIgnore" & vbCrLf & vbCrLf
        End If
        
        sString = sString & "End Select" & vbCrLf
    End If
End Sub

Private Sub ParseMsg(sVarNames As String)
    Dim sTemp As String
    Dim x As Integer
    Dim y As Integer
    Dim sMsg() As String

    y = 0
    ReDim sMsg(0 To 0) As String
    sMsg(0) = ""
    sTemp = strMsgText
    Do
        x = InStr(sTemp, Chr(13))
        If x = 0 Then
            ReDim Preserve sMsg(0 To y) As String
            sMsg(y) = """"
            sMsg(y) = sMsg(y) & sTemp
            sMsg(y) = sMsg(y) & """"
            sMsg(y) = sMsg(y) & " & vbCrLf"
            'sString = sString & sVarNames & " = "
            sString = sString & sVarNames & " = " & sVarNames & " & "
            sString = sString & sMsg(y) & vbCrLf
           Exit Do
        Else
            ReDim Preserve sMsg(0 To y) As String
            sMsg(y) = """"
            sMsg(y) = sMsg(y) & Left(sTemp, x - 1)
            sMsg(y) = sMsg(y) & """"
            sMsg(y) = sMsg(y) & " & vbCrLf"
            If y = 0 Then
                sString = sString & sVarNames & " = "
            Else
                sString = sString & sVarNames & " = " & sVarNames & " & "
            End If
            sString = sString & sMsg(y) & vbCrLf
            sTemp = Mid(sTemp, x + 2)
            If Len(sTemp) = 0 Then
                Exit Do
            End If
            y = y + 1
        End If
    Loop
End Sub
