Attribute VB_Name = "MsgWiz"
Option Explicit

Public Msg As String
Public Style As Long
Public Title As String
Public Response As Variant
Public IconStyle As String
Public sMessage As String

Public aStyle(0 To 6) As String
Public sStyle As String
Public bStyle As Long
Public gblINIFile As String

Public bUserResponse As Boolean
Public strMsgText As String
Public strMsgTitle As String
Public mbMsgBoxType As Boolean
Public lngXPos As Long
Public lngYPos As Long
Public bDeclareVars As Boolean
Public strDefValue As String
Public strHelpFile As String
Public strHelpFileID As String
Public intOptButton As Integer

Public Const WM_USER = &H400
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_SHOWDROPDOWN = WM_USER + 15
Public Const SW_SHOWNORMAL = 1

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, lParam As Any) As Long
Declare Function GetFocus Lib "USER32" () As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function OSWritePrivateProfileString% Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Declare Function OSGetPrivateProfileString% Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Integer, ByVal FileName$)

Public Function GetINIStr(cSection As String, ByVal cItem As String, ByVal cDefault As String)
    Dim ctemp As String
    Dim temp As String
    ctemp = GetSetting("MsgBWiz", cSection, cItem, cDefault)
    If Trim(UCase(ctemp)) = Trim(UCase(cDefault)) Then
        SaveSetting "MsgBWiz", cSection, cItem, cDefault
    End If
    GetINIStr = ctemp
End Function

Public Function WriteINIStr(ByVal cSection, ByVal cItem, ByVal cDefault) As Integer
    Dim temp As String
    SaveSetting "MsgBWiz", cSection, cItem, cDefault
    WriteINIStr = True
End Function

Public Function OpenINI(Optional cININame As Variant)

    OpenINI = True
    gblINIFile = "MsgBWiz"

    Dim nFile As Integer
    Dim cFile As String

    #If Win16 Then
        On Error Resume Next
        nFile = FreeFile
        cFile = gblINIFile + ".INI"
        Open cFile For Input As nFile
        If Err <> 0 Then
            If Err = 53 Then
                Open cFile For Output As nFile
            Else
                MsgBox Error$, vbCritical, "Error opening INI file - " & Str(Err)
                OpenINI = False
            End If
            Close nFile
        End If
    #End If

End Function


