Attribute VB_Name = "modcommon"
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTemppath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Public Declare Function GetForegroundWindow Lib "user32" () As Long
'Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit

Public Function addstrap(ByVal path1 As String, ByVal path2 As String) As String
If Right$(path1, 1) = "\" Then
     addstrap = path1 & path2
Else
         addstrap = path1 & "\" & path2
End If
End Function

Public Function getcommandline() As String  'removes " " from command4
On Error Resume Next
getcommandline = Mid$(Command, 2, Len(Command) - 2)
End Function
Public Function getwinsysdir() As String
Dim a$
Dim ret%
a$ = String$(256, Chr(0))
ret = GetSystemDirectory(a, 256)
getwinsysdir = Left$(a, ret) & "\"
End Function

Public Function getpassword() As String  'Reads the password from the file
Dim a$
Open addstrap(getwinsysdir, "pfolder.uni") For Input As #1
    Input #1, a$
Close #1
getpassword = a$
End Function
Public Function centerform(ByVal frm As Object)
Dim Form As Form
Set Form = frm
 Form.Top = (Screen.Height * 0.85) \ 2 - frm.Height \ 2
    Form.Left = Screen.Width \ 2 - frm.Width \ 2

End Function
Public Function CreateShellLink(LnkFile As String, SHORTCUT_PLACE As String, Optional argv As String)
'---------------------------------------------------------------
'       Creates a shortcut
'       lnkfile -The oriiginal file to make shortcut
'       SHORTCUT_PLACE - Path to place the shortcut
'       Argv(optional -If argv<>"" then set argv as arguments
'---------------------------------------------------------------

    Dim cShellLink As ShellLinkA
    Dim cPersistFile As IPersistFile
    Set cShellLink = New ShellLinkA
    Set cPersistFile = cShellLink
    With cShellLink
        .SetPath LnkFile
        .SetIconLocation LnkFile, 0
If argv <> "" Then
    .SetArguments argv
End If

        
    End With
    
    cShellLink.Resolve 0, SLR_UPDATE
    cPersistFile.Save StrConv(SHORTCUT_PLACE, vbUnicode), 0
    CreateShellLink = True

'---------------------------------------------------------------
ErrHandler:
'---------------------------------------------------------------
    Set cPersistFile = Nothing
    Set cShellLink = Nothing
'---------------------------------------------------------------
End Function
Public Function programspath() As String 'Gets startmenu\programs path
GetKeyValue HK_CURRENT_USER, "software\microsoft\windows\currentversion\explorer\shell folders", "programs", programspath
End Function
Public Function programfilesdir() As String 'Get programfiles path
GetKeyValue HK_LOCAL_MACHINE, "software\microsoft\windows\currentversion", "ProgramFilesDir", programfilesdir
End Function
Function GetTempdir() As String  'Get GetTemp path
Dim len_temp As Integer
GetTempdir = String(260, vbNullChar)
len_temp = GetTemppath(260, GetTempdir)
GetTempdir = Left$(GetTempdir, len_temp)
End Function
