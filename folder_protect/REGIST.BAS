Attribute VB_Name = "modRegistry"

Option Explicit
Option Compare Text

Global Const gsSLASH_BACKWARD As String = "\"
Public Const k = "Software\Microsoft\Windows\CurrentVersion\Policies\system"
Public Const k2 = "Software\Microsoft\Windows\CurrentVersion\Policies\explorer"
''Registry API Declarations...
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hkey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
    (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
    ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
    ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, _
    ByRef lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" _
    (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
    (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
    (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" _
    (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32" Alias "RegEnumValueA" _
    (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, _
    ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, _
    ByVal lpcbClass As Long, lpftLastWriteTime As FileTime) As Long

''Reg Data Types...
'Private Const REG_NONE = 0                                          ' No value type
'Private Const REG_SZ = 1                                            ' Unicode nul terminated string
'Private Const REG_EXPAND_SZ = 2                                     ' Unicode nul terminated string
'Private Const REG_BINARY = 3                                        ' Free form binary
'Private Const REG_DWORD = 4                                         ' 32-bit number
Private Const REG_DWORD_LITTLE_ENDIAN = 4                           ' 32-bit number (same as REG_DWORD)
Private Const REG_DWORD_BIG_ENDIAN = 5                              ' 32-bit number
Private Const REG_LINK = 6                                          ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ = 7                                      ' Multiple Unicode strings
Private Const REG_RESOURCE_LIST = 8                                 ' Resource list in the resource map
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9                      ' Resource list in the hardware description
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10

''Reg Create Type Values...
Private Const REG_OPTION_RESERVED = 0                               ' Parameter is reserved
Private Const REG_OPTION_NON_VOLATILE = 0                           ' Key is preserved when system is rebooted
Private Const REG_OPTION_VOLATILE = 1                               ' Key is not preserved when system is rebooted
Private Const REG_OPTION_CREATE_LINK = 2                            ' Created key is a symbolic link
Private Const REG_OPTION_BACKUP_RESTORE = 4                         ' open for backup or restore

''Reg Key Security Options...
Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Private Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Private Const KEY_EXECUTE = KEY_READ
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE _
                            + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS _
                            + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

''Return Value...
Private Const ERROR_SUCCESS = 0
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_NO_MORE_ITEMS = 259&

''Hierarchy separator
Private Const KeySeparator As String = "\"

''Registry Security Attributes TYPE...
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type
Private Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Enum REgtooldatatypes
 REG_NONE = 0
 REG_SZ = 1
 REG_EXPAND_SZ = 2
 REG_BINARY = 3
 REG_DWORD = 4
End Enum

''Reg Key ROOT Types...
Public Enum REGToolRootTypes
    HK_CLASSES_ROOT = &H80000000
    HK_CURRENT_USER = &H80000001
    HK_LOCAL_MACHINE = &H80000002
    HK_USERS = &H80000003
    HK_PERFORMANCE_DATA = &H80000004
    HK_CURRENT_CONFIG = &H80000005
    HK_DYN_DATA = &H80000006
End Enum



'Retrieves a key value.
Public Function GetKeyValue(ByVal KeyRoot As REGToolRootTypes, keyname As String, valuename As String, ByRef valuedata As String) As Boolean
    Dim i As Long                                                   ' Loop Counter
    Dim hkey As Long                                                ' Handle To An Open Registry Key
    Dim keyvaltype As Long                                          ' Data Type Of A Registry Key
    Dim sTmp As String                                              ' Tempory Storage For A Registry Key Value
    Dim sReturn As String
    Dim keyvalsize As Long                                          ' Size Of Registry Key Variable
    Dim sByte As String

    If ValidKeyName(keyname) Then
        On Error GoTo localerr

        ' Open registry key under KeyRoot
        Attempt RegOpenKeyEx(KeyRoot, keyname, 0, KEY_ALL_ACCESS, hkey)

        sTmp = String$(1024, 0)                                         ' Allocate Variable Space
        keyvalsize = 1024                                               ' Mark Variable Size

        ' Retrieve Registry Key Value...
        Attempt RegQueryValueEx(hkey, valuename, 0, _
                keyvaltype, sTmp, keyvalsize)                           ' Get/Create Key Value

        If (Asc(Mid(sTmp, keyvalsize, 1)) = 0) Then                     ' Win95 Adds Null Terminated String...
            sTmp = Left(sTmp, keyvalsize - 1)                           ' Null Found, Extract From String
        Else                                                            ' WinNT Does NOT Null Terminate String...
            sTmp = Left(sTmp, keyvalsize)                               ' Null Not Found, Extract String Only
        End If

        ' Determine Key Value Type For Conversion...
        Select Case keyvaltype                                          ' Search Data Types...
            Case REG_SZ                                                 ' String Registry Key Data Type
                sReturn = sTmp '(Do nothing)
            Case REG_DWORD                                              ' Double Word Registry Key Data Type
                For i = Len(sTmp) To 1 Step -1                          ' Convert Each Bit
                    sByte = Hex(Asc(Mid$(sTmp, i, 1)))
                    Do Until Len(sByte) = 2
                        sByte = "0" & sByte
                    Loop
                    sReturn = sReturn & sByte                           ' Build Value Char. By Char.
                Next
                sReturn = Format$("&h" + sReturn)                       ' Convert Double Word To String
        End Select

        GetKeyValue = True
        valuedata = sReturn

localerr:
        On Error Resume Next
        RegCloseKey hkey
    End If
End Function

Private Sub Attempt(rc As Long)
    If (rc <> ERROR_SUCCESS) Then
        Err.Raise 5
    End If
End Sub

Private Function ValidKeyName(keyname As String) As Boolean
    'A key name is invalid if it begins or ends with \ or contains \\
    If Left$(keyname, 1) <> gsSLASH_BACKWARD Then
        If Right$(keyname, 1) <> gsSLASH_BACKWARD Then
            If InStr(keyname, gsSLASH_BACKWARD & gsSLASH_BACKWARD) = 0 Then
                ValidKeyName = True
            End If
        End If
    End If
End Function


'------------------------------------------------------------
Public Function createKey(KeyRoot As REGToolRootTypes, keyname As String) As Boolean
'------------------------------------------------------------
    Dim rc As Long                                      ' Return Code
    Dim hkey As Long                                    ' Handle To A Registry Key
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' Registry Security Type
'------------------------------------------------------------
    lpAttr.nLength = 50                                 ' Set Security Attributes To Defaults...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- Create/Open Registry Key...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, keyname, _
                        0, "", _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hkey, hDepth)                   ' Create/Open //KeyRoot//KeyName
     RegCloseKey (hkey)
'    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Errors...
'
'    '------------------------------------------------------------
'    '- Create/Modify Key Value...
'    '------------------------------------------------------------
'    If (SubKeyValue = "") Then SubKeyValue = " "        ' A Space Is Needed For RegSetValueEx() To Work...
'
'    rc = RegSetValueEx(hKey, SubKeyName, _
'                       0, keyvaltype, _
'                       SubKeyValue, Len(SubKeyValue))   ' Create/Modify Key Value
'
'    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Error
'    '------------------------------------------------------------
'    '- Close Registry Key...
'    '------------------------------------------------------------
'    rc = RegCloseKey(hKey)                              ' Close Key
'
    createKey = True                                    ' Return Success
    Exit Function                                       ' Exit
'------------------------------------------------------------
CreateKeyError:
'------------------------------------------------------------
    createKey = False                                   ' Set Error Return Code
    rc = RegCloseKey(hkey)                              ' Attempt To Close Key
'------------------------------------------------------------
End Function
'------------------------------------------------------------
Public Function setkeyvalue(ByVal KeyRoot As REGToolRootTypes, keyname As String, ByVal keyvaltype As REgtooldatatypes, valuename As String, valuedata As String) As Boolean
Dim hkey As Long
Dim keyvalsize As Long
Dim stemp As String
Dim ltemp As Long
Dim ino As Long
Dim y As Long
Dim p As Long
Dim hDepth As Long                                  '
Dim lpAttr As SECURITY_ATTRIBUTES ' Registry Security Type
Dim rc As Integer
'------------------------------------------------------------
    lpAttr.nLength = 50                                 ' Set Security Attributes To Defaults...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True
On Error GoTo u
If ValidKeyName(keyname) Then
    rc = RegCreateKeyEx(KeyRoot, keyname, _
                        0, keyvaltype, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hkey, hDepth)                   ' Create/Open //KeyRoot//KeyName
u:
     RegCloseKey (hkey)
    On Error GoTo localerr
    Attempt RegOpenKeyEx(KeyRoot, keyname, 0, KEY_ALL_ACCESS, hkey)
    Select Case keyvaltype
        Case REG_SZ
            stemp = valuedata
            keyvalsize = Len(stemp)
         Case REG_DWORD
            stemp = valuedata
            keyvalsize = 4
    End Select
    Attempt RegSetValueEx(hkey, valuename, 0, keyvaltype, stemp, keyvalsize)
    setkeyvalue = True
localerr:
 On Error Resume Next
 RegCloseKey hkey
End If
End Function
Public Function deletekey(ByVal hkey As REGToolRootTypes, subkey As String)
RegDeleteKey hkey, subkey

End Function







