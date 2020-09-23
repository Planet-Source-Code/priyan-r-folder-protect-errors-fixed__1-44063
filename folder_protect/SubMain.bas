Attribute VB_Name = "SubMain"
Public clsid$
Public commandline$
Option Explicit
Public Sub Main()
'----------------------------------------------------
'    Folder Protect
'   Programmed By Priyan
'  http://priyan.netfirms.com
'  priyanrrajeevan@rediffmail.com
'----------------------------------------------------
commandline = getcommandline()
If LCase(Command) = "about" Then  'Means user select about shortcut from start menu
    If App.PrevInstance Then End
    frmabout.Show vbModal
    End
ElseIf LCase(Command) = "setpass" Then 'Means user select Change password shortcut from start menu
    If App.PrevInstance Then End
    frmchangepass.Show
ElseIf LCase(Command) = "uninstall" Then  'Means user select Uninstall shortcut from start menu or control panel
    If App.PrevInstance Then End
    Dim ch%
    ch = MsgBox("Are  you sure you want to Uninstall " & App.Title, vbYesNo + vbQuestion, "Uninstall")
    If ch = 7 Then Exit Sub
    frmuninst.Show vbModal
    End
ElseIf commandline = "" Then
If App.PrevInstance Then End
    If Dir(addstrap(getwinsysdir, "pfolder.uni"), vbNormal) = "" Then
            'if above condition ="" then means the program not installed
        frminstall.Show 'Shows the install dialogue
        Exit Sub
    Else
        registersettings
        frmlockhelp.Show  'Means program installed so shows help screen
        Exit Sub
    End If
ElseIf Right$(commandline, 38) = "{21EC2020-3AEA-1069-A2DD-08002B30309D}" Then '
    'Above means user wants a folder to be unlocked
    frmunlock.Show
ElseIf Dir(commandline, vbDirectory) <> "" Then 'user wants to lock a folder
    frmlock.Show 'Shows the lock dialogue
End If
End Sub

Public Sub registersettings(Optional remove As Boolean = False)  'Registers the folder handlers
'-------------------- --
'If remove=True then removes the folder handlers
'----------------------
clsid = "clsid\{21EC2020-3AEA-1069-A2DD-08002B30309D}" 'CLSID of control panel
If remove = False Then
    setkeyvalue HK_CLASSES_ROOT, clsid & "\progid", REG_SZ, "", "pri_protected_file"
    setkeyvalue HK_CLASSES_ROOT, "pri_protected_file", REG_SZ, "", "Protected Folder"
    setkeyvalue HK_CLASSES_ROOT, "pri_protected_file\DefaultIcon", REG_SZ, "", addstrap(App.Path, App.EXEName & ".exe")
    setkeyvalue HK_CLASSES_ROOT, "pri_protected_file\shell", REG_SZ, "", "Unlock"
    setkeyvalue HK_CLASSES_ROOT, "pri_protected_file\shell\unlock\command", REG_SZ, "", """" & addstrap(programfilesdir, "priyan\folder_protect\protect.exe") & """ ""%1"""
    setkeyvalue HK_CLASSES_ROOT, "directory\shell\Lock Folder\command", REG_SZ, "", """" & addstrap(programfilesdir, "priyan\folder_protect\protect.exe") & """ ""%1"""
Else
    deletekey HK_CLASSES_ROOT, clsid & "\progid"
    deletekey HK_CLASSES_ROOT, "pri_protected_file"
    deletekey HK_CLASSES_ROOT, "directory\shell\Lock Folder"
End If

End Sub

Public Sub install()  'Installs the programs
On Error Resume Next
MkDir addstrap(programfilesdir, "priyan")
MkDir addstrap(programfilesdir, "priyan\folder_protect")
MkDir programspath & "\" & App.Title
FileCopy addstrap(App.Path, App.EXEName & ".exe"), addstrap(programfilesdir, "priyan\folder_protect\protect.exe")
FileCopy addstrap(App.Path, App.EXEName & ".exe"), addstrap(getwinsysdir, App.Title & "_uninstall.exe")
frminstall.Label4.Caption = "Creating shortcuts"
CreateShellLink addstrap(programfilesdir, "priyan\folder_protect\protect.exe"), programspath & "\" & App.Title & "\Change Password.lnk", "setpass"
CreateShellLink addstrap(programfilesdir, "priyan\folder_protect\protect.exe"), programspath & "\" & App.Title & "\About.lnk", "About"
CreateShellLink addstrap(programfilesdir, "priyan\folder_protect\protect.exe"), programspath & "\" & App.Title & "\Help.lnk"
CreateShellLink addstrap(getwinsysdir, App.Title & "_uninstall.exe"), programspath & "\" & App.Title & "\Uninstall.lnk", "Uninstall"
setkeyvalue HK_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & App.Title, REG_SZ, _
"Displayname", App.Title 'Adds uninstall option to Add/remove Control panel
setkeyvalue HK_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & App.Title, REG_SZ, _
"Uninstallstring", addstrap(getwinsysdir, App.Title & "_uninstall.exe uninstall")
MsgBox App.Title & " Installed", vbInformation, App.Title
Shell addstrap(programfilesdir, "priyan\folder_protect\protect.exe"), vbNormalFocus
Unload frminstall
End Sub
Public Sub seticon(ByVal frm As Form)
frm.Icon = frmico.Icon
Unload frmico
End Sub
