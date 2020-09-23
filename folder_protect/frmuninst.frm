VERSION 5.00
Begin VB.Form frmuninst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Uninstall"
   ClientHeight    =   3315
   ClientLeft      =   3750
   ClientTop       =   2130
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4485
   Begin VB.CommandButton Command1 
      Caption         =   "Uninstall"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   240
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Uninstall "
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "About"
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
         Left            =   1200
         MouseIcon       =   "frmuninst.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   480
         Width           =   510
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   720
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Enter  password to uninstall"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
End
Attribute VB_Name = "frmuninst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error Resume Next
If LCase(Text1.Text) = LCase(passconv.cvtpasstostring(getpassword)) Then
    Kill addstrap(getwinsysdir, "pfolder.uni")
    Kill addstrap(programfilesdir, "priyan\folder_protect\protect.exe")
    Kill programspath & "\" & App.Title & "\Change Password.lnk"
    Kill programspath & "\" & App.Title & "\About.lnk"
    Kill programspath & "\" & App.Title & "\help.lnk"
    Kill programspath & "\" & App.Title & "\uninstall.lnk"
    RmDir programspath & "\" & App.Title
    RmDir addstrap(programfilesdir, "priyan\folder_protect")
    deletekey HK_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & App.Title
    registersettings True
    MsgBox App.Title & " Uninstalled", vbInformation, "Uninstall"
    End
Else
    MsgBox "Invalied password ", vbCritical, App.Title
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End If
End Sub

Private Sub Form_Load()
If Dir(addstrap(getwinsysdir, "pfolder.uni"), vbNormal) = "" Then
    MsgBox "The program is allready uninstalled", vbCritical, App.Title
    End
End If
seticon Me
Image1(1).Picture = Me.Icon
Me.Caption = "Uninstall "
Me.Label2.Caption = "Uninstall " & App.Title


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Label4_Click()
frmabout.Show vbModal
End Sub
