VERSION 5.00
Begin VB.Form frmunlock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unlock Folder "
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "frmunlock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Unlock Folder"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
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
      Left            =   4320
      MouseIcon       =   "frmunlock.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   240
      Width           =   510
   End
   Begin VB.Label Lblfolder 
      AutoSize        =   -1  'True
      Caption         =   "Folder name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   690
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   3315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter Password to Unlock Folder"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmunlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
If LCase(Text1.Text) = LCase(passconv.cvtpasstostring(getpassword)) Then
    unlockfolder
Else
    MsgBox "Invalied password ", vbCritical, App.Title
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End If
End Sub

Private Sub Form_Load()
centerform Me
Me.Caption = App.Title & " [ Unlock Folder ]"
Dim a$, pos%
pos = InStr(1, commandline, ".")
a = Left(commandline, pos - 1)
Me.Lblfolder.Caption = a
seticon Me
Image1.Picture = Me.Icon
End Sub
Public Sub unlockfolder()
    Dim a$
    Dim pos%
    pos = InStr(1, commandline, ".")
    a = Left(commandline, pos - 1)
    Name commandline As a
    MsgBox "Folder unlocked", vbInformation, App.Title
      ShellExecute Me.hwnd, "open", a, "", "", 1
      Unload Me
End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Label4_Click()
frmabout.Show vbModal
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.Command1.Value = True
End Sub

