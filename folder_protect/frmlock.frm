VERSION 5.00
Begin VB.Form frmlock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmlock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lock Folder"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Left            =   4440
      MouseIcon       =   "frmlock.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   360
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   1320
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
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   3315
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If LCase(Text1.Text) = LCase(passconv.cvtpasstostring(getpassword)) Then
    lockfolder
Else
    MsgBox "Invalied password ", vbCritical, App.Title
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End If

End Sub

Private Sub Form_Load()
centerform Me
seticon Me
Image1.Picture = Me.Icon
Me.Caption = App.Title & " [ Lock Folder ]"
Me.Lblfolder.Caption = commandline
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.Value = True
End Sub
Public Sub lockfolder()
   Name commandline As commandline & ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"
   'Above Renames the folder with control panel's CLSID
   Unload Me
   End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Label2_Click()
frmabout.Show vbModal
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.Command1.Value = True
End Sub
