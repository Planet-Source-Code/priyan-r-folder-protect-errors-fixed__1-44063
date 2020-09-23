VERSION 5.00
Begin VB.Form FrmSetPass 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "FrmSetPass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4695
      Begin VB.Image Image2 
         Height          =   480
         Left            =   240
         Top             =   240
         Width           =   480
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
         MouseIcon       =   "FrmSetPass.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label2 
         Caption         =   "Set your master password  this password is used to lock and unlock folders.Don't forgot the password"
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   3495
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4695
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Set Password"
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
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
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1560
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1320
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "FrmSetPass.frx":044E
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "FrmSetPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text1.Text = "" Then
    MsgBox "You should enter a valied password", vbCritical, App.Title
    Exit Sub
End If
If Text1.Text <> Text3.Text Then
    MsgBox "Confirm does not match", vbCritical, App.Title
    Text3.SetFocus
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Exit Sub
End If
Open addstrap(getwinsysdir, "pfolder.uni") For Output As #1
    Write #1, passconv.cvtstringtopass(LCase(Text1.Text))
Close #1
Unload Me
End Sub

Private Sub Form_Load()
seticon Me
Image2.Picture = Me.Icon
centerform Me
Me.Caption = App.Title & " [Set Password ]"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then
    End
End If
End Sub

Private Sub Label4_Click()
frmabout.Show vbModal
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.Command1.Value = True
End Sub
