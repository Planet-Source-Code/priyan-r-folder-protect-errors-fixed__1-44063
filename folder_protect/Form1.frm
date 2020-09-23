VERSION 5.00
Begin VB.Form frmchangepass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   4815
   ClientLeft      =   3105
   ClientTop       =   1485
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5520
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
      Height          =   2535
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   4935
      Begin VB.TextBox txtconfirm 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtold 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change Password"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtnew 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label5 
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
         TabIndex        =   10
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
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
         TabIndex        =   8
         Top             =   840
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Change your master password"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   3375
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
         MouseIcon       =   "Form1.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   480
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmchangepass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If txtconfirm.Text <> txtnew.Text Then
    MsgBox "Confirm does not match", vbCritical, App.Title
    txtconfirm.SetFocus
    txtconfirm.SelStart = 0
    txtconfirm.SelLength = Len(txtconfirm.Text)
    Exit Sub
End If
If LCase(txtold.Text) <> LCase(passconv.cvtpasstostring(getpassword)) Then
    MsgBox "Invlied old password", vbCritical
    txtold.SetFocus
    txtold.SelStart = 0
    txtold.SelLength = Len(txtnew.Text)
    Exit Sub
End If
Open addstrap(getwinsysdir, "pfolder.uni") For Output As #1
    Write #1, passconv.cvtstringtopass(LCase(txtnew.Text))
Close #1
MsgBox "Password changed", vbInformation, App.Title
End
End Sub

Private Sub Form_Load()
seticon Me
Image1.Picture = Me.Icon
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text3_Change()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Label4_Click()
frmabout.Show vbModal
End Sub

Private Sub txtconfirm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.Value = True
End If
End Sub
