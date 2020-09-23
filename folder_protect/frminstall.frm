VERSION 5.00
Begin VB.Form frminstall 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3000
   ClientLeft      =   3495
   ClientTop       =   1875
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4965
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label Label4 
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "Install"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Programmed  By Priyan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "http://priyan.netfirms.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   600
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1320
         Width           =   2610
      End
   End
End
Attribute VB_Name = "frminstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Me.Frame1.Visible = False
Me.Frame3.Visible = True
Me.Frame3.BorderStyle = 0
Frame3.Left = 0
Frame3.Top = 0
Me.Width = Frame3.Width
Me.Height = Frame3.Height
Me.Label4.Caption = "Setting Password"
FrmSetPass.Show vbModal
install
End Sub

Private Sub Form_Load()
Me.Frame1.BorderStyle = 0
Me.Caption = "Install " & App.Title
Label3.Caption = App.Title
Label3.Caption = "Installing " & App.Title
seticon Me
Me.Label2.Caption = "Version : " & App.Major & ".0"
End Sub

