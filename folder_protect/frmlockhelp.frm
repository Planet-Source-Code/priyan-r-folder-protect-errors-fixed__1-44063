VERSION 5.00
Begin VB.Form frmlockhelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   1200
   ClientTop       =   975
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10320
   Begin VB.Label Label1 
      Caption         =   "You can make a folder password protected like this "
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   5985
      Left            =   120
      Picture         =   "frmlockhelp.frx":0000
      Top             =   720
      Width           =   10140
   End
End
Attribute VB_Name = "frmlockhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Caption = App.Title
seticon Me
'Me.Icon = LoadResPicture(101, 1)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

