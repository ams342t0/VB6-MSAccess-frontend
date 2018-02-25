VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BOOKS by P.S."
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1185
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   3360
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Label1.Caption = "BOOKS" & vbCrLf & _
                        " by PIRATED SOLUTIONS" & vbCrLf & _
                        " c/o ZCHHS" & vbCrLf & _
                        " No Rights Reserved, No Patent Pending."
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub
