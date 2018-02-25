VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12240
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   18720
   LinkTopic       =   "Form1"
   ScaleHeight     =   8.5
   ScaleMode       =   5  'Inch
   ScaleWidth      =   13
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Printer.Orientation = 2
    Me.PrintForm
End Sub
