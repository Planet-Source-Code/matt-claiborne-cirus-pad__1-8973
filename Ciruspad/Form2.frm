VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Anus"
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Copyright VTC Technologies 2000"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   360
      Picture         =   "Form2.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
End Sub
