VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLive 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtTEST 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
n = 5
position = 0
Do
    position = InStr(position + 1, txtTEST.Text, "+")
    n = n - 1
Loop Until position = 0 Or n = 0
If position = Len(txtTEST.Text) Then
End
Else
Me.Show
End If
txtLive.Text = position
End Sub
